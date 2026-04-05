from __future__ import annotations

from contextlib import contextmanager
import csv
from dataclasses import dataclass
import math
import os
from pathlib import Path
import time
from typing import Callable

from mu_unscramble_bot.github_answer_sheet import GitHubAnswerSheetClient, GitHubAnswerSheetConfig
from mu_unscramble_bot.models import Puzzle, SolverResult, normalize_letters


CSV_FIELDS = [
    "scrambled_letters",
    "answer",
    "frequency",
]

LEGACY_FIELDS = {
    "first_seen_at",
    "last_seen_at",
    "use_count",
    "round_number",
    "difficulty_level",
    "hint",
    "hint_lookup_key",
    "scrambled_word",
    "answer",
    "answer_letters",
    "source_method",
    "confidence",
}


def letters_match(answer: str, scramble: str) -> bool:
    return sorted(normalize_letters(answer)) == sorted(normalize_letters(scramble))


@dataclass(slots=True)
class MemoryRecord:
    scrambled_letters: str
    answer: str
    frequency: int

    @classmethod
    def from_row(cls, row: dict[str, str]) -> "MemoryRecord | None":
        scrambled_letters = normalize_letters(
            row.get("scrambled_letters", "")
            or row.get("scrambled_word", "")
        )
        answer = normalize_letters(
            row.get("answer", "")
            or row.get("answer_letters", "")
        )
        frequency_text = row.get("frequency", "") or row.get("use_count", "") or "1"
        try:
            frequency = max(1, int(float(frequency_text or "1")))
        except Exception:
            frequency = 1
        frequency = _normalize_frequency(frequency)

        if not scrambled_letters or not answer:
            return None
        return cls(
            scrambled_letters=scrambled_letters,
            answer=answer,
            frequency=frequency,
        )

    def to_row(self) -> dict[str, str]:
        return {
            "scrambled_letters": self.scrambled_letters,
            "answer": self.answer,
            "frequency": str(self.frequency),
        }


@dataclass(frozen=True, slots=True)
class DuplicateGroup:
    kind: str
    key: str
    records: tuple[MemoryRecord, ...]

    @property
    def label(self) -> str:
        if self.kind == "scramble":
            return f"Scramble conflict: {self.key} ({len(self.records)} answers)"
        return f"Answer conflict: {self.key} ({len(self.records)} scrambles)"


class QuestionMemory:
    def __init__(
        self,
        path: str | Path,
        *,
        fuzzy_match: bool = True,
        fuzzy_cutoff: float = 0.96,
        github_sync: GitHubAnswerSheetConfig | None = None,
    ) -> None:
        self.path = Path(path)
        self.fuzzy_match = fuzzy_match
        self.fuzzy_cutoff = fuzzy_cutoff
        self.records: list[MemoryRecord] = []
        self._last_file_stamp: tuple[int, int] | None = None
        self._lock_path = self.path.with_suffix(self.path.suffix + ".lock")
        self._github_client = GitHubAnswerSheetClient(github_sync) if github_sync is not None else None
        self._last_github_sync_at = 0.0
        self._loaded_legacy_schema = False
        self._load()

    def size(self) -> int:
        self._reload_if_changed()
        return len(self.records)

    def known_answers(self) -> list[str]:
        self._reload_if_changed()
        return sorted({record.answer for record in self.records if record.answer})

    def lookup(self, puzzle: Puzzle) -> SolverResult | None:
        self._reload_if_changed()
        matches = [record for record in self.records if record.scrambled_letters == puzzle.normalized_scramble]
        if not matches:
            return None

        matches.sort(key=lambda record: (record.frequency, record.answer), reverse=True)
        best = matches[0]
        second = matches[1] if len(matches) > 1 else None
        if second is not None and second.frequency == best.frequency and second.answer != best.answer:
            return None
        return SolverResult(answer=best.answer, method="memory", confidence=1.0)

    def remember(self, puzzle: Puzzle, result: SolverResult) -> None:
        scrambled_letters = puzzle.normalized_scramble
        answer = result.normalized_answer
        if not scrambled_letters or not answer:
            return
        if not letters_match(answer, scrambled_letters):
            return

        self._reload_if_changed()
        for record in self.records:
            if record.scrambled_letters != scrambled_letters:
                continue
            if record.answer != answer:
                continue
            record.frequency += 1
            self._save()
            return

        self.records.append(MemoryRecord(scrambled_letters=scrambled_letters, answer=answer, frequency=1))
        self._save()

    def find_duplicates(self, query: str = "") -> list[str]:
        self._reload_if_changed()
        needle = normalize_letters(query)

        lines: list[str] = []
        by_scramble: dict[str, set[str]] = {}
        by_answer: dict[str, set[str]] = {}
        for record in self.records:
            by_scramble.setdefault(record.scrambled_letters, set()).add(record.answer)
            by_answer.setdefault(record.answer, set()).add(record.scrambled_letters)

        scramble_duplicates = {
            scramble: sorted(answers)
            for scramble, answers in by_scramble.items()
            if len(answers) > 1
        }
        answer_duplicates = {
            answer: sorted(scrambles)
            for answer, scrambles in by_answer.items()
            if len(scrambles) > 1
        }

        if scramble_duplicates:
            lines.append("Same scramble mapped to multiple answers:")
            for scramble, answers in sorted(scramble_duplicates.items()):
                if needle and needle not in scramble and not any(needle in answer for answer in answers):
                    continue
                lines.append(f"{scramble} -> {', '.join(answers)}")

        if answer_duplicates:
            if lines:
                lines.append("")
            lines.append("Same answer seen under multiple scramble reads:")
            for answer, scrambles in sorted(answer_duplicates.items()):
                if needle and needle not in answer and not any(needle in scramble for scramble in scrambles):
                    continue
                lines.append(f"{answer} <- {', '.join(scrambles)}")

        return lines

    def duplicate_groups(self, query: str = "") -> list[DuplicateGroup]:
        self._reload_if_changed()
        needle = normalize_letters(query)

        groups: list[DuplicateGroup] = []
        by_scramble: dict[str, list[MemoryRecord]] = {}
        by_answer: dict[str, list[MemoryRecord]] = {}
        for record in self.records:
            by_scramble.setdefault(record.scrambled_letters, []).append(record)
            by_answer.setdefault(record.answer, []).append(record)

        for scramble, records in sorted(by_scramble.items()):
            unique_answers = {record.answer for record in records}
            if len(unique_answers) <= 1:
                continue
            if needle and needle not in scramble and not any(needle in record.answer for record in records):
                continue
            groups.append(
                DuplicateGroup(
                    kind="scramble",
                    key=scramble,
                    records=tuple(sorted(records, key=lambda item: (-item.frequency, item.answer))),
                )
            )

        for answer, records in sorted(by_answer.items()):
            unique_scrambles = {record.scrambled_letters for record in records}
            if len(unique_scrambles) <= 1:
                continue
            if needle and needle not in answer and not any(needle in record.scrambled_letters for record in records):
                continue
            groups.append(
                DuplicateGroup(
                    kind="answer",
                    key=answer,
                    records=tuple(sorted(records, key=lambda item: (-item.frequency, item.scrambled_letters))),
                )
            )

        groups.sort(key=lambda group: (group.kind, group.key))
        return groups

    def delete_records(self, rows: list[tuple[str, str]]) -> int:
        keys = {
            (normalize_letters(scramble), normalize_letters(answer))
            for scramble, answer in rows
            if normalize_letters(scramble) and normalize_letters(answer)
        }
        if not keys:
            return 0

        return self._apply_mutation(
            lambda records: [
                record
                for record in records
                if (record.scrambled_letters, record.answer) not in keys
            ]
        )

    def keep_record_for_group(self, group_kind: str, group_key: str, keep_row: tuple[str, str]) -> int:
        normalized_group_key = normalize_letters(group_key)
        keep_scramble = normalize_letters(keep_row[0])
        keep_answer = normalize_letters(keep_row[1])
        if not normalized_group_key or not keep_scramble or not keep_answer:
            return 0

        def mutate(records: list[MemoryRecord]) -> list[MemoryRecord]:
            kept: list[MemoryRecord] = []
            for record in records:
                if group_kind == "scramble" and record.scrambled_letters == normalized_group_key:
                    if record.answer != keep_answer:
                        continue
                elif group_kind == "answer" and record.answer == normalized_group_key:
                    if record.scrambled_letters != keep_scramble:
                        continue
                kept.append(record)
            return kept

        return self._apply_mutation(mutate)

    def _load(self, *, force: bool = False) -> None:
        stamp = self._file_stamp()
        if not force and stamp == self._last_file_stamp:
            return

        loaded_records, used_legacy_schema = self._read_records_from_disk()
        self.records = self._canonicalize_records(loaded_records)
        self._last_file_stamp = stamp
        self._loaded_legacy_schema = used_legacy_schema
        if used_legacy_schema and self.records:
            self._save()

    def _save(self) -> None:
        self.records = self._canonicalize_records(self.records)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._acquire_lock():
            disk_records, _ = self._read_records_from_disk()
            if disk_records:
                self.records = self._canonicalize_records(self.records + disk_records)

            if self._github_client is not None:
                remote_records = self._fetch_github_records()
                if remote_records:
                    self.records = self._canonicalize_records(self.records + remote_records)

            temp_path = self.path.with_suffix(self.path.suffix + f".{os.getpid()}.tmp")
            with temp_path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
                writer.writeheader()
                for record in self.records:
                    writer.writerow(record.to_row())
            temp_path.replace(self.path)
            self._last_file_stamp = self._file_stamp()

        if self._github_client is not None:
            self._push_to_github()
            self._last_github_sync_at = time.monotonic()

    def _apply_mutation(self, mutator: Callable[[list[MemoryRecord]], list[MemoryRecord]]) -> int:
        for _ in range(3):
            snapshot_sha: str | None = None
            with self._acquire_lock():
                disk_records, _ = self._read_records_from_disk()
                base_records = self._canonicalize_records(disk_records)

                if self._github_client is not None:
                    try:
                        snapshot = self._github_client.fetch()
                    except Exception:
                        snapshot = None
                    if snapshot is not None:
                        snapshot_sha = snapshot.sha
                        remote_records = self._parse_csv_text(snapshot.text)
                        if remote_records:
                            base_records = self._canonicalize_records(base_records + remote_records)

                updated_records = self._canonicalize_records(mutator(list(base_records)))
                removed_count = max(0, len(base_records) - len(updated_records))
                if removed_count == 0:
                    self.records = updated_records
                    self._last_file_stamp = self._file_stamp()
                    return 0

                self.records = updated_records
                self._write_records_exact(updated_records)

            if self._github_client is not None and (self._github_client.config.token or "").strip():
                try:
                    self._github_client.push(self._serialize_csv_text(), sha=snapshot_sha)
                except Exception:
                    continue

            self._last_github_sync_at = time.monotonic()
            return removed_count
        return 0

    def _fetch_github_records(self) -> list[MemoryRecord]:
        if self._github_client is None:
            return []
        try:
            snapshot = self._github_client.fetch()
        except Exception:
            return []
        return self._parse_csv_text(snapshot.text)

    def _push_to_github(self) -> None:
        if self._github_client is None:
            return
        token = (self._github_client.config.token or "").strip()
        if not token:
            return

        for _ in range(3):
            try:
                snapshot = self._github_client.fetch()
            except Exception:
                return

            remote_records = self._parse_csv_text(snapshot.text)
            if remote_records:
                merged = self._canonicalize_records(self.records + remote_records)
                if len(merged) != len(self.records):
                    self.records = merged
                    self._save_local_only()

            try:
                self._github_client.push(self._serialize_csv_text(), sha=snapshot.sha)
            except Exception:
                continue

            self._last_github_sync_at = time.monotonic()
            self._last_file_stamp = self._file_stamp()
            return

    def _reload_if_changed(self) -> None:
        self._sync_from_github_if_due()
        self._load()

    def _read_records_from_disk(self) -> tuple[list[MemoryRecord], bool]:
        if not self.path.exists():
            return [], False

        with self.path.open("r", encoding="utf-8", newline="") as handle:
            reader = csv.DictReader(handle)
            fieldnames = set(reader.fieldnames or [])
            used_legacy_schema = bool(fieldnames & LEGACY_FIELDS) and "scrambled_letters" not in fieldnames
            records = [record for row in reader if (record := MemoryRecord.from_row(row)) is not None]
        return records, used_legacy_schema

    def _sync_from_github_if_due(self, *, force: bool = False) -> None:
        if self._github_client is None:
            return
        if not force and (time.monotonic() - self._last_github_sync_at) < self._github_client.config.sync_interval_seconds:
            return
        try:
            snapshot = self._github_client.fetch()
        except Exception:
            self._last_github_sync_at = time.monotonic()
            return

        remote_records = self._parse_csv_text(snapshot.text)
        if remote_records:
            merged = self._canonicalize_records(self.records + remote_records)
            if merged != self.records:
                self.records = merged
                self._save_local_only()
        self._last_github_sync_at = time.monotonic()

    def _file_stamp(self) -> tuple[int, int] | None:
        try:
            stat = self.path.stat()
        except FileNotFoundError:
            return None
        return (int(stat.st_mtime_ns), int(stat.st_size))

    @contextmanager
    def _acquire_lock(self, timeout_seconds: float = 8.0):
        self._lock_path.parent.mkdir(parents=True, exist_ok=True)
        started_at = time.monotonic()
        handle: int | None = None

        while handle is None:
            try:
                handle = os.open(str(self._lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(handle, f"{os.getpid()}|{time.time()}".encode("utf-8"))
                break
            except FileExistsError:
                self._clear_stale_lock(max_age_seconds=30.0)
                if time.monotonic() - started_at >= timeout_seconds:
                    raise TimeoutError(f"Timed out waiting for memory lock: {self._lock_path}")
                time.sleep(0.1)

        try:
            yield
        finally:
            if handle is not None:
                os.close(handle)
            try:
                self._lock_path.unlink()
            except FileNotFoundError:
                pass

    def _clear_stale_lock(self, *, max_age_seconds: float) -> None:
        try:
            stat = self._lock_path.stat()
        except FileNotFoundError:
            return
        age_seconds = time.time() - stat.st_mtime
        if age_seconds <= max_age_seconds:
            return
        try:
            self._lock_path.unlink()
        except FileNotFoundError:
            return

    def _save_local_only(self) -> None:
        self.records = self._canonicalize_records(self.records)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._write_records_exact(self.records)

    def _write_records_exact(self, records: list[MemoryRecord]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = self.path.with_suffix(self.path.suffix + f".{os.getpid()}.tmp")
        with temp_path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
            writer.writeheader()
            for record in records:
                writer.writerow(record.to_row())
        temp_path.replace(self.path)
        self._last_file_stamp = self._file_stamp()

    def _parse_csv_text(self, text: str) -> list[MemoryRecord]:
        if not text.strip():
            return []
        reader = csv.DictReader(text.splitlines())
        return [record for row in reader if (record := MemoryRecord.from_row(row)) is not None]

    def _serialize_csv_text(self) -> str:
        from io import StringIO

        buffer = StringIO()
        writer = csv.DictWriter(buffer, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for record in self.records:
            writer.writerow(record.to_row())
        return buffer.getvalue()

    @staticmethod
    def _canonicalize_records(records: list[MemoryRecord]) -> list[MemoryRecord]:
        merged: dict[tuple[str, str], MemoryRecord] = {}
        for record in records:
            scrambled_letters = normalize_letters(record.scrambled_letters)
            answer = normalize_letters(record.answer)
            if not scrambled_letters or not answer:
                continue
            key = (scrambled_letters, answer)
            existing = merged.get(key)
            if existing is None:
                merged[key] = MemoryRecord(
                    scrambled_letters=scrambled_letters,
                    answer=answer,
                    frequency=max(1, int(record.frequency)),
                )
                continue
            existing.frequency += _normalize_frequency(record.frequency)

        return sorted(
            merged.values(),
            key=lambda item: (-item.frequency, item.scrambled_letters, item.answer),
        )


def _normalize_frequency(value: int) -> int:
    value = max(1, int(value))
    if value <= 1000:
        return value
    return max(1, int(round(math.log2(value))))
