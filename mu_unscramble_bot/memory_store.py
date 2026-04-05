from __future__ import annotations

from contextlib import contextmanager
import csv
from dataclasses import dataclass
from datetime import datetime
from difflib import get_close_matches
import os
from pathlib import Path
import time

from mu_unscramble_bot.github_answer_sheet import GitHubAnswerSheetClient, GitHubAnswerSheetConfig
from mu_unscramble_bot.models import (
    Puzzle,
    SolverResult,
    normalize_letters,
    normalize_lookup_text,
    sanitize_hint_text,
)


CSV_FIELDS = [
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
]


def letters_match(answer: str, scramble: str) -> bool:
    return sorted(normalize_letters(answer)) == sorted(normalize_letters(scramble))


@dataclass(slots=True)
class MemoryRecord:
    first_seen_at: str
    last_seen_at: str
    use_count: int
    round_number: int | None
    difficulty_level: int | None
    hint: str
    hint_lookup_key: str
    scrambled_word: str
    answer: str
    answer_letters: str
    source_method: str
    confidence: float

    @classmethod
    def from_row(cls, row: dict[str, str]) -> "MemoryRecord":
        return cls(
            first_seen_at=row.get("first_seen_at", ""),
            last_seen_at=row.get("last_seen_at", ""),
            use_count=int(row.get("use_count", "0") or 0),
            round_number=_to_int(row.get("round_number", "")),
            difficulty_level=_to_int(row.get("difficulty_level", "")),
            hint=row.get("hint", ""),
            hint_lookup_key=row.get("hint_lookup_key", ""),
            scrambled_word=row.get("scrambled_word", ""),
            answer=row.get("answer", ""),
            answer_letters=row.get("answer_letters", ""),
            source_method=row.get("source_method", ""),
            confidence=float(row.get("confidence", "0") or 0),
        )

    def to_row(self) -> dict[str, str]:
        return {
            "first_seen_at": self.first_seen_at,
            "last_seen_at": self.last_seen_at,
            "use_count": str(self.use_count),
            "round_number": "" if self.round_number is None else str(self.round_number),
            "difficulty_level": "" if self.difficulty_level is None else str(self.difficulty_level),
            "hint": self.hint,
            "hint_lookup_key": self.hint_lookup_key,
            "scrambled_word": self.scrambled_word,
            "answer": self.answer,
            "answer_letters": self.answer_letters,
            "source_method": self.source_method,
            "confidence": f"{self.confidence:.3f}",
        }


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
        self._dirty = False
        self._last_save_at = 0.0
        self._last_file_stamp: tuple[int, int] | None = None
        self._lock_path = self.path.with_suffix(self.path.suffix + ".lock")
        self._github_client = GitHubAnswerSheetClient(github_sync) if github_sync is not None else None
        self._last_github_sync_at = 0.0
        self._load()

    def size(self) -> int:
        self._reload_if_changed()
        return len(self.records)

    def known_answers(self) -> list[str]:
        self._reload_if_changed()
        answers = {record.answer_letters for record in self.records if record.answer_letters}
        return sorted(answers)

    def lookup(self, puzzle: Puzzle) -> SolverResult | None:
        self._reload_if_changed()
        record = self._find_exact(puzzle)
        if record:
            self._touch_record(record)
            return SolverResult(answer=record.answer, method="memory", confidence=1.0)

        if not self.fuzzy_match or not puzzle.hint_lookup_key:
            return None

        record = self._find_fuzzy(puzzle)
        if not record:
            return None

        self._touch_record(record)
        return SolverResult(answer=record.answer, method="memory:fuzzy", confidence=0.98)

    def remember(self, puzzle: Puzzle, result: SolverResult) -> None:
        if not result.normalized_answer:
            return

        self._reload_if_changed()
        clean_hint = sanitize_hint_text(puzzle.hint)
        hint_key = normalize_lookup_text(clean_hint)
        now = _now()
        for record in self.records:
            if record.answer_letters != result.normalized_answer:
                continue
            if not letters_match(record.answer_letters, puzzle.normalized_scramble):
                continue
            if not _hint_keys_related(record.hint_lookup_key, hint_key):
                continue

            record.last_seen_at = now
            record.use_count += 1
            record.round_number = puzzle.round_number
            record.difficulty_level = puzzle.difficulty_level
            record.scrambled_word = puzzle.scrambled_word
            record.source_method = result.method
            record.confidence = result.confidence
            if len(clean_hint) >= len(record.hint):
                record.hint = clean_hint
                record.hint_lookup_key = hint_key
            self._save()
            return

        for record in self.records:
            if record.hint_lookup_key != hint_key:
                continue

            record.last_seen_at = now
            record.use_count += 1
            record.round_number = puzzle.round_number
            record.difficulty_level = puzzle.difficulty_level
            record.scrambled_word = puzzle.scrambled_word
            if record.answer_letters != result.normalized_answer:
                if result.method.startswith("observed-") or result.confidence >= record.confidence:
                    record.answer = result.normalized_answer
                    record.answer_letters = result.normalized_answer
                    record.source_method = result.method
                    record.confidence = result.confidence
            else:
                record.source_method = result.method
                record.confidence = result.confidence
            if len(clean_hint) >= len(record.hint):
                record.hint = clean_hint
                record.hint_lookup_key = hint_key
            self._save()
            return

        self.records.append(
            MemoryRecord(
                first_seen_at=now,
                last_seen_at=now,
                use_count=1,
                round_number=puzzle.round_number,
                difficulty_level=puzzle.difficulty_level,
                hint=clean_hint,
                hint_lookup_key=hint_key,
                scrambled_word=puzzle.scrambled_word,
                answer=result.normalized_answer,
                answer_letters=result.normalized_answer,
                source_method=result.method,
                confidence=result.confidence,
            )
        )
        self._save()

    def _find_exact(self, puzzle: Puzzle) -> MemoryRecord | None:
        candidates = [record for record in self.records if record.hint_lookup_key == puzzle.hint_lookup_key]
        return self._pick_best(candidates, puzzle)

    def _find_fuzzy(self, puzzle: Puzzle) -> MemoryRecord | None:
        keys = list({record.hint_lookup_key for record in self.records if record.hint_lookup_key})
        matches = get_close_matches(puzzle.hint_lookup_key, keys, n=3, cutoff=self.fuzzy_cutoff)
        candidates = [record for record in self.records if record.hint_lookup_key in matches]
        return self._pick_best(candidates, puzzle)

    def _pick_best(self, candidates: list[MemoryRecord], puzzle: Puzzle) -> MemoryRecord | None:
        valid = [record for record in candidates if letters_match(record.answer_letters, puzzle.normalized_scramble)]
        if not valid:
            return None
        valid.sort(key=lambda record: (record.use_count, record.last_seen_at), reverse=True)
        return valid[0]

    def _touch_record(self, record: MemoryRecord) -> None:
        record.last_seen_at = _now()
        record.use_count += 1

    def _load(self, *, force: bool = False) -> None:
        stamp = self._file_stamp()
        if not force and stamp == self._last_file_stamp:
            return

        loaded_records = self._read_records_from_disk()
        self.records = self._canonicalize_records(loaded_records)
        self._last_file_stamp = stamp
        if len(self.records) != len(loaded_records):
            self._save()

    def _save(self) -> None:
        self.records = self._canonicalize_records(self.records)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._acquire_lock():
            disk_records = self._read_records_from_disk()
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
        self._dirty = False
        self._last_save_at = time.monotonic()
        if self._github_client is not None:
            self._push_to_github()
            self._last_github_sync_at = time.monotonic()

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
                new_sha = self._github_client.push(self._serialize_csv_text(), sha=snapshot.sha)
            except Exception:
                continue

            self._last_github_sync_at = time.monotonic()
            if new_sha:
                self._last_file_stamp = self._file_stamp()
            return

    def _reload_if_changed(self) -> None:
        self._sync_from_github_if_due()
        if self._dirty:
            return
        self._load()

    def _read_records_from_disk(self) -> list[MemoryRecord]:
        if not self.path.exists():
            return []

        with self.path.open("r", encoding="utf-8", newline="") as handle:
            reader = csv.DictReader(handle)
            return [MemoryRecord.from_row(row) for row in reader]

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
            if len(merged) != len(self.records):
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
        temp_path = self.path.with_suffix(self.path.suffix + f".{os.getpid()}.tmp")
        with temp_path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
            writer.writeheader()
            for record in self.records:
                writer.writerow(record.to_row())
        temp_path.replace(self.path)
        self._last_file_stamp = self._file_stamp()

    def _parse_csv_text(self, text: str) -> list[MemoryRecord]:
        if not text.strip():
            return []
        reader = csv.DictReader(text.splitlines())
        return [MemoryRecord.from_row(row) for row in reader]

    def _serialize_csv_text(self) -> str:
        from io import StringIO

        buffer = StringIO()
        writer = csv.DictWriter(buffer, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for record in self.records:
            writer.writerow(record.to_row())
        return buffer.getvalue()

    def _canonicalize_records(self, records: list[MemoryRecord]) -> list[MemoryRecord]:
        normalized: list[MemoryRecord] = []
        for record in records:
            record.hint = sanitize_hint_text(record.hint)
            record.hint_lookup_key = normalize_lookup_text(record.hint or record.hint_lookup_key)
            if not record.answer_letters and record.answer:
                record.answer_letters = normalize_letters(record.answer)
            normalized.append(record)

        merged_by_key: dict[str, MemoryRecord] = {}
        passthrough: list[MemoryRecord] = []
        for record in normalized:
            if not record.hint_lookup_key:
                passthrough.append(record)
                continue

            existing = merged_by_key.get(record.hint_lookup_key)
            if existing is None:
                merged_by_key[record.hint_lookup_key] = record
                continue
            self._merge_record(existing, record)

        collapsed = list(merged_by_key.values()) + passthrough
        related_merged: list[MemoryRecord] = []
        for record in sorted(
            collapsed,
            key=lambda item: (item.answer_letters, len(item.hint_lookup_key), item.use_count),
            reverse=True,
        ):
            existing = next(
                (
                    current
                    for current in related_merged
                    if current.answer_letters == record.answer_letters
                    and _hint_keys_related(current.hint_lookup_key, record.hint_lookup_key)
                ),
                None,
            )
            if existing is None:
                related_merged.append(record)
                continue
            self._merge_record(existing, record)

        related_merged.sort(key=lambda item: (item.last_seen_at, item.hint_lookup_key))
        return related_merged

    def _merge_record(self, target: MemoryRecord, incoming: MemoryRecord) -> None:
        target.first_seen_at = min(filter(None, [target.first_seen_at, incoming.first_seen_at]), default="")
        target.last_seen_at = max(filter(None, [target.last_seen_at, incoming.last_seen_at]), default="")
        target.use_count += incoming.use_count
        if incoming.round_number is not None:
            target.round_number = incoming.round_number
        if incoming.difficulty_level is not None:
            target.difficulty_level = incoming.difficulty_level
        if len(incoming.hint) > len(target.hint):
            target.hint = incoming.hint
            target.hint_lookup_key = incoming.hint_lookup_key
        if len(incoming.scrambled_word) > len(target.scrambled_word):
            target.scrambled_word = incoming.scrambled_word
        if self._prefer_record(incoming, target):
            target.answer = incoming.answer
            target.answer_letters = incoming.answer_letters
            target.source_method = incoming.source_method
            target.confidence = incoming.confidence

    @staticmethod
    def _prefer_record(left: MemoryRecord, right: MemoryRecord) -> bool:
        return _record_rank(left) > _record_rank(right)


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _to_int(value: str) -> int | None:
    value = value.strip()
    if not value:
        return None
    return int(value)


def _hint_keys_related(left: str, right: str) -> bool:
    if not left or not right:
        return False
    return left.startswith(right) or right.startswith(left)


def _record_rank(record: MemoryRecord) -> tuple[int, float, int, str]:
    source = (record.source_method or "").lower()
    if source == "user-corrected":
        priority = 5
    elif source.startswith("observed-"):
        priority = 4
    elif source.startswith("memory"):
        priority = 3
    elif source.startswith("anagram") or source == "capital-city":
        priority = 2
    elif source.startswith("openai"):
        priority = 1
    else:
        priority = 0
    return (priority, record.confidence, record.use_count, record.last_seen_at)
