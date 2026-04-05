from __future__ import annotations

import re

from mu_unscramble_bot.models import Puzzle, normalize_letters, normalize_spacing, sanitize_hint_text


ROUND_RE = re.compile(
    r"ROUND\s*(\d+)\s*[:;]?\s*UNSCRAMB(?:LE|IE)\s*TH(?:IS|HS)\s*WORD\s*[:;]?\s*([A-Z0-9]+)",
    re.IGNORECASE,
)
ROUND_RELAXED_RE = re.compile(
    r"ROUND\s*(\d+).*?WORD\s*[:;]?\s*([A-Z0-9]{3,})",
    re.IGNORECASE,
)
DIFFICULTY_RE = re.compile(r"DIFFICULTY\s*LEVEL\s*[:;]?\s*(\d+)", re.IGNORECASE)
HINT_RE = re.compile(r"HINT\s*[:;]?\s*(.+)", re.IGNORECASE)
GUESSED_WORD_RE = re.compile(
    r"HAS\s*SUCCESSFULLY\s*GUESSED\s*WORD\s*[:;]?\s*([A-Z0-9]+)",
    re.IGNORECASE,
)


def parse_puzzle(lines: list[str]) -> Puzzle | None:
    if not lines:
        return None

    round_number: int | None = None
    scrambled_word: str | None = None
    difficulty_level: int | None = None
    hint: str | None = None

    for line in reversed(lines):
        if round_number is None or scrambled_word is None:
            match = ROUND_RE.search(line)
            if match:
                round_number = int(match.group(1))
                scrambled_word = match.group(2).lower()

        if difficulty_level is None:
            match = DIFFICULTY_RE.search(line)
            if match:
                difficulty_level = int(match.group(1))

        if hint is None:
            hint = _extract_hint_from_lines(lines)

    joined = " ".join(lines)
    if round_number is None or scrambled_word is None:
        match = _find_last(ROUND_RE, joined)
        if match:
            round_number = int(match.group(1))
            scrambled_word = match.group(2).lower()
        else:
            match = _find_last(ROUND_RELAXED_RE, joined)
            if match:
                round_number = int(match.group(1))
                scrambled_word = match.group(2).lower()

    if difficulty_level is None:
        match = _find_last(DIFFICULTY_RE, joined)
        if match:
            difficulty_level = int(match.group(1))

    if hint is None:
        hint = _extract_hint_from_lines(lines)
    if hint is None:
        match = _find_last(HINT_RE, joined)
        if match:
            hint = _clean_hint(match.group(1))

    if round_number is None or scrambled_word is None or hint is None:
        return None

    return Puzzle(
        round_number=round_number,
        scrambled_word=scrambled_word,
        hint=hint,
        difficulty_level=difficulty_level,
        raw_lines=tuple(lines),
    )


def _clean_hint(value: str) -> str:
    hint = normalize_spacing(value)
    hint = re.split(
        (
            r"(?:ROUND\s*\d+\s*:|DIFFICULTY\s*LEVEL|"
            r"[A-Z0-9_]*\s*HAS\s*SUCCESSFULLY\s*GUESSED\s*WORD|"
            r"SCRAMBLE\s*WORDS?\s*FINISHED|\[[A-Z])"
        ),
        hint,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    return sanitize_hint_text(hint)


def _find_last(pattern: re.Pattern[str], text: str) -> re.Match[str] | None:
    matches = list(pattern.finditer(text))
    return matches[-1] if matches else None


def parse_guessed_word(lines: list[str]) -> str | None:
    for line in reversed(lines):
        match = GUESSED_WORD_RE.search(line)
        if match:
            answer = normalize_letters(match.group(1))
            if answer:
                return answer
    return None


def _extract_hint_from_lines(lines: list[str]) -> str | None:
    for index in range(len(lines) - 1, -1, -1):
        line = lines[index]
        match = HINT_RE.search(line)
        if not match:
            continue

        parts = [_clean_hint_fragment(match.group(1))]
        for next_line in lines[index + 1 : index + 4]:
            if not _is_hint_continuation(next_line):
                break
            parts.append(_clean_hint_fragment(next_line))

        hint = _clean_hint(" ".join(part for part in parts if part))
        return hint or None
    return None


def _clean_hint_fragment(value: str) -> str:
    return normalize_spacing(value).strip(" -")


def _is_hint_continuation(line: str) -> bool:
    stripped = normalize_spacing(line)
    if not stripped:
        return False
    if ROUND_RE.search(stripped) or ROUND_RELAXED_RE.search(stripped):
        return False
    if DIFFICULTY_RE.search(stripped) or HINT_RE.search(stripped) or GUESSED_WORD_RE.search(stripped):
        return False
    if "[" in stripped or "]" in stripped:
        return False
    if sum(ch.isdigit() for ch in stripped) >= 2:
        return False
    lowered = stripped.lower()
    blocked_words = (
        "killed",
        "invasion",
        "minutes",
        "zen",
        "socket",
        "gloves",
        "boots",
        "bow",
        "quest",
        "warp",
        "soldier",
        "phantom",
        "skeleton",
        "pet ",
    )
    if any(word in lowered for word in blocked_words):
        return False
    return len(normalize_letters(stripped)) >= 6
