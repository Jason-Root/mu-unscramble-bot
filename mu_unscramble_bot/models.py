from __future__ import annotations

from dataclasses import dataclass, field
import re


HINT_NOISE_PATTERNS = (
    r"\b[a-z0-9_]*\s*has\s*successfully\s*guessed\s*word\s*:?\s*[a-z0-9]+\b.*$",
    r"\bscramble\s*words?\s*finished\b.*$",
    r"\bbalgass\s*has\s*join\w*.*$",
    r"\bcrywolf\s*fortress\b.*$",
)


def normalize_letters(value: str) -> str:
    return "".join(ch for ch in value.lower() if ch.isalnum())


def normalize_spacing(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def sanitize_hint_text(value: str) -> str:
    cleaned = normalize_spacing(value)
    for pattern in HINT_NOISE_PATTERNS:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)
        cleaned = normalize_spacing(cleaned)
    return cleaned.rstrip(" -:;,.?!")


def normalize_lookup_text(value: str) -> str:
    lowered = sanitize_hint_text(value).lower()
    return re.sub(r"[^a-z0-9]+", "", lowered)


@dataclass(frozen=True, slots=True)
class Puzzle:
    round_number: int
    scrambled_word: str
    hint: str
    difficulty_level: int | None = None
    raw_lines: tuple[str, ...] = field(default_factory=tuple)

    @property
    def normalized_scramble(self) -> str:
        return normalize_letters(self.scrambled_word)

    @property
    def normalized_hint(self) -> str:
        return sanitize_hint_text(self.hint).lower()

    @property
    def hint_lookup_key(self) -> str:
        return normalize_lookup_text(self.hint)

    @property
    def signature(self) -> str:
        return f"{self.round_number}|{self.normalized_scramble}|{self.normalized_hint}"

    @property
    def round_key(self) -> str:
        return f"{self.round_number}|{self.normalized_scramble}"


@dataclass(frozen=True, slots=True)
class SolverResult:
    answer: str
    method: str
    confidence: float = 0.0

    @property
    def normalized_answer(self) -> str:
        return normalize_letters(self.answer)
