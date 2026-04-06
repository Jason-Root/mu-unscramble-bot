from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from difflib import get_close_matches
import json
import re
from pathlib import Path
from typing import Protocol

from openai import OpenAI

from mu_unscramble_bot.config import BotConfig
from mu_unscramble_bot.github_answer_sheet import GitHubAnswerSheetConfig
from mu_unscramble_bot.memory_store import QuestionMemory
from mu_unscramble_bot.models import Puzzle, SolverResult, normalize_letters
from mu_unscramble_bot.paths import bundle_dir, resolve_user_path

try:
    from wordfreq import top_n_list
except Exception:  # pragma: no cover - optional dependency at runtime
    top_n_list = None


@dataclass(frozen=True, slots=True)
class ApiTestResult:
    ok: bool
    provider: str
    model: str
    reply: str = ""
    error: str = ""


class Solver(Protocol):
    name: str

    def solve(self, puzzle: Puzzle) -> SolverResult | None:
        ...


def letters_match(answer: str, scramble: str) -> bool:
    return sorted(normalize_letters(answer)) == sorted(normalize_letters(scramble))


def make_signature(value: str) -> str:
    return "".join(sorted(normalize_letters(value)))


@dataclass(slots=True)
class CapitalCitySolver:
    name: str = "capital-city"
    capitals_by_country: dict[str, str] = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self.capitals_by_country = self._load_country_capitals()
        self.capitals_by_country.update(
            {
                "usa": "washington",
                "us": "washington",
                "united states": "washington",
                "uk": "london",
                "great britain": "london",
                "britain": "london",
                "england": "london",
                "scotland": "edinburgh",
                "wales": "cardiff",
                "northern ireland": "belfast",
                "south korea": "seoul",
                "north korea": "pyongyang",
                "russia": "moscow",
                "vatican": "vaticancity",
            }
        )

    def solve(self, puzzle: Puzzle) -> SolverResult | None:
        country_name = self._extract_country_from_hint(puzzle.hint)
        if not country_name:
            return None

        capital = self._lookup_capital(country_name)
        if not capital:
            return None

        if not letters_match(capital, puzzle.normalized_scramble):
            return None

        return SolverResult(answer=capital, method=self.name, confidence=0.95)

    def _lookup_capital(self, country_name: str) -> str | None:
        normalized = self._normalize_country_name(country_name)
        if normalized in self.capitals_by_country:
            return self.capitals_by_country[normalized]

        close = get_close_matches(normalized, self.capitals_by_country.keys(), n=1, cutoff=0.87)
        if close:
            return self.capitals_by_country[close[0]]
        return None

    @staticmethod
    def _extract_country_from_hint(hint: str) -> str | None:
        patterns = [
            r"(?:what\s+is\s+)?the\s+capital(?:\s+city)?\s+of\s+(.+?)[\?!.]*$",
            r"capital(?:\s+city)?\s+of\s+(.+?)[\?!.]*$",
        ]
        for pattern in patterns:
            match = re.search(pattern, hint, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return None

    @staticmethod
    def _normalize_country_name(value: str) -> str:
        cleaned = re.sub(r"[^a-z0-9 ]+", " ", value.lower())
        return re.sub(r"\s+", " ", cleaned).strip()

    @classmethod
    def _load_country_capitals(cls) -> dict[str, str]:
        for path in cls._candidate_paths():
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                continue
            if isinstance(payload, dict):
                return {
                    cls._normalize_country_name(str(country)): normalize_letters(str(capital))
                    for country, capital in payload.items()
                    if str(country).strip() and str(capital).strip()
                }
        return {}

    @staticmethod
    def _candidate_paths() -> tuple[Path, ...]:
        return (
            resolve_user_path("data/country_capitals.json"),
            bundle_dir() / "data" / "country_capitals.json",
        )


@dataclass(slots=True)
class LocalAnagramSolver:
    max_words: int = 250000
    custom_dictionary_path: Path | None = None
    unique_only: bool = True
    seed_answers: tuple[str, ...] = ()
    extra_words: tuple[str, ...] = ()
    name: str = "anagram"
    candidates_by_signature: dict[str, tuple[str, ...]] = field(init=False, repr=False)
    base_score_by_word: dict[str, float] = field(init=False, repr=False)
    custom_words: set[str] = field(init=False, repr=False)
    seed_answer_words: set[str] = field(init=False, repr=False)
    extra_word_set: set[str] = field(init=False, repr=False)

    def __post_init__(self) -> None:
        by_signature: dict[str, set[str]] = defaultdict(set)
        self.base_score_by_word = {}
        self.custom_words = set()
        self.seed_answer_words = set()
        self.extra_word_set = set()

        if top_n_list is not None and self.max_words > 0:
            try:
                builtin_words = top_n_list("en", self.max_words)
            except LookupError:
                builtin_words = ()
            except Exception:
                builtin_words = ()

            for index, raw_word in enumerate(builtin_words):
                self._register_word(
                    by_signature,
                    raw_word,
                    base_score=max(0.0, 10.0 - (index / max(1, self.max_words / 10))),
                )

        for raw_word in self.extra_words:
            normalized = self._register_word(by_signature, raw_word, base_score=7.0)
            if normalized:
                self.extra_word_set.add(normalized)

        for raw_word in self.seed_answers:
            normalized = self._register_word(by_signature, raw_word, base_score=8.5)
            if normalized:
                self.seed_answer_words.add(normalized)

        for raw_word in self._load_custom_words():
            normalized = self._register_word(by_signature, raw_word, base_score=9.0)
            if normalized:
                self.custom_words.add(normalized)

        self.candidates_by_signature = {
            signature: tuple(sorted(words, key=lambda word: self.base_score_by_word.get(word, 0.0), reverse=True))
            for signature, words in by_signature.items()
        }

    def solve(self, puzzle: Puzzle) -> SolverResult | None:
        candidates = self.candidates_by_signature.get(make_signature(puzzle.normalized_scramble))
        if not candidates:
            return None

        if self.unique_only:
            if len(candidates) != 1:
                return None
            return SolverResult(answer=candidates[0], method=f"{self.name}:unique", confidence=0.96)

        scored = sorted(
            ((self._score_candidate(candidate, puzzle), candidate) for candidate in candidates),
            reverse=True,
        )
        best_score, best_candidate = scored[0]
        second_score = scored[1][0] if len(scored) > 1 else None

        if len(scored) == 1:
            return SolverResult(answer=best_candidate, method=self.name, confidence=0.91)

        if second_score is not None and best_score - second_score >= 2.25:
            return SolverResult(answer=best_candidate, method=f"{self.name}:ranked", confidence=0.84)

        if (
            best_candidate in self.custom_words
            or best_candidate in self.seed_answer_words
            or ("capital" in puzzle.normalized_hint and best_candidate in self.extra_word_set)
        ):
            return SolverResult(answer=best_candidate, method=f"{self.name}:seeded", confidence=0.82)

        return None

    def _register_word(
        self,
        by_signature: dict[str, set[str]],
        raw_word: str,
        *,
        base_score: float,
    ) -> str | None:
        normalized = normalize_letters(raw_word)
        if not _is_dictionary_candidate(normalized):
            return None

        signature = make_signature(normalized)
        by_signature[signature].add(normalized)
        self.base_score_by_word[normalized] = max(
            base_score,
            self.base_score_by_word.get(normalized, float("-inf")),
        )
        return normalized

    def _score_candidate(self, candidate: str, puzzle: Puzzle) -> float:
        score = self.base_score_by_word.get(candidate, 0.0)
        if candidate in self.seed_answer_words:
            score += 2.0
        if candidate in self.custom_words:
            score += 2.25
        if candidate in self.extra_word_set:
            score += 0.4
        if "capital" in puzzle.normalized_hint and candidate in self.extra_word_set:
            score += 3.0
        if candidate and candidate in puzzle.hint_lookup_key.replace(" ", ""):
            score += 1.0
        return score

    def _load_custom_words(self) -> list[str]:
        if self.custom_dictionary_path is None or not self.custom_dictionary_path.exists():
            return []

        if self.custom_dictionary_path.suffix.lower() == ".json":
            try:
                payload = json.loads(self.custom_dictionary_path.read_text(encoding="utf-8"))
            except Exception:
                return []
            if isinstance(payload, list):
                return [str(item).strip() for item in payload if str(item).strip()]
            if isinstance(payload, dict):
                return [str(key).strip() for key in payload.keys() if str(key).strip()]
            return []

        words: list[str] = []
        for raw_line in self.custom_dictionary_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            words.append(line)
        return words


@dataclass(slots=True)
class OpenAIHintSolver:
    api_key: str
    model: str
    reasoning_effort: str = ""
    request_timeout_seconds: float = 2.0
    send_hint: bool = False
    base_url: str | None = None
    http_referer: str | None = None
    app_title: str | None = None
    name: str = "openai"
    client: OpenAI = field(init=False, repr=False)

    def __post_init__(self) -> None:
        default_headers: dict[str, str] = {}
        if self.http_referer:
            default_headers["HTTP-Referer"] = self.http_referer
        if self.app_title:
            default_headers["X-Title"] = self.app_title

        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url or None,
            default_headers=default_headers or None,
        )

    def solve(self, puzzle: Puzzle) -> SolverResult | None:
        instructions = (
            "You solve one-word game anagrams. "
            "Return JSON only in the form {\"answer\":\"lowercaseletters\",\"confidence\":0.0}. "
            "The answer must use exactly the provided letters and digits once each after removing spaces and punctuation. "
            "If you are unsure, return an empty answer."
        )
        if self.send_hint:
            instructions = (
                "You solve one-word game anagram clues from arbitrary trivia hints. "
                "Use the hint to identify the target answer, then unscramble it from the provided letters. "
                "Return JSON only in the form {\"answer\":\"lowercaseletters\",\"confidence\":0.0}. "
                "The answer must use exactly the scrambled letters and digits once each after removing spaces and punctuation. "
                "If you are unsure, return an empty answer."
            )
            prompt = (
                f"Scrambled letters: {puzzle.scrambled_word}\n"
                f"Hint: {puzzle.hint}\n"
                "Return only the JSON object."
            )
        else:
            prompt = (
                "Unscramble these characters into the most likely single answer.\n"
                f"Scrambled letters: {puzzle.scrambled_word}\n"
                "Return only the JSON object."
            )
        try:
            raw_text = self._request_text(
                instructions=instructions,
                prompt=prompt,
                max_output_tokens=80,
                timeout_seconds=self.request_timeout_seconds,
            )
        except Exception:
            return None

        answer, confidence = self._parse_answer(raw_text)
        if not answer:
            return None
        return SolverResult(answer=answer, method=f"{self.name}:{self.model}", confidence=confidence)

    def startup_check(self, prompt: str, timeout_seconds: float = 20.0) -> ApiTestResult:
        try:
            reply = self._request_text(
                instructions="Answer the user's tiny connectivity check briefly and directly.",
                prompt=prompt,
                max_output_tokens=24,
                timeout_seconds=timeout_seconds,
            ).strip()
        except Exception as exc:
            return ApiTestResult(
                ok=False,
                provider=self._provider_name(),
                model=self.model,
                error=f"{type(exc).__name__}: {exc}",
            )

        if not reply:
            return ApiTestResult(
                ok=False,
                provider=self._provider_name(),
                model=self.model,
                error="Empty response from API.",
            )

        return ApiTestResult(
            ok=True,
            provider=self._provider_name(),
            model=self.model,
            reply=reply,
        )

    @staticmethod
    def _parse_answer(raw_text: str) -> tuple[str | None, float]:
        text = raw_text.strip()
        if not text:
            return None, 0.0
        try:
            payload = json.loads(text)
            answer = normalize_letters(payload.get("answer", ""))
            confidence = float(payload.get("confidence", 0.75) or 0.75)
            return answer or None, confidence
        except json.JSONDecodeError:
            answer = normalize_letters(text.splitlines()[0])
            return answer or None, 0.65

    def _provider_name(self) -> str:
        if self.base_url and "openrouter.ai" in self.base_url.lower():
            return "OpenRouter"
        return "OpenAI-compatible"

    def _request_text(
        self,
        *,
        instructions: str,
        prompt: str,
        max_output_tokens: int,
        timeout_seconds: float | None = None,
    ) -> str:
        if self._is_openrouter():
            return self._request_openrouter_text(
                instructions=instructions,
                prompt=prompt,
                max_output_tokens=max_output_tokens,
                timeout_seconds=timeout_seconds,
            )

        request = {
            "model": self.model,
            "instructions": instructions,
            "input": prompt,
            "max_output_tokens": max_output_tokens,
        }
        if self.reasoning_effort:
            request["reasoning"] = {"effort": self.reasoning_effort}
        if timeout_seconds is not None:
            request["timeout"] = timeout_seconds

        try:
            response = self.client.responses.create(**request)
            text = getattr(response, "output_text", "").strip()
            if text:
                return text
        except Exception:
            # Many local OpenAI-compatible servers expose only chat/completions.
            pass

        return self._request_chat_text(
            instructions=instructions,
            prompt=prompt,
            max_output_tokens=max_output_tokens,
            timeout_seconds=timeout_seconds,
        )

    def _is_openrouter(self) -> bool:
        return bool(self.base_url and "openrouter.ai" in self.base_url.lower())

    def _request_openrouter_text(
        self,
        *,
        instructions: str,
        prompt: str,
        max_output_tokens: int,
        timeout_seconds: float | None = None,
    ) -> str:
        token_budget = max(96, max_output_tokens)

        for _ in range(3):
            request: dict[str, object] = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": instructions},
                    {"role": "user", "content": prompt},
                ],
                "max_tokens": token_budget,
                "temperature": 0,
            }
            if timeout_seconds is not None:
                request["timeout"] = timeout_seconds

            completion = self.client.chat.completions.create(**request)
            choice = completion.choices[0] if completion.choices else None
            message = choice.message if choice else None
            content = (message.content or "").strip() if message else ""
            if content:
                return content

            token_budget *= 2

        return ""

    def _request_chat_text(
        self,
        *,
        instructions: str,
        prompt: str,
        max_output_tokens: int,
        timeout_seconds: float | None = None,
    ) -> str:
        token_budget = max(96, max_output_tokens)
        request: dict[str, object] = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": instructions},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
        }
        request["max_tokens"] = token_budget
        if timeout_seconds is not None:
            request["timeout"] = timeout_seconds

        completion = self.client.chat.completions.create(**request)
        choice = completion.choices[0] if completion.choices else None
        message = choice.message if choice else None
        content = (message.content or "").strip() if message else ""
        return content


class SolverChain:
    def __init__(
        self,
        solvers: list[Solver],
        *,
        require_letter_match: bool = True,
        question_memory: QuestionMemory | None = None,
    ) -> None:
        self.solvers = solvers
        self.require_letter_match = require_letter_match
        self.question_memory = question_memory
        self._cache: dict[str, SolverResult] = {}
        self._offline_solvers = [solver for solver in solvers if not isinstance(solver, OpenAIHintSolver)]
        self._online_solvers = [solver for solver in solvers if isinstance(solver, OpenAIHintSolver)]
        self._prefer_early_online = self._compute_prefer_early_online(solvers)

    def solve(self, puzzle: Puzzle) -> SolverResult | None:
        result = self.solve_fast(puzzle)
        if result:
            return result
        return self.solve_online(puzzle)

    def solve_fast(self, puzzle: Puzzle) -> SolverResult | None:
        cached = self._cache.get(puzzle.signature)
        if cached:
            return cached

        if self.question_memory is not None:
            memory_result = self.question_memory.lookup(puzzle)
            if memory_result:
                self._cache[puzzle.signature] = memory_result
                return memory_result

        for solver in self._offline_solvers:
            result = solver.solve(puzzle)
            if not result:
                continue
            if self.remember(puzzle, result):
                return result
        return None

    def solve_online(self, puzzle: Puzzle) -> SolverResult | None:
        cached = self._cache.get(puzzle.signature)
        if cached:
            return cached

        if self.question_memory is not None:
            memory_result = self.question_memory.lookup(puzzle)
            if memory_result:
                self._cache[puzzle.signature] = memory_result
                return memory_result

        for solver in self._online_solvers:
            result = solver.solve(puzzle)
            if not result:
                continue
            if self.remember(puzzle, result):
                return result
        return None

    def startup_check(self, prompt: str, timeout_seconds: float = 20.0) -> ApiTestResult | None:
        for solver in self._online_solvers:
            return solver.startup_check(prompt=prompt, timeout_seconds=timeout_seconds)
        return None

    def has_online_solver(self) -> bool:
        return bool(self._online_solvers)

    def prefers_early_online(self) -> bool:
        return self._prefer_early_online and bool(self._online_solvers)

    def memory_size(self) -> int:
        if self.question_memory is None:
            return 0
        return self.question_memory.size()

    def remember(self, puzzle: Puzzle, result: SolverResult) -> bool:
        if self.require_letter_match and not letters_match(result.answer, puzzle.normalized_scramble):
            return False
        if self.question_memory is not None:
            self.question_memory.remember(puzzle, result)
        self._cache[puzzle.signature] = result
        return True

    @staticmethod
    def _compute_prefer_early_online(solvers: list[Solver]) -> bool:
        first_online_index: int | None = None
        first_offline_index: int | None = None
        for index, solver in enumerate(solvers):
            if isinstance(solver, OpenAIHintSolver):
                if first_online_index is None:
                    first_online_index = index
                continue
            if first_offline_index is None:
                first_offline_index = index
        return (
            first_online_index is not None
            and first_offline_index is not None
            and first_online_index < first_offline_index
        )


def build_solver_chain(config: BotConfig) -> SolverChain:
    question_memory = None
    if config.question_memory_enabled:
        github_sync = None
        if config.github_answer_sheet_enabled and config.github_answer_sheet_repository.strip():
            github_sync = GitHubAnswerSheetConfig(
                repository=config.github_answer_sheet_repository,
                branch=config.github_answer_sheet_branch,
                path=config.github_answer_sheet_path,
                token=config.github_answer_sheet_token,
                sync_interval_seconds=config.github_answer_sheet_sync_interval_seconds,
                commit_message=config.github_answer_sheet_commit_message,
            )
        question_memory = QuestionMemory(
            path=Path(config.question_memory_path),
            fuzzy_match=config.question_memory_fuzzy_match,
            fuzzy_cutoff=config.question_memory_fuzzy_cutoff,
            github_sync=github_sync,
        )

    if config.memory_only_mode:
        return SolverChain(
            solvers=[],
            require_letter_match=config.require_letter_match,
            question_memory=question_memory,
        )

    capital_solver = CapitalCitySolver()
    available_solvers: dict[str, Solver] = {
        "capital-city": capital_solver,
    }

    if config.local_dictionary_enabled:
        available_solvers["anagram"] = LocalAnagramSolver(
            max_words=config.local_dictionary_max_words,
            custom_dictionary_path=Path(config.local_dictionary_path),
            unique_only=config.local_dictionary_unique_only,
            seed_answers=tuple(question_memory.known_answers()) if question_memory is not None else (),
            extra_words=tuple(sorted(set(capital_solver.capitals_by_country.values()))),
        )

    if config.openai_api_key:
        available_solvers["openai"] = OpenAIHintSolver(
            api_key=config.openai_api_key,
            model=config.openai_model,
            reasoning_effort=config.openai_reasoning_effort,
            request_timeout_seconds=config.online_solver_timeout_seconds,
            send_hint=config.openai_send_hint,
            base_url=config.openai_base_url,
            http_referer=config.openai_http_referer,
            app_title=config.openai_app_title,
        )

    ordered_solver_ids = list(config.solver_order)
    solvers: list[Solver] = []
    seen_solver_ids: set[str] = set()
    for solver_id in ordered_solver_ids:
        solver = available_solvers.get(solver_id)
        if solver is None or solver_id in seen_solver_ids:
            continue
        solvers.append(solver)
        seen_solver_ids.add(solver_id)

    for solver_id, solver in available_solvers.items():
        if solver_id in seen_solver_ids:
            continue
        solvers.append(solver)

    return SolverChain(
        solvers=solvers,
        require_letter_match=config.require_letter_match,
        question_memory=question_memory,
    )


def _is_dictionary_candidate(word: str) -> bool:
    if len(word) < 3 or len(word) > 18:
        return False
    return any(ch.isalpha() for ch in word)
