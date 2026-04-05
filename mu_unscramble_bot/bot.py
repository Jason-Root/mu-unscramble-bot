from __future__ import annotations

from collections import deque
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
import time
from threading import Event
from typing import Callable

from mu_unscramble_bot.config import BotConfig
from mu_unscramble_bot.models import Puzzle, SolverResult
from mu_unscramble_bot.parser import parse_guessed_word, parse_puzzle
from mu_unscramble_bot.overlay import OverlayPayload, StatusOverlay
from mu_unscramble_bot.privilege import get_window_pid, is_current_process_elevated, is_pid_elevated
from mu_unscramble_bot.screen_reader import YellowTextReader
from mu_unscramble_bot.solver import SolverChain, build_solver_chain
from mu_unscramble_bot.submitter import AnswerSubmitter
from mu_unscramble_bot.window_target import WindowSelectionError, get_target_window


@dataclass(slots=True)
class PendingOnlineSolve:
    puzzle: Puzzle
    future: Future[SolverResult | None]
    started_at: float


class MuUnscrambleBot:
    def __init__(
        self,
        config: BotConfig,
        dry_run: bool = False,
        *,
        status_callback: Callable[[OverlayPayload], None] | None = None,
        log_callback: Callable[[str], None] | None = None,
    ) -> None:
        self.config = config
        self.dry_run = dry_run
        self.status_callback = status_callback
        self.log_callback = log_callback
        self.reader = YellowTextReader(config)
        self.solver: SolverChain = build_solver_chain(config)
        self.submitter = AnswerSubmitter(config)
        self.overlay = StatusOverlay(config)
        self._last_solved_at: dict[str, float] = {}
        self._last_failed_at: dict[str, float] = {}
        self._last_idle_overlay_at = 0.0
        self._last_window_error_at = 0.0
        self._last_detected_puzzle: Puzzle | None = None
        self._last_detected_puzzle_at = 0.0
        self._last_round_activity_at = 0.0
        self._last_active_round_number: int | None = None
        self._last_observed_answer = ""
        self._recent_ocr_lines: deque[list[str]] = deque(maxlen=max(1, config.ocr_history_frames))
        self._completed_rounds: dict[str, float] = {}
        self._submitted_answers_by_round: dict[str, dict[str, float]] = {}
        self._online_executor: ThreadPoolExecutor | None = None
        if self.solver.has_online_solver():
            self._online_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="mu-online-solver")
        self._pending_online_solve: PendingOnlineSolve | None = None
        self._stop_requested = Event()

    def close(self) -> None:
        if self._pending_online_solve is not None:
            self._pending_online_solve.future.cancel()
            self._pending_online_solve = None
        if self._online_executor is not None:
            self._online_executor.shutdown(wait=False, cancel_futures=True)
        self.overlay.close()
        self.reader.close()

    def request_stop(self) -> None:
        self._stop_requested.set()

    def run_forever(self) -> None:
        self._run_startup_checks()
        self._warn_if_submit_is_blocked_by_elevation()
        self._prime_reader()
        self._log_selected_window()
        memory_count = self.solver.memory_size()
        if memory_count:
            self._log(f"Loaded {memory_count} saved question/answer rows from the spreadsheet cache.")
        self._log("Watching screen for yellow unscramble text. Press Ctrl+C to stop.")
        self._publish_status(
            status="Watching the center yellow text...",
            round_text="-",
            scramble_text="-",
            hint_text="-",
            answer_text="-",
            method_text="waiting",
            ocr_text="-",
        )
        try:
            while not self._stop_requested.is_set():
                self.run_once()
                time.sleep(self._current_capture_interval_seconds())
        finally:
            self.close()

    def run_once(self) -> tuple[Puzzle | None, SolverResult | None]:
        cycle_started_at = time.perf_counter()
        try:
            capture = self.reader.read_from_screen()
        except WindowSelectionError as exc:
            self._update_window_error(str(exc))
            return None, None

        self._recent_ocr_lines.append(capture.lines)
        merged_lines = self._merged_recent_lines()
        live_ocr_text = self._format_live_ocr_lines(capture.lines)
        observed_answer = parse_guessed_word(capture.lines) or parse_guessed_word(merged_lines)
        puzzle = parse_puzzle(capture.lines)
        if not puzzle:
            if merged_lines != capture.lines:
                puzzle = parse_puzzle(merged_lines)
        if puzzle:
            self._last_detected_puzzle = puzzle
            self._last_detected_puzzle_at = time.monotonic()
            self._last_round_activity_at = self._last_detected_puzzle_at
            self._last_active_round_number = puzzle.round_number

        if observed_answer:
            self._learn_from_observed_answer(observed_answer)
            self._mark_last_detected_round_completed()

        if not puzzle:
            pending_result = self._consume_pending_online_result_without_visible_puzzle(
                live_ocr_text=live_ocr_text,
                cycle_started_at=cycle_started_at,
            )
            if pending_result is not None:
                return pending_result
            self._update_idle_overlay(live_ocr_text, api_status=self._api_status_text())
            return None, None

        self._prune_completed_rounds(now=time.monotonic())
        self._prune_submitted_answers(now=time.monotonic())
        if self._is_round_completed(puzzle):
            self._cancel_pending_online_if_matches(puzzle.round_key)
            self._update_overlay_from_puzzle(
                puzzle,
                status="Round already solved. Waiting for the next round...",
                answer_text="-",
                method_text="round complete",
                ocr_text=live_ocr_text,
            )
            return puzzle, None

        now = time.monotonic()
        last_solved = self._last_solved_at.get(puzzle.signature, 0.0)
        if now - last_solved < self.config.submission_cooldown_seconds:
            self._update_overlay_from_puzzle(
                puzzle,
                status="Puzzle seen recently. Waiting for the next round...",
                answer_text="-",
                method_text="cooldown",
                ocr_text=live_ocr_text,
            )
            return puzzle, None

        last_failed = self._last_failed_at.get(puzzle.signature, 0.0)
        if now - last_failed < self.config.unsolved_retry_seconds:
            self._update_overlay_from_puzzle(
                puzzle,
                status="Retrying this puzzle...",
                answer_text="-",
                method_text="retrying",
                ocr_text=live_ocr_text,
            )
            return puzzle, None

        self._update_overlay_from_puzzle(
            puzzle,
            status="Puzzle detected. Solving...",
            answer_text="-",
            method_text="solving",
            ocr_text=live_ocr_text,
        )

        if observed_answer and self.solver.remember(
            puzzle,
            SolverResult(answer=observed_answer, method="observed-guess", confidence=1.0),
        ):
            self._last_failed_at.pop(puzzle.signature, None)
            self._last_solved_at[puzzle.signature] = now
            self._update_overlay_from_puzzle(
                puzzle,
                status="Someone else solved it. Learned the answer for next time.",
                answer_text=observed_answer,
                method_text="observed-guess",
                ocr_text=live_ocr_text,
            )
            self._log(
                f"Detected round {puzzle.round_number} -> {observed_answer} via observed guessed-word line "
                f"in {time.perf_counter() - cycle_started_at:.2f}s"
            )
            return puzzle, SolverResult(answer=observed_answer, method="observed-guess", confidence=1.0)

        solution = self.solver.solve_fast(puzzle)
        consumed_online_result = False
        if not solution and self._pending_online_solve is not None:
            if self._pending_online_solve.puzzle.round_key != puzzle.round_key:
                self._cancel_pending_online("New round detected before API finished.")
            elif self._pending_online_solve.future.done():
                consumed_online_result = True
                solution = self._consume_pending_online_result(puzzle)
                if solution is None and self._is_round_completed(puzzle):
                    self._update_overlay_from_puzzle(
                        puzzle,
                        status="Round already solved. Waiting for the next round...",
                        answer_text="-",
                        method_text="round complete",
                        ocr_text=live_ocr_text,
                    )
                    return puzzle, None
            else:
                self._update_overlay_from_puzzle(
                    puzzle,
                    status="Watching OCR while API solves in background...",
                    answer_text="-",
                    method_text="api pending",
                    ocr_text=live_ocr_text,
                )
                return puzzle, None

        if (
            not solution
            and not consumed_online_result
            and self._pending_online_solve is None
            and self._should_start_online_solve()
        ):
            self._start_online_solve(puzzle)
            self._update_overlay_from_puzzle(
                puzzle,
                status="Watching OCR while API solves in background...",
                answer_text="-",
                method_text="api pending",
                ocr_text=live_ocr_text,
            )
            return puzzle, None

        if not solution:
            self._last_failed_at[puzzle.signature] = now
            solver_state = "memory only" if self.config.memory_only_mode else "no answer yet"
            self._update_overlay_from_puzzle(
                puzzle,
                status="Could not solve this hint yet.",
                answer_text="-",
                method_text=solver_state,
                ocr_text=live_ocr_text,
            )
            self._log(
                f"Detected round {puzzle.round_number} but could not solve it yet. "
                f"Scramble={puzzle.scrambled_word} Hint={puzzle.hint}"
            )
            return puzzle, None

        return self._finalize_solution(
            puzzle,
            solution,
            live_ocr_text=live_ocr_text,
            cycle_started_at=cycle_started_at,
        )

    def _finalize_solution(
        self,
        puzzle: Puzzle,
        solution: SolverResult,
        *,
        live_ocr_text: str,
        cycle_started_at: float,
    ) -> tuple[Puzzle, SolverResult]:
        self._last_failed_at.pop(puzzle.signature, None)
        self._last_solved_at[puzzle.signature] = time.monotonic()
        self._update_overlay_from_puzzle(
            puzzle,
            status="Answer found.",
            answer_text=solution.normalized_answer,
            method_text=solution.method,
            ocr_text=live_ocr_text,
        )
        self._log(
            f"Detected round {puzzle.round_number} -> {solution.normalized_answer} "
            f"via {solution.method} in {time.perf_counter() - cycle_started_at:.2f}s"
        )

        if self.dry_run or not self.config.auto_submit:
            self._update_overlay_from_puzzle(
                puzzle,
                status="Dry run: answer shown only.",
                answer_text=solution.normalized_answer,
                method_text=solution.method,
                ocr_text=live_ocr_text,
            )
            return puzzle, solution

        if self._has_submitted_answer(puzzle, solution.normalized_answer):
            self._update_overlay_from_puzzle(
                puzzle,
                status="This answer was already submitted for the current round.",
                answer_text=solution.normalized_answer,
                method_text="already submitted",
                ocr_text=live_ocr_text,
            )
            self._log(
                f"Skipped duplicate submit for round {puzzle.round_number}: {solution.normalized_answer}"
            )
            return puzzle, solution

        self._update_overlay_from_puzzle(
            puzzle,
            status="Submitting answer...",
            answer_text=solution.normalized_answer,
            method_text=solution.method,
            ocr_text=live_ocr_text,
        )
        if self.submitter.submit(solution.normalized_answer):
            self._mark_answer_submitted(puzzle, solution.normalized_answer)
            self._mark_round_completed(puzzle)
            self._update_overlay_from_puzzle(
                puzzle,
                status="Answer submitted.",
                answer_text=solution.normalized_answer,
                method_text=solution.method,
                ocr_text=live_ocr_text,
            )
            self._log(f"Submitted answer: {solution.normalized_answer}")
        else:
            self._update_overlay_from_puzzle(
                puzzle,
                status="Submit failed. MU did not accept the input.",
                answer_text=solution.normalized_answer,
                method_text=solution.method,
                ocr_text=live_ocr_text,
            )
            self._log("Answer found, but MU did not accept the submit input.")
        return puzzle, solution

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}"
        print(line)
        if self.log_callback is not None:
            self.log_callback(line)

    def _publish_status(
        self,
        *,
        status: str,
        round_text: str = "-",
        scramble_text: str = "-",
        hint_text: str = "-",
        answer_text: str = "-",
        method_text: str = "-",
        ocr_text: str = "-",
    ) -> None:
        payload = OverlayPayload(
            status=status,
            round_text=round_text,
            scramble_text=scramble_text,
            hint_text=hint_text,
            answer_text=answer_text,
            method_text=method_text,
            ocr_text=ocr_text,
        )
        self.overlay.update(
            status=payload.status,
            round_text=payload.round_text,
            scramble_text=payload.scramble_text,
            hint_text=payload.hint_text,
            answer_text=payload.answer_text,
            method_text=payload.method_text,
            ocr_text=payload.ocr_text,
        )
        if self.status_callback is not None:
            self.status_callback(payload)

    def _update_idle_overlay(self, ocr_text: str, api_status: str = "idle") -> None:
        now = time.monotonic()
        if now - self._last_idle_overlay_at < 0.2:
            return
        self._last_idle_overlay_at = now
        self._publish_status(
            status="Waiting for the yellow puzzle text...",
            round_text="-",
            scramble_text="-",
            hint_text="-",
            answer_text="-",
            method_text=api_status,
            ocr_text=ocr_text,
        )

    def _update_overlay_from_puzzle(
        self,
        puzzle: Puzzle,
        *,
        status: str,
        answer_text: str,
        method_text: str,
        ocr_text: str,
    ) -> None:
        self._publish_status(
            status=status,
            round_text=str(puzzle.round_number),
            scramble_text=puzzle.scrambled_word,
            hint_text=puzzle.hint,
            answer_text=answer_text,
            method_text=method_text,
            ocr_text=ocr_text,
        )

    def _run_startup_checks(self) -> None:
        if self.config.memory_only_mode:
            self._log("Memory-only mode is enabled. Unknown questions will be skipped until learned.")
            self._publish_status(
                status="Memory-only mode enabled.",
                round_text="-",
                scramble_text="-",
                hint_text="Unknown questions will be learned from guessed-word lines.",
                answer_text="-",
                method_text="memory only",
                ocr_text="-",
            )
            return

        if not self.config.openai_api_key:
            self._log("API startup check skipped: no API key configured.")
            self._publish_status(
                status="API disabled. Only offline hint solving is available.",
                round_text="-",
                scramble_text="-",
                hint_text="-",
                answer_text="-",
                method_text="no api key",
                ocr_text="-",
            )
            return

        if not self.config.test_api_on_startup:
            self._log("API startup check skipped by config.")
            return

        self._log(f"Running API startup check using {self.config.openai_model}...")
        self._publish_status(
            status="Testing API connection...",
            round_text="-",
            scramble_text="-",
            hint_text=self.config.api_test_prompt,
            answer_text="-",
            method_text=self.config.openai_model,
            ocr_text="-",
        )

        result = self.solver.startup_check(
            prompt=self.config.api_test_prompt,
            timeout_seconds=self.config.api_test_timeout_seconds,
        )
        if result is None:
            self._log("API startup check skipped: no online solver is configured.")
            return

        if result.ok:
            self._log(
                f"API startup check passed on {result.provider} using {result.model}. "
                f"Reply: {result.reply}"
            )
            self._publish_status(
                status="API test passed. Watching the screen next...",
                round_text="-",
                scramble_text="-",
                hint_text=self.config.api_test_prompt,
                answer_text=result.reply,
                method_text=f"{result.provider} / {result.model}",
                ocr_text="-",
            )
            return

        self._log(
            f"API startup check failed on {result.provider} using {result.model}. "
            f"Error: {result.error}"
        )
        self._publish_status(
            status="API test failed. OCR will still run.",
            round_text="-",
            scramble_text="-",
            hint_text=self.config.api_test_prompt,
            answer_text="-",
            method_text="api test failed",
            ocr_text="-",
        )

    def _warn_if_submit_is_blocked_by_elevation(self) -> None:
        if self.dry_run or not self.config.auto_submit:
            return
        if self.config.capture_source.lower() != "window":
            return

        try:
            target = get_target_window(self.config)
        except WindowSelectionError:
            return

        target_pid = get_window_pid(target.window)
        current_elevated = is_current_process_elevated()
        target_elevated = is_pid_elevated(target_pid) if target_pid is not None else None
        if current_elevated is False and target_elevated is True:
            message = "Target MU client is running as administrator. Start the bot elevated for auto-submit."
            self._log(message)
            self._publish_status(
                status="Submit blocked by Windows privilege mismatch.",
                round_text="-",
                scramble_text="-",
                hint_text=message,
                answer_text="-",
                method_text="run as admin",
                ocr_text="-",
            )

    def _prime_reader(self) -> None:
        self._log("Priming OCR so the first real puzzle is faster...")
        try:
            for _ in range(3):
                self.reader.read_from_screen()
        except WindowSelectionError as exc:
            self._log(f"OCR warm-up skipped: {exc}")

    def _merged_recent_lines(self) -> list[str]:
        merged: list[str] = []
        for lines in self._recent_ocr_lines:
            for line in lines:
                merged.append(line)
        return merged[-40:]

    def _log_selected_window(self) -> None:
        if self.config.capture_source.lower() != "window":
            self._log(f"Capture source: monitor {self.config.monitor_index}")
            return

        try:
            target = get_target_window(self.config)
        except WindowSelectionError as exc:
            self._log(f"Window selection issue: {exc}")
            return

        self._log(
            f"Capture source: window #{target.match_index} "
            f"'{target.title}' at ({target.left},{target.top}) {target.width}x{target.height}"
        )

    def _update_window_error(self, message: str) -> None:
        now = time.monotonic()
        if now - self._last_window_error_at < 2.0:
            return
        self._last_window_error_at = now
        self._log(f"Window selection issue: {message}")
        self._publish_status(
            status="Target MU window not found.",
            round_text="-",
            scramble_text="-",
            hint_text=message,
            answer_text="-",
            method_text="window selection",
            ocr_text="-",
        )

    def _learn_from_observed_answer(self, observed_answer: str) -> None:
        if observed_answer == self._last_observed_answer:
            return
        self._last_observed_answer = observed_answer

        if self._last_detected_puzzle is None:
            return
        if time.monotonic() - self._last_detected_puzzle_at > 45.0:
            return
        if not self.solver.remember(
            self._last_detected_puzzle,
            SolverResult(answer=observed_answer, method="observed-guess", confidence=1.0),
        ):
            return

        self._log(
            f"Learned answer from guessed-word line for round {self._last_detected_puzzle.round_number}: "
            f"{observed_answer}"
        )
        self._last_round_activity_at = time.monotonic()
        self._last_active_round_number = self._last_detected_puzzle.round_number

    def _mark_last_detected_round_completed(self) -> None:
        if self._last_detected_puzzle is None:
            return
        if time.monotonic() - self._last_detected_puzzle_at > 45.0:
            return
        self._mark_round_completed(self._last_detected_puzzle)

    def _mark_round_completed(self, puzzle: Puzzle) -> None:
        self._completed_rounds[puzzle.round_key] = time.monotonic()
        self._last_round_activity_at = time.monotonic()
        self._last_active_round_number = puzzle.round_number
        self._cancel_pending_online_if_matches(puzzle.round_key)

    def _is_round_completed(self, puzzle: Puzzle) -> bool:
        completed_at = self._completed_rounds.get(puzzle.round_key)
        if completed_at is None:
            return False
        return (time.monotonic() - completed_at) < max(15.0, self.config.submission_cooldown_seconds)

    def _prune_completed_rounds(self, *, now: float) -> None:
        cutoff = max(15.0, self.config.submission_cooldown_seconds)
        stale_keys = [key for key, completed_at in self._completed_rounds.items() if (now - completed_at) >= cutoff]
        for key in stale_keys:
            self._completed_rounds.pop(key, None)

    def _mark_answer_submitted(self, puzzle: Puzzle, answer: str) -> None:
        answer_key = answer.strip().lower()
        if not answer_key:
            return
        submitted = self._submitted_answers_by_round.setdefault(puzzle.round_key, {})
        submitted[answer_key] = time.monotonic()

    def _has_submitted_answer(self, puzzle: Puzzle, answer: str) -> bool:
        answer_key = answer.strip().lower()
        if not answer_key:
            return False
        submitted = self._submitted_answers_by_round.get(puzzle.round_key, {})
        return answer_key in submitted

    def _prune_submitted_answers(self, *, now: float) -> None:
        cutoff = max(45.0, self.config.submission_cooldown_seconds * 4)
        stale_rounds: list[str] = []
        for round_key, answers in self._submitted_answers_by_round.items():
            stale_answers = [answer for answer, submitted_at in answers.items() if (now - submitted_at) >= cutoff]
            for answer in stale_answers:
                answers.pop(answer, None)
            if not answers:
                stale_rounds.append(round_key)
        for round_key in stale_rounds:
            self._submitted_answers_by_round.pop(round_key, None)

    def _should_start_online_solve(self) -> bool:
        return (
            not self.config.memory_only_mode
            and self.config.openai_api_key is not None
            and self._online_executor is not None
        )

    def _start_online_solve(self, puzzle: Puzzle) -> None:
        if self._online_executor is None:
            return
        self._last_round_activity_at = time.monotonic()
        self._last_active_round_number = puzzle.round_number
        future = self._online_executor.submit(self.solver.solve_online, puzzle)
        self._pending_online_solve = PendingOnlineSolve(
            puzzle=puzzle,
            future=future,
            started_at=time.monotonic(),
        )

    def _consume_pending_online_result(self, current_puzzle: Puzzle) -> SolverResult | None:
        pending = self._pending_online_solve
        self._pending_online_solve = None
        if pending is None:
            return None
        if pending.puzzle.round_key != current_puzzle.round_key:
            return None
        if self._is_round_completed(current_puzzle):
            return None
        try:
            return pending.future.result()
        except Exception as exc:
            self._log(f"Online solver failed: {type(exc).__name__}: {exc}")
            return None

    def _cancel_pending_online_if_matches(self, round_key: str) -> None:
        pending = self._pending_online_solve
        if pending is None:
            return
        if pending.puzzle.round_key != round_key:
            return
        self._cancel_pending_online("Round completed before API result was needed.")

    def _cancel_pending_online(self, reason: str) -> None:
        pending = self._pending_online_solve
        if pending is None:
            return
        pending.future.cancel()
        self._pending_online_solve = None
        self._log(reason)

    def _api_status_text(self) -> str:
        pending = self._pending_online_solve
        if pending is None:
            return "idle"
        if pending.future.done():
            return "api ready"
        elapsed = time.monotonic() - pending.started_at
        return f"api pending {elapsed:.1f}s"

    def _format_live_ocr_lines(self, lines: list[str]) -> str:
        if not self.config.show_live_ocr_overlay:
            return "-"
        if not lines:
            return "-"

        cleaned: list[str] = []
        for line in lines[: max(1, self.config.live_ocr_max_lines)]:
            text = " ".join(line.split())
            if len(text) > 90:
                text = text[:87] + "..."
            cleaned.append(text)
        return "\n".join(cleaned) if cleaned else "-"

    def _consume_pending_online_result_without_visible_puzzle(
        self,
        *,
        live_ocr_text: str,
        cycle_started_at: float,
    ) -> tuple[Puzzle, SolverResult] | None:
        pending = self._pending_online_solve
        if pending is None or not pending.future.done():
            return None

        if self._is_round_completed(pending.puzzle):
            self._pending_online_solve = None
            return None

        elapsed = time.monotonic() - pending.started_at
        if elapsed > self.config.pending_api_submit_grace_seconds:
            self._cancel_pending_online("Discarded late API result after the submit grace window expired.")
            return None

        solution = self._consume_pending_online_result(pending.puzzle)
        if solution is None:
            self._last_failed_at[pending.puzzle.signature] = time.monotonic()
            return None

        self._log("API result arrived after OCR lost the puzzle text. Submitting from the pending solve.")
        return self._finalize_solution(
            pending.puzzle,
            solution,
            live_ocr_text=live_ocr_text,
            cycle_started_at=cycle_started_at,
        )

    def _current_capture_interval_seconds(self) -> float:
        if self._is_active_round_window():
            return max(0.01, self.config.active_capture_interval_seconds)
        return max(0.01, self.config.idle_capture_interval_seconds)

    def _is_active_round_window(self) -> bool:
        if self._pending_online_solve is not None:
            return True

        round_number = self._last_active_round_number
        if round_number is None:
            return False
        if round_number < 1 or round_number > max(1, self.config.active_round_count):
            return False
        return (time.monotonic() - self._last_round_activity_at) <= max(1.0, self.config.active_round_linger_seconds)
