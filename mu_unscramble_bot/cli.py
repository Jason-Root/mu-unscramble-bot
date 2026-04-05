from __future__ import annotations

import argparse
from pathlib import Path
from typing import TYPE_CHECKING

from mu_unscramble_bot.config import load_config
from mu_unscramble_bot.privilege import get_window_pid, is_current_process_elevated, is_pid_elevated
from mu_unscramble_bot.window_target import list_matching_windows

if TYPE_CHECKING:
    from mu_unscramble_bot.config import BotConfig


def main() -> int:
    parser = argparse.ArgumentParser(description="MU Online yellow-text unscramble bot")
    parser.add_argument("--config", default="config.json", help="Path to config.json")
    parser.add_argument("--window-index", type=int, default=None, help="Override target_window_index")

    subparsers = parser.add_subparsers(dest="command", required=True)

    run_parser = subparsers.add_parser("run", help="Run the live screen watcher")
    run_parser.add_argument("--dry-run", action="store_true", help="Detect and solve but do not type")

    image_parser = subparsers.add_parser("debug-image", help="Run OCR and solving against a screenshot")
    image_parser.add_argument("--image", required=True, help="Path to an image file")

    screen_parser = subparsers.add_parser("debug-screen", help="Capture the live screen once and save debug images")
    screen_parser.add_argument("--output-dir", default="debug", help="Directory for saved debug files")

    subparsers.add_parser("list-windows", help="List matching MU windows from the current config")
    subparsers.add_parser("test-api", help="Send a tiny startup test prompt to the configured API/model")
    submit_parser = subparsers.add_parser("test-submit", help="Type a test /scramble command into the target window")
    submit_parser.add_argument("--answer", default="test", help="Answer word to send with the submit template")

    args = parser.parse_args()
    config = load_config(args.config)
    if args.window_index is not None:
        config.target_window_index = args.window_index

    if args.command == "run":
        from mu_unscramble_bot.bot import MuUnscrambleBot

        bot = MuUnscrambleBot(config=config, dry_run=args.dry_run)
        bot.run_forever()
        return 0

    if args.command == "debug-image":
        return debug_image(config=config, image_path=args.image)

    if args.command == "debug-screen":
        return debug_screen(config=config, output_dir=args.output_dir)

    if args.command == "list-windows":
        return list_windows(config=config)

    if args.command == "test-api":
        return test_api(config=config)

    if args.command == "test-submit":
        return test_submit(config=config, answer=args.answer)

    return 1


def debug_image(config: BotConfig, image_path: str) -> int:
    from mu_unscramble_bot.parser import parse_puzzle
    from mu_unscramble_bot.screen_reader import YellowTextReader
    from mu_unscramble_bot.solver import build_solver_chain

    reader = YellowTextReader(config)
    try:
        capture = reader.read_from_image(image_path)
    finally:
        reader.close()

    _print_capture(capture.lines)
    puzzle = parse_puzzle(capture.lines)
    if not puzzle:
        print("Puzzle parse: not found")
        return 1

    print(f"Puzzle parse: round={puzzle.round_number} scramble={puzzle.scrambled_word} hint={puzzle.hint}")
    solver = build_solver_chain(config)
    solution = solver.solve(puzzle)
    if solution:
        print(f"Suggested answer: {solution.normalized_answer} ({solution.method})")
    else:
        print("Suggested answer: none")
    return 0


def debug_screen(config: BotConfig, output_dir: str) -> int:
    import cv2

    from mu_unscramble_bot.parser import parse_puzzle
    from mu_unscramble_bot.screen_reader import YellowTextReader

    reader = YellowTextReader(config)
    try:
        capture = reader.read_from_screen()
    finally:
        reader.close()

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(output_path / "capture.png"), capture.frame)
    cv2.imwrite(str(output_path / "mask.png"), capture.mask)
    cv2.imwrite(str(output_path / "raw_variant.png"), capture.variants["raw"])
    cv2.imwrite(str(output_path / "yellow_only_variant.png"), capture.variants["yellow_only"])
    cv2.imwrite(str(output_path / "contrast_variant.png"), capture.variants["contrast"])

    print(f"Saved debug images to: {output_path.resolve()}")
    _print_capture(capture.lines)
    puzzle = parse_puzzle(capture.lines)
    if puzzle:
        print(f"Puzzle parse: round={puzzle.round_number} scramble={puzzle.scrambled_word} hint={puzzle.hint}")
    else:
        print("Puzzle parse: not found")
    return 0


def test_api(config: BotConfig) -> int:
    from mu_unscramble_bot.solver import build_solver_chain

    solver = build_solver_chain(config)
    result = solver.startup_check(
        prompt=config.api_test_prompt,
        timeout_seconds=config.api_test_timeout_seconds,
    )
    if result is None:
        print("API test skipped: no API key / online solver configured.")
        return 1

    if result.ok:
        print(f"API test passed: provider={result.provider} model={result.model} reply={result.reply}")
        return 0

    print(f"API test failed: provider={result.provider} model={result.model} error={result.error}")
    return 1


def list_windows(config: BotConfig) -> int:
    matches = list_matching_windows(config)
    if not matches:
        print("No matching windows found for the current capture filter.")
        return 1

    print("Matching windows used by capture:")
    for match in matches:
        minimized = " minimized" if match.is_minimized else ""
        safe_title = match.title.encode("cp1252", errors="replace").decode("cp1252")
        print(
            f"  index={match.match_index} pos=({match.left},{match.top}) "
            f"size={match.width}x{match.height}{minimized} title={safe_title}"
        )
    return 0


def test_submit(config: BotConfig, answer: str) -> int:
    from mu_unscramble_bot.submitter import AnswerSubmitter
    from mu_unscramble_bot.window_target import get_target_window

    try:
        target = get_target_window(config)
    except Exception:
        target = None

    if target is not None:
        target_pid = get_window_pid(target.window)
        current_elevated = is_current_process_elevated()
        target_elevated = is_pid_elevated(target_pid) if target_pid is not None else None
        if current_elevated is False and target_elevated is True:
            print("Test submit blocked: target MU client is running as administrator. Start the bot elevated.")
            return 1

    submitter = AnswerSubmitter(config)
    sent = submitter.submit(answer)
    if sent:
        print(f"Submitted test command using answer={answer}")
        return 0
    print("Test submit failed: target window could not be focused.")
    return 1


def _print_capture(lines: list[str]) -> None:
    print("OCR lines:")
    if not lines:
        print("  (none)")
        return
    for line in lines:
        safe_line = line.encode("cp1252", errors="replace").decode("cp1252")
        print(f"  {safe_line}")
