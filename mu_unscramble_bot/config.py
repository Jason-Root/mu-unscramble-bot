from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
from typing import Any

from dotenv import dotenv_values, load_dotenv

from mu_unscramble_bot.paths import default_config_path, default_env_path, ensure_runtime_files, resolve_user_path


DEFAULT_CONFIG_PATH = default_config_path()


@dataclass(slots=True)
class BotConfig:
    monitor_index: int = 1
    capture_source: str = "window"
    capture_width: int = 950
    capture_height: int = 280
    center_offset_x: int = 0
    center_offset_y: int = 275
    capture_interval_seconds: float = 0.08
    submission_cooldown_seconds: float = 8.0
    unsolved_retry_seconds: float = 0.25
    yellow_hsv_lower: tuple[int, int, int] = (12, 90, 90)
    yellow_hsv_upper: tuple[int, int, int] = (45, 255, 255)
    mask_dilate_iterations: int = 1
    min_ocr_confidence: float = 0.68
    target_window_title_contains: str = "Divine MU Season 21 - Powered by IGCN - Name:"
    target_window_exact_title: str = ""
    target_window_index: int = 0
    target_window_visible_only: bool = True
    focus_window_title_contains: str = "MU"
    focus_window_before_submit: bool = True
    require_window_match: bool = True
    submit_backend: str = "directinput"
    auto_submit: bool = True
    open_chat_before_submit: bool = True
    open_chat_key: str = "enter"
    submit_text_template: str = "/scramble {answer}"
    submit_key: str = "enter"
    key_hold_seconds: float = 0.05
    typing_interval_seconds: float = 0.008
    pre_submit_delay_seconds: float = 0.1
    post_submit_delay_seconds: float = 0.15
    require_letter_match: bool = True
    show_overlay: bool = True
    overlay_title: str = "MU Unscramble Bot"
    overlay_width: int = 760
    overlay_height: int = 300
    overlay_top_offset: int = 20
    overlay_left_offset: int = 0
    overlay_topmost: bool = True
    show_live_ocr_overlay: bool = True
    live_ocr_max_lines: int = 6
    test_api_on_startup: bool = True
    api_test_prompt: str = "What is 2+2? Reply with only 4."
    api_test_timeout_seconds: float = 20.0
    ocr_history_frames: int = 6
    question_memory_enabled: bool = True
    question_memory_path: str = "data/question_memory.csv"
    question_memory_fuzzy_match: bool = True
    question_memory_fuzzy_cutoff: float = 0.96
    memory_only_mode: bool = False
    local_dictionary_enabled: bool = True
    local_dictionary_max_words: int = 250000
    local_dictionary_path: str = "data/local_dictionary.txt"
    openai_model: str = "qwen/qwen3.6-plus:free"
    openai_reasoning_effort: str = ""
    online_solver_timeout_seconds: float = 2.0
    openai_send_hint: bool = False
    update_repository: str = "Jason-Root/mu-unscramble-bot"
    github_answer_sheet_enabled: bool = True
    github_answer_sheet_repository: str = "Jason-Root/mu-unscramble-bot"
    github_answer_sheet_branch: str = "main"
    github_answer_sheet_path: str = "data/question_memory.csv"
    github_answer_sheet_sync_interval_seconds: float = 30.0
    github_answer_sheet_commit_message: str = "Update community answer sheet"
    github_answer_sheet_token: str | None = None
    openai_api_key: str | None = None
    openai_base_url: str | None = None
    openai_http_referer: str | None = None
    openai_app_title: str | None = None


def _load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def load_config(path: str | Path | None = None) -> BotConfig:
    ensure_runtime_files()
    load_dotenv(dotenv_path=default_env_path(), override=True)
    config_path = resolve_user_path(path) if path else DEFAULT_CONFIG_PATH
    data = _load_json(config_path)

    if "yellow_hsv_lower" in data:
        data["yellow_hsv_lower"] = tuple(data["yellow_hsv_lower"])
    if "yellow_hsv_upper" in data:
        data["yellow_hsv_upper"] = tuple(data["yellow_hsv_upper"])

    config = BotConfig(**data)
    if config.memory_only_mode:
        config.test_api_on_startup = False
        config.openai_api_key = None

    config.openai_api_key = os.getenv("OPENAI_API_KEY", config.openai_api_key)
    config.openai_model = os.getenv("OPENAI_MODEL", config.openai_model)
    config.openai_reasoning_effort = os.getenv(
        "OPENAI_REASONING_EFFORT",
        config.openai_reasoning_effort,
    )
    config.openai_base_url = os.getenv("OPENAI_BASE_URL", config.openai_base_url)
    config.openai_http_referer = os.getenv("OPENAI_HTTP_REFERER", config.openai_http_referer)
    config.openai_app_title = os.getenv("OPENAI_APP_TITLE", config.openai_app_title)
    config.github_answer_sheet_token = os.getenv("GITHUB_TOKEN", config.github_answer_sheet_token)
    config.question_memory_path = str(resolve_user_path(config.question_memory_path))
    config.local_dictionary_path = str(resolve_user_path(config.local_dictionary_path))
    if config.memory_only_mode:
        config.test_api_on_startup = False
        config.openai_api_key = None
    return config


def save_config(config: BotConfig, path: str | Path | None = None) -> Path:
    ensure_runtime_files()
    config_path = resolve_user_path(path) if path else DEFAULT_CONFIG_PATH
    config_path.parent.mkdir(parents=True, exist_ok=True)

    data: dict[str, Any] = {
        "monitor_index": config.monitor_index,
        "capture_source": config.capture_source,
        "capture_width": config.capture_width,
        "capture_height": config.capture_height,
        "center_offset_x": config.center_offset_x,
        "center_offset_y": config.center_offset_y,
        "capture_interval_seconds": config.capture_interval_seconds,
        "submission_cooldown_seconds": config.submission_cooldown_seconds,
        "unsolved_retry_seconds": config.unsolved_retry_seconds,
        "yellow_hsv_lower": list(config.yellow_hsv_lower),
        "yellow_hsv_upper": list(config.yellow_hsv_upper),
        "mask_dilate_iterations": config.mask_dilate_iterations,
        "min_ocr_confidence": config.min_ocr_confidence,
        "target_window_title_contains": config.target_window_title_contains,
        "target_window_exact_title": config.target_window_exact_title,
        "target_window_index": config.target_window_index,
        "target_window_visible_only": config.target_window_visible_only,
        "focus_window_title_contains": config.focus_window_title_contains,
        "focus_window_before_submit": config.focus_window_before_submit,
        "require_window_match": config.require_window_match,
        "submit_backend": config.submit_backend,
        "auto_submit": config.auto_submit,
        "open_chat_before_submit": config.open_chat_before_submit,
        "open_chat_key": config.open_chat_key,
        "submit_text_template": config.submit_text_template,
        "submit_key": config.submit_key,
        "key_hold_seconds": config.key_hold_seconds,
        "typing_interval_seconds": config.typing_interval_seconds,
        "pre_submit_delay_seconds": config.pre_submit_delay_seconds,
        "post_submit_delay_seconds": config.post_submit_delay_seconds,
        "require_letter_match": config.require_letter_match,
        "show_overlay": config.show_overlay,
        "overlay_title": config.overlay_title,
        "overlay_width": config.overlay_width,
        "overlay_height": config.overlay_height,
        "overlay_top_offset": config.overlay_top_offset,
        "overlay_left_offset": config.overlay_left_offset,
        "overlay_topmost": config.overlay_topmost,
        "show_live_ocr_overlay": config.show_live_ocr_overlay,
        "live_ocr_max_lines": config.live_ocr_max_lines,
        "test_api_on_startup": config.test_api_on_startup,
        "api_test_prompt": config.api_test_prompt,
        "api_test_timeout_seconds": config.api_test_timeout_seconds,
        "ocr_history_frames": config.ocr_history_frames,
        "question_memory_enabled": config.question_memory_enabled,
        "question_memory_path": config.question_memory_path,
        "question_memory_fuzzy_match": config.question_memory_fuzzy_match,
        "question_memory_fuzzy_cutoff": config.question_memory_fuzzy_cutoff,
        "memory_only_mode": config.memory_only_mode,
        "local_dictionary_enabled": config.local_dictionary_enabled,
        "local_dictionary_max_words": config.local_dictionary_max_words,
        "local_dictionary_path": config.local_dictionary_path,
        "openai_model": config.openai_model,
        "openai_reasoning_effort": config.openai_reasoning_effort,
        "online_solver_timeout_seconds": config.online_solver_timeout_seconds,
        "openai_send_hint": config.openai_send_hint,
        "update_repository": config.update_repository,
        "github_answer_sheet_enabled": config.github_answer_sheet_enabled,
        "github_answer_sheet_repository": config.github_answer_sheet_repository,
        "github_answer_sheet_branch": config.github_answer_sheet_branch,
        "github_answer_sheet_path": config.github_answer_sheet_path,
        "github_answer_sheet_sync_interval_seconds": config.github_answer_sheet_sync_interval_seconds,
        "github_answer_sheet_commit_message": config.github_answer_sheet_commit_message,
    }
    config_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return config_path


def load_env_settings(path: str | Path | None = None) -> dict[str, str]:
    ensure_runtime_files()
    env_path = resolve_user_path(path) if path else default_env_path()
    values = dotenv_values(env_path)
    return {str(key): str(value) for key, value in values.items() if key and value is not None}


def save_env_settings(settings: dict[str, str | None], path: str | Path | None = None) -> Path:
    ensure_runtime_files()
    env_path = resolve_user_path(path) if path else default_env_path()
    current = load_env_settings(env_path)
    for key, value in settings.items():
        if value is None or value == "":
            current.pop(key, None)
            os.environ.pop(key, None)
            continue
        current[key] = value
        os.environ[key] = value

    ordered_keys = [
        "OPENAI_API_KEY",
        "OPENAI_MODEL",
        "OPENAI_REASONING_EFFORT",
        "OPENAI_BASE_URL",
        "OPENAI_HTTP_REFERER",
        "OPENAI_APP_TITLE",
        "GITHUB_TOKEN",
    ]
    lines: list[str] = []
    written: set[str] = set()
    for key in ordered_keys:
        if key not in current:
            continue
        lines.append(f"{key}={current[key]}")
        written.add(key)
    for key in sorted(current):
        if key in written:
            continue
        lines.append(f"{key}={current[key]}")
    env_path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
    return env_path
