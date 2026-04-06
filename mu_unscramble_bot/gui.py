from __future__ import annotations

from dataclasses import dataclass
import ctypes
import json
import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import font as tkfont
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urlparse
import urllib.error
import urllib.request

from mu_unscramble_bot.bot import MuUnscrambleBot
from mu_unscramble_bot.config import (
    BotConfig,
    DEFAULT_SOLVER_ORDER,
    load_config,
    load_env_settings,
    save_config,
    save_env_settings,
)
from mu_unscramble_bot.overlay import OverlayPayload
from mu_unscramble_bot.paths import APP_NAME, is_frozen, user_data_dir
from mu_unscramble_bot.privilege import is_current_process_elevated
from mu_unscramble_bot.updater import (
    UpdateCheckResult,
    check_for_updates,
    download_release_asset,
    get_app_version,
    open_release_page,
    prepare_file_update,
    stage_windows_file_update,
    stage_windows_update,
)
from mu_unscramble_bot.window_target import extract_character_name, list_matching_windows


WINDOW_BG = "#0b1620"
CARD_BG = "#12222d"
CARD_BORDER = "#244252"
TEXT_MAIN = "#eef4f8"
TEXT_SOFT = "#8ca5b5"
GREEN = "#2ea043"
BLUE = "#1f6feb"
RED = "#b94b4b"

PROVIDER_DISABLED = "Disabled"
PROVIDER_OPENROUTER = "OpenRouter"
PROVIDER_LOCAL = "Local OpenAI-Compatible"
PROVIDER_CUSTOM = "Custom OpenAI-Compatible"
SOLVER_LABELS = {
    "capital-city": "Capital-City Clues",
    "anagram": "Dictionary Anagram",
    "openai": "AI Fallback",
}


@dataclass(slots=True)
class ClientChoice:
    label: str
    match_index: int
    character_name: str
    title: str


@dataclass(slots=True)
class UpdateMessage:
    result: UpdateCheckResult
    silent_if_current: bool = False


def _speed_to_values(speed: int) -> tuple[float, float]:
    clamped = max(1, min(10, speed))
    ratio = (clamped - 1) / 9
    typing_interval = 0.03 - (0.028 * ratio)
    key_hold = 0.08 - (0.05 * ratio)
    return round(typing_interval, 4), round(key_hold, 4)


def _values_to_speed(typing_interval_seconds: float) -> int:
    slow = 0.03
    fast = 0.002
    if typing_interval_seconds <= fast:
        return 10
    if typing_interval_seconds >= slow:
        return 1
    ratio = (slow - typing_interval_seconds) / (slow - fast)
    return max(1, min(10, int(round(1 + ratio * 9))))


def _seconds_to_ms_text(value: float) -> str:
    return str(int(round(max(0.0, value) * 1000)))


def _ms_text_to_seconds(value: str, *, fallback: float) -> float:
    try:
        milliseconds = float(value.strip())
    except Exception:
        return fallback
    return max(0.01, milliseconds / 1000.0)


def _detect_provider(config: BotConfig) -> str:
    base_url = (config.openai_base_url or "").strip().lower()
    api_key = (config.openai_api_key or "").strip()
    if not base_url and not api_key:
        return PROVIDER_DISABLED
    if "openrouter.ai" in base_url:
        return PROVIDER_OPENROUTER
    if _is_local_base_url(base_url):
        return PROVIDER_LOCAL
    if base_url:
        return PROVIDER_CUSTOM
    return PROVIDER_DISABLED


def _is_local_base_url(base_url: str) -> bool:
    if not base_url:
        return False
    try:
        hostname = (urlparse(base_url).hostname or "").lower()
    except Exception:
        hostname = ""
    if not hostname:
        return False
    if hostname in {"localhost", "::1"} or hostname.startswith("127."):
        return True
    if hostname.startswith("192.168.") or hostname.startswith("10."):
        return True
    if hostname.startswith("172."):
        parts = hostname.split(".")
        if len(parts) >= 2:
            try:
                second = int(parts[1])
            except Exception:
                second = -1
            if 16 <= second <= 31:
                return True
    return False


def _normalize_provider_base_url(provider: str, base_url: str) -> str:
    cleaned = base_url.strip().rstrip("/")
    if provider != PROVIDER_LOCAL:
        return cleaned
    if not cleaned:
        return "http://127.0.0.1:11434"
    return cleaned


def _fetch_model_candidates(base_url: str, *, api_key: str = "") -> list[str]:
    normalized = base_url.strip().rstrip("/")
    if not normalized:
        raise ValueError("Base URL is empty.")

    parsed = urlparse(normalized)
    path = (parsed.path or "").rstrip("/")
    base_root = normalized[:-3] if path.endswith("/v1") else normalized
    base_root = base_root.rstrip("/")

    candidate_urls: list[str] = []
    if path.endswith("/v1"):
        candidate_urls.append(f"{normalized}/models")
    else:
        candidate_urls.append(f"{normalized}/v1/models")
    candidate_urls.append(f"{base_root}/api/tags")

    headers = {"User-Agent": "mu-unscramble-bot"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    errors: list[str] = []
    seen_models: dict[str, None] = {}
    for url in dict.fromkeys(candidate_urls):
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(request, timeout=8.0) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            errors.append(f"{url} -> HTTP {exc.code}")
            continue
        except Exception as exc:
            errors.append(f"{url} -> {type(exc).__name__}: {exc}")
            continue

        for model in _extract_model_ids(payload):
            seen_models[str(model)] = None
        if seen_models:
            return sorted(seen_models.keys())

    error_text = errors[0] if errors else "No model endpoints returned data."
    raise RuntimeError(error_text)


def _extract_model_ids(payload: object) -> list[str]:
    models: list[str] = []
    if isinstance(payload, dict):
        if isinstance(payload.get("data"), list):
            for item in payload["data"]:
                if not isinstance(item, dict):
                    continue
                model_id = str(item.get("id") or item.get("model") or item.get("name") or "").strip()
                if model_id:
                    models.append(model_id)
        if isinstance(payload.get("models"), list):
            for item in payload["models"]:
                if not isinstance(item, dict):
                    continue
                model_id = str(item.get("name") or item.get("model") or item.get("id") or "").strip()
                if model_id:
                    models.append(model_id)
    return models


def _run_connection_test(*, provider: str, base_url: str, api_key: str, model: str) -> str:
    from mu_unscramble_bot.solver import OpenAIHintSolver

    timeout_seconds = 35.0 if provider == PROVIDER_LOCAL else 10.0
    solver = OpenAIHintSolver(
        api_key=api_key or ("local" if provider == PROVIDER_LOCAL else ""),
        model=model,
        base_url=base_url,
        http_referer="http://localhost" if provider == PROVIDER_OPENROUTER else None,
        app_title=APP_NAME,
        request_timeout_seconds=timeout_seconds,
    )
    result = solver.startup_check(prompt="What is 2+2? Reply with only 4.", timeout_seconds=timeout_seconds)
    if not result.ok:
        raise RuntimeError(result.error or "The provider returned an empty response.")
    return f"Connected to {result.provider} using {result.model}. Reply: {result.reply}"


class SettingsDialog:
    def __init__(self, parent: "DesktopApp") -> None:
        self.parent = parent
        self.config = load_config()
        self.env_settings = load_env_settings()

        self.window = tk.Toplevel(parent.root)
        self.window.title(f"{APP_NAME} Settings")
        self.window.configure(bg=WINDOW_BG)
        self.window.resizable(False, False)
        self.window.transient(parent.root)
        self.window.grab_set()

        self.provider_var = tk.StringVar(value=_detect_provider(self.config))
        self.api_key_var = tk.StringVar(value=self.env_settings.get("OPENAI_API_KEY", ""))
        self.model_var = tk.StringVar(value=self.env_settings.get("OPENAI_MODEL", self.config.openai_model))
        self.base_url_var = tk.StringVar(value=self.env_settings.get("OPENAI_BASE_URL", self.config.openai_base_url or ""))
        self.community_sync_var = tk.BooleanVar(value=self.config.github_answer_sheet_enabled)
        self.github_token_var = tk.StringVar(value=self.env_settings.get("GITHUB_TOKEN", ""))
        self.speed_var = tk.IntVar(value=_values_to_speed(self.config.typing_interval_seconds))
        self.command_word_var = tk.StringVar(value=self.config.submit_command_word)
        self.active_capture_ms_var = tk.StringVar(value=_seconds_to_ms_text(self.config.active_capture_interval_seconds))
        self.idle_capture_ms_var = tk.StringVar(value=_seconds_to_ms_text(self.config.idle_capture_interval_seconds))
        self.active_round_count_var = tk.StringVar(value=str(self.config.active_round_count))
        self.active_round_linger_var = tk.StringVar(value=str(self.config.active_round_linger_seconds))
        self.dictionary_enabled_var = tk.BooleanVar(value=self.config.local_dictionary_enabled)
        self.dictionary_unique_only_var = tk.BooleanVar(value=self.config.local_dictionary_unique_only)
        self.dictionary_path_var = tk.StringVar(value=self.config.local_dictionary_path)
        self.send_hint_var = tk.BooleanVar(value=self.config.openai_send_hint)
        self.speed_text_var = tk.StringVar()
        self.detect_models_status_var = tk.StringVar(value="")
        self.test_connection_status_var = tk.StringVar(value="")
        self.solver_order = list(self.config.solver_order)

        outer = tk.Frame(self.window, bg=WINDOW_BG, padx=18, pady=18)
        outer.pack(fill="both", expand=True)

        scroll_shell = tk.Frame(outer, bg=WINDOW_BG)
        scroll_shell.pack(fill="both", expand=True)
        self.settings_canvas = tk.Canvas(
            scroll_shell,
            bg=WINDOW_BG,
            highlightthickness=0,
            bd=0,
        )
        settings_scrollbar = ttk.Scrollbar(
            scroll_shell,
            orient="vertical",
            command=self.settings_canvas.yview,
        )
        self.settings_canvas.configure(yscrollcommand=settings_scrollbar.set)
        self.settings_canvas.pack(side="left", fill="both", expand=True)
        settings_scrollbar.pack(side="right", fill="y")

        container = tk.Frame(self.settings_canvas, bg=WINDOW_BG)
        self._settings_canvas_window = self.settings_canvas.create_window((0, 0), window=container, anchor="nw")
        container.bind("<Configure>", self._sync_settings_scrollregion)
        self.settings_canvas.bind("<Configure>", self._sync_settings_canvas_width)
        self.window.bind("<MouseWheel>", self._on_settings_mousewheel)

        tk.Label(
            container,
            text="Settings",
            bg=WINDOW_BG,
            fg=TEXT_MAIN,
            font=("Segoe UI Semibold", 18),
        ).pack(anchor="w")
        tk.Label(
            container,
            text="Each PC can keep its own API setup and typing speed.",
            bg=WINDOW_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        api_card = self._card(container)
        api_card.pack(fill="x", pady=(16, 0))
        self._row_label(api_card, "API Provider").pack(anchor="w")

        provider_combo = ttk.Combobox(
            api_card,
            textvariable=self.provider_var,
            state="readonly",
            values=[PROVIDER_DISABLED, PROVIDER_OPENROUTER, PROVIDER_LOCAL, PROVIDER_CUSTOM],
            font=("Segoe UI", 10),
            width=30,
        )
        provider_combo.pack(fill="x", pady=(8, 0))
        provider_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_provider_preset())

        self._row_label(api_card, "API Key").pack(anchor="w", pady=(12, 0))
        tk.Entry(
            api_card,
            textvariable=self.api_key_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            show="*",
        ).pack(fill="x", pady=(8, 0))

        self._row_label(api_card, "Model").pack(anchor="w", pady=(12, 0))
        model_row = tk.Frame(api_card, bg=CARD_BG)
        model_row.pack(fill="x", pady=(8, 0))
        tk.Entry(
            model_row,
            textvariable=self.model_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True)
        self.detect_models_button = tk.Button(
            model_row,
            text="Detect Models",
            command=self._detect_models,
            bg="#345369",
            fg=TEXT_MAIN,
            relief="flat",
            padx=12,
            pady=6,
            font=("Segoe UI Semibold", 9),
        )
        self.detect_models_button.pack(side="left", padx=(8, 0))
        self.test_connection_button = tk.Button(
            model_row,
            text="Test Connection",
            command=self._test_connection,
            bg="#2d5f41",
            fg=TEXT_MAIN,
            relief="flat",
            padx=12,
            pady=6,
            font=("Segoe UI Semibold", 9),
        )
        self.test_connection_button.pack(side="left", padx=(8, 0))
        tk.Label(
            api_card,
            textvariable=self.detect_models_status_var,
            bg=CARD_BG,
            fg="#79c0ff",
            font=("Segoe UI", 9),
            wraplength=640,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))
        tk.Label(
            api_card,
            textvariable=self.test_connection_status_var,
            bg=CARD_BG,
            fg="#7ee787",
            font=("Segoe UI", 9),
            wraplength=640,
            justify="left",
        ).pack(anchor="w", pady=(2, 0))

        self._row_label(api_card, "Base URL").pack(anchor="w", pady=(12, 0))
        tk.Entry(
            api_card,
            textvariable=self.base_url_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(fill="x", pady=(8, 0))
        tk.Checkbutton(
            api_card,
            text="Send the full hint/question to AI instead of only the scramble letters",
            variable=self.send_hint_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor="#0b1620",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(12, 0))
        tk.Label(
            api_card,
            text="For Ollama over LAN, enter the server root URL like http://192.168.1.42:11434 . The app will try both OpenAI-style and native Ollama endpoints.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
            wraplength=640,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        solver_card = self._card(container)
        solver_card.pack(fill="x", pady=(16, 0))
        self._row_label(solver_card, "Solver Order").pack(anchor="w")
        tk.Label(
            solver_card,
            text="Memory answers stay first. Move the local solver steps up or down to experiment. If AI is first, it starts early in the background while the fast local solvers keep running.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
            wraplength=640,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))
        self.solver_order_frame = tk.Frame(solver_card, bg=CARD_BG)
        self.solver_order_frame.pack(fill="x", pady=(10, 0))
        self._render_solver_order_rows()

        speed_card = self._card(container)
        speed_card.pack(fill="x", pady=(16, 0))
        self._row_label(speed_card, "Typing Speed").pack(anchor="w")
        tk.Label(
            speed_card,
            text="1 is safest and slowest. 10 is fastest.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))

        speed_scale = tk.Scale(
            speed_card,
            from_=1,
            to=10,
            orient="horizontal",
            variable=self.speed_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            troughcolor="#244252",
            highlightthickness=0,
            activebackground=BLUE,
            command=lambda _value: self._update_speed_text(),
        )
        speed_scale.pack(fill="x", pady=(10, 0))
        tk.Label(
            speed_card,
            textvariable=self.speed_text_var,
            bg=CARD_BG,
            fg="#7ee787",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(6, 0))

        submit_card = self._card(container)
        submit_card.pack(fill="x", pady=(16, 0))
        self._row_label(submit_card, "Submission Command").pack(anchor="w")
        tk.Label(
            submit_card,
            text="Default is scramble. The app sends /<word> <answer>.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))
        tk.Entry(
            submit_card,
            textvariable=self.command_word_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(fill="x", pady=(8, 0))

        ocr_card = self._card(container)
        ocr_card.pack(fill="x", pady=(16, 0))
        self._row_label(ocr_card, "OCR Pace").pack(anchor="w")
        tk.Label(
            ocr_card,
            text="Fast scan runs during round chains. Idle scan runs between games.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))

        ocr_grid = tk.Frame(ocr_card, bg=CARD_BG)
        ocr_grid.pack(fill="x", pady=(10, 0))
        self._row_label(ocr_grid, "Round OCR (ms)").grid(row=0, column=0, sticky="w")
        tk.Entry(
            ocr_grid,
            textvariable=self.active_capture_ms_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            width=10,
        ).grid(row=1, column=0, sticky="we", padx=(0, 10), pady=(6, 0))
        self._row_label(ocr_grid, "Idle OCR (ms)").grid(row=0, column=1, sticky="w")
        tk.Entry(
            ocr_grid,
            textvariable=self.idle_capture_ms_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            width=10,
        ).grid(row=1, column=1, sticky="we", padx=(0, 10), pady=(6, 0))
        self._row_label(ocr_grid, "Rounds Per Event").grid(row=0, column=2, sticky="w")
        tk.Entry(
            ocr_grid,
            textvariable=self.active_round_count_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            width=10,
        ).grid(row=1, column=2, sticky="we", padx=(0, 10), pady=(6, 0))
        self._row_label(ocr_grid, "Fast Mode Hold (s)").grid(row=0, column=3, sticky="w")
        tk.Entry(
            ocr_grid,
            textvariable=self.active_round_linger_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            width=10,
        ).grid(row=1, column=3, sticky="we", pady=(6, 0))

        dictionary_card = self._card(container)
        dictionary_card.pack(fill="x", pady=(16, 0))
        self._row_label(dictionary_card, "Dictionary Attack").pack(anchor="w")
        tk.Checkbutton(
            dictionary_card,
            text="Enable dictionary anagram solver",
            variable=self.dictionary_enabled_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor="#0b1620",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(8, 0))
        tk.Checkbutton(
            dictionary_card,
            text="Only submit if there is exactly one possible word",
            variable=self.dictionary_unique_only_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor="#0b1620",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(6, 0))
        tk.Label(
            dictionary_card,
            text="Custom dictionary path. Supports .txt and .json files.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(10, 0))
        dictionary_row = tk.Frame(dictionary_card, bg=CARD_BG)
        dictionary_row.pack(fill="x", pady=(8, 0))
        tk.Entry(
            dictionary_row,
            textvariable=self.dictionary_path_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True)
        tk.Button(
            dictionary_row,
            text="Browse",
            command=self._browse_dictionary,
            bg="#2d5f41",
            fg=TEXT_MAIN,
            relief="flat",
            padx=12,
            pady=6,
            font=("Segoe UI Semibold", 9),
        ).pack(side="left", padx=(8, 0))

        github_card = self._card(container)
        github_card.pack(fill="x", pady=(16, 0))
        self._row_label(github_card, "Community Answer Sync").pack(anchor="w")
        tk.Label(
            github_card,
            text="Answers can sync through GitHub so every PC can pull new solves.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))
        tk.Label(
            github_card,
            text=f"Repository: {self.config.github_answer_sheet_repository}",
            bg=CARD_BG,
            fg="#79c0ff",
            font=("Consolas", 9),
        ).pack(anchor="w", pady=(6, 0))
        tk.Checkbutton(
            github_card,
            text="Enable GitHub community sync",
            variable=self.community_sync_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            activebackground=CARD_BG,
            activeforeground=TEXT_MAIN,
            selectcolor="#0b1620",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(10, 0))
        self._row_label(github_card, "GitHub Token").pack(anchor="w", pady=(12, 0))
        tk.Entry(
            github_card,
            textvariable=self.github_token_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            show="*",
        ).pack(fill="x", pady=(8, 0))
        tk.Label(
            github_card,
            text="Leave blank for read-only sync. Add a token to publish new answers back to GitHub.",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(6, 0))

        actions = tk.Frame(outer, bg=WINDOW_BG)
        actions.pack(fill="x", pady=(18, 0))
        tk.Button(
            actions,
            text="Cancel",
            command=self.window.destroy,
            bg="#2d3f4a",
            fg=TEXT_MAIN,
            relief="flat",
            padx=14,
            pady=8,
            font=("Segoe UI Semibold", 10),
        ).pack(side="right")
        tk.Button(
            actions,
            text="Save",
            command=self._save,
            bg=GREEN,
            fg=TEXT_MAIN,
            relief="flat",
            padx=16,
            pady=8,
            font=("Segoe UI Semibold", 10),
        ).pack(side="right", padx=(0, 8))

        self._apply_provider_preset(initial=True)
        self._update_speed_text()
        self.window.update_idletasks()
        width = max(760, self.parent.root.winfo_width())
        height = min(640, max(520, self.window.winfo_height()))
        x = parent.root.winfo_rootx() + max(30, (parent.root.winfo_width() - width) // 2)
        y = parent.root.winfo_rooty() + 40
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.settings_canvas.yview_moveto(0.0)

    def _card(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(parent, bg=CARD_BG, highlightbackground=CARD_BORDER, highlightthickness=1, padx=14, pady=14)

    def _row_label(self, parent: tk.Misc, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold"))

    def _sync_settings_scrollregion(self, _event: tk.Event | None = None) -> None:
        self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox("all"))

    def _sync_settings_canvas_width(self, event: tk.Event) -> None:
        self.settings_canvas.itemconfigure(self._settings_canvas_window, width=event.width)

    def _on_settings_mousewheel(self, event: tk.Event) -> str | None:
        if event.delta == 0:
            return None
        self.settings_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _apply_provider_preset(self, initial: bool = False) -> None:
        provider = self.provider_var.get()
        if provider == PROVIDER_OPENROUTER:
            if not self.base_url_var.get().strip():
                self.base_url_var.set("https://openrouter.ai/api/v1")
            if not self.model_var.get().strip():
                self.model_var.set("qwen/qwen3.6-plus:free")
        elif provider == PROVIDER_LOCAL:
            if not self.base_url_var.get().strip():
                self.base_url_var.set("http://127.0.0.1:11434")
            if not self.model_var.get().strip():
                self.model_var.set("llama3.1")
            if initial and not self.api_key_var.get().strip():
                self.api_key_var.set("local")
        elif provider == PROVIDER_DISABLED:
            if initial:
                return
            self.api_key_var.set("")
            self.model_var.set("")
            self.base_url_var.set("")

    def _update_speed_text(self) -> None:
        typing_interval, key_hold = _speed_to_values(self.speed_var.get())
        self.speed_text_var.set(
            f"Current delay: {int(round(typing_interval * 1000))} ms between keys, "
            f"{int(round(key_hold * 1000))} ms key hold"
        )

    def _browse_dictionary(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.window,
            title="Choose dictionary file",
            filetypes=[
                ("Dictionary files", "*.txt *.json"),
                ("Text files", "*.txt"),
                ("JSON files", "*.json"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.dictionary_path_var.set(path)

    def _detect_models(self) -> None:
        provider = self.provider_var.get()
        base_url = _normalize_provider_base_url(provider, self.base_url_var.get())
        if not base_url:
            messagebox.showwarning(APP_NAME, "Enter a base URL before detecting models.")
            return

        self.base_url_var.set(base_url)
        self.detect_models_status_var.set("Detecting models from the configured endpoint...")
        self.detect_models_button.config(state="disabled", text="Detecting...")
        api_key = self.api_key_var.get().strip()
        threading.Thread(
            target=self._detect_models_worker,
            args=(base_url, api_key),
            name="mu-detect-models",
            daemon=True,
        ).start()

    def _test_connection(self) -> None:
        provider = self.provider_var.get()
        if provider == PROVIDER_DISABLED:
            messagebox.showwarning(APP_NAME, "Choose an API provider first.")
            return

        model = self.model_var.get().strip()
        base_url = _normalize_provider_base_url(provider, self.base_url_var.get())
        api_key = self.api_key_var.get().strip()
        if not model:
            messagebox.showwarning(APP_NAME, "Enter a model name before testing.")
            return
        if not base_url:
            messagebox.showwarning(APP_NAME, "Enter a base URL before testing.")
            return

        self.base_url_var.set(base_url)
        self.test_connection_status_var.set("Testing API connection...")
        self.test_connection_button.config(state="disabled", text="Testing...")
        threading.Thread(
            target=self._test_connection_worker,
            args=(provider, base_url, api_key, model),
            name="mu-test-connection",
            daemon=True,
        ).start()

    def _detect_models_worker(self, base_url: str, api_key: str) -> None:
        try:
            models = _fetch_model_candidates(base_url, api_key=api_key)
        except Exception as exc:
            self.window.after(0, lambda: self._finish_detect_models(error=f"{type(exc).__name__}: {exc}"))
            return
        self.window.after(0, lambda: self._finish_detect_models(models=models))

    def _finish_detect_models(self, *, models: list[str] | None = None, error: str | None = None) -> None:
        self.detect_models_button.config(state="normal", text="Detect Models")
        if error:
            self.detect_models_status_var.set(f"Model detect failed: {error}")
            messagebox.showerror(APP_NAME, f"Could not detect models.\n\n{error}")
            return

        models = sorted(dict.fromkeys(models or []))
        if not models:
            self.detect_models_status_var.set("No models were returned by the endpoint.")
            messagebox.showinfo(APP_NAME, "No models were returned by the endpoint.")
            return

        if len(models) == 1:
            self.model_var.set(models[0])
            self.detect_models_status_var.set(f"Detected model: {models[0]}")
            return

        self.detect_models_status_var.set(f"Detected {len(models)} models. Choose one.")
        self._open_model_picker(models)

    def _test_connection_worker(self, provider: str, base_url: str, api_key: str, model: str) -> None:
        try:
            result_text = _run_connection_test(
                provider=provider,
                base_url=base_url,
                api_key=api_key,
                model=model,
            )
        except Exception as exc:
            self.window.after(0, lambda: self._finish_connection_test(error=f"{type(exc).__name__}: {exc}"))
            return
        self.window.after(0, lambda: self._finish_connection_test(result=result_text))

    def _finish_connection_test(self, *, result: str | None = None, error: str | None = None) -> None:
        self.test_connection_button.config(state="normal", text="Test Connection")
        if error:
            self.test_connection_status_var.set(f"Connection test failed: {error}")
            messagebox.showerror(APP_NAME, f"Connection test failed.\n\n{error}")
            return
        self.test_connection_status_var.set(result or "Connection test passed.")

    def _open_model_picker(self, models: list[str]) -> None:
        dialog = tk.Toplevel(self.window)
        dialog.title("Choose Model")
        dialog.configure(bg=WINDOW_BG)
        dialog.transient(self.window)
        dialog.grab_set()
        dialog.geometry("520x360")

        container = tk.Frame(dialog, bg=WINDOW_BG, padx=16, pady=16)
        container.pack(fill="both", expand=True)
        tk.Label(
            container,
            text="Detected Models",
            bg=WINDOW_BG,
            fg=TEXT_MAIN,
            font=("Segoe UI Semibold", 14),
        ).pack(anchor="w")
        tk.Label(
            container,
            text="Choose the exact model id to use for the local AI endpoint.",
            bg=WINDOW_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))

        list_frame = tk.Frame(container, bg=WINDOW_BG)
        list_frame.pack(fill="both", expand=True, pady=(12, 0))
        listbox = tk.Listbox(
            list_frame,
            bg="#0b1620",
            fg=TEXT_MAIN,
            selectbackground=BLUE,
            selectforeground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 10),
            activestyle="none",
            exportselection=False,
        )
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        for model in models:
            listbox.insert("end", model)
        listbox.selection_set(0)
        listbox.activate(0)

        def choose_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                return
            model = str(listbox.get(selection[0]))
            self.model_var.set(model)
            self.detect_models_status_var.set(f"Selected detected model: {model}")
            dialog.destroy()

        actions = tk.Frame(container, bg=WINDOW_BG)
        actions.pack(fill="x", pady=(12, 0))
        tk.Button(
            actions,
            text="Cancel",
            command=dialog.destroy,
            bg="#2d3f4a",
            fg=TEXT_MAIN,
            relief="flat",
            padx=12,
            pady=6,
            font=("Segoe UI Semibold", 9),
        ).pack(side="right")
        tk.Button(
            actions,
            text="Use Selected",
            command=choose_selected,
            bg=GREEN,
            fg=TEXT_MAIN,
            relief="flat",
            padx=12,
            pady=6,
            font=("Segoe UI Semibold", 9),
        ).pack(side="right", padx=(0, 8))

        listbox.bind("<Double-1>", lambda _event: choose_selected())

    def _render_solver_order_rows(self) -> None:
        for child in self.solver_order_frame.winfo_children():
            child.destroy()

        for index, solver_id in enumerate(self.solver_order):
            row = tk.Frame(self.solver_order_frame, bg="#0b1620", padx=10, pady=8)
            row.pack(fill="x", pady=(0, 6))
            tk.Label(
                row,
                text=f"{index + 1}.",
                bg="#0b1620",
                fg="#79c0ff",
                font=("Segoe UI Semibold", 10),
                width=3,
                anchor="w",
            ).pack(side="left")
            tk.Label(
                row,
                text=SOLVER_LABELS.get(solver_id, solver_id),
                bg="#0b1620",
                fg=TEXT_MAIN,
                font=("Segoe UI", 10),
                anchor="w",
            ).pack(side="left", fill="x", expand=True)
            tk.Button(
                row,
                text="Up",
                command=lambda idx=index: self._move_solver_row(idx, -1),
                state="normal" if index > 0 else "disabled",
                bg="#2d5f41",
                fg=TEXT_MAIN,
                relief="flat",
                padx=10,
                pady=4,
                font=("Segoe UI Semibold", 9),
            ).pack(side="left", padx=(8, 4))
            tk.Button(
                row,
                text="Down",
                command=lambda idx=index: self._move_solver_row(idx, 1),
                state="normal" if index < len(self.solver_order) - 1 else "disabled",
                bg="#345369",
                fg=TEXT_MAIN,
                relief="flat",
                padx=10,
                pady=4,
                font=("Segoe UI Semibold", 9),
            ).pack(side="left")

    def _move_solver_row(self, index: int, delta: int) -> None:
        new_index = index + delta
        if index < 0 or new_index < 0 or index >= len(self.solver_order) or new_index >= len(self.solver_order):
            return
        self.solver_order[index], self.solver_order[new_index] = self.solver_order[new_index], self.solver_order[index]
        self._render_solver_order_rows()

    def _save(self) -> None:
        provider = self.provider_var.get()
        model = self.model_var.get().strip()
        base_url = _normalize_provider_base_url(provider, self.base_url_var.get())
        api_key = self.api_key_var.get().strip()
        command_word = self.command_word_var.get().strip().lstrip("/")

        if provider != PROVIDER_DISABLED and not model:
            messagebox.showwarning(APP_NAME, "Enter a model name before saving.")
            return
        if provider in {PROVIDER_LOCAL, PROVIDER_CUSTOM, PROVIDER_OPENROUTER} and not base_url and provider != PROVIDER_DISABLED:
            messagebox.showwarning(APP_NAME, "Enter a base URL before saving.")
            return
        if not command_word:
            messagebox.showwarning(APP_NAME, "Enter the trigger command word, like scramble or answer.")
            return

        config = load_config()
        config.show_overlay = False
        config.test_api_on_startup = False
        config.typing_interval_seconds, config.key_hold_seconds = _speed_to_values(self.speed_var.get())
        config.active_capture_interval_seconds = _ms_text_to_seconds(
            self.active_capture_ms_var.get(),
            fallback=config.active_capture_interval_seconds,
        )
        config.idle_capture_interval_seconds = _ms_text_to_seconds(
            self.idle_capture_ms_var.get(),
            fallback=config.idle_capture_interval_seconds,
        )
        config.capture_interval_seconds = config.active_capture_interval_seconds
        try:
            config.active_round_count = max(1, int(float(self.active_round_count_var.get().strip())))
        except Exception:
            config.active_round_count = max(1, config.active_round_count)
        try:
            config.active_round_linger_seconds = max(1.0, float(self.active_round_linger_var.get().strip()))
        except Exception:
            config.active_round_linger_seconds = max(1.0, config.active_round_linger_seconds)
        config.github_answer_sheet_enabled = self.community_sync_var.get()
        config.submit_command_word = command_word
        config.submit_text_template = f"/{command_word} {{answer}}"
        config.local_dictionary_enabled = self.dictionary_enabled_var.get()
        config.local_dictionary_unique_only = self.dictionary_unique_only_var.get()
        config.local_dictionary_path = self.dictionary_path_var.get().strip() or config.local_dictionary_path
        config.openai_send_hint = self.send_hint_var.get()
        config.solver_order = list(self.solver_order or DEFAULT_SOLVER_ORDER)
        save_config(config)

        if provider == PROVIDER_DISABLED:
            env_updates = {
                "OPENAI_API_KEY": "",
                "OPENAI_MODEL": "",
                "OPENAI_BASE_URL": "",
                "OPENAI_HTTP_REFERER": "",
                "OPENAI_APP_TITLE": "",
            }
        elif provider == PROVIDER_OPENROUTER:
            env_updates = {
                "OPENAI_API_KEY": api_key,
                "OPENAI_MODEL": model or "qwen/qwen3.6-plus:free",
                "OPENAI_BASE_URL": "https://openrouter.ai/api/v1",
                "OPENAI_HTTP_REFERER": "http://localhost",
                "OPENAI_APP_TITLE": APP_NAME,
            }
        elif provider == PROVIDER_LOCAL:
            env_updates = {
                "OPENAI_API_KEY": api_key or "local",
                "OPENAI_MODEL": model,
                "OPENAI_BASE_URL": base_url or "http://127.0.0.1:11434/v1",
                "OPENAI_HTTP_REFERER": "",
                "OPENAI_APP_TITLE": APP_NAME,
            }
        else:
            env_updates = {
                "OPENAI_API_KEY": api_key,
                "OPENAI_MODEL": model,
                "OPENAI_BASE_URL": base_url,
                "OPENAI_HTTP_REFERER": "",
                "OPENAI_APP_TITLE": APP_NAME,
            }

        env_updates["GITHUB_TOKEN"] = self.github_token_var.get().strip()
        save_env_settings(env_updates)
        self.parent.reload_settings_badges()
        self.parent.append_log("Settings saved.")
        self.window.destroy()
        messagebox.showinfo(APP_NAME, "Settings saved. Restart the watcher if it is already running.")


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("760x320")
        self.root.minsize(680, 300)
        self.root.configure(bg=WINDOW_BG)

        self._message_queue: queue.SimpleQueue[tuple[str, object]] = queue.SimpleQueue()
        self._bot_thread: threading.Thread | None = None
        self._bot: MuUnscrambleBot | None = None
        self._launch_config: BotConfig | None = None
        self._pending_stop = threading.Event()
        self._client_choices: list[ClientChoice] = []
        self._details_visible = False
        self._log_lines: list[str] = []

        self.selected_client = tk.StringVar()
        self.status_var = tk.StringVar(value="Choose a character, then go live.")
        self.character_var = tk.StringVar(value="No character selected")
        self.answer_var = tk.StringVar(value="-")
        self.hint_var = tk.StringVar(value="-")
        self.round_var = tk.StringVar(value="-")
        self.details_button_var = tk.StringVar(value="Show Details")
        self.badges_var = tk.StringVar(value="")
        self.version_var = tk.StringVar(value=f"v{get_app_version()}")

        self._build_styles()
        self._build_ui()
        self._refresh_clients(initial=True)
        self.reload_settings_badges()
        self.root.after(120, self._pump_messages)
        if is_frozen():
            self.root.after(1600, lambda: self._start_update_check(silent_if_current=True))
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        body_font = tkfont.nametofont("TkDefaultFont").copy()
        body_font.configure(family="Segoe UI", size=10)
        heading_font = tkfont.nametofont("TkHeadingFont").copy()
        heading_font.configure(family="Segoe UI", size=10)
        row_height = max(26, body_font.metrics("linespace") + 8)
        style.configure(
            "Bot.TCombobox",
            fieldbackground=CARD_BG,
            background=CARD_BG,
            foreground=TEXT_MAIN,
            arrowcolor="#7ee787",
            bordercolor=CARD_BORDER,
            lightcolor=CARD_BORDER,
            darkcolor=CARD_BORDER,
            relief="flat",
        )
        style.configure(
            "Dup.Treeview",
            font=body_font,
            rowheight=row_height,
            fieldbackground="#f5f7fa",
            background="#f5f7fa",
            foreground="#16202a",
            bordercolor=CARD_BORDER,
            lightcolor=CARD_BORDER,
            darkcolor=CARD_BORDER,
        )
        style.map(
            "Dup.Treeview",
            background=[("selected", "#7ea6c9")],
            foreground=[("selected", "#0b1620")],
        )
        style.configure(
            "Dup.Treeview.Heading",
            font=heading_font,
            background="#dfe6ee",
            foreground="#0b1620",
            relief="flat",
        )

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, bg=WINDOW_BG, padx=18, pady=16)
        header.pack(fill="x")

        left = tk.Frame(header, bg=WINDOW_BG)
        left.pack(side="left", fill="x", expand=True)
        tk.Label(left, text=APP_NAME, bg=WINDOW_BG, fg=TEXT_MAIN, font=("Segoe UI Semibold", 20)).pack(anchor="w")
        tk.Label(
            left,
            text="Live mode only. Pick a character and let it watch.",
            bg=WINDOW_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 0))

        right = tk.Frame(header, bg=WINDOW_BG)
        right.pack(side="right")
        tk.Label(
            right,
            textvariable=self.version_var,
            bg="#0f2230",
            fg="#79c0ff",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=6,
        ).pack(anchor="e")
        button_row = tk.Frame(right, bg=WINDOW_BG)
        button_row.pack(anchor="e", pady=(8, 0))
        self._make_button(button_row, "Settings", self._open_settings, accent=BLUE).pack(side="left")
        self._make_button(button_row, "Updates", self._start_update_check, accent="#2d5f41").pack(side="left", padx=(8, 0))

        scroll_shell = tk.Frame(self.root, bg=WINDOW_BG)
        scroll_shell.pack(fill="both", expand=True)

        self.main_canvas = tk.Canvas(
            scroll_shell,
            bg=WINDOW_BG,
            highlightthickness=0,
            bd=0,
        )
        self.main_scrollbar = ttk.Scrollbar(scroll_shell, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.main_scrollbar.pack(side="right", fill="y")

        body = tk.Frame(self.main_canvas, bg=WINDOW_BG, padx=18, pady=0)
        self._main_canvas_window = self.main_canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", self._sync_main_scrollregion)
        self.main_canvas.bind("<Configure>", self._sync_main_canvas_width)
        self.root.bind("<MouseWheel>", self._on_main_mousewheel)

        control_card = self._card(body)
        control_card.pack(fill="x")
        tk.Label(control_card, text="Character", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(anchor="w")

        row = tk.Frame(control_card, bg=CARD_BG)
        row.pack(fill="x", pady=(10, 0))

        self.client_combo = ttk.Combobox(
            row,
            textvariable=self.selected_client,
            state="readonly",
            style="Bot.TCombobox",
            font=("Segoe UI", 11),
        )
        self.client_combo.pack(side="left", fill="x", expand=True)
        self.client_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_client_selected())

        self._make_button(row, "Refresh", self._refresh_clients, accent="#345369").pack(side="left", padx=(10, 0))
        self.start_button = self._make_button(row, "Go Live", self._start_bot, accent=GREEN)
        self.start_button.pack(side="left", padx=(10, 0))
        self.stop_button = self._make_button(row, "Stop", self._stop_bot, accent=RED)
        self.stop_button.pack(side="left", padx=(10, 0))
        self.stop_button.config(state="disabled")

        tk.Label(
            control_card,
            textvariable=self.badges_var,
            bg=CARD_BG,
            fg="#7ee787",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(12, 0))

        status_card = self._card(body)
        status_card.pack(fill="x", pady=(14, 0))
        tk.Label(status_card, textvariable=self.status_var, bg=CARD_BG, fg=TEXT_MAIN, font=("Segoe UI Semibold", 15)).pack(
            anchor="w"
        )
        tk.Label(status_card, textvariable=self.character_var, bg=CARD_BG, fg="#79c0ff", font=("Segoe UI", 10)).pack(
            anchor="w",
            pady=(6, 0),
        )

        actions = tk.Frame(body, bg=WINDOW_BG)
        actions.pack(fill="x", pady=(14, 0))
        self._make_button(actions, self.details_button_var.get(), self._toggle_details, accent="#38566b").pack(side="left")
        self.details_button = actions.winfo_children()[-1]
        self._make_button(actions, "Data Folder", self._open_data_folder, accent="#2d5f41").pack(side="left", padx=(8, 0))
        self._make_button(actions, "Duplicates", self._open_duplicates_window, accent="#5f4b2d").pack(side="left", padx=(8, 0))

        self.details_frame = self._card(body)
        tk.Label(self.details_frame, text="Recent Activity", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(
            anchor="w"
        )
        log_shell = tk.Frame(self.details_frame, bg=CARD_BG)
        log_shell.pack(fill="both", expand=True, pady=(8, 0))
        self.log_text = tk.Text(
            log_shell,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 9),
            wrap="word",
            height=8,
            yscrollcommand=lambda *args: self.log_scrollbar.set(*args),
        )
        self.log_scrollbar = ttk.Scrollbar(log_shell, orient="vertical", command=self.log_text.yview)
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_scrollbar.pack(side="right", fill="y")
        self.log_text.configure(state="disabled")
        tk.Label(self.details_frame, text="Live OCR", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(
            anchor="w",
            pady=(10, 0),
        )
        ocr_shell = tk.Frame(self.details_frame, bg=CARD_BG)
        ocr_shell.pack(fill="both", expand=True, pady=(8, 0))
        self.ocr_text = tk.Text(
            ocr_shell,
            bg=CARD_BG,
            fg="#79c0ff",
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Consolas", 9),
            wrap="word",
            height=6,
            yscrollcommand=lambda *args: self.ocr_scrollbar.set(*args),
        )
        self.ocr_scrollbar = ttk.Scrollbar(ocr_shell, orient="vertical", command=self.ocr_text.yview)
        self.ocr_text.pack(side="left", fill="both", expand=True)
        self.ocr_scrollbar.pack(side="right", fill="y")
        self.ocr_text.configure(state="disabled")
        self._set_text_widget(self.log_text, "-")
        self._set_text_widget(self.ocr_text, "-")

    def _card(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(parent, bg=CARD_BG, highlightbackground=CARD_BORDER, highlightthickness=1, padx=14, pady=14)

    def _make_button(self, parent: tk.Misc, text: str, command: object, *, accent: str) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=accent,
            fg=TEXT_MAIN,
            activebackground=accent,
            activeforeground=TEXT_MAIN,
            relief="flat",
            padx=14,
            pady=8,
            font=("Segoe UI Semibold", 10),
            cursor="hand2",
        )

    def _sync_main_scrollregion(self, _event: tk.Event | None = None) -> None:
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _sync_main_canvas_width(self, event: tk.Event) -> None:
        self.main_canvas.itemconfigure(self._main_canvas_window, width=event.width)

    def _on_main_mousewheel(self, event: tk.Event) -> str | None:
        if event.delta == 0:
            return None
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    @staticmethod
    def _set_text_widget(widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text.rstrip() or "-")
        widget.configure(state="disabled")

    def reload_settings_badges(self) -> None:
        config = load_config()
        submit_mode = "Auto-submit ON" if config.auto_submit else "Auto-submit OFF"
        provider = _detect_provider(config)
        provider_text = f"API: {provider}" if provider != PROVIDER_DISABLED else "API: Off"
        speed = _values_to_speed(config.typing_interval_seconds)
        community_text = "Community Sync ON" if config.github_answer_sheet_enabled else "Community Sync OFF"
        self.badges_var.set(f"{submit_mode}  |  {provider_text}  |  {community_text}  |  Typing speed {speed}/10")

    def _refresh_clients(self, initial: bool = False) -> None:
        config = load_config()
        matches = list_matching_windows(config)
        if not matches:
            self._client_choices = []
            self.client_combo["values"] = []
            self.selected_client.set("")
            self.character_var.set("No Divine MU client found. Open the client, then press Refresh.")
            if not initial:
                messagebox.showwarning(APP_NAME, "No visible Divine MU windows were found.")
            return

        choices: list[ClientChoice] = []
        name_counts: dict[str, int] = {}
        for match in matches:
            name = extract_character_name(match.title)
            name_counts[name] = name_counts.get(name, 0) + 1
        for match in matches:
            character_name = extract_character_name(match.title)
            label = character_name if name_counts[character_name] == 1 else f"{character_name} (window {match.match_index})"
            choices.append(
                ClientChoice(
                    label=label,
                    match_index=match.match_index,
                    character_name=character_name,
                    title=match.title,
                )
            )

        self._client_choices = choices
        self.client_combo["values"] = [choice.label for choice in choices]
        if self.selected_client.get() not in self.client_combo["values"]:
            self.selected_client.set(choices[0].label)
        self._on_client_selected()

    def _selected_choice(self) -> ClientChoice | None:
        label = self.selected_client.get().strip()
        for choice in self._client_choices:
            if choice.label == label:
                return choice
        return None

    def _on_client_selected(self) -> None:
        choice = self._selected_choice()
        if choice is None:
            self.character_var.set("No character selected")
            return
        self.character_var.set(f"Character: {choice.character_name}")

    def _open_settings(self) -> None:
        SettingsDialog(self)

    def _start_bot(self) -> None:
        if self._bot_thread and self._bot_thread.is_alive():
            return

        choice = self._selected_choice()
        if choice is None:
            messagebox.showwarning(APP_NAME, "Choose a character first.")
            return

        config = load_config()
        config.capture_source = "window"
        config.target_window_index = choice.match_index
        config.show_overlay = False

        self.append_log(f"Starting watcher for {choice.character_name}...")
        self.status_var.set("Starting live watcher...")
        self.character_var.set(f"Character: {choice.character_name}")
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.client_combo.config(state="disabled")
        self._launch_config = config
        self._pending_stop.clear()
        self._bot_thread = threading.Thread(target=self._run_bot, name="mu-bot-main", daemon=True)
        self._bot_thread.start()

    def _run_bot(self) -> None:
        config = self._launch_config
        if config is None:
            self._message_queue.put(("error", "No launch configuration was prepared for the bot thread."))
            return

        try:
            self._bot = MuUnscrambleBot(
                config=config,
                dry_run=False,
                status_callback=self._queue_status,
                log_callback=self._queue_log,
            )
            if self._pending_stop.is_set():
                self._bot.request_stop()
            self._bot.run_forever()
            self._message_queue.put(("stopped", "Watcher stopped."))
        except Exception:
            self._message_queue.put(("error", traceback.format_exc()))

    def _stop_bot(self) -> None:
        if self._bot is None and not (self._bot_thread and self._bot_thread.is_alive()):
            return
        self.status_var.set("Stopping...")
        self.append_log("Stopping watcher...")
        self._pending_stop.set()
        if self._bot is not None:
            self._bot.request_stop()

    def _start_update_check(self, *, silent_if_current: bool = False) -> None:
        config = load_config()
        self.append_log("Checking GitHub releases for updates...")
        thread = threading.Thread(
            target=self._run_update_check,
            args=(config.update_repository, silent_if_current),
            name="mu-update-check",
            daemon=True,
        )
        thread.start()

    def _run_update_check(self, repository: str, silent_if_current: bool) -> None:
        result = check_for_updates(repository)
        self._message_queue.put(("update", UpdateMessage(result=result, silent_if_current=silent_if_current)))

    def _start_update_install(self, result: UpdateCheckResult) -> None:
        self.append_log(f"Downloading version {result.latest_version} from GitHub...")
        thread = threading.Thread(
            target=self._run_update_install,
            args=(result,),
            name="mu-update-install",
            daemon=True,
        )
        thread.start()

    def _run_update_install(self, result: UpdateCheckResult) -> None:
        try:
            if result.manifest_asset_url:
                self._message_queue.put(("log", f"Fetching update manifest for {result.latest_version}..."))
                prepared = prepare_file_update(result)
                if prepared.changed_count == 0 and prepared.stale_count == 0:
                    self._message_queue.put(("log", "No file changes were needed for this update."))
                else:
                    self._message_queue.put(
                        (
                            "log",
                            f"Prepared {prepared.changed_count} changed files and {prepared.stale_count} cleanup items.",
                        )
                    )
                self._message_queue.put(("log", "Staging file-by-file update and preparing restart..."))
                update_log_path = stage_windows_file_update(prepared)
            else:
                self._message_queue.put(("log", f"Downloading {result.asset_name or result.latest_version}..."))
                archive_path = download_release_asset(result)
                self._message_queue.put(("log", f"Download complete: {archive_path.name}"))
                self._message_queue.put(("log", "Staging full-bundle update and preparing restart..."))
                update_log_path = stage_windows_update(archive_path)
            self._message_queue.put(("log", f"Updater log: {update_log_path}"))
        except Exception as exc:
            self._message_queue.put(("update-install-error", f"Update install failed: {type(exc).__name__}: {exc}"))
            return
        self._message_queue.put(("update-staged", result.latest_version))

    def _queue_status(self, payload: OverlayPayload) -> None:
        self._message_queue.put(("status", payload))

    def _queue_log(self, line: str) -> None:
        self._message_queue.put(("log", line))

    def _pump_messages(self) -> None:
        while True:
            try:
                kind, payload = self._message_queue.get_nowait()
            except queue.Empty:
                break

            if kind == "status":
                assert isinstance(payload, OverlayPayload)
                self.status_var.set(payload.status)
                self.answer_var.set(payload.answer_text or "-")
                self.hint_var.set(payload.hint_text or "-")
                round_text = payload.round_text or "-"
                self.round_var.set(f"Round {round_text}")
                self._set_text_widget(self.ocr_text, payload.ocr_text or "-")
            elif kind == "log":
                assert isinstance(payload, str)
                self.append_log(payload)
            elif kind == "stopped":
                assert isinstance(payload, str)
                self.append_log(payload)
                self._reset_running_state()
                self.status_var.set(payload)
            elif kind == "error":
                assert isinstance(payload, str)
                self.append_log("Bot crashed.")
                self.append_log(payload)
                self._reset_running_state()
                messagebox.showerror(APP_NAME, "The bot stopped unexpectedly. See the Activity Log for details.")
            elif kind == "update":
                assert isinstance(payload, UpdateMessage)
                self._handle_update_result(payload.result, silent_if_current=payload.silent_if_current)
            elif kind == "update-install-error":
                assert isinstance(payload, str)
                self.append_log(payload)
                messagebox.showerror(APP_NAME, payload)
            elif kind == "update-staged":
                assert isinstance(payload, str)
                self.append_log(f"Version {payload} downloaded. Restarting into the new build...")
                messagebox.showinfo(
                    APP_NAME,
                    f"Version {payload} is ready.\n\nThe app will close and relaunch into the new build now.",
                )
                if self._bot is not None:
                    self._bot.request_stop()
                self.root.after(120, self._exit_for_update)

        self.root.after(120, self._pump_messages)

    def _handle_update_result(self, result: UpdateCheckResult, *, silent_if_current: bool = False) -> None:
        if result.error:
            self.append_log(result.error)
            if not silent_if_current:
                messagebox.showinfo(APP_NAME, result.error)
            return

        if not result.available:
            if not silent_if_current:
                self.append_log(f"No update found. Current version is {result.current_version}.")
                messagebox.showinfo(APP_NAME, f"You are up to date on version {result.current_version}.")
            return

        self.append_log(f"Update available: {result.latest_version}")
        if result.asset_url and is_frozen():
            install_now = messagebox.askyesno(
                APP_NAME,
                f"Version {result.latest_version} is available.\n\nDownload and install it now?",
            )
            if install_now:
                self._start_update_install(result)
            return

        open_page = messagebox.askyesno(APP_NAME, f"Version {result.latest_version} is available.\n\nOpen the release page now?")
        if open_page:
            open_release_page(result.release_url)

    def _toggle_details(self) -> None:
        self._details_visible = not self._details_visible
        if self._details_visible:
            self.details_frame.pack(fill="x", pady=(12, 0))
            self.details_button_var.set("Hide Details")
            self.details_button.config(text=self.details_button_var.get())
            self.root.after(50, lambda: self.main_canvas.yview_moveto(1.0))
        else:
            self.details_frame.pack_forget()
            self.details_button_var.set("Show Details")
            self.details_button.config(text=self.details_button_var.get())
            self.main_canvas.yview_moveto(0.0)

    def _open_duplicates_window(self) -> None:
        config = load_config()
        from mu_unscramble_bot.github_answer_sheet import GitHubAnswerSheetConfig
        from mu_unscramble_bot.memory_store import DuplicateGroup, QuestionMemory

        github_sync = None
        if config.github_answer_sheet_enabled:
            github_sync = GitHubAnswerSheetConfig(
                repository=config.github_answer_sheet_repository,
                branch=config.github_answer_sheet_branch,
                path=config.github_answer_sheet_path,
                token=config.github_answer_sheet_token,
                sync_interval_seconds=config.github_answer_sheet_sync_interval_seconds,
                commit_message=config.github_answer_sheet_commit_message,
            )

        memory = QuestionMemory(
            path=config.question_memory_path,
            github_sync=github_sync,
            auto_sync_from_github=False,
        )
        window = tk.Toplevel(self.root)
        window.title(f"{APP_NAME} Duplicates")
        window.configure(bg=WINDOW_BG)
        window.geometry("980x620")
        window.minsize(900, 560)

        search_var = tk.StringVar()
        status_var = tk.StringVar(value="Choose a duplicate group to review.")
        count_var = tk.StringVar(value="")
        group_items: list[DuplicateGroup] = []

        container = tk.Frame(window, bg=WINDOW_BG, padx=16, pady=16)
        container.pack(fill="both", expand=True)
        tk.Label(
            container,
            text="Duplicate Review",
            bg=WINDOW_BG,
            fg=TEXT_MAIN,
            font=("Segoe UI Semibold", 16),
        ).pack(anchor="w")
        tk.Label(
            container,
            text="Find real conflicts where one letter set has been saved with multiple answers.",
            bg=WINDOW_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 9),
            wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        if config.github_answer_sheet_enabled:
            sync_note = (
                "Changes will sync back to GitHub."
                if (config.github_answer_sheet_token or "").strip()
                else "This PC is read-only for GitHub sync. Local removals may come back from the community sheet."
            )
            tk.Label(
                container,
                text=sync_note,
                bg=WINDOW_BG,
                fg="#79c0ff",
                font=("Segoe UI", 9),
                wraplength=840,
                justify="left",
            ).pack(anchor="w", pady=(6, 0))

        search_row = tk.Frame(container, bg=WINDOW_BG)
        search_row.pack(fill="x", pady=(14, 0))
        tk.Entry(
            search_row,
            textvariable=search_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True)

        tk.Button(
            search_row,
            text="Search",
            command=lambda: refresh_groups(),
            bg=BLUE,
            fg=TEXT_MAIN,
            relief="flat",
            padx=12,
            pady=6,
            font=("Segoe UI Semibold", 9),
        ).pack(side="left", padx=(8, 0))

        body = tk.Frame(container, bg=WINDOW_BG)
        body.pack(fill="both", expand=True, pady=(12, 0))

        groups_card = self._card(body)
        groups_card.pack(side="left", fill="both", expand=False)
        tk.Label(
            groups_card,
            text="Duplicate Groups",
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")
        tk.Label(
            groups_card,
            textvariable=count_var,
            bg=CARD_BG,
            fg="#79c0ff",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))

        group_list_frame = tk.Frame(groups_card, bg=CARD_BG)
        group_list_frame.pack(fill="both", expand=True, pady=(8, 0))
        group_tree = ttk.Treeview(
            group_list_frame,
            style="Dup.Treeview",
            columns=("key", "count"),
            show="headings",
            selectmode="browse",
            height=16,
        )
        group_tree.heading("key", text="Letter Set")
        group_tree.heading("count", text="Rows")
        group_tree.column("key", width=300, anchor="w")
        group_tree.column("count", width=58, anchor="center", stretch=False)
        group_scrollbar = ttk.Scrollbar(group_list_frame, orient="vertical", command=group_tree.yview)
        group_tree.configure(yscrollcommand=group_scrollbar.set)
        group_tree.pack(side="left", fill="both", expand=True)
        group_scrollbar.pack(side="right", fill="y")

        detail_card = self._card(body)
        detail_card.pack(side="left", fill="both", expand=True, padx=(12, 0))
        tk.Label(
            detail_card,
            textvariable=status_var,
            bg=CARD_BG,
            fg=TEXT_SOFT,
            font=("Segoe UI", 10, "bold"),
            wraplength=500,
            justify="left",
        ).pack(anchor="w")

        tree = ttk.Treeview(
            detail_card,
            style="Dup.Treeview",
            columns=("scramble", "answer", "frequency"),
            show="headings",
            selectmode="extended",
            height=14,
        )
        tree.heading("scramble", text="Scramble")
        tree.heading("answer", text="Answer")
        tree.heading("frequency", text="Freq")
        tree.column("scramble", width=180, anchor="w")
        tree.column("answer", width=180, anchor="w")
        tree.column("frequency", width=80, anchor="center")
        tree_scrollbar = ttk.Scrollbar(detail_card, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tree_scrollbar.set)
        tree.pack(side="left", fill="both", expand=True, pady=(10, 0))
        tree_scrollbar.pack(side="right", fill="y", pady=(10, 0))

        action_row = tk.Frame(container, bg=WINDOW_BG)
        action_row.pack(fill="x", pady=(12, 0))

        def selected_group() -> DuplicateGroup | None:
            selection = group_tree.selection()
            if not selection:
                return None
            try:
                index = int(selection[0])
            except Exception:
                return None
            if index < 0 or index >= len(group_items):
                return None
            return group_items[index]

        def selected_rows() -> list[tuple[str, str]]:
            rows: list[tuple[str, str]] = []
            for item_id in tree.selection():
                values = tree.item(item_id, "values")
                if len(values) < 2:
                    continue
                rows.append((str(values[0]), str(values[1])))
            return rows

        def refresh_group_rows() -> None:
            for item_id in tree.get_children():
                tree.delete(item_id)

            group = selected_group()
            if group is None:
                status_var.set("Choose a duplicate group to review.")
                return

            status_var.set(f"Letter set {group.key} has multiple answers. Keep one or remove bad OCR rows.")

            for record in group.records:
                tree.insert(
                    "",
                    "end",
                    iid=f"{record.scrambled_letters}|{record.answer}",
                    values=(record.scrambled_letters, record.answer, record.frequency),
                )

        def select_group(index: int) -> None:
            if index < 0 or index >= len(group_items):
                if group_tree.selection():
                    group_tree.selection_remove(group_tree.selection())
                refresh_group_rows()
                return
            item_id = str(index)
            if group_tree.exists(item_id):
                group_tree.selection_set(item_id)
                group_tree.focus(item_id)
                group_tree.see(item_id)
            refresh_group_rows()

        def refresh_groups() -> None:
            nonlocal group_items
            previous_group = selected_group()
            previous_key = (previous_group.kind, previous_group.key) if previous_group is not None else None
            try:
                group_items = memory.duplicate_groups(search_var.get().strip())
            except Exception as exc:
                status_var.set(f"Could not load duplicates: {type(exc).__name__}: {exc}")
                group_items = []

            for item_id in group_tree.get_children():
                group_tree.delete(item_id)

            count_var.set(f"{len(group_items)} groups found")
            for index, group in enumerate(group_items):
                group_tree.insert(
                    "",
                    "end",
                    iid=str(index),
                    values=(group.key, len(group.records)),
                )

            if not group_items:
                if not status_var.get().startswith("Could not load duplicates:"):
                    status_var.set("No duplicates found.")
                refresh_group_rows()
                return

            select_index = 0
            if previous_key is not None:
                for index, group in enumerate(group_items):
                    if (group.kind, group.key) == previous_key:
                        select_index = index
                        break

            select_group(select_index)

        def keep_selected() -> None:
            group = selected_group()
            rows = selected_rows()
            if group is None:
                messagebox.showwarning(APP_NAME, "Choose a duplicate group first.")
                return
            if len(rows) != 1:
                messagebox.showwarning(APP_NAME, "Select exactly one row to keep.")
                return

            removed = memory.keep_record_for_group(group.kind, group.key, rows[0])
            if removed <= 0:
                messagebox.showinfo(APP_NAME, "Nothing changed.")
                return

            self.append_log(f"Removed {removed} duplicate row(s) while keeping {rows[0][1]} for {group.key}.")
            refresh_groups()

        def remove_selected() -> None:
            rows = selected_rows()
            if not rows:
                messagebox.showwarning(APP_NAME, "Select one or more rows to remove.")
                return
            if not messagebox.askyesno(APP_NAME, f"Remove {len(rows)} selected row(s) from the answer sheet?"):
                return

            removed = memory.delete_records(rows)
            if removed <= 0:
                messagebox.showinfo(APP_NAME, "Nothing changed.")
                return

            self.append_log(f"Removed {removed} duplicate row(s) from the answer sheet.")
            refresh_groups()

        self._make_button(action_row, "Keep Selected", keep_selected, accent=GREEN).pack(side="left")
        self._make_button(action_row, "Remove Selected", remove_selected, accent=RED).pack(side="left", padx=(8, 0))
        self._make_button(action_row, "Refresh", refresh_groups, accent="#345369").pack(side="left", padx=(8, 0))

        group_tree.bind("<<TreeviewSelect>>", lambda _event: refresh_group_rows())
        tree.bind("<Double-1>", lambda _event: keep_selected())
        window.bind("<Return>", lambda _event: keep_selected())

        refresh_groups()

    def _open_data_folder(self) -> None:
        path = user_data_dir()
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def append_log(self, text: str) -> None:
        self._log_lines.append(text.rstrip())
        self._log_lines = self._log_lines[-200:]
        self._set_text_widget(self.log_text, "\n".join(self._log_lines))
        self.log_text.see("end")

    def _exit_for_update(self) -> None:
        try:
            self.root.destroy()
        finally:
            os._exit(0)

    def _reset_running_state(self) -> None:
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.client_combo.config(state="readonly")
        self._bot = None
        self._launch_config = None
        self._pending_stop.clear()
        self._bot_thread = None

    def _on_close(self) -> None:
        if self._bot is not None:
            self._bot.request_stop()
        self.root.destroy()


def _show_admin_required_message() -> None:
    ctypes.windll.user32.MessageBoxW(
        0,
        "MU Unscramble Bot must be started as Administrator.\n\nThe app will close now.",
        APP_NAME,
        0x10,
    )


def main() -> int:
    if is_current_process_elevated() is not True:
        _show_admin_required_message()
        return 1

    root = tk.Tk()
    DesktopApp(root)
    root.mainloop()
    return 0
