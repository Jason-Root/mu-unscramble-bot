from __future__ import annotations

from dataclasses import dataclass
import ctypes
import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import messagebox, ttk

from mu_unscramble_bot.bot import MuUnscrambleBot
from mu_unscramble_bot.config import BotConfig, load_config, load_env_settings, save_config, save_env_settings
from mu_unscramble_bot.overlay import OverlayPayload
from mu_unscramble_bot.paths import APP_NAME, user_data_dir
from mu_unscramble_bot.privilege import is_current_process_elevated
from mu_unscramble_bot.updater import UpdateCheckResult, check_for_updates, get_app_version, open_release_page
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


@dataclass(slots=True)
class ClientChoice:
    label: str
    match_index: int
    character_name: str
    title: str


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


def _detect_provider(config: BotConfig) -> str:
    base_url = (config.openai_base_url or "").strip().lower()
    api_key = (config.openai_api_key or "").strip()
    if not base_url and not api_key:
        return PROVIDER_DISABLED
    if "openrouter.ai" in base_url:
        return PROVIDER_OPENROUTER
    if "127.0.0.1" in base_url or "localhost" in base_url:
        return PROVIDER_LOCAL
    if base_url:
        return PROVIDER_CUSTOM
    return PROVIDER_DISABLED


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
        self.speed_var = tk.IntVar(value=_values_to_speed(self.config.typing_interval_seconds))
        self.speed_text_var = tk.StringVar()

        container = tk.Frame(self.window, bg=WINDOW_BG, padx=18, pady=18)
        container.pack(fill="both", expand=True)

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
        tk.Entry(
            api_card,
            textvariable=self.model_var,
            bg="#0b1620",
            fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN,
            relief="flat",
            font=("Segoe UI", 10),
        ).pack(fill="x", pady=(8, 0))

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

        actions = tk.Frame(container, bg=WINDOW_BG)
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
        width = 520
        height = self.window.winfo_height()
        x = parent.root.winfo_rootx() + max(30, (parent.root.winfo_width() - width) // 2)
        y = parent.root.winfo_rooty() + 40
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def _card(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(parent, bg=CARD_BG, highlightbackground=CARD_BORDER, highlightthickness=1, padx=14, pady=14)

    def _row_label(self, parent: tk.Misc, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold"))

    def _apply_provider_preset(self, initial: bool = False) -> None:
        provider = self.provider_var.get()
        if provider == PROVIDER_OPENROUTER:
            if initial or not self.base_url_var.get().strip():
                self.base_url_var.set("https://openrouter.ai/api/v1")
            if initial or not self.model_var.get().strip():
                self.model_var.set("qwen/qwen3.6-plus:free")
        elif provider == PROVIDER_LOCAL:
            if initial or not self.base_url_var.get().strip():
                self.base_url_var.set("http://127.0.0.1:11434/v1")
            if initial or not self.model_var.get().strip():
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

    def _save(self) -> None:
        provider = self.provider_var.get()
        model = self.model_var.get().strip()
        base_url = self.base_url_var.get().strip()
        api_key = self.api_key_var.get().strip()

        if provider != PROVIDER_DISABLED and not model:
            messagebox.showwarning(APP_NAME, "Enter a model name before saving.")
            return
        if provider in {PROVIDER_LOCAL, PROVIDER_CUSTOM, PROVIDER_OPENROUTER} and not base_url and provider != PROVIDER_DISABLED:
            messagebox.showwarning(APP_NAME, "Enter a base URL before saving.")
            return

        config = load_config()
        config.show_overlay = False
        config.test_api_on_startup = False
        config.typing_interval_seconds, config.key_hold_seconds = _speed_to_values(self.speed_var.get())
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

        save_env_settings(env_updates)
        self.parent.reload_settings_badges()
        self.parent.append_log("Settings saved.")
        self.window.destroy()
        messagebox.showinfo(APP_NAME, "Settings saved. Restart the watcher if it is already running.")


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("760x470")
        self.root.minsize(680, 430)
        self.root.configure(bg=WINDOW_BG)

        self._message_queue: queue.SimpleQueue[tuple[str, object]] = queue.SimpleQueue()
        self._bot_thread: threading.Thread | None = None
        self._bot: MuUnscrambleBot | None = None
        self._launch_config: BotConfig | None = None
        self._pending_stop = threading.Event()
        self._client_choices: list[ClientChoice] = []
        self._details_visible = False

        self.selected_client = tk.StringVar()
        self.status_var = tk.StringVar(value="Choose a character, then go live.")
        self.character_var = tk.StringVar(value="No character selected")
        self.answer_var = tk.StringVar(value="-")
        self.hint_var = tk.StringVar(value="-")
        self.round_var = tk.StringVar(value="-")
        self.details_button_var = tk.StringVar(value="Show Details")
        self.badges_var = tk.StringVar(value="")
        self.version_var = tk.StringVar(value=f"v{get_app_version()}")
        self.log_var = tk.StringVar(value="")
        self.ocr_var = tk.StringVar(value="-")

        self._build_styles()
        self._build_ui()
        self._refresh_clients(initial=True)
        self.reload_settings_badges()
        self.root.after(120, self._pump_messages)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
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

        body = tk.Frame(self.root, bg=WINDOW_BG, padx=18, pady=0)
        body.pack(fill="both", expand=True)

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

        answer_row = tk.Frame(body, bg=WINDOW_BG)
        answer_row.pack(fill="x", pady=(14, 0))

        answer_card = self._card(answer_row)
        answer_card.pack(side="left", fill="both", expand=True)
        tk.Label(answer_card, text="Answer", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(
            answer_card,
            textvariable=self.answer_var,
            bg=CARD_BG,
            fg="#7ee787",
            font=("Consolas", 22, "bold"),
            wraplength=300,
            justify="left",
        ).pack(anchor="w", pady=(10, 0))
        tk.Label(answer_card, textvariable=self.round_var, bg=CARD_BG, fg="#f2cc60", font=("Segoe UI", 10)).pack(
            anchor="w",
            pady=(8, 0),
        )

        hint_card = self._card(answer_row)
        hint_card.pack(side="left", fill="both", expand=True, padx=(14, 0))
        tk.Label(hint_card, text="Hint", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(
            hint_card,
            textvariable=self.hint_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            font=("Segoe UI", 11),
            justify="left",
            anchor="nw",
            wraplength=320,
            height=6,
        ).pack(fill="both", expand=True, pady=(10, 0))

        actions = tk.Frame(body, bg=WINDOW_BG)
        actions.pack(fill="x", pady=(12, 0))
        self._make_button(actions, self.details_button_var.get(), self._toggle_details, accent="#38566b").pack(side="left")
        self.details_button = actions.winfo_children()[-1]
        self._make_button(actions, "Data Folder", self._open_data_folder, accent="#2d5f41").pack(side="left", padx=(8, 0))

        self.details_frame = self._card(body)
        tk.Label(self.details_frame, text="Recent Activity", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(
            anchor="w"
        )
        self.log_label = tk.Label(
            self.details_frame,
            textvariable=self.log_var,
            bg=CARD_BG,
            fg=TEXT_MAIN,
            font=("Consolas", 9),
            justify="left",
            anchor="nw",
            wraplength=680,
            height=8,
        )
        self.log_label.pack(fill="x", pady=(8, 0))
        tk.Label(self.details_frame, text="Live OCR", bg=CARD_BG, fg=TEXT_SOFT, font=("Segoe UI", 10, "bold")).pack(
            anchor="w",
            pady=(10, 0),
        )
        tk.Label(
            self.details_frame,
            textvariable=self.ocr_var,
            bg=CARD_BG,
            fg="#79c0ff",
            font=("Consolas", 9),
            justify="left",
            anchor="nw",
            wraplength=680,
            height=6,
        ).pack(fill="x", pady=(8, 0))

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

    def reload_settings_badges(self) -> None:
        config = load_config()
        submit_mode = "Auto-submit ON" if config.auto_submit else "Auto-submit OFF"
        provider = _detect_provider(config)
        provider_text = f"API: {provider}" if provider != PROVIDER_DISABLED else "API: Off"
        speed = _values_to_speed(config.typing_interval_seconds)
        self.badges_var.set(f"{submit_mode}  |  {provider_text}  |  Typing speed {speed}/10")

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

    def _start_update_check(self) -> None:
        config = load_config()
        self.append_log("Checking GitHub releases for updates...")
        thread = threading.Thread(
            target=self._run_update_check,
            args=(config.update_repository,),
            name="mu-update-check",
            daemon=True,
        )
        thread.start()

    def _run_update_check(self, repository: str) -> None:
        result = check_for_updates(repository)
        self._message_queue.put(("update", result))

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
                self.ocr_var.set(payload.ocr_text or "-")
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
                assert isinstance(payload, UpdateCheckResult)
                self._handle_update_result(payload)

        self.root.after(120, self._pump_messages)

    def _handle_update_result(self, result: UpdateCheckResult) -> None:
        if result.error:
            self.append_log(result.error)
            messagebox.showinfo(APP_NAME, result.error)
            return

        if not result.available:
            self.append_log(f"No update found. Current version is {result.current_version}.")
            messagebox.showinfo(APP_NAME, f"You are up to date on version {result.current_version}.")
            return

        self.append_log(f"Update available: {result.latest_version}")
        open_page = messagebox.askyesno(
            APP_NAME,
            f"Version {result.latest_version} is available.\n\nOpen the release page now?",
        )
        if open_page:
            open_release_page(result.release_url)

    def _toggle_details(self) -> None:
        self._details_visible = not self._details_visible
        if self._details_visible:
            self.details_frame.pack(fill="x", pady=(12, 0))
            self.details_button_var.set("Hide Details")
            self.details_button.config(text=self.details_button_var.get())
            self.root.geometry("760x690")
        else:
            self.details_frame.pack_forget()
            self.details_button_var.set("Show Details")
            self.details_button.config(text=self.details_button_var.get())
            self.root.geometry("760x470")

    def _open_data_folder(self) -> None:
        path = user_data_dir()
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def append_log(self, text: str) -> None:
        existing = self.log_var.get().splitlines()
        existing.append(text.rstrip())
        self.log_var.set("\n".join(existing[-12:]))

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
