from __future__ import annotations

import ctypes
import time

import pyautogui
import pydirectinput

from mu_unscramble_bot.config import BotConfig
from mu_unscramble_bot.privilege import get_window_pid, is_current_process_elevated, is_pid_elevated
from mu_unscramble_bot.window_target import WindowSelectionError, get_target_window


class AnswerSubmitter:
    def __init__(self, config: BotConfig) -> None:
        self.config = config
        self.backend_name = config.submit_backend.lower()
        self.backend = pydirectinput if self.backend_name == "directinput" else pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0
        if hasattr(pydirectinput, "FAILSAFE"):
            pydirectinput.FAILSAFE = True
        if hasattr(pydirectinput, "PAUSE"):
            pydirectinput.PAUSE = 0

    def submit(self, answer: str) -> bool:
        if not self._ensure_target_window():
            return False

        submit_text = self._build_submit_text(answer)
        if self.config.open_chat_before_submit:
            self._send_key(self.config.open_chat_key)
            time.sleep(self.config.pre_submit_delay_seconds)

        self._type_submit_text(submit_text)
        time.sleep(max(0.05, self.config.pre_submit_delay_seconds / 2))
        self._send_key(self.config.submit_key)
        time.sleep(self.config.post_submit_delay_seconds)
        return True

    def _ensure_target_window(self) -> bool:
        try:
            target = get_target_window(self.config)
        except WindowSelectionError:
            return not self.config.require_window_match

        target_pid = get_window_pid(target.window)
        current_elevated = is_current_process_elevated()
        target_elevated = is_pid_elevated(target_pid) if target_pid is not None else None
        if current_elevated is False and target_elevated is True:
            return False

        active_title = self._active_window_title()
        if target.title and active_title == target.title:
            if self.config.focus_window_before_submit:
                self._click_client_body(target)
            return True

        if self.config.focus_window_before_submit:
            self._focus_window(target.window)
            self._click_client_body(target)

            active_title = self._active_window_title()
            if target.title and active_title == target.title:
                return True

            # Some MU clients accept input after a real click even when Windows
            # does not immediately report the title switch back to us.
            return True

        return not self.config.require_window_match

    def _build_submit_text(self, answer: str) -> str:
        command_word = (self.config.submit_command_word or "").strip().lstrip("/")
        if command_word:
            return f"/{command_word} {answer}".strip()

        template = self.config.submit_text_template.strip()
        if not template:
            return answer
        if "{answer}" in template:
            return template.format(answer=answer)
        return f"{template} {answer}".strip()

    @staticmethod
    def _focus_window(window: object) -> None:
        hwnd = getattr(window, "_hWnd", None)
        if hwnd:
            try:
                user32 = ctypes.windll.user32
                user32.ShowWindow(hwnd, 9)
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.15)
                return
            except Exception:
                pass

        try:
            window.activate()
            time.sleep(0.18)
        except Exception:
            pass

    def _type_submit_text(self, text: str) -> None:
        for character in text:
            key_name = self._map_character_to_key(character)
            if key_name is None:
                continue
            self._send_key(key_name)
            time.sleep(self.config.typing_interval_seconds)

    def _send_key(self, key_name: str) -> None:
        key_down = getattr(self.backend, "keyDown", None)
        key_up = getattr(self.backend, "keyUp", None)
        if key_down is None or key_up is None:
            self.backend.press(key_name)
            return

        key_down(key_name)
        time.sleep(self.config.key_hold_seconds)
        key_up(key_name)

    @staticmethod
    def _map_character_to_key(character: str) -> str | None:
        if character.isalpha():
            return character.lower()
        if character.isdigit():
            return character
        mapping = {
            " ": "space",
            "/": "/",
            "-": "-",
            ".": ".",
            ",": ",",
        }
        return mapping.get(character)

    @staticmethod
    def _click_client_body(target: object) -> None:
        left = int(getattr(target, "left", 0))
        top = int(getattr(target, "top", 0))
        width = int(getattr(target, "width", 0))
        height = int(getattr(target, "height", 0))
        if width <= 0 or height <= 0:
            return

        # Click well inside the client body so the top overlay window does not
        # intercept the focus click during live runs.
        click_x = left + max(40, min(width - 40, width // 2))
        click_y = top + max(140, min(height - 80, int(height * 0.62)))
        original_x, original_y = pyautogui.position()
        try:
            pyautogui.click(click_x, click_y)
            time.sleep(0.15)
        finally:
            pyautogui.moveTo(original_x, original_y, duration=0)

    @staticmethod
    def _active_window_title() -> str:
        try:
            import pygetwindow as gw

            active = gw.getActiveWindow() if gw is not None else None
            return active.title if active and active.title else ""
        except Exception:
            return ""
