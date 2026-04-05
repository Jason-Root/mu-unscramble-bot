from __future__ import annotations

from dataclasses import dataclass
import queue
import threading
from typing import Final

from mu_unscramble_bot.config import BotConfig

try:
    import tkinter as tk
except Exception:  # pragma: no cover - tkinter availability varies
    tk = None


EMPTY: Final[str] = "-"


@dataclass(slots=True)
class OverlayPayload:
    status: str
    round_text: str = EMPTY
    scramble_text: str = EMPTY
    hint_text: str = EMPTY
    answer_text: str = EMPTY
    method_text: str = EMPTY
    ocr_text: str = EMPTY


class StatusOverlay:
    def __init__(self, config: BotConfig) -> None:
        self.config = config
        self.enabled = bool(config.show_overlay and tk is not None)
        self._queue: queue.SimpleQueue[OverlayPayload | None] = queue.SimpleQueue()
        self._thread: threading.Thread | None = None

        if self.enabled:
            self._thread = threading.Thread(target=self._run_ui, name="mu-overlay", daemon=True)
            self._thread.start()

    def update(
        self,
        *,
        status: str,
        round_text: str = EMPTY,
        scramble_text: str = EMPTY,
        hint_text: str = EMPTY,
        answer_text: str = EMPTY,
        method_text: str = EMPTY,
        ocr_text: str = EMPTY,
    ) -> None:
        if not self.enabled:
            return
        self._queue.put(
            OverlayPayload(
                status=status,
                round_text=round_text,
                scramble_text=scramble_text,
                hint_text=hint_text,
                answer_text=answer_text,
                method_text=method_text,
                ocr_text=ocr_text,
            )
        )

    def close(self) -> None:
        if not self.enabled:
            return
        self._queue.put(None)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.5)

    def _run_ui(self) -> None:
        assert tk is not None
        root = tk.Tk()
        root.title(self.config.overlay_title)
        root.configure(bg="#0d1117")
        root.resizable(False, False)
        if self.config.overlay_topmost:
            root.attributes("-topmost", True)

        self._position_window(root)

        container = tk.Frame(root, bg="#0d1117", padx=14, pady=12)
        container.pack(fill="both", expand=True)

        status_var = tk.StringVar(value="Starting...")
        round_var = tk.StringVar(value=f"Round: {EMPTY}")
        scramble_var = tk.StringVar(value=f"Scramble: {EMPTY}")
        hint_var = tk.StringVar(value=f"Hint: {EMPTY}")
        answer_var = tk.StringVar(value=f"Answer: {EMPTY}")
        method_var = tk.StringVar(value=f"Solver: {EMPTY}")
        ocr_var = tk.StringVar(value=f"OCR:\n{EMPTY}")

        tk.Label(
            container,
            textvariable=status_var,
            bg="#0d1117",
            fg="#c9d1d9",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            container,
            textvariable=round_var,
            bg="#0d1117",
            fg="#f2cc60",
            font=("Segoe UI", 11),
            anchor="w",
        ).pack(fill="x", pady=(8, 0))
        tk.Label(
            container,
            textvariable=scramble_var,
            bg="#0d1117",
            fg="#f2cc60",
            font=("Consolas", 12, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            container,
            textvariable=hint_var,
            bg="#0d1117",
            fg="#f2cc60",
            font=("Segoe UI", 10),
            justify="left",
            wraplength=self.config.overlay_width - 30,
            anchor="w",
        ).pack(fill="x", pady=(2, 0))
        tk.Label(
            container,
            textvariable=answer_var,
            bg="#0d1117",
            fg="#7ee787",
            font=("Consolas", 14, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(10, 0))
        tk.Label(
            container,
            textvariable=method_var,
            bg="#0d1117",
            fg="#8b949e",
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            container,
            textvariable=ocr_var,
            bg="#0d1117",
            fg="#79c0ff",
            font=("Consolas", 9),
            justify="left",
            wraplength=self.config.overlay_width - 30,
            anchor="w",
        ).pack(fill="x", pady=(10, 0))

        def pump() -> None:
            alive = True
            while True:
                try:
                    payload = self._queue.get_nowait()
                except queue.Empty:
                    break

                if payload is None:
                    alive = False
                    break

                status_var.set(payload.status)
                round_var.set(f"Round: {payload.round_text}")
                scramble_var.set(f"Scramble: {payload.scramble_text}")
                hint_var.set(f"Hint: {payload.hint_text}")
                answer_var.set(f"Answer: {payload.answer_text}")
                method_var.set(f"Solver: {payload.method_text}")
                ocr_var.set(f"OCR:\n{payload.ocr_text}")

            if alive:
                root.after(120, pump)
            else:
                root.destroy()

        root.protocol("WM_DELETE_WINDOW", root.withdraw)
        root.after(120, pump)
        root.mainloop()

    def _position_window(self, root: tk.Tk) -> None:
        root.update_idletasks()
        screen_width = root.winfo_screenwidth()
        x = int((screen_width - self.config.overlay_width) / 2) + self.config.overlay_left_offset
        y = max(0, self.config.overlay_top_offset)
        root.geometry(f"{self.config.overlay_width}x{self.config.overlay_height}+{x}+{y}")
