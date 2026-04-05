from __future__ import annotations

from dataclasses import dataclass
import re

from mu_unscramble_bot.config import BotConfig

try:
    import pygetwindow as gw
except Exception:  # pragma: no cover - platform-specific import
    gw = None


@dataclass(slots=True)
class WindowMatch:
    match_index: int
    title: str
    left: int
    top: int
    width: int
    height: int
    is_minimized: bool
    window: object


class WindowSelectionError(RuntimeError):
    pass


CHARACTER_NAME_PATTERNS = (
    re.compile(r"Name:\s*\[([^\]]+)\]", re.IGNORECASE),
    re.compile(r"\[([^\]]+)\](?!.*\[)"),
)


def list_matching_windows(config: BotConfig, *, visible_only: bool | None = None) -> list[WindowMatch]:
    if gw is None:
        return []

    visible_only = config.target_window_visible_only if visible_only is None else visible_only
    exact_title = config.target_window_exact_title.strip()
    title_contains = _title_contains(config)

    matches: list[WindowMatch] = []
    for window in gw.getAllWindows():
        title = (getattr(window, "title", "") or "").strip()
        if not title:
            continue
        if exact_title:
            if title != exact_title:
                continue
        elif title_contains and title_contains.lower() not in title.lower():
            continue

        left = int(getattr(window, "left", 0))
        top = int(getattr(window, "top", 0))
        width = int(getattr(window, "width", 0))
        height = int(getattr(window, "height", 0))
        is_minimized = bool(getattr(window, "isMinimized", False)) or left <= -32000 or top <= -32000

        if visible_only and (is_minimized or width <= 0 or height <= 0):
            continue

        matches.append(
            WindowMatch(
                match_index=-1,
                title=title,
                left=left,
                top=top,
                width=width,
                height=height,
                is_minimized=is_minimized,
                window=window,
            )
        )

    matches.sort(key=lambda item: (item.top, item.left, item.title.lower()))
    for index, match in enumerate(matches):
        match.match_index = index
    return matches


def get_target_window(config: BotConfig) -> WindowMatch:
    matches = list_matching_windows(config)
    if not matches:
        title_filter = config.target_window_exact_title.strip() or _title_contains(config) or "<any>"
        raise WindowSelectionError(f"No matching window found for filter: {title_filter}")

    index = config.target_window_index
    if index < 0 or index >= len(matches):
        raise WindowSelectionError(
            f"Window index {index} is out of range. Matching windows: {len(matches)}"
        )
    return matches[index]


def extract_character_name(title: str) -> str:
    for pattern in CHARACTER_NAME_PATTERNS:
        match = pattern.search(title)
        if match:
            name = match.group(1).strip()
            if name:
                return name
    return title.strip() or "Unknown"


def _title_contains(config: BotConfig) -> str:
    title_contains = config.target_window_title_contains.strip()
    if title_contains:
        return title_contains
    return config.focus_window_title_contains.strip()
