from __future__ import annotations

from dataclasses import dataclass
import json
from importlib.metadata import PackageNotFoundError, version as package_version
import urllib.error
import urllib.request
import webbrowser

from packaging.version import InvalidVersion, Version

from mu_unscramble_bot import __version__


@dataclass(frozen=True, slots=True)
class UpdateCheckResult:
    current_version: str
    latest_version: str = ""
    available: bool = False
    release_url: str = ""
    notes: str = ""
    error: str = ""


def get_app_version() -> str:
    try:
        return package_version("mu-unscramble-bot")
    except PackageNotFoundError:
        return __version__


def check_for_updates(repository: str, *, timeout_seconds: float = 8.0) -> UpdateCheckResult:
    current = get_app_version()
    repository = repository.strip().strip("/")
    if not repository:
        return UpdateCheckResult(
            current_version=current,
            error="Update repository is not configured yet.",
        )

    url = f"https://api.github.com/repos/{repository}/releases/latest"
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "mu-unscramble-bot",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        return UpdateCheckResult(
            current_version=current,
            error=f"GitHub update check failed: HTTP {exc.code}",
        )
    except Exception as exc:
        return UpdateCheckResult(
            current_version=current,
            error=f"GitHub update check failed: {type(exc).__name__}: {exc}",
        )

    latest = str(payload.get("tag_name", "") or payload.get("name", "") or "").strip()
    latest = latest.removeprefix("v")
    current_key = _safe_version(current)
    latest_key = _safe_version(latest)
    available = bool(latest and latest_key is not None and current_key is not None and latest_key > current_key)
    return UpdateCheckResult(
        current_version=current,
        latest_version=latest,
        available=available,
        release_url=str(payload.get("html_url", "") or ""),
        notes=str(payload.get("body", "") or "").strip(),
        error="" if latest else "No release version was returned from GitHub.",
    )


def open_release_page(url: str) -> None:
    if not url.strip():
        return
    webbrowser.open(url.strip())


def _safe_version(value: str) -> Version | None:
    try:
        return Version(value)
    except InvalidVersion:
        return None
