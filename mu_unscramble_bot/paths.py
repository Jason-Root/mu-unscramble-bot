from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


APP_NAME = "MU Unscramble Bot"


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def bundle_dir() -> Path:
    if is_frozen():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def user_data_dir() -> Path:
    if not is_frozen():
        return Path(__file__).resolve().parent.parent

    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        root = Path(local_appdata)
    else:
        root = Path.home() / "AppData" / "Local"
    path = root / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


def resolve_user_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return user_data_dir() / path


def default_config_path() -> Path:
    return user_data_dir() / "config.json"


def default_env_path() -> Path:
    return user_data_dir() / ".env"


def ensure_runtime_files() -> None:
    destination_root = user_data_dir()
    source_root = bundle_dir()

    _copy_missing(source_root / "config.json", destination_root / "config.json")
    _copy_missing(source_root / ".env.example", destination_root / ".env.example")

    env_source = source_root / ".env"
    if env_source.exists():
        _copy_missing(env_source, destination_root / ".env")
    else:
        _copy_missing(source_root / ".env.example", destination_root / ".env")

    source_data_dir = source_root / "data"
    if source_data_dir.exists():
        _copy_tree_missing(source_data_dir, destination_root / "data")


def _copy_missing(source: Path, destination: Path) -> None:
    if not source.exists() or destination.exists():
        return
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def _copy_tree_missing(source_dir: Path, destination_dir: Path) -> None:
    if not source_dir.exists():
        return
    for source_path in source_dir.rglob("*"):
        if source_path.is_dir():
            continue
        relative = source_path.relative_to(source_dir)
        _copy_missing(source_path, destination_dir / relative)
