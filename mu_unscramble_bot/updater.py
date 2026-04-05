from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import textwrap
from importlib.metadata import PackageNotFoundError, version as package_version
import urllib.error
import urllib.parse
import urllib.request
import webbrowser

from packaging.version import InvalidVersion, Version

from mu_unscramble_bot import __version__
from mu_unscramble_bot.paths import is_frozen, user_data_dir


@dataclass(frozen=True, slots=True)
class UpdateCheckResult:
    current_version: str
    latest_version: str = ""
    available: bool = False
    release_url: str = ""
    notes: str = ""
    asset_name: str = ""
    asset_url: str = ""
    manifest_asset_name: str = ""
    manifest_asset_url: str = ""
    assets: tuple["ReleaseAsset", ...] = ()
    error: str = ""


@dataclass(frozen=True, slots=True)
class ReleaseAsset:
    name: str
    url: str
    size: int = 0


@dataclass(frozen=True, slots=True)
class UpdateManifestFile:
    path: str
    sha256: str
    size: int
    asset_name: str = ""


@dataclass(frozen=True, slots=True)
class UpdateManifest:
    version: str
    generated_at: str = ""
    files: tuple[UpdateManifestFile, ...] = ()


@dataclass(frozen=True, slots=True)
class PreparedFileUpdate:
    stage_root: Path
    changed_count: int
    stale_count: int
    update_log_path: Path


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
    assets = _extract_release_assets(payload)
    asset_name, asset_url = _pick_release_asset(assets)
    manifest_asset_name, manifest_asset_url = _pick_manifest_asset(assets)
    return UpdateCheckResult(
        current_version=current,
        latest_version=latest,
        available=available,
        release_url=str(payload.get("html_url", "") or ""),
        notes=str(payload.get("body", "") or "").strip(),
        asset_name=asset_name,
        asset_url=asset_url,
        manifest_asset_name=manifest_asset_name,
        manifest_asset_url=manifest_asset_url,
        assets=assets,
        error="" if latest else "No release version was returned from GitHub.",
    )


def open_release_page(url: str) -> None:
    if not url.strip():
        return
    webbrowser.open(url.strip())


def fetch_release_manifest(result: UpdateCheckResult, *, timeout_seconds: float = 45.0) -> UpdateManifest:
    manifest_url = result.manifest_asset_url.strip()
    if not manifest_url:
        raise RuntimeError("This release does not expose a file update manifest yet.")

    payload = _download_json(manifest_url, timeout_seconds=timeout_seconds)
    raw_files = payload.get("files", [])
    if not isinstance(raw_files, list):
        raise RuntimeError("The update manifest is malformed.")

    files: list[UpdateManifestFile] = []
    for entry in raw_files:
        if not isinstance(entry, dict):
            continue
        path = str(entry.get("path", "") or "").strip().replace("\\", "/")
        sha256 = str(entry.get("sha256", "") or "").strip().lower()
        asset_name = str(entry.get("asset_name", "") or "").strip()
        size_value = entry.get("size", 0)
        try:
            size = max(0, int(size_value))
        except Exception:
            size = 0
        if not path or not sha256:
            continue
        files.append(UpdateManifestFile(path=path, sha256=sha256, size=size, asset_name=asset_name))

    if not files:
        raise RuntimeError("The update manifest did not contain any files.")

    return UpdateManifest(
        version=str(payload.get("version", result.latest_version) or result.latest_version),
        generated_at=str(payload.get("generated_at", "") or ""),
        files=tuple(files),
    )


def download_release_asset(
    result: UpdateCheckResult,
    *,
    destination_dir: str | Path | None = None,
    timeout_seconds: float = 45.0,
) -> Path:
    asset_url = result.asset_url.strip()
    if not asset_url:
        raise RuntimeError("This release does not expose a downloadable Windows package yet.")

    if destination_dir is None:
        destination_root = Path(tempfile.mkdtemp(prefix="mu-unscramble-update-"))
    else:
        destination_root = Path(destination_dir)
        destination_root.mkdir(parents=True, exist_ok=True)

    file_name = result.asset_name.strip() or Path(urllib.parse.urlparse(asset_url).path).name or "update.zip"
    destination = destination_root / file_name
    _download_binary(asset_url, destination, timeout_seconds=timeout_seconds)
    return destination


def prepare_file_update(result: UpdateCheckResult, *, timeout_seconds: float = 45.0) -> PreparedFileUpdate:
    if not is_frozen():
        raise RuntimeError("Automatic file updates only work from the packaged app build.")

    manifest = fetch_release_manifest(result, timeout_seconds=timeout_seconds)
    executable = Path(sys.executable).resolve()
    install_root = executable.parent.parent
    update_log_path = user_data_dir() / "update.log"
    asset_map = {asset.name: asset for asset in result.assets}

    changed_files: list[UpdateManifestFile] = []
    expected_paths = {entry.path for entry in manifest.files}
    local_files = _list_managed_files(install_root)

    for entry in manifest.files:
        local_path = install_root / Path(*entry.path.split("/"))
        if _file_matches_manifest(local_path, entry):
            continue
        changed_files.append(entry)

    stale_paths = tuple(sorted(path for path in local_files if path not in expected_paths))

    stage_root = Path(tempfile.mkdtemp(prefix="mu-unscramble-file-update-"))
    download_root = stage_root / "files"
    manifest_path = stage_root / "update-manifest.json"
    plan_path = stage_root / "file-update-plan.json"
    download_root.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(
        json.dumps(
            {
                "version": manifest.version,
                "generated_at": manifest.generated_at,
                "files": [
                    {
                        "path": entry.path,
                        "sha256": entry.sha256,
                        "size": entry.size,
                        "asset_name": entry.asset_name,
                    }
                    for entry in manifest.files
                ],
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    for entry in changed_files:
        asset = asset_map.get(entry.asset_name)
        if entry.size == 0 or not entry.asset_name:
            continue
        if asset is None:
            raise RuntimeError(f"Update asset missing from the release: {entry.asset_name}")
        destination = download_root / Path(*entry.path.split("/"))
        _download_binary(asset.url, destination, timeout_seconds=timeout_seconds)

    plan_path.write_text(
        json.dumps(
            {
                "version": manifest.version,
                "changed_files": [entry.path for entry in changed_files],
                "stale_paths": list(stale_paths),
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return PreparedFileUpdate(
        stage_root=stage_root,
        changed_count=len(changed_files),
        stale_count=len(stale_paths),
        update_log_path=update_log_path,
    )


def stage_windows_update(zip_path: str | Path) -> Path:
    if not is_frozen():
        raise RuntimeError("Automatic update install only works from the packaged app build.")

    zip_path = Path(zip_path).resolve()
    executable = Path(sys.executable).resolve()
    install_root = executable.parent.parent
    update_log_path = user_data_dir() / "update.log"

    stage_root = Path(tempfile.mkdtemp(prefix="mu-unscramble-apply-"))
    script_path = stage_root / "apply_update.ps1"
    script_path.write_text(
        _build_apply_update_script(
            current_pid=os.getpid(),
            zip_path=zip_path,
            install_root=install_root,
            executable_name=executable.name,
            update_log_path=update_log_path,
        ),
        encoding="utf-8",
    )

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
            str(script_path),
        ],
        creationflags=creationflags,
    )
    return update_log_path


def stage_windows_file_update(prepared: PreparedFileUpdate) -> Path:
    if not is_frozen():
        raise RuntimeError("Automatic file updates only work from the packaged app build.")

    executable = Path(sys.executable).resolve()
    install_root = executable.parent.parent
    script_path = prepared.stage_root / "apply_file_update.ps1"
    script_path.write_text(
        _build_apply_file_update_script(
            current_pid=os.getpid(),
            install_root=install_root,
            executable_name=executable.name,
            stage_root=prepared.stage_root,
            update_log_path=prepared.update_log_path,
        ),
        encoding="utf-8",
    )

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
            str(script_path),
        ],
        creationflags=creationflags,
    )
    return prepared.update_log_path


def _extract_release_assets(payload: dict[str, object]) -> tuple[ReleaseAsset, ...]:
    raw_assets = payload.get("assets", [])
    if not isinstance(raw_assets, list):
        return ()

    assets: list[ReleaseAsset] = []
    for asset in raw_assets:
        if not isinstance(asset, dict):
            continue
        name = str(asset.get("name", "") or "").strip()
        url = str(asset.get("browser_download_url", "") or "").strip()
        size_value = asset.get("size", 0)
        try:
            size = max(0, int(size_value))
        except Exception:
            size = 0
        if not name or not url:
            continue
        assets.append(ReleaseAsset(name=name, url=url, size=size))
    return tuple(assets)


def _pick_release_asset(assets: tuple[ReleaseAsset, ...]) -> tuple[str, str]:
    preferred: tuple[str, str] = ("", "")
    fallback: tuple[str, str] = ("", "")
    for asset in assets:
        name = asset.name
        url = asset.url
        if not name or not url:
            continue
        lowered = name.lower()
        if lowered.endswith(".zip") and "win64" in lowered:
            return name, url
        if lowered.endswith(".zip") and not fallback[0]:
            fallback = (name, url)
        if not preferred[0]:
            preferred = (name, url)
    return fallback if fallback[0] else preferred


def _pick_manifest_asset(assets: tuple[ReleaseAsset, ...]) -> tuple[str, str]:
    for asset in assets:
        lowered = asset.name.lower()
        if lowered.endswith("-update-manifest.json") or lowered.endswith("update-manifest.json"):
            return asset.name, asset.url
    return "", ""


def _build_apply_update_script(
    *,
    current_pid: int,
    zip_path: Path,
    install_root: Path,
    executable_name: str,
    update_log_path: Path,
) -> str:
    zip_literal = str(zip_path)
    install_root_literal = str(install_root)
    executable_literal = str(executable_name)
    update_log_literal = str(update_log_path)
    return textwrap.dedent(
        f"""
        $ErrorActionPreference = "Stop"
        $targetPid = {current_pid}
        $zipPath = "{zip_literal}"
        $installRoot = "{install_root_literal}"
        $executableName = "{executable_literal}"
        $updateLogPath = "{update_log_literal}"
        $stageRoot = Split-Path -Parent $PSCommandPath
        $extractRoot = Join-Path $stageRoot "extracted"
        $bundleSource = Join-Path $extractRoot "MU Unscramble Bot"
        $bundleTarget = Join-Path $installRoot "MU Unscramble Bot"
        $launcherSource = Join-Path $extractRoot "Start MU Unscramble Bot.vbs"
        $launcherTarget = Join-Path $installRoot "Start MU Unscramble Bot.vbs"
        $relaunchPath = Join-Path $bundleTarget $executableName

        function Write-UpdateLog([string]$message) {{
            $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $line = "[{0}] {1}" -f $stamp, $message
            New-Item -ItemType Directory -Path (Split-Path -Parent $updateLogPath) -Force | Out-Null
            Add-Content -LiteralPath $updateLogPath -Value $line
        }}

        function Invoke-WithRetry([scriptblock]$operation, [string]$description) {{
            $lastError = $null
            for ($attempt = 1; $attempt -le 10; $attempt++) {{
                try {{
                    & $operation
                    Write-UpdateLog "$description succeeded on attempt $attempt."
                    return
                }} catch {{
                    $lastError = $_
                    Write-UpdateLog "$description failed on attempt $attempt: $($_.Exception.Message)"
                    Start-Sleep -Milliseconds 800
                }}
            }}
            throw $lastError
        }}

        Write-UpdateLog "Starting staged update from $zipPath into $installRoot."

        while (Get-Process -Id $targetPid -ErrorAction SilentlyContinue) {{
            Start-Sleep -Milliseconds 500
        }}
        Start-Sleep -Seconds 2

        try {{
            if (Test-Path -LiteralPath $extractRoot) {{
                Remove-Item -LiteralPath $extractRoot -Recurse -Force
            }}
            New-Item -ItemType Directory -Path $extractRoot | Out-Null
            Write-UpdateLog "Extracting release archive."
            Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force

            Invoke-WithRetry {{
                if (Test-Path -LiteralPath $bundleTarget) {{
                    Remove-Item -LiteralPath $bundleTarget -Recurse -Force
                }}
            }} "Removing old app bundle"

            Invoke-WithRetry {{
                Copy-Item -LiteralPath $bundleSource -Destination $bundleTarget -Recurse -Force
            }} "Copying new app bundle"

            if (Test-Path -LiteralPath $launcherSource) {{
                Invoke-WithRetry {{
                    Copy-Item -LiteralPath $launcherSource -Destination $launcherTarget -Force
                }} "Copying launcher script"
            }}

            if (Test-Path -LiteralPath $zipPath) {{
                Remove-Item -LiteralPath $zipPath -Force
            }}
            if (Test-Path -LiteralPath $extractRoot) {{
                Remove-Item -LiteralPath $extractRoot -Recurse -Force
            }}

            Write-UpdateLog "Relaunching updated executable from $relaunchPath."
            Start-Process -FilePath $relaunchPath
            Start-Sleep -Seconds 1
            Write-UpdateLog "Update finished successfully."
        }} catch {{
            Write-UpdateLog "Update failed: $($_.Exception.Message)"
            throw
        }} finally {{
            Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
        }}
        """
    ).strip()


def _build_apply_file_update_script(
    *,
    current_pid: int,
    install_root: Path,
    executable_name: str,
    stage_root: Path,
    update_log_path: Path,
) -> str:
    install_root_literal = str(install_root)
    executable_literal = str(executable_name)
    stage_root_literal = str(stage_root)
    update_log_literal = str(update_log_path)
    return textwrap.dedent(
        f"""
        $ErrorActionPreference = "Stop"
        $targetPid = {current_pid}
        $installRoot = "{install_root_literal}"
        $executableName = "{executable_literal}"
        $stageRoot = "{stage_root_literal}"
        $updateLogPath = "{update_log_literal}"
        $filesRoot = Join-Path $stageRoot "files"
        $planPath = Join-Path $stageRoot "file-update-plan.json"
        $manifestPath = Join-Path $stageRoot "update-manifest.json"
        $bundleTarget = Join-Path $installRoot "MU Unscramble Bot"
        $relaunchPath = Join-Path $bundleTarget $executableName

        function Write-UpdateLog([string]$message) {{
            $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $line = "[{0}] {1}" -f $stamp, $message
            New-Item -ItemType Directory -Path (Split-Path -Parent $updateLogPath) -Force | Out-Null
            Add-Content -LiteralPath $updateLogPath -Value $line
        }}

        function To-SystemPath([string]$relativePath) {{
            return Join-Path $installRoot ($relativePath -replace '/', '\\')
        }}

        function To-StagedPath([string]$relativePath) {{
            return Join-Path $filesRoot ($relativePath -replace '/', '\\')
        }}

        while (Get-Process -Id $targetPid -ErrorAction SilentlyContinue) {{
            Start-Sleep -Milliseconds 500
        }}
        Start-Sleep -Seconds 2

        try {{
            Write-UpdateLog "Applying file update from $stageRoot."
            $plan = Get-Content -LiteralPath $planPath -Raw | ConvertFrom-Json

            foreach ($relativePath in $plan.stale_paths) {{
                $targetPath = To-SystemPath $relativePath
                if (Test-Path -LiteralPath $targetPath) {{
                    Remove-Item -LiteralPath $targetPath -Force
                    Write-UpdateLog "Removed stale file: $relativePath"
                }}
            }}

            foreach ($relativePath in $plan.changed_files) {{
                $sourcePath = To-StagedPath $relativePath
                $targetPath = To-SystemPath $relativePath
                $targetParent = Split-Path -Parent $targetPath
                if (-not (Test-Path -LiteralPath $targetParent)) {{
                    New-Item -ItemType Directory -Path $targetParent -Force | Out-Null
                }}
                if (Test-Path -LiteralPath $sourcePath) {{
                    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
                    Write-UpdateLog "Updated file: $relativePath"
                }} else {{
                    Set-Content -LiteralPath $targetPath -Value $null -NoNewline
                    Write-UpdateLog "Created empty file: $relativePath"
                }}
            }}

            if (Test-Path -LiteralPath $manifestPath) {{
                Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $installRoot "update-manifest.json") -Force
            }}

            if (Test-Path -LiteralPath $bundleTarget) {{
                Get-ChildItem -LiteralPath $bundleTarget -Directory -Recurse |
                    Sort-Object FullName -Descending |
                    ForEach-Object {{
                        if (-not (Get-ChildItem -LiteralPath $_.FullName -Force)) {{
                            Remove-Item -LiteralPath $_.FullName -Force
                        }}
                    }}
            }}

            Write-UpdateLog "Relaunching updated executable from $relaunchPath."
            Start-Process -FilePath $relaunchPath
            Start-Sleep -Seconds 1
            Write-UpdateLog "File update finished successfully."
        }} catch {{
            Write-UpdateLog "File update failed: $($_.Exception.Message)"
            throw
        }} finally {{
            Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue
        }}
        """
    ).strip()


def _download_json(url: str, *, timeout_seconds: float) -> dict[str, object]:
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/json",
            "User-Agent": "mu-unscramble-bot",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
        payload = json.loads(response.read().decode("utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("Expected a JSON object from the update server.")
    return payload


def _download_binary(url: str, destination: Path, *, timeout_seconds: float) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/octet-stream",
            "User-Agent": "mu-unscramble-bot",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout_seconds) as response, destination.open("wb") as handle:
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            handle.write(chunk)


def _list_managed_files(install_root: Path) -> dict[str, Path]:
    managed: dict[str, Path] = {}
    launcher = install_root / "Start MU Unscramble Bot.vbs"
    if launcher.exists():
        managed["Start MU Unscramble Bot.vbs"] = launcher

    bundle_root = install_root / "MU Unscramble Bot"
    if bundle_root.exists():
        for path in bundle_root.rglob("*"):
            if not path.is_file():
                continue
            managed[path.relative_to(install_root).as_posix()] = path

    local_manifest = install_root / "update-manifest.json"
    if local_manifest.exists():
        managed["update-manifest.json"] = local_manifest
    return managed


def _file_matches_manifest(path: Path, entry: UpdateManifestFile) -> bool:
    try:
        stat = path.stat()
    except FileNotFoundError:
        return False
    if stat.st_size != entry.size:
        return False
    return _sha256_file(path) == entry.sha256


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _safe_version(value: str) -> Version | None:
    try:
        return Version(value)
    except InvalidVersion:
        return None
