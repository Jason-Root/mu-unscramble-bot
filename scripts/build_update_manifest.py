from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a file-by-file update manifest for the packaged app.")
    parser.add_argument("--release-root", required=True, help="Path to the packaged release folder.")
    parser.add_argument("--version", required=True, help="Release version without the leading v.")
    parser.add_argument("--manifest-output", required=True, help="Path to write the manifest JSON file.")
    parser.add_argument("--asset-output-dir", required=True, help="Directory to copy per-file update assets into.")
    return parser.parse_args()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def build_manifest(release_root: Path, version: str, manifest_output: Path, asset_output_dir: Path) -> int:
    release_root = release_root.resolve()
    manifest_output.parent.mkdir(parents=True, exist_ok=True)
    asset_output_dir.mkdir(parents=True, exist_ok=True)

    for existing in asset_output_dir.iterdir():
        if existing.is_dir():
            shutil.rmtree(existing)
        else:
            existing.unlink()

    files: list[dict[str, object]] = []
    for index, source_path in enumerate(sorted(path for path in release_root.rglob("*") if path.is_file()), start=1):
        relative_path = source_path.relative_to(release_root).as_posix()
        sha256 = sha256_file(source_path)
        asset_name = ""
        if source_path.stat().st_size > 0:
            asset_name = f"mu-update-{index:04d}-{sha256[:16]}-{source_path.name}"
            shutil.copy2(source_path, asset_output_dir / asset_name)
        files.append(
            {
                "path": relative_path,
                "sha256": sha256,
                "size": source_path.stat().st_size,
                "asset_name": asset_name,
            }
        )

    payload = {
        "version": version,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "files": files,
    }
    manifest_output.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return len(files)


def main() -> int:
    args = parse_args()
    count = build_manifest(
        release_root=Path(args.release_root),
        version=args.version,
        manifest_output=Path(args.manifest_output),
        asset_output_dir=Path(args.asset_output_dir),
    )
    print(f"Built update manifest with {count} files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
