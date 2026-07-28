from __future__ import annotations

import csv
import hashlib
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


def load_yaml_compatible(path: Path) -> dict[str, Any]:
    """Load JSON-compatible YAML without requiring a third-party YAML parser."""
    text = path.read_text(encoding="utf-8-sig")
    try:
        value = json.loads(text)
    except json.JSONDecodeError:
        try:
            import yaml  # type: ignore
        except ImportError as exc:
            raise ValueError(
                f"{path} is not JSON-compatible YAML and PyYAML is unavailable"
            ) from exc
        value = yaml.safe_load(text)
    if not isinstance(value, dict):
        raise ValueError(f"Top-level mapping required: {path}")
    return value


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def write_csv(path: Path, fieldnames: list[str], rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def read_csv(path: Path, delimiter: str = ",") -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle, delimiter=delimiter))


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def stable_source_id(path: Path, digest: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9]+", "-", path.stem).strip("-").lower() or "source"
    return f"SRC-{stem[:24]}-{digest[:12]}"


def safe_slug(value: str) -> str:
    value = value.strip()
    value = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", value, flags=re.UNICODE)
    value = value.strip("._")
    return value[:80] or "unnamed_project"


def utc_offset_timestamp() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def output_timestamp() -> str:
    return datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")


def resolve_path(
    cli_value: str | None,
    env_name: str | None,
    config_value: str | None,
    repo_root: Path,
    default_value: str,
) -> Path:
    raw = cli_value or (os.environ.get(env_name) if env_name else None) or config_value or default_value
    path = Path(raw).expanduser()
    return path.resolve() if path.is_absolute() else (repo_root / path).resolve()


def fingerprint_paths(paths: Iterable[Path]) -> str:
    digest = hashlib.sha256()
    for path in sorted(paths, key=lambda p: str(p).lower()):
        digest.update(str(path).encode("utf-8"))
        if path.is_file():
            digest.update(sha256_file(path).encode("ascii"))
    return digest.hexdigest()


def ensure_new_directory(path: Path) -> None:
    if path.exists():
        raise FileExistsError(f"Refusing to overwrite existing output directory: {path}")
    path.mkdir(parents=True, exist_ok=False)
