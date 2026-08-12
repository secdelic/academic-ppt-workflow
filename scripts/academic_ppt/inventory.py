from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from .utils import sha256_file, stable_source_id, utc_offset_timestamp, write_csv


MANIFEST_FIELDS = [
    "source_id",
    "file_name",
    "file_type",
    "relative_path",
    "sha256",
    "size_bytes",
    "modified_time",
    "parsed_successfully",
    "page_or_sheet_count",
    "parse_warning",
    "possible_sensitive_information",
    "source_role",
    "reference_mode",
]

SENSITIVE_PATTERN = re.compile(
    r"(patient[_ -]?id|medical record|mrn|identity card|phone|address|name[:：]|患者编号|住院号|身份证|手机号)",
    re.IGNORECASE,
)


def inventory_sources(
    input_root: Path,
    supported: set[str],
    manifest_path: Path,
    reference_mode: str = "style-only",
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    if not input_root.exists():
        write_csv(manifest_path, MANIFEST_FIELDS, rows)
        return rows
    for path in sorted(input_root.rglob("*")):
        relative_parts = {
            part.lower() for part in path.relative_to(input_root).parts
        }
        if "expected_ground_truth" in relative_parts:
            continue
        if not path.is_file() or path.suffix.lower() not in supported:
            continue
        digest = sha256_file(path)
        sample = path.read_bytes()[:2_000_000].decode("utf-8", errors="ignore")
        relative_path = path.relative_to(input_root).as_posix()
        source_role = (
            "style_reference"
            if relative_path.lower().startswith("style_reference/")
            else "scientific_source"
        )
        rows.append(
            {
                "source_id": stable_source_id(path, digest),
                "file_name": path.name,
                "file_type": path.suffix.lower().lstrip("."),
                "relative_path": relative_path,
                "sha256": digest,
                "size_bytes": str(path.stat().st_size),
                "modified_time": datetime.fromtimestamp(path.stat().st_mtime)
                .astimezone()
                .isoformat(timespec="seconds"),
                "parsed_successfully": "pending",
                "page_or_sheet_count": "",
                "parse_warning": "",
                "possible_sensitive_information": "yes" if SENSITIVE_PATTERN.search(sample) else "no",
                "source_role": source_role,
                "reference_mode": reference_mode if source_role == "style_reference" else "",
            }
        )
    write_csv(manifest_path, MANIFEST_FIELDS, rows)
    return rows


def verify_input_hashes(input_root: Path, rows: list[dict[str, str]]) -> list[str]:
    errors: list[str] = []
    for row in rows:
        path = input_root / row["relative_path"]
        if not path.exists():
            errors.append(f"Input disappeared during run: {row['relative_path']}")
        elif sha256_file(path) != row["sha256"]:
            errors.append(f"Input changed during run: {row['relative_path']}")
    return errors


def update_manifest(path: Path, rows: list[dict[str, str]]) -> None:
    write_csv(path, MANIFEST_FIELDS, rows)
