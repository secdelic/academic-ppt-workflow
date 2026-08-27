from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping

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

SCIENTIFIC_SOURCE = "scientific_source"
PRESENTATION_BASELINE = "presentation_baseline"
STYLE_REFERENCE = "style_reference"
FORMAL_TEMPLATE = "formal_template"
APPROVED_VISUAL_ASSET = "approved_visual_asset"
NON_SCIENTIFIC_SOURCE_ROLES = frozenset(
    {
        PRESENTATION_BASELINE,
        STYLE_REFERENCE,
        FORMAL_TEMPLATE,
        APPROVED_VISUAL_ASSET,
    }
)


def classify_source_role(
    relative_path: str,
    *,
    route: str = "generate",
    reference_mode: str = "style-only",
) -> str:
    """Classify one registered input without changing its scientific meaning.

    ``content-reference`` is the existing explicit authority that permits an
    existing deck to participate in scientific evidence.  All other enhance
    baselines remain presentation lineage only.
    """

    relative = str(relative_path).strip().replace("\\", "/").lstrip("/")
    first = relative.partition("/")[0].casefold()
    normalized_route = str(route).strip().casefold().replace("_", "-")
    normalized_mode = str(reference_mode).strip().casefold()
    if first == "style_reference":
        return STYLE_REFERENCE
    if first == "template":
        return FORMAL_TEMPLATE
    if first == "approved_assets":
        return APPROVED_VISUAL_ASSET
    if first == "existing" and normalized_route in {"enhance", "enhance-existing"}:
        if normalized_mode == "content-reference":
            return SCIENTIFIC_SOURCE
        return PRESENTATION_BASELINE
    return SCIENTIFIC_SOURCE


def source_role(value: object) -> str:
    """Return a role with legacy registry rows treated as scientific."""

    if isinstance(value, Mapping):
        return str(value.get("source_role", SCIENTIFIC_SOURCE)).strip() or SCIENTIFIC_SOURCE
    return str(value or SCIENTIFIC_SOURCE).strip()


def is_scientific_source(value: object) -> bool:
    return source_role(value) == SCIENTIFIC_SOURCE


def scientific_delta_manifest(
    delta: Mapping[str, Any],
    *,
    route: str,
    reference_mode: str = "style-only",
) -> dict[str, Any]:
    """Project a full source delta onto the scientific lineage denominator."""

    entries = delta.get("entries", [])
    if not isinstance(entries, list):
        raise ValueError("Delta manifest entries must be a list")
    selected = [
        dict(row)
        for row in entries
        if isinstance(row, Mapping)
        and classify_source_role(
            str(row.get("source_key", "")),
            route=route,
            reference_mode=reference_mode,
        )
        == SCIENTIFIC_SOURCE
    ]
    statuses = ("MODIFIED", "NEW", "REMOVED", "UNCHANGED")
    return {
        "schema_version": "academic-ppt-scientific-delta-manifest/1",
        "hash_algorithm": str(delta.get("hash_algorithm", "sha256")),
        "lineage_scope": "SCIENTIFIC_SOURCE_ONLY",
        "entries": selected,
        "counts": {
            status: sum(str(row.get("status", "")).upper() == status for row in selected)
            for status in statuses
        },
        "parse_source_keys": [
            str(row.get("source_key", ""))
            for row in selected
            if str(row.get("status", "")).upper() in {"NEW", "MODIFIED"}
        ],
        "unchanged_source_keys": [
            str(row.get("source_key", ""))
            for row in selected
            if str(row.get("status", "")).upper() == "UNCHANGED"
        ],
        "removed_source_keys": [
            str(row.get("source_key", ""))
            for row in selected
            if str(row.get("status", "")).upper() == "REMOVED"
        ],
    }


def inventory_sources(
    input_root: Path,
    supported: set[str],
    manifest_path: Path,
    reference_mode: str = "style-only",
    route: str = "generate",
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
        classified_role = classify_source_role(
            relative_path,
            route=route,
            reference_mode=reference_mode,
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
                "source_role": classified_role,
                "reference_mode": (
                    reference_mode
                    if classified_role in {STYLE_REFERENCE, PRESENTATION_BASELINE}
                    else ""
                ),
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
