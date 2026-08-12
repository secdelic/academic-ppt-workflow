"""Scoped QA for the ``fast_enhance`` workflow mode.

This module intentionally does not render slides, start PowerPoint, invoke
LibreOffice, rebuild VisualSpec objects, or rerun full-deck scientific QA.  It
combines already-produced changed-slide scientific/source/PowerPoint geometry
results and performs a shallow, read-only OPC check of the final deck.
"""

from __future__ import annotations

import hashlib
import posixpath
import urllib.parse
import zipfile
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .utils import sha256_file


PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
R_ATTR = f"{{{DOC_REL_NS}}}id"
REL_SLIDE = DOC_REL_NS + "/slide"
REL_LAYOUT = DOC_REL_NS + "/slideLayout"
MEDIA_REL_SUFFIXES = {"image", "audio", "video", "media"}
FAIL_SEVERITIES = {"critical", "major", "error", "fail", "failed", "blocking"}


class FastQAError(ValueError):
    """Raised when a fast-QA contract is incomplete or unsafe to evaluate."""


def _slide_rows(slide_specs: Mapping[str, Any] | Sequence[Mapping[str, Any]]) -> list[Mapping[str, Any]]:
    if isinstance(slide_specs, Mapping):
        rows = slide_specs.get("slides", slide_specs.get("slide_specs"))
        if rows is None and all(isinstance(value, Mapping) for value in slide_specs.values()):
            rows = list(slide_specs.values())
    else:
        rows = slide_specs
    if not isinstance(rows, Sequence) or isinstance(rows, (str, bytes)):
        raise FastQAError("slide_specs must be a slide list or an object containing one")
    result = [row for row in rows if isinstance(row, Mapping)]
    if len(result) != len(rows):
        raise FastQAError("Every SlideSpec must be an object")
    return result


def _normalize_scope(changed_slide_ids: Iterable[str]) -> list[str]:
    scope = [str(item).strip() for item in changed_slide_ids]
    if not scope or any(not item for item in scope):
        raise FastQAError("changed_slide_ids must contain at least one non-empty ID")
    if len(scope) != len(set(scope)):
        raise FastQAError("changed_slide_ids contains duplicates")
    return scope


def _normalize_issue_rows(
    rows: Iterable[Mapping[str, Any]] | None,
    *,
    kind: str,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, row in enumerate(rows or (), start=1):
        if not isinstance(row, Mapping):
            raise FastQAError(f"{kind} issue {index} is not an object")
        slide_id = str(row.get("slide_id", "")).strip()
        if not slide_id:
            raise FastQAError(
                f"{kind} issue {index} is unscoped; changed-slide QA requires slide_id"
            )
        normalized = dict(row)
        normalized["slide_id"] = slide_id
        normalized["severity"] = str(row.get("severity", "error")).strip().lower()
        normalized["issue_type"] = kind
        result.append(normalized)
    return result


def _geometry_rows(report: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    rows = report.get("slides")
    if not isinstance(rows, list):
        raise FastQAError("powerpoint_geometry must contain a slides list")
    if any(not isinstance(row, Mapping) for row in rows):
        raise FastQAError("Every powerpoint_geometry slide row must be an object")
    return rows


def _geometry_findings(row: Mapping[str, Any]) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    raw_issues = row.get("issues") or row.get("findings") or []
    if isinstance(raw_issues, list):
        for raw in raw_issues:
            if isinstance(raw, Mapping):
                severity = str(raw.get("severity", "error")).strip().lower()
                findings.append({**dict(raw), "severity": severity})
            elif raw:
                findings.append(
                    {"code": "POWERPOINT_GEOMETRY_ISSUE", "severity": "error", "message": str(raw)}
                )
    for flag, code in (
        ("overflowing", "TEXT_OVERFLOW"),
        ("off_slide", "OFF_SLIDE_GEOMETRY"),
        ("prohibited_overlap", "PROHIBITED_OVERLAP"),
        ("clipping", "TEXT_OR_OBJECT_CLIPPING"),
        ("footer_intrusion", "FOOTER_INTRUSION"),
    ):
        if row.get(flag) is True:
            findings.append({"code": code, "severity": "error", "message": code})
    status = str(row.get("status", "PASS")).strip().upper()
    if status in {"FAIL", "FAILED", "BLOCKED", "ERROR"} and not findings:
        findings.append(
            {
                "code": "POWERPOINT_GEOMETRY_FAILED",
                "severity": "error",
                "message": f"PowerPoint geometry status is {status}",
            }
        )
    return findings


def run_changed_slide_qa(
    *,
    changed_slide_ids: Iterable[str],
    slide_specs: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    source_binding_issues: Iterable[Mapping[str, Any]] | None = None,
    scientific_issues: Iterable[Mapping[str, Any]] | None = None,
    powerpoint_geometry: Mapping[str, Any],
    typography_issues: Iterable[Mapping[str, Any]] | None = None,
    visual_issues: Iterable[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Evaluate only changed slides using upstream scoped QA evidence.

    The function fails closed when a changed slide has no SlideSpec, no required
    source binding, or no PowerPoint actual-geometry row.  Findings associated
    with unchanged slide IDs are retained only as ignored counts; they cannot
    block the changed-slide gate.
    """

    scope = _normalize_scope(changed_slide_ids)
    scope_set = set(scope)
    rows = _slide_rows(slide_specs)
    spec_by_id: dict[str, Mapping[str, Any]] = {}
    for row in rows:
        slide_id = str(row.get("slide_id", "")).strip()
        if not slide_id:
            raise FastQAError("Every SlideSpec requires slide_id")
        if slide_id in spec_by_id:
            raise FastQAError(f"Duplicate SlideSpec slide_id: {slide_id}")
        spec_by_id[slide_id] = row

    issues: list[dict[str, Any]] = []
    for slide_id in scope:
        spec = spec_by_id.get(slide_id)
        if spec is None:
            issues.append(
                {
                    "slide_id": slide_id,
                    "issue_type": "contract",
                    "code": "CHANGED_SLIDESPEC_MISSING",
                    "severity": "error",
                    "message": "Changed slide has no SlideSpec",
                }
            )
            continue
        binding_required = spec.get("source_binding_required") is not False
        bindings = spec.get("source_bindings") or []
        if binding_required and (not isinstance(bindings, list) or not bindings):
            issues.append(
                {
                    "slide_id": slide_id,
                    "issue_type": "source_binding",
                    "code": "SOURCE_BINDING_MISSING",
                    "severity": "error",
                    "message": "Changed slide requires at least one canonical source binding",
                }
            )

    source_rows = _normalize_issue_rows(
        source_binding_issues, kind="source_binding"
    )
    science_rows = _normalize_issue_rows(scientific_issues, kind="scientific")
    typography_rows = _normalize_issue_rows(
        typography_issues, kind="typography"
    )
    visual_rows = _normalize_issue_rows(visual_issues, kind="visual")
    issues.extend(
        row
        for row in source_rows + science_rows + typography_rows + visual_rows
        if row["slide_id"] in scope_set
    )

    if not isinstance(powerpoint_geometry, Mapping):
        raise FastQAError("powerpoint_geometry must be an object")
    geometry_by_id: dict[str, Mapping[str, Any]] = {}
    for row in _geometry_rows(powerpoint_geometry):
        slide_id = str(row.get("slide_id", "")).strip()
        if slide_id:
            if slide_id in geometry_by_id:
                raise FastQAError(f"Duplicate PowerPoint geometry slide_id: {slide_id}")
            geometry_by_id[slide_id] = row
    for slide_id in scope:
        geometry = geometry_by_id.get(slide_id)
        if geometry is None:
            issues.append(
                {
                    "slide_id": slide_id,
                    "issue_type": "powerpoint_geometry",
                    "code": "POWERPOINT_GEOMETRY_MISSING",
                    "severity": "error",
                    "message": "Changed slide has no PowerPoint actual-geometry result",
                }
            )
            continue
        for finding in _geometry_findings(geometry):
            issues.append(
                {
                    **finding,
                    "slide_id": slide_id,
                    "issue_type": "powerpoint_geometry",
                }
            )

    blocking = [
        row for row in issues if str(row.get("severity", "error")).lower() in FAIL_SEVERITIES
    ]
    return {
        "schema_version": "1.0",
        "qa_mode": "changed_slides_only",
        "changed_slide_ids": scope,
        "checked_slide_count": len(scope),
        "ignored_unchanged_source_issues": sum(
            row["slide_id"] not in scope_set for row in source_rows
        ),
        "ignored_unchanged_scientific_issues": sum(
            row["slide_id"] not in scope_set for row in science_rows
        ),
        "ignored_unchanged_typography_issues": sum(
            row["slide_id"] not in scope_set for row in typography_rows
        ),
        "ignored_unchanged_visual_issues": sum(
            row["slide_id"] not in scope_set for row in visual_rows
        ),
        "issues": issues,
        "blocking_issue_count": len(blocking),
        "status": "PASS" if not blocking else "FAIL",
        "draft_gate": "DRAFT_READY" if not blocking else "BLOCKED",
    }


def _safe_package_names(package: zipfile.ZipFile) -> set[str]:
    names: set[str] = set()
    for item in package.infolist():
        name = item.filename.replace("\\", "/")
        normalized = posixpath.normpath(name).lstrip("/")
        if (
            not name
            or name.startswith("/")
            or normalized in {"", ".", ".."}
            or normalized.startswith("../")
        ):
            raise FastQAError(f"Unsafe OOXML package member path: {name!r}")
        if normalized in names:
            raise FastQAError(f"Duplicate OOXML package member path: {normalized}")
        names.add(normalized)
    return names


def _source_for_rels(rels_part: str) -> str | None:
    if rels_part == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(rels_part)
    if not directory.endswith("/_rels") or not filename.endswith(".rels"):
        return None
    return posixpath.join(directory[: -len("/_rels")], filename[: -len(".rels")])


def _resolve_relationship_target(source_part: str, target: str) -> str | None:
    parsed = urllib.parse.urlsplit(target)
    if parsed.scheme or parsed.netloc:
        return None
    path = urllib.parse.unquote(parsed.path).replace("\\", "/")
    if path.startswith("/"):
        resolved = posixpath.normpath(path.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), path))
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        return None
    return resolved


def _relationships(
    package: zipfile.ZipFile, names: set[str], rels_part: str
) -> list[dict[str, str]]:
    source = _source_for_rels(rels_part)
    if source is None:
        raise FastQAError(f"Unrecognized relationships part: {rels_part}")
    try:
        root = ET.fromstring(package.read(rels_part))
    except (KeyError, ET.ParseError) as exc:
        raise FastQAError(f"Malformed relationships part {rels_part}: {exc}") from exc
    rows: list[dict[str, str]] = []
    seen_ids: set[str] = set()
    for relationship in root.findall(f"{{{REL_NS}}}Relationship"):
        relationship_id = str(relationship.get("Id", ""))
        if not relationship_id or relationship_id in seen_ids:
            raise FastQAError(f"Duplicate or empty relationship ID in {rels_part}")
        seen_ids.add(relationship_id)
        target_mode = str(relationship.get("TargetMode", "Internal"))
        target = str(relationship.get("Target", ""))
        rows.append(
            {
                "id": relationship_id,
                "type": str(relationship.get("Type", "")),
                "target_mode": target_mode,
                "target": target,
                "resolved_target": (
                    "" if target_mode == "External" else (_resolve_relationship_target(source, target) or "")
                ),
            }
        )
    return rows


def _slide_has_visible_object(root: ET.Element) -> bool:
    sp_tree = root.find(f".//{{{PML_NS}}}spTree")
    if sp_tree is None:
        return False
    visible_tags = {
        f"{{{PML_NS}}}sp",
        f"{{{PML_NS}}}pic",
        f"{{{PML_NS}}}graphicFrame",
        f"{{{PML_NS}}}cxnSp",
        f"{{{PML_NS}}}grpSp",
    }
    return any(child.tag in visible_tags for child in sp_tree)


def run_whole_deck_lightweight_qa(
    pptx_path: Path,
    *,
    expected_slide_count: int,
    expected_final_order: Iterable[str] | None = None,
    operation_result: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Check open/save evidence and shallow OOXML integrity without deep QA."""

    path = pptx_path.resolve()
    if not path.is_file() or path.suffix.lower() != ".pptx":
        raise FastQAError(f"PPTX does not exist: {path}")
    if expected_slide_count < 1:
        raise FastQAError("expected_slide_count must be positive")
    before_hash = sha256_file(path)
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    try:
        package = zipfile.ZipFile(path, "r")
    except (OSError, zipfile.BadZipFile) as exc:
        raise FastQAError(f"Unreadable PPTX package: {exc}") from exc

    with package:
        names = _safe_package_names(package)
        required = {"[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
        missing_required = sorted(required - names)
        if missing_required:
            raise FastQAError("Missing required OOXML parts: " + ", ".join(missing_required))
        try:
            presentation = ET.fromstring(package.read("ppt/presentation.xml"))
        except ET.ParseError as exc:
            raise FastQAError(f"Malformed ppt/presentation.xml: {exc}") from exc

        all_relationships: dict[str, list[dict[str, str]]] = {}
        for rels_part in sorted(name for name in names if name.endswith(".rels")):
            rows = _relationships(package, names, rels_part)
            source = _source_for_rels(rels_part)
            assert source is not None
            all_relationships[source] = rows
            for row in rows:
                if row["target_mode"] == "External":
                    warnings.append(
                        {
                            "code": "EXTERNAL_RELATIONSHIP_NOT_FOLLOWED",
                            "part": rels_part,
                            "relationship_id": row["id"],
                        }
                    )
                    continue
                target = row["resolved_target"]
                if not target or target not in names:
                    errors.append(
                        {
                            "code": "DANGLING_OR_UNSAFE_RELATIONSHIP",
                            "part": rels_part,
                            "relationship_id": row["id"],
                        }
                    )
                if row["type"].rsplit("/", 1)[-1] in MEDIA_REL_SUFFIXES and target not in names:
                    errors.append(
                        {
                            "code": "MISSING_MEDIA",
                            "part": rels_part,
                            "relationship_id": row["id"],
                        }
                    )

        presentation_rels = {
            row["id"]: row
            for row in all_relationships.get("ppt/presentation.xml", [])
            if row["type"] == REL_SLIDE and row["target_mode"] != "External"
        }
        slide_list = presentation.find(f"{{{PML_NS}}}sldIdLst")
        ordered_slide_parts: list[str] = []
        if slide_list is None:
            errors.append({"code": "SLIDE_LIST_MISSING", "part": "ppt/presentation.xml"})
        else:
            for slide_id in slide_list.findall(f"{{{PML_NS}}}sldId"):
                rel_id = str(slide_id.get(R_ATTR, ""))
                relationship = presentation_rels.get(rel_id)
                if relationship is None:
                    errors.append(
                        {"code": "SLIDE_ORDER_RELATIONSHIP_MISSING", "relationship_id": rel_id}
                    )
                else:
                    ordered_slide_parts.append(relationship["resolved_target"])
        if len(ordered_slide_parts) != expected_slide_count:
            errors.append(
                {
                    "code": "SLIDE_COUNT_MISMATCH",
                    "expected": expected_slide_count,
                    "actual": len(ordered_slide_parts),
                }
            )
        if len(ordered_slide_parts) != len(set(ordered_slide_parts)):
            errors.append({"code": "DUPLICATE_SLIDE_IN_ORDER"})

        blank_slide_parts: list[str] = []
        for slide_part in ordered_slide_parts:
            if slide_part not in names:
                continue
            try:
                slide_root = ET.fromstring(package.read(slide_part))
            except ET.ParseError:
                errors.append({"code": "MALFORMED_SLIDE_XML", "slide_part": slide_part})
                continue
            if not _slide_has_visible_object(slide_root):
                blank_slide_parts.append(slide_part)
                errors.append({"code": "OBVIOUS_BLANK_SLIDE", "slide_part": slide_part})
            layout_rows = [
                row
                for row in all_relationships.get(slide_part, [])
                if row["type"] == REL_LAYOUT and row["target_mode"] != "External"
            ]
            if len(layout_rows) != 1:
                errors.append(
                    {
                        "code": "SLIDE_LAYOUT_RELATIONSHIP_COUNT",
                        "slide_part": slide_part,
                        "actual": len(layout_rows),
                    }
                )

    expected_order_hash = ""
    if expected_final_order is not None:
        expected_ids = [str(item).strip() for item in expected_final_order]
        if not expected_ids or any(not item for item in expected_ids):
            raise FastQAError("expected_final_order contains an empty slide ID")
        expected_order_hash = hashlib.sha256("\n".join(expected_ids).encode("utf-8")).hexdigest()
        observed_hash = str((operation_result or {}).get("final_order_sha256", "")).lower()
        if observed_hash != expected_order_hash:
            errors.append(
                {
                    "code": "SEMANTIC_SLIDE_ORDER_NOT_CONFIRMED",
                    "expected_sha256": expected_order_hash,
                    "observed_sha256": observed_hash,
                }
            )
    if operation_result is not None:
        if operation_result.get("powerpoint_opened") is not True:
            errors.append({"code": "POWERPOINT_OPEN_NOT_CONFIRMED"})
        try:
            backend_count = int(operation_result.get("final_slide_count", 0))
        except (TypeError, ValueError):
            backend_count = 0
        if backend_count != expected_slide_count:
            errors.append(
                {
                    "code": "POWERPOINT_SLIDE_COUNT_NOT_CONFIRMED",
                    "expected": expected_slide_count,
                    "actual": backend_count,
                }
            )

    after_hash = sha256_file(path)
    if before_hash != after_hash:
        errors.append({"code": "QA_MUTATED_PPTX"})
    return {
        "schema_version": "1.0",
        "qa_mode": "whole_deck_lightweight",
        "pptx_sha256": after_hash,
        "input_unchanged": before_hash == after_hash,
        "expected_slide_count": expected_slide_count,
        "actual_slide_count": len(ordered_slide_parts),
        "ordered_slide_parts": ordered_slide_parts,
        "expected_final_order_sha256": expected_order_hash,
        "blank_slide_parts": blank_slide_parts,
        "errors": errors,
        "warnings": warnings,
        "status": "PASS" if not errors else "FAIL",
        "libreoffice_invoked": False,
        "deep_ooxml_invoked": False,
        "full_deck_textframe2_scan_invoked": False,
    }
