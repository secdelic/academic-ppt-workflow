from __future__ import annotations

import hashlib
import json
import posixpath
import re
import shutil
import subprocess
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping
from xml.etree import ElementTree as ET

from .utils import sha256_file, write_json


class IncrementalUpdateError(RuntimeError):
    """Raised when a targeted update cannot be performed without ambiguity."""


OPERATION_ACTIONS = frozenset({"KEEP", "REPLACE", "INSERT_AFTER"})
_PRESENTATION_NS = (
    "http://schemas.openxmlformats.org/presentationml/2006/main"
)
_CONTROL_CHARACTER_RE = re.compile(r"[\x00-\x1f\x7f]")
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass(frozen=True)
class UpdateTarget:
    slide_id: str
    target_index: int
    candidate_index: int
    section_id: str
    figure_ids: tuple[str, ...]
    previous_content_hash: str
    candidate_content_hash: str
    candidate_revision: int


def _normalise_selector(value: str | None) -> str:
    return "" if value is None else str(value).strip()


def resolve_update_slide_ids(
    deck_ir: dict[str, Any],
    *,
    update_slide: str | None = None,
    update_section: str | None = None,
    update_figure: str | None = None,
) -> list[str]:
    selectors = [
        ("slide", _normalise_selector(update_slide)),
        ("section", _normalise_selector(update_section)),
        ("figure", _normalise_selector(update_figure)),
    ]
    chosen = [(kind, value) for kind, value in selectors if value]
    if len(chosen) != 1:
        raise IncrementalUpdateError(
            "Exactly one of --update-slide, --update-section, or --update-figure is required"
        )
    slides = deck_ir.get("slides")
    if not isinstance(slides, list):
        raise IncrementalUpdateError("Deck IR has no slides list")
    kind, value = chosen[0]
    matches: list[str] = []
    for slide in slides:
        if not isinstance(slide, dict):
            continue
        slide_id = str(slide.get("slide_id", ""))
        if kind == "slide" and slide_id == value:
            matches.append(slide_id)
        elif kind == "section" and str(slide.get("section_id", "")) == value:
            matches.append(slide_id)
        elif kind == "figure":
            figure_ids = slide.get("figure_ids") or []
            if isinstance(figure_ids, list) and value in {str(item) for item in figure_ids}:
                matches.append(slide_id)
    matches = [matched for matched in dict.fromkeys(matches) if matched]
    if not matches:
        raise IncrementalUpdateError(f"No Deck IR slide matches {kind} selector {value!r}")
    return matches


def _slide_index_by_id(deck_ir: dict[str, Any]) -> dict[str, int]:
    slides = deck_ir.get("slides")
    if not isinstance(slides, list):
        raise IncrementalUpdateError("Deck IR has no slides list")
    result: dict[str, int] = {}
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            raise IncrementalUpdateError(f"Deck IR slide {index} is not an object")
        slide_id = str(slide.get("slide_id", "")).strip()
        if not slide_id:
            raise IncrementalUpdateError(f"Deck IR slide {index} has no slide_id")
        if slide_id in result:
            raise IncrementalUpdateError(f"Duplicate slide_id in Deck IR: {slide_id}")
        result[slide_id] = index
    return result


def build_update_targets(
    previous_ir: dict[str, Any],
    candidate_ir: dict[str, Any],
    slide_ids: list[str],
) -> list[UpdateTarget]:
    previous_deck_id = str(previous_ir.get("deck_id", "")).strip()
    candidate_deck_id = str(candidate_ir.get("deck_id", "")).strip()
    if not previous_deck_id or not candidate_deck_id:
        raise IncrementalUpdateError(
            "Both previous and candidate Deck IR require a deck_id"
        )
    if previous_deck_id != candidate_deck_id:
        raise IncrementalUpdateError(
            "Candidate Deck IR belongs to a different deck_id"
        )
    if not slide_ids:
        raise IncrementalUpdateError("No semantic slide IDs were selected")
    if len(slide_ids) != len(set(slide_ids)):
        raise IncrementalUpdateError("Duplicate slide IDs were selected")

    previous_indexes = _slide_index_by_id(previous_ir)
    candidate_indexes = _slide_index_by_id(candidate_ir)
    previous_slides = {
        str(slide["slide_id"]): slide
        for slide in previous_ir.get("slides", [])
        if isinstance(slide, dict) and slide.get("slide_id")
    }
    candidate_slides = {
        str(slide["slide_id"]): slide
        for slide in candidate_ir.get("slides", [])
        if isinstance(slide, dict) and slide.get("slide_id")
    }
    targets: list[UpdateTarget] = []
    for slide_id in slide_ids:
        if slide_id not in previous_indexes:
            raise IncrementalUpdateError(f"Target deck does not contain slide_id {slide_id}")
        if slide_id not in candidate_indexes:
            raise IncrementalUpdateError(f"Candidate deck does not contain slide_id {slide_id}")
        previous = previous_slides[slide_id]
        candidate = candidate_slides[slide_id]
        try:
            previous_revision = int(previous.get("slide_revision", 1))
            candidate_revision = int(candidate.get("slide_revision", 1))
        except (TypeError, ValueError) as exc:
            raise IncrementalUpdateError(
                f"Invalid slide_revision for {slide_id}"
            ) from exc
        if previous_revision < 1 or candidate_revision < 1:
            raise IncrementalUpdateError(
                f"Invalid slide_revision for {slide_id}"
            )
        previous_hash = str(previous.get("content_hash", ""))
        candidate_hash = str(candidate.get("content_hash", ""))
        if previous_hash != candidate_hash and candidate_revision <= previous_revision:
            raise IncrementalUpdateError(
                f"Changed slide {slide_id} must increment slide_revision"
            )
        if candidate_revision < previous_revision:
            raise IncrementalUpdateError(
                f"Candidate revision regresses for {slide_id}"
            )
        targets.append(
            UpdateTarget(
                slide_id=slide_id,
                target_index=previous_indexes[slide_id],
                candidate_index=candidate_indexes[slide_id],
                section_id=str(candidate.get("section_id", "")),
                figure_ids=tuple(str(value) for value in candidate.get("figure_ids", []) or []),
                previous_content_hash=previous_hash,
                candidate_content_hash=candidate_hash,
                candidate_revision=candidate_revision,
            )
        )
    return targets


def backup_source_deck(source_pptx: Path, backup_dir: Path) -> dict[str, str]:
    source_pptx = source_pptx.resolve()
    if not source_pptx.is_file():
        raise IncrementalUpdateError(f"Existing PPTX is missing: {source_pptx}")
    if source_pptx.suffix.lower() != ".pptx":
        raise IncrementalUpdateError("Existing deck must be a .pptx file")
    backup_dir = backup_dir.resolve()
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir / f"{source_pptx.stem}.before_update{source_pptx.suffix}"
    if backup_path.exists():
        raise IncrementalUpdateError(f"Refusing to overwrite backup: {backup_path}")
    before = sha256_file(source_pptx)
    shutil.copy2(source_pptx, backup_path)
    after = sha256_file(backup_path)
    source_after = sha256_file(source_pptx)
    if before != source_after:
        raise IncrementalUpdateError("Source PPTX changed while creating backup")
    if before != after:
        raise IncrementalUpdateError("Backup hash mismatch")
    return {
        "source_path": str(source_pptx),
        "source_sha256": before,
        "source_sha256_after": source_after,
        "backup_path": str(backup_path),
        "backup_sha256": after,
    }


def _manifest_slide_rows(manifest: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not manifest:
        return []
    rows = manifest.get("slides")
    if not isinstance(rows, list):
        rows = manifest.get("object_manifest")
    if not isinstance(rows, list):
        return []
    return [item for item in rows if isinstance(item, dict)]


def _manifest_hash(item: dict[str, Any]) -> str:
    return str(
        item.get("managed_object_hash")
        or item.get("managed_object_sha256")
        or item.get("manifest_sha256")
        or ""
    )


def verify_no_manual_conflict(
    slide_ids: list[str],
    current_manifest: dict[str, Any] | None,
    previous_manifest: dict[str, Any] | None,
    *,
    allow_managed_overwrite: bool,
) -> list[str]:
    """Compare managed-object hashes when a prior manifest is available.

    The workflow deliberately does not guess when either manifest is missing.
    Unknown state or a detected change is blocked unless the caller supplies
    explicit managed-object overwrite permission.
    """

    warnings: list[str] = []
    if not previous_manifest or not current_manifest:
        message = (
            "MANUAL_EDIT_STATE_UNKNOWN: both current and previous object "
            "manifests are required to prove that managed objects were not "
            "changed outside the workflow."
        )
        if not allow_managed_overwrite:
            raise IncrementalUpdateError(message)
        warnings.append(message + " Explicit overwrite permission was supplied.")
        return warnings
    current_slides = {
        str(item.get("slide_id")): item
        for item in _manifest_slide_rows(current_manifest)
        if item.get("slide_id")
    }
    prior_slides = {
        str(item.get("slide_id")): item
        for item in _manifest_slide_rows(previous_manifest)
        if item.get("slide_id")
    }
    conflicts: list[str] = []
    unavailable: list[str] = []
    for slide_id in slide_ids:
        current = current_slides.get(slide_id)
        prior = prior_slides.get(slide_id)
        if not current or not prior:
            unavailable.append(slide_id)
            continue
        current_hash = _manifest_hash(current)
        prior_hash = _manifest_hash(prior)
        if not current_hash or not prior_hash:
            unavailable.append(slide_id)
            continue
        if current_hash != prior_hash:
            conflicts.append(slide_id)
    if unavailable and not allow_managed_overwrite:
        raise IncrementalUpdateError(
            "MANUAL_EDIT_STATE_UNKNOWN for: " + ", ".join(unavailable)
        )
    if unavailable:
        warnings.append(
            "Managed-object comparison unavailable; overwrite explicitly allowed for: "
            + ", ".join(unavailable)
        )
    if conflicts and not allow_managed_overwrite:
        raise IncrementalUpdateError(
            "MANUAL_EDIT_CONFLICT for managed objects on: " + ", ".join(conflicts)
        )
    if conflicts:
        warnings.append(
            "Managed-object overwrite explicitly allowed for: " + ", ".join(conflicts)
        )
    return warnings


def _validate_replacement_paths(
    existing_pptx: Path,
    candidate_pptx: Path,
    output_pptx: Path,
    script_path: Path,
) -> tuple[Path, Path, Path, Path]:
    existing = existing_pptx.resolve()
    candidate = candidate_pptx.resolve()
    output = output_pptx.resolve()
    script = script_path.resolve()
    for path, label, suffix in [
        (existing, "existing PPTX", ".pptx"),
        (candidate, "candidate PPTX", ".pptx"),
        (script, "PowerShell update script", ".ps1"),
    ]:
        if not path.is_file():
            raise IncrementalUpdateError(f"Missing {label}: {path}")
        if path.suffix.lower() != suffix:
            raise IncrementalUpdateError(f"{label} must use the {suffix} extension")
    if output.suffix.lower() != ".pptx":
        raise IncrementalUpdateError("Update output must be a .pptx file")
    if output.exists():
        raise IncrementalUpdateError(f"Refusing to overwrite update output: {output}")
    if existing == candidate:
        raise IncrementalUpdateError(
            "Existing and candidate PPTX paths must be different files"
        )
    if output in {existing, candidate, script}:
        raise IncrementalUpdateError("Output path must be distinct from every input")
    return existing, candidate, output, script


def run_powerpoint_replacement(
    *,
    existing_pptx: Path,
    candidate_pptx: Path,
    output_pptx: Path,
    targets: list[UpdateTarget],
    script_path: Path,
    timeout_seconds: int = 120,
) -> str:
    if not targets:
        raise IncrementalUpdateError("No update targets were supplied")
    if timeout_seconds <= 0:
        raise IncrementalUpdateError("timeout_seconds must be positive")
    slide_ids = [item.slide_id for item in targets]
    target_indexes = [item.target_index for item in targets]
    candidate_indexes = [item.candidate_index for item in targets]
    if len(slide_ids) != len(set(slide_ids)):
        raise IncrementalUpdateError("Duplicate slide_id values in update targets")
    if any(index < 1 for index in target_indexes + candidate_indexes):
        raise IncrementalUpdateError("Slide indexes must be positive")
    if len(target_indexes) != len(set(target_indexes)):
        raise IncrementalUpdateError("Duplicate target slide indexes")
    if len(candidate_indexes) != len(set(candidate_indexes)):
        raise IncrementalUpdateError("Duplicate candidate slide indexes")

    existing_pptx, candidate_pptx, output_pptx, script_path = (
        _validate_replacement_paths(
            existing_pptx, candidate_pptx, output_pptx, script_path
        )
    )
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    input_hashes = {
        existing_pptx: sha256_file(existing_pptx),
        candidate_pptx: sha256_file(candidate_pptx),
        script_path: sha256_file(script_path),
    }
    mapping_path = output_pptx.parent / "slide_update_mapping.json"
    if mapping_path.exists():
        raise IncrementalUpdateError(
            f"Refusing to overwrite update mapping: {mapping_path}"
        )
    shutil.copy2(existing_pptx, output_pptx)
    if sha256_file(output_pptx) != input_hashes[existing_pptx]:
        raise IncrementalUpdateError("Working-copy hash mismatch before replacement")
    write_json(
        mapping_path,
        {
            "schema_version": "1.0",
            "targets": [
                {
                    "slide_id": item.slide_id,
                    "target_index": item.target_index,
                    "candidate_index": item.candidate_index,
                    "slide_revision": item.candidate_revision,
                    "content_hash": item.candidate_content_hash,
                }
                for item in sorted(targets, key=lambda item: item.target_index, reverse=True)
            ],
        },
    )
    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-ExistingCopy",
        str(output_pptx),
        "-Candidate",
        str(candidate_pptx),
        "-Mapping",
        str(mapping_path),
    ]
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        changed_inputs = [
            str(path)
            for path, digest in input_hashes.items()
            if not path.is_file() or sha256_file(path) != digest
        ]
        if changed_inputs:
            raise IncrementalUpdateError(
                "Input changed during timed-out replacement: "
                + ", ".join(changed_inputs)
            ) from exc
        raise IncrementalUpdateError(
            f"PowerPoint incremental replacement timed out after {timeout_seconds}s"
        ) from exc
    changed_inputs = [
        str(path)
        for path, digest in input_hashes.items()
        if not path.is_file() or sha256_file(path) != digest
    ]
    if changed_inputs:
        raise IncrementalUpdateError(
            "Input changed during PowerPoint replacement: "
            + ", ".join(changed_inputs)
        )
    log = (completed.stdout + "\n" + completed.stderr).strip()
    if completed.returncode != 0:
        raise IncrementalUpdateError(
            f"PowerPoint incremental replacement failed ({completed.returncode}): {log}"
        )
    if not output_pptx.exists() or output_pptx.stat().st_size == 0:
        raise IncrementalUpdateError("PowerPoint incremental replacement produced no PPTX")
    return log


def write_slide_diff(
    path: Path,
    *,
    targets: list[UpdateTarget],
    source_hash_before: str,
    source_hash_after: str,
    untouched_differences: list[str],
    warnings: list[str],
) -> dict[str, Any]:
    report = {
        "schema_version": "1.0",
        "source_hash_before": source_hash_before,
        "source_hash_after": source_hash_after,
        "source_unchanged": source_hash_before == source_hash_after,
        "targets": [
            {
                "slide_id": item.slide_id,
                "target_index": item.target_index,
                "candidate_index": item.candidate_index,
                "previous_content_hash": item.previous_content_hash,
                "candidate_content_hash": item.candidate_content_hash,
                "content_changed": item.previous_content_hash != item.candidate_content_hash,
                "slide_revision": item.candidate_revision,
            }
            for item in targets
        ],
        "untouched_slide_differences": untouched_differences,
        "warnings": warnings,
        "manual_review_required": bool(warnings or untouched_differences),
    }
    write_json(path, report)
    return report


def _validate_semantic_slide_id(value: Any, *, label: str) -> str:
    slide_id = str(value or "").strip()
    if (
        not slide_id
        or len(slide_id) > 120
        or _CONTROL_CHARACTER_RE.search(slide_id)
    ):
        raise IncrementalUpdateError(f"Invalid semantic slide_id for {label}")
    return slide_id


def _is_unc_path(path: Path) -> bool:
    raw = str(path)
    return raw.startswith("\\\\") or raw.startswith("//")


def _local_input_file(path: Path, *, label: str, suffix: str) -> Path:
    resolved = path.resolve()
    if _is_unc_path(resolved):
        raise IncrementalUpdateError(f"{label} must be a local file, not a UNC path")
    if not resolved.is_file():
        raise IncrementalUpdateError(f"Missing {label}: {resolved}")
    if resolved.suffix.lower() != suffix:
        raise IncrementalUpdateError(f"{label} must use the {suffix} extension")
    return resolved


def _pptx_slide_count(path: Path) -> int:
    """Read only the presentation slide list; do not parse slide content."""

    try:
        with zipfile.ZipFile(path, "r") as package:
            root = ET.fromstring(package.read("ppt/presentation.xml"))
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise IncrementalUpdateError(
            f"Unreadable PPTX presentation structure: {path}: {exc}"
        ) from exc
    slide_list = root.find(f"{{{_PRESENTATION_NS}}}sldIdLst")
    if slide_list is None:
        return 0
    return len(slide_list.findall(f"{{{_PRESENTATION_NS}}}sldId"))


def _candidate_slide_workflow_contract(
    path: Path, candidate_index: int
) -> dict[str, str]:
    """Read workflow identity from one candidate slide's speaker notes.

    The contract is content-produced, not operation-plan metadata.  It binds
    the reviewed SlideSpec to the actual candidate page before PowerPoint can
    import it.  Legacy candidates without these managed markers fail closed.
    """

    try:
        with zipfile.ZipFile(path, "r") as package:
            presentation_rels = ET.fromstring(
                package.read("ppt/_rels/presentation.xml.rels")
            )
            presentation = ET.fromstring(package.read("ppt/presentation.xml"))
            rel_targets = {
                str(row.get("Id")): str(row.get("Target"))
                for row in presentation_rels.findall(
                    f"{{{_RELATIONSHIPS_NS}}}Relationship"
                )
            }
            slide_list = presentation.find(
                f"{{{_PRESENTATION_NS}}}sldIdLst"
            )
            if slide_list is None:
                raise IncrementalUpdateError("Candidate PPTX has no slides")
            slides = slide_list.findall(f"{{{_PRESENTATION_NS}}}sldId")
            if candidate_index < 1 or candidate_index > len(slides):
                raise IncrementalUpdateError("Candidate slide index is out of range")
            rel_id = slides[candidate_index - 1].get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            target = rel_targets.get(str(rel_id), "")
            if not target:
                raise IncrementalUpdateError("Candidate slide relationship is missing")
            slide_part = posixpath.normpath(
                posixpath.join("ppt", target.replace("\\", "/"))
            )
            slide_name = posixpath.basename(slide_part)
            rels_name = posixpath.join(
                posixpath.dirname(slide_part), "_rels", f"{slide_name}.rels"
            )
            slide_rels = ET.fromstring(package.read(rels_name))
            notes_target = ""
            for row in slide_rels.findall(
                f"{{{_RELATIONSHIPS_NS}}}Relationship"
            ):
                if str(row.get("Type", "")).endswith("/notesSlide"):
                    notes_target = str(row.get("Target", ""))
                    break
            if not notes_target:
                raise IncrementalUpdateError(
                    "Candidate slide has no managed speaker notes contract"
                )
            notes_name = posixpath.normpath(
                posixpath.join(
                    posixpath.dirname(slide_part),
                    notes_target.replace("\\", "/"),
                )
            )
            notes_root = ET.fromstring(package.read(notes_name))
            text = "\n".join(
                item.text or ""
                for item in notes_root.iter()
                if item.tag.endswith("}t")
            )
    except IncrementalUpdateError:
        raise
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise IncrementalUpdateError(
            "Candidate slide workflow contract is unreadable"
        ) from exc

    def marker(label: str) -> str:
        match = re.search(
            rf"\[{re.escape(label)}\]\s*\r?\n\s*([^\r\n]+)", text
        )
        return match.group(1).strip() if match else ""

    return {
        "slide_id": marker("Slide-ID"),
        "slide_revision": marker("Slide-Revision"),
        "content_hash": marker("Content-Hash").lower(),
    }


def semantic_order_sha256(slide_ids: Iterable[str]) -> str:
    normalized = [
        _validate_semantic_slide_id(item, label="expected final order")
        for item in slide_ids
    ]
    return hashlib.sha256("\n".join(normalized).encode("utf-8")).hexdigest()


def _operation_candidate_path(row: Mapping[str, Any]) -> Path:
    raw = row.get("candidate_pptx") or row.get("candidate_path")
    if not raw:
        raise IncrementalUpdateError(
            f"{row.get('action', 'operation')} for {row.get('slide_id', '[unknown]')} "
            "requires candidate_pptx"
        )
    path = Path(str(raw))
    if not path.is_absolute():
        raise IncrementalUpdateError("candidate_pptx must be an absolute local path")
    return _local_input_file(path, label="candidate PPTX", suffix=".pptx")


def build_operation_plan(
    *,
    source_pptx: Path,
    source_slide_ids: Iterable[str],
    operations: Iterable[Mapping[str, Any]],
    expected_final_order: Iterable[str],
    changed_slide_specs: Iterable[Mapping[str, Any]] | None = None,
    source_impacts: Iterable[Mapping[str, Any]] | None = None,
    source_binding_issues: Iterable[Mapping[str, Any]] | None = None,
    scientific_issues: Iterable[Mapping[str, Any]] | None = None,
    manual_review_items: Iterable[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Build a deterministic in-place-on-a-copy operation contract.

    The contract deliberately has no DELETE or MOVE operation.  Every source
    slide must remain in its original relative order and be classified as
    either KEEP or REPLACE; new semantic IDs may only be INSERT_AFTER.  This
    makes a small enhance operation auditable without reconstructing the deck.
    """

    source = _local_input_file(
        source_pptx, label="source PPTX", suffix=".pptx"
    )
    source_ids = [
        _validate_semantic_slide_id(item, label="source slide order")
        for item in source_slide_ids
    ]
    if not source_ids:
        raise IncrementalUpdateError("source_slide_ids must not be empty")
    if len(source_ids) != len(set(source_ids)):
        raise IncrementalUpdateError("source_slide_ids contains duplicates")
    actual_source_count = _pptx_slide_count(source)
    if actual_source_count != len(source_ids):
        raise IncrementalUpdateError(
            "source_slide_ids count does not match the source PPTX "
            f"({len(source_ids)} != {actual_source_count})"
        )

    expected_ids = [
        _validate_semantic_slide_id(item, label="expected final order")
        for item in expected_final_order
    ]
    if not expected_ids:
        raise IncrementalUpdateError("expected_final_order must not be empty")
    if len(expected_ids) != len(set(expected_ids)):
        raise IncrementalUpdateError("expected_final_order contains duplicates")

    raw_operations = list(operations)
    if not raw_operations:
        raise IncrementalUpdateError("operations must not be empty")
    operation_by_id: dict[str, Mapping[str, Any]] = {}
    for index, row in enumerate(raw_operations, start=1):
        if not isinstance(row, Mapping):
            raise IncrementalUpdateError(f"Operation {index} is not an object")
        slide_id = _validate_semantic_slide_id(
            row.get("slide_id"), label=f"operation {index}"
        )
        if slide_id in operation_by_id:
            raise IncrementalUpdateError(f"Duplicate operation slide_id: {slide_id}")
        operation_by_id[slide_id] = row
    if set(operation_by_id) != set(expected_ids):
        missing = sorted(set(expected_ids) - set(operation_by_id))
        unexpected = sorted(set(operation_by_id) - set(expected_ids))
        raise IncrementalUpdateError(
            "Operations must match expected_final_order exactly; "
            f"missing={missing}, unexpected={unexpected}"
        )

    source_positions = {slide_id: index for index, slide_id in enumerate(source_ids, 1)}
    source_set = set(source_ids)
    if [item for item in expected_ids if item in source_set] != source_ids:
        raise IncrementalUpdateError(
            "KEEP/REPLACE source slides must retain their original relative order"
        )

    candidate_id_by_path: dict[Path, str] = {}
    candidate_rows: list[dict[str, Any]] = []
    normalized_operations: list[dict[str, Any]] = []
    for final_index, slide_id in enumerate(expected_ids):
        row = operation_by_id[slide_id]
        action = str(row.get("action", "")).strip().upper()
        if action not in OPERATION_ACTIONS:
            raise IncrementalUpdateError(
                f"Unsupported operation action for {slide_id}: {action!r}"
            )
        normalized: dict[str, Any] = {
            "operation_id": f"OP-{final_index + 1:04d}",
            "action": action,
            "slide_id": slide_id,
        }
        if action in {"KEEP", "REPLACE"}:
            if slide_id not in source_set:
                raise IncrementalUpdateError(
                    f"{action} slide_id is not present in the source deck: {slide_id}"
                )
            expected_source_index = source_positions[slide_id]
            supplied_index = row.get("source_index", expected_source_index)
            try:
                source_index = int(supplied_index)
            except (TypeError, ValueError) as exc:
                raise IncrementalUpdateError(
                    f"Invalid source_index for {slide_id}"
                ) from exc
            if source_index != expected_source_index:
                raise IncrementalUpdateError(
                    f"source_index does not match source_slide_ids for {slide_id}"
                )
            normalized["source_index"] = source_index
            if action == "KEEP":
                lineage = str(row.get("presentation_lineage", "")).strip()
                if lineage:
                    if lineage != "PRESERVED_FROM_EXISTING_DECK":
                        raise IncrementalUpdateError(
                            f"Unsupported KEEP presentation lineage: {lineage}"
                        )
                    normalized["presentation_lineage"] = lineage
        else:
            if slide_id in source_set:
                raise IncrementalUpdateError(
                    f"INSERT_AFTER must introduce a new semantic slide_id: {slide_id}"
                )
            if final_index == 0:
                raise IncrementalUpdateError(
                    "INSERT_AFTER cannot create a slide before the first source slide"
                )
            expected_anchor = expected_ids[final_index - 1]
            anchor = _validate_semantic_slide_id(
                row.get("after_slide_id"), label=f"INSERT_AFTER {slide_id} anchor"
            )
            if anchor != expected_anchor:
                raise IncrementalUpdateError(
                    f"INSERT_AFTER {slide_id} must name its immediate final-order "
                    f"predecessor ({expected_anchor})"
                )
            normalized["after_slide_id"] = anchor

        if action in {"REPLACE", "INSERT_AFTER"}:
            candidate_path = _operation_candidate_path(row)
            candidate_count = _pptx_slide_count(candidate_path)
            try:
                candidate_index = int(row.get("candidate_index", 0))
            except (TypeError, ValueError) as exc:
                raise IncrementalUpdateError(
                    f"Invalid candidate_index for {slide_id}"
                ) from exc
            if candidate_index < 1 or candidate_index > candidate_count:
                raise IncrementalUpdateError(
                    f"candidate_index out of range for {slide_id}: {candidate_index}"
                )
            candidate_id = candidate_id_by_path.get(candidate_path)
            if candidate_id is None:
                candidate_id = f"CAND-{len(candidate_rows) + 1:03d}"
                candidate_id_by_path[candidate_path] = candidate_id
                candidate_rows.append(
                    {
                        "candidate_id": candidate_id,
                        "path": str(candidate_path),
                        "sha256": sha256_file(candidate_path),
                        "slide_count": candidate_count,
                    }
                )
            normalized.update(
                {
                    "candidate_id": candidate_id,
                    "candidate_index": candidate_index,
                }
            )
            for optional in ("slide_revision", "content_hash"):
                if row.get(optional) not in (None, ""):
                    normalized[optional] = row.get(optional)
        normalized_operations.append(normalized)

    classified_source_ids = {
        item["slide_id"]
        for item in normalized_operations
        if item["action"] in {"KEEP", "REPLACE"}
    }
    if classified_source_ids != source_set:
        missing = sorted(source_set - classified_source_ids)
        raise IncrementalUpdateError(
            f"Every source slide must be KEEP or REPLACE; missing={missing}"
        )

    changed_ids = [
        item["slide_id"]
        for item in normalized_operations
        if item["action"] != "KEEP"
    ]

    def normalize_scoped_rows(
        values: Iterable[Mapping[str, Any]] | None,
        *,
        label: str,
        require_changed_slide: bool,
    ) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for index, value in enumerate(values or (), start=1):
            if not isinstance(value, Mapping):
                raise IncrementalUpdateError(f"{label} {index} is not an object")
            row = dict(value)
            if "slide_id" in row or require_changed_slide:
                slide_id = _validate_semantic_slide_id(
                    row.get("slide_id"), label=f"{label} {index}"
                )
                if slide_id not in set(changed_ids):
                    raise IncrementalUpdateError(
                        f"{label} {index} is outside changed-slide scope: {slide_id}"
                    )
                row["slide_id"] = slide_id
            result.append(row)
        return result

    normalized_specs = normalize_scoped_rows(
        changed_slide_specs, label="changed_slide_specs", require_changed_slide=True
    )
    spec_ids = [str(row["slide_id"]) for row in normalized_specs]
    if len(spec_ids) != len(set(spec_ids)):
        raise IncrementalUpdateError("changed_slide_specs contains duplicate slide_id")
    if set(spec_ids) != set(changed_ids):
        raise IncrementalUpdateError(
            "changed_slide_specs must cover every REPLACE and INSERT_AFTER slide"
        )
    operation_by_changed_id = {
        str(row["slide_id"]): row
        for row in normalized_operations
        if row["action"] != "KEEP"
    }
    for row in normalized_specs:
        slide_id = str(row["slide_id"])
        bindings = row.get("source_bindings")
        if not isinstance(bindings, list) or not bindings:
            raise IncrementalUpdateError(
                f"Changed SlideSpec has no source_bindings: {slide_id}"
            )
        for binding_index, binding in enumerate(bindings, start=1):
            if not isinstance(binding, Mapping):
                raise IncrementalUpdateError(
                    f"Changed SlideSpec source binding {binding_index} is not an object: {slide_id}"
                )
            source_id = str(binding.get("source_id", "")).strip()
            source_file = str(binding.get("source_file", "")).strip()
            if not source_id and not source_file:
                raise IncrementalUpdateError(
                    f"Changed SlideSpec source binding has no source_id or source_file: {slide_id}"
                )

        scientific_status = str(row.get("scientific_qa_status", "")).strip().upper()
        scientific_review = row.get("scientific_review")
        if not scientific_status and isinstance(scientific_review, Mapping):
            scientific_status = str(scientific_review.get("status", "")).strip().upper()
        if scientific_status not in {"PASS", "PASSED"}:
            raise IncrementalUpdateError(
                f"Changed SlideSpec requires explicit scientific review PASS: {slide_id}"
            )

        operation = operation_by_changed_id[slide_id]
        spec_hash = str(row.get("content_hash", "")).strip().lower()
        operation_hash = str(operation.get("content_hash", "")).strip().lower()
        if not _SHA256_RE.fullmatch(spec_hash) or operation_hash != spec_hash:
            raise IncrementalUpdateError(
                f"Operation content_hash must match changed SlideSpec: {slide_id}"
            )
        try:
            spec_revision = int(row.get("slide_revision", 0))
            operation_revision = int(operation.get("slide_revision", 0))
        except (TypeError, ValueError) as exc:
            raise IncrementalUpdateError(
                f"Operation slide_revision must match changed SlideSpec: {slide_id}"
            ) from exc
        if spec_revision < 1 or operation_revision != spec_revision:
            raise IncrementalUpdateError(
                f"Operation slide_revision must match changed SlideSpec: {slide_id}"
            )
        candidate = candidate_rows[
            int(str(operation["candidate_id"]).removeprefix("CAND-")) - 1
        ]
        workflow_contract = _candidate_slide_workflow_contract(
            Path(str(candidate["path"])), int(operation["candidate_index"])
        )
        if workflow_contract["slide_id"] != slide_id:
            raise IncrementalUpdateError(
                f"Candidate slide_id does not match changed SlideSpec: {slide_id}"
            )
        if workflow_contract["content_hash"] != spec_hash:
            raise IncrementalUpdateError(
                f"Candidate content_hash does not match changed SlideSpec: {slide_id}"
            )
        if workflow_contract["slide_revision"] != str(spec_revision):
            raise IncrementalUpdateError(
                f"Candidate slide_revision does not match changed SlideSpec: {slide_id}"
            )
    normalized_impacts = normalize_scoped_rows(
        source_impacts, label="source_impacts", require_changed_slide=False
    )
    for index, row in enumerate(normalized_impacts, start=1):
        source_id = str(row.get("source_id", "")).strip()
        affected = row.get("affected_slide_ids")
        if not source_id or not isinstance(affected, list) or not affected:
            raise IncrementalUpdateError(
                f"source_impacts {index} requires source_id and affected_slide_ids"
            )
        normalized_affected = [
            _validate_semantic_slide_id(
                slide_id, label=f"source_impacts {index} affected_slide_ids"
            )
            for slide_id in affected
        ]
        if len(normalized_affected) != len(set(normalized_affected)):
            raise IncrementalUpdateError(
                f"source_impacts {index} contains duplicate affected_slide_ids"
            )
        outside = sorted(set(normalized_affected) - set(changed_ids))
        if outside:
            raise IncrementalUpdateError(
                f"source_impacts {index} is outside changed-slide scope: {outside}"
            )
        row["source_id"] = source_id
        row["affected_slide_ids"] = normalized_affected
    normalized_binding_issues = normalize_scoped_rows(
        source_binding_issues,
        label="source_binding_issues",
        require_changed_slide=True,
    )
    normalized_scientific_issues = normalize_scoped_rows(
        scientific_issues, label="scientific_issues", require_changed_slide=True
    )
    normalized_review_items = normalize_scoped_rows(
        manual_review_items,
        label="manual_review_items",
        require_changed_slide=False,
    )

    source_hash = sha256_file(source)
    return {
        "schema_version": "1.0",
        "plan_type": "fast_enhance_operations",
        "source_deck": {
            "path": str(source),
            "sha256": source_hash,
            "slide_count": actual_source_count,
        },
        "source_slide_ids": source_ids,
        "candidate_decks": candidate_rows,
        "operations": normalized_operations,
        "expected_final_order": expected_ids,
        "expected_final_order_sha256": semantic_order_sha256(expected_ids),
        "expected_final_slide_count": len(expected_ids),
        "changed_slide_ids": changed_ids,
        "changed_slide_specs": normalized_specs,
        "source_impacts": normalized_impacts,
        "source_binding_issues": normalized_binding_issues,
        "scientific_issues": normalized_scientific_issues,
        "manual_review_items": normalized_review_items,
        "input_integrity": {
            "source_sha256": source_hash,
            "candidate_sha256": {
                item["candidate_id"]: item["sha256"] for item in candidate_rows
            },
        },
    }


def validate_operation_plan(
    plan: Mapping[str, Any], *, verify_input_hashes: bool = True
) -> dict[str, Any]:
    """Validate a serialized operation plan and optionally recheck its inputs."""

    if not isinstance(plan, Mapping):
        raise IncrementalUpdateError("Operation plan must be an object")
    source = plan.get("source_deck")
    if not isinstance(source, Mapping):
        raise IncrementalUpdateError("Operation plan has no source_deck object")
    source_path = Path(str(source.get("path", "")))
    if verify_input_hashes:
        if (
            not source_path.is_file()
            or sha256_file(source_path) != str(source.get("sha256", ""))
        ):
            raise IncrementalUpdateError(
                "Source PPTX hash no longer matches operation plan"
            )
    source_ids = plan.get("source_slide_ids")
    operations = plan.get("operations")
    expected = plan.get("expected_final_order")
    if not isinstance(source_ids, list) or not isinstance(operations, list) or not isinstance(expected, list):
        raise IncrementalUpdateError(
            "Operation plan requires source_slide_ids, operations, and expected_final_order lists"
        )

    candidates = plan.get("candidate_decks")
    if not isinstance(candidates, list):
        raise IncrementalUpdateError("Operation plan candidate_decks must be a list")
    candidate_by_id: dict[str, Mapping[str, Any]] = {}
    for row in candidates:
        if not isinstance(row, Mapping) or not row.get("candidate_id"):
            raise IncrementalUpdateError("Invalid candidate_decks entry")
        candidate_id = str(row["candidate_id"])
        if candidate_id in candidate_by_id:
            raise IncrementalUpdateError(f"Duplicate candidate_id: {candidate_id}")
        candidate_by_id[candidate_id] = row
        if verify_input_hashes:
            candidate_path = Path(str(row.get("path", "")))
            if (
                not candidate_path.is_file()
                or sha256_file(candidate_path) != str(row.get("sha256", ""))
            ):
                raise IncrementalUpdateError(
                    f"Candidate PPTX hash no longer matches operation plan: {candidate_path}"
                )
    expanded_operations: list[dict[str, Any]] = []
    for row in operations:
        if not isinstance(row, Mapping):
            raise IncrementalUpdateError("Invalid operations entry")
        expanded = dict(row)
        if str(row.get("action", "")).upper() in {"REPLACE", "INSERT_AFTER"}:
            candidate = candidate_by_id.get(str(row.get("candidate_id", "")))
            if candidate is None:
                raise IncrementalUpdateError(
                    f"Unknown candidate_id for {row.get('slide_id')}: {row.get('candidate_id')}"
                )
            expanded["candidate_pptx"] = candidate.get("path")
        expanded_operations.append(expanded)

    rebuilt = build_operation_plan(
        source_pptx=source_path,
        source_slide_ids=source_ids,
        operations=expanded_operations,
        expected_final_order=expected,
        changed_slide_specs=plan.get("changed_slide_specs") or [],
        source_impacts=plan.get("source_impacts") or [],
        source_binding_issues=plan.get("source_binding_issues") or [],
        scientific_issues=plan.get("scientific_issues") or [],
        manual_review_items=plan.get("manual_review_items") or [],
    )
    invariant_keys = (
        "source_deck",
        "source_slide_ids",
        "candidate_decks",
        "operations",
        "expected_final_order",
        "expected_final_order_sha256",
        "expected_final_slide_count",
        "changed_slide_ids",
        "changed_slide_specs",
        "source_impacts",
        "source_binding_issues",
        "scientific_issues",
        "manual_review_items",
    )
    for key in invariant_keys:
        if rebuilt.get(key) != plan.get(key):
            raise IncrementalUpdateError(
                f"Serialized operation plan failed canonical validation at {key}"
            )
    return rebuilt


def write_operation_plan(path: Path, plan: Mapping[str, Any]) -> Path:
    destination = path.resolve()
    if destination.suffix.lower() != ".json":
        raise IncrementalUpdateError("Operation plan output must be a .json file")
    if destination.exists():
        raise IncrementalUpdateError(
            f"Refusing to overwrite operation plan: {destination}"
        )
    validated = validate_operation_plan(plan)
    write_json(destination, validated)
    return destination


def _parse_key_value_log(log: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for raw_line in log.splitlines():
        line = raw_line.strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().lower()
        if key and re.fullmatch(r"[a-z0-9_]+", key):
            result[key] = value.strip()
    return result


def run_powerpoint_operations(
    *,
    source_pptx: Path,
    output_pptx: Path,
    operation_plan_path: Path,
    script_path: Path,
    preview_dir: Path | None = None,
    export_pdf: Path | None = None,
    timeout_seconds: int = 180,
) -> dict[str, Any]:
    """Apply a validated plan through one PowerPoint COM application."""

    if timeout_seconds <= 0:
        raise IncrementalUpdateError("timeout_seconds must be positive")
    source = _local_input_file(source_pptx, label="source PPTX", suffix=".pptx")
    plan_path = _local_input_file(
        operation_plan_path, label="operation plan", suffix=".json"
    )
    script = _local_input_file(
        script_path, label="PowerShell operation script", suffix=".ps1"
    )
    output = output_pptx.resolve()
    if _is_unc_path(output):
        raise IncrementalUpdateError("Output PPTX must use a local path")
    if output.suffix.lower() != ".pptx":
        raise IncrementalUpdateError("Output must be a .pptx file")
    if output.exists():
        raise IncrementalUpdateError(f"Refusing to overwrite output: {output}")
    if output in {source, plan_path, script}:
        raise IncrementalUpdateError("Output path must be distinct from every input")
    preview: Path | None = None
    if preview_dir is not None:
        preview = preview_dir.resolve()
        if _is_unc_path(preview):
            raise IncrementalUpdateError("Preview directory must use a local path")
        if preview.exists():
            raise IncrementalUpdateError(
                f"Refusing to reuse preview directory: {preview}"
            )
    pdf: Path | None = None
    if export_pdf is not None:
        pdf = export_pdf.resolve()
        if _is_unc_path(pdf):
            raise IncrementalUpdateError("PDF output must use a local path")
        if pdf.suffix.lower() != ".pdf":
            raise IncrementalUpdateError("PDF output must use the .pdf extension")
        if pdf.exists():
            raise IncrementalUpdateError(f"Refusing to overwrite PDF output: {pdf}")
        if pdf in {source, plan_path, script, output}:
            raise IncrementalUpdateError("PDF output must be distinct from every input and PPTX output")
    try:
        plan_value = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise IncrementalUpdateError(f"Unreadable operation plan: {exc}") from exc
    plan = validate_operation_plan(plan_value)
    if Path(plan["source_deck"]["path"]).resolve() != source:
        raise IncrementalUpdateError(
            "Source PPTX argument does not match operation plan source_deck"
        )

    protected_paths = [source, plan_path, script] + [
        Path(row["path"]).resolve() for row in plan["candidate_decks"]
    ]
    input_hashes = {path: sha256_file(path) for path in protected_paths}
    output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-SourceDeck",
        str(source),
        "-OperationPlan",
        str(plan_path),
        "-OutputPptx",
        str(output),
    ]
    if preview is not None:
        command.extend(["-PreviewDir", str(preview)])
    if pdf is not None:
        command.extend(["-ExportPdf", str(pdf)])
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        changed_inputs = [
            str(path)
            for path, digest in input_hashes.items()
            if not path.is_file() or sha256_file(path) != digest
        ]
        if changed_inputs:
            raise IncrementalUpdateError(
                "Input changed during timed-out operation application: "
                + ", ".join(changed_inputs)
            ) from exc
        raise IncrementalUpdateError(
            f"PowerPoint operation plan timed out after {timeout_seconds}s"
        ) from exc

    changed_inputs = [
        str(path)
        for path, digest in input_hashes.items()
        if not path.is_file() or sha256_file(path) != digest
    ]
    if changed_inputs:
        raise IncrementalUpdateError(
            "Input changed during operation application: " + ", ".join(changed_inputs)
        )
    log = (completed.stdout + "\n" + completed.stderr).strip()
    if completed.returncode != 0:
        raise IncrementalUpdateError(
            f"PowerPoint operation plan failed ({completed.returncode}): {log}"
        )
    if not output.is_file() or output.stat().st_size == 0:
        raise IncrementalUpdateError("Operation plan produced no output PPTX")
    values = _parse_key_value_log(log)
    try:
        final_slide_count = int(values.get("final_slides", "0"))
    except ValueError as exc:
        raise IncrementalUpdateError("Invalid FINAL_SLIDES value from backend") from exc
    expected_count = int(plan["expected_final_slide_count"])
    if final_slide_count != expected_count or _pptx_slide_count(output) != expected_count:
        raise IncrementalUpdateError(
            "PowerPoint output slide count does not match operation plan"
        )
    final_order_hash = values.get("final_order_sha256", "").lower()
    expected_order_hash = str(plan["expected_final_order_sha256"]).lower()
    if final_order_hash != expected_order_hash:
        raise IncrementalUpdateError(
            "PowerPoint backend did not confirm the expected semantic slide order"
        )
    if values.get("powerpoint_open", "").upper() != "PASS":
        raise IncrementalUpdateError("PowerPoint backend did not confirm open/save success")
    changed_count = len(plan["changed_slide_ids"])
    try:
        applied_operations = int(values.get("applied_operations", "0"))
    except ValueError as exc:
        raise IncrementalUpdateError("Invalid APPLIED_OPERATIONS value from backend") from exc
    if applied_operations != changed_count:
        raise IncrementalUpdateError(
            "PowerPoint backend applied-operation count does not match changed slide count"
        )
    try:
        preview_count = int(values.get("changed_preview_count", "0"))
    except ValueError as exc:
        raise IncrementalUpdateError("Invalid CHANGED_PREVIEW_COUNT value from backend") from exc
    if preview is not None:
        previews = sorted(preview.glob("slide_*.png")) if preview.is_dir() else []
        if preview_count != changed_count or len(previews) != changed_count:
            raise IncrementalUpdateError(
                "PowerPoint backend did not export every changed-slide preview"
            )
    if pdf is not None:
        if values.get("pdf_export", "").upper() != "PASS" or not pdf.is_file() or pdf.stat().st_size == 0:
            raise IncrementalUpdateError("PowerPoint backend did not produce the requested PDF")
    return {
        "status": "PASS",
        "powerpoint_opened": True,
        "source_unchanged": sha256_file(source) == input_hashes[source],
        "input_hashes_unchanged": True,
        "applied_operations": applied_operations,
        "final_slide_count": final_slide_count,
        "final_order_sha256": final_order_hash,
        "output_pptx": str(output),
        "output_sha256": sha256_file(output),
        "changed_preview_count": preview_count,
        "preview_dir": str(preview) if preview is not None else "",
        "pdf_exported": pdf is not None,
        "pdf_path": str(pdf) if pdf is not None else "",
        "log": log,
    }
