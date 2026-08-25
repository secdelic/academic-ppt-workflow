from __future__ import annotations

import csv
import copy
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import traceback
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import Any, Mapping
from xml.etree import ElementTree as ET

from .deck_ir import (
    canonical_json_hash,
    to_storyboard_rows,
    validate_deck_ir,
)
from .evidence import build_evidence
from .extractors import extract_all
from .incremental import (
    backup_source_deck,
    build_update_targets,
    resolve_update_slide_ids,
    run_powerpoint_replacement,
    verify_no_manual_conflict,
    write_slide_diff,
)
from .inventory import inventory_sources, update_manifest
from .ooxml_qa import inspect_package
from .qa import run_qa, write_traceability_maps
from .rendering import create_contact_sheet, render_pptx
from .routing import RouteRequest, validate_route_request
from .storyboard import STORYBOARD_FIELDS, create_storyboard, write_speaker_notes
from .style_reference import parse_reference_style
from .template_fill import run_native_template_fill
from .utils import (
    ensure_new_directory,
    load_yaml_compatible,
    load_local_runtime_config,
    output_timestamp,
    read_csv,
    resolve_runtime_path,
    resolve_workflow_home,
    safe_slug,
    sha256_file,
    utc_offset_timestamp,
    write_csv,
    write_json,
)
from .visual_layout_qa import VisualLayoutQAError, inspect_powerpoint_text_layout
from .visual_density import write_density_reports


STAGES = [
    "preflight",
    "inventory",
    "extract",
    "evidence",
    "outline",
    "storyboard",
    "generate",
    "render",
    "ooxml_qa",
    "visual_qa",
    "scientific_qa",
    "package",
]


class WorkflowError(RuntimeError):
    pass


def _arg(args: Any, name: str, default: Any = None) -> Any:
    """Read a CLI attribute while preserving compatibility with v1 callers."""

    return getattr(args, name, default)


def _resolve_declared_file(
    raw: str | Path | None,
    *,
    repo_root: Path,
    input_root: Path,
    brief_path: Path,
) -> Path | None:
    if raw in {None, ""}:
        return None
    path = Path(raw).expanduser()
    if path.is_absolute():
        return path.resolve()
    candidates = [
        (repo_root / path).resolve(),
        (input_root / path).resolve(),
        (brief_path.parent / path).resolve(),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _first_brief_style_reference(brief: Mapping[str, Any]) -> str | Path | None:
    values = brief.get("style_reference_files", [])
    if isinstance(values, (str, Path)):
        return values
    if isinstance(values, list):
        return next((item for item in values if str(item).strip()), None)
    return None


def _load_json_mapping(path: Path, label: str) -> dict[str, Any]:
    if not path.is_file():
        raise WorkflowError(f"{label} is missing: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise WorkflowError(f"{label} is not valid JSON: {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise WorkflowError(f"{label} must contain a top-level mapping: {path}")
    return value


def _build_object_manifest_document(
    report: Mapping[str, Any],
    deck_ir: Mapping[str, Any],
) -> dict[str, Any]:
    """Bind privacy-preserving OOXML object hashes to stable slide IDs."""

    identity_rows = (
        report.get("ir_identity_check", {}).get("identity_map", [])
        if isinstance(report.get("ir_identity_check"), Mapping)
        else []
    )
    id_by_index = {
        int(row["slide_index"]): str(row["slide_id"])
        for row in identity_rows
        if isinstance(row, Mapping)
        and str(row.get("slide_index", "")).isdigit()
        and row.get("slide_id")
    }
    slides: list[dict[str, Any]] = []
    for row in report.get("object_manifest", []):
        if not isinstance(row, Mapping):
            continue
        try:
            slide_index = int(row.get("slide_index", 0))
        except (TypeError, ValueError):
            slide_index = 0
        slides.append(
            {
                "slide_id": id_by_index.get(slide_index, "UNRESOLVED"),
                "slide_index": slide_index,
                "slide_part": row.get("slide_part", "UNRESOLVED"),
                "managed_object_hash": row.get("manifest_sha256", "UNRESOLVED"),
                "object_kind_counts": row.get("object_kind_counts", {}),
                "object_count": row.get("object_count", 0),
                "chart_count": row.get("chart_count", 0),
                "table_count": row.get("table_count", 0),
                "picture_count": (
                    row.get("object_kind_counts", {}).get("picture", 0)
                    if isinstance(row.get("object_kind_counts"), Mapping)
                    else 0
                ),
            }
        )
    return {
        "schema_version": "2.0",
        "deck_id": deck_ir.get("deck_id", "UNKNOWN"),
        "deck_version": deck_ir.get("deck_version", "UNKNOWN"),
        "deck_ir_canonical_hash": deck_ir.get("canonical_hash", "UNKNOWN"),
        "source_pptx_sha256": report.get("source", {}).get("sha256", "UNKNOWN"),
        "slides": slides,
        "privacy_contract": {
            "slide_text_returned": False,
            "notes_text_returned": False,
            "object_hashes_only": True,
        },
    }


def _identity_hashes(report: Mapping[str, Any]) -> dict[str, str]:
    check = report.get("ir_identity_check", {})
    rows = check.get("identity_map", []) if isinstance(check, Mapping) else []
    return {
        str(row.get("slide_id")): str(row.get("ooxml_canonical_sha256"))
        for row in rows
        if isinstance(row, Mapping) and row.get("slide_id")
    }


def _pptx_slide_id_order(path: Path) -> list[str]:
    """Read only workflow-owned Slide-ID note markers in presentation order."""

    presentation_ns = (
        "http://schemas.openxmlformats.org/presentationml/2006/main"
    )
    drawing_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    relationship_ns = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    package_relationship_ns = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )

    def relationships(archive: zipfile.ZipFile, member: str) -> dict[str, dict[str, str]]:
        root = ET.fromstring(archive.read(member))
        return {
            str(element.get("Id")): {
                "type": str(element.get("Type", "")),
                "target": str(element.get("Target", "")),
                "mode": str(element.get("TargetMode", "")),
            }
            for element in root.findall(
                f"{{{package_relationship_ns}}}Relationship"
            )
            if element.get("Id")
        }

    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = set(archive.namelist())
            presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
            presentation_rels = relationships(
                archive, "ppt/_rels/presentation.xml.rels"
            )
            result: list[str] = []
            for slide_element in presentation.findall(
                f".//{{{presentation_ns}}}sldId"
            ):
                relationship_id = slide_element.get(
                    f"{{{relationship_ns}}}id", ""
                )
                relationship = presentation_rels.get(relationship_id)
                if not relationship or relationship["mode"].lower() == "external":
                    raise WorkflowError(
                        "Existing PPTX slide order has an unresolved relationship"
                    )
                slide_part = posixpath.normpath(
                    posixpath.join("ppt", relationship["target"])
                )
                slide_dir, slide_name = posixpath.split(slide_part)
                slide_rels_part = posixpath.join(
                    slide_dir, "_rels", slide_name + ".rels"
                )
                if slide_rels_part not in names:
                    raise WorkflowError(
                        f"Existing slide has no relationship part: {slide_part}"
                    )
                slide_rels = relationships(archive, slide_rels_part)
                notes_relationship = next(
                    (
                        value
                        for value in slide_rels.values()
                        if value["type"].endswith("/notesSlide")
                        and value["mode"].lower() != "external"
                    ),
                    None,
                )
                if notes_relationship is None:
                    raise WorkflowError(
                        f"Existing slide has no workflow notes identity: {slide_part}"
                    )
                notes_part = posixpath.normpath(
                    posixpath.join(
                        posixpath.dirname(slide_part),
                        notes_relationship["target"],
                    )
                )
                if notes_part not in names:
                    raise WorkflowError(
                        f"Existing slide notes are missing: {notes_part}"
                    )
                notes_root = ET.fromstring(archive.read(notes_part))
                notes_text = "\n".join(
                    element.text or ""
                    for element in notes_root.findall(
                        f".//{{{drawing_ns}}}t"
                    )
                )
                match = re.search(
                    r"\[Slide-ID\]\s*(SLD-[A-Za-z0-9-]+)",
                    notes_text,
                )
                if not match:
                    raise WorkflowError(
                        f"Existing slide lacks a stable [Slide-ID] marker: {slide_part}"
                    )
                result.append(match.group(1))
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise WorkflowError(
            f"Cannot verify stable slide identity in existing PPTX: {exc}"
        ) from exc
    if not result or len(result) != len(set(result)):
        raise WorkflowError(
            "Existing PPTX has missing or duplicate stable slide identities"
        )
    return result


def _reorder_deck_ir(
    deck_ir: Mapping[str, Any],
    slide_id_order: list[str],
    *,
    bump_version_for_reorder: bool = False,
) -> dict[str, Any]:
    prepared = copy.deepcopy(dict(deck_ir))
    slides = prepared.get("slides", [])
    by_id = {
        str(slide.get("slide_id")): slide
        for slide in slides
        if isinstance(slide, Mapping) and slide.get("slide_id")
    }
    if set(by_id) != set(slide_id_order) or len(by_id) != len(slide_id_order):
        raise WorkflowError(
            "Existing PPTX Slide-ID set does not match the previous Deck IR"
        )
    previous_order = [
        str(slide.get("slide_id")) for slide in slides if isinstance(slide, Mapping)
    ]
    prepared["slides"] = [copy.deepcopy(by_id[slide_id]) for slide_id in slide_id_order]
    if bump_version_for_reorder and previous_order != slide_id_order:
        prepared["deck_version"] = int(prepared.get("deck_version", 1)) + 1
    prepared.pop("canonical_hash", None)
    prepared["canonical_hash"] = canonical_json_hash(prepared)
    validate_deck_ir(prepared)
    return prepared


def _rewrite_storyboard_artifacts(
    staging_root: Path,
    deck_ir: Mapping[str, Any],
    project: str,
) -> list[dict[str, str]]:
    slides = to_storyboard_rows(deck_ir)
    write_json(staging_root / "deck_ir.json", deck_ir)
    write_csv(staging_root / "storyboard.csv", STORYBOARD_FIELDS, slides)
    outline = ["# Deck Outline", "", f"Project: {project}", ""]
    for slide in slides:
        outline.extend(
            [
                f"## {slide['slide_id']} - {slide['slide_title']}",
                "",
                f"- Purpose: {slide['slide_purpose']}",
                f"- Key message: {slide['single_key_message']}",
                f"- Sources: {slide['source_ids'] or 'INFORMATION_REQUIRED'}",
                f"- Layout: {slide['proposed_layout']}",
                "",
            ]
        )
    (staging_root / "deck_outline.md").write_text(
        "\n".join(outline), encoding="utf-8"
    )
    write_speaker_notes(staging_root, slides)
    return slides


def _log(log_path: Path, stage: str, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"{utc_offset_timestamp()}\t{stage}\t{message}\n")


def _copy_required(source: Path, destination: Path) -> None:
    if not source.exists():
        raise WorkflowError(f"Required artifact missing: {source}")
    shutil.copy2(source, destination)


def _detect_node(explicit: str | None) -> str:
    if explicit:
        if Path(explicit).exists():
            return str(Path(explicit).resolve())
        located = shutil.which(explicit)
        if located:
            return located
    located = shutil.which("node") or shutil.which("node.exe")
    if located:
        return located
    raise WorkflowError("Node.js unavailable; pass --node or set ACADEMIC_PPT_NODE")


def _run_generator(
    backend: str,
    node: str,
    node_modules: str | None,
    scripts_root: Path,
    spec: Path,
    pptx: Path,
    artifact_preview: Path,
    timeout: int,
) -> str:
    generator = (
        scripts_root / "generate_deck_artifact.mjs"
        if backend == "artifact_tool"
        else scripts_root / "generate_deck_pptxgen.mjs"
    )
    if not generator.exists():
        raise WorkflowError(f"Generator unavailable: {generator}")
    env = os.environ.copy()
    if node_modules:
        env["NODE_PATH"] = node_modules
    command = [node, str(generator), str(spec), str(pptx), str(artifact_preview)]
    result = subprocess.run(
        command,
        check=False,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=env,
    )
    if result.returncode != 0 or not pptx.exists() or pptx.stat().st_size == 0:
        raise WorkflowError(
            f"{backend} generation failed ({result.returncode}): {result.stdout}\n{result.stderr}"
        )
    return (result.stdout + "\n" + result.stderr).strip()


def _read_inventory_csv(path: Path) -> list[dict[str, str]]:
    return read_csv(path)


def _execute_create_style_profile(
    *,
    reference_path: Path,
    reference_mode: str,
    project: str,
    run_id: str,
    staging_root: Path,
    output_dir: Path,
) -> Path:
    """Execute the metadata-only route without invoking deck generation."""

    log_path = staging_root / "generation_log.md"
    state = {
        "status": "BLOCKED",
        "project_name": project,
        "run_id": run_id,
        "route": "create-style-profile",
        "backend": "read-only-ooxml",
        "started_at": utc_offset_timestamp(),
        "output_dir": str(output_dir),
        "network_access": False,
        "external_services": False,
        "stages": {},
    }
    write_json(staging_root / "run_state.json", state)
    before_hash = sha256_file(reference_path)
    try:
        _log(
            log_path,
            "preflight",
            f"route=create-style-profile; reference={reference_path}; mode={reference_mode}",
        )
        profile = parse_reference_style(reference_path, mode=reference_mode)
        if sha256_file(reference_path) != before_hash:
            raise WorkflowError("Reference file changed during style-only inspection")
        staging_profile = staging_root / "reference_style_profile.json"
        write_json(staging_profile, profile)
        shutil.copy2(staging_profile, output_dir / staging_profile.name)
        state["status"] = "READY_FOR_ASSISTED_USE"
        state["reference_sha256"] = before_hash
        state["reference_mode"] = reference_mode
        state["stages"] = {
            "preflight": "passed",
            "reference_style_import": "passed",
            "package": "passed",
        }
        state["finished_at"] = utc_offset_timestamp()
        write_json(output_dir / "runtime_manifest.json", state)
        write_json(staging_root / "run_state.json", state)
        _log(log_path, "package", f"status={state['status']}; output={output_dir}")
        shutil.copy2(log_path, output_dir / "generation_log.md")
        (output_dir / "dependency_manifest.txt").write_text(
            "Runtime dependencies used or declared\n"
            f"Python={sys.version.split()[0]}\n"
            "Reference parser=Python standard library OOXML (read-only)\n"
            "Network access=none\n"
            "External services=none\n",
            encoding="utf-8",
        )
        return output_dir
    except Exception as exc:
        state["status"] = "BLOCKED"
        state["finished_at"] = utc_offset_timestamp()
        state["error"] = f"{type(exc).__name__}: {exc}"
        state["traceback"] = traceback.format_exc()
        write_json(staging_root / "run_state.json", state)
        write_json(output_dir / "runtime_manifest.json", state)
        _log(log_path, "error", state["error"])
        shutil.copy2(log_path, output_dir / "generation_log.md")
        raise


def execute_full_validation(args: Any) -> Path:
    repo_root = resolve_workflow_home(
        _arg(args, "repo_root"), Path(__file__).resolve().parents[2]
    )
    workflow_config = load_yaml_compatible(repo_root / "config" / "workflow.yaml")
    local_cfg = load_local_runtime_config(repo_root)
    raw_brief = _arg(args, "brief", "brief/presentation_brief.yaml")
    brief_path = (
        Path(raw_brief).resolve()
        if Path(raw_brief).is_absolute()
        else (repo_root / raw_brief).resolve()
    )
    brief = load_yaml_compatible(brief_path)
    paths_cfg = workflow_config["paths"]
    env_cfg = workflow_config["environment"]
    workspace_home_raw = os.environ.get("PPT_WORKSPACE_HOME") or local_cfg.get("workspace_home")
    workspace_home = Path(workspace_home_raw).expanduser().resolve() if workspace_home_raw else None
    input_root = resolve_runtime_path(cli_value=_arg(args,"input_root"),env_name="PPT_INPUT_HOME",legacy_env_name=env_cfg["input_root"],local_value=local_cfg.get("input_root"),workspace_home=workspace_home,workspace_child="input",repo_root=repo_root,repo_default=paths_cfg["input_root"])
    staging_base = resolve_runtime_path(cli_value=_arg(args,"staging_root"),env_name="PPT_STAGING_HOME",legacy_env_name=env_cfg["staging_root"],local_value=local_cfg.get("staging_root"),workspace_home=workspace_home,workspace_child="staging",repo_root=repo_root,repo_default=paths_cfg["staging_root"])
    output_base = resolve_runtime_path(cli_value=_arg(args,"output_root"),env_name="PPT_OUTPUT_HOME",legacy_env_name=env_cfg["output_root"],local_value=local_cfg.get("output_root"),workspace_home=workspace_home,workspace_child="output",repo_root=repo_root,repo_default=paths_cfg["output_root"])
    audit_root = resolve_runtime_path(cli_value=_arg(args,"audit_root"),env_name="PPT_AUDIT_HOME",legacy_env_name=env_cfg.get("audit_root"),local_value=local_cfg.get("audit_root"),workspace_home=workspace_home,workspace_child="audit",repo_root=repo_root,repo_default=paths_cfg["audit_root"])
    project = safe_slug(str(brief.get("project_name", "INFORMATION_REQUIRED")))
    run_id = _arg(args, "run_id") or output_timestamp()
    staging_root = staging_base / project / run_id
    output_dir = output_base / project / run_id

    route_name = str(_arg(args, "route", "generate"))
    reference_raw = _arg(args, "reference_pptx")
    if not reference_raw and route_name in {"create-style-profile", "fill-template"}:
        reference_raw = _first_brief_style_reference(brief)
    reference_path = _resolve_declared_file(
        reference_raw,
        repo_root=repo_root,
        input_root=input_root,
        brief_path=brief_path,
    )
    style_profile_path = _resolve_declared_file(
        _arg(args, "style_profile"),
        repo_root=repo_root,
        input_root=input_root,
        brief_path=brief_path,
    )
    existing_pptx = _resolve_declared_file(
        _arg(args, "existing_pptx"),
        repo_root=repo_root,
        input_root=input_root,
        brief_path=brief_path,
    )
    reference_mode = str(_arg(args, "reference_mode", "style-only"))
    if route_name == "fill-template":
        reference_mode = "template-fill"
    backend = _arg(args, "backend") or workflow_config.get(
        "default_backend", "pptxgenjs"
    )
    if backend not in {"pptxgenjs", "artifact_tool"}:
        raise WorkflowError(f"Unknown backend: {backend}")
    if route_name in {"fill-template", "enhance-existing"} and backend == "artifact_tool":
        raise WorkflowError(
            f"Route {route_name!r} requires the stable PptxGenJS/native PowerPoint path"
        )

    route_request = validate_route_request(
        RouteRequest(
            route=route_name,
            brief=brief,
            reference_path=(
                reference_path
                if route_name != "fill-template"
                else None
            ),
            template_path=reference_path if route_name == "fill-template" else None,
            existing_deck_path=(
                existing_pptx if route_name == "enhance-existing" else None
            ),
            style_profile_path=style_profile_path,
            output_path=output_dir / f"{project}.pptx",
            options={
                "backend": (
                    "native-template-fill"
                    if route_name == "fill-template"
                    else ("pptxgenjs" if backend == "pptxgenjs" else None)
                ),
                "reference_mode": reference_mode,
                "update_slide": _arg(args, "update_slide"),
                "update_section": _arg(args, "update_section"),
                "update_figure": _arg(args, "update_figure"),
                "allow_managed_overwrite": bool(
                    _arg(args, "allow_managed_overwrite", False)
                ),
                "preserve_user_edits": True,
                "create_backup": True,
                "resume_from": _arg(args, "from_stage", "preflight"),
                "narrative_mode": _arg(args, "narrative_mode"),
                "use_network": False,
                "workflow_mode": "full_validation",
                "audit_full": bool(_arg(args, "audit_full", False)),
                "clinical_privacy_mode": bool(
                    _arg(args, "clinical_privacy_mode", False)
                ),
                "cross_renderer_validation": bool(
                    _arg(args, "cross_renderer_validation", False)
                ),
                "final_delivery": bool(_arg(args, "final_delivery", False)),
                "manual_review_approved": bool(
                    _arg(args, "manual_review_approved", False)
                ),
                "retry_slide": _arg(args, "retry_slide"),
                "retry_stage": _arg(args, "retry_stage"),
            },
        )
    )
    if route_name == "enhance-existing":
        chosen_updates = [
            value
            for value in (
                _arg(args, "update_slide"),
                _arg(args, "update_section"),
                _arg(args, "update_figure"),
            )
            if value
        ]
        if len(chosen_updates) != 1:
            raise WorkflowError(
                "enhance-existing requires exactly one of --update-slide, "
                "--update-section, or --update-figure"
            )

    resume = bool(_arg(args, "resume", False))
    from_stage = str(_arg(args, "from_stage", "preflight"))
    if resume:
        if not staging_root.exists() or not output_dir.exists():
            raise WorkflowError(
                "--resume requires an existing --run-id with both staging and output directories"
            )
    else:
        if from_stage != "preflight":
            raise WorkflowError("--from-stage after preflight requires --resume and an existing --run-id")
        ensure_new_directory(staging_root)
        ensure_new_directory(output_dir)
    if route_name == "create-style-profile":
        if reference_path is None or not reference_path.is_file():
            raise WorkflowError(
                "create-style-profile requires an existing .pptx or .potx reference"
            )
        if resume:
            raise WorkflowError(
                "create-style-profile is atomic; start a new run instead of using --resume"
            )
        return _execute_create_style_profile(
            reference_path=reference_path,
            reference_mode=reference_mode,
            project=project,
            run_id=run_id,
            staging_root=staging_root,
            output_dir=output_dir,
        )

    log_path = staging_root / "generation_log.md"
    scripts_root = repo_root / "scripts"
    audit_root.mkdir(parents=True, exist_ok=True)
    manifest_path = audit_root / "source_manifest.csv"
    supported = set(workflow_config["supported_extensions"])

    existing_state = staging_root / "run_state.json"
    state: dict[str, Any]
    if resume and existing_state.exists():
        state = json.loads(existing_state.read_text(encoding="utf-8"))
        state["status"] = "BLOCKED"
        state["resumed_at"] = utc_offset_timestamp()
        state["backend"] = backend
        state["route"] = route_name
    else:
        state = {
            "status": "BLOCKED",
            "project_name": project,
            "run_id": run_id,
            "backend": backend,
            "route": route_name,
            "reference_mode": reference_mode,
            "started_at": utc_offset_timestamp(),
            "repo_root": str(repo_root),
            "input_root": str(input_root),
            "output_dir": str(output_dir),
            "network_access": False,
            "external_services": False,
            "stages": {},
        }
    write_json(staging_root / "run_state.json", state)
    start_index = STAGES.index(from_stage)

    try:
        _log(
            log_path,
            "preflight",
            (
                f"repo_root={repo_root}; input_root={input_root}; "
                f"route={route_request.route}; backend={backend}"
            ),
        )
        if not input_root.exists():
            raise WorkflowError(f"Input root does not exist: {input_root}")
        if brief.get("project_name") == "INFORMATION_REQUIRED":
            raise WorkflowError("Brief project_name is INFORMATION_REQUIRED")
        if route_name == "fill-template" and (
            reference_path is None or not reference_path.is_file()
        ):
            raise WorkflowError("fill-template requires an existing .pptx or .potx")
        if route_name == "enhance-existing" and (
            existing_pptx is None or not existing_pptx.is_file()
        ):
            raise WorkflowError("enhance-existing requires an existing .pptx")
        state["stages"]["preflight"] = "passed"

        run_manifest_path = staging_root / "source_manifest.csv"
        if start_index <= STAGES.index("inventory"):
            manifest = inventory_sources(
                input_root,
                supported,
                manifest_path,
                reference_mode=reference_mode,
                route=route_name,
            )
            if not manifest:
                raise WorkflowError(f"No supported input files found under {input_root}")
            shutil.copy2(manifest_path, run_manifest_path)
            state["stages"]["inventory"] = "passed"
            _log(log_path, "inventory", f"sources={len(manifest)}")
        else:
            if not run_manifest_path.exists():
                raise WorkflowError("Cannot resume: run-specific source_manifest.csv is missing")
            manifest = _read_inventory_csv(run_manifest_path)

        style_profile: dict[str, Any] | None = None
        staging_style_profile = staging_root / "reference_style_profile.json"
        if style_profile_path is not None:
            style_profile = _load_json_mapping(style_profile_path, "Style profile")
            if start_index <= STAGES.index("inventory"):
                write_json(staging_style_profile, style_profile)
        elif route_name == "fill-template" and reference_path is not None:
            if start_index <= STAGES.index("inventory"):
                style_profile = parse_reference_style(
                    reference_path, mode="template-fill"
                )
                write_json(staging_style_profile, style_profile)
            elif staging_style_profile.exists():
                style_profile = _load_json_mapping(
                    staging_style_profile, "Staged style profile"
                )
        elif route_name == "generate":
            style_rows = [
                row
                for row in manifest
                if row.get("source_role") == "style_reference"
                and row.get("file_type") in {"pptx", "potx"}
            ]
            if style_rows and start_index <= STAGES.index("inventory"):
                first_reference = input_root / style_rows[0]["relative_path"]
                style_profile = parse_reference_style(
                    first_reference, mode=reference_mode
                )
                write_json(staging_style_profile, style_profile)
                _log(
                    log_path,
                    "inventory",
                    (
                        "reference_style_import="
                        f"{style_rows[0]['source_id']}; mode={reference_mode}"
                    ),
                )
            elif staging_style_profile.exists():
                style_profile = _load_json_mapping(
                    staging_style_profile, "Staged style profile"
                )
        if style_profile is not None:
            state["reference_style_profile"] = {
                "mode": style_profile.get("reference_mode", reference_mode),
                "source_sha256": style_profile.get("source", {}).get(
                    "sha256", "UNKNOWN"
                ),
                "application": (
                    "APPLIED_BY_NATIVE_TEMPLATE"
                    if route_name == "fill-template"
                    else (
                        "APPLIED_TO_PPTXGENJS_THEME_TOKENS"
                        if route_name == "generate" and backend == "pptxgenjs"
                        else "REGISTERED_FOR_AUDIT"
                    )
                ),
            }

        if start_index <= STAGES.index("extract"):
            extracted = extract_all(input_root, staging_root, manifest)
            update_manifest(manifest_path, manifest)
            update_manifest(run_manifest_path, manifest)
            if not any(row["parsed_successfully"] == "yes" for row in manifest):
                raise WorkflowError("All source parsers failed")
            state["stages"]["extract"] = "passed"
        else:
            extracted = {}
            for row in manifest:
                extracted_path = staging_root / "extracted" / f"{row['source_id']}.txt"
                if not extracted_path.exists():
                    raise WorkflowError(f"Cannot resume: missing extracted source {extracted_path.name}")
                extracted[row["source_id"]] = extracted_path.read_text(encoding="utf-8")

        if start_index <= STAGES.index("evidence"):
            claims, unresolved = build_evidence(extracted, manifest, staging_root, brief)
        else:
            claims = read_csv(staging_root / "evidence_inventory.csv")
            unresolved = []
        unresolved_path = audit_root / "unresolved_items.md"
        if start_index <= STAGES.index("evidence"):
            unresolved_lines = [
                "# Unresolved Items",
                "",
                f"Run: `{run_id}`",
                "",
                *(f"- INFORMATION_REQUIRED: {item}" for item in unresolved),
                *(["- None detected by automated checks; final scientific approval remains manual."] if not unresolved else []),
            ]
            unresolved_path.write_text("\n".join(unresolved_lines), encoding="utf-8")
        state["stages"]["evidence"] = "passed"

        previous_ir: dict[str, Any] | None = None
        previous_ir_path = _resolve_declared_file(
            _arg(args, "previous_deck_ir"),
            repo_root=repo_root,
            input_root=input_root,
            brief_path=brief_path,
        )
        if route_name == "enhance-existing":
            if previous_ir_path is None and existing_pptx is not None:
                sibling_ir = existing_pptx.parent / "deck_ir.json"
                if sibling_ir.is_file():
                    previous_ir_path = sibling_ir
            if previous_ir_path is None:
                raise WorkflowError(
                    "enhance-existing requires --previous-deck-ir or a deck_ir.json "
                    "beside the existing PPTX"
                )
            previous_ir = _load_json_mapping(previous_ir_path, "Previous Deck IR")
            validate_deck_ir(previous_ir)

        if start_index <= STAGES.index("storyboard"):
            storyboard = create_storyboard(
                brief,
                claims,
                manifest,
                input_root,
                staging_root,
                existing_ir=previous_ir,
                narrative_mode=_arg(args, "narrative_mode"),
                template_layout_id=_arg(args, "template_layout_id"),
                style_profile=style_profile,
                repo_root=repo_root,
            )
            write_speaker_notes(staging_root, storyboard)
        else:
            storyboard = read_csv(staging_root / "storyboard.csv")
        if not (staging_root / "deck_outline.md").exists():
            raise WorkflowError("Mandatory deck outline missing")
        state["stages"]["outline"] = "passed"
        if not (staging_root / "storyboard.csv").exists():
            raise WorkflowError("Mandatory storyboard missing")
        state["stages"]["storyboard"] = "passed"

        deck_ir_path = staging_root / "deck_ir.json"
        deck_ir = _load_json_mapping(deck_ir_path, "Deck IR")
        validate_deck_ir(deck_ir)
        state["deck_ir"] = {
            "deck_id": deck_ir["deck_id"],
            "deck_version": deck_ir["deck_version"],
            "canonical_hash": deck_ir["canonical_hash"],
        }

        incremental_targets = []
        incremental_previous_report: dict[str, Any] | None = None
        incremental_warnings: list[str] = []
        incremental_source_hash = ""
        pptx_path = output_dir / f"{project}.pptx"
        artifact_preview = staging_root / "artifact_render"
        if start_index <= STAGES.index("generate"):
            if route_name == "fill-template":
                assert reference_path is not None
                generator_log = run_native_template_fill(
                    template_path=reference_path,
                    deck_ir_path=deck_ir_path,
                    output_pptx=pptx_path,
                    script_path=scripts_root / "fill_native_template.ps1",
                    timeout_seconds=int(
                        workflow_config.get("render_timeout_seconds", 120)
                    ),
                )
                state["backend"] = "native-template-fill"
            elif route_name == "enhance-existing":
                assert existing_pptx is not None and previous_ir is not None
                previous_ids = [
                    str(item.get("slide_id")) for item in previous_ir["slides"]
                ]
                candidate_ids = [
                    str(item.get("slide_id")) for item in deck_ir["slides"]
                ]
                if previous_ids != candidate_ids:
                    raise WorkflowError(
                        "Incremental replacement requires the same stable slide_id "
                        "set and order; use generate for structural deck changes"
                    )
                node = _detect_node(
                    _arg(args, "node")
                    or os.environ.get(env_cfg["node_executable"])
                )
                node_modules = _arg(args, "node_modules") or os.environ.get(
                    env_cfg["node_modules"]
                )
                candidate_pptx = staging_root / "incremental_candidate.pptx"
                candidate_log = _run_generator(
                    backend,
                    node,
                    node_modules,
                    scripts_root,
                    staging_root / "build_spec.json",
                    candidate_pptx,
                    artifact_preview,
                    int(workflow_config.get("render_timeout_seconds", 120)),
                )
                existing_slide_order = _pptx_slide_id_order(existing_pptx)
                if set(existing_slide_order) != set(previous_ids):
                    raise WorkflowError(
                        "Existing PPTX Slide-ID set differs from the previous Deck IR"
                    )
                selected_ids = resolve_update_slide_ids(
                    deck_ir,
                    update_slide=_arg(args, "update_slide"),
                    update_section=_arg(args, "update_section"),
                    update_figure=_arg(args, "update_figure"),
                )
                targets = build_update_targets(previous_ir, deck_ir, selected_ids)
                actual_index_by_id = {
                    slide_id: index
                    for index, slide_id in enumerate(
                        existing_slide_order, start=1
                    )
                }
                targets = [
                    replace(
                        target,
                        target_index=actual_index_by_id[target.slide_id],
                    )
                    for target in targets
                ]
                physical_previous_ir = _reorder_deck_ir(
                    previous_ir, existing_slide_order
                )
                existing_report = inspect_package(
                    existing_pptx,
                    known_source_ids={
                        str(row["source_id"]) for row in manifest
                    },
                    deck_ir=physical_previous_ir,
                )
                current_manifest = _build_object_manifest_document(
                    existing_report, physical_previous_ir
                )
                previous_manifest_path = _resolve_declared_file(
                    _arg(args, "previous_object_manifest"),
                    repo_root=repo_root,
                    input_root=input_root,
                    brief_path=brief_path,
                )
                if previous_manifest_path is None and previous_ir_path is not None:
                    sibling_manifest = previous_ir_path.parent / "object_manifest.json"
                    if sibling_manifest.is_file():
                        previous_manifest_path = sibling_manifest
                previous_manifest = (
                    _load_json_mapping(
                        previous_manifest_path, "Previous object manifest"
                    )
                    if previous_manifest_path is not None
                    else None
                )
                update_warnings = verify_no_manual_conflict(
                    selected_ids,
                    current_manifest,
                    previous_manifest,
                    allow_managed_overwrite=bool(
                        _arg(args, "allow_managed_overwrite", False)
                    ),
                )
                backup = backup_source_deck(
                    existing_pptx, staging_root / "backups"
                )
                replacement_log = run_powerpoint_replacement(
                    existing_pptx=existing_pptx,
                    candidate_pptx=candidate_pptx,
                    output_pptx=pptx_path,
                    targets=targets,
                    script_path=scripts_root / "replace_pptx_slides.ps1",
                    timeout_seconds=int(
                        workflow_config.get("render_timeout_seconds", 120)
                    ),
                )
                source_after = sha256_file(existing_pptx)
                if source_after != backup["source_sha256"]:
                    raise WorkflowError(
                        "Existing PPTX changed during incremental replacement"
                    )
                deck_ir = _reorder_deck_ir(
                    deck_ir,
                    existing_slide_order,
                    bump_version_for_reorder=True,
                )
                storyboard = _rewrite_storyboard_artifacts(
                    staging_root, deck_ir, project
                )
                state["deck_ir"] = {
                    "deck_id": deck_ir["deck_id"],
                    "deck_version": deck_ir["deck_version"],
                    "canonical_hash": deck_ir["canonical_hash"],
                }
                generator_log = candidate_log + "\n" + replacement_log
                state["incremental_update"] = {
                    "selected_slide_ids": selected_ids,
                    "backup_path": backup["backup_path"],
                    "source_unchanged": True,
                    "warnings": update_warnings,
                }
                incremental_targets = targets
                incremental_previous_report = existing_report
                incremental_warnings = update_warnings
                incremental_source_hash = backup["source_sha256"]
            else:
                node = _detect_node(
                    _arg(args, "node")
                    or os.environ.get(env_cfg["node_executable"])
                )
                node_modules = _arg(args, "node_modules") or os.environ.get(
                    env_cfg["node_modules"]
                )
                generator_log = _run_generator(
                    backend,
                    node,
                    node_modules,
                    scripts_root,
                    staging_root / "build_spec.json",
                    pptx_path,
                    artifact_preview,
                    int(workflow_config.get("render_timeout_seconds", 120)),
                )
            _log(log_path, "generate", generator_log)
        elif not pptx_path.exists():
            raise WorkflowError("Cannot resume: generated PPTX is missing")
        state["stages"]["generate"] = "passed"

        pdf_path = output_dir / f"{project}.pdf"
        preview_dir = output_dir / "preview"
        if start_index <= STAGES.index("render"):
            render_method, previews, render_log = render_pptx(
                pptx_path,
                pdf_path,
                preview_dir,
                scripts_root,
                int(workflow_config.get("render_timeout_seconds", 120)),
                _arg(args, "soffice")
                or os.environ.get(env_cfg["soffice_executable"]),
                _arg(args, "pdftoppm")
                or os.environ.get(env_cfg["pdftoppm_executable"]),
            )
            create_contact_sheet(previews, preview_dir / "contact_sheet.png")
            _log(log_path, "render", render_log)
        else:
            if not pdf_path.exists() or not preview_dir.exists():
                raise WorkflowError("Cannot resume: rendered PDF or previews are missing")
            previews = sorted(preview_dir.glob("slide_*.png"))
            prior_runtime = output_dir / "runtime_manifest.json"
            render_method = "previous successful render"
            if prior_runtime.exists():
                render_method = json.loads(prior_runtime.read_text(encoding="utf-8")).get(
                    "renderer", render_method
                )
        state["stages"]["render"] = "passed"

        powerpoint_layout_path = staging_root / "powerpoint_layout_manifest.json"
        if start_index <= STAGES.index("visual_qa"):
            if render_method == "PowerPoint COM":
                try:
                    powerpoint_layout_report, layout_log = (
                        inspect_powerpoint_text_layout(
                            pptx_path=pptx_path,
                            output_json=powerpoint_layout_path,
                            script_path=scripts_root / "inspect_pptx_layout.ps1",
                            timeout_seconds=int(
                                workflow_config.get(
                                    "render_timeout_seconds", 120
                                )
                            ),
                        )
                    )
                    _log(log_path, "visual_qa", layout_log)
                except VisualLayoutQAError as exc:
                    powerpoint_layout_report = {
                        "schema_version": "1.0",
                        "status": "ERROR",
                        "error": str(exc),
                        "privacy_contract": {
                            "raw_text_returned": False,
                            "shape_names_returned": False,
                            "text_hashes_only": True,
                        },
                    }
                    write_json(
                        powerpoint_layout_path, powerpoint_layout_report
                    )
            else:
                powerpoint_layout_report = {
                    "schema_version": "1.0",
                    "status": "NOT_RUN",
                    "reason": (
                        "Renderer was not PowerPoint COM; rendered-image QA "
                        "remains active"
                    ),
                    "privacy_contract": {
                        "raw_text_returned": False,
                        "shape_names_returned": False,
                        "text_hashes_only": True,
                    },
                }
                write_json(powerpoint_layout_path, powerpoint_layout_report)
        else:
            if not powerpoint_layout_path.exists():
                raise WorkflowError(
                    "Cannot resume: PowerPoint layout manifest is missing"
                )
            powerpoint_layout_report = _load_json_mapping(
                powerpoint_layout_path, "PowerPoint layout manifest"
            )

        ooxml_report_path = staging_root / "ooxml_qa_report.json"
        object_manifest_path = staging_root / "object_manifest.json"
        if start_index <= STAGES.index("ooxml_qa"):
            ooxml_report = inspect_package(
                pptx_path,
                known_source_ids={str(row["source_id"]) for row in manifest},
                deck_ir=deck_ir,
            )
            write_json(ooxml_report_path, ooxml_report)
            object_manifest = _build_object_manifest_document(
                ooxml_report, deck_ir
            )
            write_json(object_manifest_path, object_manifest)
            state["stages"]["ooxml_qa"] = (
                "passed" if ooxml_report.get("valid") else "failed"
            )
            _log(
                log_path,
                "ooxml_qa",
                (
                    f"status={ooxml_report.get('status')}; "
                    f"errors={len(ooxml_report.get('errors', []))}; "
                    f"warnings={len(ooxml_report.get('warnings', []))}"
                ),
            )
        else:
            if not ooxml_report_path.exists() or not object_manifest_path.exists():
                raise WorkflowError(
                    "Cannot resume: OOXML QA report or object manifest is missing"
                )
            ooxml_report = _load_json_mapping(
                ooxml_report_path, "OOXML QA report"
            )
            object_manifest = _load_json_mapping(
                object_manifest_path, "Object manifest"
            )

        if (
            route_name == "enhance-existing"
            and incremental_targets
            and incremental_previous_report is not None
            and existing_pptx is not None
        ):
            previous_hashes = _identity_hashes(incremental_previous_report)
            current_hashes = _identity_hashes(ooxml_report)
            target_ids = {target.slide_id for target in incremental_targets}
            untouched_differences = sorted(
                slide_id
                for slide_id, previous_hash in previous_hashes.items()
                if slide_id not in target_ids
                and current_hashes.get(slide_id) != previous_hash
            )
            write_slide_diff(
                staging_root / "slide_diff.json",
                targets=incremental_targets,
                source_hash_before=incremental_source_hash,
                source_hash_after=sha256_file(existing_pptx),
                untouched_differences=untouched_differences,
                warnings=incremental_warnings,
            )

        figures = read_csv(staging_root / "figure_inventory.csv")
        density_issues = write_density_reports(output_dir, deck_ir)
        status, scientific_issues, visual_issues, file_issues = run_qa(
            pptx_path,
            pdf_path,
            previews,
            artifact_preview / "layouts",
            storyboard,
            claims,
            manifest,
            input_root,
            output_dir,
            render_method,
            ooxml_report=ooxml_report,
            powerpoint_layout_report=powerpoint_layout_report,
            density_issues=density_issues,
        )
        state["status"] = status
        state["stages"]["visual_qa"] = "passed" if not visual_issues else "failed"
        state["stages"]["scientific_qa"] = "passed" if not scientific_issues else "failed"

        for filename in [
            "deck_outline.md",
            "storyboard.csv",
            "speaker_notes.md",
            "deck_ir.json",
            "ooxml_qa_report.json",
            "object_manifest.json",
            "powerpoint_layout_manifest.json",
            "planned_geometry_manifest.csv",
            "planned_geometry_qa.json",
            "chart_data_contracts.json",
        ]:
            _copy_required(staging_root / filename, output_dir / filename)
        for optional_filename in [
            "reference_style_profile.json",
            "slide_diff.json",
        ]:
            optional_path = staging_root / optional_filename
            if optional_path.exists():
                shutil.copy2(optional_path, output_dir / optional_filename)
        _copy_required(run_manifest_path, output_dir / "source_manifest.csv")
        write_traceability_maps(output_dir, storyboard, claims, figures)
        _copy_required(log_path, output_dir / "generation_log.md")

        dependency_lines = [
            "Runtime dependencies used or declared",
            f"Python={sys.version.split()[0]}",
            "pypdf=6.10.0",
            "python-docx=1.2.0",
            "openpyxl=3.1.5",
            "Pillow=12.2.0",
            "PptxGenJS=4.0.1 (MIT; default project backend)",
            "@oai/artifact-tool=2.8.33 (proprietary; internal evaluation/testing only)",
            f"Selected route={route_name}",
            f"Selected backend={state.get('backend', backend)}",
            f"Renderer={render_method}",
            "OOXML QA=Python standard library (read-only)",
            "Network access=none",
            "External services=none",
        ]
        (output_dir / "dependency_manifest.txt").write_text("\n".join(dependency_lines) + "\n", encoding="utf-8")
        manual = [
            "# Manual Review Checklist",
            "",
            "- [ ] Confirm every title and key message matches the source meaning.",
            "- [ ] Confirm all numbers, units, sample sizes, effect estimates, uncertainty, and labels.",
            "- [ ] Confirm observational findings are not phrased causally.",
            "- [ ] Confirm computational inferences are not presented as experimental validation.",
            "- [ ] Confirm citations are real, complete, and correctly located.",
            "- [ ] Confirm negative findings, uncertainty, and limitations are not omitted.",
            "- [ ] Confirm no patient-level or directly identifying data appears.",
            "- [ ] Review every slide at presentation scale and in speaker-notes view.",
            "- [ ] Approve the final scientific conclusion.",
            "",
            "Automated status does not replace user approval.",
        ]
        (output_dir / "manual_review_checklist.md").write_text("\n".join(manual), encoding="utf-8")

        state["finished_at"] = utc_offset_timestamp()
        state["renderer"] = render_method
        state["ooxml_qa"] = {
            "status": ooxml_report.get("status", "UNKNOWN"),
            "error_count": len(ooxml_report.get("errors", [])),
            "warning_count": len(ooxml_report.get("warnings", [])),
        }
        state["powerpoint_layout_qa"] = {
            "status": powerpoint_layout_report.get("status", "UNKNOWN"),
        }
        state["stages"]["package"] = "passed"
        try:
            from .fast_enhance import promote_full_validation_cache

            configured_cache = _arg(args, "project_cache_root") or paths_cfg.get(
                "project_cache_root", ".cache/project_state"
            )
            cache_root = Path(configured_cache).expanduser()
            if not cache_root.is_absolute():
                cache_root = (repo_root / cache_root).resolve()
            cache_pointers = promote_full_validation_cache(
                repo_root=repo_root,
                project_identity=str(brief.get("project_name", project)),
                input_root=input_root,
                manifest=manifest,
                extracted=extracted,
                claims=claims,
                deck_ir=deck_ir,
                storyboard=storyboard,
                style_profile=style_profile,
                figures=figures,
                qa_status={
                    "status": status,
                    "scientific_issue_count": len(scientific_issues),
                    "visual_issue_count": len(visual_issues),
                    "file_issue_count": len(file_issues),
                },
                output_pptx=pptx_path,
                brief=brief,
                cache_root=cache_root,
                project_root=Path(_arg(args, "project_root")).expanduser().resolve()
                if _arg(args, "project_root")
                else None,
                cache_home=Path(os.environ["PPT_CACHE_HOME"]).expanduser()
                if os.environ.get("PPT_CACHE_HOME")
                else None,
                clinical_privacy_mode=bool(
                    _arg(args, "clinical_privacy_mode", False)
                    or brief.get("clinical_privacy_mode") is True
                    or "clinical" in str(
                        brief.get("presentation_type", "")
                    ).casefold()
                    or "mdt" in str(
                        brief.get("presentation_type", "")
                    ).casefold()
                ),
            )
            state["project_cache"] = {
                "status": "PROMOTED",
                "current_generation": cache_pointers.get("current"),
                "storage_policy": "local_private_cache_only",
            }
        except Exception as cache_exc:
            # Cache promotion must never falsify a validated deck result.  It
            # is recorded fail-closed: later fast_enhance will refuse to run
            # until a valid full-validation cache exists.
            state["project_cache"] = {
                "status": "NOT_PROMOTED",
                "error_type": type(cache_exc).__name__,
                "reason": str(cache_exc),
            }
            _log(log_path, "package", f"project_cache_not_promoted={cache_exc}")
        write_json(output_dir / "runtime_manifest.json", state)
        write_json(staging_root / "run_state.json", state)
        _log(log_path, "package", f"status={status}; output={output_dir}")
        shutil.copy2(log_path, output_dir / "generation_log.md")
        return output_dir
    except Exception as exc:
        state["status"] = "BLOCKED"
        state["finished_at"] = utc_offset_timestamp()
        state["error"] = f"{type(exc).__name__}: {exc}"
        state["traceback"] = traceback.format_exc()
        write_json(staging_root / "run_state.json", state)
        write_json(output_dir / "runtime_manifest.json", state)
        _log(log_path, "error", state["error"])
        shutil.copy2(log_path, output_dir / "generation_log.md")
        raise


def execute(args: Any) -> Path:
    """Route one request to the explicitly selected validation depth.

    ``enhance-existing`` is deliberately fast by default.  Callers that need
    the historical end-to-end validation path must request
    ``--workflow-mode full_validation``.  Imports remain lazy so the frozen
    full-validation path does not acquire fast-runtime dependencies.
    """

    route = str(_arg(args, "route", "generate"))
    requested_mode = _arg(args, "workflow_mode")
    mode = str(requested_mode or (
        "fast_enhance" if route == "enhance-existing" else "full_validation"
    ))
    if mode == "full_validation":
        return execute_full_validation(args)
    if mode != "fast_enhance":
        raise WorkflowError(
            "workflow_mode must be 'full_validation' or 'fast_enhance'"
        )
    if route != "enhance-existing":
        raise WorkflowError(
            "fast_enhance is restricted to route 'enhance-existing'"
        )
    from .fast_enhance import execute_fast_enhance
    if (
        _arg(args, "existing_pptx")
        and (not _arg(args, "operation_plan") or not _arg(args, "changed_candidate_pptx"))
    ):
        from .fast_enhance import execute_fast_production

        return execute_fast_production(args)
    return execute_fast_enhance(args)
