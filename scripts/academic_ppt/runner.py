from __future__ import annotations

import csv
import json
import os
import shutil
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any

from .evidence import build_evidence
from .extractors import extract_all
from .inventory import inventory_sources, update_manifest
from .qa import run_qa, write_traceability_maps
from .rendering import create_contact_sheet, render_pptx
from .storyboard import create_storyboard, write_speaker_notes
from .utils import (
    ensure_new_directory,
    load_yaml_compatible,
    output_timestamp,
    read_csv,
    resolve_path,
    safe_slug,
    utc_offset_timestamp,
    write_json,
)


STAGES = [
    "preflight",
    "inventory",
    "extract",
    "evidence",
    "outline",
    "storyboard",
    "generate",
    "render",
    "visual_qa",
    "scientific_qa",
    "package",
]


class WorkflowError(RuntimeError):
    pass


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


def execute(args: Any) -> Path:
    repo_root = Path(args.repo_root).resolve() if args.repo_root else Path(__file__).resolve().parents[2]
    workflow_config = load_yaml_compatible(repo_root / "config" / "workflow.yaml")
    brief_path = Path(args.brief).resolve() if Path(args.brief).is_absolute() else (repo_root / args.brief).resolve()
    brief = load_yaml_compatible(brief_path)
    paths_cfg = workflow_config["paths"]
    env_cfg = workflow_config["environment"]
    input_root = resolve_path(args.input_root, env_cfg["input_root"], paths_cfg["input_root"], repo_root, "input")
    staging_base = resolve_path(args.staging_root, env_cfg["staging_root"], paths_cfg["staging_root"], repo_root, "staging")
    output_base = resolve_path(args.output_root, env_cfg["output_root"], paths_cfg["output_root"], repo_root, "output")
    project = safe_slug(str(brief.get("project_name", "INFORMATION_REQUIRED")))
    run_id = args.run_id or output_timestamp()
    staging_root = staging_base / project / run_id
    output_dir = output_base / project / run_id
    if args.resume:
        if not staging_root.exists() or not output_dir.exists():
            raise WorkflowError(
                "--resume requires an existing --run-id with both staging and output directories"
            )
    else:
        if args.from_stage != "preflight":
            raise WorkflowError("--from-stage after preflight requires --resume and an existing --run-id")
        ensure_new_directory(staging_root)
        ensure_new_directory(output_dir)
    log_path = staging_root / "generation_log.md"
    scripts_root = repo_root / "scripts"
    manifest_path = repo_root / "audit" / "source_manifest.csv"
    supported = set(workflow_config["supported_extensions"])
    backend = args.backend or workflow_config.get("default_backend", "pptxgenjs")
    if backend not in {"pptxgenjs", "artifact_tool"}:
        raise WorkflowError(f"Unknown backend: {backend}")

    existing_state = staging_root / "run_state.json"
    state: dict[str, Any]
    if args.resume and existing_state.exists():
        state = json.loads(existing_state.read_text(encoding="utf-8"))
        state["status"] = "BLOCKED"
        state["resumed_at"] = utc_offset_timestamp()
        state["backend"] = backend
    else:
        state = {
            "status": "BLOCKED",
            "project_name": project,
            "run_id": run_id,
            "backend": backend,
            "started_at": utc_offset_timestamp(),
            "repo_root": str(repo_root),
            "input_root": str(input_root),
            "output_dir": str(output_dir),
            "stages": {},
        }
    write_json(staging_root / "run_state.json", state)
    start_index = STAGES.index(args.from_stage)

    try:
        _log(log_path, "preflight", f"repo_root={repo_root}; input_root={input_root}; backend={backend}")
        if not input_root.exists():
            raise WorkflowError(f"Input root does not exist: {input_root}")
        if brief.get("project_name") == "INFORMATION_REQUIRED":
            raise WorkflowError("Brief project_name is INFORMATION_REQUIRED")
        state["stages"]["preflight"] = "passed"

        run_manifest_path = staging_root / "source_manifest.csv"
        if start_index <= STAGES.index("inventory"):
            manifest = inventory_sources(input_root, supported, manifest_path)
            if not manifest:
                raise WorkflowError(f"No supported input files found under {input_root}")
            shutil.copy2(manifest_path, run_manifest_path)
            state["stages"]["inventory"] = "passed"
            _log(log_path, "inventory", f"sources={len(manifest)}")
        else:
            if not run_manifest_path.exists():
                raise WorkflowError("Cannot resume: run-specific source_manifest.csv is missing")
            manifest = _read_inventory_csv(run_manifest_path)

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
        unresolved_path = repo_root / "audit" / "unresolved_items.md"
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

        if start_index <= STAGES.index("storyboard"):
            storyboard = create_storyboard(brief, claims, manifest, input_root, staging_root)
            write_speaker_notes(staging_root, storyboard)
        else:
            storyboard = read_csv(staging_root / "storyboard.csv")
        if not (staging_root / "deck_outline.md").exists():
            raise WorkflowError("Mandatory deck outline missing")
        state["stages"]["outline"] = "passed"
        if not (staging_root / "storyboard.csv").exists():
            raise WorkflowError("Mandatory storyboard missing")
        state["stages"]["storyboard"] = "passed"

        node = _detect_node(args.node or os.environ.get(env_cfg["node_executable"]))
        node_modules = args.node_modules or os.environ.get(env_cfg["node_modules"])
        pptx_path = output_dir / f"{project}.pptx"
        artifact_preview = staging_root / "artifact_render"
        if start_index <= STAGES.index("generate"):
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
                args.soffice or os.environ.get(env_cfg["soffice_executable"]),
                args.pdftoppm or os.environ.get(env_cfg["pdftoppm_executable"]),
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

        figures = read_csv(staging_root / "figure_inventory.csv")
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
        )
        state["status"] = status
        state["stages"]["visual_qa"] = "passed" if not visual_issues else "failed"
        state["stages"]["scientific_qa"] = "passed" if not scientific_issues else "failed"

        for filename in [
            "deck_outline.md",
            "storyboard.csv",
            "speaker_notes.md",
        ]:
            _copy_required(staging_root / filename, output_dir / filename)
        _copy_required(manifest_path, output_dir / "source_manifest.csv")
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
            f"Selected backend={backend}",
            f"Renderer={render_method}",
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
        state["stages"]["package"] = "passed"
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
