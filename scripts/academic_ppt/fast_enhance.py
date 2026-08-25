from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import traceback
import zipfile
from pathlib import Path
from typing import Any, Iterable, Mapping

from pypdf import PdfReader
from PIL import Image, ImageStat

from .change_impact import ChangeImpactGraph
from .evidence import build_evidence
from .extractors import extract_selected
from .fast_qa import run_changed_slide_qa, run_whole_deck_lightweight_qa
from .incremental import (
    run_powerpoint_operations,
    validate_operation_plan,
)
from .inventory import inventory_sources, verify_input_hashes
from .project_cache import (
    CacheStateError,
    ProjectCache,
    build_cache_contract_fingerprints,
    build_project_state,
    build_source_manifest,
    resolve_project_cache_root,
    validate_project_state,
)
from .style_reference import parse_reference_style
from .runtime_profile import RuntimeProfiler
from .utils import (
    ensure_new_directory,
    load_yaml_compatible,
    load_local_runtime_config,
    output_timestamp,
    read_csv,
    resolve_path,
    resolve_runtime_path,
    resolve_workflow_home,
    safe_slug,
    sha256_file,
    utc_offset_timestamp,
    write_json,
)
from .visual_layout_qa import (
    inspect_powerpoint_text_layout,
    inspect_text_geometry,
)


class FastEnhanceError(RuntimeError):
    """Raised when a fast update cannot be proven safe and local."""


RETRYABLE_FAST_STAGES = frozenset(
    {
        "changed_slide_generation",
        "powerpoint_apply",
        "changed_slide_render",
        "changed_slide_qa",
        "whole_deck_light_qa",
    }
)
_CHECKPOINT_STAGE_ORDER = (
    "preflight",
    "source_snapshot",
    "cache_load",
    "delta_classification",
    "source_parse",
    "evidence_rebuild",
    "impact_resolution",
    "slide_planning",
    "changed_slide_generation",
    "powerpoint_apply",
    "changed_slide_render",
    "changed_slide_qa",
    "whole_deck_light_qa",
)
_FAST_CACHE_APPROVED_STATUSES = frozenset(
    {
        "PASS",
        "PASSED",
        "READY_FOR_ASSISTED_USE",
        "DRAFT_READY",
        "FINAL_DELIVERY_READY",
    }
)


def _arg(args: Any, name: str, default: Any = None) -> Any:
    return getattr(args, name, default)


def _load_json(path: Path, label: str) -> dict[str, Any]:
    if not path.is_file():
        raise FastEnhanceError(f"{label} is missing: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise FastEnhanceError(f"{label} is not valid JSON: {path}") from exc
    if not isinstance(value, dict):
        raise FastEnhanceError(f"{label} must contain a JSON object")
    return value


def _resolve_declared(raw: str | Path | None, repo_root: Path) -> Path | None:
    if raw in {None, ""}:
        return None
    path = Path(raw).expanduser()
    return path.resolve() if path.is_absolute() else (repo_root / path).resolve()


def _assert_clinical_cache_is_ignored(
    repo_root: Path,
    cache_root: Path | None,
    *,
    project_root: Path | None = None,
    cache_home: Path | None = None,
    production: bool = False,
) -> Path:
    """Validate local placement and Git exclusion before creating cache state.

    The check intentionally runs before :class:`ProjectCache` creates its
    project-key secret.  A clinical run therefore cannot materialise even
    metadata in an unprotected repository tree.
    """

    resolved = resolve_project_cache_root(
        repo_root,
        project_root=project_root,
        explicit_cache_root=cache_root,
        cache_home=cache_home,
        production=production,
    )
    repo = repo_root.resolve()
    try:
        resolved.relative_to(repo)
    except ValueError:
        # An external project cache or configured PPT_CACHE_HOME is already
        # contained by the canonical production policy and is not publishable
        # through this repository.
        return resolved
    gitignore = repo_root / ".gitignore"
    if not gitignore.is_file():
        raise FastEnhanceError(
            "Clinical fast cache is not protected by repository .gitignore"
        )
    lines = [
        line.strip().replace("\\", "/")
        for line in gitignore.read_text(encoding="utf-8-sig").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]
    # A later negation can expose the cache even when an earlier broad rule
    # ignored it, so fail closed for any explicit .cache negation.
    if any(line.startswith("!") and ".cache" in line for line in lines):
        raise FastEnhanceError(
            "Clinical fast cache has a conflicting .gitignore negation"
        )
    try:
        relative = resolved.relative_to(repo).as_posix()
    except ValueError:  # pragma: no cover - returned above
        relative = ""
    protected = subprocess.run(
        ["git", "-C", str(repo), "check-ignore", "--no-index", "-q", relative],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    ).returncode == 0
    if not protected:
        normalized = {
            line.lstrip("/").removesuffix("/**").rstrip("/")
            for line in lines
            if not line.startswith("!")
        }
        protected = any(
            relative == pattern or relative.startswith(pattern + "/")
            for pattern in normalized
            if pattern
        )
    if not protected:
        raise FastEnhanceError(
            "Clinical fast cache is not protected by repository .gitignore"
        )
    return resolved


def validate_fast_runtime_controls(
    args: Any, changed_slide_ids: Iterable[str]
) -> dict[str, Any]:
    """Validate recovery and delivery switches without touching the deck."""

    changed_ids = [str(item).strip() for item in changed_slide_ids]
    if not changed_ids or any(not item for item in changed_ids):
        raise FastEnhanceError("Operation plan contains no valid changed slides")
    retry_slide = str(_arg(args, "retry_slide", "") or "").strip()
    retry_stage = str(_arg(args, "retry_stage", "") or "").strip()
    if retry_slide and retry_stage:
        raise FastEnhanceError(
            "--retry-slide and --retry-stage are mutually exclusive"
        )
    if retry_slide and changed_ids != [retry_slide]:
        raise FastEnhanceError(
            "--retry-slide requires an isolated one-slide operation plan for that slide_id"
        )
    if retry_stage and retry_stage not in RETRYABLE_FAST_STAGES:
        raise FastEnhanceError(
            "--retry-stage must name a changed-only generation, PowerPoint, or QA stage"
        )
    final_delivery = bool(_arg(args, "final_delivery", False))
    manual_approved = bool(_arg(args, "manual_review_approved", False))
    if manual_approved and not final_delivery:
        raise FastEnhanceError(
            "--manual-review-approved requires --final-delivery"
        )
    return {
        "retry_slide": retry_slide,
        "retry_stage": retry_stage,
        "final_delivery": final_delivery,
        "manual_review_approved": manual_approved,
        "cross_renderer_validation": bool(
            _arg(args, "cross_renderer_validation", False)
        ),
        "audit_full": bool(_arg(args, "audit_full", False)),
    }


def retry_stage_execution_policy(retry_stage: str) -> dict[str, bool]:
    """Return the deterministic stage skip/reuse policy for one retry."""

    stage = str(retry_stage or "").strip()
    if stage and stage not in RETRYABLE_FAST_STAGES:
        raise FastEnhanceError(f"Unsupported retry stage: {stage}")
    return {
        "run_inventory": True,
        "run_extract": not bool(stage),
        "run_evidence": not bool(stage),
        "run_slide_planning": not bool(stage),
        "run_candidate_validation": stage in {"", "changed_slide_generation"},
        "run_powerpoint_apply": stage in {
            "",
            "changed_slide_generation",
            "powerpoint_apply",
        },
        "run_powerpoint_geometry": stage in {
            "",
            "changed_slide_generation",
            "powerpoint_apply",
            "changed_slide_render",
        },
        "run_changed_slide_qa": stage != "whole_deck_light_qa",
        "run_whole_deck_light_qa": True,
    }


def _checkpoint_path(staging_root: Path) -> Path:
    return staging_root / "fast_enhance_checkpoint.json"


def _build_failure_record(
    exc: Exception, *, clinical_privacy_mode: bool
) -> dict[str, Any]:
    """Build a failure record without leaking clinical paths or source prose."""

    record: dict[str, Any] = {
        "status": "BLOCKED",
        "workflow_mode": "fast_enhance",
        "error_type": type(exc).__name__,
        "finished_at": utc_offset_timestamp(),
    }
    if clinical_privacy_mode:
        trace = traceback.format_exc()
        record.update(
            {
                "error": "REDACTED_CLINICAL_FAILURE",
                "traceback_sha256": hashlib.sha256(
                    trace.encode("utf-8")
                ).hexdigest(),
                "clinical_detail_redacted": True,
            }
        )
    else:
        record.update(
            {
                "error": str(exc),
                "traceback": traceback.format_exc(),
                "clinical_detail_redacted": False,
            }
        )
    return record


def validate_fast_cache_baseline(
    cached_state: Mapping[str, Any],
    *,
    source_pptx: Path,
    expected_slide_count: int,
    current_contract_fingerprints: Mapping[str, str],
) -> None:
    """Require a passed QA baseline bound to the exact source deck."""

    qa = cached_state.get("previous_qa_status")
    status = str(qa.get("status", "")).upper() if isinstance(qa, Mapping) else ""
    if status not in _FAST_CACHE_APPROVED_STATUSES:
        raise FastEnhanceError(
            "Cached baseline QA is not approved; run workflow_mode=full_validation"
        )
    pptx_state = cached_state.get("pptx_state")
    if not isinstance(pptx_state, Mapping):
        raise FastEnhanceError(
            "Cached baseline has no PPTX state; run workflow_mode=full_validation"
        )
    try:
        cached_count = int(pptx_state.get("slide_count", -1))
    except (TypeError, ValueError):
        cached_count = -1
    if (
        str(pptx_state.get("sha256", "")).lower() != sha256_file(source_pptx).lower()
        or cached_count != expected_slide_count
    ):
        raise FastEnhanceError(
            "Cached PPTX state does not match the operation-plan source deck; "
            "run workflow_mode=full_validation"
        )
    cached_fingerprints = cached_state.get("cache_contract_fingerprints")
    if not isinstance(cached_fingerprints, Mapping) or dict(
        cached_fingerprints
    ) != dict(current_contract_fingerprints):
        raise FastEnhanceError(
            "Cached brief, scientific definition, extractor, config, or code "
            "fingerprint differs from the current runtime; run "
            "workflow_mode=full_validation"
        )


def _validate_cached_source_slide_order(
    cached_state: Mapping[str, Any], source_slide_ids: Iterable[str]
) -> None:
    """Bind operation-plan source identities to the validated cached deck order."""

    cached_specs = cached_state.get("slide_specs")
    if not isinstance(cached_specs, list):
        raise FastEnhanceError("Cached SlideSpec collection is invalid")
    cached_ids = [
        str(item.get("slide_id", "")).strip()
        for item in cached_specs
        if isinstance(item, Mapping)
    ]
    if len(cached_ids) != len(cached_specs) or any(not item for item in cached_ids):
        raise FastEnhanceError("Cached SlideSpec order contains an invalid slide_id")
    if len(cached_ids) != len(set(cached_ids)):
        raise FastEnhanceError("Cached SlideSpec order contains duplicate slide_id values")
    planned_ids = [str(item).strip() for item in source_slide_ids]
    if cached_ids != planned_ids:
        raise FastEnhanceError(
            "Operation-plan source_slide_ids do not match the validated cached SlideSpec order"
        )


def _validate_changed_spec_source_registry(
    changed_specs: Iterable[Mapping[str, Any]],
    source_registry: Iterable[Mapping[str, Any]],
) -> None:
    """Require every changed-slide binding to resolve in the current registry."""

    by_id: dict[str, str] = {}
    by_file: dict[str, str] = {}
    for row in source_registry:
        if not isinstance(row, Mapping):
            raise FastEnhanceError("Current SourceRegistry contains a non-object entry")
        source_id = str(row.get("source_id", "")).strip()
        source_file = str(row.get("relative_path", "")).strip().replace("\\", "/")
        if not source_id or not source_file:
            raise FastEnhanceError("Current SourceRegistry entry lacks source_id or relative_path")
        if source_id in by_id or source_file in by_file:
            raise FastEnhanceError("Current SourceRegistry contains duplicate identity")
        by_id[source_id] = source_file
        by_file[source_file] = source_id

    for spec in changed_specs:
        slide_id = str(spec.get("slide_id", "")).strip()
        bindings = spec.get("source_bindings")
        if not isinstance(bindings, list) or not bindings:
            raise FastEnhanceError(f"Changed slide {slide_id} has no source bindings")
        for binding in bindings:
            if not isinstance(binding, Mapping):
                raise FastEnhanceError(
                    f"Changed slide {slide_id} contains an invalid source binding"
                )
            source_id = str(binding.get("source_id", "")).strip()
            source_file = str(binding.get("source_file", "")).strip().replace("\\", "/")
            if not source_id and not source_file:
                raise FastEnhanceError(
                    f"Changed slide {slide_id} source binding has no source identity"
                )
            if source_id and source_id not in by_id:
                raise FastEnhanceError(
                    f"Changed slide {slide_id} binds an unknown current source_id"
                )
            if source_file and source_file not in by_file:
                raise FastEnhanceError(
                    f"Changed slide {slide_id} binds an unknown current source_file"
                )
            if source_id and source_file and by_id[source_id] != source_file:
                raise FastEnhanceError(
                    f"Changed slide {slide_id} source_id/source_file binding is inconsistent"
                )
def _write_checkpoint(
    path: Path,
    *,
    completed_stage: str,
    existing_pptx: Path,
    operation_plan_path: Path,
    cache_generation: str,
    current_source_manifest: Mapping[str, Any],
    final_delivery: bool = False,
    cross_renderer_validation: bool = False,
    updated_pptx: Path | None = None,
    geometry_path: Path | None = None,
) -> None:
    if completed_stage not in _CHECKPOINT_STAGE_ORDER:
        raise FastEnhanceError(f"Unsupported checkpoint stage: {completed_stage}")
    artifacts: dict[str, dict[str, str]] = {}
    for name, candidate in (
        ("retry_context", path.parent / "fast_retry_context.json"),
        ("changed_qa", path.parent / "changed_slide_qa.json"),
        ("updated_pptx", updated_pptx),
        ("geometry", geometry_path),
    ):
        if candidate is not None and candidate.is_file():
            artifacts[name] = {
                "relative_name": candidate.name,
                "sha256": sha256_file(candidate),
            }
    preview_dir = path.parent / "changed_preview"
    preview_hashes = {
        candidate.name: sha256_file(candidate)
        for candidate in sorted(preview_dir.glob("*.png"))
        if candidate.is_file()
    }
    retry_intermediate_hashes: dict[str, str] = {}
    for pattern in ("extracted/*.txt", "evidence_delta/evidence_inventory.csv"):
        for candidate in sorted(path.parent.glob(pattern)):
            if candidate.is_file():
                retry_intermediate_hashes[
                    candidate.relative_to(path.parent).as_posix()
                ] = sha256_file(candidate)
    write_json(
        path,
        {
            "schema_version": "academic-ppt-fast-checkpoint/1",
            "completed_stage": completed_stage,
            "existing_pptx_sha256": sha256_file(existing_pptx),
            "operation_plan_sha256": sha256_file(operation_plan_path),
            "cache_generation": cache_generation,
            "source_manifest_sha256": hashlib.sha256(
                json.dumps(
                    current_source_manifest,
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                ).encode("utf-8")
            ).hexdigest(),
            "runtime_controls": {
                "final_delivery": bool(final_delivery),
                "cross_renderer_validation": bool(cross_renderer_validation),
            },
            "artifacts": artifacts,
            "changed_preview_hashes": preview_hashes,
            "retry_intermediate_hashes": retry_intermediate_hashes,
            "privacy": {
                "anonymous_stage_state_only": True,
                "contains_source_text": False,
                "contains_private_paths": False,
            },
        },
    )


def _load_retry_checkpoint(
    path: Path,
    *,
    retry_stage: str,
    existing_pptx: Path,
    operation_plan_path: Path,
    cache_generation: str,
    current_source_manifest: Mapping[str, Any],
    final_delivery: bool = False,
    cross_renderer_validation: bool = False,
    output_dir: Path,
    staging_root: Path,
) -> dict[str, Path]:
    document = _load_json(path, "Fast enhance checkpoint")
    if document.get("schema_version") != "academic-ppt-fast-checkpoint/1":
        raise FastEnhanceError("Fast enhance checkpoint schema is unsupported")
    completed = str(document.get("completed_stage", ""))
    if completed not in _CHECKPOINT_STAGE_ORDER:
        raise FastEnhanceError("Fast enhance checkpoint stage is invalid")
    required_previous = {
        "changed_slide_generation": "slide_planning",
        "powerpoint_apply": "changed_slide_generation",
        "changed_slide_render": "powerpoint_apply",
        "changed_slide_qa": "changed_slide_render",
        "whole_deck_light_qa": "changed_slide_qa",
    }[retry_stage]
    if _CHECKPOINT_STAGE_ORDER.index(completed) < _CHECKPOINT_STAGE_ORDER.index(
        required_previous
    ):
        raise FastEnhanceError(
            "Fast enhance checkpoint has insufficient completed upstream stages"
        )
    expected_manifest_hash = hashlib.sha256(
        json.dumps(
            current_source_manifest,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()
    checks = {
        "existing_pptx_sha256": sha256_file(existing_pptx),
        "operation_plan_sha256": sha256_file(operation_plan_path),
        "cache_generation": cache_generation,
        "source_manifest_sha256": expected_manifest_hash,
    }
    if any(str(document.get(key, "")) != value for key, value in checks.items()):
        raise FastEnhanceError(
            "Fast enhance checkpoint no longer matches inputs, plan, or cache generation"
        )
    if document.get("runtime_controls") != {
        "final_delivery": bool(final_delivery),
        "cross_renderer_validation": bool(cross_renderer_validation),
    }:
        raise FastEnhanceError(
            "Fast enhance retry controls differ from the checkpointed run"
        )
    artifacts = document.get("artifacts", {})
    if not isinstance(artifacts, Mapping):
        raise FastEnhanceError("Fast enhance checkpoint artifact map is invalid")
    allowed = {
        "retry_context": staging_root / "fast_retry_context.json",
        "changed_qa": staging_root / "changed_slide_qa.json",
        "updated_pptx": output_dir / "updated.pptx",
        "geometry": staging_root / "powerpoint_geometry.json",
    }
    verified: dict[str, Path] = {}
    for name, expected_path in allowed.items():
        row = artifacts.get(name)
        if not isinstance(row, Mapping):
            continue
        if row.get("relative_name") != expected_path.name or not expected_path.is_file():
            raise FastEnhanceError(f"Checkpoint artifact is missing: {name}")
        if sha256_file(expected_path) != str(row.get("sha256", "")):
            raise FastEnhanceError(f"Checkpoint artifact hash mismatch: {name}")
        verified[name] = expected_path
    if "retry_context" not in verified:
        raise FastEnhanceError("Retry requires a verified private retry context")
    intermediate_hashes = document.get("retry_intermediate_hashes")
    if not isinstance(intermediate_hashes, Mapping):
        raise FastEnhanceError("Retry intermediate manifest is invalid")
    for relative, digest in intermediate_hashes.items():
        candidate = (staging_root / str(relative)).resolve()
        try:
            candidate.relative_to(staging_root.resolve())
        except ValueError as exc:
            raise FastEnhanceError("Retry intermediate path escapes staging") from exc
        if not candidate.is_file() or sha256_file(candidate) != str(digest):
            raise FastEnhanceError("Retry intermediate hash mismatch")
    if retry_stage in {"changed_slide_render", "changed_slide_qa", "whole_deck_light_qa"} and "updated_pptx" not in verified:
        raise FastEnhanceError("Retry requires a verified updated.pptx checkpoint")
    if retry_stage in {"changed_slide_qa", "whole_deck_light_qa"} and "geometry" not in verified:
        raise FastEnhanceError("Retry requires a verified geometry checkpoint")
    if retry_stage == "whole_deck_light_qa" and "changed_qa" not in verified:
        raise FastEnhanceError("Retry requires a verified changed-slide QA checkpoint")
    if retry_stage == "changed_slide_qa":
        preview_hashes = document.get("changed_preview_hashes")
        if not isinstance(preview_hashes, Mapping) or not preview_hashes:
            raise FastEnhanceError("Retry requires verified changed-slide previews")
        preview_dir = staging_root / "changed_preview"
        for name, digest in preview_hashes.items():
            candidate = preview_dir / str(name)
            if (
                candidate.parent.resolve() != preview_dir.resolve()
                or not candidate.is_file()
                or sha256_file(candidate) != str(digest)
            ):
                raise FastEnhanceError("Changed-slide preview checkpoint mismatch")
    return verified


def _quarantine_failed_powerpoint_outputs(
    *, output_dir: Path, staging_root: Path, preview_dir: Path
) -> None:
    """Move only known failed-attempt outputs aside before a COM retry."""

    candidates = [output_dir / "updated.pptx", output_dir / "updated.pdf", preview_dir]
    existing = [path for path in candidates if path.exists()]
    if not existing:
        return
    quarantine = staging_root / "retry_previous_powerpoint_apply"
    if quarantine.exists():
        raise FastEnhanceError(
            "Retry quarantine already exists; use a new run-id or inspect manually"
        )
    quarantine.mkdir(parents=True, exist_ok=False)
    for path in existing:
        shutil.move(str(path), str(quarantine / path.name))


def _write_runtime_outputs(
    profiler: RuntimeProfiler,
    *,
    staging_root: Path,
    output_dir: Path,
    audit_full: bool,
) -> tuple[dict[str, Path | None], dict[str, Any], Path | None]:
    """Keep profiling private by default; publish only required exceptions."""

    runtime_paths = profiler.write(staging_root)
    profile = _load_json(Path(runtime_paths["runtime_profile"]), "Runtime profile")
    bottleneck_path = runtime_paths.get("runtime_bottleneck_report")
    if bottleneck_path is not None:
        shutil.copy2(
            Path(bottleneck_path), output_dir / "runtime_bottleneck_report.md"
        )
    audit_dir: Path | None = None
    if audit_full:
        audit_dir = output_dir / "audit_full"
        audit_dir.mkdir(parents=True, exist_ok=False)
        shutil.copy2(
            Path(runtime_paths["runtime_profile"]),
            audit_dir / "runtime_profile.json",
        )
    return runtime_paths, profile, audit_dir


def _operation_metadata(raw_plan: Mapping[str, Any]) -> dict[str, Any]:
    """Return canonical optional planning evidence from the operation plan."""

    return {
        key: raw_plan.get(key, [])
        for key in (
            "changed_slide_specs",
            "source_impacts",
            "source_binding_issues",
            "scientific_issues",
            "manual_review_items",
        )
    }


def _source_paths(input_root: Path, manifest: Iterable[Mapping[str, Any]]) -> dict[str, Path]:
    result: dict[str, Path] = {}
    for row in manifest:
        relative = str(row.get("relative_path", "")).strip()
        if not relative:
            raise FastEnhanceError("Current source manifest contains an empty relative_path")
        path = (input_root / relative).resolve()
        try:
            path.relative_to(input_root.resolve())
        except ValueError as exc:
            raise FastEnhanceError("Source path escapes input_root") from exc
        result[relative] = path
    return result


def _source_registry_index(state: Mapping[str, Any]) -> tuple[dict[str, str], dict[str, str]]:
    registry = state.get("source_registry", [])
    if not isinstance(registry, list):
        registry = []
    relative_by_id: dict[str, str] = {}
    id_by_relative: dict[str, str] = {}
    for row in registry:
        if not isinstance(row, Mapping):
            continue
        source_id = str(row.get("source_id", "")).strip()
        relative = str(row.get("relative_path", "")).strip()
        if source_id and relative:
            relative_by_id[source_id] = relative
            id_by_relative[relative] = source_id
    return relative_by_id, id_by_relative


def build_cached_change_impact_graph(
    *,
    source_registry: list[Mapping[str, Any]],
    evidence_registry: list[Mapping[str, Any]],
    slide_specs: list[Mapping[str, Any]],
) -> ChangeImpactGraph:
    """Build one conservative lineage graph from already-canonical objects."""

    graph = ChangeImpactGraph()
    relative_by_id = {
        str(row.get("source_id", "")): str(row.get("relative_path", ""))
        for row in source_registry
        if str(row.get("source_id", "")).strip()
        and str(row.get("relative_path", "")).strip()
    }
    for relative in relative_by_id.values():
        graph.register_source(relative)
    claims_by_id: dict[str, Mapping[str, Any]] = {}
    for claim in evidence_registry:
        if not isinstance(claim, Mapping):
            continue
        claim_id = str(claim.get("claim_id", "")).strip()
        source_id = str(claim.get("source_id", "")).strip()
        if not claim_id:
            continue
        claims_by_id[claim_id] = claim
        graph.register_claim(claim_id)
        relative = relative_by_id.get(source_id)
        if relative:
            graph.add_source_claim(relative, claim_id)
    for slide in slide_specs:
        if not isinstance(slide, Mapping):
            continue
        slide_id = str(slide.get("slide_id", "")).strip()
        if not slide_id:
            continue
        graph.register_slide(slide_id)
        visual_id = f"VIS-CACHED-{hashlib.sha256(slide_id.encode('utf-8')).hexdigest()[:20].upper()}"
        graph.register_visual(visual_id)
        graph.add_visual_slide(visual_id, slide_id)
        claim_ids = slide.get("claim_ids") or slide.get("claim_bindings") or []
        for raw_claim in claim_ids if isinstance(claim_ids, list) else []:
            claim_id = (
                str(raw_claim.get("claim_id", "")).strip()
                if isinstance(raw_claim, Mapping)
                else str(raw_claim).strip()
            )
            if not claim_id:
                continue
            if claim_id not in claims_by_id:
                graph.register_claim(claim_id)
            graph.add_claim_visual(claim_id, visual_id)
        bindings = slide.get("source_bindings") or []
        for binding in bindings if isinstance(bindings, list) else []:
            if not isinstance(binding, Mapping):
                continue
            source_id = str(binding.get("source_id", "")).strip()
            relative = str(binding.get("source_file", "")).strip() or relative_by_id.get(source_id, "")
            if not relative:
                continue
            claim_id = f"CLM-BIND-{hashlib.sha256((relative + '|' + slide_id).encode('utf-8')).hexdigest()[:20].upper()}"
            graph.add_source_claim(relative, claim_id)
            graph.add_claim_visual(claim_id, visual_id)
    return graph


def _extend_graph_from_declared_impacts(
    graph: ChangeImpactGraph,
    source_impacts: list[Mapping[str, Any]],
    *,
    allowed_slides: set[str],
    relative_by_source_id: Mapping[str, str],
) -> None:
    for row in source_impacts:
        source_id = str(row.get("source_id", "")).strip()
        source = relative_by_source_id.get(source_id, source_id)
        raw_slides = row.get("affected_slide_ids")
        if not source or not isinstance(raw_slides, list) or not raw_slides:
            raise FastEnhanceError("Every source_impacts entry requires a source key and slides")
        graph.register_source(source)
        for raw_slide in raw_slides:
            slide_id = str(raw_slide).strip()
            if slide_id not in allowed_slides:
                raise FastEnhanceError(
                    f"source_impacts references a slide outside the operation plan: {slide_id}"
                )
            claim_id = f"CLM-DELTA-{hashlib.sha256((source + '|' + slide_id).encode('utf-8')).hexdigest()[:20].upper()}"
            visual_id = f"VIS-DELTA-{hashlib.sha256((slide_id + '|' + source).encode('utf-8')).hexdigest()[:20].upper()}"
            graph.add_source_claim(source, claim_id)
            graph.add_claim_visual(claim_id, visual_id)
            graph.add_visual_slide(visual_id, slide_id)


def _changed_slide_specs(
    metadata: Mapping[str, Any],
    changed_slide_ids: list[str],
) -> list[dict[str, Any]]:
    declared = metadata.get("changed_slide_specs", [])
    if declared is None:
        declared = []
    if not isinstance(declared, list):
        raise FastEnhanceError("changed_slide_specs must be a list")
    by_id = {
        str(item.get("slide_id", "")).strip(): dict(item)
        for item in declared
        if isinstance(item, Mapping) and str(item.get("slide_id", "")).strip()
    }
    result: list[dict[str, Any]] = []
    for slide_id in changed_slide_ids:
        spec = by_id.get(slide_id)
        if spec is None:
            raise FastEnhanceError(
                f"Changed slide {slide_id} has no changed_slide_specs entry"
            )
        if str(spec.get("slide_id", "")) != slide_id:
            raise FastEnhanceError("Changed SlideSpec identity mismatch")
        bindings = spec.get("source_bindings")
        if spec.get("source_binding_required") is not False and (
            not isinstance(bindings, list) or not bindings
        ):
            raise FastEnhanceError(f"Changed slide {slide_id} has no source bindings")
        result.append(spec)
    unexpected = sorted(set(by_id) - set(changed_slide_ids))
    if unexpected:
        raise FastEnhanceError(
            "changed_slide_specs contains unchanged slides: " + ", ".join(unexpected)
        )
    return result


def _scientific_issue_rows(
    metadata: Mapping[str, Any], changed_specs: list[Mapping[str, Any]]
) -> list[dict[str, Any]]:
    raw = metadata.get("scientific_issues", []) or []
    if not isinstance(raw, list):
        raise FastEnhanceError("scientific_issues must be a list")
    issues = [dict(item) for item in raw if isinstance(item, Mapping)]
    for spec in changed_specs:
        slide_id = str(spec["slide_id"])
        status = str(
            spec.get("scientific_qa_status")
            or (spec.get("scientific_review") or {}).get("status", "")
            if isinstance(spec.get("scientific_review"), Mapping)
            else spec.get("scientific_qa_status", "")
        ).upper()
        if status not in {"PASS", "PASSED", "APPROVED_FOR_DRAFT"}:
            issues.append(
                {
                    "slide_id": slide_id,
                    "severity": "error",
                    "code": "CHANGED_SLIDE_SCIENTIFIC_REVIEW_MISSING",
                    "message": "Changed slide lacks an explicit source-bound scientific review status",
                }
            )
    return issues


def _geometry_gate(raw_report: Mapping[str, Any]) -> dict[str, Any]:
    issues_by_index: dict[int, list[dict[str, str]]] = {}
    for message in _blocking_geometry_messages(raw_report):
        match = re.match(r"Slide\s+(\d+):\s*(.+)", message)
        if not match:
            continue
        issues_by_index.setdefault(int(match.group(1)), []).append(
            {
                "severity": "error",
                "code": "POWERPOINT_ACTUAL_GEOMETRY",
                "message": match.group(2),
            }
        )
    slides: list[dict[str, Any]] = []
    for row in raw_report.get("slides", []):
        if not isinstance(row, Mapping):
            continue
        index = int(row.get("slide_index", 0) or 0)
        findings = issues_by_index.get(index, [])
        slides.append(
            {
                "slide_index": index,
                "slide_id": str(row.get("slide_id", "")),
                "status": "FAIL" if findings else "PASS",
                "issues": findings,
            }
        )
    return {"schema_version": "1.0", "slides": slides}


def _blocking_geometry_messages(raw_report: Mapping[str, Any]) -> list[str]:
    # The inspector already excludes footer/page-number roles.  Unnamed
    # pictures and scientific shapes are intentionally classified as content,
    # so suppressing generic content findings would hide real footer intrusion.
    # Explicit background exemptions must be declared by the inspector/layout
    # contract, never inferred here from the broad content role.
    return list(inspect_text_geometry(raw_report))


def _preview_visual_issues(
    preview_dir: Path,
    changed_slide_ids: list[str],
    final_index: Mapping[str, int],
) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for slide_id in changed_slide_ids:
        preview = preview_dir / f"slide_{int(final_index[slide_id]):03d}.png"
        if not preview.is_file():
            issues.append(
                {
                    "slide_id": slide_id,
                    "severity": "error",
                    "code": "CHANGED_SLIDE_PREVIEW_MISSING",
                }
            )
            continue
        with Image.open(preview) as image:
            if image.width < 1000 or image.height < 500:
                issues.append(
                    {
                        "slide_id": slide_id,
                        "severity": "error",
                        "code": "CHANGED_SLIDE_PREVIEW_LOW_RESOLUTION",
                    }
                )
            variance = sum(ImageStat.Stat(image.convert("RGB")).var) / 3.0
            if variance < 2.0:
                issues.append(
                    {
                        "slide_id": slide_id,
                        "severity": "error",
                        "code": "CHANGED_SLIDE_PREVIEW_NEAR_BLANK",
                    }
                )
    return issues


def _typography_issues(raw_report: Mapping[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    thresholds = {"title": 28.0, "chart": 15.0, "content": 18.0}
    for slide in raw_report.get("slides", []):
        if not isinstance(slide, Mapping):
            continue
        slide_id = str(slide.get("slide_id", "")).strip()
        for shape in slide.get("text_shapes", []):
            if not isinstance(shape, Mapping):
                continue
            role = str(shape.get("shape_role", "content"))
            if role in {"footer", "page_number"} or shape.get("footer_like"):
                continue
            minimum = thresholds.get(role, 18.0)
            try:
                font_size = float(shape.get("font_size_pt"))
            except (TypeError, ValueError):
                continue
            if font_size > 0 and font_size + 0.01 < minimum:
                issues.append(
                    {
                        "slide_id": slide_id,
                        "severity": "error",
                        "code": "MINIMUM_FONT_SIZE_VIOLATION",
                    }
                )
    return issues


def _write_changed_qa(path: Path, report: Mapping[str, Any]) -> None:
    lines = [
        "# Changed-slide QA",
        "",
        f"- Scope: changed slides only ({report.get('checked_slide_count', 0)})",
        f"- Status: `{report.get('status', 'UNKNOWN')}`",
        f"- Blocking issues: `{report.get('blocking_issue_count', 0)}`",
        "",
    ]
    for issue in report.get("issues", []):
        if isinstance(issue, Mapping):
            lines.append(
                f"- [{issue.get('severity', 'error')}] {issue.get('slide_id', '')}: "
                f"{issue.get('code', 'ISSUE')}"
            )
    if not report.get("issues"):
        lines.append("- No automated changed-slide blocking issue detected.")
    lines.extend(["", "Automated QA does not replace user scientific approval.", ""])
    path.write_text("\n".join(lines), encoding="utf-8")


def _write_manual_checklist(
    path: Path, metadata: Mapping[str, Any], *, final_delivery: bool
) -> None:
    declared = metadata.get("manual_review_items", []) or []
    if not isinstance(declared, list):
        raise FastEnhanceError("manual_review_items must be a list")
    lines = [
        "# Manual Review Checklist",
        "",
        "- [ ] Confirm changed-slide claims and numbers against registered sources.",
        "- [ ] Confirm uncertainty, conflicts, and unresolved markers remain explicit.",
        "- [ ] Review every changed slide in PowerPoint at presentation scale.",
        "- [ ] Confirm no unchanged user-authored slide was altered.",
    ]
    for item in declared:
        text = (
            str(item.get("text") or item.get("comment") or item.get("issue") or "").strip()
            if isinstance(item, Mapping)
            else str(item).strip()
        )
        if text:
            lines.append(f"- [ ] {text}")
    if final_delivery:
        lines.append("- [ ] Provide explicit final scientific and visual approval.")
    lines.extend(["", "User approval is mandatory for final delivery.", ""])
    path.write_text("\n".join(lines), encoding="utf-8")


def _libreoffice_page_count(
    pptx_path: Path, output_dir: Path, soffice: str | None, timeout: int
) -> dict[str, Any]:
    executable = None
    if soffice:
        explicit = Path(soffice)
        executable = str(explicit) if explicit.is_file() else shutil.which(soffice)
    if not executable:
        executable = shutil.which("soffice.com") or shutil.which("soffice")
    if not executable:
        raise FastEnhanceError("LibreOffice was explicitly requested but is unavailable")
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="academic-ppt-fast-lo-") as profile:
        command = [
            executable,
            f"-env:UserInstallation={Path(profile).resolve().as_uri()}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(pptx_path),
        ]
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    pdf = output_dir / f"{pptx_path.stem}.pdf"
    if completed.returncode != 0 or not pdf.is_file() or pdf.stat().st_size == 0:
        raise FastEnhanceError("LibreOffice cross-renderer export failed")
    return {"status": "PASS", "pdf": str(pdf), "page_count": len(PdfReader(str(pdf)).pages)}


def _merge_slide_specs(
    cached: list[Mapping[str, Any]],
    changed: list[Mapping[str, Any]],
    expected_order: list[str],
) -> list[dict[str, Any]]:
    merged = {
        str(item.get("slide_id", "")): copy.deepcopy(dict(item))
        for item in cached
        if isinstance(item, Mapping) and str(item.get("slide_id", "")).strip()
    }
    merged.update({str(item["slide_id"]): copy.deepcopy(dict(item)) for item in changed})
    missing = [slide_id for slide_id in expected_order if slide_id not in merged]
    if missing:
        raise FastEnhanceError("Cache promotion lacks SlideSpec for: " + ", ".join(missing))
    return [merged[slide_id] for slide_id in expected_order]


def _merge_evidence(
    cached: Any,
    changed_claims: list[Mapping[str, Any]],
    changed_relative_paths: set[str],
    cached_state: Mapping[str, Any],
) -> list[dict[str, Any]]:
    existing = [dict(item) for item in cached if isinstance(item, Mapping)] if isinstance(cached, list) else []
    relative_by_id, _ = _source_registry_index(cached_state)
    retained = [
        row
        for row in existing
        if relative_by_id.get(str(row.get("source_id", "")), "") not in changed_relative_paths
    ]
    return retained + [dict(item) for item in changed_claims if isinstance(item, Mapping)]


def _refreshed_pptx_asset_state(
    output_pptx: Path,
) -> tuple[
    dict[str, Any],
    dict[str, Any],
    dict[str, Any] | list[dict[str, Any]],
]:
    """Re-scan the updated package; never carry forward stale asset metadata."""

    source_hash = sha256_file(output_pptx)
    try:
        profile = parse_reference_style(output_pptx, mode="style-only")
        master_layout_map: dict[str, Any] = {
            "cache_status": "REFRESHED_FROM_OUTPUT_PPTX",
            "source_pptx_sha256": source_hash,
            "slide_dimensions": profile.get("slide_dimensions"),
            "slide_masters": profile.get("slide_masters", []),
            "slide_layouts": profile.get("slide_layouts", []),
            "common_page_roles": profile.get("common_page_roles", []),
            "footer_and_page_number_rules": profile.get(
                "footer_and_page_number_rules", {}
            ),
        }
        font_map: dict[str, Any] = {
            "cache_status": "REFRESHED_FROM_OUTPUT_PPTX",
            "source_pptx_sha256": source_hash,
            "theme_fonts": profile.get("theme_fonts", "UNKNOWN"),
        }
        images: list[dict[str, Any]] = []
        with zipfile.ZipFile(output_pptx, mode="r") as package:
            for name in sorted(package.namelist()):
                if not name.startswith("ppt/media/") or name.endswith("/"):
                    continue
                payload = package.read(name)
                images.append(
                    {
                        "part": name,
                        "sha256": hashlib.sha256(payload).hexdigest(),
                        "size_bytes": len(payload),
                    }
                )
        return master_layout_map, font_map, images
    except Exception as exc:
        # Honest invalidation is safer than reusing a map bound to the previous
        # package.  No exception prose or private path is persisted.
        marker = {
            "cache_status": "INVALIDATED_AFTER_FAST_UPDATE",
            "source_pptx_sha256": source_hash,
            "reason_code": type(exc).__name__,
        }
        return dict(marker), dict(marker), dict(marker)


def _promote_fast_cache(
    *,
    cache: ProjectCache,
    cached_state: Mapping[str, Any],
    current_source_manifest: Mapping[str, Any],
    current_registry: list[Mapping[str, Any]],
    parsed_objects: Mapping[str, Any],
    evidence_registry: list[Mapping[str, Any]],
    slide_specs: list[Mapping[str, Any]],
    qa_status: Mapping[str, Any],
    output_pptx: Path,
    contract_fingerprints: Mapping[str, str],
) -> dict[str, Any]:
    promoted_graph = build_cached_change_impact_graph(
        source_registry=current_registry,
        evidence_registry=evidence_registry,
        slide_specs=slide_specs,
    )
    visual_specs: list[Mapping[str, Any]] = []
    for slide in slide_specs:
        values = slide.get("visual_specs", []) if isinstance(slide, Mapping) else []
        if isinstance(values, list):
            visual_specs.extend(item for item in values if isinstance(item, Mapping))
    master_layout_map, font_map, image_registry = _refreshed_pptx_asset_state(
        output_pptx
    )
    state = build_project_state(
        cache_contract_fingerprints=contract_fingerprints,
        source_manifest=current_source_manifest,
        parsed_document_objects=parsed_objects,
        evidence_registry=evidence_registry,
        slide_specs=slide_specs,
        visual_specs=visual_specs or list(cached_state.get("visual_specs", [])),
        master_layout_map=master_layout_map,
        font_map=font_map,
        image_registry=image_registry,
        source_bindings=slide_specs,
        previous_qa_status=dict(qa_status),
        extra={
            "source_registry": [dict(row) for row in current_registry],
            "change_impact_graph": promoted_graph.to_dict(),
            "pptx_state": {
                "sha256": sha256_file(output_pptx),
                "slide_count": len(slide_specs),
            },
        },
    )
    return cache.commit_generation(state)


def promote_full_validation_cache(
    *,
    repo_root: Path,
    project_identity: str,
    input_root: Path,
    manifest: list[Mapping[str, Any]],
    extracted: Mapping[str, Any],
    claims: list[Mapping[str, Any]],
    deck_ir: Mapping[str, Any],
    storyboard: list[Mapping[str, Any]],
    style_profile: Mapping[str, Any] | None,
    figures: list[Mapping[str, Any]],
    qa_status: Mapping[str, Any],
    output_pptx: Path,
    brief: Mapping[str, Any],
    cache_root: Path | None = None,
    project_root: Path | None = None,
    cache_home: Path | None = None,
    clinical_privacy_mode: bool = False,
) -> dict[str, Any]:
    """Promote one successful full-validation run for later fast updates."""

    # The privacy/local-cache gate must run before ProjectCache can create its
    # key secret or project directory.
    resolved_cache_root = (
        _assert_clinical_cache_is_ignored(
            repo_root,
            cache_root,
            project_root=project_root,
            cache_home=cache_home,
            production=project_root is not None,
        )
        if clinical_privacy_mode
        else resolve_project_cache_root(
            repo_root,
            project_root=project_root,
            explicit_cache_root=cache_root,
            cache_home=cache_home,
            production=project_root is not None,
        )
    )
    contract_fingerprints = build_cache_contract_fingerprints(
        repo_root=repo_root,
        brief=brief,
    )

    source_paths = _source_paths(input_root, manifest)
    source_manifest = build_source_manifest(source_paths, relative_to=input_root)
    slide_specs = [
        dict(item) for item in deck_ir.get("slides", []) if isinstance(item, Mapping)
    ]
    evidence = [dict(item) for item in claims if isinstance(item, Mapping)]
    graph = build_cached_change_impact_graph(
        source_registry=manifest,
        evidence_registry=evidence,
        slide_specs=slide_specs,
    )
    visual_specs: list[Mapping[str, Any]] = []
    for slide in slide_specs:
        if isinstance(slide.get("visual_specs"), list):
            visual_specs.extend(
                item for item in slide["visual_specs"] if isinstance(item, Mapping)
            )
    profile = dict(style_profile or {})
    state = build_project_state(
        cache_contract_fingerprints=contract_fingerprints,
        source_manifest=source_manifest,
        parsed_document_objects=dict(extracted),
        evidence_registry=evidence,
        slide_specs=slide_specs,
        visual_specs=visual_specs,
        master_layout_map=profile,
        font_map=dict(profile.get("theme", {}).get("fonts", {}))
        if isinstance(profile.get("theme"), Mapping)
        else {},
        image_registry=figures,
        source_bindings=storyboard,
        previous_qa_status=dict(qa_status),
        extra={
            "source_registry": [dict(row) for row in manifest],
            "change_impact_graph": graph.to_dict(),
            "deck_ir": dict(deck_ir),
            "pptx_state": {
                "sha256": sha256_file(output_pptx),
                "slide_count": len(slide_specs),
            },
        },
    )
    cache = ProjectCache.from_identity(
        repo_root,
        project_identity,
        cache_root=resolved_cache_root,
        project_root=project_root,
        cache_home=cache_home,
        production=project_root is not None,
        clinical_privacy_mode=clinical_privacy_mode,
    )
    return cache.commit_generation(state)


def execute_fast_enhance(args: Any) -> Path:
    repo_root = resolve_workflow_home(
        _arg(args, "repo_root"), Path(__file__).resolve().parents[2]
    )
    workflow_config = load_yaml_compatible(repo_root / "config" / "workflow.yaml")
    local_cfg = load_local_runtime_config(repo_root)
    raw_brief = _arg(args, "brief", "brief/presentation_brief.yaml")
    brief_path = _resolve_declared(raw_brief, repo_root)
    if brief_path is None or not brief_path.is_file():
        raise FastEnhanceError(f"Brief is missing: {brief_path}")
    brief = load_yaml_compatible(brief_path)
    current_contract_fingerprints = build_cache_contract_fingerprints(
        repo_root=repo_root,
        brief=brief,
    )
    project_name = str(brief.get("project_name", "")).strip()
    if not project_name or project_name == "INFORMATION_REQUIRED":
        raise FastEnhanceError("brief.project_name is required")
    paths_cfg = workflow_config["paths"]
    env_cfg = workflow_config["environment"]
    workspace_home_raw = os.environ.get("PPT_WORKSPACE_HOME") or local_cfg.get("workspace_home")
    project_root = _resolve_declared(_arg(args, "project_root"), repo_root)
    cache_home_raw = os.environ.get("PPT_CACHE_HOME") or local_cfg.get("cache_home")
    cache_home = Path(cache_home_raw).expanduser() if cache_home_raw else None
    workspace_home = Path(workspace_home_raw).expanduser().resolve() if workspace_home_raw else None
    input_root = resolve_runtime_path(cli_value=_arg(args,"input_root"),env_name="PPT_INPUT_HOME",legacy_env_name=env_cfg["input_root"],local_value=local_cfg.get("input_root"),workspace_home=workspace_home,workspace_child="input",repo_root=repo_root,repo_default=paths_cfg["input_root"])
    staging_base = resolve_runtime_path(cli_value=_arg(args,"staging_root"),env_name="PPT_STAGING_HOME",legacy_env_name=env_cfg["staging_root"],local_value=local_cfg.get("staging_root"),workspace_home=workspace_home,workspace_child="staging",repo_root=repo_root,repo_default=paths_cfg["staging_root"])
    output_base = resolve_runtime_path(cli_value=_arg(args,"output_root"),env_name="PPT_OUTPUT_HOME",legacy_env_name=env_cfg["output_root"],local_value=local_cfg.get("output_root"),workspace_home=workspace_home,workspace_child="output",repo_root=repo_root,repo_default=paths_cfg["output_root"])
    project = safe_slug(project_name)
    run_id = str(_arg(args, "run_id") or output_timestamp())
    staging_root = staging_base / project / run_id
    output_dir = output_base / project / run_id
    resume = bool(_arg(args, "resume", False))
    requested_retry = bool(
        str(_arg(args, "retry_slide", "") or "").strip()
        or str(_arg(args, "retry_stage", "") or "").strip()
    )
    if requested_retry and not resume:
        raise FastEnhanceError(
            "Fast retry requires --resume with an existing --run-id"
        )
    if resume:
        if not staging_root.is_dir() or not output_dir.is_dir():
            raise FastEnhanceError("Fast retry requires the existing run-id staging and output directories")
    else:
        ensure_new_directory(staging_root)
        ensure_new_directory(output_dir)

    operation_plan_path = _resolve_declared(_arg(args, "operation_plan"), repo_root)
    existing_pptx = _resolve_declared(_arg(args, "existing_pptx"), repo_root)
    if operation_plan_path is None or existing_pptx is None:
        raise FastEnhanceError("fast_enhance requires --existing-pptx and --operation-plan")
    raw_plan = _load_json(operation_plan_path, "Operation plan")
    plan = validate_operation_plan(raw_plan)
    if Path(plan["source_deck"]["path"]).resolve() != existing_pptx.resolve():
        raise FastEnhanceError("Operation plan source deck differs from --existing-pptx")
    declared_candidate = _resolve_declared(
        _arg(args, "changed_candidate_pptx"), repo_root
    )
    if declared_candidate is not None:
        candidate_paths = {
            Path(str(item["path"])).resolve()
            for item in plan.get("candidate_decks", [])
            if isinstance(item, Mapping) and item.get("path")
        }
        if candidate_paths != {declared_candidate.resolve()}:
            raise FastEnhanceError(
                "--changed-candidate-pptx must match the operation plan's sole candidate deck"
            )
    metadata = _operation_metadata(raw_plan)
    changed_ids = [str(item) for item in plan["changed_slide_ids"]]
    controls = validate_fast_runtime_controls(args, changed_ids)
    retry_slide = str(controls["retry_slide"])
    requested_retry_stage = str(controls["retry_stage"])
    retry_stage = requested_retry_stage or (
        "changed_slide_generation" if retry_slide else ""
    )
    retry_policy = retry_stage_execution_policy(retry_stage)
    final_delivery = bool(controls["final_delivery"])
    manual_approved = bool(controls["manual_review_approved"])

    presentation_type = str(brief.get("presentation_type", "")).casefold()
    clinical_privacy_mode = bool(
        _arg(args, "clinical_privacy_mode", False)
        or brief.get("clinical_privacy_mode") is True
        or "clinical" in presentation_type
        or "mdt" in presentation_type
    )
    profiler = RuntimeProfiler(
        workflow_mode="fast_enhance",
        original_slide_count=len(plan["source_slide_ids"]),
        changed_slide_count=len(changed_ids),
        source_count=0,
    )
    # This runtime consumes an already-reviewed changed-slide operation plan;
    # it performs deterministic cache, OOXML, PowerPoint, and QA work only.
    profiler.record_model_telemetry(
        model_calls=0,
        input_tokens=0,
        output_tokens=0,
        model_elapsed_seconds=0,
    )
    runtime_paths: dict[str, Path | None] = {}
    status = "BLOCKED"
    try:
        with profiler.stage("preflight"):
            if _arg(args, "route", "enhance-existing") != "enhance-existing":
                raise FastEnhanceError("fast_enhance requires route enhance-existing")
            if not input_root.is_dir():
                raise FastEnhanceError(f"Input root does not exist: {input_root}")
            cache_root = _resolve_declared(_arg(args, "project_cache_root"), repo_root)
            if clinical_privacy_mode:
                cache_root = _assert_clinical_cache_is_ignored(
                    repo_root,
                    cache_root,
                    project_root=project_root,
                    cache_home=cache_home,
                    production=project_root is not None,
                )
            cache = ProjectCache.from_identity(
                repo_root,
                project_name,
                cache_root=cache_root,
                project_root=project_root,
                cache_home=cache_home,
                production=project_root is not None,
                clinical_privacy_mode=clinical_privacy_mode,
            )

        with profiler.stage("source_snapshot") as stage:
            manifest = inventory_sources(
                input_root,
                set(workflow_config["supported_extensions"]),
                staging_root / "current_source_manifest.csv",
                reference_mode=str(_arg(args, "reference_mode", "style-only")),
            )
            if not manifest:
                raise FastEnhanceError("No supported input sources were found")
            stage.add_counts(files_snapshotted=len(manifest))
            current_paths = _source_paths(input_root, manifest)
            current_source_manifest = build_source_manifest(current_paths, relative_to=input_root)
            profiler.source_count = len(manifest)

        with profiler.stage("cache_load") as stage:
            cached_state = cache.load_current()
            if cached_state is None:
                raise FastEnhanceError(
                    "No validated project cache exists; run workflow_mode=full_validation first"
                )
            try:
                validate_project_state(cached_state)
            except CacheStateError as exc:
                raise FastEnhanceError(
                    "Cached project state is incompatible; run "
                    "workflow_mode=full_validation"
                ) from exc
            validate_fast_cache_baseline(
                cached_state,
                source_pptx=existing_pptx,
                expected_slide_count=len(plan["source_slide_ids"]),
                current_contract_fingerprints=current_contract_fingerprints,
            )
            _validate_cached_source_slide_order(
                cached_state, plan["source_slide_ids"]
            )
            stage.add_counts(cache_generations_loaded=1)
            cache_generation = str(cache.current_generation_id() or "")
            if not cache_generation:
                raise FastEnhanceError("Validated cache has no current generation")

        with profiler.stage("delta_classification") as stage:
            delta = cache.compute_delta(current_source_manifest)
            write_json(staging_root / "delta_manifest.json", delta)
            stage.add_counts(
                unchanged_files=int(delta["counts"].get("UNCHANGED", 0)),
                changed_files=int(delta["counts"].get("MODIFIED", 0))
                + int(delta["counts"].get("NEW", 0)),
                removed_files=int(delta["counts"].get("REMOVED", 0)),
            )

        checkpoint = _checkpoint_path(staging_root)
        retry_artifacts: dict[str, Path] = {}
        if retry_stage:
            retry_artifacts = _load_retry_checkpoint(
                checkpoint,
                retry_stage=retry_stage,
                existing_pptx=existing_pptx,
                operation_plan_path=operation_plan_path,
                cache_generation=cache_generation,
                current_source_manifest=current_source_manifest,
                final_delivery=final_delivery,
                cross_renderer_validation=bool(controls["cross_renderer_validation"]),
                output_dir=output_dir,
                staging_root=staging_root,
            )
            _load_json(retry_artifacts["retry_context"], "Private retry context")
            cached_parsed = cached_state.get("parsed_document_objects", {})
            if not isinstance(cached_parsed, Mapping):
                raise FastEnhanceError("Cached parsed_document_objects is invalid")
            changed_relatives = set(delta["parse_source_keys"])
            changed_source_ids = {
                str(row["source_id"])
                for row in manifest
                if str(row.get("relative_path", "")) in changed_relatives
            }
            parsed_objects = dict(cached_parsed)
            for source_id in changed_source_ids:
                extraction = staging_root / "extracted" / f"{source_id}.txt"
                if not extraction.is_file():
                    raise FastEnhanceError(
                        "Retry extraction is missing; full fast run required"
                    )
                parsed_objects[source_id] = extraction.read_text(encoding="utf-8")
            evidence_path = staging_root / "evidence_delta" / "evidence_inventory.csv"
            changed_claims = read_csv(evidence_path) if evidence_path.is_file() else []
            unresolved = []
            cached_specs = cached_state.get("slide_specs", [])
            if not isinstance(cached_specs, list):
                raise FastEnhanceError("Cached SlideSpec collection is invalid")
            changed_specs = _changed_slide_specs(
                metadata, changed_ids
            )
            merged_specs = _merge_slide_specs(
                cached_specs, changed_specs, list(plan["expected_final_order"])
            )
            for name in (
                "source_parse",
                "evidence_rebuild",
                "impact_resolution",
                "slide_planning",
            ):
                profiler.mark_skipped(name)
        else:
            cached_parsed = cached_state.get("parsed_document_objects", {})
            if not isinstance(cached_parsed, Mapping):
                raise FastEnhanceError("Cached parsed_document_objects is invalid")
            changed_relatives = set(delta["parse_source_keys"])
            changed_source_ids = {
                str(row["source_id"])
                for row in manifest
                if str(row.get("relative_path", "")) in changed_relatives
            }
            with profiler.stage("source_parse") as stage:
                parsed_objects, parse_stats = extract_selected(
                    input_root,
                    staging_root,
                    manifest,
                    changed_source_ids,
                    cached_extractions=dict(cached_parsed),
                )
                if parse_stats["missing"]:
                    raise FastEnhanceError(
                        "Validated cache is missing an unchanged parsed source; full validation required"
                    )
                stage.add_counts(
                    files_parsed=parse_stats["parsed"],
                    files_reused=parse_stats["reused"],
                )
            with profiler.stage("evidence_rebuild") as stage:
                changed_manifest = [
                    row for row in manifest
                    if str(row.get("source_id", "")) in changed_source_ids
                ]
                if changed_manifest:
                    (staging_root / "evidence_delta").mkdir(parents=True, exist_ok=True)
                    changed_extracted = {
                        source_id: parsed_objects[source_id]
                        for source_id in changed_source_ids
                        if source_id in parsed_objects
                    }
                    changed_claims, unresolved = build_evidence(
                        changed_extracted,
                        changed_manifest,
                        staging_root / "evidence_delta",
                        brief,
                    )
                else:
                    changed_claims, unresolved = [], []
                    stage.status = "SKIPPED"
                stage.add_counts(
                    claims_rebuilt=len(changed_claims),
                    unresolved_items=len(unresolved),
                )
            with profiler.stage("impact_resolution") as stage:
                graph_document = cached_state.get("change_impact_graph")
                if not isinstance(graph_document, Mapping):
                    raise FastEnhanceError("Validated cache has no ChangeImpactGraph")
                graph = ChangeImpactGraph.from_dict(graph_document)
                source_impacts = metadata.get("source_impacts", []) or []
                if not isinstance(source_impacts, list) or any(
                    not isinstance(item, Mapping) for item in source_impacts
                ):
                    raise FastEnhanceError("source_impacts must be a list of objects")
                relative_by_current_id = {
                    str(row.get("source_id", "")): str(row.get("relative_path", ""))
                    for row in manifest
                    if str(row.get("source_id", "")).strip()
                    and str(row.get("relative_path", "")).strip()
                }
                _extend_graph_from_declared_impacts(
                    graph,
                    source_impacts,
                    allowed_slides=set(plan["expected_final_order"]),
                    relative_by_source_id=relative_by_current_id,
                )
                impact = graph.analyze_delta(delta, changed_slides=changed_ids)
                write_json(staging_root / "change_impact_result.json", impact.to_dict())
                if impact.requires_full_rebuild:
                    raise FastEnhanceError(
                        "Change impact is not completely bounded; use workflow_mode=full_validation: "
                        + ", ".join(impact.escalation_reasons)
                    )
                if not set(impact.affected_slides).issubset(set(changed_ids)):
                    raise FastEnhanceError(
                        "Operation plan omits one or more source-affected slides"
                    )
                stage.add_counts(affected_slides=len(impact.affected_slides))
            with profiler.stage("slide_planning") as stage:
                cached_specs = cached_state.get("slide_specs", [])
                if not isinstance(cached_specs, list):
                    raise FastEnhanceError("Cached SlideSpec collection is invalid")
                changed_specs = _changed_slide_specs(
                    metadata, changed_ids
                )
                merged_specs = _merge_slide_specs(
                    cached_specs, changed_specs, list(plan["expected_final_order"])
                )
                stage.add_counts(
                    slides_reused=len(merged_specs) - len(changed_specs),
                    slides_planned=len(changed_specs),
                )
            write_json(
                staging_root / "fast_retry_context.json",
                {
                    "schema_version": "academic-ppt-fast-retry-context/1",
                    "changed_source_count": len(changed_source_ids),
                    "changed_claim_count": len(changed_claims),
                    "changed_slide_count": len(changed_specs),
                    "contains_source_text": False,
                    "contains_private_paths": False,
                },
            )
            _write_checkpoint(
                checkpoint,
                completed_stage="slide_planning",
                existing_pptx=existing_pptx,
                operation_plan_path=operation_plan_path,
                cache_generation=cache_generation,
                current_source_manifest=current_source_manifest,
                final_delivery=final_delivery,
                cross_renderer_validation=bool(controls["cross_renderer_validation"]),
            )

        _validate_changed_spec_source_registry(changed_specs, manifest)

        if not retry_policy["run_candidate_validation"]:
            profiler.mark_skipped("changed_slide_generation")
        else:
            with profiler.stage("changed_slide_generation") as stage:
                # The reviewed operation plan contains only candidate slides.
                for candidate in plan["candidate_decks"]:
                    path = Path(str(candidate["path"]))
                    if sha256_file(path) != str(candidate["sha256"]):
                        raise FastEnhanceError("Changed-slide candidate hash mismatch")
                stage.add_counts(
                    candidate_decks=len(plan["candidate_decks"]),
                    slides_generated=len(changed_ids),
                )
            _write_checkpoint(
                checkpoint,
                completed_stage="changed_slide_generation",
                existing_pptx=existing_pptx,
                operation_plan_path=operation_plan_path,
                cache_generation=cache_generation,
                current_source_manifest=current_source_manifest,
                final_delivery=final_delivery,
                cross_renderer_validation=bool(controls["cross_renderer_validation"]),
            )

        change_plan_path = output_dir / "change_plan.json"
        write_json(change_plan_path, raw_plan)
        updated_pptx = output_dir / "updated.pptx"
        # Previews are QA intermediates by default.  They stay in staging so
        # the formal fast package remains the six-file minimal contract.
        preview_dir = staging_root / "changed_preview"
        pdf_path = output_dir / "updated.pdf" if (final_delivery or bool(_arg(args, "export_pdf", False))) else None
        reuse_powerpoint = not retry_policy["run_powerpoint_apply"]
        if reuse_powerpoint:
            operation_result = _load_json(
                staging_root / "operation_result.json", "Operation result"
            )
            profiler.mark_skipped("powerpoint_apply", powerpoint_calls=0)
        else:
            if retry_stage == "powerpoint_apply":
                _quarantine_failed_powerpoint_outputs(
                    output_dir=output_dir,
                    staging_root=staging_root,
                    preview_dir=preview_dir,
                )
            with profiler.stage("powerpoint_apply") as stage:
                operation_result = run_powerpoint_operations(
                    source_pptx=existing_pptx,
                    output_pptx=updated_pptx,
                    operation_plan_path=operation_plan_path,
                    script_path=repo_root / "scripts" / "apply_pptx_operations.ps1",
                    preview_dir=preview_dir,
                    export_pdf=pdf_path,
                    timeout_seconds=int(workflow_config.get("render_timeout_seconds", 120)) + 60,
                )
                stage.add_counts(powerpoint_calls=1, operations_applied=operation_result["applied_operations"])
                write_json(staging_root / "operation_result.json", operation_result)
            _write_checkpoint(
                checkpoint,
                completed_stage="powerpoint_apply",
                existing_pptx=existing_pptx,
                operation_plan_path=operation_plan_path,
                cache_generation=cache_generation,
                current_source_manifest=current_source_manifest,
                final_delivery=final_delivery,
                cross_renderer_validation=bool(controls["cross_renderer_validation"]),
                updated_pptx=updated_pptx,
            )

        final_order = list(plan["expected_final_order"])
        final_index = {slide_id: index for index, slide_id in enumerate(final_order, 1)}
        geometry_scope = None if final_delivery else [final_index[item] for item in changed_ids]
        reuse_geometry = not retry_policy["run_powerpoint_geometry"]
        if reuse_geometry:
            raw_geometry = _load_json(
                staging_root / "powerpoint_geometry.json", "PowerPoint geometry"
            )
            geometry = _geometry_gate(raw_geometry)
            profiler.mark_skipped(
                "changed_slide_render", powerpoint_calls=0, slides_rendered=0
            )
        else:
            with profiler.stage("changed_slide_render") as stage:
                raw_geometry, geometry_log = inspect_powerpoint_text_layout(
                    pptx_path=updated_pptx,
                    output_json=staging_root / "powerpoint_geometry.json",
                    script_path=repo_root / "scripts" / "inspect_pptx_layout.ps1",
                    timeout_seconds=int(workflow_config.get("render_timeout_seconds", 120)) + 60,
                    slide_indexes=geometry_scope,
                )
                geometry = _geometry_gate(raw_geometry)
                write_json(staging_root / "powerpoint_geometry_gate.json", geometry)
                stage.add_counts(
                    powerpoint_calls=1,
                    slides_rendered=(len(final_order) if final_delivery else len(changed_ids)),
                    pdf_exports=1 if pdf_path else 0,
                )
            _write_checkpoint(
                checkpoint,
                completed_stage="changed_slide_render",
                existing_pptx=existing_pptx,
                operation_plan_path=operation_plan_path,
                cache_generation=cache_generation,
                current_source_manifest=current_source_manifest,
                final_delivery=final_delivery,
                cross_renderer_validation=bool(controls["cross_renderer_validation"]),
                updated_pptx=updated_pptx,
                geometry_path=staging_root / "powerpoint_geometry.json",
            )

        if pdf_path:
            profiler.record_stage("optional_pdf", 0.0, pdf_exports=1)
        else:
            profiler.mark_skipped("optional_pdf", pdf_exports=0)

        if not retry_policy["run_changed_slide_qa"]:
            changed_qa = _load_json(
                staging_root / "changed_slide_qa.json", "Changed-slide QA"
            )
            if changed_qa.get("status") != "PASS":
                raise FastEnhanceError("Checkpoint changed-slide QA is not passing")
            profiler.mark_skipped("changed_slide_qa", slides_checked=0)
        else:
            with profiler.stage("changed_slide_qa") as stage:
                source_binding_issues = metadata.get("source_binding_issues", []) or []
                if not isinstance(source_binding_issues, list):
                    raise FastEnhanceError("source_binding_issues must be a list")
                scientific_issues = _scientific_issue_rows(metadata, changed_specs)
                visual_issues = _preview_visual_issues(
                    preview_dir, changed_ids, final_index
                )
                typography_issues = _typography_issues(raw_geometry)
                changed_qa = run_changed_slide_qa(
                    changed_slide_ids=changed_ids,
                    slide_specs=changed_specs,
                    source_binding_issues=source_binding_issues,
                    scientific_issues=scientific_issues,
                    powerpoint_geometry=geometry,
                    typography_issues=typography_issues,
                    visual_issues=visual_issues,
                )
                if changed_qa["status"] != "PASS":
                    raise FastEnhanceError("Changed-slide QA failed")
                write_json(staging_root / "changed_slide_qa.json", changed_qa)
                _write_changed_qa(output_dir / "changed_slide_qa.md", changed_qa)
                stage.add_counts(slides_checked=len(changed_ids), blocking_issues=0)
            _write_checkpoint(
                checkpoint,
                completed_stage="changed_slide_qa",
                existing_pptx=existing_pptx,
                operation_plan_path=operation_plan_path,
                cache_generation=cache_generation,
                current_source_manifest=current_source_manifest,
                final_delivery=final_delivery,
                cross_renderer_validation=bool(controls["cross_renderer_validation"]),
                updated_pptx=updated_pptx,
                geometry_path=staging_root / "powerpoint_geometry.json",
            )

        with profiler.stage("whole_deck_light_qa") as stage:
            light_qa = run_whole_deck_lightweight_qa(
                updated_pptx,
                expected_slide_count=len(final_order),
                expected_final_order=final_order,
                operation_result=operation_result,
            )
            if light_qa["status"] != "PASS":
                raise FastEnhanceError("Whole-deck lightweight QA failed")
            write_json(staging_root / "whole_deck_lightweight_qa.json", light_qa)
            if final_delivery:
                full_geometry_issues = _blocking_geometry_messages(raw_geometry)
                if full_geometry_issues:
                    raise FastEnhanceError("Final-delivery whole-deck PowerPoint geometry QA failed")
            stage.add_counts(slides_checked=len(final_order))
        _write_checkpoint(
            checkpoint,
            completed_stage="whole_deck_light_qa",
            existing_pptx=existing_pptx,
            operation_plan_path=operation_plan_path,
            cache_generation=cache_generation,
            current_source_manifest=current_source_manifest,
            final_delivery=final_delivery,
            cross_renderer_validation=bool(controls["cross_renderer_validation"]),
            updated_pptx=updated_pptx,
            geometry_path=staging_root / "powerpoint_geometry.json",
        )

        if bool(controls["cross_renderer_validation"]):
            with profiler.stage("optional_cross_renderer") as stage:
                lo = _libreoffice_page_count(
                    updated_pptx,
                    staging_root / "libreoffice",
                    _arg(args, "soffice") or os.environ.get(env_cfg["soffice_executable"]),
                    int(workflow_config.get("render_timeout_seconds", 120)),
                )
                if int(lo["page_count"]) != len(final_order):
                    raise FastEnhanceError("LibreOffice page count differs from PowerPoint")
                write_json(staging_root / "libreoffice_qa.json", lo)
                stage.add_counts(libreoffice_calls=1, pages_rendered=lo["page_count"])
        else:
            profiler.mark_skipped("optional_cross_renderer", libreoffice_calls=0)

        input_errors = verify_input_hashes(input_root, manifest)
        if input_errors:
            raise FastEnhanceError("Input integrity failed: " + "; ".join(input_errors))

        with profiler.stage("package"):
            slide_diff = {
                "schema_version": "1.0",
                "workflow_mode": "fast_enhance",
                "unchanged_slide_count": sum(
                    item["action"] == "KEEP" for item in plan["operations"]
                ),
                "changed_slides": [
                    {
                        "slide_id": item["slide_id"],
                        "action": item["action"],
                        "slide_revision": item.get("slide_revision"),
                        "content_hash": item.get("content_hash", ""),
                    }
                    for item in plan["operations"]
                    if item["action"] != "KEEP"
                ],
            }
            write_json(output_dir / "slide_diff.json", slide_diff)
            _write_manual_checklist(
                output_dir / "manual_review_checklist.md",
                metadata,
                final_delivery=final_delivery,
            )
            if final_delivery and not manual_approved:
                status = "DRAFT_READY"
            elif final_delivery:
                status = "FINAL_DELIVERY_READY"
            else:
                status = "DRAFT_READY"

        current_registry = [dict(row) for row in manifest]
        evidence = _merge_evidence(
            cached_state.get("evidence_registry", []),
            changed_claims,
            changed_relatives,
            cached_state,
        )
        with profiler.stage("cache_promote") as stage:
            pointers = _promote_fast_cache(
                cache=cache,
                cached_state=cached_state,
                current_source_manifest=current_source_manifest,
                current_registry=current_registry,
                parsed_objects=parsed_objects,
                evidence_registry=evidence,
                slide_specs=merged_specs,
                qa_status={"status": status, "changed_slide_qa": "PASS", "whole_deck_lightweight_qa": "PASS"},
                output_pptx=updated_pptx,
                contract_fingerprints=current_contract_fingerprints,
            )
            stage.add_counts(cache_generations_promoted=1)

        runtime_paths, profile, audit_dir = _write_runtime_outputs(
            profiler,
            staging_root=staging_root,
            output_dir=output_dir,
            audit_full=bool(controls["audit_full"]),
        )
        if audit_dir is not None:
            for path in (
                staging_root / "current_source_manifest.csv",
                staging_root / "delta_manifest.json",
                staging_root / "change_impact_result.json",
                staging_root / "operation_result.json",
                staging_root / "powerpoint_geometry.json",
                staging_root / "powerpoint_geometry_gate.json",
                staging_root / "whole_deck_lightweight_qa.json",
            ):
                if path.is_file():
                    shutil.copy2(path, audit_dir / path.name)
        stage_by_name = {row["name"]: row for row in profile.get("stages", [])}
        parse_attempts = stage_by_name.get("source_parse", {}).get("attempts", [])
        parse_counts = parse_attempts[-1].get("counts", {}) if parse_attempts else {}
        def runtime_count(stage_name: str, counter: str) -> int:
            attempts = stage_by_name.get(stage_name, {}).get("attempts", [])
            return int(
                sum(
                    float(item.get("counts", {}).get(counter, 0) or 0)
                    for item in attempts
                )
            )
        powerpoint_calls = runtime_count("powerpoint_apply", "powerpoint_calls") + runtime_count(
            "changed_slide_render", "powerpoint_calls"
        )
        libreoffice_calls = runtime_count(
            "optional_cross_renderer", "libreoffice_calls"
        )
        summary = [
            "# Fast Enhance Execution Summary",
            "",
            f"- Status: `{status}`",
            "- Workflow mode: `fast_enhance`",
            f"- Original slides: `{len(plan['source_slide_ids'])}`",
            f"- Changed/new slides: `{len(changed_ids)}`",
            f"- Unchanged slides reused: `{slide_diff['unchanged_slide_count']}`",
            f"- Files parsed: `{parse_counts.get('files_parsed', 0)}`",
            f"- Files reused from cache: `{parse_counts.get('files_reused', 0)}`",
            f"- PowerPoint calls: `{powerpoint_calls}`",
            f"- LibreOffice calls: `{libreoffice_calls}`",
            f"- Total runtime: `{profile['total_seconds']}` seconds",
            "- Model calls: `0` (deterministic fast runtime)",
            "- Model tokens: `0`",
            f"- Retry request: `{retry_slide or requested_retry_stage or 'none'}`",
            "- Input hashes unchanged: `yes`",
            "- Cache policy: `local_private_cache_only`",
            "",
            "FINAL_DELIVERY_READY requires explicit user manual approval.",
            "",
        ]
        (output_dir / "execution_summary.md").write_text("\n".join(summary), encoding="utf-8")
        return output_dir
    except Exception as exc:
        try:
            runtime_paths = profiler.write(staging_root)
            failure = _build_failure_record(
                exc, clinical_privacy_mode=clinical_privacy_mode
            )
            write_json(staging_root / "fast_enhance_failure.json", failure)
        except Exception:
            pass
        raise


def execute_fast_production(args: Any) -> Path:
    """Close the candidate-free changed-source to draft-ready fast chain."""

    from time import perf_counter
    from .fast_production import (
        FastStateMachine,
        build_automatic_operation_plan,
        build_keep_baseline_fingerprints,
        build_changed_slide_plans,
        candidate_attempt_record,
        content_review_status,
        generate_candidate_deck,
        preflight_changed_slide_plans,
        validate_keep_baseline_fingerprints,
        write_changed_content_review,
    )

    wall_started = perf_counter()
    repo_root = Path(_arg(args, "repo_root")).resolve() if _arg(args, "repo_root") else Path(__file__).resolve().parents[2]
    config = load_yaml_compatible(repo_root / "config" / "workflow.yaml")
    brief_path = _resolve_declared(_arg(args, "brief", "brief/presentation_brief.yaml"), repo_root)
    existing_pptx = _resolve_declared(_arg(args, "existing_pptx"), repo_root)
    if brief_path is None or not brief_path.is_file() or existing_pptx is None:
        raise FastEnhanceError("fast_enhance requires --ppt and --brief")
    if not existing_pptx.is_file() or existing_pptx.suffix.lower() != ".pptx":
        raise FastEnhanceError("--ppt must identify an existing local PPTX")
    brief = load_yaml_compatible(brief_path)
    project_name = str(brief.get("project_name", "")).strip()
    if not project_name or project_name == "INFORMATION_REQUIRED":
        raise FastEnhanceError("brief.project_name is required")
    paths_cfg, env_cfg = config["paths"], config["environment"]
    input_root = resolve_path(_arg(args, "input_root"), env_cfg["input_root"], paths_cfg["input_root"], repo_root, "input")
    staging_base = resolve_path(_arg(args, "staging_root"), env_cfg["staging_root"], paths_cfg["staging_root"], repo_root, "staging")
    output_base = resolve_path(_arg(args, "output_root"), env_cfg["output_root"], paths_cfg["output_root"], repo_root, "output")
    project = safe_slug(project_name)
    run_id = str(_arg(args, "run_id") or output_timestamp())
    staging_root, gate_output = staging_base / project / run_id, output_base / project / run_id
    if staging_root.exists() or gate_output.exists():
        raise FastEnhanceError("Automatic production requires a new run-id")
    ensure_new_directory(staging_root)
    ensure_new_directory(gate_output)

    presentation_type = str(brief.get("presentation_type", "")).casefold()
    clinical_privacy_mode = bool(_arg(args, "clinical_privacy_mode", False) or brief.get("clinical_privacy_mode") is True or "clinical" in presentation_type or "mdt" in presentation_type)
    project_root = _resolve_declared(_arg(args, "project_root"), repo_root)
    cache_home_raw = os.environ.get("PPT_CACHE_HOME")
    cache_home = Path(cache_home_raw).expanduser() if cache_home_raw else None
    cache_root = _resolve_declared(_arg(args, "project_cache_root"), repo_root)
    if clinical_privacy_mode:
        cache_root = _assert_clinical_cache_is_ignored(
            repo_root,
            cache_root,
            project_root=project_root,
            cache_home=cache_home,
            production=project_root is not None,
        )
    cache = ProjectCache.from_identity(
        repo_root,
        project_name,
        cache_root=cache_root,
        project_root=project_root,
        cache_home=cache_home,
        production=project_root is not None,
        clinical_privacy_mode=clinical_privacy_mode,
    )
    cached_state = cache.load_current()
    if cached_state is None:
        raise FastEnhanceError("No payload/2 cache exists; run full_validation once")
    try:
        validate_project_state(cached_state)
    except CacheStateError as exc:
        write_json(staging_root / "cache_migration_report.json", {
            "schema_version": "academic-ppt-cache-migration-report/1",
            "status": "CACHE_SCHEMA_MISMATCH", "legacy_cache_reused": False,
            "action": "Run full_validation to create an isolated payload/2 generation.",
        })
        raise FastEnhanceError("CACHE_SCHEMA_MISMATCH; full_validation required") from exc
    cached_specs = cached_state.get("slide_specs", [])
    if not isinstance(cached_specs, list):
        raise FastEnhanceError("Validated cache has no SlideSpec collection")
    validate_fast_cache_baseline(
        cached_state, source_pptx=existing_pptx,
        expected_slide_count=len(cached_specs),
        current_contract_fingerprints=build_cache_contract_fingerprints(repo_root=repo_root, brief=brief),
    )
    manifest = inventory_sources(input_root, set(config["supported_extensions"]), staging_root / "current_source_manifest.csv", reference_mode=str(_arg(args, "reference_mode", "style-only")))
    if not manifest:
        raise FastEnhanceError("No supported input sources were found")
    current_manifest = build_source_manifest(_source_paths(input_root, manifest), relative_to=input_root)
    input_hash_bundle = {
        row["relative_path"]: str(row["sha256"]).lower() for row in manifest
    }
    machine = FastStateMachine(staging_root / "fast_state_machine.json", input_hash_bundle)
    machine.checkpoint("PREFLIGHT", {"brief": sha256_file(brief_path).lower(), "pptx": sha256_file(existing_pptx).lower()})
    machine.checkpoint("CACHE_VALIDATE", {"cache_state": hashlib.sha256(json.dumps(cached_state, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()})
    delta = cache.compute_delta(current_manifest)
    write_json(staging_root / "delta_manifest.json", delta)
    machine.checkpoint("DELTA_DISCOVERY", {"delta_manifest": sha256_file(staging_root / "delta_manifest.json").lower()})
    changed_relatives = set(delta["parse_source_keys"])
    if not changed_relatives:
        raise FastEnhanceError("No NEW or MODIFIED input source was detected")
    changed_source_ids = {str(row["source_id"]) for row in manifest if str(row.get("relative_path", "")) in changed_relatives}
    cached_parsed = cached_state.get("parsed_document_objects", {})
    if not isinstance(cached_parsed, Mapping):
        raise FastEnhanceError("Cached parsed_document_objects is invalid")
    parsed_objects, parse_stats = extract_selected(input_root, staging_root, manifest, changed_source_ids, cached_extractions=dict(cached_parsed))
    if parse_stats["missing"]:
        raise FastEnhanceError("Validated cache lacks an unchanged parse; full_validation required")
    changed_manifest = [row for row in manifest if row["source_id"] in changed_source_ids]
    (staging_root / "evidence_delta").mkdir(parents=True, exist_ok=True)
    changed_claims, unresolved = build_evidence({sid: parsed_objects[sid] for sid in changed_source_ids}, changed_manifest, staging_root / "evidence_delta", brief)
    evidence_inventory = staging_root / "evidence_delta" / "evidence_inventory.csv"
    machine.checkpoint("EVIDENCE_PATCH", {"evidence_patch": sha256_file(evidence_inventory).lower()})
    plans = build_changed_slide_plans(
        brief=brief, cached_state=cached_state, source_registry=manifest,
        changed_claims=changed_claims, unresolved_items=unresolved,
        retry_slide=str(_arg(args, "retry_slide", "") or "").strip(),
    )
    time_to_change_plan = round(perf_counter() - wall_started, 6)
    write_json(staging_root / "changed_slide_plans.json", {"schema_version": "academic-ppt-changed-slide-plan/1", "plans": plans})
    machine.checkpoint("CHANGE_PLAN", {"changed_slide_plans": sha256_file(staging_root / "changed_slide_plans.json").lower()})
    review_status = content_review_status(brief=brief, auto_approve_content=bool(_arg(args, "auto_approve_content", False)), clinical_privacy_mode=clinical_privacy_mode)
    review_path = write_changed_content_review(staging_root / "changed_content_review.md", plans, status=review_status)
    machine.checkpoint("CONTENT_REVIEW", {"changed_content_review": sha256_file(review_path).lower()})
    time_to_content_review = round(perf_counter() - wall_started, 6)
    if review_status != "CONTENT_APPROVED":
        shutil.copy2(review_path, gate_output / "changed_content_review.md")
        (gate_output / "execution_summary.md").write_text("# Fast Enhance Execution Summary\n\n- Status: `CONTENT_REVIEW_REQUIRED`\n- Candidate generation: `NOT_RUN`\n", encoding="utf-8")
        return gate_output
    preflight = preflight_changed_slide_plans(plans)
    write_json(staging_root / "preflight_layout_qa.json", preflight)
    machine.checkpoint("PREFLIGHT_LAYOUT_QA", {"preflight_layout_qa": sha256_file(staging_root / "preflight_layout_qa.json").lower()})
    if preflight["status"] != "PASS":
        raise FastEnhanceError("Changed-slide deterministic layout preflight failed")
    candidate_path, candidate_spec, _ = generate_candidate_deck(
        plans=plans, staging_root=staging_root, repo_root=repo_root,
        node_executable=_arg(args, "node"), node_modules=_arg(args, "node_modules"),
        timeout_seconds=int(config.get("render_timeout_seconds", 120)), brief=brief,
    )
    write_json(staging_root / "candidate_attempts.json", candidate_attempt_record(plans, attempt=1))
    machine.checkpoint("CANDIDATE_GENERATION", {"candidate_pptx": sha256_file(candidate_path).lower(), "candidate_spec": sha256_file(candidate_spec).lower()})
    machine.checkpoint("CANDIDATE_RENDER", {"candidate_manifest": sha256_file(staging_root / "changed_candidate_manifest.json").lower()})
    plan_path = staging_root / "operation_plan.json"
    plan = build_automatic_operation_plan(source_pptx=existing_pptx, cached_state=cached_state, plans=plans, candidate_pptx=candidate_path, output_path=plan_path)
    keep_baseline = build_keep_baseline_fingerprints(
        source_pptx=existing_pptx, operation_plan=plan
    )
    write_json(staging_root / "keep_fingerprint_baseline.json", keep_baseline)
    shutil.rmtree(gate_output)
    downstream = copy.copy(args)
    downstream.operation_plan, downstream.changed_candidate_pptx = str(plan_path), str(candidate_path)
    downstream.existing_pptx = str(existing_pptx)
    downstream.staging_root, downstream.output_root = str(staging_base), str(output_base)
    downstream.run_id, downstream.resume = run_id + "_apply", False
    downstream.retry_slide, downstream.retry_stage = None, None
    produced = execute_fast_enhance(downstream)
    updated_pptx = produced / "updated.pptx"
    keep_integrity = validate_keep_baseline_fingerprints(
        baseline=keep_baseline,
        updated_pptx=updated_pptx,
        expected_final_order=plan["expected_final_order"],
        operation_plan=plan,
    )
    write_json(staging_root / "keep_fingerprint_report.json", keep_integrity)
    if keep_integrity["status"] != "PASS":
        raise FastEnhanceError("KEEP slide structural fingerprint changed")
    machine.checkpoint("CHANGED_SLIDE_QA", {"changed_slide_qa": sha256_file(produced / "changed_slide_qa.md").lower()})
    machine.checkpoint("INCREMENTAL_APPLY", {"updated_pptx": sha256_file(updated_pptx).lower()})
    machine.checkpoint("WHOLE_DECK_LIGHT_QA", {"keep_fingerprint_report": sha256_file(staging_root / "keep_fingerprint_report.json").lower()})
    machine.checkpoint("DRAFT_READY", {"execution_summary": sha256_file(produced / "execution_summary.md").lower()})
    elapsed = round(perf_counter() - wall_started, 6)
    runtime = {
        "schema_version": "2.5.2", "runtime_boundary": "changed_source_to_draft_ready",
        "total_wall_time_seconds": elapsed,
        "time_to_change_plan_seconds": time_to_change_plan,
        "time_to_content_review_seconds": time_to_content_review,
        "time_to_first_preview_seconds": elapsed,
        "time_to_draft_ready_seconds": elapsed,
        "time_to_final_delivery_seconds": "NOT_ASSESSED",
        "model_calls": 0, "input_tokens": 0, "output_tokens": 0,
        "candidate_attempts": 1, "fallback_count": 0,
        "changed_slides": len(plans), "keep_slides": sum(row["action"] == "KEEP" for row in plan["operations"]),
        "cache_hit_rate": round(parse_stats["reused"] / max(1, parse_stats["reused"] + parse_stats["parsed"]), 6),
        "content_review_status": review_status, "candidate_spec_sha256": sha256_file(candidate_spec),
        "candidate_pptx_sha256": sha256_file(candidate_path), "input_hashes_unchanged": not verify_input_hashes(input_root, manifest),
    }
    write_json(staging_root / "production_runtime_profile.json", runtime)
    with (produced / "execution_summary.md").open("a", encoding="utf-8") as handle:
        handle.write("\n## Production closure\n\n- Automatic ChangedSlidePlan: `PASS`\n- Prebuilt candidate required: `no`\n" f"- End-to-end wall time: `{elapsed}` seconds\n- Tier 2 model calls: `0`\n")
    return produced


__all__ = [
    "FastEnhanceError",
    "build_cached_change_impact_graph",
    "execute_fast_enhance",
    "execute_fast_production",
    "promote_full_validation_cache",
    "validate_fast_runtime_controls",
]
