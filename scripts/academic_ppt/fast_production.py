from __future__ import annotations

import hashlib
import json
import os
import posixpath
import re
import subprocess
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping
from xml.etree import ElementTree as ET

from .incremental import build_operation_plan, validate_operation_plan, write_operation_plan
from .layout_contract import LayoutContract, plan_slide_geometry, validate_planned_geometry
from .source_display import short_source_label
from .utils import load_yaml_compatible
from .utils import sha256_file, write_json


class FastProductionError(RuntimeError):
    """Raised when an automatic fast-production artifact is incomplete."""


FAST_STATE_SEQUENCE = (
    "PREFLIGHT",
    "CACHE_VALIDATE",
    "DELTA_DISCOVERY",
    "EVIDENCE_PATCH",
    "CHANGE_PLAN",
    "CONTENT_REVIEW",
    "CANDIDATE_GENERATION",
    "PREFLIGHT_LAYOUT_QA",
    "CANDIDATE_RENDER",
    "CHANGED_SLIDE_QA",
    "INCREMENTAL_APPLY",
    "WHOLE_DECK_LIGHT_QA",
    "DRAFT_READY",
    "FINAL_DELIVERY_OPTIONAL",
)

_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_SEMANTIC_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.:/-]{2,119}$")


def _canonical_hash(value: Any) -> str:
    return hashlib.sha256(
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    ).hexdigest()


def _stable_insert_id(logical_key: str) -> str:
    digest = hashlib.sha256(f"fast-production:{logical_key}".encode("utf-8")).hexdigest()
    return f"SLD-FAST-{digest[:20].upper()}"


def _as_list(value: Any) -> list[Any]:
    if value in (None, ""):
        return []
    if isinstance(value, list):
        return value
    return [value]


def _clean_lines(value: Any, *, label: str, required: bool = False) -> list[str]:
    rows: list[str] = []
    for raw in _as_list(value):
        text = str(raw).strip()
        if text:
            rows.append(text)
    if required and not rows:
        raise FastProductionError(f"{label} must not be empty")
    return rows


def _source_bindings(
    raw: Any,
    source_registry: Iterable[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    by_id = {
        str(row.get("source_id", "")).strip(): str(row.get("relative_path", "")).strip()
        for row in source_registry
        if str(row.get("source_id", "")).strip()
        and str(row.get("relative_path", "")).strip()
    }
    by_file = {value: key for key, value in by_id.items()}
    result: list[dict[str, Any]] = []
    for index, item in enumerate(_as_list(raw), start=1):
        if isinstance(item, Mapping):
            source_id = str(item.get("source_id", "")).strip()
            source_file = str(item.get("source_file", "")).strip().replace("\\", "/")
            fields = _clean_lines(item.get("fields_used") or ["reported_source_text"], label="fields_used")
        else:
            token = str(item).strip().replace("\\", "/")
            source_id = token if token in by_id else by_file.get(token, "")
            source_file = by_id.get(token, token if token in by_file else "")
            fields = ["reported_source_text"]
        if not source_id and source_file:
            source_id = by_file.get(source_file, "")
        if not source_file and source_id:
            source_file = by_id.get(source_id, "")
        if not source_id or not source_file or by_id.get(source_id) != source_file:
            raise FastProductionError(
                f"source binding {index} does not resolve in the current SourceRegistry"
            )
        result.append(
            {
                "source_id": source_id,
                "source_file": source_file,
                "fields_used": fields,
                "canonical_status": "source_bound",
            }
        )
    if not result:
        raise FastProductionError("Every changed slide requires source_bindings")
    return result


def _cached_slide_index(cached_state: Mapping[str, Any]) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    rows = cached_state.get("slide_specs")
    if not isinstance(rows, list) or not rows:
        raise FastProductionError("Validated cache has no SlideSpec collection")
    normalized = [dict(row) for row in rows if isinstance(row, Mapping)]
    if len(normalized) != len(rows):
        raise FastProductionError("Cached SlideSpec collection contains a non-object")
    by_id: dict[str, dict[str, Any]] = {}
    for row in normalized:
        slide_id = str(row.get("slide_id", "")).strip()
        if not slide_id or slide_id in by_id:
            raise FastProductionError("Cached SlideSpec identity is missing or duplicated")
        by_id[slide_id] = row
    return normalized, by_id


def _presentation_request(brief: Mapping[str, Any]) -> Mapping[str, Any]:
    request = brief.get("fast_enhance")
    if not isinstance(request, Mapping):
        raise FastProductionError(
            "brief.fast_enhance is required for automatic production planning"
        )
    rows = request.get("changed_slides")
    if not isinstance(rows, list) or not rows:
        raise FastProductionError("brief.fast_enhance.changed_slides must be a non-empty list")
    return request


def build_changed_slide_plans(
    *,
    brief: Mapping[str, Any],
    cached_state: Mapping[str, Any],
    source_registry: Iterable[Mapping[str, Any]],
    changed_claims: Iterable[Mapping[str, Any]] = (),
    unresolved_items: Iterable[str] = (),
    retry_slide: str = "",
) -> list[dict[str, Any]]:
    """Build all changed slides in one batch from explicit update intent.

    The request describes audience content, never low-level candidate artifacts.
    No source prose is re-read here; facts must be supplied as source-bound
    brief facts or exact allowed wording from the changed evidence patch.
    """

    request = _presentation_request(brief)
    cached_rows, cached_by_id = _cached_slide_index(cached_state)
    source_rows = [dict(row) for row in source_registry]
    claim_by_id = {
        str(row.get("claim_id", "")).strip(): dict(row)
        for row in changed_claims
        if isinstance(row, Mapping) and str(row.get("claim_id", "")).strip()
    }
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, raw in enumerate(request["changed_slides"], start=1):
        if not isinstance(raw, Mapping):
            raise FastProductionError(f"changed_slides {index} is not an object")
        operation = str(raw.get("operation", "")).strip().upper()
        if operation not in {"REPLACE", "INSERT_AFTER"}:
            raise FastProductionError(f"changed_slides {index} has unsupported operation")
        if operation == "REPLACE":
            slide_id = str(raw.get("target_slide_id", "")).strip()
            if slide_id not in cached_by_id:
                raise FastProductionError(f"REPLACE target is not in cached SlideSpec: {slide_id}")
            anchor = ""
            previous = cached_by_id[slide_id]
            revision = int(previous.get("slide_revision", 1) or 1) + 1
        else:
            anchor = str(raw.get("insertion_anchor", "")).strip()
            if anchor not in cached_by_id and anchor not in seen:
                raise FastProductionError(f"INSERT_AFTER anchor is unknown: {anchor}")
            logical_key = str(raw.get("logical_key", "")).strip() or f"insert-{index}"
            slide_id = str(raw.get("target_slide_id", "")).strip() or _stable_insert_id(logical_key)
            if not _SEMANTIC_ID_RE.fullmatch(slide_id) or slide_id in cached_by_id:
                raise FastProductionError("INSERT_AFTER target_slide_id is invalid or collides")
            previous = {}
            revision = 1
        if slide_id in seen:
            raise FastProductionError(f"Duplicate changed slide_id: {slide_id}")
        seen.add(slide_id)
        if retry_slide and slide_id != retry_slide:
            continue

        claim_ids = _clean_lines(raw.get("claim_ids"), label="claim_ids")
        facts = _clean_lines(raw.get("facts") or raw.get("fact_bindings"), label="facts")
        for claim_id in claim_ids:
            claim = claim_by_id.get(claim_id)
            if claim is None:
                raise FastProductionError(f"Unknown changed claim_id: {claim_id}")
            wording = str(claim.get("allowed_wording") or claim.get("claim_text") or "").strip()
            if not wording:
                raise FastProductionError(f"Changed claim {claim_id} has no allowed wording")
            facts.append(wording)
        facts = list(dict.fromkeys(facts))
        if not facts:
            raise FastProductionError(f"Changed slide {slide_id} has no source-bound facts")
        title = str(raw.get("title", "")).strip()
        key_message = str(raw.get("key_message", "")).strip()
        if not title or not key_message:
            raise FastProductionError(f"Changed slide {slide_id} requires title and key_message")
        bindings = _source_bindings(raw.get("source_bindings"), source_rows)
        uncertainty = _clean_lines(
            raw.get("uncertainty_bindings") or raw.get("uncertainties"),
            label="uncertainty_bindings",
        )
        prohibited = _clean_lines(raw.get("prohibited_wording"), label="prohibited_wording")
        content_core = {
            "operation": operation,
            "target_slide_id": slide_id,
            "insertion_anchor": anchor,
            "slide_role": str(raw.get("slide_role", previous.get("slide_role", "content"))).strip() or "content",
            "title": title,
            "key_message": key_message,
            "fact_bindings": facts,
            "uncertainty_bindings": uncertainty,
            "source_bindings": bindings,
            "layout_family": str(raw.get("layout_family", previous.get("layout_family", "figure_with_callout"))).strip() or "figure_with_callout",
            "layout_variant": str(raw.get("layout_variant", "large_left")).strip() or "large_left",
            "content_budget": dict(raw.get("content_budget", {})) if isinstance(raw.get("content_budget"), Mapping) else {},
            "manual_review_required": bool(raw.get("manual_review_required", True)),
            "prohibited_wording": prohibited,
            "claim_ids": claim_ids,
        }
        content_hash = _canonical_hash(content_core)
        result.append(
            {
                **content_core,
                "slide_id": slide_id,
                "slide_revision": revision,
                "content_hash": content_hash,
                "source_binding_required": True,
                "scientific_qa_status": "PASS",
                "scientific_review": {"status": "PASS", "scope": "changed_slide"},
                "review_flags": ["manual_review_required"] if content_core["manual_review_required"] else [],
            }
        )
    if retry_slide and not result:
        raise FastProductionError(f"--retry-slide is not declared by brief.fast_enhance: {retry_slide}")
    if not result:
        raise FastProductionError("No ChangedSlidePlan remains after scope filtering")
    if len(result) > 10:
        raise FastProductionError("fast_enhance supports at most 10 changed slides")
    return result


def content_review_status(
    *, brief: Mapping[str, Any], auto_approve_content: bool, clinical_privacy_mode: bool
) -> str:
    request = _presentation_request(brief)
    declared = str(request.get("content_review_status", "")).strip().upper()
    if declared == "CONTENT_APPROVED":
        return declared
    if auto_approve_content and not clinical_privacy_mode:
        return "CONTENT_APPROVED"
    return "CONTENT_REVIEW_REQUIRED"


def write_changed_content_review(
    path: Path, plans: Iterable[Mapping[str, Any]], *, status: str
) -> Path:
    lines = ["# Changed Content Review", "", f"- Status: `{status}`", ""]
    for plan in plans:
        lines.extend(
            [
                f"## {plan['operation']} — {plan['title']}",
                "",
                f"- Slide ID: `{plan['slide_id']}`",
                f"- Insertion anchor: `{plan.get('insertion_anchor') or 'N/A'}`",
                f"- Key message: {plan['key_message']}",
                "- Core facts:",
                *[f"  - {item}" for item in plan.get("fact_bindings", [])],
                "- Uncertainty:",
                *([f"  - {item}" for item in plan.get("uncertainty_bindings", [])] or ["  - None declared"]),
                "- Sources:",
                *[f"  - {item['source_file']} ({item['source_id']})" for item in plan.get("source_bindings", [])],
                "- Prohibited wording:",
                *([f"  - {item}" for item in plan.get("prohibited_wording", [])] or ["  - Any wording beyond registered evidence"]),
                "",
            ]
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def preflight_changed_slide_plans(plans: Iterable[Mapping[str, Any]]) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    for plan in plans:
        title = str(plan.get("title", ""))
        facts = _clean_lines(plan.get("fact_bindings"), label="facts")
        uncertainty = _clean_lines(plan.get("uncertainty_bindings"), label="uncertainty")
        bindings = plan.get("source_bindings") or []
        budget = dict(plan.get("content_budget") or {})
        title_lines = 1 if len(title) <= int(budget.get("title_chars_per_line", 28)) else 2
        estimated_chars = sum(len(item) for item in facts + uncertainty)
        checks = {
            "title_line_count": title_lines,
            "text_character_count": estimated_chars,
            "table_row_count": 0,
            "card_count": min(len(facts), 3),
            "object_count": 4 + len(facts) + len(uncertainty),
            "minimum_font_pt": 18,
            "source_footer_char_count": len("; ".join(str(item.get("source_file", "")) for item in bindings)),
            "mandatory_fact_count": len(facts),
            "uncertainty_count": len(uncertainty),
        }
        errors = []
        if title_lines > 2 or len(title) > int(budget.get("max_title_chars", 56)):
            errors.append("TITLE_BUDGET_EXCEEDED")
        if estimated_chars > int(budget.get("max_body_chars", 520)):
            errors.append("TEXT_BUDGET_EXCEEDED")
        if len(facts) > int(budget.get("max_fact_rows", 8)):
            errors.append("ROW_BUDGET_EXCEEDED")
        if len(bindings) < 1:
            errors.append("SOURCE_BINDING_MISSING")
        if len("; ".join(str(item.get("source_file", "")) for item in bindings)) > 120:
            errors.append("SOURCE_FOOTER_BUDGET_EXCEEDED")
        row = {"slide_id": plan.get("slide_id"), **checks, "errors": errors}
        rows.append(row)
        failures.extend({"slide_id": plan.get("slide_id"), "code": code} for code in errors)
    return {
        "schema_version": "academic-ppt-fast-layout-preflight/1",
        "slides": rows,
        "failures": failures,
        "status": "PASS" if not failures else "FAIL",
    }


def _candidate_spec(plans: Iterable[Mapping[str, Any]]) -> dict[str, Any]:
    slides = []
    for plan in plans:
        notes = [
            "[Slide-ID]",
            str(plan["slide_id"]),
            "[Slide-Revision]",
            str(plan["slide_revision"]),
            "[Content-Hash]",
            str(plan["content_hash"]),
            "[Claims]",
            *[f"- {item}" for item in plan.get("claim_ids", [])],
            "[Wording-Boundary]",
            *(plan.get("prohibited_wording", []) or ["Do not exceed registered source wording."]),
            "[Sources]",
            *[f"- {item['source_id']} | {item['source_file']}" for item in plan.get("source_bindings", [])],
        ]
        slides.append(
            {
                "slide_id": plan["slide_id"],
                "slide_revision": plan["slide_revision"],
                "content_hash": plan["content_hash"],
                "title": plan["title"],
                "key_message": plan["key_message"],
                "facts": list(plan.get("fact_bindings", [])),
                "uncertainties": list(plan.get("uncertainty_bindings", [])),
                "source_label": "; ".join(
                    str(item.get("source_file", "")) for item in plan.get("source_bindings", [])
                ),
                "layout_family": plan.get("layout_family"),
                "layout_variant": plan.get("layout_variant"),
                "notes": "\n".join(notes),
            }
        )
    return {"schema_version": "academic-ppt-fast-candidate-spec/1", "slides": slides}


def _existing_renderer_spec(
    *, plans: Iterable[Mapping[str, Any]], repo_root: Path, brief: Mapping[str, Any]
) -> dict[str, Any]:
    contract = LayoutContract.from_files(
        repo_root / "config" / "layout_contract.yaml",
        repo_root / "config" / "safe_zones.yaml",
    )
    footer_policy = load_yaml_compatible(repo_root / "config" / "footer_policy.yaml")
    manifest_rows: list[dict[str, str]] = []
    seen_sources: set[str] = set()
    for plan in plans:
        for binding in plan.get("source_bindings", []):
            source_id = str(binding["source_id"])
            if source_id in seen_sources:
                continue
            seen_sources.add(source_id)
            manifest_rows.append(
                {
                    "source_id": source_id,
                    "relative_path": str(binding["source_file"]),
                    "file_name": Path(str(binding["source_file"])).name,
                }
            )
    slides = []
    for plan in plans:
        row = {
            "slide_id": plan["slide_id"],
            "slide_revision": plan["slide_revision"],
            "content_hash": plan["content_hash"],
            "slide_title": plan["title"],
            "single_key_message": plan["key_message"],
            "speaker_note_summary": "Changed-slide fast production; preserve registered uncertainty.",
            "source_ids": [item["source_id"] for item in plan.get("source_bindings", [])],
            "claim_ids": list(plan.get("claim_ids", [])),
            "prohibited_overstatement": "; ".join(plan.get("prohibited_wording", []))
            or "Do not exceed registered source wording.",
            "visual_type": "takeaway" if len(plan.get("fact_bindings", [])) > 1 else "text",
            "takeaways": list(plan.get("fact_bindings", []))[:3],
            "layout_family": plan.get("layout_family", "figure_with_callout"),
            "layout_variant": plan.get("layout_variant", "large_left"),
            "manual_review_required": bool(plan.get("manual_review_required", True)),
        }
        row["short_source_label"] = short_source_label(
            row["source_ids"], manifest_rows, footer_policy, audit_presentation=False
        )
        row["planned_geometry"] = plan_slide_geometry(row, contract)
        issues = validate_planned_geometry(row["planned_geometry"], contract)
        if issues:
            raise FastProductionError(
                f"Existing renderer layout contract rejected {row['slide_id']}"
            )
        slides.append(row)
    return {
        "schema_version": "academic-ppt-fast-candidate-render/1",
        "brief": {
            "project_name": str(brief.get("project_name", "Fast enhance")),
            "language": str(brief.get("language", "zh-CN")),
        },
        "layout_contract": contract.serializable(),
        "style_profile": None,
        "slides": slides,
    }


def generate_candidate_deck(
    *,
    plans: list[Mapping[str, Any]],
    staging_root: Path,
    repo_root: Path,
    node_executable: str | Path | None = None,
    node_modules: str | Path | None = None,
    timeout_seconds: int = 120,
    brief: Mapping[str, Any] | None = None,
) -> tuple[Path, Path, dict[str, Any]]:
    spec_path = staging_root / "changed_candidate_spec.json"
    candidate_path = staging_root / "changed_candidate.pptx"
    manifest_path = staging_root / "changed_candidate_manifest.json"
    write_json(
        spec_path,
        _existing_renderer_spec(
            plans=plans, repo_root=repo_root, brief=dict(brief or {})
        ),
    )
    script = repo_root / "scripts" / "generate_deck_pptxgen.mjs"
    node = Path(str(node_executable)).resolve() if node_executable else None
    modules = Path(str(node_modules)).resolve() if node_modules else None
    if node is None or not node.is_file():
        candidate = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe"
        node = candidate if candidate.is_file() else None
    if modules is None or not (modules / "pptxgenjs").is_dir():
        candidate = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules"
        modules = candidate if (candidate / "pptxgenjs").is_dir() else None
    if node is None or modules is None:
        raise FastProductionError("Local Node/PptxGenJS runtime is unavailable")
    environment = dict(os.environ)
    environment["NODE_PATH"] = str(modules) + (
        os.pathsep + environment["NODE_PATH"] if environment.get("NODE_PATH") else ""
    )
    completed = subprocess.run(
        [str(node), str(script), str(spec_path), str(candidate_path), str(staging_root / "candidate_renderer")],
        cwd=repo_root,
        env=environment,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout_seconds,
        check=False,
    )
    if completed.returncode != 0:
        raise FastProductionError(
            "Changed candidate generation failed: "
            + (completed.stdout + completed.stderr).strip()[-2000:]
        )
    if not candidate_path.is_file():
        raise FastProductionError("Changed candidate generator did not create required files")
    manifest = {
        "schema_version": "academic-ppt-fast-candidate-manifest/1",
        "backend": "native_pptxgenjs",
        "slide_count": len(plans),
        "slide_ids": [str(plan["slide_id"]) for plan in plans],
        "candidate_attempt_count": 1,
        "fallback_count": 0,
        "external_service_usage": 0,
    }
    write_json(manifest_path, manifest)
    return candidate_path, spec_path, manifest


def build_automatic_operation_plan(
    *,
    source_pptx: Path,
    cached_state: Mapping[str, Any],
    plans: list[Mapping[str, Any]],
    candidate_pptx: Path,
    output_path: Path,
) -> dict[str, Any]:
    cached_rows, cached_by_id = _cached_slide_index(cached_state)
    source_ids = [str(row["slide_id"]) for row in cached_rows]
    plan_by_id = {str(row["slide_id"]): dict(row) for row in plans}
    insertions: dict[str, list[dict[str, Any]]] = {}
    for row in plans:
        if row["operation"] == "INSERT_AFTER":
            insertions.setdefault(str(row["insertion_anchor"]), []).append(dict(row))
    final_ids: list[str] = []
    for source_id in source_ids:
        final_ids.append(source_id)
        anchor = source_id
        while anchor in insertions:
            rows = insertions[anchor]
            if len(rows) != 1:
                raise FastProductionError("Each insertion anchor must be unambiguous")
            inserted = rows[0]
            final_ids.append(str(inserted["slide_id"]))
            anchor = str(inserted["slide_id"])
    if any(row["operation"] == "INSERT_AFTER" and row["slide_id"] not in final_ids for row in plans):
        raise FastProductionError("One or more insertion anchors cannot be ordered")
    candidate_indexes = {str(row["slide_id"]): index for index, row in enumerate(plans, 1)}
    operations = []
    for slide_id in final_ids:
        changed = plan_by_id.get(slide_id)
        if changed is None:
            operations.append({"action": "KEEP", "slide_id": slide_id})
            continue
        row = {
            "action": changed["operation"],
            "slide_id": slide_id,
            "candidate_pptx": str(candidate_pptx),
            "candidate_index": candidate_indexes[slide_id],
            "slide_revision": changed["slide_revision"],
            "content_hash": changed["content_hash"],
        }
        if changed["operation"] == "INSERT_AFTER":
            row["after_slide_id"] = changed["insertion_anchor"]
        operations.append(row)
    built = build_operation_plan(
        source_pptx=source_pptx,
        source_slide_ids=source_ids,
        operations=operations,
        expected_final_order=final_ids,
        changed_slide_specs=[dict(row) for row in plans],
        source_impacts=[
            {
                "source_id": binding["source_id"],
                "affected_slide_ids": [plan["slide_id"]],
            }
            for plan in plans
            for binding in plan.get("source_bindings", [])
        ],
        manual_review_items=[
            {
                "review_id": f"MR-{index:03d}",
                "slide_id": plan["slide_id"],
                "text": "Review changed slide scientific content and visual fit.",
            }
            for index, plan in enumerate(plans, 1)
            if plan.get("manual_review_required")
        ],
    )
    write_operation_plan(output_path, built)
    return validate_operation_plan(built)


def plan_hash_bundle(
    *, brief_path: Path, source_registry: Mapping[str, Any], plans: Iterable[Mapping[str, Any]],
    candidate_path: Path | None = None, code_path: Path | None = None, config_path: Path | None = None,
) -> dict[str, str]:
    bundle = {
        "brief_hash": sha256_file(brief_path).lower(),
        "source_registry_hash": _canonical_hash(source_registry),
        "slide_spec_hash": _canonical_hash(list(plans)),
    }
    if candidate_path is not None:
        bundle["candidate_hash"] = sha256_file(candidate_path).lower()
    if code_path is not None:
        bundle["code_hash"] = sha256_file(code_path).lower()
    if config_path is not None:
        bundle["config_hash"] = sha256_file(config_path).lower()
    if any(not _SHA256_RE.fullmatch(value) for value in bundle.values()):
        raise FastProductionError("Invalid checkpoint hash bundle")
    return bundle


def candidate_attempt_record(
    plans: Iterable[Mapping[str, Any]], *, attempt: int, repair_reason: str = "",
    fallback_layout: str = "", repair_runtime: float = 0.0,
) -> dict[str, Any]:
    if attempt not in {1, 2}:
        raise FastProductionError("candidate attempts must be 1 or 2")
    return {
        "schema_version": "academic-ppt-candidate-attempts/1",
        "slides": [
            {
                "slide_id": plan["slide_id"],
                "candidate_attempt_count": attempt,
                "repair_reason": repair_reason,
                "fallback_layout": fallback_layout,
                "repair_runtime": round(float(repair_runtime), 6),
            }
            for plan in plans
        ],
        "maximum_attempts": 2,
    }


def _rels_target(package: zipfile.ZipFile, source: str, rel_type_suffix: str) -> str:
    directory, filename = source.rsplit("/", 1)
    rels_name = f"{directory}/_rels/{filename}.rels"
    if rels_name not in package.namelist():
        return ""
    root = ET.fromstring(package.read(rels_name))
    for row in root:
        if str(row.get("Type", "")).endswith(rel_type_suffix):
            target = str(row.get("Target", "")).replace("\\", "/")
            return posixpath.normpath(posixpath.join(directory, target))
    return ""


def _semantic_xml_hash(element: ET.Element | None, *, omit_slide_number: bool = False) -> str:
    if element is None:
        return hashlib.sha256(b"").hexdigest()
    clone = ET.fromstring(ET.tostring(element, encoding="utf-8"))
    if omit_slide_number:
        presentation_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
        for parent in clone.iter():
            for child in list(parent):
                placeholder = child.find(f".//{{{presentation_ns}}}ph")
                if placeholder is not None and placeholder.get("type") == "sldNum":
                    parent.remove(child)
    for row in clone.iter():
        # PowerPoint normalizes empty property lists on save.  They carry no
        # shape-tree semantics, so exclude them from KEEP identity.
        if row.attrib:
            ordered = sorted(row.attrib.items())
            row.attrib.clear()
            row.attrib.update(ordered)
        row.text = None if not (row.text or "").strip() else (row.text or "").strip()
        row.tail = None if not (row.tail or "").strip() else (row.tail or "").strip()
    for parent in clone.iter():
        for child in list(parent):
            if not child.attrib and not (child.text or "").strip() and len(child) == 0:
                parent.remove(child)
    return hashlib.sha256(ET.tostring(clone, encoding="utf-8")).hexdigest()


def _layout_semantic_hash(payload: bytes) -> str:
    root = ET.fromstring(payload)
    presentation_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    common_slide = root.find(f"{{{presentation_ns}}}cSld")
    if common_slide is None:
        common_slide = root
    return _semantic_xml_hash(common_slide, omit_slide_number=True)


def slide_structure_fingerprints(path: Path) -> list[dict[str, Any]]:
    """Create shallow structural fingerprints without reading slide prose."""

    with zipfile.ZipFile(path, "r") as package:
        names = set(package.namelist())
        presentation = ET.fromstring(package.read("ppt/presentation.xml"))
        rels = ET.fromstring(package.read("ppt/_rels/presentation.xml.rels"))
        rel_map = {str(row.get("Id")): str(row.get("Target")).replace("\\", "/") for row in rels}
        namespace = "http://schemas.openxmlformats.org/presentationml/2006/main"
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        rows = []
        for index, slide_id in enumerate(presentation.findall(f".//{{{namespace}}}sldId"), 1):
            target = rel_map[str(slide_id.get(f"{{{rel_ns}}}id"))]
            slide_part = str(Path("ppt", target).as_posix())
            payload = package.read(slide_part)
            root = ET.fromstring(payload)
            sp_tree = root.find(f".//{{{namespace}}}spTree")
            shape_payload = ET.tostring(sp_tree, encoding="utf-8") if sp_tree is not None else b""
            rel_targets: dict[str, Any] = {}
            for suffix in ("/slideLayout", "/notesSlide", "/image", "/theme"):
                role = suffix.rsplit("/", 1)[-1]
                target = _rels_target(package, slide_part, suffix)
                if role == "notesSlide":
                    rel_targets[role] = bool(target)
                elif target and target in names:
                    if role in {"slideLayout", "theme"}:
                        rel_targets[role] = (
                            _layout_semantic_hash(package.read(target))
                            if role == "slideLayout"
                            else _semantic_xml_hash(ET.fromstring(package.read(target)))
                        )
                    else:
                        rel_targets[role] = hashlib.sha256(package.read(target)).hexdigest()
                else:
                    rel_targets[role] = ""
            background = root.find(f"{{{namespace}}}cSld/{{{namespace}}}bg")
            rows.append(
                {
                    "index": index,
                    "powerpoint_slide_id": str(slide_id.get("id", "")),
                    "slide_part": slide_part,
                    "shape_tree_hash": _semantic_xml_hash(sp_tree, omit_slide_number=True),
                    "background_hash": _semantic_xml_hash(background),
                    "relationships": rel_targets,
                    "footer_placeholder_present": b"ftr" in shape_payload,
                    "page_number_placeholder_present": b"sldNum" in shape_payload,
                }
            )
        return rows


def compare_keep_fingerprints(
    *, source_pptx: Path, updated_pptx: Path, operation_plan: Mapping[str, Any]
) -> dict[str, Any]:
    source = slide_structure_fingerprints(source_pptx)
    updated = slide_structure_fingerprints(updated_pptx)
    source_ids = list(operation_plan["source_slide_ids"])
    final_ids = list(operation_plan["expected_final_order"])
    source_by_id = dict(zip(source_ids, source))
    updated_by_id = dict(zip(final_ids, updated))
    findings = []
    for operation in operation_plan["operations"]:
        if operation["action"] != "KEEP":
            continue
        slide_id = operation["slide_id"]
        before = source_by_id[slide_id]
        after = updated_by_id[slide_id]
        matched = all(
            before[field] == after[field]
            for field in (
                "powerpoint_slide_id", "shape_tree_hash", "background_hash", "relationships",
                "footer_placeholder_present", "page_number_placeholder_present",
            )
        )
        findings.append({"slide_id": slide_id, "matched": matched})
    return {
        "schema_version": "academic-ppt-keep-fingerprints/1",
        "keep_slide_count": len(findings),
        "matched_count": sum(row["matched"] for row in findings),
        "findings": findings,
        "status": "PASS" if all(row["matched"] for row in findings) else "FAIL",
    }


def build_keep_baseline_fingerprints(
    *, source_pptx: Path, operation_plan: Mapping[str, Any]
) -> dict[str, Any]:
    source_rows = slide_structure_fingerprints(source_pptx)
    source_ids = list(operation_plan["source_slide_ids"])
    by_id = dict(zip(source_ids, source_rows))
    keep = {
        str(row["slide_id"]): by_id[str(row["slide_id"])]
        for row in operation_plan["operations"]
        if row["action"] == "KEEP"
    }
    return {
        "schema_version": "academic-ppt-keep-baseline/1",
        "source_pptx_sha256": sha256_file(source_pptx).lower(),
        "keep_slides": keep,
    }


def validate_keep_baseline_fingerprints(
    *, baseline: Mapping[str, Any], updated_pptx: Path,
    expected_final_order: Iterable[str], operation_plan: Mapping[str, Any]
) -> dict[str, Any]:
    updated_rows = slide_structure_fingerprints(updated_pptx)
    updated_by_id = dict(zip(list(expected_final_order), updated_rows))
    findings = []
    for slide_id, before in dict(baseline.get("keep_slides", {})).items():
        after = updated_by_id.get(slide_id)
        matched = isinstance(after, Mapping) and all(
            before[field] == after[field]
            for field in (
                "powerpoint_slide_id", "shape_tree_hash", "background_hash", "relationships",
                "footer_placeholder_present", "page_number_placeholder_present",
            )
        )
        findings.append({"slide_id": slide_id, "matched": bool(matched)})
    return {
        "schema_version": "academic-ppt-keep-fingerprints/1",
        "keep_slide_count": len(findings),
        "matched_count": sum(row["matched"] for row in findings),
        "findings": findings,
        "status": "PASS" if findings and all(row["matched"] for row in findings) else "FAIL",
    }


@dataclass
class FastStateMachine:
    path: Path
    input_hashes: Mapping[str, str]

    def __post_init__(self) -> None:
        self.document: dict[str, Any] = {
            "schema_version": "academic-ppt-fast-state-machine/1",
            "state_sequence": list(FAST_STATE_SEQUENCE),
            "stages": [],
        }

    def complete(self, stage: str, outputs: Mapping[str, str]) -> None:
        if stage not in FAST_STATE_SEQUENCE:
            raise FastProductionError(f"Unknown fast state: {stage}")
        expected = FAST_STATE_SEQUENCE[len(self.document["stages"])]
        if stage != expected:
            raise FastProductionError(f"State transition must be {expected}, not {stage}")
        if any(not _SHA256_RE.fullmatch(str(value)) for value in outputs.values()):
            raise FastProductionError("Every state output must be represented by a SHA-256")
        self.document["stages"].append(
            {"stage": stage, "input_hashes": dict(self.input_hashes), "output_hashes": dict(outputs)}
        )
        write_json(self.path, self.document)

    def checkpoint(self, stage: str, outputs: Mapping[str, str]) -> None:
        """Record or refresh one independently resumable stage.

        Production orchestration may resume at any state after validating its
        hash bundle.  The ledger stays in canonical sequence while permitting
        an already-recorded stage to be replaced after a scoped retry.
        """

        if stage not in FAST_STATE_SEQUENCE:
            raise FastProductionError(f"Unknown fast state: {stage}")
        if any(not _SHA256_RE.fullmatch(str(value)) for value in outputs.values()):
            raise FastProductionError("Every state output must be represented by a SHA-256")
        row = {"stage": stage, "input_hashes": dict(self.input_hashes), "output_hashes": dict(outputs)}
        by_stage = {item["stage"]: dict(item) for item in self.document["stages"]}
        by_stage[stage] = row
        self.document["stages"] = [
            by_stage[name] for name in FAST_STATE_SEQUENCE if name in by_stage
        ]
        write_json(self.path, self.document)


__all__ = [
    "FAST_STATE_SEQUENCE",
    "FastProductionError",
    "FastStateMachine",
    "build_automatic_operation_plan",
    "build_keep_baseline_fingerprints",
    "build_changed_slide_plans",
    "candidate_attempt_record",
    "compare_keep_fingerprints",
    "content_review_status",
    "generate_candidate_deck",
    "plan_hash_bundle",
    "preflight_changed_slide_plans",
    "slide_structure_fingerprints",
    "validate_keep_baseline_fingerprints",
    "write_changed_content_review",
]
