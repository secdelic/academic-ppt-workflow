from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import time
from pathlib import Path
from types import SimpleNamespace


REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "scripts"))
sys.path.insert(0, str(REPO_ROOT))

from academic_ppt.fast_enhance import build_cached_change_impact_graph  # noqa: E402
from academic_ppt.incremental import build_operation_plan, write_operation_plan  # noqa: E402
from academic_ppt.inventory import inventory_sources  # noqa: E402
from academic_ppt.project_cache import (  # noqa: E402
    ProjectCache,
    build_cache_contract_fingerprints,
    build_project_state,
    build_source_manifest,
)
from academic_ppt.runner import execute  # noqa: E402
from academic_ppt.utils import sha256_file  # noqa: E402
from academic_ppt.utils import stable_source_id  # noqa: E402
from tests.fast_enhance.fixture_factory import (  # noqa: E402
    build_fixture_definition,
    generate_synthetic_pptx,
    materialize_synthetic_fixture,
)


SUPPORTED = {".csv", ".txt"}


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _create_inputs(root: Path) -> tuple[Path, Path]:
    prior = root / "prior_input"
    current = root / "current_input"
    prior.mkdir(parents=True)
    current.mkdir(parents=True)
    unchanged = "component,state\nparser,ready\nrenderer,ready\n"
    (prior / "component_matrix.csv").write_text(unchanged, encoding="utf-8")
    (current / "component_matrix.csv").write_text(unchanged, encoding="utf-8")
    (prior / "supplement.txt").write_text(
        "Synthetic supplement revision one. Structure only.\n", encoding="utf-8"
    )
    (current / "supplement.txt").write_text(
        "Synthetic supplement revision two. Changed structure only.\n", encoding="utf-8"
    )
    (current / "addendum.txt").write_text(
        "Synthetic addendum. Changed-slide QA and retry scope.\n", encoding="utf-8"
    )
    return prior, current


def _registry(input_root: Path, manifest_path: Path) -> list[dict[str, str]]:
    return inventory_sources(input_root, SUPPORTED, manifest_path)


def _source_manifest(input_root: Path, registry: list[dict[str, str]]) -> dict:
    return build_source_manifest(
        {
            str(row["relative_path"]): input_root / str(row["relative_path"])
            for row in registry
        },
        relative_to=input_root,
    )


def _binding(source_id: str, source_file: str) -> list[dict[str, object]]:
    return [
        {
            "source_id": source_id,
            "source_file": source_file,
            "fields_used": ["synthetic_structure"],
        }
    ]


def run(root: Path) -> Path:
    wall_started = time.perf_counter()
    if root.exists():
        raise FileExistsError(f"Refusing to overwrite performance run: {root}")
    root.mkdir(parents=True)
    project_identity = (
        "SYNTHETIC_FAST_ENHANCE_"
        + __import__("hashlib").sha256(root.name.encode("utf-8")).hexdigest()[:12].upper()
    )
    brief_document = {
        "project_name": project_identity,
        "presentation_type": "software_verification",
        "language": "en-US",
    }
    fixture = generate_synthetic_pptx(
        materialize_synthetic_fixture(root / "fixture", render_pptx=False)
    )
    if fixture.baseline_pptx is None or fixture.candidate_pptx is None:
        raise RuntimeError("Synthetic PPTX generation failed")
    definition = build_fixture_definition()
    prior_input, current_input = _create_inputs(root)
    prior_registry = _registry(prior_input, root / "prior_manifest.csv")
    current_registry = _registry(current_input, root / "current_manifest.csv")
    prior_by_relative = {row["relative_path"]: row for row in prior_registry}
    current_by_relative = {row["relative_path"]: row for row in current_registry}
    stable_source = str(prior_by_relative["component_matrix.csv"]["source_id"])

    baseline_specs = []
    for row in definition["baseline_slides"]:
        baseline_specs.append(
            {
                **row,
                "source_bindings": _binding(stable_source, "component_matrix.csv"),
                "source_binding_required": True,
                "scientific_qa_status": "PASS",
            }
        )
    graph = build_cached_change_impact_graph(
        source_registry=prior_registry,
        evidence_registry=[],
        slide_specs=baseline_specs,
    )
    prior_source_manifest = _source_manifest(prior_input, prior_registry)
    parsed = {
        row["source_id"]: (prior_input / row["relative_path"]).read_text(encoding="utf-8")
        for row in prior_registry
    }
    cache_state = build_project_state(
        cache_contract_fingerprints=build_cache_contract_fingerprints(
            repo_root=REPO_ROOT,
            brief=brief_document,
        ),
        source_manifest=prior_source_manifest,
        parsed_document_objects=parsed,
        evidence_registry=[],
        slide_specs=baseline_specs,
        visual_specs=[],
        master_layout_map={"registered": True},
        font_map={"latin": "Aptos"},
        image_registry=[],
        source_bindings=baseline_specs,
        previous_qa_status={"status": "PASS"},
        extra={
            "source_registry": prior_registry,
            "change_impact_graph": graph.to_dict(),
            "pptx_state": {
                "sha256": sha256_file(fixture.baseline_pptx),
                "slide_count": 21,
            },
        },
    )
    cache_root = Path(os.environ.get("PPT_CACHE_HOME", root / "cache")) / "project_state"
    cache = ProjectCache.from_identity(
        REPO_ROOT,
        project_identity,
        cache_root=cache_root,
        project_root=root,
        production=True,
        clinical_privacy_mode=True,
    )
    cache.commit_generation(cache_state)

    changed_specs = []
    source_impacts: dict[str, list[str]] = {}
    for index, row in enumerate(definition["candidate_slides"]):
        relative = "supplement.txt" if index < 4 else "addendum.txt"
        source_id = str(current_by_relative[relative]["source_id"])
        changed_specs.append(
            {
                **row,
                "source_bindings": _binding(source_id, relative),
                "source_binding_required": True,
                "scientific_qa_status": "PASS",
            }
        )
        source_impacts.setdefault(source_id, []).append(str(row["slide_id"]))
    operations = []
    for row in definition["operations"]:
        item = dict(row)
        if item["action"] in {"REPLACE", "INSERT_AFTER"}:
            item["candidate_pptx"] = str(fixture.candidate_pptx)
            changed = definition["candidate_slides"][int(item["candidate_index"]) - 1]
            item["slide_revision"] = changed["slide_revision"]
            item["content_hash"] = changed["content_hash"]
        operations.append(item)
    plan = build_operation_plan(
        source_pptx=fixture.baseline_pptx,
        source_slide_ids=[row["slide_id"] for row in definition["baseline_slides"]],
        operations=operations,
        expected_final_order=[row["slide_id"] for row in definition["final_slides"]],
        changed_slide_specs=changed_specs,
        source_impacts=[
            {"source_id": source_id, "affected_slide_ids": slide_ids}
            for source_id, slide_ids in source_impacts.items()
        ],
        manual_review_items=[
            {"review_id": "MR-SYN-001", "text": "Confirm the synthetic delta structure."}
        ],
    )
    plan_path = write_operation_plan(root / "operation_plan.json", plan)
    brief = root / "brief.json"
    _write_json(brief, brief_document)
    staging = root / "staging"
    output = root / "output"
    args = SimpleNamespace(
        brief=str(brief),
        route="enhance-existing",
        repo_root=str(REPO_ROOT),
        input_root=str(current_input),
        staging_root=str(staging),
        output_root=str(output),
        audit_root=None,
        run_id="synthetic_m_7_slides",
        backend="pptxgenjs",
        node=None,
        node_modules=None,
        soffice=None,
        pdftoppm=None,
        from_stage="preflight",
        resume=False,
        workflow_mode="fast_enhance",
        final_delivery=False,
        cross_renderer_validation=False,
        audit_full=False,
        retry_slide=None,
        retry_stage=None,
        operation_plan=str(plan_path),
        changed_candidate_pptx=str(fixture.candidate_pptx),
        project_root=str(root),
        project_cache_root=str(cache_root),
        clinical_privacy_mode=True,
        export_pdf=False,
        manual_review_approved=False,
        reference_pptx=None,
        style_profile=None,
        reference_mode="style-only",
        existing_pptx=str(fixture.baseline_pptx),
        previous_deck_ir=None,
        previous_object_manifest=None,
        update_slide=None,
        update_section=None,
        update_figure=None,
        allow_managed_overwrite=False,
        template_layout_id=None,
        narrative_mode=None,
    )
    execute_started = time.perf_counter()
    output_dir = execute(args)
    execute_finished = time.perf_counter()
    report = {
        "schema_version": "2.5.1",
        "synthetic_only": True,
        "original_slide_count": 21,
        "changed_slide_count": 7,
        "final_slide_count": 25,
        "source_count": len(current_registry),
        "model_calls": 0,
        "planning_telemetry": {
            "model_calls": 0,
            "input_tokens": 0,
            "output_tokens": 0,
            "reason": "fully_deterministic_synthetic_fixture",
        },
        "preparation_seconds": round(execute_started - wall_started, 6),
        "workflow_execute_seconds": round(execute_finished - execute_started, 6),
        "outer_wall_clock_seconds": round(execute_finished - wall_started, 6),
        "runtime_boundary": {
            "candidate_and_plan_prepared_before_workflow_execute": True,
            "scientific_model_planning_included": False,
            "benchmark_role": "incremental_apply_and_qa_engine",
        },
        "input_hashes_unchanged": all(
            sha256_file(current_input / row["relative_path"]) == row["sha256"]
            for row in current_registry
        ),
        "output_dir": str(output_dir),
    }
    _write_json(root / "performance_result.json", report)
    return output_dir


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    output = run(Path(args.out).resolve())
    print(f"OUTPUT_DIR={output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
