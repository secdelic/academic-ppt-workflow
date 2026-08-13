from __future__ import annotations

import argparse
import copy
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

from academic_ppt.fast_enhance import (  # noqa: E402
    build_cached_change_impact_graph,
    execute_fast_production,
)
from academic_ppt.inventory import inventory_sources  # noqa: E402
from academic_ppt.project_cache import (  # noqa: E402
    ProjectCache,
    build_cache_contract_fingerprints,
    build_project_state,
    build_source_manifest,
)
from academic_ppt.utils import sha256_file, stable_source_id  # noqa: E402
from tests.fast_enhance.fixture_factory import (  # noqa: E402
    build_fixture_definition,
    generate_synthetic_pptx,
    materialize_synthetic_fixture,
)


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _binding(source_id: str, source_file: str) -> list[dict[str, object]]:
    return [{"source_id": source_id, "source_file": source_file, "fields_used": ["synthetic_structure"]}]


def _materialize_sources(root: Path) -> tuple[Path, Path]:
    prior, current = root / "prior_input", root / "current_input"
    prior.mkdir(parents=True)
    current.mkdir(parents=True)
    matrix = "component,state\nparser,ready\nrenderer,ready\n"
    for target in (prior, current):
        (target / "component_matrix.csv").write_text(matrix, encoding="utf-8")
    (prior / "supplement.txt").write_text("Synthetic supplement revision one.\n", encoding="utf-8")
    (current / "supplement.txt").write_text(
        "Synthetic supplement revision two. Changed structure only.\n", encoding="utf-8"
    )
    (current / "addendum.txt").write_text(
        "Synthetic addendum for changed-slide QA and retry.\n", encoding="utf-8"
    )
    return prior, current


def _brief(project_name: str, definition: dict, source_ids: dict[str, str], *, clinical: bool = False, retry_slide: str = "") -> dict:
    changes = []
    replacement_ids = {
        row["slide_id"]: row for row in definition["candidate_slides"][:3]
    }
    insertion_rows = definition["candidate_slides"][3:]
    operation_by_id = {row["slide_id"]: row for row in definition["operations"]}
    for index, row in enumerate(definition["candidate_slides"]):
        if retry_slide and row["slide_id"] != retry_slide:
            continue
        operation = operation_by_id[row["slide_id"]]
        relative = "supplement.txt" if index < 4 else "addendum.txt"
        changes.append(
            {
                "operation": operation["action"],
                "target_slide_id": row["slide_id"],
                "insertion_anchor": operation.get("after_slide_id", ""),
                "logical_key": row["logical_key"],
                "title": row["title"],
                "key_message": row["key_message"],
                "facts": [row["key_message"]],
                "uncertainties": ["Synthetic structure only; no real-world conclusion."],
                "source_bindings": [source_ids[relative]],
                "prohibited_wording": ["Do not describe this synthetic fixture as a real result."],
                "manual_review_required": True,
            }
        )
    return {
        "project_name": project_name,
        "presentation_type": "clinical_simulation" if clinical else "software_verification",
        "clinical_privacy_mode": clinical,
        "language": "en-US",
        "fast_enhance": {
            "content_review_status": "CONTENT_APPROVED",
            "changed_slides": changes,
        },
    }


def _retry_brief(brief: dict, retry_slide: str) -> dict:
    result = copy.deepcopy(brief)
    rows = result["fast_enhance"]["changed_slides"]
    selected = [row for row in rows if row.get("target_slide_id") == retry_slide]
    if not selected:
        raise ValueError("retry_slide is not a declared changed slide")
    result["fast_enhance"]["changed_slides"] = selected
    return result


def run(root: Path, *, clinical: bool = False, retry_slide: str = "") -> dict:
    if root.exists():
        raise FileExistsError(root)
    root.mkdir(parents=True)
    fixture = generate_synthetic_pptx(materialize_synthetic_fixture(root / "fixture"))
    definition = build_fixture_definition()
    prior_input, current_input = _materialize_sources(root)
    if retry_slide:
        # A true one-slide retry must not leave an unrelated NEW source in the
        # delta manifest; doing so correctly escalates as missing lineage.
        (current_input / "addendum.txt").unlink()
    prior_manifest = inventory_sources(prior_input, {".csv", ".txt"}, root / "prior_manifest.csv")
    current_manifest = inventory_sources(current_input, {".csv", ".txt"}, root / "current_manifest.csv")
    current_ids = {row["relative_path"]: row["source_id"] for row in current_manifest}
    project_name = "SYNTHETIC_PRODUCTION_CLOSURE_" + root.name[-8:].upper()
    brief = _brief(project_name, definition, current_ids, clinical=clinical, retry_slide=retry_slide)
    if retry_slide:
        brief = _retry_brief(brief, retry_slide)
    brief_path = root / "brief.json"
    _write_json(brief_path, brief)
    prior_ids = {row["relative_path"]: row["source_id"] for row in prior_manifest}
    baseline_specs = []
    for row in definition["baseline_slides"]:
        baseline_specs.append(
            {
                **row,
                "source_bindings": _binding(prior_ids["component_matrix.csv"], "component_matrix.csv"),
                "source_binding_required": True,
                "scientific_qa_status": "PASS",
            }
        )
    parsed = {
        row["source_id"]: (prior_input / row["relative_path"]).read_text(encoding="utf-8")
        for row in prior_manifest
    }
    graph = build_cached_change_impact_graph(
        source_registry=prior_manifest, evidence_registry=[], slide_specs=baseline_specs
    )
    cache_state = build_project_state(
        cache_contract_fingerprints=build_cache_contract_fingerprints(repo_root=REPO_ROOT, brief=brief),
        source_manifest=build_source_manifest(
            {row["relative_path"]: prior_input / row["relative_path"] for row in prior_manifest},
            relative_to=prior_input,
        ),
        parsed_document_objects=parsed,
        evidence_registry=[],
        slide_specs=baseline_specs,
        visual_specs=[],
        master_layout_map={"registered": True},
        font_map={"latin": "Arial"},
        image_registry=[],
        source_bindings=baseline_specs,
        previous_qa_status={"status": "PASS"},
        extra={
            "source_registry": prior_manifest,
            "change_impact_graph": graph.to_dict(),
            "pptx_state": {"sha256": sha256_file(fixture.baseline_pptx), "slide_count": 21},
        },
    )
    cache_root = Path(os.environ.get("PPT_CACHE_HOME", root / "cache")) / "project_state"
    cache = ProjectCache.from_identity(REPO_ROOT, project_name, cache_root=cache_root, clinical_privacy_mode=True)
    cache.commit_generation(cache_state)
    args = SimpleNamespace(
        repo_root=str(REPO_ROOT), brief=str(brief_path), existing_pptx=str(fixture.baseline_pptx),
        input_root=str(current_input), staging_root=str(root / "staging"), output_root=str(root / "output"),
        run_id="production_e2e", project_cache_root=str(cache_root), clinical_privacy_mode=False,
        reference_mode="style-only", node=None, node_modules=None, soffice=None,
        route="enhance-existing", workflow_mode="fast_enhance", auto_approve_content=False,
        final_delivery=False, cross_renderer_validation=False, audit_full=False,
        retry_slide=retry_slide or None, retry_stage=None, resume=False, export_pdf=False,
        manual_review_approved=False, changed_candidate_pptx=None, operation_plan=None,
    )
    started = time.perf_counter()
    output = execute_fast_production(args)
    elapsed = time.perf_counter() - started
    summary = {
        "schema_version": "2.5.2",
        "synthetic_only": True,
        "runtime_boundary": "changed_source_to_draft_ready",
        "elapsed_seconds": round(elapsed, 6),
        "output": str(output),
        "original_slides": 21,
        "changed_slides": 1 if retry_slide else 7,
        "keep_slides": 20 if retry_slide else 18,
        "model_calls": 0,
        "tokens": 0,
        "clinical_simulation": clinical,
        "retry_slide": retry_slide,
        "input_hashes_unchanged": all(
            sha256_file(current_input / row["relative_path"]) == row["sha256"]
            for row in current_manifest
        ),
    }
    _write_json(root / "production_result.json", summary)
    return summary


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True)
    parser.add_argument("--clinical-simulation", action="store_true")
    parser.add_argument("--retry-slide")
    args = parser.parse_args()
    print(json.dumps(run(Path(args.out).resolve(), clinical=args.clinical_simulation, retry_slide=args.retry_slide or ""), indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
