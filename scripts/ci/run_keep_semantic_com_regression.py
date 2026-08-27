from __future__ import annotations

import argparse
import json
import subprocess
import time
from pathlib import Path

from academic_ppt.fast_production import (
    build_keep_baseline_fingerprints,
    build_keep_render_identity_evidence,
    slide_structure_fingerprints,
    validate_keep_baseline_fingerprints,
)
from academic_ppt.utils import write_json


def _run(command: list[str], *, cwd: Path, timeout: int = 240) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        command, cwd=cwd, capture_output=True, text=True, timeout=timeout, check=False
    )
    if result.returncode != 0:
        raise RuntimeError("KEEP COM regression subprocess failed")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the real PowerPoint KEEP semantic regression")
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo_root.resolve()
    root = args.output.resolve()
    if root.exists():
        raise RuntimeError("COM regression output directory must be new")
    root.mkdir(parents=True)
    generator = repo / "tests" / "fixtures" / "synthetic" / "keep_semantics" / "generate_fixture.cjs"
    source = root / "synthetic_existing_before_com.pptx"
    generated_changed = root / "synthetic_changed_before_com.pptx"
    updated = root / "synthetic_changed_after_com.pptx"
    _run(["node", str(generator), str(source)], cwd=repo)
    _run(["node", str(generator), str(generated_changed), "changed-chart"], cwd=repo)
    saveas = _run(
        [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
            str(repo / "scripts" / "powerpoint_saveas_copy.ps1"),
            "-InputPptx", str(generated_changed), "-OutputPptx", str(updated),
        ],
        cwd=repo,
    )
    ids = [f"SYN-KEEP-{index}" for index in range(1, 7)]
    plan = {
        "source_slide_ids": ids,
        "expected_final_order": ids,
        "operations": [
            {"action": "KEEP" if index < 6 else "REPLACE", "slide_id": slide_id}
            for index, slide_id in enumerate(ids, 1)
        ],
    }
    raw_started = time.perf_counter()
    source_raw = slide_structure_fingerprints(source)
    updated_raw = slide_structure_fingerprints(updated)
    raw_runtime = round(time.perf_counter() - raw_started, 6)
    baseline_started = time.perf_counter()
    baseline = build_keep_baseline_fingerprints(source_pptx=source, operation_plan=plan)
    first = validate_keep_baseline_fingerprints(
        baseline=baseline,
        updated_pptx=updated,
        expected_final_order=ids,
        operation_plan=plan,
    )
    semantic_runtime = round(time.perf_counter() - baseline_started, 6)
    render = build_keep_render_identity_evidence(
        source_pptx=source,
        updated_pptx=updated,
        baseline=baseline,
        expected_final_order=ids,
        slide_ids=first["render_required_slide_ids"],
        output_dir=root / "render_identity",
        powershell_script=repo / "scripts" / "powerpoint_compare_keep_render.ps1",
    )
    final = validate_keep_baseline_fingerprints(
        baseline=baseline,
        updated_pptx=updated,
        expected_final_order=ids,
        operation_plan=plan,
        render_identity=render["slides"],
    )
    changed_distinguished = (
        source_raw[5]["semantic_fingerprint_sha256"]
        != updated_raw[5]["semantic_fingerprint_sha256"]
    )
    raw_differences = sum(
        left["raw_shape_tree_hash"] != right["raw_shape_tree_hash"]
        for left, right in zip(source_raw[:5], updated_raw[:5])
    )
    report = {
        "schema_version": "academic-ppt-com-keep-semantic-regression/1",
        "fixture_id": "SECOND_DEVICE_COM_KEEP_TABLE_NORMALIZATION",
        "fixture_provenance": "SYNTHETIC",
        "powerpoint": saveas.stdout.strip(),
        "before_keep_result": {
            "status": first["status"],
            "matched_count": first["matched_count"],
            "keep_slide_count": first["keep_slide_count"],
        },
        "after_keep_result": {
            "status": final["status"],
            "matched_count": final["matched_count"],
            "keep_slide_count": final["keep_slide_count"],
        },
        "text_table_image_diagram_mixed_keep": final["status"] == "PASS" and final["matched_count"] == 5,
        "changed_slide_distinguished": changed_distinguished,
        "raw_ooxml_changed_keep_slides": raw_differences,
        "semantic_fingerprint_differences_keep_slides": sum(
            row["semantic_editable_identity"] != "PASS" for row in final["findings"]
        ),
        "render_differences_keep_slides": sum(
            row.get("status") != "PASS" for row in render["slides"].values()
        ),
        "raw_hash_runtime_seconds": raw_runtime,
        "semantic_manifest_runtime_seconds": semantic_runtime,
        "render_validation_runtime_seconds": render["runtime_seconds"],
        "status": "PASS" if final["status"] == "PASS" and changed_distinguished else "FAIL",
    }
    write_json(root / "keep_semantic_com_regression.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
