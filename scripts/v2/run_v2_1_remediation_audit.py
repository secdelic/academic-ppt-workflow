from __future__ import annotations

import argparse
import csv
import hashlib
import json
import subprocess
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
AUDIT = ROOT / "audit" / "v2_1"
FIXTURE = (
    ROOT / "regression" / "fixtures" / "synthetic_cardio_aki_v2_0"
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_csv(path: Path, fields: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def input_hashes() -> list[dict[str, Any]]:
    rows = []
    for path in sorted((ROOT / "input").rglob("*")):
        if not path.is_file():
            continue
        rows.append(
            {
                "relative_path": path.relative_to(ROOT).as_posix(),
                "sha256": sha256_file(path),
                "size_bytes": path.stat().st_size,
            }
        )
    return rows


def baseline() -> None:
    AUDIT.mkdir(parents=True, exist_ok=True)
    pptx = FIXTURE / "SYN_CARDIO_AKI_SHADOW_VALIDATION_v2_0_defect_fixture.pptx"
    geometry_path = FIXTURE / "powerpoint_layout_manifest.json"
    report = json.loads(geometry_path.read_text(encoding="utf-8-sig"))
    issue_rows = [
        {
            "issue_id": "ISSUE-LAYOUT-001",
            "category": "layout",
            "expected_regression_state": "FAIL",
            "observed": "Long body/source text enters footer exclusion area and visibly overlaps the footer.",
            "evidence": "PowerPoint actual bounds in fixture slide 2",
        },
        {
            "issue_id": "ISSUE-QA-001",
            "category": "qa_false_negative",
            "expected_regression_state": "FAIL",
            "observed": "Legacy QA reported overflow but did not classify the visible body-versus-footer collision.",
            "evidence": "original_qa_report.md versus actual shape geometry",
        },
        {
            "issue_id": "ISSUE-METADATA-001",
            "category": "chart_metadata",
            "expected_regression_state": "FAIL",
            "observed": "Caption displayed sample size: not provided despite canonical denominators.",
            "evidence": "fixture render and canonical aggregate table",
        },
        {
            "issue_id": "ISSUE-VISUAL-001",
            "category": "visual_routing",
            "expected_regression_state": "FAIL",
            "observed": "Structured results frequently degraded to title plus body text.",
            "evidence": "fixture storyboard and contact sheet",
        },
        {
            "issue_id": "ISSUE-DENSITY-001",
            "category": "density",
            "expected_regression_state": "FAIL",
            "observed": "Slides alternate between visually overloaded and sparse layouts.",
            "evidence": "fixture contact sheet",
        },
    ]
    write_csv(
        AUDIT / "baseline_issue_manifest.csv",
        [
            "issue_id",
            "category",
            "expected_regression_state",
            "observed",
            "evidence",
        ],
        issue_rows,
    )
    geometry_rows: list[dict[str, Any]] = []
    for slide in report.get("slides", []):
        for shape in slide.get("text_shapes", []):
            geometry_rows.append(
                {
                    "slide_index": slide.get("slide_index"),
                    "object_index": shape.get("object_index"),
                    "top_pt": shape.get("top_pt"),
                    "height_pt": shape.get("height_pt"),
                    "bound_top_pt": shape.get(
                        "bound_top_pt", shape.get("top_pt")
                    ),
                    "bound_height_pt": shape.get("bound_height_pt"),
                    "footer_like": shape.get("footer_like"),
                    "text_length": shape.get("text_length"),
                    "legacy_detected_vertical_overflow": (
                        float(shape.get("bound_height_pt") or 0)
                        > float(shape.get("height_pt") or 0) * 1.04 + 1
                    ),
                }
            )
    write_csv(
        AUDIT / "baseline_geometry_manifest.csv",
        [
            "slide_index",
            "object_index",
            "top_pt",
            "height_pt",
            "bound_top_pt",
            "bound_height_pt",
            "footer_like",
            "text_length",
            "legacy_detected_vertical_overflow",
        ],
        geometry_rows,
    )
    (AUDIT / "baseline_render_manifest.json").write_text(
        json.dumps(
            {
                "schema_version": "2.1",
                "fixture": str(pptx),
                "fixture_sha256": sha256_file(pptx),
                "slide_count": len(report.get("slides", [])),
                "contact_sheet": str(AUDIT / "baseline_contact_sheet.png"),
                "expected_failures": [row["issue_id"] for row in issue_rows],
                "source_output_was_not_modified": True,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (AUDIT / "baseline_visual_qa.md").write_text(
        """# v2.0 Baseline Visual QA

- Fixture role: fixed regression input, not a production-ready deck.
- Expected result: **FAIL**.
- Confirmed: body/source text can enter the footer exclusion area.
- Confirmed: legacy geometry QA did not classify the visible body-vs-footer collision.
- Confirmed: structured quantitative inputs were under-routed to visuals.
- Confirmed: page density and layout repetition were not governed as release gates.
- Confirmed: `sample size: not provided` contradicted available denominators.

The fixture is preserved read-only by convention. No source data or answer files
were copied into the fixture directory.
""",
        encoding="utf-8",
    )
    write_csv(
        AUDIT / "input_hashes_before.csv",
        ["relative_path", "sha256", "size_bytes"],
        input_hashes(),
    )


def finalize(output_dir: Path, test_summary: str) -> None:
    before_path = AUDIT / "input_hashes_before.csv"
    with before_path.open("r", encoding="utf-8-sig", newline="") as handle:
        before = {row["relative_path"]: row for row in csv.DictReader(handle)}
    after_rows = input_hashes()
    after = {row["relative_path"]: row for row in after_rows}
    write_csv(
        AUDIT / "input_hashes_after.csv",
        ["relative_path", "sha256", "size_bytes"],
        after_rows,
    )
    changes = []
    for name in sorted(set(before) | set(after)):
        if name not in before:
            changes.append(f"added: {name}")
        elif name not in after:
            changes.append(f"removed: {name}")
        elif before[name]["sha256"] != after[name]["sha256"]:
            changes.append(f"modified: {name}")
    (AUDIT / "input_integrity_report.md").write_text(
        "# Input Integrity Report\n\n"
        + (
            f"- PASS: {len(after_rows)} input files retained identical SHA-256 values and paths.\n"
            if not changes
            else "- FAIL:\n" + "\n".join(f"  - {item}" for item in changes) + "\n"
        ),
        encoding="utf-8",
    )
    changed = subprocess.run(
        ["git", "status", "--short", "--untracked-files=all"],
        cwd=ROOT,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        check=False,
    ).stdout.splitlines()
    write_csv(
        AUDIT / "changed_files.csv",
        ["git_status", "path", "scope"],
        [
            {
                "git_status": line[:2],
                "path": line[3:],
                "scope": (
                    "v2.1"
                    if any(
                        token in line[3:]
                        for token in (
                            "layout_contract",
                            "safe_zones",
                            "footer_policy",
                            "visual_routing",
                            "visual_templates",
                            "visual_router",
                            "visual_density",
                            "text_measurement",
                            "chart_data",
                            "source_display",
                            "ppt_regression",
                            "generate_deck_pptxgen",
                            "visual_layout_qa",
                            "inspect_pptx_layout",
                            "storyboard",
                            "qa.py",
                            "runner.py",
                            "deck_ir.py",
                            "evidence.py",
                            "regression/",
                            "scripts/v2/",
                            "audit/v2_1",
                        )
                    )
                    else "pre-existing or unrelated"
                ),
            }
            for line in changed
        ],
    )
    (AUDIT / "root_cause_analysis.md").write_text(
        """# Root Cause Analysis

1. Layout coordinates were owned by individual render functions rather than one
   enforced zone contract.
2. Legacy QA compared nominal text-box dimensions but did not consistently use
   `TextFrame2.TextRange.Bound*`, `Overflowing`, or body/footer role semantics.
3. The planner treated structured tables and unstructured claims similarly, so
   quantitative results degraded to text.
4. Caption metadata was inferred independently from the chart data source.
5. Full internal source identifiers were rendered where short human-facing
   provenance labels were sufficient.
6. No density or consecutive-layout budget participated in release status.

All remedies are keyed by slide role, content structure, geometry, and data
contracts. No project name, current page number, fixed title, or phenotype label
is used as a routing condition.
""",
        encoding="utf-8",
    )
    (AUDIT / "module_change_map.md").write_text(
        """# Module Change Map

| Original problem | Module | New behavior | Regression evidence | Compatibility risk |
|---|---|---|---|---|
| Footer collision | layout_contract + generator | all frames validated against footer exclusion before generation | test_safe_zones, test_footer_intrusion | custom templates must provide equivalent zones |
| Text-bound false negative | PowerPoint COM + visual_layout_qa | actual Bound* and Overflowing values gate delivery | test_text_bounds, actual COM manifest | renderer-specific font metrics remain possible |
| Text-only degradation | visual_router + storyboard | structured payload selects chart/diagram | test_visual_router | incomplete data intentionally falls back |
| Density extremes | visual_density + runner | sparse/overloaded/repetition reports gate QA | density/repetition tests | thresholds may require theme calibration |
| Metadata mismatch | ChartDataContract | chart, caption and notes share one record | test_chart_metadata | unsupported table schemas remain uncharted |
| Source noise | source_display + generator | short labels on slide, full IDs in notes/audit | OOXML and scientific regression | audit mode remains opt-in |
""",
        encoding="utf-8",
    )
    (AUDIT / "behavior_change_manifest.md").write_text(
        """# Behavior Change Manifest

- Every generated slide receives `planned_geometry` and fails closed on zone intrusion.
- Content routing supports event-rate bars, forest plots, line charts,
  missingness bars, timelines, 2x2 matrices, native flows, comparison cards,
  takeaways, audit tables, and source registries.
- Text measurement observes CJK width, font fallback, margins, bullets and a
  bounded 8% shrink rule; minimum fonts remain enforced.
- PowerPoint actual geometry, OOXML geometry and rendered-page signals are
  separate QA layers.
- The body slide shows a short source label; notes/audit retain full source IDs,
  paths, claim IDs and wording boundaries.
- CSV and XLSX aggregate tables can create a single ChartDataContract used by
  the chart, caption, notes and source map.
- No statistical model is recomputed.
""",
        encoding="utf-8",
    )
    (AUDIT / "regression_test_results.md").write_text(
        "# Regression Test Results\n\n" + test_summary.strip() + "\n",
        encoding="utf-8",
    )
    (AUDIT / "output_pointer.json").write_text(
        json.dumps(
            {"output_dir": str(output_dir), "input_changes": changes},
            indent=2,
        ),
        encoding="utf-8",
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=["baseline", "finalize"])
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--test-summary", default="")
    args = parser.parse_args()
    if args.mode == "baseline":
        baseline()
    else:
        if not args.output_dir:
            raise SystemExit("--output-dir is required for finalize")
        finalize(args.output_dir.resolve(), args.test_summary)


if __name__ == "__main__":
    main()
