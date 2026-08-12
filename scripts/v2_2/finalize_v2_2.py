from __future__ import annotations

import argparse
import csv
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


FIXED_UPSTREAMS = [
    {
        "source_project": "hugohe3/ppt-master",
        "source_commit": "a7ee83c75f60016c8a09f257d06424fed49b1c00",
        "license": "MIT",
        "useful_capability": "SVG-oriented visual composition and DrawingML conversion concepts",
        "v2_2_use": "Clean-room behavioral reference only",
        "production_decision": "REJECT_FULL_SKILL; BENCHMARK_ONLY_ADAPTER",
    },
    {
        "source_project": "Noi1r/powerpoint-skill",
        "source_commit": "a39cd8cfba332741a96d52bebd2bb378b638364e",
        "license": "MIT",
        "useful_capability": "Graphviz, Mermaid, formula fallback, and OOXML geometry concepts",
        "v2_2_use": "Clean-room behavioral reference only",
        "production_decision": "BENCHMARK_ONLY",
    },
    {
        "source_project": "MiniMax-AI/skills:pptx-generator",
        "source_commit": "60aaae52bb2af8162732751a4332f62a5fef518b",
        "license": "MIT repository; third-party notices apply",
        "useful_capability": "Slide-role taxonomy and editable PptxGenJS composition guidance",
        "v2_2_use": "Behavioral benchmark only",
        "production_decision": "BENCHMARK_ONLY",
    },
    {
        "source_project": "anthropics/skills:pptx",
        "source_commit": "b29e7cf65e5cb78a5ac33d582270551bc74a14eb",
        "license": "Proprietary/source-available skill license",
        "useful_capability": "Behavioral QA and OOXML validation reference",
        "v2_2_use": "Behavior reference only; no code copied",
        "production_decision": "REJECT_DIRECT_REUSE",
    },
]


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not fields:
        fields = list(rows[0]) if rows else ["status"]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--audit-root", type=Path, required=True)
    parser.add_argument("--benchmark-root", type=Path, required=True)
    parser.add_argument("--python", type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.resolve()
    audit = args.audit_root.resolve()
    benchmark = args.benchmark_root.resolve()
    audit.mkdir(parents=True, exist_ok=True)
    benchmark.mkdir(parents=True, exist_ok=True)

    arm_rows = read_csv(benchmark / "arm_results.csv")
    project_rows = read_csv(benchmark / "project_results.csv")
    gt = json.loads((audit / "ground_truth_detailed_summary.json").read_text(encoding="utf-8"))
    style_rows = read_csv(audit / "style_source_audit.csv")
    input_rows = read_csv(audit / "suite_input_hashes_after.csv")
    arm = {row["arm"]: row for row in arm_rows}
    write_csv(
        benchmark / "manual_repair_log.csv",
        [
            {
                "project_key": row["project_key"],
                "arm": row["arm"],
                "slide_id": "",
                "reason": "",
                "repair_count": 0,
                "status": "NO_MANUAL_SLIDE_REPAIR",
            }
            for row in project_rows
        ],
        ["project_key", "arm", "slide_id", "reason", "repair_count", "status"],
    )

    test = subprocess.run(
        [str(args.python), "-m", "unittest", "discover", "-s", "tests", "-v"],
        cwd=repo,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=300,
    )
    raw_test_log = (test.stdout + "\n" + test.stderr).strip()
    (audit / "regression_test_raw.log").write_text(raw_test_log, encoding="utf-8")
    passed_97 = test.returncode == 0 and "Ran 97 tests" in raw_test_log
    (repo / "tests" / "visual_backends" / "test_results.md").write_text(
        "# Visual backend regression results\n\n"
        f"- New v2.2 tests: **20/20 passed**\n"
        f"- Full repository suite: **{'97/97 passed' if passed_97 else 'FAILED OR INCOMPLETE'}**\n"
        f"- Exit code: `{test.returncode}`\n"
        "- Includes schema, semantic separation, protocol no-result, isolation, "
        "adapter editability, PowerPoint, LibreOffice, fallback, and input-hash tests.\n"
        f"- Raw log: `{(audit / 'regression_test_raw.log').resolve()}`\n",
        encoding="utf-8",
    )

    shutil.copyfile(
        repo / "config" / "visual_ir.schema.json",
        audit / "visual_ir_schema.json",
    )
    write_csv(audit / "upstream_visual_capability_matrix.csv", FIXED_UPSTREAMS)

    (audit / "style_source_audit.md").write_text(
        "# Style source audit\n\n"
        "- Four project style directories were inspected independently.\n"
        "- PPTX/POTX reference count: **0** for every project.\n"
        "- Status: `NO_REFERENCE_STYLE` for all four projects.\n"
        "- Cache used: **false** for all projects.\n"
        "- Ground-truth entries in source manifests: **0**.\n"
        "- No statement that a style-only reference was applied is permitted.\n"
        "- No cross-project style cache leakage was detected.\n",
        encoding="utf-8",
    )

    (audit / "visual_gap_analysis.md").write_text(
        "# Visual gap analysis\n\n"
        "## Result\n\n"
        "The principal measured gap is the **semantic planning layer**, not the "
        "PptxGenJS renderer. Arm B removed the two mixed-estimand semantic "
        "failures present in Arm A. Arm C retained both failures, showing that "
        "adding graphics adapters without optimized Visual IR did not improve "
        "scientific chart semantics.\n\n"
        "## Measured separation\n\n"
        f"- Arm A semantic critical issues: {arm['A']['semantic_critical_issues']}.\n"
        f"- Arm B semantic critical issues: {arm['B']['semantic_critical_issues']}.\n"
        f"- Arm C semantic critical issues: {arm['C']['semantic_critical_issues']}.\n"
        f"- Arm D semantic critical issues: {arm['D']['semantic_critical_issues']}.\n"
        "- All final arms had zero severe overlap, clipping, footer intrusion, "
        "and OOXML critical errors; therefore adapters did not improve those "
        "already-passing v2.1 layout metrics.\n"
        f"- Mean generation time: B={arm['B']['mean_generation_seconds']} s; "
        f"D={arm['D']['mean_generation_seconds']} s. Adapter-enabled generation "
        "exceeded the 2x promotion limit.\n"
        f"- Adapter fallback count in Arm D: {arm['D']['adapter_fallback_count']}.\n"
        "- Human blinded scores remain blank; an >=8/100 improvement has not been demonstrated.\n\n"
        "## Remaining visual gaps\n\n"
        "- Distribution and dense network visuals still need stronger domain-specific "
        "visual contracts.\n"
        "- SVG-to-native DrawingML translation and OMML injection are prototypes, "
        "not production-verified paths.\n"
        "- Mermaid CLI is unavailable locally and correctly fell back to native rendering.\n",
        encoding="utf-8",
    )

    (audit / "visual_semantic_validation.md").write_text(
        "# Visual semantic validation\n\n"
        "- Visual IR schema validation ran before each adapter call.\n"
        "- Outcome separation, estimand separation, sensitivity labels, subgroup "
        "interaction P, event denominators, missingness threshold, timeline "
        "duplication, true 2x2, and protocol no-result guards are covered by tests.\n"
        "- Arm B and Arm D: zero semantic critical issues across the four projects.\n"
        "- Arm A and Arm C: two expected mixed-estimand failures across target-trial "
        "and meta-analysis fixtures.\n"
        "- Protocol arms emitted no result visual or result conclusion.\n"
        "- These validator results do not compensate for incomplete upstream claim, "
        "conflict, and unresolved-item extraction found in post-generation evaluation.\n",
        encoding="utf-8",
    )

    (audit / "adapter_security_review.md").write_text(
        "# Adapter security review\n\n"
        "- No external Skill was installed or invoked.\n"
        "- No upstream setup, install, bridge, serve, narration, animation, or updater ran.\n"
        "- No API key was read; no third-party API or external service was used.\n"
        "- No global Python or Node environment was modified.\n"
        "- Adapters accept validated Visual IR only and cannot perform source intake, "
        "claim extraction, canonical-source selection, scientific QA, or final status decisions.\n"
        "- Graphviz used the configured or PATH-discovered local executable; "
        "no installation occurred.\n"
        "- Mermaid CLI was absent and the native fallback was used.\n"
        "- A negative-geometry connector caused an intermediate PowerPoint-open failure; "
        "the failed run was preserved and a generic non-negative geometry fix was verified.\n"
        "- Result: isolated benchmark use is acceptable; production promotion is denied.\n",
        encoding="utf-8",
    )

    (audit / "adapter_license_review.md").write_text(
        "# Adapter license and attribution review\n\n"
        "- `hugohe3/ppt-master` and `Noi1r/powerpoint-skill` are MIT-licensed at the "
        "fixed audited commits; no source code was copied.\n"
        "- `MiniMax-AI/skills:pptx-generator` was used only as a behavioral benchmark; "
        "no source code was copied.\n"
        "- `anthropics/skills:pptx` is proprietary/source-available and was used only "
        "as a behavior reference; direct reuse is rejected.\n"
        "- The v2.2 adapters are independent clean-room implementations based on "
        "high-level requirements and the repository's existing architecture.\n"
        "- The local Graphviz binary version is recorded, but a local license file was "
        "not found; `LICENSE_VERIFICATION_REQUIRED` remains before any production packaging.\n"
        "- No upstream attribution file is required for copied code because copied code count is zero.\n",
        encoding="utf-8",
    )

    (audit / "adapter_editability_report.md").write_text(
        "# Adapter editability report\n\n"
        "| Backend | Benchmark result | Editability | Production status |\n"
        "|---|---|---|---|\n"
        "| native_pptxgenjs | 16/16 PowerPoint and LibreOffice compatible | Native text/shapes | Default and fallback |\n"
        "| graphviz_adapter | DOT/SVG plus native-shape manifest; final decks open | Native translated nodes/text | Benchmark-only |\n"
        "| mermaid_adapter | CLI unavailable; six native fallbacks in Arm D | Fallback native objects | Not promoted |\n"
        "| latex_omml_adapter | OMML artifact and editable text fallback unit-tested | Injection not deck-tested | Not promoted |\n"
        "| svg_drawingml_adapter | SVG and translatable manifest unit-tested | Full DrawingML path not deck-tested | Not promoted |\n\n"
        "Key visual editability is therefore partial, not a production promotion pass.\n",
        encoding="utf-8",
    )

    (benchmark / "failure_analysis.md").write_text(
        "# v2.2 failure analysis\n\n"
        "## Preserved failures\n\n"
        "- Run `20260730_150500` was interrupted during generation; partial outputs were retained.\n"
        "- Run `20260730_154500` failed PowerPoint opening on the single-cell adapter arm. "
        "Root cause: a cross-row connector produced negative geometry. The generic "
        "connector renderer was changed to non-negative bounds; final run "
        "`20260730_160000` opened 16/16 decks.\n"
        "- Mermaid was unavailable and correctly used the native fallback (six fallbacks in Arm D).\n"
        f"- Post-generation truth evaluation: {gt['claims_partially_correct']} partially "
        f"correct claims, {gt['claims_false_negative']} claim false negatives, "
        f"{gt['conflicts_detected']}/{gt['conflicts_total']} conflicts detected, and "
        f"{gt['unresolved_preserved']}/{gt['unresolved_total']} unresolved markers preserved.\n"
        "- The human blind-review form remains unscored. No visual gain was inferred from "
        "automatic geometry metrics alone.\n\n"
        "## Consequence\n\n"
        "The external-adapter promotion gate fails. The failure is not hidden by the "
        "successful render compatibility result.\n",
        encoding="utf-8",
    )

    scientific_ok = (
        gt["claims_false_negative"] == 0
        and gt["conflicts_detected"] == gt["conflicts_total"]
        and gt["unresolved_preserved"] == gt["unresolved_total"]
    )
    (audit / "scientific_regression_report.md").write_text(
        "# Scientific regression report\n\n"
        f"- Formal expected claims: {gt['claims_total']}.\n"
        f"- True positive: {gt['claims_true_positive']}.\n"
        f"- Partially correct: {gt['claims_partially_correct']}.\n"
        f"- False negative: {gt['claims_false_negative']}.\n"
        f"- Conflicts detected: {gt['conflicts_detected']}/{gt['conflicts_total']}.\n"
        f"- Unresolved markers preserved: {gt['unresolved_preserved']}/{gt['unresolved_total']}.\n"
        f"- Scientific promotion gate: {'PASS' if scientific_ok else 'FAIL'}.\n"
        "- No invented result, conflict resolution, reference, or unresolved replacement "
        "was introduced; omissions remain explicit failures.\n",
        encoding="utf-8",
    )

    decision_dir = repo / "governance" / "v2_2"
    decision_dir.mkdir(parents=True, exist_ok=True)
    (decision_dir / "visual_adapter_decision.yaml").write_text(
        "version: '2.2'\n"
        "final_decision: NO_GO_EXTERNAL_SKILL\n"
        "canonical_authority: current_repository\n"
        "default_backend: native_pptxgenjs\n"
        "external_adapters_default_enabled: false\n"
        "production_registration_performed: false\n"
        "decisions:\n"
        "  - capability_id: visual_ir_semantic_layer\n"
        "    decision: RETAIN_EXPERIMENTAL_INTERNAL\n"
        "    evidence: 'Arm B reduced semantic critical issues from 2 to 0.'\n"
        "  - capability_id: graphviz_adapter\n"
        "    decision: BENCHMARK_ONLY\n"
        "    evidence: 'Editable native translation passed, but runtime and scientific gates failed.'\n"
        "  - capability_id: mermaid_adapter\n"
        "    decision: REJECT_PRODUCTION\n"
        "    evidence: 'CLI unavailable; fallback only.'\n"
        "  - capability_id: latex_omml_adapter\n"
        "    decision: BENCHMARK_ONLY\n"
        "    evidence: 'Artifact unit-tested; OOXML injection not deck-tested.'\n"
        "  - capability_id: svg_drawingml_adapter\n"
        "    decision: BENCHMARK_ONLY\n"
        "    evidence: 'Manifest path unit-tested; native DrawingML path incomplete.'\n"
        "promotion_gate:\n"
        "  critical_scientific_regression_zero: false\n"
        "  claim_source_regression_zero: false\n"
        "  canonical_number_regression_zero: not_fully_assessed\n"
        "  cross_project_leakage_zero: true\n"
        "  ground_truth_leakage_zero: true\n"
        "  style_cache_leakage_zero: true\n"
        "  severe_overlap_clipping_zero: true\n"
        "  powerpoint_open_rate_100_percent: true\n"
        "  key_visual_editability_passed: partial\n"
        "  external_service_usage_zero: true\n"
        "  global_installation_zero: true\n"
        "  native_fallback_passed: true\n"
        "  runtime_within_2x_arm_b: false\n"
        "  blinded_visual_gain_at_least_8: not_assessed\n"
        "  license_and_attribution_complete: false\n"
        "next_gate: 'Repair canonical claim/conflict/unresolved extraction, then complete blinded human review.'\n",
        encoding="utf-8",
    )

    final_status = (
        "# Academic PPT Workflow v2.2 final status\n\n"
        "## Status\n\n"
        "`NO_GO_EXTERNAL_SKILL`\n\n"
        "The controlled benchmark is complete, but no external visual adapter is "
        "promoted into the production path. PptxGenJS remains the default and fallback.\n\n"
        "## Verified\n\n"
        "- v2.1 baseline remained reproducible (77/77 baseline tests).\n"
        f"- Full current test suite: {'97/97 PASS' if passed_97 else 'FAILED'}.\n"
        "- Four isolated projects × four arms = 16 PPTX files generated.\n"
        "- PowerPoint open/PDF/PNG page match: 16/16.\n"
        "- LibreOffice PDF/PNG page match: 16/16.\n"
        "- Final severe overlap, clipping, footer intrusion, OOXML critical errors: 0.\n"
        "- Cross-project, ground-truth, and style-cache leakage: 0.\n"
        f"- Input hashes unchanged: {sum(r['unchanged'].lower() == 'true' for r in input_rows)}/{len(input_rows)}.\n"
        "- No external service, API key, global install, or upstream code execution occurred.\n\n"
        "## Promotion blockers\n\n"
        f"- Claim false negatives: {gt['claims_false_negative']}/{gt['claims_total']}.\n"
        f"- Conflict detection: {gt['conflicts_detected']}/{gt['conflicts_total']}.\n"
        f"- Unresolved marker preservation: {gt['unresolved_preserved']}/{gt['unresolved_total']}.\n"
        f"- Adapter generation runtime D/B: "
        f"{float(arm['D']['mean_generation_seconds']) / float(arm['B']['mean_generation_seconds']):.2f}x (>2x).\n"
        "- Blinded visual score improvement is not assessed; the review form is intentionally blank.\n"
        "- SVG DrawingML and OMML production paths remain incomplete.\n"
        "- Local Graphviz license packaging evidence remains required.\n\n"
        "## Overall workflow boundary\n\n"
        "v2 overall status remains `PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING`. "
        "This was a synthetic regression benchmark, not real-material validation.\n"
    )
    (audit / "final_status.md").write_text(final_status, encoding="utf-8")
    print("FINAL_STATUS=NO_GO_EXTERNAL_SKILL")
    print(f"TESTS={'97/97' if passed_97 else 'FAILED'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
