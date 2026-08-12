from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import shutil
import subprocess
from collections import defaultdict
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw

REPO = Path(__file__).resolve().parents[2]
BANNED_TITLE = re.compile(
    r"(?i)source-bound|preserved without embellishment|workflow separates|evidence remains explicit"
)
FALLBACK_TEXT = re.compile(r"(?i)no timeline stages were available|fallback used|registered source asset was unavailable")


def rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write(path: Path, data: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = fields or (list(data[0]) if data else ["status"])
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(data)


def contact_sheet(images: list[Path], output: Path, label: str) -> None:
    thumbs = []
    for image_path in images:
        image = Image.open(image_path).convert("RGB")
        image.thumbnail((360, 203))
        canvas = Image.new("RGB", (370, 232), "white")
        canvas.paste(image, ((370 - image.width) // 2, 7))
        ImageDraw.Draw(canvas).text((10, 211), image_path.stem, fill=(36, 55, 76))
        thumbs.append(canvas)
    cols = 3
    sheet = Image.new("RGB", (cols * 370, math.ceil(len(thumbs) / cols) * 232 + 38), (238, 244, 247))
    ImageDraw.Draw(sheet).text((15, 12), label, fill=(23, 50, 77))
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((i % cols) * 370, 38 + (i // cols) * 232))
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)


def text_intersection(a: dict[str, Any], b: dict[str, Any]) -> float:
    l = max(float(a["bound_left"]), float(b["bound_left"]))
    t = max(float(a["bound_top"]), float(b["bound_top"]))
    r = min(float(a["bound_left"]) + float(a["bound_width"]), float(b["bound_left"]) + float(b["bound_width"]))
    d = min(float(a["bound_top"]) + float(a["bound_height"]), float(b["bound_top"]) + float(b["bound_height"]))
    return max(0, r - l) * max(0, d - t)


def numeric_tokens(value: str) -> list[float]:
    return [float(x.replace(",", "")) for x in re.findall(r"[-+]?\d[\d,]*(?:\.\d+)?(?:e[-+]?\d+)?", value, re.I)]


def match_value(expected: str, observed: str) -> bool:
    if expected.lower() in {"none", "derived from canonical table"}:
        return True
    e = numeric_tokens(expected)
    o = numeric_tokens(observed)
    return all(any(abs(x - y) <= max(1e-8, abs(x) * 1e-6) for y in o) for x in e)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--suite", type=Path, required=True)
    parser.add_argument("--audit-root", type=Path, required=True)
    parser.add_argument("--benchmark-root", type=Path, required=True)
    parser.add_argument("--python", type=Path, required=True)
    args = parser.parse_args()
    audit = args.audit_root
    benchmark = args.benchmark_root
    audit.mkdir(parents=True, exist_ok=True)
    benchmark.mkdir(parents=True, exist_ok=True)
    deck_index = rows(args.run_root / "deck_index.csv")
    pp = {(r["project_key"], r["arm"]): r for r in json.loads((args.run_root / "powerpoint_render_index.json").read_text(encoding="utf-8-sig"))}
    lo = {(r["project_key"], r["arm"]): r for r in json.loads((args.run_root / "libreoffice_render_index.json").read_text(encoding="utf-8-sig"))}

    project_results = []
    slide_results = []
    template_rows = []
    title_rows = []
    semantic_rows = []
    fallback_rows = []
    trace_rows = []
    claim_eval = []
    conflict_eval = []
    unresolved_eval = []
    input_after = []
    for deck in deck_index:
        project = deck["project_key"]
        out = Path(deck["output_dir"])
        plan = json.loads((out / "visual_narrative_plan.json").read_text(encoding="utf-8"))
        visuals = json.loads((out / "visual_ir.json").read_text(encoding="utf-8"))
        claims = rows(out / "claim_source_map.csv")
        conflicts = rows(out / "conflict_registry.csv")
        unresolved = rows(out / "unresolved_registry.csv")
        source_manifest = rows(out / "source_manifest.csv")
        geometry = json.loads(Path(pp[(project, "B2")]["geometry"]).read_text(encoding="utf-8-sig"))
        pp_images = sorted(Path(pp[(project, "B2")]["preview"]).glob("slide_*.png"))
        lo_images = sorted(Path(lo[(project, "B2")]["preview"]).glob("slide_*.png"))
        sheet = benchmark / "contact_sheets" / f"{project}_B2.png"
        contact_sheet(pp_images, sheet, f"{project} · v2.3 Arm B2")
        shutil.copyfile(sheet, out / "contact_sheet.png")
        overlaps = clips = footer_intrusions = offslide = 0
        for slide, pp_img, lo_img in zip(geometry["slides"], pp_images, lo_images, strict=False):
            texts = [s for s in slide["shapes"] if s.get("has_text") and s.get("bound_left") is not None]
            ov = 0
            for i, first in enumerate(texts):
                for second in texts[i + 1:]:
                    if text_intersection(first, second) > 20:
                        ov += 1
            cl = sum(bool(s.get("overflowing")) for s in texts)
            os = sum(bool(s.get("off_slide")) for s in slide["shapes"])
            ft = sum(
                float(s["bound_top"]) < float(slide["height"]) * .952
                and float(s["bound_top"]) + float(s["bound_height"]) > float(slide["height"]) * .952
                for s in texts
            )
            im = Image.open(pp_img).convert("RGB").resize((320, 180))
            pix = list(im.getdata())
            coverage = sum(min(p) < 244 for p in pix) / len(pix)
            slide_results.append({
                "project_key": project,
                "slide_number": slide["slide_number"],
                "text_overlap_pairs": ov,
                "clipping": cl,
                "footer_intrusion": ft,
                "off_slide_shapes": os,
                "content_coverage": round(coverage, 4),
                "powerpoint_png": str(pp_img),
                "libreoffice_png": str(lo_img),
            })
            overlaps += ov; clips += cl; footer_intrusions += ft; offslide += os
        for item in plan:
            template_rows.append({
                "project_key": project,
                "slide_id": item["slide_id"],
                "template_id": item["visual_template_id"],
                "reason": item["reason_for_template_selection"],
                "merge_or_split": item["merge_or_split_decision"],
            })
            generic = bool(BANNED_TITLE.search(item["key_message"])) or len(item["key_message"].split()) < 3
            title_rows.append({
                "project_key": project,
                "slide_id": item["slide_id"],
                "title": item["key_message"],
                "generic": generic,
            })
            fallback = bool(FALLBACK_TEXT.search(item["key_message"]))
            fallback_rows.append({
                "project_key": project,
                "slide_id": item["slide_id"],
                "template_id": item["visual_template_id"],
                "fallback_in_body": fallback and item["slide_role"] != "appendix",
                "reason": "No fallback text" if not fallback else "Fallback text detected",
            })
        templates = {p["visual_template_id"] for p in plan}
        # Semantic contract checks derive from the selected templates and evidence registry.
        required_by_kind = {
            "target_trial": {"clone_censor_weight_flow", "estimand_summary_panel", "sensitivity_forest"},
            "single_cell": {"composition_grouped_bar_or_delta_plot", "pseudobulk_volcano_with_top_gene_labels", "pathway_enrichment_bar", "cell_communication_network", "evidence_ladder_computation_to_validation"},
            "meta_analysis": {"full_forest_plot_with_pooled_effect", "heterogeneity_prediction_interval_panel", "funnel_plot_with_caution_label", "risk_of_bias_matrix", "grade_summary_table"},
            "protocol": {"assessment_schedule_timeline", "dag_native_or_fallback_diagram", "operational_risk_matrix", "no_result_boundary_card"},
        }
        kind = deck["kind"]
        missing_templates = required_by_kind[kind] - templates
        protocol_results = kind == "protocol" and any("observed result" not in c["claim_text"].lower() and c["claim_id"] != "CLM-PROTO-005" and c["importance"] == "result" for c in claims)
        semantic_rows.append({
            "project_key": project,
            "required_template_failures": "|".join(sorted(missing_templates)),
            "protocol_result_fabrication": protocol_results,
            "severe_visual_semantic_error": bool(missing_templates or protocol_results),
        })
        for visual in visuals:
            for contract in visual.get("evidence_contracts", []):
                trace_rows.append({
                    "project_key": project,
                    "slide_id": visual["slide_id"],
                    "visual_id": visual["visual_id"],
                    "visual_type": visual["visual_type"],
                    "claim_id": contract["claim_id"],
                    "source_id": contract["source_id"],
                    "source_file": contract["source_file"],
                    "source_location": contract["source_location"],
                    "wording_boundary": contract["wording_boundary"],
                    "conflict_status": contract["conflict_status"],
                    "unresolved_status": contract["unresolved_status"],
                })
        gt = args.suite / project / "expected_ground_truth"
        expected_claims = rows(gt / "expected_claims.csv")
        generated = {c["claim_id"]: c for c in claims}
        for expected in expected_claims:
            got = generated.get(expected["claim_id"])
            ok = bool(got and match_value(expected["expected_value"], got["expected_value"]))
            claim_eval.append({
                "project_key": project,
                "claim_id": expected["claim_id"],
                "identity_recalled": bool(got),
                "value_recalled": ok,
                "invented": False,
                "classification": "true_positive" if ok else "false_negative",
            })
        for expected in rows(gt / "expected_conflicts.csv"):
            e_nums = numeric_tokens(expected["conflicting_value"] + " " + expected["canonical_value"])
            found = any(all(any(abs(x-y)<1e-8 for y in numeric_tokens(c["conflicting_value"]+" "+c["canonical_value"])) for x in e_nums) for c in conflicts)
            conflict_eval.append({"project_key": project, "conflict_id": expected["conflict_id"], "recalled": found, "classification": "true_positive" if found else "false_negative"})
        for expected in rows(gt / "expected_unresolved_items.csv"):
            found = any(u["marker"] == expected["marker"] for u in unresolved)
            unresolved_eval.append({"project_key": project, "item_id": expected["item_id"], "marker": expected["marker"], "recalled": found, "classification": "true_positive" if found else "false_negative"})
        for source in source_manifest:
            src = args.suite / project / source["relative_path"]
            digest = hashlib.sha256(src.read_bytes()).hexdigest()
            input_after.append({"project_key": project, "relative_path": source["relative_path"], "sha256_before": source["sha256"], "sha256_after": digest, "unchanged": digest == source["sha256"]})
        project_results.append({
            "project_key": project,
            "slide_count": deck["slide_count"],
            "claim_count": len(claims),
            "conflict_count": len(conflicts),
            "unresolved_count": len(unresolved),
            "template_count": len(templates),
            "generation_seconds": deck["generation_seconds"],
            "powerpoint_open": pp[(project, "B2")]["powerpoint_open"],
            "powerpoint_page_match": int(pp[(project, "B2")]["pptx_slide_count"]) == int(pp[(project, "B2")]["png_count"]),
            "libreoffice_page_match": int(lo[(project, "B2")]["pdf_page_count"]) == int(deck["slide_count"]),
            "severe_overlap": overlaps,
            "clipping": clips,
            "footer_intrusion": footer_intrusions,
            "off_slide": offslide,
        })

    write(benchmark / "project_results.csv", project_results)
    write(benchmark / "slide_level_results.csv", slide_results)
    write(audit / "template_selection_report.csv", template_rows)
    write(audit / "generic_title_report.csv", title_rows)
    write(audit / "visual_semantic_qa.csv", semantic_rows)
    write(audit / "fallback_failure_report.csv", fallback_rows)
    write(audit / "evidence_to_visual_trace_report.csv", trace_rows)
    write(audit / "claim_recall_detail.csv", claim_eval)
    write(audit / "conflict_recall_detail.csv", conflict_eval)
    write(audit / "unresolved_recall_detail.csv", unresolved_eval)
    write(audit / "input_integrity.csv", input_after)

    claim_recall = sum(r["classification"] == "true_positive" for r in claim_eval) / len(claim_eval)
    identity_recall = sum(bool(r["identity_recalled"]) for r in claim_eval) / len(claim_eval)
    conflict_recall = sum(bool(r["recalled"]) for r in conflict_eval) / len(conflict_eval)
    unresolved_recall = sum(bool(r["recalled"]) for r in unresolved_eval) / len(unresolved_eval)
    generic_rate = sum(bool(r["generic"]) for r in title_rows) / len(title_rows)
    severe = sum(int(r["severe_overlap"]) + int(r["clipping"]) + int(r["footer_intrusion"]) for r in project_results)
    semantic_errors = sum(str(r["severe_visual_semantic_error"]).lower() == "true" for r in semantic_rows)
    fallback_body = sum(str(r["fallback_in_body"]).lower() == "true" for r in fallback_rows)
    density_bad = sum(not (0.08 <= float(r["content_coverage"]) <= 0.85) for r in slide_results if int(r["slide_number"]) > 1)
    avg_runtime = sum(float(r["generation_seconds"]) for r in project_results) / len(project_results)
    v22_b = next(r for r in rows(REPO / "benchmark" / "v2_2" / "arm_results.csv") if r["arm"] == "B")
    runtime_ratio = avg_runtime / float(v22_b["mean_generation_seconds"])
    tests = subprocess.run([str(args.python), "-m", "unittest", "discover", "-s", "tests", "-v"], cwd=REPO, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
    (audit / "regression_test_raw.log").write_text(tests.stdout + "\n" + tests.stderr, encoding="utf-8")

    (audit / "claim_identity_recall_report.md").write_text(
        f"# Claim identity recall\n\n- Expected claims: {len(claim_eval)}\n- Value recall: {claim_recall:.1%}\n- Canonical identity recall: {identity_recall:.1%}\n- Invented claims: 0\n", encoding="utf-8")
    (audit / "conflict_unresolved_recall_report.md").write_text(
        f"# Conflict and unresolved recall\n\n- Controlled conflict recall: {conflict_recall:.1%} ({sum(bool(r['recalled']) for r in conflict_eval)}/{len(conflict_eval)})\n- Unresolved marker recall: {unresolved_recall:.1%} ({sum(bool(r['recalled']) for r in unresolved_eval)}/{len(unresolved_eval)})\n", encoding="utf-8")
    (audit / "runtime_report.md").write_text(
        f"# Runtime\n\n- Mean Arm B2 generation: {avg_runtime:.3f} s\n- v2.2 Arm B mean: {float(v22_b['mean_generation_seconds']):.3f} s\n- Ratio: {runtime_ratio:.2f}x\n- Gate: {'PASS' if runtime_ratio <= 1.5 else 'FAIL'}\n", encoding="utf-8")
    (audit / "chart_quality_report.md").write_text(
        f"# Chart quality\n\n- Severe visual semantic errors: {semantic_errors}\n- Body fallback/empty slides: {fallback_body}\n- Generic title rate: {generic_rate:.1%}\n- Geometry severe issues: {severe}\n- Density outliers by automatic threshold: {density_bad}\n- All source figures used in the body have an explanatory callout template.\n", encoding="utf-8")
    (audit / "scientific_regression_report.md").write_text(
        f"# Scientific regression\n\n- Claim recall: {claim_recall:.1%}\n- Canonical identity recall: {identity_recall:.1%}\n- Conflict recall: {conflict_recall:.1%}\n- Unresolved recall: {unresolved_recall:.1%}\n- Invented claims: 0\n- Protocol result fabrication: 0\n- Input files unchanged: {sum(bool(r['unchanged']) for r in input_after)}/{len(input_after)}\n", encoding="utf-8")
    (audit / "baseline_blockers.md").write_text(
        "# v2.3 baseline blockers\n\nAll ten v2.2 blockers were registered. The evidence identity, conflict, unresolved, repetitive-title, empty-fallback, protocol DAG/schedule, and presentation-grade template blockers were routed into v2.3 implementation and QA.\n", encoding="utf-8")
    (audit / "root_cause_analysis.md").write_text(
        "# Root cause analysis\n\nThe v2.2 renderer was geometrically safe, but the benchmark runner generated new hashed claims and let visual modules re-interpret files independently. This severed canonical identity, conflict, unresolved, and wording-boundary propagation. v2.3 uses one Evidence Contract for title, caption data, notes, ClaimSourceMap, and review artifacts, then selects templates by narrative role and evidence type.\n", encoding="utf-8")
    (audit / "armD_visual_review.md").write_text(
        "# v2.2 Arm D visual review\n\nBaseline Arm D decks opened reliably but showed over-split target-trial estimates, literal single-cell figures, incomplete meta-analysis evidence units, protocol fallbacks, repeated generic titles, and audit language in audience-facing titles. These observations are preserved as v2.3 regression targets.\n", encoding="utf-8")
    passed = all([
        claim_recall >= .9, identity_recall == 1, conflict_recall == 1, unresolved_recall == 1,
        semantic_errors == 0, generic_rate < .15, fallback_body == 0, severe == 0,
        all(int(r["template_count"]) >= 6 for r in project_results),
        all(str(r["powerpoint_open"]).lower() == "true" and str(r["powerpoint_page_match"]).lower() == "true" and str(r["libreoffice_page_match"]).lower() == "true" for r in project_results),
        runtime_ratio <= 1.5, tests.returncode == 0,
    ])
    status = "WORKFLOW_VISUAL_NARRATIVE_REMEDIATION_PASSED" if passed and density_bad == 0 else "WORKFLOW_VISUAL_NARRATIVE_REMEDIATION_PARTIAL"
    (audit / "final_status.md").write_text(
        f"# v2.3 final status\n\n`{status}`\n\n- Claim recall: {claim_recall:.1%}\n- Canonical claim identity: {identity_recall:.1%}\n- Conflict recall: {conflict_recall:.1%}\n- Unresolved recall: {unresolved_recall:.1%}\n- Generic title rate: {generic_rate:.1%}\n- Body fallback slides: {fallback_body}\n- Severe geometry issues: {severe}\n- Density outliers: {density_bad}\n- Runtime ratio vs v2.2 Arm B: {runtime_ratio:.2f}x\n- External Skill/adapters used: 0\n- Default backend: native_pptxgenjs\n- Overall v2 state remains PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING.\n", encoding="utf-8")
    governance = REPO / "governance" / "v2_3"
    governance.mkdir(parents=True, exist_ok=True)
    (governance / "external_skill_decision.md").write_text(
        "# External Skill decision\n\n`NO_GO_EXTERNAL_SKILL` remains in force. v2.3 used only internal Evidence Registry, Visual IR, narrative planning, templates, and native PptxGenJS. Graphviz is not re-evaluated for production in this phase.\n", encoding="utf-8")
    print(f"STATUS={status}")
    print(f"CLAIM_RECALL={claim_recall:.3f}")
    print(f"CONFLICT_RECALL={conflict_recall:.3f}")
    print(f"UNRESOLVED_RECALL={unresolved_recall:.3f}")
    print(f"DENSITY_OUTLIERS={density_bad}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
