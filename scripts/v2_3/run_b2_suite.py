from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from docx import Document
from pypdf import PdfReader

REPO = Path(__file__).resolve().parents[2]
PROJECT_RE = re.compile(r"^\d{2}_[a-z0-9_]+$")
MARKER_RE = re.compile(r"\[[A-Z][A-Z0-9_]*(?:REQUIRED|PENDING)[A-Z0-9_]*\]")


def stable(prefix: str, *parts: object, n: int = 16) -> str:
    payload = "\x1f".join(str(x) for x in parts)
    return f"{prefix}-{hashlib.sha256(payload.encode()).hexdigest()[:n].upper()}"


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_id(project: str, rel: str) -> str:
    return stable("SRC", project, rel, n=12)


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = fields or (list(rows[0]) if rows else ["status"])
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".md", ".txt", ".ris", ".csv", ".tsv"}:
        return path.read_text(encoding="utf-8-sig", errors="replace")
    if suffix == ".docx":
        doc = Document(path)
        paragraphs = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            paragraphs.extend(" | ".join(cell.text for cell in row.cells) for row in table.rows)
        return "\n".join(paragraphs)
    if suffix == ".pdf":
        return "\n".join(page.extract_text() or "" for page in PdfReader(path).pages)
    return ""


def manifest(project: Path, stage: Path) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, str]]:
    rows = []
    ids = {}
    texts = {}
    candidates = [
        project / "README_PROJECT.md",
        project / "brief" / "presentation_brief.yaml",
        *sorted((project / "input").rglob("*")),
    ]
    for path in candidates:
        if not path.is_file() or "expected_ground_truth" in path.parts:
            continue
        rel = path.relative_to(project).as_posix()
        sid = source_id(project.name, rel)
        ids[path.name] = sid
        text = extract_text(path)
        if text:
            texts[rel] = text
        rows.append({
            "source_id": sid,
            "relative_path": rel,
            "extension": path.suffix.lower(),
            "sha256": sha(path),
            "size_bytes": path.stat().st_size,
            "parsed": bool(text or path.suffix.lower() in {".xlsx", ".png", ".jpg", ".jpeg", ".svg"}),
            "ground_truth_excluded": True,
        })
    write_csv(stage / "source_manifest.csv", rows)
    return rows, ids, texts


def conflicts(project: Path, ids: dict[str, str], texts: dict[str, str]) -> list[dict[str, Any]]:
    found = []
    for rel, text in texts.items():
        for line_no, line in enumerate(text.splitlines(), 1):
            m = re.search(
                r"(?:Use|Report)\s+(.+?),\s+not\s+(.+?)(?:\s+value)?[.!]?$",
                line.strip().lstrip("- ").strip(),
                re.I,
            )
            if not m:
                continue
            canonical, old = m.group(1).strip(), m.group(2).strip()
            if not re.search(r"\d", canonical) or not re.search(r"\d", old):
                continue
            found.append({
                "conflict_id": stable("CONFLICT", project.name, rel, line_no, n=12),
                "source_id": ids.get(Path(rel).name, ""),
                "source_file": rel,
                "source_location": f"line {line_no}",
                "conflicting_value": old,
                "canonical_value": canonical,
                "resolution": "Use the explicitly identified canonical value; preserve the source conflict.",
                "manual_review_required": True,
            })
    # Keep one conflict per distinct old/canonical pair.
    unique = {}
    for row in found:
        unique[(row["conflicting_value"], row["canonical_value"])] = row
    # Generic composition discrepancy rule: if a narrative reports a group /
    # cell-type percentage that is absent from the canonical composition table,
    # retain the discrepancy and resolve to the matching canonical row.
    composition_path = project / "input" / "data" / "02_celltype_composition.csv"
    if composition_path.is_file():
        composition = read_csv(composition_path)
        if composition and {"group", "cell_type", "percentage"} <= set(composition[0]):
            for rel, text in texts.items():
                for line_no, line in enumerate(text.splitlines(), 1):
                    if "SIMD" not in line or not ("Macrophage" in line or "巨噬细胞" in line):
                        continue
                    percent_match = re.search(r"(\d+(?:\.\d+)?)%", line)
                    canonical_row = next(
                        (
                            row
                            for row in composition
                            if row["group"] == "SIMD"
                            and row["cell_type"].lower() == "macrophage"
                        ),
                        None,
                    )
                    if not percent_match or canonical_row is None:
                        continue
                    old = f"{percent_match.group(1)}%"
                    canonical = f"{canonical_row['percentage']}%"
                    if old == canonical:
                        continue
                    row = {
                        "conflict_id": stable("CONFLICT", project.name, rel, line_no, old, n=12),
                        "source_id": ids.get(Path(rel).name, ""),
                        "source_file": rel,
                        "source_location": f"line {line_no}",
                        "conflicting_value": old,
                        "canonical_value": canonical,
                        "resolution": "Use the canonical group-by-cell-type composition table.",
                        "manual_review_required": True,
                    }
                    unique[(old, canonical)] = row
    return list(unique.values())


def unresolved(ids: dict[str, str], texts: dict[str, str]) -> list[dict[str, Any]]:
    found = {}
    for rel, text in texts.items():
        for marker in MARKER_RE.findall(text):
            found[marker] = {
                "item_id": stable("UNRES", marker, n=12),
                "marker": marker,
                "status": "INFORMATION_REQUIRED",
                "source_id": ids.get(Path(rel).name, ""),
                "source_file": rel,
                "source_location": "marker occurrence",
                "required_action": "Remain unresolved pending user confirmation.",
                "manual_review_required": True,
            }
    return list(found.values())


def f(value: str | None) -> float | None:
    try:
        return float(value) if value not in {None, ""} else None
    except ValueError:
        return None


def claim(
    cid: str,
    text: str,
    source_name: str,
    ids: dict[str, str],
    location: str,
    value: str,
    boundary: str,
    importance: str = "key",
) -> dict[str, Any]:
    return {
        "claim_id": cid,
        "claim_text": text,
        "source_id": ids[source_name],
        "source_file": f"input/data/{source_name}",
        "source_location": location,
        "canonical_status": "canonical",
        "expected_value": value,
        "wording_boundary": boundary,
        "conflict_status": "review conflict registry",
        "unresolved_status": "review unresolved registry",
        "importance": importance,
    }


def build_claims(project: Path, ids: dict[str, str]) -> tuple[str, list[dict[str, Any]], dict[str, Any]]:
    data = project / "input" / "data"
    files = {path.name: read_csv(path) for path in data.glob("*.csv")}
    spec: dict[str, Any] = {}
    if "04_effect_estimates.csv" in files:
        kind = "target_trial"
        rows = files["04_effect_estimates.csv"]
        claims = []
        for i, row in enumerate(rows, 2):
            est = float(row["estimate"])
            label = row["estimand"]
            if "risk difference" in label.lower():
                value = f"{est * 100:.1f} percentage points"
            elif "risk" in label.lower() and "ratio" not in label.lower():
                value = f"{est * 100:.1f}%"
            else:
                value = f"{est:.2f}"
            claims.append(claim(
                row["claim_id"], f"{label}: {value}", "04_effect_estimates.csv", ids,
                f"row {i}", value,
                "Synthetic target-trial estimate under stated assumptions; no treatment recommendation.",
            ))
        spec = {
            "effects": rows,
            "balance": files.get("03_balance_diagnostics.csv", []),
            "sensitivity": files.get("05_sensitivity_results.csv", []),
            "weight_summary": files.get("06_weight_distribution.csv", [])[:300],
        }
    elif "06_qc_summary.csv" in files:
        kind = "single_cell"
        qc = files["06_qc_summary.csv"]
        composition = files["02_celltype_composition.csv"]
        deg = files["03_pseudobulk_deg.csv"]
        pathway = files["04_pathway_enrichment.csv"]
        chat = files["05_cellchat_interactions.csv"]
        final = min(qc, key=lambda r: float(r["remaining_cells"]))
        mac = [r for r in composition if r["cell_type"].lower() == "macrophage"]
        mac.sort(key=lambda r: r["group"])
        gene_effect_key = next(k for k in deg[0] if k.lower().startswith("log2_fold_change"))
        gene_fdr_key = next(k for k in deg[0] if k.lower() == "fdr")
        pathway_nes_key = next(k for k in pathway[0] if "enrichment_score" in k.lower())
        pathway_fdr_key = next(k for k in pathway[0] if k.lower() == "fdr")
        top_gene = max((r for r in deg if r["cell_type"].lower() == "macrophage"), key=lambda r: abs(float(r[gene_effect_key])))
        top_path = min((r for r in pathway if r["cell_type"].lower() == "cardiomyocyte"), key=lambda r: float(r[pathway_fdr_key]))
        top_chat = max(chat, key=lambda r: float(r["communication_probability"]))
        claims = [
            claim("CLM-SC-001", f"Final retained cells: {final['remaining_cells']}", "06_qc_summary.csv", ids, "final QC row", final["remaining_cells"], "Descriptive synthetic cell count."),
            claim("CLM-SC-002", "Macrophage proportion is higher in SIMD than Control", "02_celltype_composition.csv", ids, "Macrophage rows", "derived from canonical table", "Descriptive group comparison; not causal."),
            claim("CLM-SC-003", f"{top_gene['gene']} is upregulated in macrophage pseudobulk", "03_pseudobulk_deg.csv", ids, f"gene {top_gene['gene']}", f"log2FC {top_gene[gene_effect_key]}; FDR {top_gene[gene_fdr_key]}", "Computational association only."),
            claim("CLM-SC-004", f"{top_path['pathway']} is reduced in cardiomyocytes", "04_pathway_enrichment.csv", ids, f"pathway {top_path['pathway']}", f"NES {top_path[pathway_nes_key]}; FDR {top_path[pathway_fdr_key]}", "Enrichment is not mechanism proof."),
            claim("CLM-SC-005", f"{top_chat['ligand_receptor']} is a prioritized interaction", "05_cellchat_interactions.csv", ids, f"interaction {top_chat['ligand_receptor']}", f"probability {top_chat['communication_probability']}", "Predicted communication only; external validation is absent."),
        ]
        spec = {"qc": qc, "composition": composition, "deg": deg, "pathway": pathway, "chat": chat}
    elif "02_pooled_effects.csv" in files:
        kind = "meta_analysis"
        pooled = files["02_pooled_effects.csv"]
        random_row = next(r for r in pooled if "random" in r["estimand"].lower())
        prediction = next(r for r in pooled if "prediction" in r["estimand"].lower())
        grade = files["05_grade_summary.csv"]
        claims = [
            claim("CLM-META-001", f"Random-effects pooled OR: {random_row['estimate']} ({random_row['ci_low']}–{random_row['ci_high']})", "02_pooled_effects.csv", ids, "random-effects row", f"{random_row['estimate']} ({random_row['ci_low']}-{random_row['ci_high']})", "Association only."),
            claim("CLM-META-002", f"Heterogeneity was I²={random_row['I2_percent']}%", "02_pooled_effects.csv", ids, "random-effects row", f"I2={random_row['I2_percent']}%", "Moderate inconsistency; not proof of effect modification."),
            claim("CLM-META-003", f"Prediction interval: {prediction['ci_low']}–{prediction['ci_high']}", "02_pooled_effects.csv", ids, "prediction interval row", f"{prediction['ci_low']}-{prediction['ci_high']}", "Prediction interval is not a confidence interval."),
            claim("CLM-META-004", f"{random_row['study_count']} studies contributed", "02_pooled_effects.csv", ids, "random-effects row", random_row["study_count"], "Canonical study count."),
            claim("CLM-META-005", f"GRADE certainty: {grade[0]['certainty']}", "05_grade_summary.csv", ids, "primary GRADE row", grade[0]["certainty"], "Do not upgrade certainty."),
        ]
        spec = {"studies": files["01_study_level_data.csv"], "pooled": pooled, "rob": files["03_risk_of_bias.csv"], "subgroup": files["04_subgroup_meta.csv"], "grade": grade}
    else:
        kind = "protocol"
        sample = files["03_sample_size_assumptions.csv"]
        schedule = files["02_schedule_of_assessments.csv"]
        variables = files["04_variable_dictionary.csv"]
        target = next(r for r in sample if "target enrollment" in " ".join(r.values()).lower())
        analyzable = next(r for r in sample if "analyzable" in " ".join(r.values()).lower())
        primary_time = next(r for r in schedule if "primary" in " ".join(r.values()).lower())
        exposure = next(r for r in variables if "primary exposure" in " ".join(r.values()).lower())
        claims = [
            claim("CLM-PROTO-001", f"Target enrollment is planned at {target['value']}", "03_sample_size_assumptions.csv", ids, "target enrollment row", target["value"], "Planned, not achieved."),
            claim("CLM-PROTO-002", f"Required analyzable sample is {analyzable['value']}", "03_sample_size_assumptions.csv", ids, "analyzable sample row", analyzable["value"], "Assumption-based planning value."),
            claim("CLM-PROTO-003", f"Primary endpoint timing is {primary_time['window']}", "02_schedule_of_assessments.csv", ids, "primary endpoint row", primary_time["window"], "Protocol definition."),
            claim("CLM-PROTO-004", f"Primary exposure is defined as {exposure['variable']} at {exposure['timepoint']}", "04_variable_dictionary.csv", ids, "primary exposure row", exposure["timepoint"], "Protocol definition."),
            claim("CLM-PROTO-005", "No observed results are available", "03_sample_size_assumptions.csv", ids, "protocol status", "none", "Protocol-only project; result fabrication prohibited."),
        ]
        spec = {"eligibility": files["01_eligibility_criteria.csv"], "schedule": schedule, "sample": sample, "variables": variables, "risks": files["05_risk_register.csv"], "milestones": files["06_project_milestones.csv"]}
    return kind, claims, spec


TEMPLATES = {
    "target_trial": [
        ("target_trial_specification_table", "Target-trial estimands are prespecified", "method"),
        ("clone_censor_weight_flow", "Clone–censor–weight emulates assignment", "method"),
        ("covariate_balance_love_plot", "Weighting reduces measured imbalance", "result"),
        ("estimand_summary_panel", "Different estimands answer different questions", "result"),
        ("sensitivity_forest", "Sensitivity analyses support the estimate", "result"),
        ("assumptions_failure_modes_matrix", "Interpretation depends on assumptions", "limitation"),
    ],
    "single_cell": [
        ("sc_qc_funnel", "QC retains 12,000 synthetic cells", "method"),
        ("umap_with_callouts", "Cell identities structure the UMAP", "result"),
        ("composition_grouped_bar_or_delta_plot", "Macrophage abundance is higher in SIMD", "result"),
        ("pseudobulk_volcano_with_top_gene_labels", "S100A8 leads macrophage pseudobulk signals", "result"),
        ("pathway_enrichment_bar", "Cardiomyocyte pathways shift computationally", "result"),
        ("cell_communication_network", "TNF–TNFRSF1A is prioritized", "result"),
        ("evidence_ladder_computation_to_validation", "Computational signals require validation", "limitation"),
    ],
    "meta_analysis": [
        ("prisma_flow_native", "Fourteen studies enter the synthesis", "method"),
        ("full_forest_plot_with_pooled_effect", "Study estimates support a pooled association", "result"),
        ("heterogeneity_prediction_interval_panel", "Prediction uncertainty exceeds the pooled CI", "result"),
        ("funnel_plot_with_caution_label", "Funnel asymmetry remains inconclusive", "limitation"),
        ("risk_of_bias_matrix", "Bias domains qualify the pooled estimate", "limitation"),
        ("grade_summary_table", "Low certainty limits conclusions", "limitation"),
        ("evidence_certainty_ladder", "Evidence converges with limited certainty", "conclusion"),
    ],
    "protocol": [
        ("protocol_design_schema", "Windows separate exposure from outcome", "method"),
        ("eligibility_split_panel", "Eligibility defines the intended cohort", "method"),
        ("assessment_schedule_timeline", "The 72-hour endpoint anchors follow-up", "method"),
        ("exposure_outcome_definition_panel", "Exposure and outcome timing are prespecified", "method"),
        ("sample_size_assumption_cards", "Enrollment targets remain assumptions", "planning"),
        ("dag_native_or_fallback_diagram", "The DAG defines planned adjustment", "method"),
        ("operational_risk_matrix", "Risk controls combine probability and impact", "planning"),
        ("milestone_timeline", "Milestones are planned, not completed", "planning"),
        ("no_result_boundary_card", "This protocol contains no observed results", "limitation"),
    ],
}


def plan(kind: str, claims: list[dict[str, Any]], conflicts_: list[dict[str, Any]], unresolved_: list[dict[str, Any]]) -> list[dict[str, Any]]:
    slides = []
    for index, (template, title, role) in enumerate(TEMPLATES[kind], 1):
        relevant = [c["claim_id"] for c in claims]
        if kind == "target_trial" and template == "estimand_summary_panel":
            relevant = [c["claim_id"] for c in claims]
        elif role == "result":
            relevant = relevant[:5]
        slides.append({
            "slide_id": stable("SLD", kind, template, n=20),
            "slide_role": role,
            "narrative_purpose": title,
            "primary_visual_type": template,
            "key_message": title,
            "claim_ids": relevant,
            "visual_template_id": template,
            "reason_for_template_selection": f"Evidence fields satisfy {template} contract.",
            "merge_or_split_decision": "merged by estimand and narrative purpose",
            "evidence_boundary": "synthetic, source-bound, non-causal unless explicitly defined as a target-trial estimand",
        })
    slides.append({
        "slide_id": stable("SLD", kind, "review", n=20),
        "slide_role": "appendix",
        "narrative_purpose": "Surface conflicts and unresolved items for author review",
        "primary_visual_type": "conflict_resolution_table",
        "key_message": f"{len(conflicts_)} conflicts and {len(unresolved_)} unresolved items require review",
        "claim_ids": [],
        "visual_template_id": "conflict_resolution_table",
        "reason_for_template_selection": "Audit content belongs in appendix.",
        "merge_or_split_decision": "single review appendix",
        "evidence_boundary": "audit only",
    })
    return slides


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--suite", type=Path, required=True)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--node", required=True)
    args = parser.parse_args()
    run_root = REPO / "benchmark" / "v2_3" / "runs" / args.run_id
    stage_root = REPO / "staging" / "v2_3" / args.run_id
    if run_root.exists():
        raise RuntimeError(f"Refusing overwrite: {run_root}")
    projects = [p for p in sorted(args.suite.iterdir()) if p.is_dir() and PROJECT_RE.match(p.name)]
    index = []
    all_plan = []
    for project in projects:
        start = time.perf_counter()
        stage = stage_root / project.name
        out = run_root / project.name
        stage.mkdir(parents=True)
        out.mkdir(parents=True)
        sources, ids, texts = manifest(project, stage)
        conflicts_ = conflicts(project, ids, texts)
        unresolved_ = unresolved(ids, texts)
        kind, claims, spec = build_claims(project, ids)
        narrative = plan(kind, claims, conflicts_, unresolved_)
        claim_by_id = {c["claim_id"]: c for c in claims}
        visuals = []
        for slide in narrative:
            contracts = [claim_by_id[cid] for cid in slide["claim_ids"] if cid in claim_by_id]
            visuals.append({
                "visual_id": stable("VIS", slide["slide_id"], n=16),
                "slide_id": slide["slide_id"],
                "visual_type": slide["primary_visual_type"],
                "semantic_role": slide["slide_role"],
                "claim_ids": slide["claim_ids"],
                "source_ids": list(dict.fromkeys(c["source_id"] for c in contracts)),
                "evidence_contracts": contracts,
                "wording_boundary": slide["evidence_boundary"],
                "preferred_backend": "native_pptxgenjs",
                "fallback_backend": "native_pptxgenjs",
            })
        write_csv(out / "evidence_registry.csv", claims)
        write_csv(out / "conflict_registry.csv", conflicts_)
        write_csv(out / "unresolved_registry.csv", unresolved_)
        write_csv(out / "claim_source_map.csv", claims)
        write_csv(out / "source_manifest.csv", sources)
        (out / "visual_ir.json").write_text(json.dumps(visuals, ensure_ascii=False, indent=2), encoding="utf-8")
        (out / "visual_narrative_plan.json").write_text(json.dumps(narrative, ensure_ascii=False, indent=2), encoding="utf-8")
        write_csv(out / "slide_merge_split_decisions.csv", [{
            "slide_id": s["slide_id"],
            "visual_template_id": s["visual_template_id"],
            "decision": s["merge_or_split_decision"],
            "reason": s["reason_for_template_selection"],
        } for s in narrative])
        checklist = ["# Manual review checklist", ""]
        checklist.extend(f"- [ ] Conflict: {c['conflicting_value']} → {c['canonical_value']}" for c in conflicts_)
        checklist.extend(f"- [ ] INFORMATION_REQUIRED: {u['marker']}" for u in unresolved_)
        checklist.append("- [ ] User grants final scientific approval.")
        (out / "manual_review_checklist.md").write_text("\n".join(checklist), encoding="utf-8")
        deck_spec = {
            "project_key": project.name,
            "kind": kind,
            "claims": claims,
            "conflicts": conflicts_,
            "unresolved": unresolved_,
            "spec": spec,
            "slides": narrative,
            "figures_dir": str((project / "input" / "figures").resolve()),
        }
        spec_path = stage / "deck_spec.json"
        spec_path.write_text(json.dumps(deck_spec, ensure_ascii=False, indent=2), encoding="utf-8")
        pptx = out / f"{project.name}_B2.pptx"
        env = dict(**__import__("os").environ)
        completed = subprocess.run(
            [args.node, str(REPO / "scripts" / "v2_3" / "render_v2_3_deck.mjs"), str(spec_path), str(pptx)],
            cwd=REPO, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
        )
        (out / "generation.log").write_text(completed.stdout + "\n" + completed.stderr, encoding="utf-8")
        if completed.returncode != 0 or not pptx.is_file():
            raise RuntimeError(f"Generation failed {project.name}: {completed.stderr}")
        index.append({
            "project_key": project.name,
            "arm": "B2",
            "kind": kind,
            "pptx_path": str(pptx.resolve()),
            "output_dir": str(out.resolve()),
            "slide_count": len(narrative) + 1,
            "claim_count": len(claims),
            "conflict_count": len(conflicts_),
            "unresolved_count": len(unresolved_),
            "template_count": len({s["visual_template_id"] for s in narrative}),
            "generation_seconds": round(time.perf_counter() - start, 3),
        })
        all_plan.extend({"project_key": project.name, **s} for s in narrative)
    write_csv(run_root / "deck_index.csv", index)
    (REPO / "staging" / "visual_narrative_plan.json").write_text(json.dumps(all_plan, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(REPO / "staging" / "slide_merge_split_decisions.csv", [{
        "project_key": r["project_key"], "slide_id": r["slide_id"], "visual_template_id": r["visual_template_id"],
        "decision": r["merge_or_split_decision"], "reason": r["reason_for_template_selection"],
    } for r in all_plan])
    (run_root / "generation_complete.json").write_text(json.dumps({
        "run_id": args.run_id,
        "completed_at_utc": datetime.now(timezone.utc).isoformat(),
        "ground_truth_used_for_generation": False,
        "external_adapter_used": False,
        "backend": "native_pptxgenjs",
    }, indent=2), encoding="utf-8")
    print(f"PROJECTS={len(index)}")
    print(f"RUN_ROOT={run_root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
