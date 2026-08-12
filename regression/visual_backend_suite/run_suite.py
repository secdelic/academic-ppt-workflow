from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

from extensions.visual_backends.registry import render_visual
from scripts.academic_ppt.visual_ir import (
    make_visual_ir,
    stable_id,
    timeline_similarity,
    validate_visual_ir_schema,
    validate_visual_semantics,
)


PROJECT_PATTERN = re.compile(r"^\d{2}_[a-z0-9_]+$")
GROUND_TRUTH = "expected_ground_truth"
ARMS = {
    "A": {"visual_ir": False, "adapters": False},
    "B": {"visual_ir": True, "adapters": False},
    "C": {"visual_ir": False, "adapters": True},
    "D": {"visual_ir": True, "adapters": True},
}
RESULT_VISUALS = {
    "forest_plot",
    "sensitivity_plot",
    "subgroup_plot",
    "event_rate_plot",
    "weighted_risk_curve",
    "love_plot",
    "volcano",
    "enrichment",
    "funnel_plot",
    "distribution",
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def source_id(project_key: str, rel_path: str) -> str:
    token = hashlib.sha256(f"{project_key}\x1f{rel_path}".encode()).hexdigest()
    return f"SRC-{token[:12].upper()}"


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return list(reader.fieldnames or []), [dict(row) for row in reader]


def read_brief(path: Path) -> dict[str, str]:
    """Read the scalar brief fields needed by the benchmark without PyYAML."""
    result: dict[str, str] = {}
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        if not raw or raw.lstrip().startswith("#") or raw[:1].isspace() or ":" not in raw:
            continue
        key, value = raw.split(":", 1)
        value = value.strip().strip("\"'")
        if key.strip() and value and value not in {"|", ">"}:
            result[key.strip()] = value
    return result


def f(value: Any) -> float | None:
    if value in (None, "", "NA", "N/A", "NR", "not reported"):
        return None
    try:
        return float(str(value).replace("%", "").strip())
    except ValueError:
        return None


def find_column(headers: list[str], *patterns: str) -> str | None:
    lowered = {h.lower(): h for h in headers}
    for pattern in patterns:
        for low, original in lowered.items():
            if re.search(pattern, low):
                return original
    return None


def list_values(rows: list[dict[str, str]], column: str | None) -> list[Any]:
    return [row.get(column, "") for row in rows] if column else []


def numeric_values(rows: list[dict[str, str]], column: str | None) -> list[float | None]:
    return [f(row.get(column)) for row in rows] if column else []


def build_manifest(project_dir: Path, isolated_root: Path) -> tuple[list[dict[str, Any]], dict[str, str]]:
    project_key = project_dir.name
    rows: list[dict[str, Any]] = []
    mapping: dict[str, str] = {}
    for path in sorted((project_dir / "input").rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(project_dir).as_posix()
        if GROUND_TRUTH in path.parts:
            continue
        sid = source_id(project_key, rel)
        mapping[rel] = sid
        rows.append(
            {
                "source_id": sid,
                "project_key": project_key,
                "relative_path": rel,
                "extension": path.suffix.lower(),
                "size_bytes": path.stat().st_size,
                "sha256": sha256(path),
                "ground_truth_excluded": True,
                "style_reference": path.suffix.lower() in {".pptx", ".potx"},
            }
        )
    isolated_root.mkdir(parents=True, exist_ok=True)
    with (isolated_root / "source_manifest.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]) if rows else ["source_id"])
        writer.writeheader()
        writer.writerows(rows)
    return rows, mapping


def _backend_for(visual_type: str, adapters: bool) -> str:
    if not adapters:
        return "native_pptxgenjs"
    if visual_type in {"flow_diagram", "dag", "network"}:
        return "graphviz_adapter"
    if visual_type in {"timeline"}:
        return "mermaid_adapter"
    if visual_type in {"formula"}:
        return "latex_omml_adapter"
    if visual_type in {"evidence_ladder", "conceptual_framework"}:
        return "svg_drawingml_adapter"
    return "native_pptxgenjs"


def _visual(
    *,
    key: str,
    visual_type: str,
    source: str,
    role: str = "result",
    categories: list[str] | None = None,
    series: list[dict[str, Any]] | None = None,
    estimate: list[float | None] | None = None,
    ci_low: list[float | None] | None = None,
    ci_high: list[float | None] | None = None,
    numerator: list[float | None] | None = None,
    denominator: list[float | None] | None = None,
    interaction_p: list[float | None] | None = None,
    analysis_label: list[str] | None = None,
    sample_size: float | None = None,
    outcome: str | None = None,
    estimand: str | None = None,
    reference_group: str | None = None,
    reference_value: float | None = None,
    annotation_rule: dict[str, Any] | None = None,
    wording_boundary: str = "Association or descriptive evidence only; do not imply causality.",
) -> dict[str, Any]:
    return make_visual_ir(
        slide_key=key,
        visual_type=visual_type,
        semantic_role=role,
        source_ids=[source],
        data_contract_id=stable_id("DATA", source, key, length=12),
        categories=categories or [],
        series=series or [],
        estimate=estimate or [],
        ci_low=ci_low or [],
        ci_high=ci_high or [],
        numerator=numerator or [],
        denominator=denominator or [],
        interaction_p=interaction_p or [],
        analysis_label=analysis_label or [],
        sample_size=sample_size,
        outcome=outcome,
        estimand=estimand,
        reference_group=reference_group,
        reference_value=reference_value,
        annotation_rule=annotation_rule or {},
        wording_boundary=wording_boundary,
    )


def route_csv(path: Path, sid: str, protocol: bool) -> list[dict[str, Any]]:
    headers, rows = read_csv(path)
    lows = {header.lower() for header in headers}
    visuals: list[dict[str, Any]] = []
    if not rows:
        return visuals
    key = path.stem

    if {"covariate", "absolute_smd_unweighted", "absolute_smd_weighted"} <= lows:
        labels = list_values(rows, find_column(headers, r"^covariate$"))
        visuals.append(_visual(
            key=key, visual_type="love_plot", source=sid, categories=labels,
            series=[
                {"name": "Unweighted", "values": numeric_values(rows, find_column(headers, "unweighted"))},
                {"name": "Weighted", "values": numeric_values(rows, find_column(headers, "weighted"))},
            ],
            role="diagnostic", reference_value=0.1,
            wording_boundary="Balance diagnostic only; not an outcome estimate.",
        ))
    elif {"analysis", "risk_difference", "ci_low", "ci_high"} <= lows:
        labels = list_values(rows, find_column(headers, r"^analysis$"))
        visuals.append(_visual(
            key=key, visual_type="sensitivity_plot", source=sid,
            categories=labels, analysis_label=labels,
            estimate=numeric_values(rows, find_column(headers, "risk_difference")),
            ci_low=numeric_values(rows, find_column(headers, r"^ci_low$")),
            ci_high=numeric_values(rows, find_column(headers, r"^ci_high$")),
            estimand="risk difference", reference_value=0.0,
            wording_boundary="Sensitivity analyses support robustness assessment, not proof.",
        ))
    elif "stabilized_weight" in lows:
        values = [v for v in numeric_values(rows, find_column(headers, "stabilized_weight")) if v is not None]
        visuals.append(_visual(
            key=key, visual_type="distribution", source=sid, role="diagnostic",
            categories=["Stabilized weights"], estimate=values[:500],
            wording_boundary="Weight distribution is a diagnostic, not an outcome.",
        ))
    elif {"stage", "remaining_cells", "removed_at_stage"} <= lows:
        visuals.append(_visual(
            key=key, visual_type="flow_diagram", source=sid, role="method",
            categories=list_values(rows, find_column(headers, r"^stage$")),
            estimate=numeric_values(rows, find_column(headers, "remaining_cells")),
            wording_boundary="Synthetic QC workflow; cell counts are descriptive.",
        ))
    elif {"group", "cell_type", "cell_count", "percentage"} <= lows:
        categories = [f"{r['group']} · {r['cell_type']}" for r in rows]
        visuals.append(_visual(
            key=key, visual_type="composition_bar", source=sid,
            categories=categories,
            estimate=numeric_values(rows, find_column(headers, "percentage")),
            numerator=numeric_values(rows, find_column(headers, "cell_count")),
            denominator=[sum(f(x.get("cell_count")) or 0 for x in rows if x.get("group") == r.get("group")) for r in rows],
            wording_boundary="Synthetic composition is descriptive; not a clinical event rate.",
        ))
    elif {"gene", "log2_fold_change", "fdr", "cell_type"} <= lows or {"gene", "log2fc", "fdr", "cell_type"} <= lows:
        visuals.append(_visual(
            key=key, visual_type="volcano", source=sid,
            categories=list_values(rows, find_column(headers, r"^gene$")),
            estimate=numeric_values(rows, find_column(headers, "log2")),
            series=[{"fdr": numeric_values(rows, find_column(headers, "fdr"))}],
            wording_boundary="Computational differential expression; not mechanistic validation.",
        ))
    elif {"pathway", "nes", "fdr", "cell_type"} <= lows:
        visuals.append(_visual(
            key=key, visual_type="enrichment", source=sid,
            categories=list_values(rows, find_column(headers, "pathway")),
            estimate=numeric_values(rows, find_column(headers, r"^nes$")),
            series=[{"fdr": numeric_values(rows, find_column(headers, "fdr"))}],
            wording_boundary="Computational pathway enrichment only.",
        ))
    elif {"sender", "receiver", "ligand_receptor", "communication_probability"} <= lows:
        labels: list[str] = []
        for row in rows:
            labels.extend([row.get("sender", ""), row.get("receiver", "")])
        visuals.append(_visual(
            key=key, visual_type="network", source=sid,
            categories=list(dict.fromkeys(filter(None, labels))),
            series=rows[:20],
            wording_boundary="Computational communication inference; not confirmed signaling.",
        ))
    elif {"study_label", "odds_ratio", "ci_low", "ci_high"} <= lows:
        visuals.append(_visual(
            key=key, visual_type="forest_plot", source=sid,
            categories=list_values(rows, find_column(headers, "study_label")),
            estimate=numeric_values(rows, find_column(headers, "odds_ratio")),
            ci_low=numeric_values(rows, find_column(headers, "ci_low")),
            ci_high=numeric_values(rows, find_column(headers, "ci_high")),
            estimand="odds ratio", reference_value=1.0,
            wording_boundary="Study-level associations; heterogeneity and bias must remain visible.",
        ))
    elif {"subgroup", "pooled_or", "ci_low", "ci_high", "interaction_p"} <= lows:
        visuals.append(_visual(
            key=key, visual_type="subgroup_plot", source=sid,
            categories=list_values(rows, find_column(headers, "subgroup")),
            estimate=numeric_values(rows, find_column(headers, "pooled_or")),
            ci_low=numeric_values(rows, find_column(headers, "ci_low")),
            ci_high=numeric_values(rows, find_column(headers, "ci_high")),
            interaction_p=numeric_values(rows, find_column(headers, "interaction_p")),
            estimand="odds ratio", reference_value=1.0,
            wording_boundary="Point-estimate differences do not establish effect modification.",
        ))
    elif any("risk_of_bias" in h or h.startswith("rob_") for h in lows):
        label_col = find_column(headers, "study")
        visuals.append(_visual(
            key=key, visual_type="risk_matrix", source=sid, role="limitation",
            categories=list_values(rows, label_col), series=rows,
            wording_boundary="Risk-of-bias judgments are source-bound assessments.",
        ))
    elif "certainty" in lows or "grade" in lows:
        label_col = find_column(headers, "outcome", "domain")
        visuals.append(_visual(
            key=key, visual_type="grade_summary", source=sid, role="limitation",
            categories=list_values(rows, label_col), series=rows,
            wording_boundary="Certainty statements must preserve source grading.",
        ))
    elif protocol and ("eligibility" in key or {"criterion", "definition"} <= lows):
        label_col = find_column(headers, "criterion", "domain", "item")
        visuals.append(_visual(
            key=key, visual_type="flow_diagram", source=sid, role="method",
            categories=list_values(rows, label_col)[:8],
            wording_boundary="Protocol specification only; no observed recruitment or results.",
        ))
    elif protocol and ("schedule" in key or any("visit" in h or "timepoint" in h for h in lows)):
        label_col = find_column(headers, "assessment", "procedure", "visit")
        visuals.append(_visual(
            key=key, visual_type="timeline", source=sid, role="method",
            categories=list_values(rows, label_col)[:8], series=rows,
            wording_boundary="Planned assessment schedule; not completed observations.",
        ))
    elif protocol and ("sample_size" in key or any("assumption" in h for h in lows)):
        label_col = find_column(headers, "parameter", "assumption", "item")
        visuals.append(_visual(
            key=key, visual_type="formula", source=sid, role="method",
            categories=list_values(rows, label_col)[:8], series=rows,
            wording_boundary="Planning assumptions only; no observed effect estimate.",
        ))
    elif protocol and "risk" in key:
        label_col = find_column(headers, "risk")
        visuals.append(_visual(
            key=key, visual_type="risk_matrix", source=sid, role="planning",
            categories=list_values(rows, label_col)[:10], series=rows,
            wording_boundary="Operational risk register, not observed study outcomes.",
        ))
    elif protocol and "milestone" in key:
        label_col = find_column(headers, "milestone", "activity")
        visuals.append(_visual(
            key=key, visual_type="timeline", source=sid, role="planning",
            categories=list_values(rows, label_col)[:8], series=rows,
            wording_boundary="Planned project milestones; not completed events.",
        ))
    else:
        estimate_col = find_column(headers, r"^estimate$", "pooled_effect", "odds_ratio", "risk_difference")
        low_col = find_column(headers, r"^ci_low$")
        high_col = find_column(headers, r"^ci_high$")
        if estimate_col and low_col and high_col and not protocol:
            label_col = find_column(headers, "analysis", "estimand", "contrast", "outcome")
            outcomes = list_values(rows, find_column(headers, r"^outcome$"))
            estimands = list_values(rows, find_column(headers, r"^estimand$"))
            visuals.append(_visual(
                key=key, visual_type="forest_plot", source=sid,
                categories=list_values(rows, label_col) or [f"Row {i+1}" for i in range(len(rows))],
                estimate=numeric_values(rows, estimate_col),
                ci_low=numeric_values(rows, low_col),
                ci_high=numeric_values(rows, high_col),
                series=[
                    {"outcome": outcomes[i] if i < len(outcomes) else None,
                     "estimand": estimands[i] if i < len(estimands) else None}
                    for i in range(len(rows))
                ],
                outcome=outcomes[0] if outcomes and len(set(outcomes)) == 1 else None,
                estimand=estimands[0] if estimands and len(set(estimands)) == 1 else None,
                reference_value=1.0,
                wording_boundary="Frozen estimates only; do not recompute or imply causality.",
            ))
    return visuals


def split_mixed_forests(visuals: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for visual in visuals:
        if visual["visual_type"] != "forest_plot" or not visual.get("series"):
            result.append(visual)
            continue
        keys = []
        for row in visual["series"]:
            keys.append((row.get("outcome") or visual.get("outcome"), row.get("estimand") or visual.get("estimand")))
        unique = list(dict.fromkeys(keys))
        if len(unique) <= 1:
            result.append(visual)
            continue
        for number, key in enumerate(unique, start=1):
            idx = [i for i, value in enumerate(keys) if value == key]
            clone = json.loads(json.dumps(visual))
            clone["categories"] = [visual["categories"][i] for i in idx]
            for field in ("estimate", "ci_low", "ci_high", "p_value", "analysis_label"):
                clone[field] = [visual.get(field, [])[i] for i in idx if i < len(visual.get(field, []))]
            clone["series"] = [visual["series"][i] for i in idx]
            clone["outcome"], clone["estimand"] = key
            clone["slide_id"] = stable_id("SLD", visual["slide_id"], key, length=20)
            clone["visual_id"] = stable_id("VIS", clone["slide_id"], number, length=16)
            result.append(clone)
    return result


def add_figure_visuals(project_dir: Path, source_map: dict[str, str], existing: list[dict[str, Any]], protocol: bool) -> None:
    existing_types = {v["visual_type"] for v in existing}
    for path in sorted((project_dir / "input" / "figures").glob("*")):
        rel = path.relative_to(project_dir).as_posix()
        name = path.stem.lower()
        visual_type = "source_figure"
        if "timeline" in name:
            visual_type = "timeline"
        elif "dag" in name:
            visual_type = "dag"
        elif "prisma" in name or "recruit" in name:
            visual_type = "flow_diagram"
        elif "funnel" in name:
            visual_type = "funnel_plot"
        elif "risk_curve" in name:
            visual_type = "weighted_risk_curve"
        elif "umap" in name:
            visual_type = "umap"
        if visual_type in existing_types and visual_type not in {"source_figure", "weighted_risk_curve", "umap", "funnel_plot"}:
            continue
        role = "method" if protocol else ("result" if visual_type in RESULT_VISUALS else "evidence")
        visual = _visual(
            key=path.stem, visual_type=visual_type, source=source_map[rel], role=role,
            categories=[path.name],
            wording_boundary=(
                "Protocol figure only; no observed results."
                if protocol else
                "Use the registered synthetic source figure without scientific upgrading."
            ),
        )
        visual["annotation_rule"]["source_asset"] = str(path.resolve())
        existing.append(visual)


def build_visuals(project_dir: Path, source_map: dict[str, str], optimized: bool) -> tuple[list[dict[str, Any]], bool]:
    brief = read_brief(project_dir / "brief" / "presentation_brief.yaml")
    presentation_type = str(brief.get("presentation_type", "")).lower()
    protocol = "protocol" in presentation_type or "protocol" in project_dir.name.lower()
    visuals: list[dict[str, Any]] = []
    for path in sorted((project_dir / "input" / "data").glob("*.csv")):
        rel = path.relative_to(project_dir).as_posix()
        visuals.extend(route_csv(path, source_map[rel], protocol))
    add_figure_visuals(project_dir, source_map, visuals, protocol)
    if optimized:
        visuals = split_mixed_forests(visuals)
        # High-similarity timelines are kept only when semantic roles differ.
        filtered: list[dict[str, Any]] = []
        for visual in visuals:
            duplicate = False
            if visual["visual_type"] == "timeline":
                for prior in filtered:
                    if prior["visual_type"] == "timeline" and timeline_similarity(prior, visual) >= 0.82:
                        if prior["semantic_role"] == visual["semantic_role"]:
                            duplicate = True
                            break
            if not duplicate:
                filtered.append(visual)
        visuals = filtered
    return visuals, protocol


def assign_backends(visuals: list[dict[str, Any]], adapters: bool) -> None:
    for visual in visuals:
        backend = _backend_for(visual["visual_type"], adapters)
        visual["preferred_backend"] = backend
        visual["fallback_backend"] = "native_pptxgenjs"
        if backend in {"graphviz_adapter", "latex_omml_adapter"}:
            visual["editability_requirement"] = "native_required"
        elif backend in {"svg_drawingml_adapter", "mermaid_adapter"}:
            visual["editability_requirement"] = "vector_acceptable"


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not fields:
        fields = list(rows[0]) if rows else ["status"]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def make_claims(visuals: list[dict[str, Any]], protocol: bool) -> list[dict[str, Any]]:
    claims = []
    for visual in visuals:
        if protocol:
            text = f"Planned {visual['visual_type'].replace('_', ' ')} is specified in the synthetic protocol."
        elif visual["visual_type"] in RESULT_VISUALS:
            text = f"Synthetic {visual['visual_type'].replace('_', ' ')} is available from the registered source."
        else:
            text = f"{visual['visual_type'].replace('_', ' ').title()} is registered for this synthetic project."
        claims.append({
            "claim_id": stable_id("CLM", visual["visual_id"], length=16),
            "claim_text": text,
            "slide_id": visual["slide_id"],
            "visual_id": visual["visual_id"],
            "source_ids": "|".join(visual["source_ids"]),
            "wording_boundary": visual["wording_boundary"],
            "protocol_without_results": protocol,
        })
    return claims


def build_deck_spec(project_dir: Path, arm: str, visuals: list[dict[str, Any]], artifacts: dict[str, dict[str, Any]], protocol: bool) -> dict[str, Any]:
    brief = read_brief(project_dir / "brief" / "presentation_brief.yaml")
    project_title = brief.get("project_name") or project_dir.name
    slides = [{
        "slide_id": stable_id("SLD", project_dir.name, "cover", length=20),
        "title": str(project_title).replace("_", " "),
        "subtitle": f"Controlled synthetic visual benchmark · Arm {arm}",
        "role": "cover",
        "synthetic_warning": "COMPLEX SYNTHETIC FIXTURE · NOT REAL CLINICAL EVIDENCE",
        "sources": [],
    }]
    for index, visual in enumerate(visuals[:10], start=1):
        title_map = {
            "forest_plot": "Effect estimates remain source-bound and uncertainty-visible",
            "sensitivity_plot": "Sensitivity analyses are distinguished by analysis label",
            "subgroup_plot": "Subgroups require interaction evidence before heterogeneity claims",
            "love_plot": "Weighting improves measured covariate balance",
            "weighted_risk_curve": "Weighted risk trajectories remain synthetic associations",
            "distribution": "Weight distribution is checked for instability",
            "flow_diagram": "The workflow separates design stages explicitly",
            "timeline": "The planned temporal sequence is explicit",
            "dag": "The causal diagram defines planned adjustment logic",
            "network": "Communication links remain computational hypotheses",
            "volcano": "Differential signals remain computational evidence",
            "enrichment": "Pathway signals are hypothesis-generating",
            "event_rate_plot": "Group composition is shown with denominator-aware labels",
            "risk_matrix": "Risk domains remain visible in the interpretation",
            "grade_summary": "Evidence certainty is summarized without upgrading conclusions",
            "formula": "Sample-size assumptions are explicit and editable",
            "funnel_plot": "Small-study patterns require cautious interpretation",
            "umap": "Embedding structure is descriptive, not mechanistic proof",
            "source_figure": "The registered source figure is preserved without embellishment",
        }
        slides.append({
            "slide_id": visual["slide_id"],
            "title": title_map.get(visual["visual_type"], visual["visual_type"].replace("_", " ").title()),
            "role": visual["semantic_role"],
            "visual": visual,
            "artifact": artifacts.get(visual["visual_id"], {}),
            "sources": [f"Source {sid[-6:]}" for sid in visual["source_ids"]],
            "notes_sources": visual["source_ids"],
            "claim_ids": [stable_id("CLM", visual["visual_id"], length=16)],
            "wording_boundary": visual["wording_boundary"],
            "protocol_without_results": protocol,
        })
    slides.append({
        "slide_id": stable_id("SLD", project_dir.name, "boundary", length=20),
        "title": "Evidence boundaries remain explicit",
        "role": "conclusion",
        "cards": [
            "All materials and findings are synthetic.",
            "Every displayed result traces to a registered source.",
            "Protocol-only projects do not emit results.",
        ],
        "sources": [],
    })
    return {
        "project_key": project_dir.name,
        "project_title": str(project_title),
        "arm": arm,
        "protocol_without_results": protocol,
        "slides": slides,
        "theme": {"accent": "167C80", "navy": "17324D", "teal": "3AAFA9"},
    }


def run(args: argparse.Namespace) -> int:
    suite = args.suite.resolve()
    run_id = args.run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    benchmark_root = (REPO / "benchmark" / "v2_2" / "runs" / run_id).resolve()
    staging_root = (REPO / "staging" / "v2_2" / run_id).resolve()
    projects = [path for path in sorted(suite.iterdir()) if path.is_dir() and PROJECT_PATTERN.match(path.name)]
    if len(projects) != 4:
        raise RuntimeError(f"Expected four isolated projects, found {len(projects)}")

    index_rows: list[dict[str, Any]] = []
    initial_qa: list[dict[str, Any]] = []
    for project_dir in projects:
        project_stage = staging_root / project_dir.name
        manifest, mapping = build_manifest(project_dir, project_stage)
        style_files = [row for row in manifest if row["style_reference"]]
        (project_stage / "style_status.json").write_text(
            json.dumps({
                "status": "NO_REFERENCE_STYLE" if not style_files else "REFERENCE_STYLE_PRESENT",
                "pptx_or_potx_count": len(style_files),
                "cache_used": False,
            }, indent=2),
            encoding="utf-8",
        )
        base_visuals, protocol = build_visuals(project_dir, mapping, optimized=False)
        optimized_visuals, _ = build_visuals(project_dir, mapping, optimized=True)
        for arm, config in ARMS.items():
            arm_start = time.perf_counter()
            arm_stage = project_stage / f"arm_{arm}"
            arm_output = benchmark_root / project_dir.name / f"arm_{arm}"
            if arm_output.exists():
                raise RuntimeError(f"Refusing to overwrite {arm_output}")
            arm_stage.mkdir(parents=True, exist_ok=True)
            arm_output.mkdir(parents=True, exist_ok=False)
            visuals = json.loads(json.dumps(optimized_visuals if config["visual_ir"] else base_visuals))
            assign_backends(visuals, config["adapters"])
            artifacts: dict[str, dict[str, Any]] = {}
            semantic_rows = []
            for visual in visuals:
                schema_errors = validate_visual_ir_schema(visual)
                issues = validate_visual_semantics(visual, protocol_without_results=protocol)
                semantic_rows.append({
                    "visual_id": visual["visual_id"],
                    "visual_type": visual["visual_type"],
                    "schema_errors": len(schema_errors),
                    "critical_issues": len(issues),
                    "issue_codes": "|".join(item["code"] for item in issues),
                })
                artifact = render_visual(visual, arm_stage / "artifacts", backend=visual["preferred_backend"])
                artifacts[visual["visual_id"]] = asdict(artifact)
            (arm_output / "visual_ir.json").write_text(json.dumps(visuals, ensure_ascii=False, indent=2), encoding="utf-8")
            write_csv(arm_output / "visual_semantic_validation.csv", semantic_rows)
            claims = make_claims(visuals, protocol)
            write_csv(arm_output / "claim_source_map.csv", claims)
            deck_spec = build_deck_spec(project_dir, arm, visuals, artifacts, protocol)
            spec_path = arm_stage / "deck_spec.json"
            spec_path.write_text(json.dumps(deck_spec, ensure_ascii=False, indent=2), encoding="utf-8")
            pptx_path = arm_output / f"{project_dir.name}_arm_{arm}.pptx"
            completed = subprocess.run(
                [args.node, str(REPO / "extensions" / "visual_backends" / "render_benchmark_deck.mjs"), str(spec_path), str(pptx_path)],
                cwd=REPO, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
            )
            (arm_output / "generation.log").write_text(
                f"exit={completed.returncode}\nstdout:\n{completed.stdout}\nstderr:\n{completed.stderr}",
                encoding="utf-8",
            )
            if completed.returncode != 0 or not pptx_path.is_file():
                raise RuntimeError(f"Deck generation failed: {project_dir.name} arm {arm}: {completed.stderr}")
            elapsed = time.perf_counter() - arm_start
            deck_ir = {
                "deck_id": stable_id("DECK", project_dir.name, arm, length=16),
                "arm": arm,
                "project_key": project_dir.name,
                "visual_ir_enabled": config["visual_ir"],
                "external_adapters_enabled": config["adapters"],
                "slides": [{"slide_id": slide["slide_id"], "revision": 1} for slide in deck_spec["slides"]],
            }
            (arm_output / "deck_ir.json").write_text(json.dumps(deck_ir, indent=2), encoding="utf-8")
            index_rows.append({
                "project_key": project_dir.name,
                "arm": arm,
                "pptx_path": str(pptx_path),
                "output_dir": str(arm_output),
                "slide_count": len(deck_spec["slides"]),
                "visual_count": len(visuals),
                "semantic_critical_issues": sum(row["critical_issues"] for row in semantic_rows),
                "fallback_count": sum(1 for artifact in artifacts.values() if artifact["fallback_used"]),
                "generation_seconds": round(elapsed, 3),
                "protocol_without_results": protocol,
            })
            initial_qa.extend({
                "project_key": project_dir.name,
                "arm": arm,
                **row,
            } for row in semantic_rows)
    write_csv(benchmark_root / "deck_index.csv", index_rows)
    write_csv(benchmark_root / "initial_semantic_qa.csv", initial_qa)
    (benchmark_root / "generation_complete.json").write_text(
        json.dumps({
            "run_id": run_id,
            "completed_at_utc": datetime.now(timezone.utc).isoformat(),
            "ground_truth_content_read": False,
            "deck_count": len(index_rows),
        }, indent=2),
        encoding="utf-8",
    )
    print(f"RUN_ID={run_id}")
    print(f"DECK_COUNT={len(index_rows)}")
    print(f"BENCHMARK_ROOT={benchmark_root}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--suite", type=Path, required=True)
    parser.add_argument("--run-id")
    parser.add_argument("--node", required=True)
    return run(parser.parse_args())


if __name__ == "__main__":
    raise SystemExit(main())
