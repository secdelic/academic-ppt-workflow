"""Declarative, evidence-aware scientific visual specifications.

This module does not execute external renderers and does not modify PPTX files.
It produces deterministic geometry that can be consumed by PowerPoint native
shape code or rendered as a self-contained, local SVG.
"""

from __future__ import annotations

import copy
import hashlib
import json
import shutil
from html import escape
from pathlib import Path
from typing import Any, Callable, Mapping


TEMPLATE_IDS = (
    "cohort_flow",
    "consort_flow",
    "strobe_flow",
    "dag",
    "target_trial_framework",
    "bioinformatics_workflow",
    "single_cell_workflow",
    "multi_omics_integration",
    "mechanism_hypothesis",
    "timeline",
    "study_design_schematic",
)

MECHANISM_EVIDENCE_STATUSES = (
    "demonstrated_mechanism",
    "supported_interpretation",
    "proposed_mechanism",
    "hypothesis_only",
)

OPTIONAL_RENDERER_EXECUTABLES = {
    "graphviz": "dot",
    "mermaid": "mmdc",
}

DEFAULT_RENDER_TIMEOUT_SECONDS = 30


def _node(
    node_id: str,
    label: str,
    x: float,
    y: float,
    *,
    w: float = 0.20,
    h: float = 0.11,
    shape: str = "roundRect",
    role: str = "process",
) -> dict[str, Any]:
    return {
        "node_id": node_id,
        "label": label,
        "role": role,
        "shape": shape,
        "x": x,
        "y": y,
        "w": w,
        "h": h,
    }


def _edge(
    source: str,
    target: str,
    label: str = "",
    *,
    relationship: str = "directed",
) -> dict[str, str]:
    return {
        "source": source,
        "target": target,
        "label": label,
        "relationship": relationship,
    }


def _template_definitions() -> dict[str, dict[str, Any]]:
    return {
        "cohort_flow": {
            "layout_kind": "vertical_flow",
            "nodes": [
                _node("source_population", "Source population", 0.40, 0.08),
                _node("eligible", "Eligible records", 0.40, 0.27),
                _node("excluded", "Excluded records", 0.72, 0.27, role="exclusion"),
                _node("included", "Included cohort", 0.40, 0.48),
                _node("analysis", "Analysis population", 0.40, 0.70, role="outcome"),
            ],
            "edges": [
                _edge("source_population", "eligible"),
                _edge("eligible", "excluded", "Excluded"),
                _edge("eligible", "included"),
                _edge("included", "analysis"),
            ],
        },
        "consort_flow": {
            "layout_kind": "branched_flow",
            "nodes": [
                _node("assessed", "Assessed for eligibility", 0.40, 0.05),
                _node("excluded", "Excluded", 0.72, 0.20, role="exclusion"),
                _node("randomized", "Randomized", 0.40, 0.24),
                _node("arm_a", "Allocated to arm A", 0.17, 0.45),
                _node("arm_b", "Allocated to arm B", 0.63, 0.45),
                _node("analyzed_a", "Analyzed: arm A", 0.17, 0.70, role="outcome"),
                _node("analyzed_b", "Analyzed: arm B", 0.63, 0.70, role="outcome"),
            ],
            "edges": [
                _edge("assessed", "excluded", "Excluded"),
                _edge("assessed", "randomized"),
                _edge("randomized", "arm_a"),
                _edge("randomized", "arm_b"),
                _edge("arm_a", "analyzed_a"),
                _edge("arm_b", "analyzed_b"),
            ],
        },
        "strobe_flow": {
            "layout_kind": "vertical_flow",
            "nodes": [
                _node("setting", "Study setting", 0.40, 0.06),
                _node("participants", "Potential participants", 0.40, 0.25),
                _node("excluded", "Excluded or missing", 0.72, 0.25, role="exclusion"),
                _node("observed", "Observed cohort", 0.40, 0.47),
                _node("analysis", "Final analytic sample", 0.40, 0.70, role="outcome"),
            ],
            "edges": [
                _edge("setting", "participants"),
                _edge("participants", "excluded", "Not included"),
                _edge("participants", "observed"),
                _edge("observed", "analysis"),
            ],
        },
        "dag": {
            "layout_kind": "causal_graph",
            "nodes": [
                _node("exposure", "Exposure", 0.10, 0.38, shape="ellipse", role="exposure"),
                _node("confounder", "Confounder", 0.39, 0.08, shape="diamond", role="covariate"),
                _node("mediator", "Mediator", 0.39, 0.38, shape="ellipse", role="mediator"),
                _node("outcome", "Outcome", 0.70, 0.38, shape="ellipse", role="outcome"),
            ],
            "edges": [
                _edge("confounder", "exposure"),
                _edge("confounder", "outcome"),
                _edge("exposure", "mediator"),
                _edge("mediator", "outcome"),
                _edge("exposure", "outcome"),
            ],
        },
        "target_trial_framework": {
            "layout_kind": "horizontal_framework",
            "nodes": [
                _node("eligibility", "Eligibility", 0.03, 0.36, w=0.16),
                _node("strategies", "Treatment strategies", 0.22, 0.36, w=0.16),
                _node("assignment", "Assignment", 0.41, 0.36, w=0.16),
                _node("follow_up", "Follow-up", 0.60, 0.36, w=0.16),
                _node("estimand", "Outcome and estimand", 0.79, 0.36, w=0.18, role="outcome"),
            ],
            "edges": [
                _edge("eligibility", "strategies"),
                _edge("strategies", "assignment"),
                _edge("assignment", "follow_up"),
                _edge("follow_up", "estimand"),
            ],
        },
        "bioinformatics_workflow": {
            "layout_kind": "horizontal_pipeline",
            "nodes": [
                _node("data", "Input data", 0.03, 0.36, w=0.15),
                _node("qc", "Quality control", 0.21, 0.36, w=0.15),
                _node("preprocess", "Preprocessing", 0.39, 0.36, w=0.15),
                _node("analysis", "Computational analysis", 0.57, 0.36, w=0.17),
                _node("validation", "Validation plan", 0.78, 0.36, w=0.17, role="outcome"),
            ],
            "edges": [
                _edge("data", "qc"),
                _edge("qc", "preprocess"),
                _edge("preprocess", "analysis"),
                _edge("analysis", "validation"),
            ],
        },
        "single_cell_workflow": {
            "layout_kind": "horizontal_pipeline",
            "nodes": [
                _node("sample", "Sample", 0.02, 0.36, w=0.12),
                _node("dissociation", "Dissociation", 0.16, 0.36, w=0.13),
                _node("library", "Library preparation", 0.31, 0.36, w=0.15),
                _node("qc", "Cell-level QC", 0.48, 0.36, w=0.13),
                _node("clustering", "Clustering", 0.63, 0.36, w=0.13),
                _node("annotation", "Annotation", 0.78, 0.36, w=0.13),
                _node("downstream", "Downstream inference", 0.78, 0.62, w=0.17, role="outcome"),
            ],
            "edges": [
                _edge("sample", "dissociation"),
                _edge("dissociation", "library"),
                _edge("library", "qc"),
                _edge("qc", "clustering"),
                _edge("clustering", "annotation"),
                _edge("annotation", "downstream"),
            ],
        },
        "multi_omics_integration": {
            "layout_kind": "converging_flow",
            "nodes": [
                _node("genomics", "Genomics", 0.05, 0.10, w=0.16),
                _node("transcriptomics", "Transcriptomics", 0.05, 0.29, w=0.16),
                _node("proteomics", "Proteomics", 0.05, 0.48, w=0.16),
                _node("metabolomics", "Metabolomics", 0.05, 0.67, w=0.16),
                _node("integration", "Integration model", 0.40, 0.38, w=0.20, shape="hexagon"),
                _node("interpretation", "Integrated interpretation", 0.73, 0.38, w=0.21, role="outcome"),
            ],
            "edges": [
                _edge("genomics", "integration"),
                _edge("transcriptomics", "integration"),
                _edge("proteomics", "integration"),
                _edge("metabolomics", "integration"),
                _edge("integration", "interpretation"),
            ],
        },
        "mechanism_hypothesis": {
            "layout_kind": "evidence_labeled_chain",
            "nodes": [
                _node("initiator", "Initiating factor", 0.08, 0.36, w=0.20, role="exposure"),
                _node("intermediate", "Proposed intermediate", 0.39, 0.36, w=0.22, role="mediator"),
                _node("response", "Observed response", 0.72, 0.36, w=0.20, role="outcome"),
            ],
            "edges": [
                _edge("initiator", "intermediate", "Interpreted link"),
                _edge("intermediate", "response", "Interpreted link"),
            ],
        },
        "timeline": {
            "layout_kind": "timeline",
            "nodes": [
                _node("baseline", "Baseline", 0.06, 0.38, w=0.16, shape="ellipse"),
                _node("timepoint_1", "Timepoint 1", 0.29, 0.38, w=0.16, shape="ellipse"),
                _node("timepoint_2", "Timepoint 2", 0.52, 0.38, w=0.16, shape="ellipse"),
                _node("endpoint", "Endpoint", 0.75, 0.38, w=0.16, shape="ellipse", role="outcome"),
            ],
            "edges": [
                _edge("baseline", "timepoint_1"),
                _edge("timepoint_1", "timepoint_2"),
                _edge("timepoint_2", "endpoint"),
            ],
        },
        "study_design_schematic": {
            "layout_kind": "horizontal_framework",
            "nodes": [
                _node("question", "Research question", 0.03, 0.36, w=0.16),
                _node("design", "Study design", 0.22, 0.36, w=0.16),
                _node("population", "Study population", 0.41, 0.36, w=0.16),
                _node("analysis", "Analysis plan", 0.60, 0.36, w=0.16),
                _node("interpretation", "Evidence-bounded interpretation", 0.79, 0.36, w=0.19, role="outcome"),
            ],
            "edges": [
                _edge("question", "design"),
                _edge("design", "population"),
                _edge("population", "analysis"),
                _edge("analysis", "interpretation"),
            ],
        },
    }


def _canonical_hash(value: Mapping[str, Any]) -> str:
    payload = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def available_visual_templates() -> tuple[str, ...]:
    """Return the stable template identifiers."""

    return TEMPLATE_IDS


def build_visual_spec(
    template_id: str,
    *,
    title: str | None = None,
    labels: Mapping[str, str] | None = None,
    evidence_status: str | None = None,
) -> dict[str, Any]:
    """Build a deterministic, renderer-neutral scientific visual specification."""

    definitions = _template_definitions()
    if template_id not in definitions:
        raise ValueError(f"Unsupported scientific visual template: {template_id}")
    if evidence_status is not None and template_id != "mechanism_hypothesis":
        raise ValueError("evidence_status is only valid for mechanism_hypothesis")
    if template_id == "mechanism_hypothesis":
        evidence_status = evidence_status or "hypothesis_only"
        if evidence_status not in MECHANISM_EVIDENCE_STATUSES:
            raise ValueError(
                "Mechanism evidence_status must be one of: "
                + ", ".join(MECHANISM_EVIDENCE_STATUSES)
            )

    definition = copy.deepcopy(definitions[template_id])
    node_ids = {node["node_id"] for node in definition["nodes"]}
    for node_id, label in (labels or {}).items():
        if node_id not in node_ids:
            raise ValueError(f"Unknown node label override: {node_id}")
        if not isinstance(label, str):
            raise TypeError(f"Label for {node_id} must be text")
        for node in definition["nodes"]:
            if node["node_id"] == node_id:
                node["label"] = label
                break

    spec: dict[str, Any] = {
        "template_id": template_id,
        "title": title or template_id.replace("_", " ").title(),
        "nodes": definition["nodes"],
        "edges": definition["edges"],
        "layout": {
            "kind": definition["layout_kind"],
            "coordinate_system": "normalized_0_to_1",
            "canvas_aspect_ratio": "16:9",
            "safe_margin": 0.03,
            "reading_order": [node["node_id"] for node in definition["nodes"]],
        },
        "metadata": {
            "native_shape_compatible": True,
            "safe_local_svg_compatible": True,
            "external_assets_allowed": False,
            "editable_object_requirement": "powerpoint_native_shapes_preferred",
            "evidence_status": evidence_status,
        },
    }
    spec["content_hash"] = _canonical_hash(spec)
    return spec


def _node_center(node: Mapping[str, Any], width: int, height: int) -> tuple[float, float]:
    return (
        (float(node["x"]) + float(node["w"]) / 2.0) * width,
        (float(node["y"]) + float(node["h"]) / 2.0) * height,
    )


def render_safe_svg(
    spec: Mapping[str, Any], *, width: int = 1200, height: int = 675
) -> str:
    """Render a self-contained SVG without images, scripts, links, or URLs."""

    if width <= 0 or height <= 0:
        raise ValueError("SVG dimensions must be positive")
    nodes = list(spec.get("nodes", []))
    edges = list(spec.get("edges", []))
    node_map = {str(node["node_id"]): node for node in nodes}
    if len(node_map) != len(nodes):
        raise ValueError("Duplicate node_id in visual specification")

    parts = [
        f'<svg viewBox="0 0 {width} {height}" width="{width}" height="{height}" '
        'role="img">',
        f'<title>{escape(str(spec.get("title", "")), quote=True)}</title>',
        '<rect x="0" y="0" width="100%" height="100%" fill="#FFFFFF"/>',
    ]
    for edge in edges:
        source = node_map.get(str(edge["source"]))
        target = node_map.get(str(edge["target"]))
        if source is None or target is None:
            raise ValueError("Edge references an unknown node")
        x1, y1 = _node_center(source, width, height)
        x2, y2 = _node_center(target, width, height)
        parts.append(
            f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" '
            'stroke="#52606D" stroke-width="3"/>'
        )
        if edge.get("label"):
            parts.append(
                f'<text x="{(x1 + x2) / 2:.2f}" y="{(y1 + y2) / 2 - 8:.2f}" '
                'text-anchor="middle" font-family="Arial" font-size="16" '
                f'fill="#334E68">{escape(str(edge["label"]), quote=True)}</text>'
            )

    for node in nodes:
        x = float(node["x"]) * width
        y = float(node["y"]) * height
        w = float(node["w"]) * width
        h = float(node["h"]) * height
        shape = str(node.get("shape", "roundRect"))
        fill = "#E8F1F8" if node.get("role") != "exclusion" else "#FDECEC"
        if shape == "ellipse":
            parts.append(
                f'<ellipse cx="{x + w / 2:.2f}" cy="{y + h / 2:.2f}" '
                f'rx="{w / 2:.2f}" ry="{h / 2:.2f}" fill="{fill}" '
                'stroke="#1F4E79" stroke-width="2"/>'
            )
        elif shape == "diamond":
            points = (
                f"{x + w / 2:.2f},{y:.2f} {x + w:.2f},{y + h / 2:.2f} "
                f"{x + w / 2:.2f},{y + h:.2f} {x:.2f},{y + h / 2:.2f}"
            )
            parts.append(
                f'<polygon points="{points}" fill="{fill}" '
                'stroke="#1F4E79" stroke-width="2"/>'
            )
        elif shape == "hexagon":
            inset = w * 0.18
            points = (
                f"{x + inset:.2f},{y:.2f} {x + w - inset:.2f},{y:.2f} "
                f"{x + w:.2f},{y + h / 2:.2f} {x + w - inset:.2f},{y + h:.2f} "
                f"{x + inset:.2f},{y + h:.2f} {x:.2f},{y + h / 2:.2f}"
            )
            parts.append(
                f'<polygon points="{points}" fill="{fill}" '
                'stroke="#1F4E79" stroke-width="2"/>'
            )
        else:
            radius = 12 if shape == "roundRect" else 0
            parts.append(
                f'<rect x="{x:.2f}" y="{y:.2f}" width="{w:.2f}" height="{h:.2f}" '
                f'rx="{radius}" fill="{fill}" stroke="#1F4E79" stroke-width="2"/>'
            )
        parts.append(
            f'<text x="{x + w / 2:.2f}" y="{y + h / 2 + 6:.2f}" '
            'text-anchor="middle" font-family="Arial" font-size="18" '
            f'fill="#102A43">{escape(str(node["label"]), quote=True)}</text>'
        )

    evidence_status = spec.get("metadata", {}).get("evidence_status")
    if evidence_status:
        parts.append(
            f'<text x="{width - 20}" y="{height - 16}" text-anchor="end" '
            'font-family="Arial" font-size="14" fill="#52606D">'
            f'Evidence status: {escape(str(evidence_status), quote=True)}</text>'
        )
    parts.append("</svg>")
    return "".join(parts)


def probe_visual_capabilities(
    executable_finder: Callable[[str], str | None] | None = None,
) -> dict[str, dict[str, Any]]:
    """Probe allowlisted optional renderers without executing or installing them."""

    finder = executable_finder or shutil.which
    result: dict[str, dict[str, Any]] = {
        "powerpoint_shapes": {
            "available": True,
            "execution_attempted": False,
            "executable": None,
            "timeout_seconds": 0,
        },
        "svg": {
            "available": True,
            "execution_attempted": False,
            "executable": None,
            "timeout_seconds": 0,
        },
    }
    for backend, executable in OPTIONAL_RENDERER_EXECUTABLES.items():
        resolved = finder(executable)
        result[backend] = {
            "available": bool(resolved),
            "execution_attempted": False,
            "executable": executable,
            "resolved_path": str(Path(resolved)) if resolved else None,
            "timeout_seconds": DEFAULT_RENDER_TIMEOUT_SECONDS,
        }
    return result


def resolve_optional_renderer(
    backend: str,
    executable_finder: Callable[[str], str | None] | None = None,
) -> str | None:
    """Resolve only an allowlisted renderer name; this never executes it."""

    if backend not in OPTIONAL_RENDERER_EXECUTABLES:
        raise ValueError(f"Renderer is not allowlisted: {backend}")
    finder = executable_finder or shutil.which
    return finder(OPTIONAL_RENDERER_EXECUTABLES[backend])
