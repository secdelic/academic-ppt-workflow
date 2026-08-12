from __future__ import annotations

import re
from collections.abc import Mapping
from typing import Any

from .chart_data import validate_chart_contract


CONTRACT_VISUALS = {
    "categorical_event_rates": "event_rate_chart",
    "effect_estimates": "forest_plot",
    "longitudinal_measurements": "line_chart",
    "sensitivity_analyses": "forest_plot",
    "subgroup_results": "subgroup_forest_plot",
    "missingness": "missingness_chart",
}


def route_content_type(
    content_type: str,
    payload: Mapping[str, Any],
) -> tuple[str, list[str]]:
    """Route a structured payload to a visual or a safe structured fallback."""

    content_type = str(content_type).strip()
    required_by_type = {
        "effect_estimates": ["categories", "values", "confidence_interval", "reference_value"],
        "sensitivity_analyses": ["categories", "values", "confidence_interval", "reference_value"],
        "subgroup_results": ["categories", "values", "confidence_interval", "reference_value"],
        "categorical_event_rates": ["categories", "values", "denominator"],
        "longitudinal_measurements": ["categories", "values"],
        "missingness": ["categories", "values"],
        "temporal_study_design": ["milestones"],
        "two_binary_dimensions": ["row_dimension", "column_dimension", "cells"],
        "cohort_inclusion_exclusion": ["nodes"],
    }
    missing = []
    for field in required_by_type.get(content_type, []):
        value = payload.get(field)
        if value is None or value == "" or value == []:
            missing.append(field)
    if missing:
        return "structured_table", missing
    if content_type == "temporal_study_design":
        return "timeline", []
    if content_type == "two_binary_dimensions":
        return "matrix_2x2", []
    if content_type == "cohort_inclusion_exclusion":
        return "native_flow", []
    return CONTRACT_VISUALS.get(content_type, "structured_table"), []


def _message_for_claims(claims: list[dict[str, str]]) -> str:
    if not claims:
        return "INFORMATION_REQUIRED"
    candidates = [
        str(claim.get("claim_text", "")).strip()
        for claim in claims
        if str(claim.get("claim_text", "")).strip()
    ]
    if not candidates:
        return "INFORMATION_REQUIRED"
    if len(candidates) == 1:
        return candidates[0]
    compact = "；".join(candidates[:2])
    # A chart already carries the complete structured values. Keep the spoken
    # takeaway concise enough for the reserved message band and retain every
    # linked claim in notes/source maps rather than forcing two conclusions
    # into the visible text box.
    return compact if len(compact) <= 110 else candidates[0]


def _title_for_contract(contract: Mapping[str, Any]) -> str:
    return {
        "categorical_event_rates": "组间事件率差异由 canonical 汇总数据直接呈现",
        "effect_estimates": "调整后效应估计同时展示方向与不确定性",
        "sensitivity_analyses": "敏感性分析的效应方向与区间集中展示",
        "subgroup_results": "亚组估计必须结合置信区间和交互检验解释",
        "missingness": "缺失比例决定变量能否进入主要模型",
        "longitudinal_measurements": "纵向测量以时间序列展示变化模式",
    }.get(str(contract.get("content_type")), "结构化结果采用可比较的视觉形式")


def chart_caption(contract: Mapping[str, Any]) -> str:
    units = str(contract.get("units", "source unit"))
    denominators = list(contract.get("denominator", []))
    numerators = list(contract.get("numerator", []))
    total = contract.get("sample_size_total")
    groups = dict(contract.get("sample_size_by_group", {}))
    if denominators:
        group_text = "；".join(
            f"{label}: {int(value) if float(value).is_integer() else value}"
            for label, value in groups.items()
        )
        total_text = (
            f"N={int(total) if float(total).is_integer() else total}"
            if total is not None
            else f"各组样本量：{group_text}"
        )
        event_text = ""
        if numerators and all(value is not None for value in numerators):
            event_text = "；事件数/样本量见图中标签"
        return f"{total_text}；单位：{units}{event_text}"
    if contract.get("confidence_interval"):
        sample = (
            f"N={total}"
            if total is not None
            else "sample size not reported in source"
        )
        reference = contract.get("reference_value")
        return (
            f"{contract.get('estimate_type') or 'effect estimate'}；"
            f"参考线={reference if reference is not None else 'source-defined'}；{sample}"
        )
    if str(contract.get("content_type")) == "missingness":
        return "缺失例数与缺失比例来自同一 canonical ChartDataContract"
    return f"单位：{units}"


def route_chart_contracts(
    claims: list[dict[str, str]],
    contracts: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], set[str]]:
    slides: list[dict[str, Any]] = []
    consumed: set[str] = set()
    claims_by_source: dict[str, list[dict[str, str]]] = {}
    for claim in claims:
        claims_by_source.setdefault(claim["source_id"], []).append(claim)
    for contract in contracts:
        if validate_chart_contract(contract):
            continue
        visual_type, missing = route_content_type(
            str(contract.get("content_type", "")), contract
        )
        if missing or visual_type == "structured_table":
            continue
        related = claims_by_source.get(str(contract["source_id"]), [])
        if not related:
            continue
        claim_ids = [claim["claim_id"] for claim in related]
        consumed.update(claim_ids)
        slides.append(
            {
                "logical_slide_key": f"chart:{contract['chart_id']}",
                "slide_role": "result",
                "section_id": "results",
                "slide_title": _title_for_contract(contract),
                "slide_purpose": "Compare source-bound structured results",
                "single_key_message": _message_for_claims(related),
                "source_ids": [contract["source_id"]],
                "claim_ids": claim_ids,
                "proposed_layout": visual_type,
                "layout_family": visual_type,
                "visual_type": visual_type,
                "chart_data": dict(contract),
                "chart_caption": chart_caption(contract),
                "citation_requirement": "required",
                "speaker_note_summary": (
                    "The chart, caption, sample-size metadata, and labels use "
                    f"ChartDataContract {contract['chart_id']}."
                ),
                "confidence": min(
                    (claim.get("confidence", "medium") for claim in related),
                    default="medium",
                ),
                "manual_review_required": "yes",
                "evidence_strength": (
                    "observational_association"
                    if contract["content_type"]
                    in {
                        "effect_estimates",
                        "sensitivity_analyses",
                        "subgroup_results",
                    }
                    else "reported_fact"
                ),
                "prohibited_overstatement": "Do not exceed the wording boundary of any linked claim.",
            }
        )
    return slides, consumed


def route_unstructured_claim(
    claim: Mapping[str, str],
) -> dict[str, Any]:
    text = str(claim.get("claim_text", ""))
    requested = str(claim.get("visual_type", "")).strip()
    if requested == "source_figure" and claim.get("visual_asset_path"):
        visual = "source_figure"
        extra = {"visual_asset_path": claim.get("visual_asset_path", "")}
    elif re.search(r"(?i)\b(?:landmark|baseline|follow[- ]?up|timepoint)\b|小时|随访|时间点", text):
        times = re.findall(r"(?i)(\d+(?:\.\d+)?)\s*(?:h|hr|hour|小时|天|day)", text)
        visual = "timeline" if len(times) >= 2 else "takeaway"
        extra = {
            "timeline_spec": {
                "milestones": [
                    {"label": f"时间点 {index + 1}", "value": value}
                    for index, value in enumerate(times[:5])
                ]
            }
        }
    elif re.search(r"(?i)conflict|canonical|unresolved|information_required|冲突|待确认", text):
        visual = "audit_table"
        extra = {
            "audit_rows": [
                {
                    "item": str(claim.get("claim_id", "")),
                    "decision": text,
                    "status": str(claim.get("conflict_status", "manual review")),
                }
            ]
        }
    else:
        visual = "takeaway"
        extra = {
            "takeaways": [
                text,
                str(
                    claim.get(
                        "wording_boundary",
                        claim.get(
                            "prohibited_overstatement",
                            "Retain source evidence strength.",
                        ),
                    )
                ),
            ]
        }
    return {
        "visual_type": visual,
        "layout_family": visual,
        **extra,
    }


__all__ = [
    "CONTRACT_VISUALS",
    "chart_caption",
    "route_chart_contracts",
    "route_content_type",
    "route_unstructured_claim",
]
