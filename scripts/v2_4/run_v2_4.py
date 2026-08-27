from __future__ import annotations

import argparse
import csv
import json
import math
import os
import statistics
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from model import (
    asset_binding,
    detect_conflicts,
    detect_unresolved,
    discover_project,
    document_binding,
    dump_json,
    infer_kind,
    merge_bindings,
    source_binding,
    stable,
    validate_slide_semantics,
    validate_slide_spec,
    write_csv,
)


REPO = Path(__file__).resolve().parents[2]
PROJECT_PATTERN = "[0-9][0-9]_*"

PROFILE_BY_KIND = {
    "target_trial": "clinical_methods",
    "single_cell": "bioinformatics",
    "meta_analysis": "evidence_synthesis",
    "protocol": "clinical_protocol",
}


# Audience-facing localization is deliberately separate from the canonical
# source rows.  Source values, field bindings and row keys remain byte-for-byte
# faithful to the registered inputs; only SlideSpec narrative strings use these
# labels.  The renderer applies the same principle to chart labels.
ZH_ENTITY_LABELS = {
    "Control": "对照",
    "SIMD": "SIMD",
    "T cell": "T细胞",
    "B cell": "B细胞",
    "NK cell": "NK细胞",
    "Monocyte": "单核细胞",
    "Macrophage": "巨噬细胞",
    "Dendritic": "树突状细胞",
    "Endothelial": "内皮细胞",
    "Fibroblast": "成纤维细胞",
    "Pericyte": "周细胞",
    "Cardiomyocyte": "心肌细胞",
    "Biventricular dysfunction phenotype": "双心室功能障碍表型",
    "AKI progression by 72 h": "72小时内AKI进展",
    "SOFA change at 48 h": "48小时SOFA变化",
    "28-day mortality": "28天死亡",
    "Age": "年龄",
    "SOFA score": "SOFA评分",
    "Lactate": "乳酸",
    "Norepinephrine-equivalent dose": "去甲肾上腺素等效剂量",
    "Mechanical ventilation": "机械通气",
    "TTE image quality": "TTE图像质量",
}


def zh_entity_label(value: Any) -> str:
    return ZH_ENTITY_LABELS.get(str(value), str(value))


def claim(
    claim_id: str,
    claim_text: str,
    dataset: dict[str, Any],
    source_location: str,
    expected_value: str,
    wording_boundary: str,
    fields_used: list[str],
) -> dict[str, Any]:
    binding = source_binding(dataset, fields_used, source_location)
    return {
        "claim_id": claim_id,
        "claim_text": claim_text,
        "expected_value": expected_value,
        "wording_boundary": wording_boundary,
        "canonical_status": "canonical",
        "source_bindings": [binding],
    }


def build_claims(
    kind: str,
    datasets: dict[str, dict[str, Any]],
    sources: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    if kind == "target_trial":
        result: list[dict[str, Any]] = []
        data = datasets["effects"]
        for index, row in enumerate(data["rows"], 2):
            estimate = float(row["estimate"])
            if row["claim_id"] in {"CLM-TTE-001", "CLM-TTE-002"}:
                value = f"{estimate * 100:.1f}%"
            elif row["claim_id"] == "CLM-TTE-003":
                value = f"{estimate * 100:.1f} percentage points"
            else:
                value = f"{estimate:.2f}"
            boundary = {
                "CLM-TTE-001": "在目标试验假设下的估计风险；不得解释为真实随机证据。",
                "CLM-TTE-002": "在目标试验假设下的估计风险；不得解释为真实随机证据。",
                "CLM-TTE-003": "不得转化为输血建议。",
                "CLM-TTE-004": "不得宣称来自随机试验。",
                "CLM-TTE-005": "仅为支持性分析。",
            }[row["claim_id"]]
            text = {
                "CLM-TTE-001": f"24小时内输血策略的28天死亡风险估计为{value}",
                "CLM-TTE-002": f"24小时内未输血策略的28天死亡风险估计为{value}",
                "CLM-TTE-003": f"风险差为{estimate * 100:.1f}个百分点",
                "CLM-TTE-004": f"风险比为{value}",
                "CLM-TTE-005": f"加权风险比为{value}",
            }[row["claim_id"]]
            result.append(claim(
                row["claim_id"], text, data, f"row {index}", value, boundary,
                ["claim_id", "estimand", "contrast", "estimate", "ci_low", "ci_high", "status"],
            ))
        return result

    if kind == "single_cell":
        qc = datasets["qc"]
        composition = datasets["composition"]
        deg = datasets["deg"]
        pathway = datasets["pathway"]
        communication = datasets["communication"]
        final = min(qc["rows"], key=lambda row: float(row["remaining_cells"]))
        effect_field = next(field for field in deg["columns"] if field.lower().startswith("log2_fold_change"))
        by_cell_type: dict[str, list[dict[str, Any]]] = {}
        for row in composition["rows"]:
            by_cell_type.setdefault(row["cell_type"], []).append(row)
        featured_cell_type, featured_rows = max(
            by_cell_type.items(),
            key=lambda item: max(float(row["percentage"]) for row in item[1]) - min(float(row["percentage"]) for row in item[1]),
        )
        high_group = max(featured_rows, key=lambda row: float(row["percentage"]))
        low_group = min(featured_rows, key=lambda row: float(row["percentage"]))
        delta = float(high_group["percentage"]) - float(low_group["percentage"])
        featured_deg = max(deg["rows"], key=lambda row: abs(float(row[effect_field])))
        negative_pathways = [row for row in pathway["rows"] if float(row["normalized_enrichment_score"]) < 0]
        # When both enrichment directions exist, register the strongest
        # negative program as the canonical direction-preservation claim; the
        # full visual still renders every positive and negative source row.
        featured_pathway = min(
            negative_pathways or pathway["rows"],
            key=lambda row: float(row["normalized_enrichment_score"]),
        )
        featured_interaction = max(communication["rows"], key=lambda row: float(row["communication_probability"]))
        pathway_direction = "上调" if float(featured_pathway["normalized_enrichment_score"]) > 0 else "下调"
        return [
            claim("CLM-SC-001", f"质控后保留{final['remaining_cells']}个合成细胞", qc, "final QC row", final["remaining_cells"], "描述性合成细胞计数。", ["stage", "remaining_cells", "removed_at_stage"]),
            claim("CLM-SC-002", f"{high_group['group']}组{featured_cell_type}比例较{low_group['group']}组高{delta:.2f}个百分点", composition, f"{featured_cell_type} rows", "derived from canonical table", "描述性组间比较，不代表因果关系。", ["group", "cell_type", "cell_count", "percentage"]),
            claim("CLM-SC-003", f"{featured_deg['cell_type']} pseudobulk中{featured_deg['gene']}呈最强绝对变化", deg, f"gene {featured_deg['gene']}", f"log2FC {featured_deg[effect_field]}; FDR {featured_deg['FDR']}", "计算关联，不代表实验验证。", ["gene", effect_field, "FDR", "cell_type"]),
            claim("CLM-SC-004", f"{featured_pathway['cell_type']}的{featured_pathway['pathway']}通路{pathway_direction}", pathway, f"pathway {featured_pathway['pathway']}", f"NES {featured_pathway['normalized_enrichment_score']}; FDR {featured_pathway['FDR']}", "富集结果不是机制证明。", ["pathway", "normalized_enrichment_score", "FDR", "cell_type", "evidence_level"]),
            claim("CLM-SC-005", f"{featured_interaction['ligand_receptor']}为最高通信概率的预测互作", communication, f"interaction {featured_interaction['ligand_receptor']}", f"probability {featured_interaction['communication_probability']}", "仅为预测性细胞通讯，缺少外部验证。", ["sender", "receiver", "ligand_receptor", "communication_probability", "adjusted_p"]),
        ]

    if kind == "meta_analysis":
        pooled = datasets["pooled"]
        grade = datasets["grade"]
        random_effect = next(row for row in pooled["rows"] if "Random-effects" in row["estimand"])
        prediction = next(row for row in pooled["rows"] if "Prediction interval" in row["estimand"])
        return [
            claim("CLM-META-001", f"随机效应合并OR为{random_effect['estimate']}（95%CI {random_effect['ci_low']}–{random_effect['ci_high']}）", pooled, "random-effects row", f"{random_effect['estimate']} ({random_effect['ci_low']}-{random_effect['ci_high']})", "仅表述观察性关联。", ["claim_id", "estimand", "estimate", "ci_low", "ci_high", "I2_percent", "study_count"]),
            claim("CLM-META-002", f"异质性I²为{random_effect['I2_percent']}%", pooled, "random-effects row", f"I2={random_effect['I2_percent']}%", "中等不一致性，不代表效应修饰。", ["estimand", "I2_percent"]),
            claim("CLM-META-003", f"预测区间为{prediction['ci_low']}–{prediction['ci_high']}", pooled, "prediction interval row", f"{prediction['ci_low']}-{prediction['ci_high']}", "预测区间不得标记为置信区间。", ["estimand", "ci_low", "ci_high"]),
            claim("CLM-META-004", f"共有{random_effect['study_count']}项研究进入合并", pooled, "random-effects row", random_effect["study_count"], "使用canonical研究数量。", ["estimand", "study_count"]),
            claim("CLM-META-005", f"主要结局GRADE证据确定性为{grade['rows'][0]['certainty']}", grade, "primary outcome row", grade["rows"][0]["certainty"], "不得提高证据确定性。", ["outcome", "studies", "participants", "risk_of_bias", "inconsistency", "indirectness", "imprecision", "other", "certainty"]),
        ]

    sample = datasets["sample_size"]
    schedule = datasets["schedule"]
    variables = datasets["variables"]
    target = next(row for row in sample["rows"] if row["parameter"] == "Target enrollment")
    analyzable = next(row for row in sample["rows"] if row["parameter"] == "Required analyzable sample")
    primary_endpoint = next(row for row in schedule["rows"] if row["status"] == "Primary endpoint")
    exposure = next(row for row in variables["rows"] if row["role"] == "Primary exposure")
    protocol_status_binding = document_binding(sources or [], "full_protocol", "protocol status") if sources else source_binding(sample, ["parameter", "value", "unit"], "protocol status")
    return [
        claim("CLM-PROTO-001", f"计划目标入组{target['value']}例", sample, "Target enrollment row", target["value"], "计划值，不得表述为已完成。", ["parameter", "value", "unit"]),
        claim("CLM-PROTO-002", f"所需可分析样本量为{analyzable['value']}例", sample, "Required analyzable sample row", analyzable["value"], "基于假设的样本量规划。", ["parameter", "value", "unit"]),
        claim("CLM-PROTO-003", f"主要终点时间为{primary_endpoint['window']}", schedule, "Primary endpoint row", primary_endpoint["window"], "研究方案定义。", ["visit", "assessment", "window", "status"]),
        claim("CLM-PROTO-004", f"主要暴露在{exposure['timepoint']}测量", variables, "Primary exposure row", exposure["timepoint"], "研究方案定义。", ["role", "variable", "timepoint", "type"]),
        {
            "claim_id": "CLM-PROTO-005",
            "claim_text": "本研究方案尚无任何观察结果",
            "expected_value": "none",
            "wording_boundary": "严禁生成任何研究结果。",
            "canonical_status": "canonical",
            "source_bindings": [protocol_status_binding],
        },
    ]


def make_visual(
    kind: str,
    slide_key: str,
    visual_type: str,
    scientific_role: str,
    bindings: list[dict[str, Any]],
    payload: dict[str, Any],
    expected_row_count: int,
    rendered_row_count: int | None = None,
    **fields: Any,
) -> dict[str, Any]:
    # A VisualSpec is planned before PowerPoint objects exist.  Keep the actual
    # rendered row count unknown until the renderer reports it independently;
    # this prevents a planner-side value from masquerading as render evidence.
    source_row_count = expected_row_count
    planned_count = expected_row_count
    if visual_type in {"umap_figure", "source_figure"}:
        planned_count = 1
    elif isinstance(payload.get("rows"), list) and (payload["rows"] or expected_row_count == 0):
        planned_count = len(payload["rows"])
    elif isinstance(payload.get("bins"), list):
        planned_count = len(payload["bins"])
    elif isinstance(payload.get("items"), list):
        planned_count = len(payload["items"])
    elif isinstance(payload.get("nodes"), list):
        planned_count = len(payload["nodes"])
    elif isinstance(payload.get("events"), list):
        planned_count = len(payload["events"])
    elif isinstance(payload.get("steps"), list):
        planned_count = len(payload["steps"])
    elif isinstance(payload.get("takeaways"), list):
        planned_count = len(payload["takeaways"])
    elif isinstance(payload.get("allowed"), list) or isinstance(payload.get("prohibited"), list):
        planned_count = len(payload.get("allowed", [])) + len(payload.get("prohibited", []))
    elif isinstance(payload.get("planned"), list) or isinstance(payload.get("not_available"), list):
        planned_count = len(payload.get("planned", [])) + len(payload.get("not_available", []))
    elif isinstance(payload.get("available"), list) or isinstance(payload.get("missing"), list):
        planned_count = len(payload.get("available", [])) + len(payload.get("missing", []))
    default_filter = "; ".join(
        f"{binding['source_id']}:{binding.get('row_filter', 'all rows')}" for binding in bindings
    )
    default_aggregation = "; ".join(
        f"{binding['source_id']}:{binding.get('aggregation', 'none')}" for binding in bindings
    )
    visual_id = stable("VIS", kind, slide_key, visual_type, n=16)
    row_items: list[Any]
    if visual_type in {"umap_figure", "source_figure"}:
        row_items = [{"image_path": payload.get("image_path", "")}]
    elif isinstance(payload.get("rows"), list):
        row_items = payload["rows"]
    elif isinstance(payload.get("bins"), list):
        row_items = payload["bins"]
    elif isinstance(payload.get("items"), list):
        row_items = payload["items"]
    elif isinstance(payload.get("nodes"), list):
        row_items = payload["nodes"]
    elif isinstance(payload.get("events"), list):
        row_items = payload["events"]
    elif isinstance(payload.get("steps"), list):
        row_items = payload["steps"]
    elif isinstance(payload.get("takeaways"), list):
        row_items = payload["takeaways"]
    elif isinstance(payload.get("allowed"), list) or isinstance(payload.get("prohibited"), list):
        row_items = payload.get("allowed", []) + payload.get("prohibited", [])
    elif isinstance(payload.get("planned"), list) or isinstance(payload.get("not_available"), list):
        row_items = payload.get("planned", []) + payload.get("not_available", [])
    elif isinstance(payload.get("available"), list) or isinstance(payload.get("missing"), list):
        row_items = payload.get("available", []) + payload.get("missing", [])
    else:
        row_items = [{"singleton_visual": visual_type}] if planned_count else []
    if len(row_items) != planned_count:
        raise ValueError(
            f"Visual row contract mismatch before rendering: {visual_type} has "
            f"{len(row_items)} row identities for {planned_count} expected rows"
        )
    expected_row_keys = [
        stable("ROW", visual_id, index, json.dumps(item, ensure_ascii=False, sort_keys=True), n=12)
        for index, item in enumerate(row_items)
    ]
    visual = {
        "visual_id": visual_id,
        "visual_type": visual_type,
        "scientific_role": scientific_role,
        "data_contract_id": "",
        "source_bindings": bindings,
        "row_filter": fields.pop("row_filter", default_filter),
        "facet_filter": fields.pop("facet_filter", "none"),
        "aggregation": fields.pop("aggregation", default_aggregation),
        "entity_scope": fields.pop("entity_scope", "all registered entities"),
        "outcome": fields.pop("outcome", ""),
        "estimand": fields.pop("estimand", ""),
        "unit": fields.pop("unit", ""),
        "x_field": fields.pop("x_field", ""),
        "y_field": fields.pop("y_field", ""),
        "group_field": fields.pop("group_field", ""),
        "facet_field": fields.pop("facet_field", ""),
        "color_field": fields.pop("color_field", ""),
        "label_field": fields.pop("label_field", ""),
        "ci_low_field": fields.pop("ci_low_field", ""),
        "ci_high_field": fields.pop("ci_high_field", ""),
        "reference_value": fields.pop("reference_value", None),
        "time_field": fields.pop("time_field", ""),
        "edge_source_field": fields.pop("edge_source_field", ""),
        "edge_target_field": fields.pop("edge_target_field", ""),
        "expected_row_count": planned_count,
        "source_row_count": source_row_count,
        "rendered_row_count": rendered_row_count,
        "expected_row_keys": expected_row_keys,
        "rendered_row_keys": [],
        "top_k_policy": fields.pop("top_k_policy", "none"),
        "omitted_rows": fields.pop("omitted_rows", []),
        "editability_requirement": fields.pop("editability_requirement", "native_shapes_or_chart"),
        "design_variant": fields.pop("design_variant", PROFILE_BY_KIND[kind]),
        "sequential": fields.pop("sequential", None),
        "payload": payload,
    }
    visual.update(fields)
    contract_payload = {
        "visual_type": visual_type,
        "scientific_role": scientific_role,
        "bindings": [
            {
                "source_id": binding["source_id"],
                "fields_used": binding.get("fields_used", []),
                "row_filter": binding.get("row_filter", ""),
                "aggregation": binding.get("aggregation", ""),
                "canonical_status": binding.get("canonical_status", ""),
            }
            for binding in bindings
        ],
        "row_filter": visual["row_filter"],
        "facet_filter": visual["facet_filter"],
        "aggregation": visual["aggregation"],
        "entity_scope": visual["entity_scope"],
        "outcome": visual["outcome"],
        "estimand": visual["estimand"],
        "unit": visual["unit"],
        "fields": {
            name: visual[name]
            for name in (
                "x_field", "y_field", "group_field", "facet_field", "color_field",
                "label_field", "ci_low_field", "ci_high_field", "time_field",
                "edge_source_field", "edge_target_field",
            )
        },
        "expected_row_keys": expected_row_keys,
    }
    visual["data_contract_id"] = stable(
        "DATA", json.dumps(contract_payload, ensure_ascii=False, sort_keys=True), n=16
    )
    return visual


def make_slide(
    kind: str,
    key: str,
    title: str,
    role: str,
    purpose: str,
    layout_family: str,
    visuals: list[dict[str, Any]],
    claims: list[dict[str, Any]],
    claim_ids: list[str] | None = None,
    conflicts: list[dict[str, Any]] | None = None,
    unresolved: list[dict[str, Any]] | None = None,
    coverage_items: list[str] | None = None,
    appendix: bool = False,
    card_grid: bool | None = None,
) -> dict[str, Any]:
    claim_ids = claim_ids or []
    conflicts = conflicts or []
    unresolved = unresolved or []
    claim_by_id = {item["claim_id"]: item for item in claims}
    claim_bindings = [claim_by_id[item] for item in claim_ids]
    registry_bindings = [
        {
            "source_id": item["source_id"],
            "source_file": item["source_file"],
            "source_location": item["source_location"],
            "fields_used": [],
            "row_filter": item.get("conflict_id", item.get("item_id", "registry item")),
            "aggregation": "conflict registry" if "conflict_id" in item else "unresolved registry",
            "canonical_status": "noncanonical_conflict" if "conflict_id" in item else "information_required",
            "short_label_zh": "冲突来源" if "conflict_id" in item else "待确认来源",
        }
        for item in [*conflicts, *unresolved]
    ]
    bindings = merge_bindings(
        [binding for visual in visuals for binding in visual["source_bindings"]]
        + [binding for item in claim_bindings for binding in item["source_bindings"]]
        + registry_bindings
    )
    return {
        "slide_id": stable("SLD", kind, key, n=20),
        "slide_role": role,
        "narrative_purpose": purpose,
        "title": title,
        "language": "zh-CN",
        "key_message": title,
        "claim_bindings": claim_bindings,
        "source_bindings": bindings,
        "conflict_bindings": conflicts,
        "unresolved_bindings": unresolved,
        "visual_specs": visuals,
        "layout_family": layout_family,
        "design_profile": PROFILE_BY_KIND[kind],
        "appendix_status": "appendix" if appendix else "body",
        "review_flags": [
            *("conflict_review" for _ in conflicts or []),
            *("information_required" for _ in unresolved or []),
        ],
        "coverage_items": coverage_items or [],
        "card_grid": bool(card_grid) if card_grid is not None else any(
            visual["visual_type"] in {"specification_matrix", "assumption_warning_matrix", "comparison_cards"}
            for visual in visuals
        ),
        "text_only": len(visuals) == 0 and not appendix,
    }


def histogram(values: list[float], bin_count: int = 18) -> list[dict[str, float]]:
    low, high = min(values), max(values)
    width = (high - low) / bin_count if high > low else 1.0
    counts = [0] * bin_count
    for value in values:
        index = min(bin_count - 1, int((value - low) / width))
        counts[index] += 1
    return [
        {"bin_low": low + i * width, "bin_high": low + (i + 1) * width, "count": count}
        for i, count in enumerate(counts)
    ]


def covered_items(brief: dict[str, Any], *keywords: str) -> list[str]:
    result = []
    for item in brief.get("must_include", []):
        normalized = item.lower()
        if any(keyword.lower() in normalized for keyword in keywords):
            result.append(item)
    return result


def narrative_visual(kind: str, key: str, visual_type: str, binding: dict[str, Any], payload: dict[str, Any], role: str = "context") -> dict[str, Any]:
    return make_visual(kind, key, visual_type, role, [binding], payload, 1, entity_scope="source-bound narrative")


def build_target_trial_slides(
    brief: dict[str, Any], sources: list[dict[str, Any]], datasets: dict[str, dict[str, Any]],
    claims: list[dict[str, Any]], conflicts: list[dict[str, Any]], unresolved: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    kind = "target_trial"
    protocol = document_binding(sources, "target_trial_protocol", "target-trial specification")
    memo = document_binding(sources, "analysis_memo", "analysis specification")
    effect_rows = datasets["effects"]["rows"]
    absolute_effect_rows = [row for row in effect_rows if row["claim_id"] in {"CLM-TTE-001", "CLM-TTE-002", "CLM-TTE-003"}]
    ratio_effect_rows = [row for row in effect_rows if row not in absolute_effect_rows]
    risk_difference_row = next(row for row in absolute_effect_rows if "difference" in row["estimand"].lower())
    balance_rows = sorted(datasets["balance"]["rows"], key=lambda row: float(row["absolute_smd_unweighted"]), reverse=True)
    sensitivity_rows = datasets["sensitivity"]["rows"]
    weights = [float(row["stabilized_weight"]) for row in datasets["weight_distribution"]["rows"]]
    slides = [
        make_slide(kind, "cover", "合成目标试验模拟：输血策略与28天死亡风险", "cover", "声明合成测试边界并提出研究问题", "cover", [], claims, coverage_items=covered_items(brief, "synthetic")),
        make_slide(kind, "question", "目标试验先定义问题，再解释估计量", "context", "说明目标试验模拟的科学问题与学习目标", "hero_insight", [
            narrative_visual(kind, "question", "hero_question", protocol, {
                "question": "若在入组后24小时内实施不同输血策略，28天死亡风险如何比较？",
                "gap": "观察性资料需要显式模拟资格、分配、随访与分析规则。",
                "objective": "展示预先冻结的目标试验组成、诊断与效应估计。",
            })
        ], claims, coverage_items=covered_items(brief, "target trial components")),
        make_slide(kind, "specification", "八项规格共同定义可解释的目标试验", "method", "完整呈现目标试验规格", "comparison_matrix", [
            narrative_visual(kind, "specification", "specification_matrix", protocol, {
                "items": [
                    ["资格标准", "合成ICU队列与预设纳排标准"],
                    ["治疗策略", "24小时内输血 vs 24小时内未输血"],
                    ["分配规则", "在time zero按策略克隆"],
                    ["宽限期", "canonical窗口为24小时"],
                    ["随访", "从time zero至28天"],
                    ["结局", "28天死亡风险"],
                    ["估计量", "风险、风险差、风险比及支持性HR"],
                    ["分析", "Clone–Censor–Weight"],
                ]
            }, "method")
        ], claims, coverage_items=covered_items(brief, "eligibility", "strategies", "assignment", "grace period", "follow-up", "outcome", "estimand")),
        make_slide(kind, "ccw", "Clone–Censor–Weight把策略偏离纳入估计", "method", "解释克隆、删失和加权的分支逻辑", "process_branch", [
            narrative_visual(kind, "ccw", "clone_censor_weight_branch", memo, {
                "nodes": ["符合资格", "克隆至两种策略", "监测策略依从", "偏离时人工删失", "估计稳定权重", "比较加权结局"],
                "branches": [[0, 1], [1, 2], [2, 3], [3, 4], [4, 5]],
                "warning": "权重不能消除未测量混杂。",
            }, "method")
        ], claims, coverage_items=covered_items(brief, "clone-censor-weight")),
        make_slide(kind, "grace_period", "24小时宽限期连接基线、策略与随访", "method", "按真实时间顺序呈现界标与宽限期", "process_branch", [
            narrative_visual(kind, "grace_period", "time_window_timeline", protocol, {
                "events": [
                    {"label": "time zero / 入组", "hour": 0},
                    {"label": "策略评估开始", "hour": 0},
                    {"label": "宽限期结束", "hour": 24},
                    {"label": "随访结局", "hour": 672},
                ],
                "unit": "小时",
            }, "method")
        ], claims, coverage_items=covered_items(brief, "grace period", "follow-up")),
    ]

    balance_binding = source_binding(
        datasets["balance"],
        ["covariate", "absolute_smd_unweighted", "absolute_smd_weighted"],
        "rows 2-8", "all 7 covariates", "ordered by unweighted absolute SMD",
    )
    slides.append(make_slide(kind, "balance", "加权后所有已测协变量的|SMD|均低于0.10", "method", "使用定量协变量平衡图展示改善", "chart_led", [
        make_visual(kind, "balance", "love_plot", "diagnostic", [balance_binding], {"rows": balance_rows, "threshold": 0.10}, len(balance_rows),
                    x_field="absolute_smd_unweighted", y_field="absolute_smd_weighted", label_field="covariate", reference_value=0.10,
                    unit="absolute standardized mean difference", entity_scope="all measured covariates")
    ], claims, coverage_items=covered_items(brief, "covariate balance")))

    weight_binding = source_binding(
        datasets["weight_distribution"], ["stabilized_weight"], f"all {len(weights)} rows", "all rows", "18 equal-width histogram bins",
    )
    bins = histogram(weights)
    slides.append(make_slide(kind, "weights", "稳定权重集中在1附近，但尾部仍需诊断", "method", "完整汇总权重分布并提示极端权重", "chart_led", [
        make_visual(kind, "weights", "weight_histogram", "diagnostic", [weight_binding], {
            "bins": bins,
            "summary": {"n": len(weights), "median": statistics.median(weights), "p99": sorted(weights)[math.floor(len(weights) * .99) - 1], "max": max(weights)},
        }, len(weights), aggregation="18 equal-width histogram bins", x_field="stabilized_weight", y_field="count", label_field="bin", unit="stabilized weight", entity_scope="all clone-interval weights")
    ], claims, coverage_items=covered_items(brief, "weight distribution")))

    effect_binding = source_binding(
        datasets["effects"], ["claim_id", "estimand", "contrast", "estimate", "ci_low", "ci_high", "status"],
        "rows 2-4", "absolute risk and risk-difference rows", "none",
    )
    direction = "负向" if float(risk_difference_row["estimate"]) < 0 else "正向"
    slides.append(make_slide(kind, "absolute_effect", f"两种策略的估计风险相差{abs(float(risk_difference_row['estimate'])) * 100:.1f}个百分点（风险差{direction}）", "result", "用绝对风险作为主结果视觉", "hero_insight", [
        make_visual(kind, "absolute_effect", "absolute_risk_comparison", "primary_result", [effect_binding], {"rows": absolute_effect_rows}, len(absolute_effect_rows),
                    outcome="28-day mortality", estimand="absolute risk and risk difference", unit="percent and percentage points",
                    x_field="estimate", label_field="estimand", ci_low_field="ci_low", ci_high_field="ci_high", reference_value=0,
                    entity_scope="absolute-risk estimands")
    ], claims, ["CLM-TTE-001", "CLM-TTE-002", "CLM-TTE-003"], coverage_items=covered_items(brief, "absolute 28-day risks")))

    ratio_binding = source_binding(
        datasets["effects"], ["claim_id", "estimand", "contrast", "estimate", "ci_low", "ci_high", "status"],
        "rows 5-6", "ratio estimand rows", "none",
    )
    slides.append(make_slide(kind, "ratio_effects", "风险比与加权HR方向一致，但回答不同问题", "result", "分区呈现不同比分尺度估计量", "split_screen", [
        make_visual(kind, "ratio_effects", "estimand_comparison", "supportive_result", [ratio_binding], {"rows": ratio_effect_rows}, len(ratio_effect_rows),
                    outcome="28-day mortality", estimand="risk ratio and weighted hazard ratio", unit="ratio",
                    x_field="estimate", label_field="estimand", ci_low_field="ci_low", ci_high_field="ci_high", reference_value=1,
                    entity_scope="ratio estimands, separated panels")
    ], claims, ["CLM-TTE-004", "CLM-TTE-005"], coverage_items=covered_items(brief, "risk ratio")))

    sensitivity_binding = source_binding(
        datasets["sensitivity"], ["analysis", "risk_difference", "ci_low", "ci_high"],
        "rows 2-6", "all sensitivity analyses", "none",
    )
    direction_label = "负向" if all(float(row["risk_difference"]) < 0 for row in sensitivity_rows) else "方向不完全一致的"
    slides.append(make_slide(kind, "sensitivity", f"{len(sensitivity_rows)}项敏感性分析呈{direction_label}风险差", "result", "用分析标签和共同参考线检验稳健性", "chart_led", [
        make_visual(kind, "sensitivity", "forest_plot", "sensitivity_analysis", [sensitivity_binding], {"rows": sensitivity_rows, "scale": "linear", "axis_label": "风险差（比例）"}, len(sensitivity_rows),
                    outcome="28-day mortality", estimand="risk difference", unit="proportion", x_field="risk_difference",
                    label_field="analysis", ci_low_field="ci_low", ci_high_field="ci_high", reference_value=0,
                    entity_scope="all registered sensitivity analyses")
    ], claims, coverage_items=covered_items(brief, "sensitivity")))

    slides.extend([
        make_slide(kind, "assumptions", "可交换性、阳性与一致性决定解释边界", "limitation", "把方法假设与失败模式成对呈现", "comparison_matrix", [
            narrative_visual(kind, "assumptions", "assumption_warning_matrix", memo, {
                "items": [
                    ["可交换性", "残余混杂仍可能存在", "高"],
                    ["阳性", "极端权重提示策略可及性不足", "中"],
                    ["一致性", "输血策略必须有可执行定义", "中"],
                    ["模型依赖", "结果依赖权重模型与删失规则", "中"],
                ]
            }, "limitation")
        ], claims, coverage_items=covered_items(brief, "assumptions", "failure modes"), card_grid=True),
        make_slide(kind, "boundary", "合成估计用于方法教学，不形成治疗建议", "limitation", "明确科学表达允许与禁止边界", "split_screen", [
            narrative_visual(kind, "boundary", "evidence_boundary", protocol, {
                "allowed": ["在既定假设下的估计风险", "与较低风险一致", "结果对敏感性设定稳健"],
                "prohibited": ["输血可降低真实患者死亡", "随机证据已经证明", "可直接制定临床策略"],
            }, "limitation")
        ], claims),
        make_slide(kind, "conclusion", "先定义目标试验，再解读冻结估计与诊断", "conclusion", "回扣方法教学主线与限制", "conclusion_synthesis", [
            narrative_visual(kind, "conclusion", "takeaway_synthesis", memo, {
                "takeaways": ["规格决定估计量含义", "诊断决定估计稳定性", "合成结果不得外推为治疗结论"],
            }, "conclusion")
        ], claims, [item["claim_id"] for item in claims]),
        make_slide(kind, "appendix_review", f"{len(conflicts)}项冲突与{len(unresolved)}项待确认内容保留审计", "appendix", "集中呈现冲突与未解决事项", "appendix_audit", [], claims,
                   conflicts=conflicts, unresolved=unresolved, appendix=True),
    ])
    return slides


def build_single_cell_slides(
    project: Path, brief: dict[str, Any], sources: list[dict[str, Any]], datasets: dict[str, dict[str, Any]],
    claims: list[dict[str, Any]], conflicts: list[dict[str, Any]], unresolved: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    kind = "single_cell"
    manuscript = document_binding(sources, "single_cell_manuscript", "study question and interpretation")
    methods = document_binding(sources, "computational_methods", "computational workflow")
    qc = datasets["qc"]
    final_qc = min(qc["rows"], key=lambda row: float(row["remaining_cells"]))
    composition = datasets["composition"]
    group_labels = list(dict.fromkeys(row["group"] for row in composition["rows"]))
    by_cell_type: dict[str, list[dict[str, Any]]] = {}
    for row in composition["rows"]:
        by_cell_type.setdefault(row["cell_type"], []).append(row)
    featured_cell_type, featured_rows = max(
        by_cell_type.items(),
        key=lambda item: max(float(row["percentage"]) for row in item[1]) - min(float(row["percentage"]) for row in item[1]),
    )
    high_group = max(featured_rows, key=lambda row: float(row["percentage"]))
    low_group = min(featured_rows, key=lambda row: float(row["percentage"]))
    composition_delta = float(high_group["percentage"]) - float(low_group["percentage"])
    group_comparison = "与".join(zh_entity_label(label) for label in group_labels)
    featured_cell_type_zh = zh_entity_label(featured_cell_type)
    high_group_zh = zh_entity_label(high_group["group"])
    low_group_zh = zh_entity_label(low_group["group"])
    slides: list[dict[str, Any]] = [
        make_slide(kind, "cover", "合成单细胞研究：细胞组成、状态与预测通讯", "cover", "声明合成数据与计算推断边界", "cover", [], claims, coverage_items=covered_items(brief, "synthetic")),
        make_slide(kind, "question", f"{group_comparison}的细胞组成与状态差异需要分层解释", "context", "提出单细胞研究问题与证据层级", "hero_insight", [
            narrative_visual(kind, "question", "hero_question", manuscript, {
                "question": f"{group_comparison}的细胞组成、转录状态与预测通讯如何变化？",
                "gap": "计算信号不能自动升级为机制或临床标志物。",
                "objective": "串联QC、细胞注释、pseudobulk、通路和通讯证据。",
            })
        ], claims),
        make_slide(kind, "workflow", "分析流程把细胞级QC连接到样本级推断", "method", "解释scRNA-seq分析层级", "process_branch", [
            narrative_visual(kind, "workflow", "single_cell_workflow", methods, {
                "nodes": ["原始细胞", "QC与双细胞过滤", "细胞类型注释", "样本级聚合", "通路富集", "预测细胞通讯", "外部验证"],
                "branches": [[0, 1], [1, 2], [2, 3], [3, 4], [3, 5], [4, 6], [5, 6]],
                "warning": "计算推断止于待验证假设。",
            }, "method")
        ], claims, coverage_items=covered_items(brief, "scrna-seq workflow")),
    ]

    qc_binding = source_binding(qc, ["stage", "remaining_cells", "removed_at_stage"], f"rows 2-{len(qc['rows']) + 1}", "all QC stages", "none")
    slides.append(make_slide(kind, "qc", f"{len(qc['rows'])}步QC保留{int(float(final_qc['remaining_cells'])):,}个合成细胞", "method", "显示每步保留与剔除数量", "process_branch", [
        make_visual(kind, "qc", "qc_funnel", "quality_control", [qc_binding], {"rows": qc["rows"]}, len(qc["rows"]),
                    x_field="stage", y_field="remaining_cells", label_field="removed_at_stage", unit="cells", entity_scope="all QC stages")
    ], claims, ["CLM-SC-001"], coverage_items=covered_items(brief, "qc flow")))

    metadata = datasets["cell_metadata"]
    umap_asset = asset_binding(sources, "umap", "registered UMAP figure")
    metadata_binding = source_binding(metadata, ["cell_id", "group", "cell_type", "UMAP_1", "UMAP_2"], f"all {len(metadata['rows'])} rows", "all retained cells", "none")
    umap_source = next(source for source in sources if source["source_id"] == umap_asset["source_id"])
    slides.append(make_slide(kind, "umap", "UMAP展示细胞类型结构，但不代表谱系或机制", "result", "用全幅来源图加解释性callout", "figure_with_callout", [
        make_visual(kind, "umap", "umap_figure", "descriptive_embedding", [metadata_binding, umap_asset], {
            "image_path": str((project / umap_source["source_file"]).resolve()),
            "callout": "邻近关系仅描述低维嵌入；不能解释为谱系、因果或机制。",
        }, len(metadata["rows"]), x_field="UMAP_1", y_field="UMAP_2", group_field="cell_type", color_field="cell_type", label_field="cell_type", unit="embedding coordinate", entity_scope="all retained cells")
    ], claims, coverage_items=covered_items(brief, "umap")))

    composition_binding = source_binding(composition, ["group", "cell_type", "cell_count", "percentage"], f"rows 2-{len(composition['rows']) + 1}", f"{group_comparison}的全部细胞类型", "none")
    slides.append(make_slide(kind, "composition", f"{high_group_zh}组{featured_cell_type_zh}比例高{composition_delta:.2f}个百分点", "result", "定量比较两组全部细胞类型组成", "chart_led", [
        make_visual(kind, "composition", "grouped_composition", "descriptive_comparison", [composition_binding], {"rows": composition["rows"], "sort": "absolute_delta_desc"}, len(composition["rows"]),
                    x_field="cell_type", y_field="percentage", group_field="group", label_field="cell_type", unit="percent", entity_scope=f"all cell types in {group_comparison}")
    ], claims, ["CLM-SC-002"], coverage_items=covered_items(brief, "cell composition")))

    deg = datasets["deg"]
    effect_field = next(field for field in deg["columns"] if field.lower().startswith("log2_fold_change"))
    deg_binding = source_binding(deg, ["gene", effect_field, "FDR", "cell_type"], f"rows 2-{len(deg['rows']) + 1}", "all genes, faceted by cell type", "none")
    slides.append(make_slide(kind, "deg", "Pseudobulk差异信号按细胞类型分层呈现", "result", "保留所有方向与细胞类型范围", "split_screen", [
        make_visual(kind, "deg", "faceted_volcano", "computational_association", [deg_binding], {"rows": deg["rows"], "top_label_policy": "top 2 labels per facet by FDR, collision-adjusted"}, len(deg["rows"]),
                    x_field=effect_field, y_field="FDR", facet_field="cell_type", label_field="gene", reference_value=0,
                    unit="log2 fold change and -log10 FDR", entity_scope="all registered cell-type pseudobulk rows", top_k_policy="top 2 labels per facet by FDR, collision-adjusted; all points retained")
    ], claims, ["CLM-SC-003"], coverage_items=covered_items(brief, "pseudobulk deg")))

    pathway = datasets["pathway"]
    pathway_binding = source_binding(pathway, ["pathway", "normalized_enrichment_score", "FDR", "cell_type", "evidence_level"], f"rows 2-{len(pathway['rows']) + 1}", "all pathways, faceted by cell type", "none")
    slides.append(make_slide(kind, "pathway", "不同细胞类型呈现方向相反的通路富集", "result", "用发散轴和细胞类型分面保留NES方向", "chart_led", [
        make_visual(kind, "pathway", "diverging_enrichment", "supported_interpretation", [pathway_binding], {"rows": pathway["rows"]}, len(pathway["rows"]),
                    x_field="normalized_enrichment_score", y_field="pathway", facet_field="cell_type", color_field="normalized_enrichment_score", label_field="pathway",
                    reference_value=0, unit="normalized enrichment score", entity_scope="all pathways, explicitly faceted by cell type")
    ], claims, ["CLM-SC-004"], coverage_items=covered_items(brief, "pathway enrichment")))

    communication = datasets["communication"]
    communication_binding = source_binding(communication, ["sender", "receiver", "ligand_receptor", "communication_probability", "adjusted_p"], f"rows 2-{len(communication['rows']) + 1}", "all independent interactions", "none")
    slides.append(make_slide(kind, "communication", f"{len(communication['rows'])}条预测互作构成独立的有向网络", "result", "以发送者—接收者边而非连续流程呈现通讯", "figure_with_callout", [
        make_visual(kind, "communication", "communication_network", "predicted_interaction", [communication_binding], {"rows": communication["rows"], "prediction_badge": True}, len(communication["rows"]),
                    edge_source_field="sender", edge_target_field="receiver", label_field="ligand_receptor", color_field="adjusted_p", unit="communication probability",
                    entity_scope="all independent sender-receiver interactions", sequential=False)
    ], claims, ["CLM-SC-005"], coverage_items=covered_items(brief, "cell-cell communication")))

    multi_bindings = [deg_binding, pathway_binding, communication_binding]
    slides.extend([
        make_slide(kind, "evidence_ladder", "计算证据越接近机制，验证要求越高", "limitation", "建立从描述到外部验证的证据阶梯", "evidence_ladder", [
            make_visual(kind, "evidence_ladder", "evidence_ladder", "evidence_boundary", multi_bindings, {
                "steps": ["细胞组成", "pseudobulk差异", "通路富集", "预测互作", "外部实验验证"],
                "status": ["描述性", "计算关联", "支持性解释", "预测性", "尚未完成"],
            }, len(deg["rows"]) + len(pathway["rows"]) + len(communication["rows"]), aggregation="evidence-level synthesis", entity_scope="all computational evidence layers")
        ], claims, [item["claim_id"] for item in claims], coverage_items=covered_items(brief, "evidence ladder")),
        make_slide(kind, "validation_gap", "外部数据与实验验证仍是关键缺口", "limitation", "保留未验证边界与待确认标记", "split_screen", [
            narrative_visual(kind, "validation_gap", "validation_gap", manuscript, {
                "available": ["合成细胞组成", "pseudobulk统计", "通路与通讯计算推断"],
                "missing": ["外部队列复现", "蛋白或功能实验", "临床效用评估"],
            }, "limitation")
        ], claims, coverage_items=covered_items(brief, "external validation gap")),
        make_slide(kind, "interpretation", "共现的炎症与代谢信号不等于机制链条", "limitation", "防止计算结果机制化", "split_screen", [
            narrative_visual(kind, "interpretation", "evidence_boundary", manuscript, {
                "allowed": ["共同出现", "与……一致", "优先验证"],
                "prohibited": ["证明机制", "实验已验证", "可作为临床生物标志物"],
            }, "limitation")
        ], claims),
        make_slide(kind, "conclusion", "合成信号覆盖组成、状态与预测通讯三个证据层级", "conclusion", "综合主要计算发现与验证边界", "conclusion_synthesis", [
            make_visual(kind, "conclusion", "takeaway_synthesis", "conclusion", multi_bindings, {
                "takeaways": [f"{high_group_zh}组{featured_cell_type_zh}比例较{low_group_zh}组高{composition_delta:.2f}个百分点", "细胞类型特异的转录与通路方向", "预测通讯需要外部验证"],
            }, len(deg["rows"]) + len(pathway["rows"]) + len(communication["rows"]), aggregation="claim-level synthesis", entity_scope="registered computational findings")
        ], claims, [item["claim_id"] for item in claims]),
        make_slide(kind, "appendix_review", f"{len(conflicts)}项冲突与{len(unresolved)}项待确认内容保留审计", "appendix", "集中呈现冲突与未解决事项", "appendix_audit", [], claims,
                   conflicts=conflicts, unresolved=unresolved, appendix=True),
    ])
    return slides


def build_meta_slides(
    project: Path, brief: dict[str, Any], sources: list[dict[str, Any]], datasets: dict[str, dict[str, Any]],
    claims: list[dict[str, Any]], conflicts: list[dict[str, Any]], unresolved: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    kind = "meta_analysis"
    manuscript = document_binding(sources, "meta_analysis_manuscript", "review question and PRISMA narrative")
    methods = document_binding(sources, "review_methods", "review methods")
    studies = datasets["study_level"]
    pooled = datasets["pooled"]
    rob = datasets["risk_of_bias"]
    subgroup = datasets["subgroup"]
    grade = datasets["grade"]
    study_count = len(studies["rows"])
    rob_domain_count = len({row["domain"] for row in rob["rows"]})
    study_binding = source_binding(studies, ["study_id", "study_label", "year", "sample_size", "death_events", "odds_ratio", "ci_low", "ci_high", "risk_of_bias"], f"rows 2-{study_count + 1}", f"all {study_count} studies", "none")
    pooled_binding = source_binding(pooled, ["claim_id", "estimand", "estimate", "ci_low", "ci_high", "I2_percent", "study_count", "status"], "rows 2-4", "all pooled estimands", "none")
    rob_binding = source_binding(rob, ["study_id", "study_label", "domain", "judgment"], f"rows 2-{len(rob['rows']) + 1}", f"all {study_count} studies and {rob_domain_count} domains", "pivot to study-by-domain matrix")
    subgroup_binding = source_binding(subgroup, ["subgroup", "pooled_or", "ci_low", "ci_high", "interaction_p"], "rows 2-7", "all subgroup analyses", "none")
    grade_binding = source_binding(grade, ["outcome", "studies", "participants", "risk_of_bias", "inconsistency", "indirectness", "imprecision", "other", "certainty"], "rows 2-4", "all GRADE outcomes", "none")

    prisma_asset = asset_binding(sources, "prisma", "registered PRISMA figure")
    prisma_source = next(source for source in sources if source["source_id"] == prisma_asset["source_id"])
    funnel_asset = asset_binding(sources, "funnel", "registered funnel figure")
    funnel_source = next(source for source in sources if source["source_id"] == funnel_asset["source_id"])

    total_n = sum(int(row["sample_size"]) for row in studies["rows"])
    total_events = sum(int(row["death_events"]) for row in studies["rows"])
    random_effect = next(row for row in pooled["rows"] if "Random-effects" in row["estimand"])
    fixed_effect = next(row for row in pooled["rows"] if "Fixed-effect" in row["estimand"])
    prediction = next(row for row in pooled["rows"] if "Prediction interval" in row["estimand"])

    slides: list[dict[str, Any]] = [
        make_slide(kind, "cover", "合成Meta分析：脓毒症右心功能障碍与死亡", "cover", "声明合成证据与观察性边界", "cover", [], claims),
        make_slide(kind, "question", "评估目标是关联强度与证据确定性，而非因果证明", "context", "明确PICO式综述问题和解释目标", "hero_insight", [
            narrative_visual(kind, "question", "hero_question", manuscript, {
                "question": "脓毒症患者右心功能障碍是否与更高死亡风险相关？",
                "population": "成人脓毒症或脓毒性休克队列",
                "exposure": "超声定义的右心功能障碍",
                "outcome": "短期或ICU死亡",
            })
        ], claims, coverage_items=covered_items(brief, "review question")),
        make_slide(kind, "prisma", f"{study_count}项研究进入定量合并", "method", "呈现来源PRISMA流程并避免补造缺失计数", "figure_with_callout", [
            make_visual(kind, "prisma", "source_figure", "study_selection", [prisma_asset, manuscript, study_binding], {
                "image_path": str((project / prisma_source["source_file"]).resolve()),
                "callout": f"规范来源的合并数据包含{study_count}项研究；未在结构化来源中提供的筛选计数不补造。",
            }, study_count, entity_scope=f"registered PRISMA figure and {study_count} included studies")
        ], claims, ["CLM-META-004"], coverage_items=covered_items(brief, "prisma flow")),
        make_slide(kind, "characteristics", f"{study_count}项研究共纳入{total_n:,}例并报告{total_events:,}例死亡", "method", "概括研究规模、年份与偏倚分布", "split_screen", [
            make_visual(kind, "characteristics", "study_characteristics", "study_inventory", [study_binding], {"rows": studies["rows"], "total_n": total_n, "total_events": total_events}, len(studies["rows"]),
                        x_field="year", y_field="sample_size", group_field="risk_of_bias", label_field="study_label", unit="participants", entity_scope=f"all {study_count} included studies")
        ], claims, ["CLM-META-004"], coverage_items=covered_items(brief, "study characteristics")),
        make_slide(kind, "forest", f"{study_count}项研究的方向总体支持正向关联", "result", "显示完整研究级森林图与合并菱形", "chart_led", [
            make_visual(kind, "forest", "forest_plot", "primary_meta_analysis", [study_binding, pooled_binding], {"rows": studies["rows"], "pooled": random_effect, "scale": "log", "axis_label": "比值比（OR）"}, len(studies["rows"]),
                        outcome="mortality", estimand="study-level odds ratio and random-effects pooled odds ratio", unit="ratio", x_field="odds_ratio", label_field="study_label",
                        ci_low_field="ci_low", ci_high_field="ci_high", reference_value=1, entity_scope=f"all {study_count} studies plus pooled diamond")
        ], claims, ["CLM-META-001", "CLM-META-004"], coverage_items=covered_items(brief, "forest plot")),
        make_slide(kind, "pooled", "合并OR为2.12，但预测区间更宽", "result", "把合并CI、异质性和预测区间作为统一证据单元", "hero_insight", [
            make_visual(kind, "pooled", "pooled_evidence_panel", "evidence_interpretation", [pooled_binding], {"rows": pooled["rows"], "fixed": fixed_effect, "random": random_effect, "prediction": prediction}, len(pooled["rows"]),
                        outcome="mortality", estimand="random-effects pooled OR and prediction interval", unit="ratio", x_field="estimate", label_field="estimand",
                        ci_low_field="ci_low", ci_high_field="ci_high", reference_value=1, entity_scope="all pooled estimands; prediction interval labeled separately")
        ], claims, ["CLM-META-001", "CLM-META-002", "CLM-META-003"], coverage_items=covered_items(brief, "heterogeneity", "prediction interval")),
        make_slide(kind, "subgroup", "亚组点估计不同，但交互检验未支持异质性", "result", "显示亚组OR、CI和交互检验P值", "chart_led", [
            make_visual(kind, "subgroup", "subgroup_forest", "subgroup_analysis", [subgroup_binding], {"rows": subgroup["rows"], "scale": "log", "axis_label": "合并OR"}, len(subgroup["rows"]),
                        outcome="mortality", estimand="subgroup pooled odds ratio", unit="ratio", x_field="pooled_or", label_field="subgroup",
                        ci_low_field="ci_low", ci_high_field="ci_high", reference_value=1, entity_scope="all registered subgroup analyses")
        ], claims, coverage_items=covered_items(brief, "subgroups")),
        make_slide(kind, "funnel", "漏斗图只能提示不对称，不能排除发表偏倚", "limitation", "使用来源图并显示谨慎解释", "figure_with_callout", [
            make_visual(kind, "funnel", "source_figure", "publication_bias_assessment", [funnel_asset, study_binding], {
                "image_path": str((project / funnel_source["source_file"]).resolve()),
                "callout": "视觉对称或不对称均不能单独证明或排除发表偏倚。",
            }, len(studies["rows"]), entity_scope="all study-level funnel points")
        ], claims, coverage_items=covered_items(brief, "funnel plot")),
        make_slide(kind, "rob", "偏倚风险在研究与领域间分布不均", "limitation", f"以{study_count}×{rob_domain_count}交通灯矩阵完整呈现偏倚判断", "risk_heatmap", [
            make_visual(kind, "rob", "risk_of_bias_matrix", "risk_of_bias", [rob_binding], {"rows": rob["rows"]}, len(rob["rows"]),
                        x_field="domain", y_field="study_label", group_field="study_label", facet_field="domain", color_field="judgment", label_field="judgment",
                        unit="judgment category", entity_scope=f"{study_count} studies by {rob_domain_count} domains")
        ], claims, coverage_items=covered_items(brief, "risk of bias")),
        make_slide(kind, "grade", "GRADE为低或极低，限制结论强度", "limitation", "将结局、降级领域和确定性形成一张证据表", "evidence_ladder", [
            make_visual(kind, "grade", "grade_summary", "certainty_assessment", [grade_binding], {"rows": grade["rows"]}, len(grade["rows"]),
                        x_field="outcome", y_field="certainty", group_field="certainty", label_field="outcome", unit="certainty category", entity_scope="all GRADE outcomes")
        ], claims, ["CLM-META-005"], coverage_items=covered_items(brief, "grade certainty")),
        make_slide(kind, "limitations", "观察性设计、偏倚与间接性共同限制推断", "limitation", "合并研究设计与GRADE边界", "split_screen", [
            make_visual(kind, "limitations", "evidence_boundary", "evidence_limitations", [rob_binding, grade_binding, methods], {
                "allowed": ["存在统计关联", "结果方向总体一致", "未来研究可能有差异"],
                "prohibited": ["右心功能障碍导致死亡", "漏斗图排除发表偏倚", "证据支持治疗建议"],
            }, len(rob["rows"]) + len(grade["rows"]), aggregation="limitation synthesis", entity_scope="risk-of-bias and certainty evidence")
        ], claims, coverage_items=covered_items(brief, "evidence limitations")),
        make_slide(kind, "conclusion", "关联信号一致，但低确定性要求克制解读", "conclusion", "综合合并效应、异质性与确定性", "conclusion_synthesis", [
            make_visual(kind, "conclusion", "takeaway_synthesis", "conclusion", [pooled_binding, grade_binding], {
                "takeaways": [f"合并OR {random_effect['estimate']}（95%CI {random_effect['ci_low']}–{random_effect['ci_high']}）", f"预测区间 {prediction['ci_low']}–{prediction['ci_high']}", "主要结局GRADE：低"],
            }, len(pooled["rows"]) + len(grade["rows"]), aggregation="claim-level synthesis", entity_scope="pooled and certainty evidence")
        ], claims, [item["claim_id"] for item in claims]),
        make_slide(kind, "appendix_review", f"{len(conflicts)}项冲突与{len(unresolved)}项待确认内容保留审计", "appendix", "集中呈现冲突与未解决事项", "appendix_audit", [], claims,
                   conflicts=conflicts, unresolved=unresolved, appendix=True),
    ]
    return slides


def schedule_hour(window: str) -> float:
    text = window.lower().strip()
    if "day 28" in text:
        return 28 * 24
    match = __import__("re").search(r"(\d+)(?:\s*[-+]\s*(\d+))?\s*h", text)
    if not match:
        return 0.0
    if match.group(2):
        return (float(match.group(1)) + float(match.group(2))) / 2
    return float(match.group(1))


def build_protocol_slides(
    brief: dict[str, Any], sources: list[dict[str, Any]], datasets: dict[str, dict[str, Any]],
    claims: list[dict[str, Any]], conflicts: list[dict[str, Any]], unresolved: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    kind = "protocol"
    protocol = document_binding(sources, "full_protocol", "study design and analysis plan")
    operations = document_binding(sources, "operations_manual", "operations and quality control")
    eligibility = datasets["eligibility"]
    schedule = datasets["schedule"]
    sample = datasets["sample_size"]
    variables = datasets["variables"]
    risks = datasets["risk_register"]
    milestones = datasets["milestones"]
    eligibility_binding = source_binding(eligibility, ["category", "criterion", "operational_status"], "rows 2-9", "all 4 inclusion and 4 exclusion criteria", "none")
    schedule_binding = source_binding(schedule, ["visit", "assessment", "window", "status"], "rows 2-8", "all scheduled assessments", "time-position derivation from window")
    sample_binding = source_binding(sample, ["parameter", "value", "unit"], "rows 2-8", "all sample-size assumptions", "none")
    variable_binding = source_binding(variables, ["role", "variable", "timepoint", "type"], "rows 2-11", "all registered variables", "grouped by role")
    risk_binding = source_binding(risks, ["risk", "domain", "probability", "impact", "mitigation"], "rows 2-7", "all operational risks", "probability-impact matrix")
    milestone_binding = source_binding(milestones, ["milestone", "start_date", "end_date", "status"], "rows 2-8", "all project milestones", "date-scaled Gantt")

    schedule_rows = [{**row, "position_hours": schedule_hour(row["window"])} for row in schedule["rows"]]
    risk_level = {"Low": 1, "Medium": 2, "High": 3}
    risk_rows = [{**row, "probability_score": risk_level[row["probability"]], "impact_score": risk_level[row["impact"]]} for row in risks["rows"]]
    exposure_row = next(row for row in variables["rows"] if "exposure" in row["role"].lower())
    outcome_row = next(row for row in variables["rows"] if row["role"].lower().startswith("primary outcome"))
    prespecified_confounders = [
        row for row in variables["rows"]
        if row["role"].lower() == "covariate"
        and ("baseline" in row["timepoint"].lower() or "before" in row["timepoint"].lower())
    ]
    same_window_variables = [
        row for row in variables["rows"]
        if row["role"].lower() == "covariate" and row not in prespecified_confounders
    ]
    exposure_label = zh_entity_label(exposure_row["variable"])
    outcome_label = zh_entity_label(outcome_row["variable"])
    nodes = [
        {
            "id": "confounder_pre_exposure",
            "label": "暴露前混杂因素\n" + "、".join(zh_entity_label(row["variable"]) for row in prespecified_confounders),
            "role": "confounder",
        },
        {"id": "exposure", "label": f"{exposure_label}\n{exposure_row['timepoint']}", "role": "exposure"},
        {"id": "outcome", "label": f"{outcome_label}\n{outcome_row['timepoint']}", "role": "outcome"},
    ]
    edges = [
        {"source": "confounder_pre_exposure", "target": "exposure"},
        {"source": "confounder_pre_exposure", "target": "outcome"},
        {"source": "exposure", "target": "outcome"},
    ]

    target = next(row for row in sample["rows"] if row["parameter"] == "Target enrollment")
    first_assessment = min(schedule_rows, key=lambda row: row["position_hours"])
    last_assessment = max(schedule_rows, key=lambda row: row["position_hours"])
    primary_endpoint_schedule = next(row for row in schedule_rows if row["status"] == "Primary endpoint")
    slides: list[dict[str, Any]] = [
        make_slide(kind, "cover", f"前瞻性研究方案：{exposure_label}与{outcome_label}", "cover", "声明方案状态与无结果边界", "cover", [], claims),
        make_slide(kind, "question", "明确时间窗是检验预后关联的前提", "context", "呈现研究问题、设计与站点状态", "hero_insight", [
            narrative_visual(kind, "question", "protocol_question", protocol, {
                "question": f"{exposure_label}是否与{outcome_label}相关？",
                "design": "前瞻性、多中心、观察性队列",
                "sites": "计划使用合成站点；最终站点清单待确认",
                "boundary": "研究方案尚无任何观察结果",
            })
        ], claims, coverage_items=covered_items(brief, "study question", "design and sites")),
        make_slide(kind, "design", f"暴露在{exposure_row['timepoint']}测量，主要结局在{primary_endpoint_schedule['window']}判定", "method", "用时间顺序隔离暴露、协变量和结局", "process_branch", [
            make_visual(kind, "design", "protocol_design", "study_design", [schedule_binding, variable_binding], {
                "nodes": [
                    f"{first_assessment['visit']} / time zero",
                    f"{exposure_label}\n{exposure_row['timepoint']}",
                    "预设基线与暴露前协变量",
                    f"{outcome_label}\n{primary_endpoint_schedule['window']}",
                    f"{last_assessment['assessment']}\n{last_assessment['window']}",
                ],
                "branches": [[0, 1], [0, 2], [1, 3], [2, 3], [3, 4]],
            }, len(schedule["rows"]) + len(variables["rows"]), aggregation="design synthesis", entity_scope="registered exposure, covariates and outcomes")
        ], claims, ["CLM-PROTO-003", "CLM-PROTO-004"]),
        make_slide(kind, "eligibility", "四项纳入与四项排除标准共同定义目标队列", "method", "完整呈现所有纳排标准", "split_screen", [
            make_visual(kind, "eligibility", "eligibility_split", "cohort_definition", [eligibility_binding], {"rows": eligibility["rows"]}, len(eligibility["rows"]),
                        group_field="category", label_field="criterion", color_field="operational_status", entity_scope="all inclusion and exclusion criteria")
        ], claims, coverage_items=covered_items(brief, "eligibility")),
        make_slide(kind, "schedule", "评估节点按0小时至28天的真实时间展开", "method", "按真实时间窗定位所有评估", "chart_led", [
            make_visual(kind, "schedule", "assessment_timeline", "assessment_schedule", [schedule_binding], {"rows": schedule_rows}, len(schedule_rows),
                        time_field="position_hours", label_field="assessment", group_field="status", unit="hours from enrollment", entity_scope="all scheduled assessments")
        ], claims, ["CLM-PROTO-003"], coverage_items=covered_items(brief, "assessment schedule")),
        make_slide(kind, "definitions", "变量字典把暴露、结局与调整变量分层", "method", "完整呈现变量角色和测量时间", "comparison_matrix", [
            make_visual(kind, "definitions", "variable_definition_matrix", "variable_contract", [variable_binding], {"rows": variables["rows"]}, len(variables["rows"]),
                        group_field="role", label_field="variable", time_field="timepoint", color_field="role", entity_scope="all registered variables")
        ], claims, ["CLM-PROTO-004"], coverage_items=covered_items(brief, "exposure and outcome definitions")),
        make_slide(kind, "sample_size", f"计划入组{target['value']}例，以获得{next(row for row in sample['rows'] if row['parameter'] == 'Required analyzable sample')['value']}例可分析样本", "planning", "把全部样本量假设与目标入组关联", "hero_insight", [
            make_visual(kind, "sample_size", "sample_size_waterfall", "sample_size_planning", [sample_binding], {"rows": sample["rows"]}, len(sample["rows"]),
                        x_field="parameter", y_field="value", group_field="unit", label_field="parameter", unit="parameter-specific", entity_scope="all sample-size assumptions")
        ], claims, ["CLM-PROTO-001", "CLM-PROTO-002"], coverage_items=covered_items(brief, "sample-size assumptions")),
        make_slide(kind, "dag", "预设DAG同时连接混杂因素、暴露与结局", "method", "使用node/edge规范表达调整逻辑", "process_branch", [
            make_visual(kind, "dag", "dag", "confounding_plan", [variable_binding, protocol], {
                "nodes": nodes,
                "edges": edges,
                "adjustment_set": [zh_entity_label(row["variable"]) for row in prespecified_confounders],
                "review_note": "同期变量需单独进行因果角色审核：" + "、".join(zh_entity_label(row["variable"]) for row in same_window_variables),
                "explanation": "图中仅把明确位于暴露前的协变量作为预设混杂因素；同期管理/严重度变量不自动纳入调整集。",
            }, len(variables["rows"]), aggregation="semantic node/edge specification", entity_scope="prespecified exposure-outcome confounding graph",
                        edge_source_field="source", edge_target_field="target", label_field="label", sequential=False)
        ], claims, coverage_items=covered_items(brief, "dag", "confounding plan")),
        make_slide(kind, "tte_qc", "核心实验室培训与重复判读控制TTE测量误差", "method", "呈现TTE质量控制和图像质量变量", "evidence_ladder", [
            make_visual(kind, "tte_qc", "quality_control_ladder", "measurement_quality", [operations, risk_binding, variable_binding], {
                "steps": ["标准化采集", "核心实验室培训", "重复判读", "记录图像质量", "图像质量敏感性分析"],
                "status": ["计划", "计划", "计划", "预设变量", "预设分析"],
            }, len(risks["rows"]) + len(variables["rows"]), aggregation="quality-control synthesis", entity_scope="TTE acquisition and reading quality")
        ], claims, coverage_items=covered_items(brief, "tte quality control")),
        make_slide(kind, "analysis", "主分析报告调整OR，并预设中心与图像质量敏感性分析", "method", "呈现分析模型、输出和敏感性边界", "process_branch", [
            narrative_visual(kind, "analysis", "analysis_plan", protocol, {
                "nodes": ["编码暴露与结局", "检查缺失与事件数", "多变量Logistic回归", "报告调整OR与95%CI", "图像质量敏感性", "中心效应敏感性"],
                "branches": [[0, 1], [1, 2], [2, 3], [3, 4], [3, 5]],
                "warning": "所有分析均为预设计划，未执行结果分析。",
            }, "method")
        ], claims, coverage_items=covered_items(brief, "analysis plan")),
        make_slide(kind, "risk_matrix", "高概率与高影响风险需要优先缓解", "planning", "将所有风险定位到概率×影响二维矩阵", "risk_heatmap", [
            make_visual(kind, "risk_matrix", "risk_matrix", "operational_risk", [risk_binding], {"rows": risk_rows}, len(risk_rows),
                        x_field="probability_score", y_field="impact_score", label_field="risk", color_field="domain", entity_scope="all registered operational risks")
        ], claims, coverage_items=covered_items(brief, "operational risks")),
        make_slide(kind, "milestones", f"分析与报告阶段延续至{max(row['end_date'] for row in milestones['rows'])}", "planning", "按真实日期和持续时间绘制全部里程碑", "chart_led", [
            make_visual(kind, "milestones", "timeline_gantt", "project_plan", [milestone_binding], {"rows": milestones["rows"]}, len(milestones["rows"]),
                        time_field="start_date", label_field="milestone", group_field="status", unit="calendar date", entity_scope="all project milestones", end_time_field="end_date")
        ], claims, coverage_items=covered_items(brief, "milestones")),
        make_slide(kind, "no_results", "当前仅有方案与计划值，没有任何观察结果", "limitation", "用显式页面阻止协议结果化", "split_screen", [
            make_visual(kind, "no_results", "no_result_boundary", "protocol_boundary", [protocol, sample_binding], {
                "planned": ["目标入组与样本量假设", "评估时间窗", "分析模型与敏感性方案"],
                "not_available": ["事件率", "效应值", "趋势", "招募完成状态"],
            }, len(sample["rows"]), aggregation="protocol boundary synthesis", entity_scope="protocol-only evidence")
        ], claims, ["CLM-PROTO-005"]),
        make_slide(kind, "conclusion", "方案以明确时间、质量控制与预设分析保护可解释性", "conclusion", "综合研究方案的可执行要点与边界", "conclusion_synthesis", [
            make_visual(kind, "conclusion", "takeaway_synthesis", "conclusion", [schedule_binding, sample_binding, risk_binding, protocol], {
                "takeaways": ["0–6小时定义主要暴露", "72小时定义主要结局", f"计划入组{target['value']}例", "结果与伦理/资金信息仍待未来确认"],
            }, len(schedule["rows"]) + len(sample["rows"]) + len(risks["rows"]), aggregation="protocol synthesis", entity_scope="registered protocol components")
        ], claims, [item["claim_id"] for item in claims]),
        make_slide(kind, "appendix_review", f"{len(conflicts)}项冲突与{len(unresolved)}项待确认内容保留审计", "appendix", "集中呈现冲突与未解决事项", "appendix_audit", [], claims,
                   conflicts=conflicts, unresolved=unresolved, appendix=True, coverage_items=covered_items(brief, "unresolved ethics", "funding")),
    ]
    return slides


def resolve_projects(project: Path | None, suite: Path | None) -> list[Path]:
    if suite is not None:
        projects = [path for path in sorted(suite.glob(PROJECT_PATTERN)) if (path / "brief" / "presentation_brief.yaml").is_file()]
        if not projects:
            raise ValueError("NO_PROJECTS_DETECTED")
        return projects
    root = (project or REPO).resolve()
    briefs = list(root.rglob("presentation_brief.yaml"))
    if len(briefs) > 1:
        raise ValueError("MULTIPLE_PROJECTS_DETECTED")
    if len(briefs) == 0:
        raise ValueError("NO_PROJECT_BRIEF_DETECTED")
    return [briefs[0].parent.parent]


def build_slides(
    project: Path, kind: str, brief: dict[str, Any], sources: list[dict[str, Any]], datasets: dict[str, dict[str, Any]],
    claims: list[dict[str, Any]], conflicts: list[dict[str, Any]], unresolved: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if brief.get("language") != "zh-CN":
        raise ValueError(f"UNSUPPORTED_GENERATION_LANGUAGE:{brief.get('language')}")
    if kind == "target_trial":
        slides = build_target_trial_slides(brief, sources, datasets, claims, conflicts, unresolved)
    elif kind == "single_cell":
        slides = build_single_cell_slides(project, brief, sources, datasets, claims, conflicts, unresolved)
    elif kind == "meta_analysis":
        slides = build_meta_slides(project, brief, sources, datasets, claims, conflicts, unresolved)
    else:
        slides = build_protocol_slides(brief, sources, datasets, claims, conflicts, unresolved)
    for slide in slides:
        slide["language"] = brief["language"]
        for visual in slide["visual_specs"]:
            visual["slide_id"] = slide["slide_id"]
    return slides


def coverage_contracts(brief: dict[str, Any], slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    contracts: list[dict[str, Any]] = []
    for item in brief.get("must_include", []):
        matched = [slide for slide in slides if item in slide.get("coverage_items", [])]
        contracts.append({
            "must_include_item": item,
            "expected_claim": "source-bound coverage",
            "expected_visual": "|".join(sorted({visual["visual_type"] for slide in matched for visual in slide["visual_specs"]})) or "cover warning",
            "expected_row_count": sum(visual["expected_row_count"] for slide in matched for visual in slide["visual_specs"]),
            "target_slide_count": brief.get("target_slide_count", ""),
            "rendered_slide": "|".join(slide["slide_id"] for slide in matched),
            "appendix_location": "|".join(slide["slide_id"] for slide in matched if slide["appendix_status"] == "appendix"),
            "omitted": not bool(matched),
            "omission_reason": "" if matched else "No SlideSpec declared coverage",
            "manual_decision_required": not bool(matched),
        })
    return contracts


def derived_artifacts(
    out: Path,
    slides: list[dict[str, Any]],
    claims: list[dict[str, Any]],
    conflicts: list[dict[str, Any]],
    unresolved: list[dict[str, Any]],
    coverage: list[dict[str, Any]],
) -> None:
    slide_manifest = []
    source_rows = []
    claim_rows = []
    notes_lines = ["# Speaker notes", ""]
    for index, slide in enumerate(slides, 1):
        visual_ids = [visual["visual_id"] for visual in slide["visual_specs"]]
        def visual_ids_for_binding(binding: dict[str, Any]) -> list[str]:
            result = []
            expected_fields = set(binding.get("fields_used", []))
            for visual in slide["visual_specs"]:
                for visual_binding in visual["source_bindings"]:
                    if visual_binding["source_id"] != binding["source_id"]:
                        continue
                    actual_fields = set(visual_binding.get("fields_used", []))
                    if (not expected_fields and not actual_fields) or expected_fields <= actual_fields:
                        result.append(visual["visual_id"])
                        break
            return result
        slide_manifest.append({
            "slide_number": index,
            "slide_id": slide["slide_id"],
            "slide_role": slide["slide_role"],
            "title": slide["title"],
            "language": slide["language"],
            "layout_family": slide["layout_family"],
            "design_profile": slide["design_profile"],
            "claim_ids": "|".join(item["claim_id"] for item in slide["claim_bindings"]),
            "visual_ids": "|".join(visual_ids),
            "appendix_status": slide["appendix_status"],
            "review_flags": "|".join(slide["review_flags"]),
        })
        notes_lines.extend([f"## {slide['slide_id']} — {slide['title']}", "", "[Sources]"])
        for binding in slide["source_bindings"]:
            source_rows.append({
                "slide_id": slide["slide_id"],
                "visual_ids": "|".join(visual_ids_for_binding(binding)),
                "source_id": binding["source_id"],
                "source_file": binding["source_file"],
                "source_location": binding["source_location"],
                "fields_used": "|".join(binding["fields_used"]),
                "row_filter": binding["row_filter"],
                "aggregation": binding["aggregation"],
                "canonical_status": binding["canonical_status"],
            })
            notes_lines.append(
                f"- {binding['source_id']} | {binding['source_file']} | {binding['source_location']} | "
                f"fields={','.join(binding['fields_used'])} | filter={binding['row_filter']} | aggregation={binding['aggregation']}"
            )
        notes_lines.append("[Claims]")
        for item in slide["claim_bindings"]:
            notes_lines.append(f"- {item['claim_id']} | {item['claim_text']} | boundary={item['wording_boundary']}")
            for binding in item["source_bindings"]:
                claim_rows.append({
                    "claim_id": item["claim_id"],
                    "slide_id": slide["slide_id"],
                    "visual_ids": "|".join(visual_ids_for_binding(binding)),
                    "claim_text": item["claim_text"],
                    "expected_value": item["expected_value"],
                    "source_id": binding["source_id"],
                    "source_file": binding["source_file"],
                    "source_location": binding["source_location"],
                    "fields_used": "|".join(binding["fields_used"]),
                    "wording_boundary": item["wording_boundary"],
                    "canonical_status": item["canonical_status"],
                })
        if slide["conflict_bindings"]:
            notes_lines.append("[Conflicts]")
            notes_lines.extend(f"- {item['conflict_id']} | {item['conflicting_value']} -> {item['canonical_value']}" for item in slide["conflict_bindings"])
        if slide["unresolved_bindings"]:
            notes_lines.append("[Unresolved]")
            notes_lines.extend(f"- {item['marker']} | INFORMATION_REQUIRED" for item in slide["unresolved_bindings"])
        notes_lines.append("")
    write_csv(out / "slide_manifest.csv", slide_manifest)
    write_csv(out / "source_binding_map.csv", source_rows)
    write_csv(out / "claim_source_map.csv", claim_rows)
    write_csv(out / "narrative_coverage_report.csv", coverage)
    (out / "speaker_notes.md").write_text("\n".join(notes_lines), encoding="utf-8")
    checklist = ["# Manual review checklist", ""]
    checklist.extend(f"- [ ] Conflict {item['conflict_id']}: {item['conflicting_value']} -> {item['canonical_value']}" for item in conflicts)
    checklist.extend(f"- [ ] INFORMATION_REQUIRED: {item['marker']}" for item in unresolved)
    checklist.extend([
        "- [ ] Confirm every slide-level source footer against speaker notes.",
        "- [ ] Confirm the deck contains no real clinical claims.",
        "- [ ] Complete the blinded visual review form.",
        "- [ ] User grants final scientific approval.",
    ])
    (out / "manual_review_checklist.md").write_text("\n".join(checklist), encoding="utf-8")
    review_rows = [{
        "anonymous_deck_code": "",
        "scientific_clarity_20": "",
        "visual_semantic_correctness_20": "",
        "hierarchy_15": "",
        "composition_15": "",
        "chart_professionalism_10": "",
        "consistency_10": "",
        "readability_5": "",
        "presentation_readiness_5": "",
        "total_100": "",
        "reviewer": "",
        "review_date": "",
        "comments": "",
    }]
    write_csv(out / "manual_visual_review_form.csv", review_rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Academic PPT Workflow v2.4 canonical design runner")
    parser.add_argument("--project", type=Path, help="Single project root (default input mode)")
    parser.add_argument("--batch-regression-suite", type=Path, help="Explicit isolated suite mode")
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--node", required=True)
    parser.add_argument("--node-modules", type=Path, help="Project/bundled Node module directory; exported only to the child renderer")
    args = parser.parse_args()
    try:
        projects = resolve_projects(args.project, args.batch_regression_suite)
    except ValueError as exc:
        print(str(exc))
        return 2
    run_root = REPO / "benchmark" / "v2_4" / "runs" / args.run_id
    stage_root = REPO / "staging" / "v2_4" / args.run_id
    if run_root.exists() or stage_root.exists():
        raise RuntimeError(f"Refusing to overwrite existing run: {args.run_id}")
    run_root.mkdir(parents=True)
    stage_root.mkdir(parents=True)
    deck_index = []
    for project in projects:
        started = time.perf_counter()
        stage = stage_root / project.name
        out = run_root / project.name
        stage.mkdir()
        out.mkdir()
        brief, sources, datasets, texts = discover_project(project)
        kind = infer_kind(datasets)
        claims = build_claims(kind, datasets, sources)
        conflicts = detect_conflicts(project.name, sources, datasets, texts)
        unresolved = detect_unresolved(sources, texts)
        slides = build_slides(project, kind, brief, sources, datasets, claims, conflicts, unresolved)
        validation_errors = [
            {"slide_id": slide["slide_id"], "error": error}
            for slide in slides for error in [*validate_slide_spec(slide), *validate_slide_semantics(slide)]
        ]
        if validation_errors:
            dump_json(out / "schema_validation_errors.json", validation_errors)
            raise RuntimeError(f"Schema validation failed for {project.name}: {validation_errors[:3]}")
        coverage = coverage_contracts(brief, slides)
        if any(item["omitted"] for item in coverage):
            write_csv(out / "narrative_coverage_report.csv", coverage)
            raise RuntimeError(f"Mandatory narrative coverage failed for {project.name}")
        story_graph = {
            "communication_job": f"演示结束时，{brief['audience']}应理解：{brief['key_message']}",
            "presentation_type": brief["presentation_type"],
            "narrative_mode": brief.get("narrative_mode", "scientific_problem"),
            "language": brief["language"],
            "target_slide_count": brief["target_slide_count"],
            "must_include": brief.get("must_include", []),
            "must_exclude": brief.get("must_exclude", []),
        }
        visual_specs = [visual for slide in slides for visual in slide["visual_specs"]]
        graph = {
            "graph_version": "2.4",
            "project_key": project.name,
            "kind": kind,
            "brief": brief,
            "source_registry": sources,
            "canonical_evidence_graph": {"claims": claims, "conflicts": conflicts, "unresolved": unresolved},
            "story_graph": story_graph,
            # SlideSpec.visual_specs is the sole canonical VisualSpec state.
            # The flat visual_spec.json below is a derived review artifact.
            "slide_specs": slides,
            "backend": "native_pptxgenjs",
            "external_skill_used": False,
            "ground_truth_used_for_generation": False,
        }
        write_csv(out / "source_manifest.csv", sources)
        write_csv(out / "evidence_registry.csv", claims)
        write_csv(out / "conflict_registry.csv", conflicts)
        write_csv(out / "unresolved_registry.csv", unresolved)
        dump_json(out / "canonical_object_graph.json", graph)
        dump_json(out / "slide_spec.json", slides)
        dump_json(out / "visual_spec.json", visual_specs)
        dump_json(out / "story_graph.json", story_graph)
        derived_artifacts(out, slides, claims, conflicts, unresolved, coverage)
        spec_path = stage / "canonical_object_graph.json"
        dump_json(spec_path, graph)
        pptx = out / f"{project.name}_v2_4.pptx"
        env = dict(os.environ)
        if args.node_modules:
            env["NODE_PATH"] = str(args.node_modules.resolve())
        completed = subprocess.run(
            [args.node, str(REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs"), str(spec_path), str(pptx)],
            cwd=REPO,
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=240,
        )
        (out / "generation.log").write_text(completed.stdout + "\n" + completed.stderr, encoding="utf-8")
        if completed.returncode != 0 or not pptx.is_file():
            raise RuntimeError(f"Deck generation failed for {project.name}: {completed.stderr}")
        object_manifest_path = pptx.with_suffix(".object_manifest.json")
        if not object_manifest_path.is_file():
            raise RuntimeError(f"Renderer did not create object manifest for {project.name}")
        object_manifest = json.loads(object_manifest_path.read_text(encoding="utf-8"))
        actual_by_visual = {item["visual_id"]: item for item in object_manifest}
        row_validation = []
        for slide in slides:
            for visual in slide["visual_specs"]:
                artifact = actual_by_visual.get(visual["visual_id"], {})
                actual = artifact.get("rendered_row_count")
                actual_keys = artifact.get("rendered_row_keys", [])
                visual["rendered_row_count"] = actual
                visual["rendered_row_keys"] = actual_keys
                passed = (
                    actual == visual["expected_row_count"]
                    and set(actual_keys) == set(visual["expected_row_keys"])
                ) or bool(visual["omitted_rows"])
                row_validation.append({
                    "slide_id": slide["slide_id"],
                    "visual_id": visual["visual_id"],
                    "visual_type": visual["visual_type"],
                    "expected_row_count": visual["expected_row_count"],
                    "rendered_row_count": actual,
                    "expected_row_keys": "|".join(visual["expected_row_keys"]),
                    "rendered_row_keys": "|".join(actual_keys),
                    "omitted_rows": len(visual["omitted_rows"]),
                    "status": "PASS" if passed else "FAIL",
                })
        write_csv(out / "rendered_row_validation.csv", row_validation)
        if any(item["status"] == "FAIL" for item in row_validation):
            raise RuntimeError(f"Rendered row coverage failed for {project.name}")
        visual_specs = [visual for slide in slides for visual in slide["visual_specs"]]
        dump_json(out / "canonical_object_graph.json", graph)
        dump_json(out / "slide_spec.json", slides)
        dump_json(out / "visual_spec.json", visual_specs)
        deck_index.append({
            "project_key": project.name,
            "arm": "V24",
            "kind": kind,
            "pptx_path": str(pptx.resolve()),
            "output_dir": str(out.resolve()),
            "slide_count": len(slides),
            "target_slide_count": brief["target_slide_count"],
            "claim_count": len(claims),
            "conflict_count": len(conflicts),
            "unresolved_count": len(unresolved),
            "layout_family_count": len({slide["layout_family"] for slide in slides}),
            "generation_seconds": round(time.perf_counter() - started, 3),
        })
    write_csv(run_root / "deck_index.csv", deck_index)
    dump_json(run_root / "generation_complete.json", {
        "run_id": args.run_id,
        "completed_at_utc": datetime.now(timezone.utc).isoformat(),
        "input_mode": "batch_regression_suite" if args.batch_regression_suite else "single_project",
        "project_count": len(projects),
        "ground_truth_used_for_generation": False,
        "external_skill_used": False,
        "backend": "native_pptxgenjs",
    })
    print(f"PROJECTS={len(deck_index)}")
    print(f"RUN_ROOT={run_root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
