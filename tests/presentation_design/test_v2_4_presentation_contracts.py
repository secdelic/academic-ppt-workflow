from __future__ import annotations

import csv
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any


REPO = Path(__file__).resolve().parents[2]
V24 = REPO / "scripts" / "v2_4"
sys.path.insert(0, str(V24))

from model import (  # noqa: E402
    asset_binding,
    source_binding,
    validate_slide_semantics,
    validate_slide_spec,
    validate_visual_semantics,
    validate_visual_spec,
    visible_cjk_ratio,
)
from run_v2_4 import coverage_contracts, derived_artifacts, make_slide, make_visual  # noqa: E402
from finalize_v2_4 import canonical_source_matches  # noqa: E402


def dataset(role: str, columns: list[str], rows: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """Create a project-neutral canonical dataset fixture."""
    return {
        "role": role,
        "source_id": f"SRC-{role.upper().replace('_', '-')}",
        "source_file": f"input/data/{role}.csv",
        "absolute_path": f"synthetic/{role}.csv",
        "columns": columns,
        "rows": rows or [],
    }


def binding_for(role: str, columns: list[str], rows: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    return source_binding(dataset(role, columns, rows), columns)


def visual_fixture(
    visual_type: str,
    *,
    role: str = "generic",
    columns: list[str] | None = None,
    rows: list[dict[str, Any]] | None = None,
    payload: dict[str, Any] | None = None,
    kind: str = "meta_analysis",
    expected: int | None = None,
    rendered: int | None = None,
    **fields: Any,
) -> dict[str, Any]:
    columns = columns or ["label", "value"]
    rows = rows or []
    if expected and not rows and payload is None:
        rows = [{columns[0]: f"row-{index + 1}", columns[-1]: index + 1} for index in range(expected)]
    if expected and payload is not None and isinstance(payload.get("rows"), list) and not payload["rows"]:
        payload = {**payload, "rows": [{columns[0]: f"row-{index + 1}", columns[-1]: index + 1} for index in range(expected)]}
    binding = binding_for(role, columns, rows)
    visual = make_visual(
        kind,
        "generic-contract",
        visual_type,
        "test-contract",
        [binding],
        payload or {"rows": rows},
        len(rows) if expected is None else expected,
        rendered_row_count=rendered,
        **fields,
    )
    if rendered is not None:
        visual["rendered_row_keys"] = visual["expected_row_keys"][:rendered]
    return visual


def slide_fixture(
    title: str,
    *,
    key: str = "generic-slide",
    layout_family: str = "chart_led",
    visuals: list[dict[str, Any]] | None = None,
    coverage_items: list[str] | None = None,
    card_grid: bool = False,
) -> dict[str, Any]:
    return make_slide(
        "meta_analysis",
        key,
        title,
        "result",
        "表达一个来源绑定且可核查的核心信息",
        layout_family,
        visuals or [],
        [],
        coverage_items=coverage_items,
        card_grid=card_grid,
    )


def semantic_codes(spec: dict[str, Any]) -> set[str]:
    return set(validate_visual_semantics(spec))


def parse_flat_yaml(path: Path) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        value = value.strip()
        if value.lower() in {"true", "false"}:
            result[key] = value.lower() == "true"
        elif re.fullmatch(r"-?\d+(?:\.\d+)?", value):
            result[key] = float(value) if "." in value else int(value)
        else:
            result[key] = value
    return result


def generic_title_rate(titles: list[str]) -> float:
    generic = {
        "背景", "方法", "结果", "讨论", "结论", "background", "methods",
        "results", "discussion", "conclusion", "slide", "overview",
    }
    count = sum(title.strip().lower() in generic for title in titles)
    return count / len(titles) if titles else 0.0


def duplicate_semantic_signatures(slides: list[dict[str, Any]]) -> set[tuple[str, str, tuple[str, ...]]]:
    seen: set[tuple[str, str, tuple[str, ...]]] = set()
    duplicates: set[tuple[str, str, tuple[str, ...]]] = set()
    for slide in slides:
        signature = (
            slide["title"].strip(),
            slide["narrative_purpose"].strip(),
            tuple(sorted(visual["scientific_role"] for visual in slide["visual_specs"])),
        )
        if signature in seen:
            duplicates.add(signature)
        seen.add(signature)
    return duplicates


def protocol_result_violations(slides: list[dict[str, Any]]) -> list[str]:
    forbidden_visuals = {
        "event_rate_plot", "forest_plot", "subgroup_forest", "risk_curve",
        "effect_estimate", "observed_trend",
    }
    forbidden_phrases = {"观察到", "结果显示", "显著降低", "显著升高", "observed", "resulted in"}
    violations: list[str] = []
    for slide in slides:
        for visual in slide["visual_specs"]:
            if visual["visual_type"] in forbidden_visuals:
                violations.append(f"visual:{visual['visual_type']}")
        text = f"{slide['title']} {slide['key_message']}".lower()
        if any(phrase in text for phrase in forbidden_phrases):
            violations.append(f"wording:{slide['slide_id']}")
    return violations


class PresentationDesignContracts(unittest.TestCase):
    def test_data_contract_id_sensitive_to_semantic_binding(self):
        data = dataset(
            "semantic_contract",
            ["entity", "estimate", "outcome", "stratum"],
            [{"entity": "甲", "estimate": "1.2", "outcome": "结局甲", "stratum": "层一"}],
        )

        def build(
            *,
            fields: list[str] | None = None,
            row_filter: str = "all rows",
            aggregation: str = "none",
            entity_scope: str = "all entities",
            outcome: str = "结局甲",
        ) -> dict[str, Any]:
            binding = source_binding(
                data,
                fields or ["entity", "estimate"],
                "row 1",
                row_filter,
                aggregation,
            )
            return make_visual(
                "meta_analysis", "same-semantic-slot", "native_bar_chart", "effect", [binding],
                {"rows": data["rows"]}, 1,
                row_filter=row_filter,
                aggregation=aggregation,
                entity_scope=entity_scope,
                outcome=outcome,
                x_field="entity",
                y_field="estimate",
            )

        variants = [
            build(),
            build(fields=["entity", "estimate", "outcome"]),
            build(row_filter="stratum == 层一"),
            build(aggregation="mean by entity"),
            build(entity_scope="prespecified subgroup"),
            build(outcome="结局乙"),
        ]
        contract_ids = {visual["data_contract_id"] for visual in variants}
        self.assertEqual(len(contract_ids), len(variants))

    def test_exact_field_binding_requires_expected_subset(self):
        data = dataset(
            "shared_source",
            ["entity", "estimate", "stratum"],
            [{"entity": "甲", "estimate": "1.2", "stratum": "层一"}],
        )
        claim_binding = source_binding(data, ["entity", "estimate"], "row 1")
        visual_binding = source_binding(data, ["entity", "stratum"], "row 1")
        visual = make_visual(
            "meta_analysis", "partial-field-overlap", "native_bar_chart", "stratification",
            [visual_binding], {"rows": data["rows"]}, 1,
            x_field="entity", y_field="stratum",
        )
        claim = {
            "claim_id": "CLM-FIELD-SUBSET",
            "claim_text": "估计值需要entity与estimate两个字段共同支持",
            "expected_value": "1.2",
            "wording_boundary": "不得以仅共享entity字段的视觉替代精确绑定。",
            "canonical_status": "canonical",
            "source_bindings": [claim_binding],
        }
        slide = make_slide(
            "meta_analysis", "field-subset", "字段子集决定claim与visual的精确映射", "result",
            "验证任意字段交集不构成完整来源绑定", "chart_led", [visual], [claim],
            claim_ids=[claim["claim_id"]],
        )
        with tempfile.TemporaryDirectory(dir=REPO / "staging") as tmp:
            out = Path(tmp)
            derived_artifacts(out, [slide], [claim], [], [], [])
            with (out / "claim_source_map.csv").open("r", encoding="utf-8-sig", newline="") as handle:
                claim_rows = list(csv.DictReader(handle))
        self.assertEqual(len(claim_rows), 1)
        self.assertEqual(claim_rows[0]["visual_ids"], "")

    def test_risk_severity_bilingual_warning_mapping(self):
        source = (REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs").read_text(encoding="utf-8")
        body = source.split("function renderMatrix", 1)[1].split("function renderBranch", 1)[0]
        self.assertIn('severity.includes("high")', body)
        self.assertIn('severity.includes("高")', body)
        self.assertRegex(body, r"severity\.includes\(\"high\"\).*severity\.includes\(\"高\"\).*SHARED\.warning")

    def test_protocol_timeline_neighbor_lane_collision_safety(self):
        source = (REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs").read_text(encoding="utf-8")
        body = source.split("function renderAssessmentTimeline", 1)[1].split("function renderVariables", 1)[0]
        self.assertIn("sort((a, b) => Number(a.row.position_hours) - Number(b.row.position_hours))", body)
        self.assertGreaterEqual(body.count("lastRight: -Infinity"), 4)
        self.assertIn("labelX >= candidate.lastRight + 0.08", body)
        self.assertIn("Assessment timeline label budget exceeded", body)
        self.assertRegex(body, r"Math\.max\(0\.76,\s*Math\.min\(11\.76,\s*x - labelW / 2\)\)")

    def test_volcano_top_k_label_limit(self):
        source = (REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs").read_text(encoding="utf-8")
        body = source.split("function renderFacetedVolcano", 1)[1].split("function renderEnrichment", 1)[0]
        self.assertIn("const facets = [...new Set", body)
        self.assertRegex(body, r"const labelRows = new Set\([\s\S]*?\.slice\(0, 2\)\)")
        self.assertIn("if (labelRows.has(item))", body)
        self.assertIn("placedLabels", body)
        # The current canonical VisualSpec declares the same per-facet cap;
        # a renderer-side label may not exceed this declared top-k budget.
        data = dataset(
            "volcano",
            ["gene", "cell_type", "log2fc", "FDR"],
            [{"gene": f"G{i}", "cell_type": "细胞甲", "log2fc": str(i / 10), "FDR": str((i + 1) / 1000)} for i in range(8)],
        )
        binding = source_binding(data, data["columns"], "all rows")
        visual = make_visual(
            "single_cell", "volcano-top-k", "faceted_volcano", "computational_association",
            [binding], {"rows": data["rows"], "top_label_policy": "top 2 labels per facet"},
            len(data["rows"]), x_field="log2fc", y_field="FDR", facet_field="cell_type",
            label_field="gene", top_k_policy="top 2 labels per facet; all points retained",
        )
        declared_top_k = int(re.search(r"top\s+(\d+)", visual["top_k_policy"], re.I).group(1))
        self.assertEqual(declared_top_k, 2)

    def test_prerender_row_count_is_unknown_not_pass(self):
        visual = visual_fixture(
            "native_bar_chart",
            columns=["entity", "value"],
            rows=[{"entity": "甲", "value": "1"}, {"entity": "乙", "value": "2"}],
            x_field="entity",
            y_field="value",
        )
        self.assertIsNone(visual["rendered_row_count"])
        self.assertNotEqual(visual["rendered_row_count"], visual["expected_row_count"])
        self.assertFalse(
            visual["rendered_row_count"] == visual["expected_row_count"]
            or bool(visual["omitted_rows"])
        )
        # Schema validation permits a pre-render unknown; only renderer evidence
        # may subsequently turn this state into PASS.
        self.assertNotIn("silent_truncation", validate_visual_spec(visual))

    def test_source_row_count_is_distinct_from_expected_visual_rows(self):
        source_rows = 12
        displayed_items = [[f"条目{i}", "摘要"] for i in range(5)]
        visual = visual_fixture(
            "specification_matrix",
            columns=["field", "value"],
            payload={"items": displayed_items},
            expected=source_rows,
        )
        self.assertEqual(visual["source_row_count"], source_rows)
        self.assertEqual(visual["expected_row_count"], len(displayed_items))
        self.assertIsNone(visual["rendered_row_count"])

    def test_appendix_registry_bindings_are_nonempty(self):
        conflict = {
            "conflict_id": "CONFLICT-GENERIC",
            "source_id": "SRC-CONFLICT",
            "source_file": "input/documents/older_version.docx",
            "source_location": "paragraph 7",
            "conflicting_value": "旧值",
            "canonical_value": "冻结值",
        }
        unresolved = {
            "item_id": "UNRES-GENERIC",
            "marker": "[INFORMATION_REQUIRED]",
            "source_id": "SRC-UNRESOLVED",
            "source_file": "input/documents/protocol.docx",
            "source_location": "ethics section",
        }
        slide = make_slide(
            "meta_analysis", "appendix-registry", "冲突与待确认事项保留审计", "appendix",
            "让人工审核能够追溯冲突与未解决来源", "appendix_audit", [], [],
            conflicts=[conflict], unresolved=[unresolved], appendix=True,
        )
        self.assertTrue(slide["source_bindings"])
        self.assertEqual(
            {binding["source_id"] for binding in slide["source_bindings"]},
            {"SRC-CONFLICT", "SRC-UNRESOLVED"},
        )
        self.assertTrue(all(binding["source_file"] and binding["source_location"] for binding in slide["source_bindings"]))

    def test_source_and_claim_map_visual_ids_are_binding_exact(self):
        first_data = dataset("endpoint_a", ["entity", "estimate"], [{"entity": "甲", "estimate": "1.2"}])
        second_data = dataset("endpoint_b", ["category", "rate"], [{"category": "乙", "rate": "25.0"}])
        first_binding = source_binding(first_data, ["entity", "estimate"], "row 1")
        second_binding = source_binding(second_data, ["category", "rate"], "row 1")
        first_visual = make_visual(
            "meta_analysis", "binding-a", "native_bar_chart", "effect", [first_binding],
            {"rows": first_data["rows"]}, 1, x_field="entity", y_field="estimate",
        )
        second_visual = make_visual(
            "meta_analysis", "binding-b", "native_bar_chart", "rate", [second_binding],
            {"rows": second_data["rows"]}, 1, x_field="category", y_field="rate",
        )
        claim = {
            "claim_id": "CLM-GENERIC",
            "claim_text": "来源甲支持该效应估计",
            "expected_value": "1.2",
            "wording_boundary": "仅表述合成关联。",
            "canonical_status": "canonical",
            "source_bindings": [first_binding],
        }
        slide = make_slide(
            "meta_analysis", "binding-map", "两个视觉使用各自的精确来源", "result",
            "验证派生映射不会把同页全部visual_id复制给每个来源", "chart_led",
            [first_visual, second_visual], [claim], claim_ids=[claim["claim_id"]],
        )
        with tempfile.TemporaryDirectory(dir=REPO / "staging") as tmp:
            out = Path(tmp)
            derived_artifacts(out, [slide], [claim], [], [], [])
            with (out / "source_binding_map.csv").open("r", encoding="utf-8-sig", newline="") as handle:
                source_rows = list(csv.DictReader(handle))
            with (out / "claim_source_map.csv").open("r", encoding="utf-8-sig", newline="") as handle:
                claim_rows = list(csv.DictReader(handle))
        source_visuals = {row["source_id"]: row["visual_ids"] for row in source_rows}
        self.assertEqual(source_visuals[first_data["source_id"]], first_visual["visual_id"])
        self.assertEqual(source_visuals[second_data["source_id"]], second_visual["visual_id"])
        self.assertEqual(claim_rows[0]["visual_ids"], first_visual["visual_id"])
        self.assertNotIn(second_visual["visual_id"], claim_rows[0]["visual_ids"])

    def test_renderer_emits_independent_render_evidence(self):
        source = (REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs").read_text(encoding="utf-8")
        self.assertIn("const renderEvidence = {}", source)
        self.assertIn("created_object_count", source)
        self.assertIn("renderVisual(slide, visual, spec)", source)
        self.assertIn("createdObjects", source)
        self.assertNotIn("actualRenderedRowCount", source)
        self.assertRegex(source, r"rendered_row_count:\s*renderEvidence\[visual\.visual_id\]\?\.renderedRows")

    def test_grouped_composition_uses_data_driven_group_names(self):
        source = (REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs").read_text(encoding="utf-8")
        body = source.split("function renderComposition", 1)[1].split("function renderFacetedVolcano", 1)[0]
        self.assertIn("new Set(rows.map(row => row.group))", body)
        self.assertIn("const [groupA, groupB] = groups", body)
        self.assertNotRegex(body, r"\b(?:Control|SIMD|Treatment|Placebo)\b")

    def test_dag_renderer_supports_arbitrary_node_ids(self):
        payload = {
            "nodes": [
                {"id": "baseline-severity::α", "label": "基线严重度", "role": "confounder"},
                {"id": "exposure/window-0h", "label": "暴露", "role": "exposure"},
                {"id": "outcome@72h", "label": "结局", "role": "outcome"},
            ],
            "edges": [
                {"source": "baseline-severity::α", "target": "exposure/window-0h"},
                {"source": "baseline-severity::α", "target": "outcome@72h"},
                {"source": "exposure/window-0h", "target": "outcome@72h"},
            ],
        }
        visual = visual_fixture("dag", columns=["node", "edge"], payload=payload, expected=3)
        self.assertEqual(semantic_codes(visual), set())
        source = (REPO / "scripts" / "v2_4" / "render_v2_4_deck.mjs").read_text(encoding="utf-8")
        body = source.split("function renderDAG", 1)[1].split("function renderRiskMatrix", 1)[0]
        self.assertIn("positions[node.id]", body)
        self.assertIn("positions[edge.source]", body)
        self.assertIn("positions[edge.target]", body)
        self.assertNotRegex(body, r"positions\.(?:C|E|Y|exposure|outcome)")

    def test_slide_specific_source_binding(self):
        first = dataset("analysis_a", ["entity", "estimate"], [{"entity": "A", "estimate": "1.2"}])
        second = dataset("analysis_b", ["entity", "estimate"], [{"entity": "B", "estimate": "1.5"}])
        first_binding = source_binding(first, ["entity", "estimate"], "row 1")
        second_binding = source_binding(second, ["entity", "estimate"], "row 1")
        first_visual = make_visual("meta_analysis", "a", "native_bar_chart", "comparison", [first_binding], {"rows": first["rows"]}, 1, x_field="entity", y_field="estimate")
        second_visual = make_visual("meta_analysis", "b", "native_bar_chart", "comparison", [second_binding], {"rows": second["rows"]}, 1, x_field="entity", y_field="estimate")
        first_slide = slide_fixture("分析A的估计值高于参考线", key="a", visuals=[first_visual])
        second_slide = slide_fixture("分析B的估计值高于参考线", key="b", visuals=[second_visual])
        self.assertEqual({item["source_id"] for item in first_slide["source_bindings"]}, {first["source_id"]})
        self.assertEqual({item["source_id"] for item in second_slide["source_bindings"]}, {second["source_id"]})
        self.assertNotEqual(first_slide["source_bindings"], second_slide["source_bindings"])

    def test_language_contract_zh_cn(self):
        slide = slide_fixture("预设研究问题决定图表结构")
        self.assertEqual(validate_slide_spec(slide), [])
        visible_text = [slide["title"], slide["key_message"], slide["narrative_purpose"]]
        self.assertGreaterEqual(visible_cjk_ratio(visible_text), 0.90)
        self.assertEqual(slide["language"], "zh-CN")

    def test_title_entity_scope(self):
        rows = [
            {"pathway": "通路甲", "cell_type": "巨噬细胞", "nes": "1.8"},
            {"pathway": "通路乙", "cell_type": "心肌细胞", "nes": "-1.4"},
        ]
        visual = visual_fixture(
            "diverging_enrichment", role="pathway", columns=list(rows[0]), rows=rows,
            x_field="nes", facet_field="cell_type", label_field="pathway", reference_value=0,
        )
        mismatched = slide_fixture("巨噬细胞通路呈双向变化", visuals=[visual])
        matched = slide_fixture("不同细胞类型的通路呈双向变化", key="matched", visuals=[visual])
        self.assertTrue(any("title_entity_scope_mismatch" in error for error in validate_slide_semantics(mismatched)))
        self.assertFalse(any("title_entity_scope_mismatch" in error for error in validate_slide_semantics(matched)))

    def test_facet_consistency(self):
        rows = [
            {"pathway": "上调通路", "cell_type": "细胞甲", "nes": "1.2"},
            {"pathway": "下调通路", "cell_type": "细胞乙", "nes": "-1.1"},
        ]
        visual = visual_fixture(
            "diverging_enrichment", role="pathway", columns=list(rows[0]), rows=rows,
            x_field="nes", facet_field="cell_type", label_field="pathway", reference_value=0,
        )
        self.assertNotIn("enrichment_facets_missing", semantic_codes(visual))
        collapsed = dict(visual)
        collapsed["payload"] = {"rows": [dict(rows[0]), {**rows[1], "cell_type": "细胞甲"}]}
        self.assertIn("enrichment_facets_missing", semantic_codes(collapsed))

    def test_direction_sign_preserved(self):
        rows = [
            {"pathway": "正向", "cell_type": "细胞甲", "nes": "2.4"},
            {"pathway": "负向", "cell_type": "细胞乙", "nes": "-1.7"},
        ]
        visual = visual_fixture(
            "diverging_enrichment", role="pathway", columns=list(rows[0]), rows=rows,
            x_field="nes", facet_field="cell_type", label_field="pathway", reference_value=0,
        )
        rendered_values = [float(row[visual["x_field"]]) for row in visual["payload"]["rows"]]
        self.assertEqual(rendered_values, [2.4, -1.7])
        self.assertNotIn("diverging_sign_not_represented", semantic_codes(visual))

    def test_independent_network_not_sequence(self):
        rows = [
            {"sender": "A", "receiver": "B", "interaction": "L1-R1", "weight": "0.8"},
            {"sender": "C", "receiver": "D", "interaction": "L2-R2", "weight": "0.5"},
        ]
        visual = visual_fixture(
            "communication_network", role="network", columns=list(rows[0]), rows=rows,
            edge_source_field="sender", edge_target_field="receiver", label_field="interaction",
            color_field="weight", sequential=False,
        )
        self.assertEqual(validate_visual_spec(visual), [])
        self.assertNotIn("independent_network_marked_sequence", semantic_codes(visual))
        invalid = dict(visual, sequential=True)
        self.assertIn("network_marked_sequential", validate_visual_spec(invalid))
        self.assertIn("independent_network_marked_sequence", semantic_codes(invalid))

    def test_expected_vs_rendered_rows(self):
        visual = visual_fixture("native_bar_chart", expected=4, rendered=4)
        self.assertEqual(visual["expected_row_count"], visual["rendered_row_count"])
        self.assertNotIn("silent_truncation", validate_visual_spec(visual))

    def test_no_silent_truncation(self):
        invalid = visual_fixture("native_bar_chart", expected=5, rendered=3)
        self.assertIn("silent_truncation", validate_visual_spec(invalid))
        declared = dict(invalid, omitted_rows=[{"row": 4, "reason": "moved to appendix"}, {"row": 5, "reason": "moved to appendix"}])
        self.assertNotIn("silent_truncation", validate_visual_spec(declared))

    def test_brief_must_include_coverage(self):
        items = ["研究问题", "敏感性分析", "研究局限性"]
        slides = [
            slide_fixture("研究问题定义分析边界", key="q", coverage_items=[items[0]]),
            slide_fixture("敏感性分析支持主要方向", key="s", coverage_items=[items[1]]),
            slide_fixture("局限性限制结论强度", key="l", coverage_items=[items[2]]),
        ]
        contracts = coverage_contracts({"must_include": items, "target_slide_count": 3}, slides)
        self.assertEqual(len(contracts), len(items))
        self.assertTrue(all(not row["omitted"] and row["rendered_slide"] for row in contracts))

    def test_slide_budget(self):
        target = 10
        slides = [slide_fixture(f"结论式标题{i}", key=f"budget-{i}") for i in range(8)]
        ratio = len(slides) / target
        self.assertGreaterEqual(ratio, 0.80)
        contracts = coverage_contracts({"must_include": ["核心结果"], "target_slide_count": target}, [
            slide_fixture("核心结果保持完整", key="core", coverage_items=["核心结果"]),
        ])
        self.assertTrue(all(row["target_slide_count"] == target for row in contracts))

    def test_true_rob_matrix(self):
        rows = [
            {"study": study, "domain": domain, "judgment": judgment}
            for study, judgment in [("研究甲", "Low"), ("研究乙", "Some concerns")]
            for domain in ["随机化", "缺失"]
        ]
        visual = visual_fixture(
            "risk_of_bias_matrix", role="rob", columns=list(rows[0]), rows=rows,
            group_field="study", facet_field="domain", color_field="judgment",
        )
        self.assertNotIn("rob_matrix_incomplete", semantic_codes(visual))
        incomplete = dict(visual)
        incomplete["payload"] = {"rows": rows[:-1]}
        self.assertIn("rob_matrix_incomplete", semantic_codes(incomplete))

    def test_true_risk_matrix(self):
        rows = [
            {"risk": "招募延迟", "probability": "3", "impact": "4"},
            {"risk": "图像缺失", "probability": "2", "impact": "3"},
        ]
        visual = visual_fixture(
            "risk_matrix", role="risk", columns=list(rows[0]), rows=rows,
            x_field="probability", y_field="impact", label_field="risk",
        )
        self.assertNotIn("risk_matrix_axes_missing", semantic_codes(visual))
        invalid = dict(visual)
        invalid["payload"] = {"rows": [{"risk": "招募延迟", "probability": "3", "impact": ""}]}
        self.assertIn("risk_matrix_axes_missing", semantic_codes(invalid))

    def test_true_dag(self):
        payload = {
            "nodes": [
                {"id": "C", "role": "confounder"},
                {"id": "E", "role": "exposure"},
                {"id": "Y", "role": "outcome"},
            ],
            "edges": [
                {"source": "C", "target": "E"},
                {"source": "C", "target": "Y"},
                {"source": "E", "target": "Y"},
            ],
        }
        visual = visual_fixture("dag", role="dag", columns=["node", "edge"], payload=payload, expected=3)
        self.assertEqual(semantic_codes(visual), set())
        false_flow = dict(visual)
        false_flow["payload"] = {
            "nodes": [{"id": "E", "role": "exposure"}, {"id": "Y", "role": "outcome"}],
            "edges": [{"source": "E", "target": "Y"}],
        }
        self.assertIn("dag_semantic_roles_missing", semantic_codes(false_flow))

    def test_forest_axis_and_reference(self):
        rows = [{"label": "研究甲", "estimate": "1.5", "low": "1.1", "high": "2.0"}]
        visual = visual_fixture(
            "forest_plot", role="forest", columns=list(rows[0]), rows=rows,
            payload={"rows": rows, "axis_label": "比值比（OR）", "scale": "log"},
            x_field="estimate", ci_low_field="low", ci_high_field="high", label_field="label",
            reference_value=1, unit="OR",
        )
        self.assertEqual(validate_visual_spec(visual), [])
        self.assertEqual(semantic_codes(visual), set())
        invalid = dict(visual, reference_value=None)
        invalid["payload"] = {"rows": rows, "scale": "log"}
        self.assertIn("forest_reference_missing", semantic_codes(invalid))
        self.assertIn("forest_axis_missing", semantic_codes(invalid))

    def test_prediction_interval_not_ci(self):
        visual = visual_fixture(
            "pooled_evidence_panel", role="pooled", columns=["estimand", "low", "high"],
            payload={"rows": [], "prediction": {"estimand": "95% prediction interval", "low": "0.8", "high": "2.4"}},
            expected=1,
        )
        self.assertNotIn("prediction_interval_not_explicit", semantic_codes(visual))
        mislabeled = dict(visual)
        mislabeled["payload"] = {"rows": [], "prediction": {"estimand": "95% confidence interval", "low": "0.8", "high": "2.4"}}
        self.assertIn("prediction_interval_not_explicit", semantic_codes(mislabeled))

    def test_native_chart_data_binding(self):
        data = dataset("event_rates", ["group", "numerator", "denominator", "rate"], [
            {"group": "甲", "numerator": "10", "denominator": "40", "rate": "25.0"},
        ])
        binding = source_binding(data, ["group", "numerator", "denominator", "rate"], "all rows")
        visual = make_visual(
            "target_trial", "event-rate", "native_bar_chart", "event_rate", [binding],
            {"rows": data["rows"]}, 1, x_field="group", y_field="rate", unit="%",
            editability_requirement="native_chart",
        )
        self.assertEqual(visual["source_bindings"][0]["fields_used"], data["columns"])
        self.assertEqual(visual["editability_requirement"], "native_chart")
        self.assertTrue(visual["data_contract_id"].startswith("DATA-"))
        with self.assertRaises(ValueError):
            source_binding(data, ["group", "unregistered_field"])

    def test_minimum_font(self):
        typography = parse_flat_yaml(REPO / "design_system" / "tokens" / "typography.yaml")
        self.assertGreaterEqual(typography["slide_title_pt"], 30)
        self.assertGreaterEqual(typography["body_pt"], 18)
        self.assertGreaterEqual(typography["card_detail_pt"], 16)
        self.assertGreaterEqual(typography["chart_label_pt"], 15)
        self.assertGreaterEqual(typography["footer_pt"], 10)
        self.assertLessEqual(typography["footer_pt"], 11)

    def test_layout_family_diversity(self):
        masters = json.loads((REPO / "design_system" / "masters" / "families.json").read_text(encoding="utf-8"))["families"]
        rules = parse_flat_yaml(REPO / "design_system" / "composition_rules" / "rules.yaml")
        selected = ["hero_insight", "chart_led", "split_screen", "comparison_matrix", "conclusion_synthesis"]
        slides = [slide_fixture(f"视觉职责{i}形成不同节奏", key=f"layout-{i}", layout_family=family) for i, family in enumerate(selected)]
        self.assertTrue(set(selected) <= set(masters))
        self.assertGreaterEqual(len({slide["layout_family"] for slide in slides}), rules["minimum_layout_families_per_deck"])

    def test_card_grid_ratio(self):
        rules = parse_flat_yaml(REPO / "design_system" / "composition_rules" / "rules.yaml")
        slides = [slide_fixture(f"页面{i}承担独立视觉职责", key=f"ratio-{i}", card_grid=i < 3) for i in range(10)]
        ratio = sum(bool(slide["card_grid"]) for slide in slides) / len(slides)
        self.assertLessEqual(ratio, rules["maximum_card_grid_ratio"])
        overloaded = [dict(slide, card_grid=i < 4) for i, slide in enumerate(slides)]
        overloaded_ratio = sum(bool(slide["card_grid"]) for slide in overloaded) / len(overloaded)
        self.assertGreater(overloaded_ratio, rules["maximum_card_grid_ratio"])

    def test_generic_title_rate(self):
        titles = [
            "预设问题决定可解释的估计量", "权重诊断支持正值假设", "主要估计跨越无效线",
            "敏感性分析维持效应方向", "证据确定性限制结论强度", "亚组差异未证明异质性",
            "缺失数据仍可能造成偏倚", "风险矩阵确定优先缓解项", "里程碑按真实日期排布", "结果",
        ]
        self.assertLess(generic_title_rate(titles), 0.15)

    def test_duplicate_slide_semantics(self):
        first = slide_fixture("权重诊断支持正值假设", key="diagnostic")
        second = slide_fixture("敏感性分析维持效应方向", key="sensitivity")
        self.assertEqual(duplicate_semantic_signatures([first, second]), set())
        duplicate = dict(first, slide_id="SLD-DIFFERENT")
        self.assertTrue(duplicate_semantic_signatures([first, duplicate]))

    def test_source_figure_callout(self):
        sources = [{
            "source_id": "SRC-FIGURE",
            "source_file": "input/figures/registered_figure.png",
            "source_name": "registered_figure.png",
            "source_role": "figure",
        }]
        binding = asset_binding(sources, "registered_figure", "entire registered figure")
        visual = make_visual(
            "single_cell", "registered-figure", "registered_figure_callout", "source_figure",
            [binding], {"asset_path": "input/figures/registered_figure.png", "callouts": ["主要结构", "证据边界"]},
            1, entity_scope="registered source figure", editability_requirement="image_plus_native_callouts",
        )
        self.assertEqual(validate_visual_spec(visual), [])
        self.assertEqual(visual["source_bindings"][0]["source_id"], "SRC-FIGURE")
        self.assertTrue(visual["payload"]["callouts"])

    def test_protocol_no_results(self):
        planned = visual_fixture(
            "timeline_gantt", role="milestones", columns=["milestone", "start_date", "end_date"],
            rows=[
                {"milestone": "启动", "start_date": "2027-01-01", "end_date": "2027-02-01"},
                {"milestone": "分析", "start_date": "2028-01-01", "end_date": "2028-04-01"},
            ], time_field="start_date", label_field="milestone", end_time_field="end_date",
        )
        boundary = visual_fixture("no_result_boundary", role="protocol", expected=1, payload={"planned": ["样本量假设"], "not_available": ["事件率", "效应值"]})
        slides = [
            slide_fixture("研究里程碑按计划日期排列", key="plan", visuals=[planned]),
            slide_fixture("当前仅有方案与计划值，没有观察结果", key="boundary", visuals=[boundary]),
        ]
        self.assertEqual(protocol_result_violations(slides), [])
        fabricated = slide_fixture("结果显示事件率显著降低", key="fabricated", visuals=[visual_fixture("event_rate_plot")])
        self.assertTrue(protocol_result_violations([fabricated]))

    def test_canonical_source_category_is_registered_not_fabricated(self):
        binding = {
            "source_file": "input/documents/01_full_protocol_synthetic.docx",
            "source_location": "protocol status",
            "canonical_status": "supporting",
            "short_label_zh": "研究方案",
        }
        self.assertTrue(canonical_source_matches("protocol materials", binding))
        self.assertFalse(canonical_source_matches("model results", binding))


if __name__ == "__main__":
    unittest.main()
