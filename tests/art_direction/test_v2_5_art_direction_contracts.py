from __future__ import annotations

import hashlib
import importlib.util
import json
import re
import unittest
from pathlib import Path
from types import ModuleType
from typing import Any, Callable


REPO = Path(__file__).resolve().parents[2]
V25_ROOT = REPO / "scripts" / "v2_5"
REQUIRED_INTERFACE = {
    "build_art_direction_spec",
    "validate_art_direction_spec",
    "build_deck_rhythm_plan",
    "validate_deck_rhythm_plan",
    "plan_slide_presentation",
    "route_visual_render_mode",
    "decide_figure_rebuild",
    "empty_human_review_form",
    "scientific_contract_identity",
    "verify_hash_manifest",
}


def load_v25_modules() -> tuple[list[ModuleType], str]:
    """Load v2.5 modules without assuming a single production filename."""
    if not V25_ROOT.is_dir():
        return [], "V2_5_PRESENTATION_LAYER_INTERFACE_PENDING: scripts/v2_5 is absent"
    modules: list[ModuleType] = []
    failures: list[str] = []
    for path in sorted(V25_ROOT.glob("*.py")):
        if path.name.startswith("_"):
            continue
        name = f"academic_ppt_v25_contract_{path.stem}"
        try:
            spec = importlib.util.spec_from_file_location(name, path)
            if spec is None or spec.loader is None:
                failures.append(f"{path.name}: no loader")
                continue
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            modules.append(module)
        except Exception as exc:  # pragma: no cover - diagnostic path
            failures.append(f"{path.name}: {type(exc).__name__}: {exc}")
    if not modules:
        detail = "; ".join(failures) if failures else "no importable Python modules"
        return [], f"V2_5_PRESENTATION_LAYER_INTERFACE_PENDING: {detail}"
    return modules, ""


V25_MODULES, V25_LOAD_ERROR = load_v25_modules()


def require_api(test: unittest.TestCase, *names: str) -> dict[str, Callable[..., Any]]:
    found: dict[str, Callable[..., Any]] = {}
    missing: list[str] = []
    for name in names:
        value = next((getattr(module, name) for module in V25_MODULES if hasattr(module, name)), None)
        if not callable(value):
            missing.append(name)
        else:
            found[name] = value
    if missing:
        reason = V25_LOAD_ERROR or f"V2_5_PRESENTATION_LAYER_INTERFACE_PENDING: missing {', '.join(missing)}"
        test.fail(reason)
    return found


def generic_brief(project_name: str = "synthetic-alpha") -> dict[str, Any]:
    return {
        "project_name": project_name,
        "presentation_type": "research_report",
        "domain": "meta_analysis",
        "narrative_mode": "scientific_problem",
        "presentation_objective": "解释来源绑定的合成研究证据与限制",
        "audience": "医学科研同行",
        "language": "zh-CN",
        "duration_minutes": 18,
        "target_slide_count": 15,
        "theme_profile": "medical_academic",
        "design_intent": ["克制", "证据优先", "视觉层级明确"],
        "must_include": ["研究问题", "主要结果", "局限性", "结论"],
        "must_exclude": ["无来源装饰图", "因果化措辞"],
    }


def generic_visual(
    visual_id: str,
    visual_type: str,
    *,
    role: str = "supporting",
    source_id: str = "SRC-SYNTHETIC",
) -> dict[str, Any]:
    return {
        "visual_id": visual_id,
        "visual_type": visual_type,
        "scientific_role": role,
        "data_contract_id": f"DATA-{visual_id}",
        "source_bindings": [{
            "source_id": source_id,
            "source_file": "input/data/synthetic.csv",
            "source_location": "registered rows",
            "fields_used": ["entity", "estimate"],
            "row_filter": "all rows",
            "aggregation": "none",
            "canonical_status": "canonical",
        }],
        "editability_requirement": "native_editable",
        "payload": {
            "rows": [
                {"entity": "甲", "estimate": 1.2},
                {"entity": "乙", "estimate": 0.8},
            ]
        },
    }


def generic_slides(count: int = 12) -> list[dict[str, Any]]:
    roles = ["cover", "context", "method", "result", "limitation", "conclusion"]
    families = [
        "cover", "hero_insight", "process_branch", "chart_led",
        "split_screen", "conclusion_synthesis", "figure_with_callout",
        "comparison_matrix", "evidence_ladder", "risk_heatmap",
        "section_divider", "appendix_audit",
    ]
    return [
        {
            "slide_id": f"SLD-{role.upper()}-{index:02d}",
            "slide_role": role,
            "narrative_purpose": f"{role}阶段承担独立沟通职责",
            "title": f"{role}阶段的来源绑定信息形成可核查判断",
            "language": "zh-CN",
            "key_message": f"第{index + 1}个语义节拍只表达一个主要判断",
            "layout_family": families[index % len(families)],
            "appendix_status": "body",
            "visual_specs": [generic_visual(f"VIS-{index:02d}", "native_bar_chart", role="primary" if role == "result" else "supporting")],
        }
        for index, role in enumerate((roles * ((count + len(roles) - 1) // len(roles)))[:count])
    ]


def canonical_scientific_graph() -> dict[str, Any]:
    return {
        "graph_version": "2.4",
        "source_registry": [{"source_id": "SRC-SYNTHETIC", "sha256": "a" * 64}],
        "canonical_evidence_graph": {
            "claims": [{"claim_id": "CLM-001", "expected_value": "1.2", "source_ids": ["SRC-SYNTHETIC"]}],
            "conflicts": [{"conflict_id": "CONFLICT-001", "canonical_value": "1.2"}],
            "unresolved": [{"item_id": "UNRES-001", "status": "INFORMATION_REQUIRED"}],
        },
        "slide_specs": generic_slides(6),
        "backend": "native_pptxgenjs",
        "external_skill_used": False,
    }


class ArtDirectionContracts(unittest.TestCase):
    def test_art_direction_spec_schema_is_brief_driven_not_project_named(self):
        api = require_api(self, "build_art_direction_spec", "validate_art_direction_spec")
        slides = generic_slides()
        visual_types = [visual["visual_type"] for slide in slides for visual in slide["visual_specs"]]
        first = api["build_art_direction_spec"](
            generic_brief("synthetic-alpha"), domain="meta_analysis",
            slide_count=len(slides), available_visuals=visual_types,
        )
        second = api["build_art_direction_spec"](
            generic_brief("synthetic-beta"), domain="meta_analysis",
            slide_count=len(slides), available_visuals=visual_types,
        )
        self.assertEqual(api["validate_art_direction_spec"](first), [])
        self.assertEqual(api["validate_art_direction_spec"](second), [])
        required = {
            "art_direction_id", "presentation_type", "audience", "domain_profile",
            "visual_tone", "visual_motif", "primary_palette", "secondary_palette",
            "semantic_colors", "background_strategy", "section_backgrounds",
            "hero_style", "figure_treatment", "chart_treatment", "annotation_style",
            "typography_personality", "card_style", "divider_style",
            "conclusion_style", "appendix_style", "density_wave",
            "signature_components", "allowed_variants", "prohibited_styles",
            "backend_authority",
        }
        self.assertTrue(required <= set(first))
        presentation_fields = required - {"art_direction_id"}
        self.assertEqual(
            {key: first[key] for key in presentation_fields},
            {key: second[key] for key in presentation_fields},
        )

    def test_deck_rhythm_plan_limits_intensity_run_and_changes_every_three_to_five_slides(self):
        api = require_api(self, "build_art_direction_spec", "build_deck_rhythm_plan", "validate_deck_rhythm_plan")
        slides = generic_slides(15)
        spec = api["build_art_direction_spec"](
            generic_brief(), domain="meta_analysis", slide_count=len(slides),
            available_visuals=[visual["visual_type"] for slide in slides for visual in slide["visual_specs"]],
        )
        plan = api["build_deck_rhythm_plan"](slides, spec)
        self.assertEqual(api["validate_deck_rhythm_plan"](plan), [])
        beats = plan["slides"]
        max_run = 1
        current_run = 1
        for previous, current in zip(beats, beats[1:]):
            current_run = current_run + 1 if current["visual_intensity"] == previous["visual_intensity"] else 1
            max_run = max(max_run, current_run)
        self.assertLessEqual(max_run, 3)
        # Every 3–5 page passage must contain a perceptible rhythm change.  The
        # contract does not require an implementation-specific change marker.
        for window_size in (3, 4, 5):
            for start in range(0, len(beats) - window_size + 1):
                window = beats[start:start + window_size]
                signature = {(item["visual_intensity"], item["background_variant"]) for item in window}
                self.assertGreaterEqual(len(signature), 2)

    def test_layout_variant_diversity_is_role_driven(self):
        api = require_api(self, "build_art_direction_spec", "build_deck_rhythm_plan", "plan_slide_presentation")
        slides = generic_slides(12)
        spec = api["build_art_direction_spec"](generic_brief(), domain="meta_analysis", slide_count=len(slides))
        rhythm = api["build_deck_rhythm_plan"](slides, spec)
        plans = api["plan_slide_presentation"](slides, spec, rhythm)
        variants = {slide["layout_variant"] for slide in plans}
        self.assertGreaterEqual(len(variants), 5)
        self.assertLessEqual(
            max(sum(1 for _ in group) for group in _consecutive_groups(slide["layout_variant"] for slide in plans)),
            2,
        )

    def test_hero_area_is_35_to_65_percent_with_single_focus(self):
        api = require_api(self, "build_art_direction_spec", "build_deck_rhythm_plan", "plan_slide_presentation", "validate_presentation_specs")
        slides = generic_slides(12)
        slides[3]["slide_role"] = "result"
        slides[3]["visual_specs"] = [generic_visual("VIS-HERO", "forest_plot", role="primary_result")]
        spec = api["build_art_direction_spec"](generic_brief(), domain="meta_analysis", slide_count=len(slides))
        rhythm = api["build_deck_rhythm_plan"](slides, spec)
        planned = api["plan_slide_presentation"](slides, spec, rhythm)
        self.assertEqual(api["validate_presentation_specs"](planned, spec, rhythm), [])
        hero_slides = [slide for slide in planned if slide["hero_visual"]["enabled"]]
        self.assertTrue(hero_slides)
        for slide in hero_slides:
            self.assertGreaterEqual(slide["hero_visual"]["target_area_ratio"], 0.35)
            self.assertLessEqual(slide["hero_visual"]["target_area_ratio"], 0.65)
            self.assertEqual(slide["hero_visual"]["max_primary_anchors"], 1)

    def test_annotation_contract_has_one_evidence_and_one_interpretation(self):
        api = require_api(self, "build_art_direction_spec", "build_deck_rhythm_plan", "plan_slide_presentation")
        slides = generic_slides(12)
        pooled = generic_visual("VIS-ANNOTATION", "pooled_evidence_panel", role="primary_meta_analysis")
        pooled["payload"] = {
            "rows": [],
            "random": {"estimand": "Random-effects", "estimate": "1.42", "ci_low": "1.10", "ci_high": "1.83"},
            "prediction": {"estimand": "Prediction interval", "ci_low": "0.82", "ci_high": "2.47"},
        }
        slides[3]["visual_specs"] = [pooled]
        spec = api["build_art_direction_spec"](generic_brief(), domain="meta_analysis", slide_count=len(slides))
        rhythm = api["build_deck_rhythm_plan"](slides, spec)
        planned = api["plan_slide_presentation"](slides, spec, rhythm)
        annotations = planned[3]["visual_specs"][0]["annotation_specs"]
        priorities = [item["priority"] for item in annotations]
        self.assertEqual(priorities.count("primary"), 1)
        self.assertEqual(priorities.count("secondary"), 1)
        self.assertLessEqual(len(annotations), 2)

    def test_hybrid_render_routing_uses_only_internal_modes(self):
        api = require_api(self, "route_visual_render_mode")
        cases = [
            (generic_visual("VIS-CHART", "bar_chart"), "native_chart"),
            (generic_visual("VIS-GRAPH", "dag"), "editable_shapes"),
            (generic_visual("VIS-FIGURE", "source_figure"), "source_figure"),
            (generic_visual("VIS-SVG", "mechanism_svg"), "svg"),
        ]
        for visual, expected in cases:
            with self.subTest(visual_type=visual["visual_type"]):
                self.assertEqual(api["route_visual_render_mode"](visual), expected)

    def test_native_chart_privacy_excludes_raw_records_and_local_paths(self):
        api = require_api(self, "build_art_direction_spec", "route_visual_render_mode")
        visual = generic_visual("VIS-PRIVATE", "bar_chart")
        spec = api["build_art_direction_spec"](generic_brief(), domain="meta_analysis", slide_count=1)
        self.assertEqual(api["route_visual_render_mode"](visual), "native_chart")
        self.assertIn("patient_level_data_in_chart_workbook", spec["prohibited_styles"])
        renderer = (REPO / "scripts" / "v2_5" / "render_v2_5_deck.mjs").read_text(encoding="utf-8")
        self.assertNotIn("raw_records", renderer)
        self.assertNotIn("patient_id", renderer)

    def test_figure_rebuild_decision_is_evidence_and_editability_driven(self):
        api = require_api(self, "decide_figure_rebuild")
        umap = generic_visual("VIS-UMAP", "umap_figure", role="dimensional_reduction")
        umap["privacy_safe_aggregate"] = True
        umap["presentation_payload"] = {
            "aggregate_rows": [{"UMAP_1": 0.1, "UMAP_2": -0.2, "cell_type": "A"}],
            "embedded_fields": ["UMAP_1", "UMAP_2", "cell_type"],
        }
        source_figure = generic_visual("VIS-SOURCE", "source_figure", role="microscopy")
        selection = generic_visual("VIS-SELECTION", "source_figure", role="study_selection")
        selection["payload"]["callout"] = "仅保留来源图并解释筛选流程"
        allowed = {
            "REDRAW_FROM_STRUCTURED_DATA", "PRESERVE_SOURCE_FIGURE",
            "PRESERVE_WITH_CALLOUT", "MANUAL_REVIEW_REQUIRED",
        }
        decisions = {
            api["decide_figure_rebuild"](umap, {"UMAP_1", "UMAP_2", "cell_type"}, True),
            api["decide_figure_rebuild"](source_figure, set(), True),
            api["decide_figure_rebuild"](selection, set(), True),
            api["decide_figure_rebuild"](selection, set(), False),
        }
        self.assertTrue(decisions <= allowed)
        self.assertIn("REDRAW_FROM_STRUCTURED_DATA", decisions)
        self.assertIn("PRESERVE_SOURCE_FIGURE", decisions)
        self.assertIn("PRESERVE_WITH_CALLOUT", decisions)
        self.assertIn("MANUAL_REVIEW_REQUIRED", decisions)

    def test_unicode_chinese_typography_preserves_text_and_font_contract(self):
        api = require_api(self, "normalize_unicode_typography", "validate_unicode_typography")
        text = "目标试验模拟（Target Trial Emulation）在二十八天结局前定义暴露窗口"
        normalized = api["normalize_unicode_typography"](text + "\u200b")
        self.assertEqual(normalized, text)
        self.assertEqual(api["validate_unicode_typography"]([normalized]), [])
        self.assertNotRegex(normalized, r"(?:ME T HO D|RES U LT|\ufffd|锟斤拷)")
        renderer = (REPO / "scripts" / "v2_5" / "render_v2_5_deck.mjs").read_text(encoding="utf-8")
        self.assertRegex(renderer, r"Microsoft YaHei|SimHei|Noto Sans CJK|Source Han Sans")

    def test_title_is_at_most_two_lines_without_silent_semantic_loss(self):
        api = require_api(self, "normalize_unicode_typography", "validate_unicode_typography")
        first = "在预设来源绑定和证据边界下，主要结果保持可核查"
        second = "统计不确定性与人工审核责任同时保留"
        title = api["normalize_unicode_typography"](first + "\n" + second)
        self.assertEqual(api["validate_unicode_typography"]([title]), [])
        errors = api["validate_unicode_typography"]([first + "\n" + second + "\n第三行不得出现"])
        self.assertTrue(any(item["issue"] == "title_more_than_two_lines" for item in errors))
        renderer = (REPO / "scripts" / "v2_5" / "render_v2_5_deck.mjs").read_text(encoding="utf-8")
        self.assertRegex(renderer, r"title[^\n]{0,120}(?:30|3[1-9]|4\d|5\d)")

    def test_visual_diversity_threshold_is_met_without_decorative_assets(self):
        api = require_api(
            self, "build_art_direction_spec", "build_deck_rhythm_plan",
            "plan_slide_presentation", "visual_diversity_metrics",
        )
        slides = generic_slides(12)
        visual_types = [
            "hero_question", "timeline_gantt", "native_bar_chart", "forest_plot",
            "risk_matrix", "source_figure", "evidence_ladder", "dag",
            "grouped_composition", "takeaway_synthesis", "risk_of_bias_matrix", "appendix_audit",
        ]
        for slide, visual_type in zip(slides, visual_types):
            slide["visual_specs"] = [generic_visual(f"VIS-{visual_type}", visual_type, role="primary")]
        spec = api["build_art_direction_spec"](
            generic_brief(), domain="meta_analysis", slide_count=len(slides),
            available_visuals=visual_types,
        )
        rhythm = api["build_deck_rhythm_plan"](slides, spec)
        planned = api["plan_slide_presentation"](slides, spec, rhythm)
        metrics = api["visual_diversity_metrics"](planned)
        self.assertGreaterEqual(metrics["composition_variant_count"], 5)
        self.assertGreaterEqual(metrics["background_variant_count"], 2)
        self.assertGreaterEqual(metrics["hero_slide_count"], 2)
        self.assertGreater(metrics["chart_led_count"] + metrics["figure_led_count"], 0)
        self.assertNotIn("decorative_gradient", spec.get("allowed_variants", {}))
        self.assertIn("decorative_gradient", spec["prohibited_styles"])

    def test_human_visual_review_form_is_never_auto_filled(self):
        api = require_api(self, "empty_human_review_form")
        form = api["empty_human_review_form"]("DECK-SYNTHETIC")
        score_fields = [key for key in form if key.endswith("_score") or re.search(r"_\d+$", key)]
        self.assertTrue(score_fields)
        self.assertTrue(all(form[key] in {None, ""} for key in score_fields))
        self.assertNotIn("automatic_score", form)
        self.assertFalse(form.get("reviewer"))
        self.assertFalse(form.get("review_date"))

    def test_frozen_v24_scientific_contract_identity_ignores_presentation_layer(self):
        api = require_api(self, "scientific_contract_identity")
        frozen = canonical_scientific_graph()
        redesigned = json.loads(json.dumps(frozen, ensure_ascii=False))
        redesigned["art_direction_spec"] = {"style_keywords": ["克制", "结构化"]}
        redesigned["deck_rhythm_plan"] = {"slides": [{"slide_id": "SLD-X", "intensity": "high"}]}
        redesigned["presentation_artifacts"] = {"layout_variant": "asymmetric_hero"}
        self.assertEqual(api["scientific_contract_identity"](frozen), api["scientific_contract_identity"](redesigned))

    def test_no_external_skill_and_pptxgenjs_remains_default(self):
        api = require_api(self, "build_art_direction_spec", "build_deck_rhythm_plan", "plan_slide_presentation")
        slides = generic_slides(1)
        visual = slides[0]["visual_specs"][0]
        spec = api["build_art_direction_spec"](generic_brief(), domain="meta_analysis", slide_count=1)
        rhythm = api["build_deck_rhythm_plan"](slides, spec)
        planned = api["plan_slide_presentation"](slides, spec, rhythm)
        self.assertEqual(spec["backend_authority"], "native_pptxgenjs")
        self.assertIn("external_skill_backend", spec["prohibited_styles"])
        self.assertTrue(all(
            item["preferred_backend"] == "native_pptxgenjs"
            for slide in planned for item in slide["visual_specs"]
        ))
        self.assertTrue(all(
            item["fallback_backend"] == "native_pptxgenjs_editable_shapes"
            for slide in planned for item in slide["visual_specs"]
        ))

    def test_input_hash_contract_detects_any_change(self):
        api = require_api(self, "verify_hash_manifest")
        before = {
            "input/data/a.csv": hashlib.sha256(b"alpha").hexdigest(),
            "input/documents/b.docx": hashlib.sha256(b"beta").hexdigest(),
        }
        self.assertEqual(api["verify_hash_manifest"](before, dict(before)), [])
        changed = dict(before)
        changed["input/data/a.csv"] = hashlib.sha256(b"changed").hexdigest()
        errors = api["verify_hash_manifest"](before, changed)
        self.assertTrue(errors)
        self.assertTrue(any("input/data/a.csv" in str(error) for error in errors))


def _consecutive_groups(values):
    iterator = iter(values)
    try:
        first = next(iterator)
    except StopIteration:
        return []
    groups: list[list[Any]] = [[first]]
    for value in iterator:
        if value == groups[-1][-1]:
            groups[-1].append(value)
        else:
            groups.append([value])
    return groups


if __name__ == "__main__":
    unittest.main()
