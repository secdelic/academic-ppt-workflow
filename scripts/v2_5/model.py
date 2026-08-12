from __future__ import annotations

import copy
import csv
import hashlib
import json
import math
import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable


ART_DIRECTION_REQUIRED = {
    "art_direction_id",
    "presentation_type",
    "audience",
    "domain_profile",
    "visual_tone",
    "visual_motif",
    "primary_palette",
    "secondary_palette",
    "semantic_colors",
    "background_strategy",
    "section_backgrounds",
    "hero_style",
    "figure_treatment",
    "chart_treatment",
    "annotation_style",
    "typography_personality",
    "card_style",
    "divider_style",
    "conclusion_style",
    "appendix_style",
    "density_wave",
    "emphasis_slides",
    "section_breaks",
    "signature_components",
    "allowed_variants",
    "prohibited_styles",
}


DOMAIN_ART = {
    "target_trial": {
        "domain_profile": "clinical_methods",
        "visual_tone": "precise, method-forward, restrained",
        "visual_motif": "dual_strategy_branch",
        "primary_palette": ["16324F", "008A8A", "E76F51"],
        "secondary_palette": ["DCECEE", "FCE6E0", "F4F7FA"],
        "semantic_colors": {"strategy_a": "008A8A", "strategy_b": "E76F51", "warning": "C26A1B", "neutral": "64748B"},
        "signature_components": ["dual_strategy_branch", "absolute_risk_hero", "assumption_warning_band"],
    },
    "single_cell": {
        "domain_profile": "bioinformatics",
        "visual_tone": "analytical, molecular, evidence-layered",
        "visual_motif": "cell_identity_molecular_state",
        "primary_palette": ["17324D", "4C78A8", "B35C9B"],
        "secondary_palette": ["59A14F", "F28E2B", "E5E9F0"],
        "semantic_colors": {"positive": "C44E52", "negative": "4C78A8", "prediction": "8E6C8A", "neutral": "A6AAB2"},
        "signature_components": ["cell_identity_strip", "diverging_effect_axis", "prediction_boundary_badge"],
    },
    "meta_analysis": {
        "domain_profile": "evidence_synthesis",
        "visual_tone": "evidence-centered, authoritative, cautious",
        "visual_motif": "evidence_aggregation",
        "primary_palette": ["1F3B5B", "2A7F9E", "6B7280"],
        "secondary_palette": ["DDECF2", "F3E8C8", "F5F7F9"],
        "semantic_colors": {"low_risk": "4E9F6D", "some_concern": "E6A23C", "high_risk": "C94C4C", "uncertain": "9CA3AF"},
        "signature_components": ["pooled_diamond", "evidence_axis", "certainty_hierarchy"],
    },
    "protocol": {
        "domain_profile": "clinical_protocol",
        "visual_tone": "planned, operational, non-resultative",
        "visual_motif": "planned_pathway_timeline",
        "primary_palette": ["173B5E", "1C8C8C", "607D8B"],
        "secondary_palette": ["DCEFF1", "E8EEF3", "F5F8FA"],
        "semantic_colors": {"planned": "2A7F9E", "required": "1C8C8C", "review": "C47B23", "not_observed": "7B8794"},
        "signature_components": ["planning_band", "time_axis", "review_marker"],
    },
}


LAYOUT_VARIANTS = {
    "cover": ["minimal_dark", "motif_dark"],
    "section_divider": ["dark_band", "tinted_field"],
    "hero_insight": ["hero_left", "hero_right", "hero_number", "full_width"],
    "chart_led": ["standard", "full_width", "hero_left", "hero_right", "chart_plus_insight", "chart_plus_secondary_metric"],
    "figure_with_callout": ["large_left", "large_right", "full_bleed", "top_figure_bottom_callout"],
    "split_screen": ["balanced", "asymmetric_left", "asymmetric_right", "boundary_contrast"],
    "comparison_matrix": ["2_column", "3_column", "asymmetric", "delta_focus"],
    "process_branch": ["horizontal", "vertical", "central_branch", "swimlane"],
    "evidence_ladder": ["ascending", "central_spine", "certainty_focus"],
    "risk_heatmap": ["matrix_led", "matrix_plus_mitigation", "full_width"],
    "conclusion_synthesis": ["numbered_takeaways", "hero_number_plus_takeaways", "evidence_boundary_summary"],
    "appendix_audit": ["audit_table", "review_register"],
}


NATIVE_CHART_TYPES = {
    "bar_chart",
    "grouped_bar",
    "line_chart",
    "area_chart",
    "scatter_plot",
    "histogram",
    "simple_time_series",
    "weight_histogram",
    "grouped_composition",
    "umap_scatter",
    "funnel_plot",
}

EDITABLE_SHAPE_TYPES = {
    "forest_plot",
    "subgroup_forest",
    "love_plot",
    "risk_of_bias_matrix",
    "risk_matrix",
    "dag",
    "clone_censor_weight_branch",
    "single_cell_workflow",
    "protocol_design",
    "analysis_plan",
    "evidence_ladder",
    "quality_control_ladder",
    "timeline_gantt",
    "assessment_timeline",
    "time_window_timeline",
    "communication_network",
    "faceted_volcano",
    "diverging_enrichment",
    "sample_size_waterfall",
    "pooled_evidence_panel",
}


CONTROL_CHARS = re.compile(r"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F\u200B\u200C\u200D\u2060\uFEFF]")
REPLACEMENT_CHAR = "\uFFFD"


def stable(prefix: str, *parts: object, n: int = 16) -> str:
    raw = "\u241f".join(str(part) for part in parts)
    return f"{prefix}-{hashlib.sha256(raw.encode('utf-8')).hexdigest()[:n].upper()}"


def canonical_json_hash(value: Any) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def canonical_scientific_projection(graph: dict[str, Any]) -> dict[str, Any]:
    """Return only the frozen v2.4 scientific/governance state.

    Presentation fields and renderer evidence are deliberately excluded.  The
    projection is used to prove that v2.5 did not redesign the evidence layer.
    """
    slides = []
    for slide in graph.get("slide_specs", []):
        slide_copy = {
            key: copy.deepcopy(value)
            for key, value in slide.items()
            if key not in {
                "art_direction_id", "layout_variant", "background_variant",
                "visual_intensity", "density_target", "hero_priority",
                "transition_role", "section_role", "hero_visual",
                "presentation_revision",
            }
        }
        for visual in slide_copy.get("visual_specs", []):
            for key in (
                "render_mode", "annotation_specs", "figure_rebuild_decision",
                "presentation_variant", "chart_treatment", "hero_area_ratio",
                "preferred_backend", "fallback_backend", "presentation_revision",
                "render_visual_type", "presentation_payload", "native_chart_privacy",
            ):
                visual.pop(key, None)
            # Actual rendering evidence is not part of the scientific contract.
            visual["rendered_row_count"] = None
            visual["rendered_row_keys"] = []
        slides.append(slide_copy)
    return {
        "project_key": graph.get("project_key"),
        "kind": graph.get("kind"),
        "brief": copy.deepcopy(graph.get("brief", {})),
        "source_registry": copy.deepcopy(graph.get("source_registry", [])),
        "canonical_evidence_graph": copy.deepcopy(graph.get("canonical_evidence_graph", {})),
        "story_graph": copy.deepcopy(graph.get("story_graph", {})),
        "slide_specs": slides,
        "backend": graph.get("backend", "native_pptxgenjs"),
        "external_skill_used": bool(graph.get("external_skill_used", False)),
        "ground_truth_used_for_generation": bool(graph.get("ground_truth_used_for_generation", False)),
    }


def scientific_contract_identity(graph: dict[str, Any]) -> str:
    return canonical_json_hash(canonical_scientific_projection(graph))


def _domain_config(domain: str) -> dict[str, Any]:
    if domain not in DOMAIN_ART:
        raise ValueError(f"UNSUPPORTED_DOMAIN_PROFILE:{domain}")
    return copy.deepcopy(DOMAIN_ART[domain])


def build_art_direction_spec(
    brief: dict[str, Any],
    presentation_type: str | None = None,
    narrative_mode: str | None = None,
    domain: str | None = None,
    slide_count: int | None = None,
    available_visuals: Iterable[str] | None = None,
) -> dict[str, Any]:
    domain = domain or str(brief.get("domain") or brief.get("kind") or "")
    config = _domain_config(domain)
    count = int(slide_count or brief.get("target_slide_count") or 1)
    visuals = sorted(set(available_visuals or []))
    dark_target = max(1, min(math.floor(count * 0.25), round(count * 0.16)))
    presentation_type = presentation_type or str(brief.get("presentation_type", "academic_presentation"))
    narrative_mode = narrative_mode or str(brief.get("narrative_mode", "scientific_problem"))
    payload = {
        "schema_version": "2.5",
        "language": str(brief.get("language", "zh-CN")),
        "presentation_type": presentation_type,
        "audience": str(brief.get("audience", "academic audience")),
        "domain_profile": config["domain_profile"],
        "visual_tone": config["visual_tone"],
        "visual_motif": config["visual_motif"],
        "primary_palette": config["primary_palette"],
        "secondary_palette": config["secondary_palette"],
        "semantic_colors": config["semantic_colors"],
        "background_strategy": {
            "light": "default analytical canvas",
            "tinted": "section or interpretive transition",
            "dark": "cover and evidence climax only",
            "target_dark_slide_count": dark_target,
            "dark_ratio_range": [0.10, 0.25],
        },
        "section_backgrounds": {
            "context": "tinted",
            "method": "light",
            "result": "light_or_dark_emphasis",
            "limitation": "tinted",
            "conclusion": "dark_emphasis",
            "appendix": "light_audit",
        },
        "hero_style": "single_anchor_with_direct_evidence_annotation",
        # One canonical field controls both the ArtDirectionSpec and renderer.
        # The renderer must not silently substitute a second figure-treatment
        # vocabulary.
        "figure_treatment": "figure_with_callout",
        "chart_treatment": "native_chart_for_standard_quantitative_graphics_else_editable_shapes",
        "annotation_style": "one_primary_plus_one_secondary_maximum",
        "typography_personality": "modern_medical_sans_with_zh_cn_optical_spacing",
        "card_style": "flat_low_chrome_no_dashboard",
        "divider_style": "tonal_field_with_signature_motif",
        "conclusion_style": "one_synthesis_anchor_plus_evidence_boundary",
        "appendix_style": "neutral_audit_register",
        "density_wave": "semantic_low_to_high_to_release",
        "emphasis_slides": [],
        "section_breaks": [],
        "signature_components": config["signature_components"],
        "allowed_variants": copy.deepcopy(LAYOUT_VARIANTS),
        "prohibited_styles": [
            "decorative_gradient", "glow", "complex_shadow", "entertainment_illustration",
            "irrelevant_photo", "patient_level_data_in_chart_workbook", "external_skill_backend",
            "result_semantics_in_protocol", "multiple_competing_hero_anchors",
        ],
        "narrative_mode": narrative_mode,
        "slide_count": count,
        "available_visuals": visuals,
        "backend_authority": "native_pptxgenjs",
        "style_keywords": ["学术", "克制", "证据优先", config["visual_motif"]],
        "layout_variant_policy": "semantic_role_and_visual_structure",
        "rhythm_policy": "no_more_than_three_equal_intensity_and_semantic_change_every_three_to_five_slides",
        "hero_policy": "two_to_four_single_anchor_slides_at_35_to_65_percent_content_area",
        "annotation_policy": "one_primary_and_one_secondary_maximum",
        "render_policy": {"default_backend": "native_pptxgenjs", "external_skill_enabled": False, "fallback": "editable_shapes"},
        "privacy_policy": "aggregate_only_native_chart_workbook_no_patient_or_cell_level_rows",
    }
    payload["art_direction_id"] = stable("ART", canonical_json_hash(payload), n=16)
    return payload


def validate_art_direction_spec(spec: dict[str, Any]) -> list[str]:
    errors = [f"missing:{field}" for field in sorted(ART_DIRECTION_REQUIRED - set(spec))]
    if spec.get("backend_authority") != "native_pptxgenjs":
        errors.append("non_native_backend_authority")
    dark_range = spec.get("background_strategy", {}).get("dark_ratio_range")
    if dark_range != [0.10, 0.25]:
        errors.append("dark_background_contract_missing")
    if not spec.get("signature_components"):
        errors.append("signature_components_empty")
    return errors


def _hero_score(slide: dict[str, Any]) -> int:
    if slide.get("appendix_status") == "appendix" or slide.get("slide_role") == "cover":
        return -1
    visuals = slide.get("visual_specs", [])
    roles = {visual.get("scientific_role") for visual in visuals}
    types = {visual.get("visual_type") for visual in visuals}
    score = 0
    if "primary_result" in roles or "primary_meta_analysis" in roles:
        score += 120
    if "sample_size_planning" in roles or "evidence_interpretation" in roles:
        score += 110
    if slide.get("slide_role") == "conclusion":
        score += 90
    if types & {"clone_censor_weight_branch", "dag", "single_cell_workflow", "protocol_design"}:
        score += 80
    if slide.get("layout_family") == "hero_insight":
        score += 50
    if slide.get("slide_role") == "result":
        score += 35
    return score


def build_deck_rhythm_plan(slides: list[dict[str, Any]], art_direction: dict[str, Any]) -> dict[str, Any]:
    content_slides = [slide for slide in slides if slide.get("appendix_status") != "appendix" and slide.get("slide_role") != "cover"]
    hero_count = min(4, max(2, round(len(content_slides) / 5)))
    ranked = sorted(content_slides, key=lambda slide: (-_hero_score(slide), slide.get("slide_id", "")))
    # Two-to-four hero slides is a deck-level target, not permission to promote
    # semantically ordinary pages.  Only the user-defined hero roles/types may
    # become a hero; a deck with fewer eligible slides reports fewer heroes.
    eligible = [slide for slide in ranked if _hero_score(slide) > 0]
    hero_ids = {slide["slide_id"] for slide in eligible[:hero_count]}

    target_dark = int(art_direction.get("background_strategy", {}).get("target_dark_slide_count", 2))
    dark_ids: set[str] = set()
    cover = next((slide for slide in slides if slide.get("slide_role") == "cover"), None)
    if cover:
        dark_ids.add(cover["slide_id"])
    for slide in ranked:
        if len(dark_ids) >= target_dark:
            break
        if slide["slide_id"] in hero_ids:
            dark_ids.add(slide["slide_id"])

    result = []
    prior_role = None
    run_intensity = None
    run_length = 0
    for index, slide in enumerate(slides):
        role = slide.get("slide_role", "content")
        visual_roles = {visual.get("scientific_role") for visual in slide.get("visual_specs", [])}
        is_hero = slide["slide_id"] in hero_ids
        if role in {"cover", "appendix"}:
            intensity = 1
        elif is_hero:
            intensity = 4
        elif role == "result" or visual_roles & {"diagnostic", "subgroup_analysis", "supported_interpretation"}:
            intensity = 3
        else:
            intensity = 2
        if intensity == run_intensity:
            run_length += 1
        else:
            run_intensity, run_length = intensity, 1
        if run_length > 3:
            intensity = 3 if intensity == 2 else 2
            run_intensity, run_length = intensity, 1
        section_break = prior_role is not None and role != prior_role and role not in {"appendix"}
        if slide["slide_id"] in dark_ids:
            background = "dark_emphasis"
        elif section_break or role in {"context", "limitation"}:
            background = "tinted"
        else:
            background = "light"
        result.append({
            "slide_id": slide["slide_id"],
            "visual_intensity": intensity,
            "density_target": {1: "low", 2: "moderate", 3: "data_dense", 4: "hero_focused"}[intensity],
            "background_variant": background,
            "hero_priority": "required" if is_hero else "none",
            "transition_role": "section_transition" if section_break else "continuation",
            "section_role": role,
            "is_hero": is_hero,
            "hero_area_ratio": 0.52 if is_hero else 0.0,
        })
        prior_role = role
    art_direction["emphasis_slides"] = [item["slide_id"] for item in result if item["is_hero"]]
    art_direction["section_breaks"] = [item["slide_id"] for item in result if item["transition_role"] == "section_transition"]
    # art_direction_id identifies the final, replayable deck-level specification
    # after emphasis and section breaks are known.
    identity_payload = {key: value for key, value in art_direction.items() if key != "art_direction_id"}
    art_direction["art_direction_id"] = stable("ART", canonical_json_hash(identity_payload), n=16)
    return {
        "art_direction_id": art_direction["art_direction_id"],
        "slides": result,
        "hero_slide_count": sum(item["is_hero"] for item in result),
        "dark_slide_count": sum(item["background_variant"] == "dark_emphasis" for item in result),
        "dark_slide_ratio": round(sum(item["background_variant"] == "dark_emphasis" for item in result) / max(1, len([s for s in slides if s.get("appendix_status") != "appendix"])), 4),
    }


def validate_deck_rhythm_plan(plan: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    slides = plan.get("slides", [])
    if not slides:
        return ["rhythm_plan_empty"]
    run = 0
    previous = None
    for item in slides:
        level = item.get("visual_intensity")
        if level not in {1, 2, 3, 4}:
            errors.append(f"invalid_intensity:{item.get('slide_id')}")
        run = run + 1 if level == previous else 1
        if run > 3:
            errors.append(f"intensity_run_gt_3:{item.get('slide_id')}")
        previous = level
    if len(slides) >= 12 and not 2 <= plan.get("hero_slide_count", 0) <= 4:
        errors.append("hero_count_outside_2_4")
    ratio = float(plan.get("dark_slide_ratio", 0))
    if len(slides) >= 10 and not 0.10 <= ratio <= 0.25:
        errors.append("dark_background_ratio_outside_10_25pct")
    return errors


def _variant_for(slide: dict[str, Any], occurrence: int, is_hero: bool) -> str:
    family = slide.get("layout_family", "chart_led")
    variants = LAYOUT_VARIANTS.get(family, ["standard"])
    visual_types = {visual.get("visual_type") for visual in slide.get("visual_specs", [])}
    if is_hero:
        for preferred in ("full_width", "hero_left", "hero_number_plus_takeaways", "delta_focus", "matrix_led"):
            if preferred in variants:
                return preferred
    if visual_types & {"source_figure", "umap_figure"}:
        for preferred in ("large_left", "large_right", "full_bleed"):
            if preferred in variants:
                return variants[(occurrence + variants.index(preferred)) % len(variants)]
    return variants[occurrence % len(variants)]


def route_visual_render_mode(visual: dict[str, Any]) -> str:
    visual_type = str(visual.get("render_visual_type") or visual.get("visual_type", ""))
    if visual_type in NATIVE_CHART_TYPES:
        return "native_chart"
    if visual_type in {"source_figure", "umap_figure"}:
        return "source_figure"
    if visual_type.endswith("_svg"):
        return "svg"
    return "editable_shapes"


def decide_figure_rebuild(
    visual: dict[str, Any],
    structured_fields: Iterable[str] | None = None,
    source_available: bool = True,
) -> str:
    fields = {str(item) for item in (structured_fields or [])}
    role = str(visual.get("scientific_role", ""))
    visual_type = str(visual.get("visual_type", ""))
    if visual_type == "umap_figure" and {"UMAP_1", "UMAP_2"} <= fields and fields & {"cell_type", "group"}:
        # A PowerPoint scatter workbook would expose one row per cell.  Keep the
        # registered source figure unless the presentation planner supplies an
        # explicit privacy-safe aggregate or approved de-identified subset.
        approved_rows = visual.get("presentation_payload", {}).get("aggregate_rows")
        approved_fields = set(visual.get("presentation_payload", {}).get("embedded_fields", []))
        safe_redraw = (
            visual.get("privacy_safe_aggregate") is True
            and isinstance(approved_rows, list)
            and bool(approved_rows)
            and {"UMAP_1", "UMAP_2"} <= approved_fields
        )
        return "REDRAW_FROM_STRUCTURED_DATA" if safe_redraw else "PRESERVE_WITH_CALLOUT"
    if role == "publication_bias_assessment" and ({"odds_ratio", "ci_low", "ci_high"} <= fields or {"effect", "standard_error"} <= fields):
        return "REDRAW_FROM_STRUCTURED_DATA"
    if role == "study_selection":
        return "PRESERVE_WITH_CALLOUT" if source_available else "MANUAL_REVIEW_REQUIRED"
    if source_available:
        return "PRESERVE_WITH_CALLOUT" if visual.get("payload", {}).get("callout") else "PRESERVE_SOURCE_FIGURE"
    return "MANUAL_REVIEW_REQUIRED"


def derive_annotations(visual: dict[str, Any]) -> list[dict[str, Any]]:
    rows = visual.get("payload", {}).get("rows", [])
    visual_type = visual.get("visual_type")
    annotations: list[dict[str, Any]] = []
    if visual_type in {"source_figure", "umap_figure"} and visual.get("payload", {}).get("callout"):
        annotations.append({
            "annotation_type": "evidence_boundary_annotation",
            "priority": "primary",
            "text": str(visual["payload"]["callout"]),
        })
    elif visual_type == "absolute_risk_comparison":
        row = next((item for item in rows if "difference" in str(item.get("estimand", "")).lower()), None)
        if row:
            value = float(row.get("estimate", 0)) * 100
            annotations.append({"annotation_type": "delta_annotation", "priority": "primary", "text": f"{value:+.1f} pp", "value": value})
    elif visual_type == "grouped_composition":
        group_field = visual.get("group_field") or "group"
        entity_field = visual.get("label_field") or "cell_type"
        value_field = visual.get("y_field") or "percentage"
        by_entity: dict[str, list[dict[str, Any]]] = {}
        for row in rows:
            by_entity.setdefault(str(row.get(entity_field, "")), []).append(row)
        candidates = []
        for entity, items in by_entity.items():
            if len({item.get(group_field) for item in items}) >= 2:
                values = [float(item[value_field]) for item in items]
                candidates.append((abs(max(values) - min(values)), max(values) - min(values), entity))
        if candidates:
            _, delta, entity = max(candidates)
            annotations.append({"annotation_type": "delta_annotation", "priority": "primary", "text": f"{entity} {delta:+.1f} pp", "value": delta, "entity": entity})
    elif visual_type == "pooled_evidence_panel":
        random_row = visual.get("payload", {}).get("random") or next((item for item in rows if "Random" in str(item.get("estimand", ""))), None)
        prediction = visual.get("payload", {}).get("prediction") or next((item for item in rows if "Prediction" in str(item.get("estimand", ""))), None)
        if random_row:
            annotations.append({"annotation_type": "key_value_callout", "priority": "primary", "text": f"OR {float(random_row['estimate']):.2f}", "value": float(random_row["estimate"])})
        if prediction:
            annotations.append({"annotation_type": "confidence_annotation", "priority": "secondary", "text": f"PI {prediction['ci_low']}–{prediction['ci_high']}"})
    elif visual_type == "sample_size_waterfall":
        target = next((item for item in rows if "Target enrollment" in str(item.get("parameter", ""))), None)
        analyzable = next((item for item in rows if "Required analyzable" in str(item.get("parameter", ""))), None)
        if target and analyzable:
            annotations.append({"annotation_type": "bracket_comparison", "priority": "primary", "text": f"{target['value']} → {analyzable['value']}"})
    elif visual_type in {"forest_plot", "subgroup_forest", "love_plot"} and visual.get("reference_value") is not None:
        annotations.append({"annotation_type": "reference_line_label", "priority": "secondary", "text": f"参考值 {visual['reference_value']}"})
    return annotations[:2]


ANNOTATION_TYPES = {
    "direct_label", "key_value_callout", "delta_annotation",
    "threshold_annotation", "confidence_annotation", "caution_annotation",
    "evidence_boundary_annotation", "reference_line_label",
    "highlight_band", "bracket_comparison",
}


def validate_annotation_specs(annotations: list[dict[str, Any]]) -> list[str]:
    errors = []
    if len(annotations) > 2:
        errors.append("annotation_count_gt_2")
    if sum(item.get("priority") == "primary" for item in annotations) > 1:
        errors.append("primary_annotation_gt_1")
    if sum(item.get("priority") == "secondary" for item in annotations) > 1:
        errors.append("secondary_annotation_gt_1")
    for index, item in enumerate(annotations):
        annotation_type = item.get("annotation_type")
        if annotation_type not in ANNOTATION_TYPES:
            errors.append(f"annotation_type_invalid:{index}:{annotation_type}")
        if item.get("priority") not in {"primary", "secondary"}:
            errors.append(f"annotation_priority_invalid:{index}")
        if not str(item.get("text", "")).strip():
            errors.append(f"annotation_text_missing:{index}")
    return errors


def plan_slide_presentation(
    slides: list[dict[str, Any]],
    art_direction: dict[str, Any],
    rhythm_plan: dict[str, Any],
    structured_fields_by_source: dict[str, set[str]] | None = None,
) -> list[dict[str, Any]]:
    result = copy.deepcopy(slides)
    rhythm_by_id = {item["slide_id"]: item for item in rhythm_plan["slides"]}
    family_occurrence: dict[str, int] = {}
    structured_fields_by_source = structured_fields_by_source or {}
    for slide in result:
        rhythm = rhythm_by_id[slide["slide_id"]]
        family = slide.get("layout_family", "chart_led")
        occurrence = family_occurrence.get(family, 0)
        family_occurrence[family] = occurrence + 1
        visual_candidates = slide.get("visual_specs", [])
        primary_visual_id = None
        if rhythm["is_hero"] and visual_candidates:
            def visual_hero_score(item: dict[str, Any]) -> tuple[int, str]:
                role = str(item.get("scientific_role", ""))
                visual_type = str(item.get("visual_type", ""))
                score = 0
                if role in {"primary_result", "primary_meta_analysis", "sample_size_planning", "evidence_interpretation"}:
                    score += 100
                if visual_type in {"clone_censor_weight_branch", "dag", "single_cell_workflow", "protocol_design", "pooled_evidence_panel", "sample_size_waterfall"}:
                    score += 60
                return score, str(item.get("visual_id", ""))
            primary_visual_id = max(visual_candidates, key=visual_hero_score).get("visual_id")
        slide.update({
            "art_direction_id": art_direction["art_direction_id"],
            "layout_variant": _variant_for(slide, occurrence, rhythm["is_hero"]),
            "background_variant": rhythm["background_variant"],
            "visual_intensity": rhythm["visual_intensity"],
            "density_target": rhythm["density_target"],
            "hero_priority": rhythm["hero_priority"],
            "transition_role": rhythm["transition_role"],
            "section_role": rhythm["section_role"],
            "hero_visual": {
                "enabled": rhythm["is_hero"],
                "target_area_ratio": rhythm["hero_area_ratio"],
                "max_primary_anchors": 1,
                "max_auxiliary_annotations": 2,
                "primary_visual_id": primary_visual_id,
            },
            "presentation_revision": 25,
        })
        for visual in slide.get("visual_specs", []):
            fields: set[str] = set()
            for binding in visual.get("source_bindings", []):
                fields.update(binding.get("fields_used", []))
                fields.update(structured_fields_by_source.get(binding.get("source_id", ""), set()))
            requires_figure_decision = visual.get("visual_type") in {"source_figure", "umap_figure"}
            decision = (
                decide_figure_rebuild(visual, fields, bool(visual.get("payload", {}).get("image_path")))
                if requires_figure_decision else None
            )
            render_visual_type = visual.get("visual_type")
            if decision == "REDRAW_FROM_STRUCTURED_DATA" and visual.get("scientific_role") == "publication_bias_assessment":
                render_visual_type = "funnel_plot"
            elif decision == "REDRAW_FROM_STRUCTURED_DATA" and visual.get("visual_type") == "umap_figure":
                # Cell-level coordinates must not be embedded in a chart workbook.
                # The runner may only activate this route after a privacy-safe
                # aggregation or an explicitly approved de-identified subset.
                render_visual_type = "umap_scatter"
            visual["render_visual_type"] = render_visual_type
            visual.update({
                "render_mode": route_visual_render_mode(visual),
                "annotation_specs": derive_annotations(visual),
                "figure_rebuild_decision": decision,
                "presentation_variant": slide["layout_variant"],
                "chart_treatment": art_direction["chart_treatment"],
                "hero_area_ratio": rhythm["hero_area_ratio"],
                "preferred_backend": "native_pptxgenjs",
                "fallback_backend": "native_pptxgenjs_editable_shapes",
                "presentation_revision": 25,
            })
            errors = validate_annotation_specs(visual["annotation_specs"])
            if errors:
                raise ValueError(f"Invalid annotations for {visual['visual_id']}: {errors}")
        slide_annotations = [
            annotation
            for visual in slide.get("visual_specs", [])
            for annotation in visual.get("annotation_specs", [])
        ]
        errors = validate_annotation_specs(slide_annotations)
        if errors:
            raise ValueError(f"Invalid slide annotations for {slide['slide_id']}: {errors}")
    return result


def normalize_unicode_typography(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    text = unicodedata.normalize("NFC", value)
    text = CONTROL_CHARS.sub("", text).replace(REPLACEMENT_CHAR, "")
    text = text.replace("\u2012", "–").replace("\u2013", "–").replace("\u2014", "—")
    text = re.sub(r"\b(OR|HR|RR|CI|FDR|NES)\s*[:：]?\s*", lambda m: m.group(1) + " ", text)
    text = re.sub(r"(?<=\d)\s+(?=(?:h|d|天|小时|例|项|个)\b)", " ", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def validate_unicode_typography(texts: Iterable[str]) -> list[dict[str, str]]:
    issues = []
    for index, raw in enumerate(texts):
        text = str(raw)
        if CONTROL_CHARS.search(text):
            issues.append({"index": str(index), "issue": "invisible_unicode_control", "text": text})
        if REPLACEMENT_CHAR in text:
            issues.append({"index": str(index), "issue": "replacement_character", "text": text})
        if re.search(r"\b(?:M\s+E\s+T\s+H\s+O\s+D|R\s+E\s+S\s+U\s+L\s+T)\b", text, re.I):
            issues.append({"index": str(index), "issue": "abnormal_letter_spacing", "text": text})
        if len(text.splitlines()) > 2:
            issues.append({"index": str(index), "issue": "title_more_than_two_lines", "text": text})
    return issues


def empty_human_review_form(anonymous_code: str = "") -> dict[str, str]:
    return {
        "anonymous_deck_code": anonymous_code,
        "scientific_clarity_20": "",
        "visual_semantic_correctness_20": "",
        "hierarchy_15": "",
        "composition_15": "",
        "chart_professionalism_10": "",
        "consistency_8": "",
        "readability_5": "",
        "visual_richness_5": "",
        "presentation_readiness_2": "",
        "total_100": "",
        "reviewer": "",
        "review_date": "",
        "comments": "",
        "status": "PENDING_HUMAN_REVIEW",
        "completed": "",
    }


def verify_hash_manifest(before: dict[str, str], after: dict[str, str]) -> list[str]:
    errors: list[str] = []
    before_paths, after_paths = set(before), set(after)
    for path in sorted(before_paths - after_paths):
        errors.append(f"missing_after:{path}")
    for path in sorted(after_paths - before_paths):
        errors.append(f"new_after:{path}")
    for path in sorted(before_paths & after_paths):
        if before[path].lower() != after[path].lower():
            errors.append(f"hash_changed:{path}")
    return errors


HUMAN_REVIEW_SCORE_LIMITS = {
    "scientific_clarity_20": 20,
    "visual_semantic_correctness_20": 20,
    "hierarchy_15": 15,
    "composition_15": 15,
    "chart_professionalism_10": 10,
    "consistency_8": 8,
    "readability_5": 5,
    "visual_richness_5": 5,
    "presentation_readiness_2": 2,
}


def read_completed_human_reviews(paths: Iterable[Path]) -> list[dict[str, str]]:
    completed: list[dict[str, str]] = []
    for path in paths:
        if not path.is_file():
            continue
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle):
                try:
                    values = {field: float(row.get(field, "")) for field in HUMAN_REVIEW_SCORE_LIMITS}
                    total = float(row.get("total_100", ""))
                except (TypeError, ValueError):
                    continue
                if not all(0 <= values[field] <= limit for field, limit in HUMAN_REVIEW_SCORE_LIMITS.items()):
                    continue
                if abs(sum(values.values()) - total) > 0.01 or not 0 <= total <= 100:
                    continue
                if not row.get("reviewer", "").strip() or not row.get("review_date", "").strip():
                    continue
                if str(row.get("completed", "")).strip().lower() not in {"1", "true", "yes", "completed"}:
                    continue
                if str(row.get("status", "")).strip().upper() not in {"COMPLETED", "HUMAN_REVIEW_COMPLETED"}:
                    continue
                completed.append({**row, "review_path": str(path.resolve())})
    return completed


def visual_diversity_metrics(slides: list[dict[str, Any]]) -> dict[str, Any]:
    main = [slide for slide in slides if slide.get("appendix_status") != "appendix"]
    variants = [slide.get("layout_variant", "") for slide in main]
    backgrounds = [slide.get("background_variant", "") for slide in main]
    hero_count = sum(bool(slide.get("hero_visual", {}).get("enabled")) for slide in main)
    chart_led = sum(any(visual.get("render_mode") == "native_chart" or "plot" in visual.get("visual_type", "") or "forest" in visual.get("visual_type", "") for visual in slide.get("visual_specs", [])) for slide in main)
    figure_led = sum(any(visual.get("render_mode") in {"source_figure", "svg"} or visual.get("visual_type") in {"dag", "communication_network", "clone_censor_weight_branch", "protocol_design"} for visual in slide.get("visual_specs", [])) for slide in main)
    cards = sum(slide.get("layout_family") in {"comparison_matrix"} for slide in main)
    text_only = sum(not slide.get("visual_specs") for slide in main if slide.get("slide_role") != "cover")
    repeated = sum(variants[index] == variants[index - 1] for index in range(1, len(variants)))
    n = max(1, len(main))
    return {
        "main_slide_count": len(main),
        "background_variant_count": len(set(backgrounds)),
        "composition_variant_count": len(set(variants)),
        "hero_slide_count": hero_count,
        "figure_led_count": figure_led,
        "chart_led_count": chart_led,
        "card_based_count": cards,
        "text_only_count": text_only,
        "chart_led_ratio": round(chart_led / n, 4),
        "figure_led_ratio": round(figure_led / n, 4),
        "card_based_ratio": round(cards / n, 4),
        "text_only_ratio": round(text_only / n, 4),
        "repeated_structure_ratio": round(repeated / max(1, len(variants) - 1), 4),
    }


def validate_presentation_specs(slides: list[dict[str, Any]], art_direction: dict[str, Any], rhythm_plan: dict[str, Any]) -> list[str]:
    errors = [*validate_art_direction_spec(art_direction), *validate_deck_rhythm_plan(rhythm_plan)]
    for slide in slides:
        if slide.get("art_direction_id") != art_direction.get("art_direction_id"):
            errors.append(f"art_direction_mismatch:{slide.get('slide_id')}")
        if not slide.get("layout_variant"):
            errors.append(f"layout_variant_missing:{slide.get('slide_id')}")
        hero = slide.get("hero_visual", {})
        if hero.get("enabled") and not 0.35 <= float(hero.get("target_area_ratio", 0)) <= 0.65:
            errors.append(f"hero_area_outside_35_65pct:{slide.get('slide_id')}")
        if hero.get("enabled"):
            visual_ids = [visual.get("visual_id") for visual in slide.get("visual_specs", [])]
            if hero.get("primary_visual_id") not in visual_ids:
                errors.append(f"hero_primary_visual_missing:{slide.get('slide_id')}")
            if int(hero.get("max_primary_anchors", 0)) != 1:
                errors.append(f"hero_primary_anchor_contract_invalid:{slide.get('slide_id')}")
            if int(hero.get("max_auxiliary_annotations", -1)) not in {0, 1, 2}:
                errors.append(f"hero_auxiliary_annotation_contract_invalid:{slide.get('slide_id')}")
        slide_annotations = [
            annotation
            for visual in slide.get("visual_specs", [])
            for annotation in visual.get("annotation_specs", [])
        ]
        errors.extend(
            f"{slide.get('slide_id')}:{error}"
            for error in validate_annotation_specs(slide_annotations)
        )
        for visual in slide.get("visual_specs", []):
            if visual.get("preferred_backend") != "native_pptxgenjs":
                errors.append(f"backend_not_native:{visual.get('visual_id')}")
            errors.extend(f"{visual.get('visual_id')}:{error}" for error in validate_annotation_specs(visual.get("annotation_specs", [])))
    return errors
