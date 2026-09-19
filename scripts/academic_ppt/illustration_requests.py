"""Default-off, local-only controlled illustration request builder. No generation or deck dispatch.

The caller supplies a source-bound planning record; this module projects only
controlled enums, fixed conceptual summaries and validated numeric regions.
"""
from __future__ import annotations

import copy
import hashlib
import json
import math
import re
from pathlib import Path

VERSION = "content-aware-illustration-request/1"
SOURCE_PRIORITY = ["EXISTING_SCIENTIFIC_SOURCE_FIGURE", "APPROVED_PROJECT_ASSET",
                   "APPROVED_DETERMINISTIC_LIBRARY", "AI_CONCEPTUAL_CANDIDATE", "NONE"]
PALETTE = ["#5B7F72", "#7EA7B3", "#C97F63", "#C9B28A", "#F6F2EA", "#E8E1D7", "#1F3A36", "#4D5B58"]
STYLE = "WARM_ACADEMIC_EDITORIAL"
STYLES = (STYLE, "RESTRAINED_EDITORIAL", "CLINICAL_RESTRAINED")
PROHIBITED = ["ECG", "EEG", "ULTRASOUND", "CT", "MRI", "PATHOLOGY", "PATIENT_IMAGING",
              "STATISTICAL_CHART", "REAL_MONITORING_DATA", "CLINICAL_EVIDENCE", "PATIENT_IDENTIFIERS",
              "GENERATED_TEXT_OR_NUMBERS", "BRAIN_LESION", "VASCULAR_ANATOMY", "HEMISPHERE_FUNCTION_DIFFERENCES",
              "OXYGEN_MAP", "PERFUSION_IMAGE", "VENTILATOR_WAVEFORM", "BLOOD_GAS_DISPLAY",
              "TREATMENT_OUTCOME", "MANDATORY_PROTOCOL", "NEW_STAGE_OR_TIMEPOINT", "NEW_CAUSAL_RELATION",
              "AI_BCI_EFFICACY", "NEON", "HEAVY_GRADIENT", "3D", "GLASSMORPHISM", "CARTOON",
              "STOCK_PHOTO", "BUSY_HOSPITAL_COLLAGE", "SPECIFIC_GEOGRAPHY", "ALTITUDE_MEASUREMENT",
              "PATIENT_EXPOSURE", "PROPORTIONAL_HETEROGENEOUS_VALUES"]

# These are conceptual art descriptions, not text extracted from a clinical file.
PROFILES = {
    "EDITORIAL_HORIZON": {
        "summary": "Abstract editorial horizon with layered organic forms and restrained texture; no real geography or measured quantities.",
        "boundary": "ABSTRACT_ATMOSPHERE_ONLY", "role": "BACKGROUND_ATMOSPHERE",
        "composition": "RIGHT_LOWER_SCENE_LEFT_TITLE", "weight": "SUBORDINATE",
        "depth": "FOREGROUND_MIDGROUND_BACKGROUND", "types": ["ABSTRACT_HORIZON", "EDITORIAL_ATMOSPHERE"],
        "reason": "Support the opening atmosphere while preserving a large title area and native text dominance."},
    "RESEARCH_PROCESS": {
        "summary": "A coherent editorial icon family for research search, selection, consultation, review and document output; equal scale and no data marks.",
        "boundary": "NATIVE_SEQUENCE_ONLY", "role": "PROCESS_ICON_FAMILY",
        "composition": "EQUAL_CUES_ABOVE_NATIVE_SEQUENCE", "weight": "LOW",
        "depth": "FLAT_EDITORIAL_CUES", "types": ["SEARCH_CUE", "DOCUMENT_CUE", "CONSULTATION_CUE", "REVIEW_CUE"],
        "reason": "Differentiate process roles with equal visual weight; all counts and relations remain native."},
    "CONCEPTUAL_FRAMEWORK": {
        "summary": "An abstract conceptual silhouette with low-contrast layers and balanced category cues; no anatomical structures, diagnostic imaging or implied directional relation.",
        "boundary": "NATIVE_EQUAL_DIMENSIONS_ONLY", "role": "ABSTRACT_FRAMEWORK_ATMOSPHERE",
        "composition": "SUBORDINATE_CENTER_EQUAL_NATIVE_DIMENSIONS", "weight": "LOW",
        "depth": "LOW_CONTRAST_LAYERED", "types": ["ABSTRACT_CONCEPTUAL_SILHOUETTE", "BALANCED_CATEGORY_CUES"],
        "reason": "Support the existing central concept without anatomical claims or unequal emphasis among dimensions."},
    "IMPLEMENTATION_CUES": {
        "summary": "Restrained editorial cues for safety, coordination, implementation checks and team support; no monitor display, waveform or treatment outcome.",
        "boundary": "NATIVE_PARAMETERS_AND_LIMITATIONS_ONLY", "role": "IMPLEMENTATION_CUE_FAMILY",
        "composition": "MARGIN_CUES_NATIVE_PARAMETERS_DOMINANT", "weight": "LOW",
        "depth": "FLAT_EDITORIAL_CUES", "types": ["SAFETY_CUE", "COORDINATION_CUE", "CHECK_CUE"],
        "reason": "Keep native parameters, units and limitations dominant; optional cues do not prescribe a protocol."},
    "CONTINUITY_BACKGROUND": {
        "summary": "Subtle editorial atmosphere suggesting continuity, a journey and multidisciplinary support; no new stage, time point, recovery outcome or technology efficacy.",
        "boundary": "NATIVE_TIMELINE_ONLY_FUTURE_EXPLORATION", "role": "CONTINUUM_ATMOSPHERE",
        "composition": "LOWER_ATMOSPHERE_NATIVE_TIMELINE_DOMINANT", "weight": "LOW",
        "depth": "FOREGROUND_MIDGROUND_BACKGROUND", "types": ["ABSTRACT_CONTINUITY_ATMOSPHERE"],
        "reason": "Support continuity while leaving the native timeline authoritative and future exploration unpromoted."},
    "CLOSING_EQUAL_STATEMENTS": {
        "summary": "No illustration: preserve equal emphasis among the existing closing statements.",
        "boundary": "NATIVE_EQUAL_STATEMENTS_ONLY", "role": "NONE", "composition": "NATIVE_CONTENT_ONLY",
        "weight": "NONE", "depth": "NONE", "types": [],
        "reason": "Illustration would compete with equal statements; microcopy remains unapplied."},
    "NUMERIC_NATIVE_ANCHOR": {
        "summary": "No illustration: the existing native numeric concept is the visual anchor.",
        "boundary": "NATIVE_NUMERIC_CONCEPT_ONLY", "role": "NONE", "composition": "NATIVE_CONTENT_ONLY",
        "weight": "NONE", "depth": "NONE", "types": [],
        "reason": "A strong native numeric anchor needs no decorative artwork; retain the negative control."},
    "UNRESOLVED_SEMANTICS": {
        "summary": "No visual execution: unresolved semantic mapping requires a human decision.",
        "boundary": "HUMAN_SEMANTIC_DECISION", "role": "NONE", "composition": "NO_VISUAL_EXECUTION",
        "weight": "NONE", "depth": "NONE", "types": [],
        "reason": "Unresolved semantics prohibit a decorative substitute or a newly implied relationship."},
}
ROLES = ["HERO", "CONTEXT", "EVIDENCE", "FRAMEWORK", "RESULTS", "COMPARISON", "IMPLEMENTATION", "TIMELINE", "CLOSING", "APPENDIX"]
STATUS = ["ILLUSTRATION_CANDIDATE_PENDING", "NO_ILLUSTRATION_NEEDED", "NO_VISUAL_EXECUTION",
          "SOURCE_FIGURE_REQUIRED", "REUSE_SOURCE_FIGURE", "REUSE_APPROVED_PROJECT_ASSET", "REUSE_APPROVED_LIBRARY"]


def canonical_hash(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()).hexdigest()


def require(condition, message):
    if not condition:
        raise ValueError(message)


def closed(value, keys, label):
    require(isinstance(value, dict) and set(value) == set(keys), "CLOSED_FIELDS_REQUIRED: " + label)


def rectangle(region):
    closed(region, ("x", "y", "w", "h"), "region")
    require(all(type(v) in (int, float) and math.isfinite(v) for v in region.values()), "FINITE_REGION_REQUIRED")
    require(0 <= region["x"] < 1 and 0 <= region["y"] < 1 and 0 < region["w"] <= 1 and 0 < region["h"] <= 1
            and region["x"] + region["w"] <= 1.0000001 and region["y"] + region["h"] <= 1.0000001,
            "NORMALIZED_REGION_REQUIRED")


def overlap(a, b):
    return min(a["x"]+a["w"], b["x"]+b["w"]) - max(a["x"], b["x"]) > .000001 and min(a["y"]+a["h"], b["y"]+b["h"]) - max(a["y"], b["y"]) > .000001


def validate_style(style, palette):
    require(style in STYLES, "CONTROLLED_STYLE_REQUIRED")
    require(isinstance(palette, list) and 2 <= len(palette) <= 12 and all(isinstance(c, str) and re.fullmatch(r"#[0-9a-fA-F]{6}", c) for c in palette), "CONTROLLED_PALETTE_REQUIRED")


def build_request(record, context, brief, sources, *, enabled=False):
    """No file/network access. A disabled call returns before examining inputs."""
    if not enabled:
        return None
    closed(context, ("subject_class", "semantic_boundary", "focal_region", "negative_space_region", "source_figure_required"), "context")
    closed(brief, ("approved", "style_family", "palette"), "brief")
    require(brief["approved"] is True, "APPROVED_STYLE_REQUIRED")
    validate_style(brief["style_family"], brief["palette"])
    closed(sources, ("scientific_source_figure", "approved_project_asset", "approved_deterministic_library"), "source availability")
    require(all(type(v) is bool for v in sources.values()), "BOOLEAN_SOURCE_STATUS_REQUIRED")
    require(type(context["source_figure_required"]) is bool, "BOOLEAN_SOURCE_REQUIREMENT_REQUIRED")
    subject = context["subject_class"]
    require(subject in PROFILES, "CONTROLLED_SUBJECT_REQUIRED")
    profile = PROFILES[subject]
    require(context["semantic_boundary"] == profile["boundary"], "SEMANTIC_BOUNDARY_MISMATCH")
    require(bool(re.fullmatch(r"SLD-[A-F0-9]{20}|S[0-9]{3,6}", record["slide_id"])), "OPAQUE_SLIDE_ID_REQUIRED")
    require(record["page_role"] in ROLES and record["visual_anchor_role"] in ("ANCHOR", "SUPPORTING", "UTILITY")
            and record["recommended_ambition_level"] in ("FLOOR", "TARGET", "STRETCH"), "PLANNER_ENUM_INVALID")
    require(record["enrichment_opportunity"] in ("NONE", "LOW", "MEDIUM", "HIGH", "REFERENCE_GRADE_OPPORTUNITY"), "PLANNER_OPPORTUNITY_INVALID")
    require(record["quality_floor_status"] in ("PASS", "WARN", "HUMAN_REVIEW", "BLOCK"), "FLOOR_STATUS_INVALID")
    require(record["deck_quality_floor_status"] in ("VISUAL_QUALITY_PASS", "VISUAL_QUALITY_PASS_WITH_WARNINGS", "VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED", "VISUAL_QUALITY_BLOCKED"), "FLOOR_STATUS_INVALID")
    require(record["source_priority"] == SOURCE_PRIORITY, "SOURCE_PRIORITY_IMMUTABLE")
    for region in (context["focal_region"], context["negative_space_region"]):
        if region is not None:
            rectangle(region)
    require(context["negative_space_region"] is not None, "NEGATIVE_SPACE_REQUIRED")
    if context["focal_region"] is not None:
        require(not overlap(context["focal_region"], context["negative_space_region"]), "FOCAL_INTRUDES_NEGATIVE_SPACE")
    source_required = context["source_figure_required"] or record["source_figure_required"]
    unresolved = record["recommended_next_action"] == "HUMAN_SEMANTIC_DECISION"
    require(unresolved == (subject == "UNRESOLVED_SEMANTICS"), "UNRESOLVED_SEMANTICS_CANNOT_BE_RECLASSIFIED")
    if unresolved:
        status, selected = "NO_VISUAL_EXECUTION", "NONE"
    elif source_required:
        status, selected = "SOURCE_FIGURE_REQUIRED", "EXISTING_SCIENTIFIC_SOURCE_FIGURE"
    elif profile["role"] == "NONE":
        status, selected = "NO_ILLUSTRATION_NEEDED", "NONE"
    elif sources["scientific_source_figure"]:
        status, selected = "REUSE_SOURCE_FIGURE", SOURCE_PRIORITY[0]
    elif sources["approved_project_asset"]:
        status, selected = "REUSE_APPROVED_PROJECT_ASSET", SOURCE_PRIORITY[1]
    elif sources["approved_deterministic_library"]:
        status, selected = "REUSE_APPROVED_LIBRARY", SOURCE_PRIORITY[2]
    else:
        status, selected = "ILLUSTRATION_CANDIDATE_PENDING", SOURCE_PRIORITY[3]
    needed = status not in ("NO_VISUAL_EXECUTION", "NO_ILLUSTRATION_NEEDED", "SOURCE_FIGURE_REQUIRED", "REUSE_SOURCE_FIGURE")
    require(not needed or context["focal_region"] is not None, "FOCAL_REGION_REQUIRED")
    result = {
        "schema_version": VERSION, "slide_id": record["slide_id"], "page_role": record["page_role"],
        "anchor_role": record["visual_anchor_role"], "ambition_level": record["recommended_ambition_level"],
        "enrichment_opportunity": record["enrichment_opportunity"], "illustration_needed": needed,
        "illustration_role": profile["role"] if needed else "NONE", "subject_class": subject,
        "subject_summary": profile["summary"], "semantic_boundary": profile["boundary"],
        "style_family": brief["style_family"], "palette": copy.deepcopy(brief["palette"]), "composition_role": profile["composition"],
        "focal_region": copy.deepcopy(context["focal_region"] if needed else None),
        "negative_space_region": copy.deepcopy(context["negative_space_region"]),
        "visual_weight": profile["weight"] if needed else "NONE", "depth_role": profile["depth"] if needed else "NONE",
        "allowed_visual_types": copy.deepcopy(profile["types"] if needed else []), "prohibited_visual_types": copy.deepcopy(PROHIBITED),
        "scientific_evidence": False, "presentation_only": True, "patient_data_included": False,
        "human_approval_required": True, "human_approved": False, "source_priority": copy.deepcopy(SOURCE_PRIORITY),
        "source_selection": selected, "source_figure_required": bool(source_required),
        "output_preference": "SEPARATELY_REPLACEABLE_PNG_CANDIDATE" if needed else "NONE",
        "aspect_ratio": "16:9", "transparency_preference": "TRANSPARENT_PREFERRED" if needed else "NOT_APPLICABLE",
        "reasoning_summary": profile["reason"], "status": status,
        "quality_floor_status": record["quality_floor_status"], "deck_quality_floor_status": record["deck_quality_floor_status"],
        "original_illustration_decision": record["illustration_decision"],
        "floor_override_allowed": False, "production_insertion_allowed": False, "external_dispatch_authorized": False,
        "generation_authorized": False, "candidate_only": True, "canonical": False,
        "microcopy_applied": False, "semantic_graph_change_allowed": False,
        "native_content_occlusion_allowed": False, "native_text_only": True,
    }
    result["request_id"] = "IRQ-" + canonical_hash(result)[:24].upper()
    validate_request(result)
    return result


def validate_request(request):
    """Validate every public field against finite contracts; no free text slot."""
    require(isinstance(request, dict), "REQUEST_OBJECT_REQUIRED")
    schema = request_schema()
    closed(request, schema["required"], "request")
    for name, spec in schema["properties"].items():
        value = request[name]
        if "const" in spec:
            require(type(value) is type(spec["const"]) and value == spec["const"], "FIXED_CONTRACT: " + name)
        if "enum" in spec:
            require(value in spec["enum"], "CONTROLLED_FIELD: " + name)
        if spec.get("type") == "boolean":
            require(type(value) is bool, "BOOLEAN_REQUIRED: " + name)
        if "pattern" in spec:
            require(isinstance(value, str) and bool(re.fullmatch(spec["pattern"], value)), "OPAQUE_ID_REQUIRED: " + name)
    validate_style(request["style_family"], request["palette"])
    profile = PROFILES[request["subject_class"]]
    for field, key in (("subject_summary", "summary"), ("reasoning_summary", "reason"), ("semantic_boundary", "boundary"), ("composition_role", "composition")):
        require(request[field] == profile[key], "CONTROLLED_TEMPLATE_MISMATCH: " + field)
    needed = request["illustration_needed"]
    require(profile["role"] != "NONE" or not needed, "NO_FORCED_ILLUSTRATION")
    for field, key, absent in (("illustration_role", "role", "NONE"), ("depth_role", "depth", "NONE"),
                               ("visual_weight", "weight", "NONE"), ("allowed_visual_types", "types", [])):
        require(request[field] == (profile[key] if needed else absent), "SUBJECT_CONTRACT_MISMATCH: " + field)
    rectangle(request["negative_space_region"])
    if needed:
        rectangle(request["focal_region"])
        require(not overlap(request["focal_region"], request["negative_space_region"]), "FOCAL_INTRUDES_NEGATIVE_SPACE")
    else:
        require(request["focal_region"] is None, "NO_FOCAL_REGION_FOR_NO_ILLUSTRATION")
    require(needed == (request["status"] in ("ILLUSTRATION_CANDIDATE_PENDING", "REUSE_APPROVED_PROJECT_ASSET", "REUSE_APPROVED_LIBRARY")), "STATUS_NEED_MISMATCH")
    require(request["output_preference"] == ("SEPARATELY_REPLACEABLE_PNG_CANDIDATE" if needed else "NONE")
            and request["transparency_preference"] == ("TRANSPARENT_PREFERRED" if needed else "NOT_APPLICABLE"), "OUTPUT_PREFERENCE_MISMATCH")
    if request["source_figure_required"]:
        require(request["status"] in ("SOURCE_FIGURE_REQUIRED", "NO_VISUAL_EXECUTION") and not needed, "SOURCE_FIGURE_CANNOT_BE_REPLACED")
    if request["subject_class"] == "UNRESOLVED_SEMANTICS":
        require(request["status"] == "NO_VISUAL_EXECUTION", "NO_SEMANTIC_EXECUTION")
    source_by_status = dict(zip(STATUS, [SOURCE_PRIORITY[3], "NONE", "NONE", SOURCE_PRIORITY[0], SOURCE_PRIORITY[0], SOURCE_PRIORITY[1], SOURCE_PRIORITY[2]]))
    require(request["source_selection"] == source_by_status[request["status"]], "SOURCE_SELECTION_MISMATCH")
    require(request["request_id"] == "IRQ-" + canonical_hash({k: v for k, v in request.items() if k != "request_id"})[:24].upper(), "REQUEST_HASH_MISMATCH")


def request_schema():
    rect = {"type": "object", "additionalProperties": False, "required": ["x", "y", "w", "h"],
            "properties": {k: {"type": "number", "minimum": 0, "maximum": 1} for k in ("x", "y", "w", "h")}}
    props = {
        "schema_version": {"const": VERSION}, "request_id": {"type": "string", "pattern": r"^IRQ-[A-F0-9]{24}$"},
        "slide_id": {"type": "string", "pattern": r"^(SLD-[A-F0-9]{20}|S[0-9]{3,6})$"},
        "page_role": {"enum": ROLES}, "anchor_role": {"enum": ["ANCHOR", "SUPPORTING", "UTILITY"]},
        "ambition_level": {"enum": ["FLOOR", "TARGET", "STRETCH"]},
        "enrichment_opportunity": {"enum": ["NONE", "LOW", "MEDIUM", "HIGH", "REFERENCE_GRADE_OPPORTUNITY"]},
        "illustration_needed": {"type": "boolean"}, "source_figure_required": {"type": "boolean"},
        "subject_class": {"enum": list(PROFILES)}, "style_family": {"enum": list(STYLES)}, "palette": {"type": "array", "minItems": 2, "maxItems": 12, "items": {"type": "string", "pattern": "^#[0-9A-Fa-f]{6}$"}},
        "source_priority": {"const": SOURCE_PRIORITY}, "prohibited_visual_types": {"const": PROHIBITED},
        "source_selection": {"enum": SOURCE_PRIORITY}, "status": {"enum": STATUS},
        "focal_region": {"anyOf": [rect, {"type": "null"}]}, "negative_space_region": rect,
        "aspect_ratio": {"const": "16:9"}, "output_preference": {"enum": ["SEPARATELY_REPLACEABLE_PNG_CANDIDATE", "NONE"]},
        "transparency_preference": {"enum": ["TRANSPARENT_PREFERRED", "NOT_APPLICABLE"]},
        "quality_floor_status": {"enum": ["PASS", "WARN", "HUMAN_REVIEW", "BLOCK"]},
        "deck_quality_floor_status": {"enum": ["VISUAL_QUALITY_PASS", "VISUAL_QUALITY_PASS_WITH_WARNINGS", "VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED", "VISUAL_QUALITY_BLOCKED"]},
        "original_illustration_decision": {"enum": ["APPROVED_ASSET_SELECTED", "NO_ILLUSTRATION_NEEDED", "SOURCE_FIGURE_REQUIRED", "AI_CANDIDATE_ALLOWED", "HUMAN_REVIEW_REQUIRED"]},
    }
    for field, key in (("subject_summary", "summary"), ("semantic_boundary", "boundary"), ("reasoning_summary", "reason"),
                       ("composition_role", "composition"), ("illustration_role", "role"), ("depth_role", "depth"), ("visual_weight", "weight")):
        props[field] = {"enum": sorted({p[key] for p in PROFILES.values()})}
    props["allowed_visual_types"] = {"enum": [p["types"] for p in PROFILES.values()]}
    for field in ("scientific_evidence", "patient_data_included", "human_approved", "floor_override_allowed", "production_insertion_allowed", "external_dispatch_authorized", "generation_authorized", "canonical", "microcopy_applied", "semantic_graph_change_allowed", "native_content_occlusion_allowed"):
        props[field] = {"const": False}
    for field in ("presentation_only", "human_approval_required", "candidate_only", "native_text_only"):
        props[field] = {"const": True}
    return {"$schema": "https://json-schema.org/draft/2020-12/schema", "title": VERSION,
            "description": "Closed export-safe data. Runtime validation additionally binds templates, geometry, source selection and request ID.",
            "type": "object", "additionalProperties": False, "required": sorted(props), "properties": props}


def validate_candidate(asset, request, preview_authorization, *, file_sha256):
    """Local registry check; approval to preview is never scientific approval."""
    validate_request(request)
    closed(asset, ("candidate_id", "slide_id", "request_id", "request_hash", "asset_sha256", "kind", "status",
                   "scientific_evidence", "presentation_only", "human_approved", "generator_identity", "generation_timestamp"), "candidate")
    closed(preview_authorization, ("slide_id", "request_id", "asset_sha256", "candidate_preview_approved", "production_approved"), "preview authorization")
    require(request["illustration_needed"], "NO_ASSET_FOR_NO_ILLUSTRATION")
    require(asset["kind"] in ("PLACEHOLDER", "GENERATED_CANDIDATE"), "CANDIDATE_KIND_INVALID")
    require(asset["status"] == "CANDIDATE" and asset["scientific_evidence"] is False and asset["presentation_only"] is True and asset["human_approved"] is False, "CANDIDATE_ONLY_REQUIRED")
    require(asset["slide_id"] == request["slide_id"] == preview_authorization["slide_id"] and asset["request_id"] == request["request_id"] == preview_authorization["request_id"], "SLIDE_AUTHORIZATION_MISMATCH")
    require(asset["request_hash"] == canonical_hash(request), "ASSET_REQUEST_HASH_MISMATCH")
    require(bool(re.fullmatch(r"[a-f0-9]{64}", file_sha256)) and asset["asset_sha256"] == preview_authorization["asset_sha256"] == file_sha256, "ASSET_HASH_MISMATCH")
    require(preview_authorization["candidate_preview_approved"] is True and preview_authorization["production_approved"] is False, "CANDIDATE_PREVIEW_AUTHORIZATION_REQUIRED")
    if asset["kind"] == "GENERATED_CANDIDATE":
        require(request["generation_authorized"] is True, "GENERATION_NOT_AUTHORIZED")
        require(bool(asset["generator_identity"]) and bool(asset["generation_timestamp"]), "GENERATION_PROVENANCE_REQUIRED")
    else:
        require(asset["generator_identity"] == "native_pptxgenjs_request_box" and asset["generation_timestamp"] is None, "PLACEHOLDER_NOT_GENERATED")


def write_json_new(path, value):
    path = Path(path)
    with path.open("x", encoding="utf8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2, allow_nan=False)
        stream.write("\n")
