"""Independent, default-off ambition planning over hash-bound floor outputs.

No rendering, scientific reassessment, request construction, or deck modification.
Use --enabled to opt into local read-only planning.
"""
from __future__ import annotations

import argparse
from collections import Counter
import copy
import csv
import json
from pathlib import Path
import re

from .visual_quality_floor import ROOT, digest, pointer_value, read_json, sha256

CONFIG = ROOT / "config/visual_ambition"
FLOOR_CONFIG = ROOT / "config/visual_quality"
VERSION = "visual-ambition/1"
LEVELS = ("FLOOR", "TARGET", "STRETCH")
OPPORTUNITIES = ("NONE", "LOW", "MEDIUM", "HIGH", "REFERENCE_GRADE_OPPORTUNITY")
CHANNELS = ("COMPOSITION_UPGRADE", "ILLUSTRATION_ENRICHMENT", "ICONOGRAPHY", "TYPOGRAPHY_UPGRADE",
            "EDITORIAL_COMPONENTS", "MICROCOPY_OPPORTUNITY", "BACKGROUND_SCENE", "ANNOTATION_ENRICHMENT")
ACTIONS = ("NO_ACTION", "COMPOSITION_REVIEW", "ILLUSTRATION_CANDIDATE", "ICON_SET_CANDIDATE",
           "ART_DIRECTION_REVIEW", "MICROCOPY_PROPOSAL", "REFERENCE_GRADE_SCENE_TRIAL",
           "HUMAN_SEMANTIC_DECISION", "SOURCE_FIGURE_REQUIRED")
MICROCOPY = ("MICROCOPY_NOT_NEEDED", "MICROCOPY_MAY_HELP", "MICROCOPY_HIGH_VALUE", "MICROCOPY_UNSAFE")
NARRATIVE = {
    "OPENING": ("ANCHOR", "The opening establishes the presentation's visual direction."),
    "EVIDENCE_METHOD": ("ANCHOR", "The evidence/method page carries the source formation story and its limits."),
    "CORE_FRAMEWORK": ("ANCHOR", "A source-bound core framework organizes the narrative."),
    "KEY_RESULT": ("ANCHOR", "A registered key result carries the central evidence message."),
    "KEY_CONCEPT": ("ANCHOR", "A key concept is a narrative reference point, not necessarily a measured result."),
    "MAJOR_IMPLEMENTATION": ("ANCHOR", "A major implementation page connects the framework with its conditions of use."),
    "MAJOR_TIMELINE": ("ANCHOR", "A major timeline organizes stages without inventing durations or relations."),
    "TAKEAWAYS": ("ANCHOR", "The closing supports recall while retaining equal statements and limitations."),
    "SUPPORTING_CONTEXT": ("SUPPORTING", "Context supports the narrative without requiring a scene treatment."),
    "SUPPORTING_CONCEPT": ("SUPPORTING", "This concept supports the main framework; unresolved meaning is retained."),
    "SUPPORTING_IMPLEMENTATION": ("SUPPORTING", "These detailed actions support implementation and retain their limits."),
    "UTILITY": ("UTILITY", "Reference/navigation utility is the primary purpose; no enrichment quota applies."),
}
NEEDS = ("ALREADY_CLEAR", "READING_ORDER", "ROLE_DISTINCTION", "SCENE_INTEGRATION",
         "LABEL_HIERARCHY", "RELATIONSHIP_UNRESOLVED", "SOURCE_FIGURE_GAP")
NEED_REASON = {
    "ALREADY_CLEAR": "The existing supporting composition is clear; further decoration may add little communication value.",
    "READING_ORDER": "Review grouping, hierarchy and whitespace within the existing source-bound structure.",
    "ROLE_DISTINCTION": "Make this communicative purpose more distinct while preserving every source statement and relation.",
    "SCENE_INTEGRATION": "Review the interaction of native content and subordinate scene elements against the human reference benchmark.",
    "LABEL_HIERARCHY": "Review the hierarchy of labels and conditional wording without deleting or changing protected fragments.",
    "RELATIONSHIP_UNRESOLVED": "Potential communication gain is conditional on the author resolving the existing semantic restriction.",
    "SOURCE_FIGURE_GAP": "A genuine source figure is required; decorative or generated imagery cannot provide that evidence.",
}
SUBJECTS = {
    "ABSTRACT_MEDICAL_CONCEPT": "Abstract conceptual support; no clinical evidence, new mechanism, or diagnostic implication.",
    "BACKGROUND_SCENE": "Muted editorial background atmosphere with protected space for native content; no geographic evidence.",
    "ORGAN_SILHOUETTE": "Subordinate conceptual silhouette; no anatomical truth, laterality, diagnostic or causal claim.",
    "RESEARCH_ENVIRONMENT": "Abstract research-process editorial cues; no invented study results or evidence-strength badges.",
    "CARE_PATHWAY_ATMOSPHERE": "Abstract care-pathway atmosphere; preserve source-defined stages and conditions.",
    "REHABILITATION_CONCEPT": "Abstract recovery-stage atmosphere; no promised outcome, patient portrayal, or invented duration.",
    "EDITORIAL_ILLUSTRATION": "Restrained editorial cue supporting reading order; no scientific or clinical evidence.",
}
PROTECTED = ("NUMBERS", "UNITS", "NEGATIONS", "CONDITIONS", "RECOMMENDATION_STRENGTH", "CAUSAL_LANGUAGE")
TRAITS = ("visual_depth", "illustration_richness", "editorial_hierarchy", "composition_variety",
          "scene_integration", "component_vocabulary", "intentional_whitespace")
OUTPUTS = ("visual_ambition_plan.json", "visual_ambition_plan.md", "slide_visual_opportunities.json",
           "anchor_slide_plan.json", "illustration_opportunity_report.json", "microcopy_opportunity_report.json",
           "composition_opportunity_report.json", "human_ambition_review_queue.csv", "quality_floor_plus_ambition_summary.json")
CONTEXT_FIELDS = {"slide_id", "narrative_function", "communication_need", "illustration_potential", "subject_class",
                  "space_status", "reference_scene_opportunity", "source_figure_required", "requested_visual_type",
                  "microcopy_requires_protected_rewrite", "repeated_wording", "negative_space_hint", "evidence_refs", "override"}


def authority():
    lock = read_json(CONFIG / "authority.json")
    from .visual_quality_floor import load_authority
    load_authority()
    matrix = read_json(FLOOR_CONFIG / "quality_floor_QA_matrix.json")
    safety = read_json(FLOOR_CONFIG / "ai_illustration_safety_contract.json")
    priority = read_json(FLOOR_CONFIG / "illustration_decision.schema.json")["properties"]["source_priority"]["const"]
    return lock, {c["check_id"]: c["severity"] for c in matrix["checks"]}, safety, priority


def validate_context(context):
    required = CONTEXT_FIELDS - {"override", "negative_space_hint", "requested_visual_type"}
    if set(context) - CONTEXT_FIELDS or required - set(context):
        raise ValueError("INVALID_AMBITION_CONTEXT_FIELDS")
    choices = {"narrative_function": NARRATIVE, "communication_need": NEEDS,
               "illustration_potential": OPPORTUNITIES, "subject_class": (*SUBJECTS, "NONE"),
               "space_status": ("AVAILABLE_CANDIDATE", "TIGHT", "NOT_ASSESSED"),
               "negative_space_hint": (None, "LEFT_CANDIDATE_REQUIRES_REVIEW", "RIGHT_CANDIDATE_REQUIRES_REVIEW",
                                        "BACKGROUND_CANDIDATE_REQUIRES_REVIEW")}
    for field, allowed in choices.items():
        if context.get(field) not in allowed:
            raise ValueError("INVALID_AMBITION_CONTEXT_ENUM: " + field)
    for field in ("reference_scene_opportunity", "source_figure_required", "microcopy_requires_protected_rewrite", "repeated_wording"):
        if type(context[field]) is not bool:
            raise ValueError("INVALID_AMBITION_CONTEXT_BOOLEAN: " + field)
    if context["illustration_potential"] != "NONE" and context["subject_class"] == "NONE":
        raise ValueError("ILLUSTRATION_SUBJECT_CLASS_REQUIRED")
    if not re.fullmatch(r"(?:SLD-[A-F0-9]{20}|S[0-9]{3,6})", context["slide_id"]):
        raise ValueError("UNSAFE_HANDOFF_SLIDE_ID: use an opaque workflow ID")
    if not isinstance(context["evidence_refs"], list) or not context["evidence_refs"]:
        raise ValueError("MISSING_AMBITION_REASONING_PROVENANCE")
    override = context.get("override", {})
    if not isinstance(override, dict) or set(override) - {
            "recommended_ambition_level", "primary_enrichment_channels", "secondary_enrichment_channels", "recommended_next_action"}:
        raise ValueError("INVALID_AMBITION_OVERRIDE")
    for key in ("primary_enrichment_channels", "secondary_enrichment_channels"):
        if key in override:
            values = override[key]
            if (not isinstance(values, list) or any(v not in CHANNELS for v in values) or
                    len(set(values)) != len(values) or key == "primary_enrichment_channels" and len(values) > 2):
                raise ValueError("INVALID_PRIMARY_ENRICHMENT_CHANNELS: select 0-2 unique channels")


def floor_status(findings):
    if any(f["severity"] == "BLOCK" for f in findings):
        return "BLOCK"
    if any(f["severity"] == "HUMAN_REVIEW" for f in findings):
        return "HUMAN_REVIEW"
    return "WARN" if findings else "PASS"


def audit_ambition(records):
    """Audit the second axis only. Never emits a Quality Floor finding or BLOCK."""
    findings = []

    def flag(code, row, severity="WARN"):
        findings.append({"code": code, "severity": severity, "slide_id": row["slide_id"] if row else None,
                         "axis": "VISUAL_AMBITION", "floor_override_allowed": False})

    for r in records:
        primary, secondary = r["primary_enrichment_channels"], r["secondary_enrichment_channels"]
        if (len(primary) > 2 or len(set(primary)) != len(primary) or
                any(c not in CHANNELS for c in primary+secondary) or set(primary) & set(secondary)):
            raise ValueError("INVALID_PRIMARY_ENRICHMENT_CHANNELS: select 0-2 distinct primary channels")
        if r["recommended_ambition_level"] not in LEVELS or r["recommended_next_action"] not in ACTIONS:
            raise ValueError("INVALID_AMBITION_LEVEL_OR_ACTION")
        high = r["enrichment_opportunity"] in ("HIGH", "REFERENCE_GRADE_OPPORTUNITY")
        if high and r["recommended_ambition_level"] == "FLOOR":
            flag("AMBITION_TOO_LOW", r)
        if r["is_anchor"] and high and (not primary or r["recommended_ambition_level"] == "FLOOR"):
            flag("ANCHOR_SLIDE_UNDERDESIGNED", r)
        visual = set(primary) & {"ILLUSTRATION_ENRICHMENT", "ICONOGRAPHY", "BACKGROUND_SCENE"}
        if r["illustration_enrichment_opportunity"] in ("HIGH", "REFERENCE_GRADE_OPPORTUNITY") and not visual and not r["illustration_deferral_reason"]:
            flag("ILLUSTRATION_OPPORTUNITY_IGNORED", r)
        if r["microcopy_opportunity"] == "MICROCOPY_HIGH_VALUE" and "MICROCOPY_OPPORTUNITY" not in primary+secondary:
            flag("MICROCOPY_OPPORTUNITY_IGNORED", r)
        if (r["enrichment_opportunity"] == "REFERENCE_GRADE_OPPORTUNITY" and
                (r["recommended_ambition_level"] != "STRETCH" or not (set(primary) & {"COMPOSITION_UPGRADE", "BACKGROUND_SCENE", "ILLUSTRATION_ENRICHMENT"}))):
            flag("REFERENCE_SCENE_OPPORTUNITY_IGNORED", r)
        if visual and (r["density"] == "HIGH" or r["space_status"] != "AVAILABLE_CANDIDATE" or r["existing_illustration_present"]):
            flag("OVER_DECORATION_RISK", r, "HUMAN_REVIEW")
        if r["scientific_risk"] != "PROTECTED_SOURCE_CONTENT" or r["semantic_risk"] != "SOURCE_BOUND_REVIEW_REQUIRED":
            flag("SCIENTIFIC_VISUAL_CONFUSION_RISK", r, "HUMAN_REVIEW")
    if len(records) >= 3 and len({r["narrative_function"] for r in records}) > 1:
        signatures = {(r["recommended_ambition_level"], tuple(r["primary_enrichment_channels"]), r["composition_upgrade_candidate"]["type"]) for r in records}
        if len(signatures) == 1:
            flag("EVERY_PAGE_TREATED_THE_SAME", None)
    return findings


def plan(bundle, *, enabled=False):
    """Default off precedes all reads, validation, allocation of plans and writes."""
    if enabled is not True:
        return None
    lock, severities, safety, priority = authority()
    before = digest(bundle)
    floor, slides, contexts = bundle["floor_report"], bundle["floor_records"], bundle["contexts"]
    n = floor["summary"]["slide_count"]
    ids = [s["slide_id"] for s in slides]
    if n < 1 or len(slides) != n or len(set(ids)) != n or [s["slide_index"] for s in slides] != list(range(1, n+1)):
        raise ValueError("FROZEN_FLOOR_RECORD_COVERAGE_INVALID")
    page_roles = {r["page_role"] for r in read_json(FLOOR_CONFIG / "page_role_and_composition_taxonomy.json")["roles"]}
    for slide in slides:
        if slide["page_role"] not in page_roles:
            raise ValueError("INFORMATION_REQUIRED: unresolved primary page role")
    context_ids = [c["slide_id"] for c in contexts]
    if len(contexts) != n or len(set(context_ids)) != n or set(context_ids) != set(ids):
        raise ValueError("MISSING_OR_DUPLICATE_AMBITION_RECORD")
    for c in contexts:
        validate_context(c)
        if c.get("requested_visual_type") not in (None, *safety["allowed_subject_classes"], *safety["forbidden_evidence_types"]):
            raise ValueError("INVALID_REQUESTED_VISUAL_TYPE")
    for f in floor["findings"]:
        if (severities.get(f["check_id"]) != f["severity"] or f.get("slide_id") not in (None, *ids) or
                f.get("result") not in ("FAIL", "PENDING", "NOT_ASSESSED")):
            raise ValueError("VISUAL_AMBITION_REQUIRES_FLOOR_CONTRACT_CHANGE: inconsistent frozen findings")
    aggregate = {"BLOCK": "VISUAL_QUALITY_BLOCKED", "HUMAN_REVIEW": "VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED",
                 "WARN": "VISUAL_QUALITY_PASS_WITH_WARNINGS", "PASS": "VISUAL_QUALITY_PASS"}[floor_status(floor["findings"])]
    if floor["status"] != aggregate:
        raise ValueError("FROZEN_FLOOR_STATUS_INCONSISTENT")
    brief = bundle["visual_brief"]
    if brief.get("approval", {}).get("status") != "approved":
        raise ValueError("APPROVED_VISUAL_BRIEF_REQUIRED")
    direction_name = brief.get("direction", {}).get("direction_name", "").casefold()
    if not direction_name.strip():
        raise ValueError("APPROVED_DIRECTION_REQUIRED")
    palette = brief.get("direction", {}).get("primary_palette", []) + brief.get("direction", {}).get("secondary_palette", [])
    if not palette or any(not isinstance(c, str) or not re.fullmatch(r"#?[0-9a-fA-F]{6}", c) for c in palette):
        raise ValueError("INVALID_HANDOFF_PALETTE")
    if type(bundle.get("third_party_upload_allowed")) is not bool:
        raise ValueError("EXPLICIT_PROJECT_PRIVACY_POLICY_REQUIRED")
    by_id = {c["slide_id"]: c for c in contexts}
    records = []
    for slide in slides:
        sid = slide["slide_id"]
        c = by_id[sid]
        local = [f for f in floor["findings"] if f.get("slide_id") in (sid, None)]
        status = floor_status(local)
        features = slide["design_record"]["composition_family"]["selector_inputs"]
        density, longest = features["content_density"], features["max_text_block_length"]
        anchor_role, anchor_reason = NARRATIVE[c["narrative_function"]]
        anchor = anchor_role == "ANCHOR"
        unresolved = slide["semantic_mapping_status"] != "SOURCE_BOUND"
        source_required = c["source_figure_required"] or slide["illustration_decision"] == "SOURCE_FIGURE_REQUIRED"
        unsafe_visual = c.get("requested_visual_type") in safety["forbidden_evidence_types"]
        need = c["communication_need"]
        ill = c["illustration_potential"]
        opportunity = ("NONE" if need == "ALREADY_CLEAR" and ill == "NONE" else
                       "HIGH" if need in ("ROLE_DISTINCTION", "RELATIONSHIP_UNRESOLVED", "SOURCE_FIGURE_GAP") or
                           ill in ("HIGH", "REFERENCE_GRADE_OPPORTUNITY") or density == "HIGH" or longest > 60 else
                       "LOW" if need == "ALREADY_CLEAR" and ill == "LOW" else "MEDIUM")
        reference = anchor and c["reference_scene_opportunity"] and not unresolved and not source_required and not unsafe_visual
        if reference:
            opportunity = "REFERENCE_GRADE_OPPORTUNITY"
        level = "FLOOR" if opportunity == "NONE" else "STRETCH" if reference else "TARGET"
        micro = ("MICROCOPY_UNSAFE" if unresolved or c["microcopy_requires_protected_rewrite"] else
                 "MICROCOPY_HIGH_VALUE" if density == "HIGH" or longest > 60 or c["repeated_wording"] else
                 "MICROCOPY_MAY_HELP" if need == "LABEL_HIERARCHY" else "MICROCOPY_NOT_NEEDED")
        composition_type = ("EXISTING_COMPOSITION_FAMILY" if opportunity == "NONE" else
                            "FIGURE_LED_COMPOSITION" if source_required else "SCENE_LEVEL_COMPOSITION" if reference else
                            "EDITORIAL_COMPOSITION" if slide["layout_repetition_family"] == "GENERIC_CARD_GRID" else
                            "VARIANT_OF_EXISTING_FAMILY")
        channels, secondary = [], []
        if opportunity != "NONE":
            if unresolved:
                channels = ["COMPOSITION_UPGRADE", "ANNOTATION_ENRICHMENT"]
            elif source_required:
                channels = ["COMPOSITION_UPGRADE", "ANNOTATION_ENRICHMENT"]
            elif reference:
                channels = ["COMPOSITION_UPGRADE", "EDITORIAL_COMPONENTS" if ill == "NONE" else
                    "BACKGROUND_SCENE" if c["subject_class"] in (
                    "BACKGROUND_SCENE", "CARE_PATHWAY_ATMOSPHERE", "REHABILITATION_CONCEPT") else "ILLUSTRATION_ENRICHMENT"]
            elif micro == "MICROCOPY_HIGH_VALUE":
                channels = ["MICROCOPY_OPPORTUNITY", "TYPOGRAPHY_UPGRADE"]
            elif ill in ("HIGH", "REFERENCE_GRADE_OPPORTUNITY") and not unsafe_visual:
                channels = ["ILLUSTRATION_ENRICHMENT", "COMPOSITION_UPGRADE"]
            elif c["narrative_function"] == "EVIDENCE_METHOD":
                channels = ["ICONOGRAPHY", "EDITORIAL_COMPONENTS"]
            elif c["narrative_function"] == "KEY_CONCEPT":
                channels = ["TYPOGRAPHY_UPGRADE", "ANNOTATION_ENRICHMENT"]
            else:
                channels = ["COMPOSITION_UPGRADE", "EDITORIAL_COMPONENTS"]
            if micro == "MICROCOPY_MAY_HELP":
                secondary = ["MICROCOPY_OPPORTUNITY"]
        if "COMPOSITION_UPGRADE" not in channels:
            composition_type = "EXISTING_COMPOSITION_FAMILY"
        forbidden_illustration = unresolved or source_required or unsafe_visual
        reason = ("HOLD: original semantic restriction must be resolved by the author" if unresolved else
                  "Only an existing registered scientific/source figure can meet this need" if source_required else
                  "Prohibited evidence-like medical visual cannot be a conceptual/AI candidate" if unsafe_visual else
                  "No illustrative gain identified; preserve the native content focal point" if ill == "NONE" else
                  "Protect text and native scientific objects; space must be reviewed before an optional candidate" if c["space_status"] != "AVAILABLE_CANDIDATE" else None)
        eligible = ill != "NONE" and not forbidden_illustration
        next_action = ("HUMAN_SEMANTIC_DECISION" if unresolved else "SOURCE_FIGURE_REQUIRED" if source_required else
                       "ART_DIRECTION_REVIEW" if unsafe_visual else "NO_ACTION" if opportunity == "NONE" else
                       "REFERENCE_GRADE_SCENE_TRIAL" if reference else "MICROCOPY_PROPOSAL" if micro == "MICROCOPY_HIGH_VALUE" else
                       "ILLUSTRATION_CANDIDATE" if "ILLUSTRATION_ENRICHMENT" in channels else
                       "ICON_SET_CANDIDATE" if "ICONOGRAPHY" in channels else
                       "COMPOSITION_REVIEW" if "COMPOSITION_UPGRADE" in channels else "ART_DIRECTION_REVIEW")
        row = {"schema_version": VERSION, "slide_id": sid, "slide_index": slide["slide_index"], "page_role": slide["page_role"],
               "quality_floor_status": status, "quality_floor_hard_checks_status": "BLOCK" if status == "BLOCK" else "PASS",
               "deck_quality_floor_status": floor["status"],
               "quality_floor_findings": copy.deepcopy(local), "floor_record_digest": digest(slide),
               "visual_ambition_status": "NO_MEANINGFUL_ENRICHMENT_NEEDED" if opportunity == "NONE" else
                   "ENRICHMENT_REQUIRES_HUMAN_DECISION" if forbidden_illustration else "ENRICHMENT_RECOMMENDED",
               "visual_anchor_role": anchor_role, "is_anchor": anchor, "anchor_reason": anchor_reason,
               "narrative_function": c["narrative_function"], "recommended_ambition_level": level,
               "ambition_level_achieved": "NOT_ASSESSED", "enrichment_opportunity": opportunity,
               "primary_enrichment_channels": channels, "secondary_enrichment_channels": secondary,
               "composition_upgrade_candidate": {"type": composition_type, "existing_family": slide["composition_family"],
                   "new_family_selected": None, "semantic_graph_change_allowed": False, "renderer_change_allowed": False},
               "illustration_decision": slide["illustration_decision"], "illustration_enrichment_opportunity": ill,
               "illustration_role_candidate": ("BACKGROUND_ATMOSPHERE" if c["subject_class"] in (
                   "BACKGROUND_SCENE", "CARE_PATHWAY_ATMOSPHERE", "REHABILITATION_CONCEPT") else
                   "EDITORIAL_CUE" if c["subject_class"] in ("EDITORIAL_ILLUSTRATION", "RESEARCH_ENVIRONMENT") else "CONCEPT_SUPPORT") if eligible else None,
               "illustration_subject_summary": SUBJECTS.get(c["subject_class"]) if eligible else None,
               "illustration_expected_gain": "Optional support for narrative focus and reading order; validate on a future composite preview" if eligible else
                   "Preserve native/source content; no decorative replacement is authorized",
               "illustration_risk": "SOURCE_OR_SEMANTIC_GATE" if forbidden_illustration else
                   "OVER_DECORATION_OR_OCCLUSION_REVIEW" if density == "HIGH" or c["space_status"] != "AVAILABLE_CANDIDATE" else "HUMAN_APPROPRIATENESS_REVIEW",
               "illustration_candidate_allowed": eligible, "illustration_deferral_reason": reason,
               "allowed_visual_types": [c["subject_class"]] if eligible else [], "prohibited_visual_types": safety["forbidden_evidence_types"],
               "source_priority": priority, "source_figure_required": source_required,
               "microcopy_opportunity": micro, "microcopy_reason": "Semantic/protected wording cannot be safely changed" if micro == "MICROCOPY_UNSAFE" else
                   "Density/long-label screening only; no wording proposed" if micro != "MICROCOPY_NOT_NEEDED" else "No label-length improvement identified",
               "protected_content": list(PROTECTED), "microcopy_text": None,
               "reference_style_relevance": {"benchmark": "APPROVED_VISUAL_BRIEF", "traits": list(TRAITS),
                   "scene_candidate": reference, "reference_grade_achieved": False, "image_similarity_scoring": False},
               "scientific_risk": "SOURCE_FIGURE_PRIORITY" if source_required else "PROHIBITED_MEDICAL_VISUAL" if unsafe_visual else
                   "CONCEPTUAL_ANATOMY_REVIEW" if eligible and c["subject_class"] == "ORGAN_SILHOUETTE" else "PROTECTED_SOURCE_CONTENT",
               "semantic_risk": slide["semantic_mapping_status"] if unresolved else "SOURCE_BOUND_REVIEW_REQUIRED",
               "privacy_risk": "EXTERNAL_DISPATCH_PROHIBITED" if not bundle["third_party_upload_allowed"] else "SEPARATE_EXPORT_APPROVAL_REQUIRED",
               "human_approval_required": True, "human_review_required": True,
               "reasoning_summary": [anchor_reason, "Communication need: " + need,
                   NEED_REASON[need],
                   "Enrichment is an unexecuted opportunity, independent of existing Floor findings",
                   "Original illustration decision retained; optional enrichment is considered separately"],
               "recommended_next_action": next_action, "execution_eligibility": "HOLD_FLOOR_BLOCK" if status == "BLOCK" else
                   "HOLD_DECK_FLOOR_BLOCK" if floor["status"] == "VISUAL_QUALITY_BLOCKED" else "HUMAN_GATE_REQUIRED",
               "automatic_execution_allowed": False, "floor_override_allowed": False,
               "density": density, "space_status": c["space_status"], "existing_illustration_present": bool(slide["approved_asset_ids"]),
               "visual_brief_style_family": "APPROVED_VISUAL_BRIEF", "palette": palette,
               "negative_space_hint": c.get("negative_space_hint"), "audit_provenance": copy.deepcopy(c["evidence_refs"])}
        row.update(copy.deepcopy(c.get("override", {})))
        # Human-adjusted channels/actions still cannot bypass a source or semantic gate.
        if unresolved:
            row["recommended_next_action"] = "HUMAN_SEMANTIC_DECISION"
        elif source_required:
            row["recommended_next_action"] = "SOURCE_FIGURE_REQUIRED"
        elif unsafe_visual:
            row["recommended_next_action"] = "ART_DIRECTION_REVIEW"
        elif micro == "MICROCOPY_UNSAFE" and row["recommended_next_action"] == "MICROCOPY_PROPOSAL":
            row["recommended_next_action"] = "ART_DIRECTION_REVIEW"
        requested_action = c.get("override", {}).get("recommended_next_action")
        row["action_override_withheld_by_safety_gate"] = requested_action if requested_action and requested_action != row["recommended_next_action"] else None
        records.append(row)
    warnings = audit_ambition(records)
    for r in records:
        r["ambition_findings"] = [f for f in warnings if f["slide_id"] in (None, r["slide_id"])]
    high = lambda r: r["enrichment_opportunity"] in ("HIGH", "REFERENCE_GRADE_OPPORTUNITY")
    summary = {"slide_count": n, "anchor_slide_count": sum(r["is_anchor"] for r in records),
               "anchor_slide_ids": [r["slide_id"] for r in records if r["is_anchor"]],
               "supporting_slide_count": sum(r["visual_anchor_role"] == "SUPPORTING" for r in records),
               "utility_slide_count": sum(r["visual_anchor_role"] == "UTILITY" for r in records),
               "high_enrichment_opportunity_count": sum(high(r) for r in records),
               "illustration_candidate_count": sum(r["illustration_candidate_allowed"] and bool(set(r["primary_enrichment_channels"]) &
                   {"ILLUSTRATION_ENRICHMENT", "BACKGROUND_SCENE", "ICONOGRAPHY"}) for r in records),
               "optional_illustration_opportunity_count": sum(r["illustration_candidate_allowed"] for r in records),
               "microcopy_candidate_count": sum(r["microcopy_opportunity"] in ("MICROCOPY_MAY_HELP", "MICROCOPY_HIGH_VALUE") for r in records),
               "reference_grade_candidate_count": sum(r["enrichment_opportunity"] == "REFERENCE_GRADE_OPPORTUNITY" for r in records),
               "composition_upgrade_count": sum(r["composition_upgrade_candidate"]["type"] != "EXISTING_COMPOSITION_FAMILY" for r in records),
               "quality_floor_status": floor["status"], "quality_floor_slide_status_distribution": dict(Counter(r["quality_floor_status"] for r in records)),
               "full_floor_pass_with_high_opportunity": [r["slide_id"] for r in records if r["quality_floor_status"] == "PASS" and high(r)],
               "hard_checks_pass_with_high_opportunity": [r["slide_id"] for r in records if r["quality_floor_hard_checks_status"] == "PASS" and high(r)],
               "no_illustration_needed_with_high_enrichment": [r["slide_id"] for r in records if r["illustration_decision"] == "NO_ILLUSTRATION_NEEDED" and r["illustration_enrichment_opportunity"] in ("HIGH", "REFERENCE_GRADE_OPPORTUNITY")],
               "level_distribution": dict(Counter(r["recommended_ambition_level"] for r in records)),
               "visual_ambition_status": "HIGH_VALUE_ENRICHMENT_OPPORTUNITIES_PRESENT" if any(high(r) for r in records) else
                   "ENRICHMENT_OPPORTUNITIES_PRESENT" if any(r["enrichment_opportunity"] != "NONE" for r in records) else "NO_MEANINGFUL_ENRICHMENT_NEEDED",
               "human_review_required": True, "illustration_quota": None, "reference_grade_achieved": False,
               "canonical_promotion": "BLOCKED_PENDING_SECOND_PROJECT_AND_HUMAN_APPROVAL", "independent_real_projects_evaluated": bundle.get("independent_real_projects_evaluated", 0)}
    handoff_keys = ("slide_id", "page_role", "illustration_enrichment_opportunity", "illustration_role_candidate",
                    "illustration_subject_summary", "allowed_visual_types", "prohibited_visual_types", "source_priority",
                    "visual_brief_style_family", "palette", "negative_space_hint", "human_approval_required",
                    "quality_floor_status", "deck_quality_floor_status", "execution_eligibility", "illustration_candidate_allowed", "source_figure_required")
    handoff = [{**{k: r[k] for k in handoff_keys}, "local_only": True, "external_export_approved": False,
                "generation_authorized": False, "is_generation_request": False,
                "generated_candidate_scientific_evidence": False, "may_replace_source_figure": False} for r in records]
    if digest(bundle) != before:
        raise AssertionError("READ_ONLY_INPUT_MUTATION")
    return {"schema_version": VERSION, "experimental": False, "mode": "PLANNING", "default_off": True,
            "canonical": True, "production_active": True, "authority": lock,
            "quality_floor_status": floor["status"], "quality_floor_report_digest": digest(floor),
            "quality_floor_records_digest": digest(slides), "visual_ambition_status": summary["visual_ambition_status"],
            "summary": summary, "records": records, "handoff_records": handoff, "ambition_findings": warnings,
            "scientific_content_changed": False, "deck_modified": False, "real_deck_rendered": False,
            "images_generated": 0, "generation_requests_created": 0, "microcopy_generated_or_applied": 0}


def read_bundle(manifest_path):
    manifest = read_json(manifest_path)
    verified = {str(Path(manifest_path).resolve()): sha256(manifest_path)}

    def load(ref, yaml=False):
        path = Path(ref["path"])
        if sha256(path) != ref["sha256"]:
            raise ValueError("FROZEN_INPUT_CHANGED: " + str(path))
        verified[str(path.resolve())] = ref["sha256"]
        if yaml:
            import yaml as yaml_module
            return yaml_module.safe_load(path.read_text(encoding="utf-8-sig"))
        return read_json(path)

    bundle = {"floor_report": load(manifest["floor_report"]), "floor_records": load(manifest["floor_records"]),
              "contexts": load(manifest["contexts"]), "visual_brief": load(manifest["visual_brief"], yaml=True),
              "third_party_upload_allowed": manifest["third_party_upload_allowed"],
              "independent_real_projects_evaluated": manifest.get("independent_real_projects_evaluated", 0)}
    policy = load(manifest["privacy_policy_evidence"], yaml=True)
    if policy["privacy"]["third_party_upload_allowed"] is not bundle["third_party_upload_allowed"]:
        raise ValueError("PROJECT_PRIVACY_POLICY_MISMATCH")
    for context in bundle["contexts"]:
        for ref in context["evidence_refs"]:
            pointer_value(load(ref), ref.get("pointer", ""))
    return bundle, verified


def write_outputs(report, output):
    output = Path(output).resolve()
    if ROOT in output.parents and (ROOT / "staging").resolve() not in output.parents:
        raise ValueError("Use an external project output or a new staging directory")
    output.mkdir(parents=True, exist_ok=False)
    records = report["records"]
    outputs = {
        "visual_ambition_plan.json": {k: v for k, v in report.items() if k not in ("records", "handoff_records")},
        "slide_visual_opportunities.json": records,
        "anchor_slide_plan.json": {"anchors": [{k: r[k] for k in ("slide_id", "page_role", "is_anchor", "anchor_reason", "recommended_ambition_level", "human_review_required")} for r in records if r["is_anchor"]],
                                   "all_slide_roles": [{k: r[k] for k in ("slide_id", "visual_anchor_role", "narrative_function")} for r in records]},
        "illustration_opportunity_report.json": {"interface_version": VERSION, "handoff_records": report["handoff_records"],
            "records": [{k: r[k] for k in ("slide_id", "illustration_decision", "illustration_enrichment_opportunity", "illustration_expected_gain", "illustration_risk", "illustration_deferral_reason")} for r in records]},
        "microcopy_opportunity_report.json": {"text_generated": False, "records": [{k: r[k] for k in ("slide_id", "microcopy_opportunity", "microcopy_reason", "protected_content", "microcopy_text", "human_approval_required")} for r in records]},
        "composition_opportunity_report.json": {"renderer_changed": False, "records": [{k: r[k] for k in ("slide_id", "composition_upgrade_candidate", "recommended_ambition_level", "execution_eligibility")} for r in records]},
        "quality_floor_plus_ambition_summary.json": report["summary"],
    }
    for name, value in outputs.items():
        (output / name).write_text(json.dumps(value, ensure_ascii=False, indent=2)+"\n", encoding="utf8")
    lines = ["# Quality Floor + Visual Ambition", "", "Quality Floor: **"+report["quality_floor_status"]+"**",
             "", "Visual Ambition: **"+report["visual_ambition_status"]+"**", "",
             "Planning only. STRETCH is an attempt, never an achievement. No generation, microcopy or deck mutation.", "",
             "| Slide | Floor | Anchor role | Opportunity | Level | Primary channels | Primary next action |",
             "|---|---|---|---|---|---|---|"]
    lines += [f"| S{r['slide_index']:02d} | {r['quality_floor_status']} | {r['visual_anchor_role']} | {r['enrichment_opportunity']} | {r['recommended_ambition_level']} | {', '.join(r['primary_enrichment_channels']) or 'NONE'} | {r['recommended_next_action']} |" for r in records]
    lines += ["", "## Ambition review findings", ""] + [f"- {f['severity']}: {f['slide_id'] or 'DECK'} / {f['code']}" for f in report["ambition_findings"]]
    lines += ["", "Second independent real-project validation is still required; STRETCH is not an achieved score."]
    (output / "visual_ambition_plan.md").write_text("\n".join(lines)+"\n", encoding="utf8")
    with (output / "human_ambition_review_queue.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=["slide_id", "quality_floor_status", "ambition_findings", "recommended_next_action", "execution_eligibility", "human_approval_required"])
        writer.writeheader()
        for r in records:
            writer.writerow({**{k: r[k] for k in ("slide_id", "quality_floor_status", "recommended_next_action", "execution_eligibility", "human_approval_required")},
                             "ambition_findings": ";".join(f["code"] for f in r["ambition_findings"]) or "HUMAN_ART_DIRECTION_BENCHMARK_REVIEW"})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--enabled", action="store_true")
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if not args.enabled:
        return
    if not args.manifest or not args.output:
        parser.error("Explicit manifest and new output required")
    bundle, verified = read_bundle(args.manifest)
    report = plan(bundle, enabled=True)
    for path, expected in verified.items():
        if sha256(path) != expected:
            raise ValueError("INPUT_CHANGED_DURING_READ_ONLY_PLANNING: " + path)
    report["verified_input_sha256"] = verified
    write_outputs(report, args.output)
    print(report["quality_floor_status"] + " / " + report["visual_ambition_status"])


if __name__ == "__main__":
    main()
