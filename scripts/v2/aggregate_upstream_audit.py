#!/usr/bin/env python3
"""Aggregate isolated upstream PPT Skill audits into deterministic Phase 1 reports.

This script is deliberately read-only with respect to upstream checkouts. It consumes
only the normalized evidence recorded in ``audit/upstream_review/raw/*_audit.json``.
It does not execute upstream code, install dependencies, access the network, or infer
undocumented capabilities from repository popularity or prose.

Missing source fields are rendered as ``not documented``. The required outputs are
written only when all six requested upstream targets have one distinct audit record.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


NOT_DOCUMENTED = "not documented"

EXPECTED_PROJECTS: dict[str, str] = {
    "hugohe3_ppt-master": "hugohe3/ppt-master",
    "kdnsna_ultimate-ppt-master-skill": "kdnsna/ultimate-ppt-master-skill",
    "gabberflast_academic-pptx-skill": "Gabberflast/academic-pptx-skill",
    "noi1r_powerpoint-skill": "Noi1r/powerpoint-skill",
    "minimax_ai_skills_pptx_generator": "MiniMax-AI/skills:skills/pptx-generator",
    "anthropics_skills_pptx": "anthropics/skills:skills/pptx",
}


def parse_args() -> argparse.Namespace:
    script_path = Path(__file__).resolve()
    project_root = script_path.parents[2]
    parser = argparse.ArgumentParser(
        description="Aggregate six static upstream audit JSON records."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=project_root,
        help="Academic PPT workflow root (default: inferred from this script).",
    )
    parser.add_argument(
        "--raw-dir",
        type=Path,
        default=None,
        help="Raw audit directory (default: audit/upstream_review/raw).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Phase 1 output directory (default: audit/upstream_review).",
    )
    return parser.parse_args()


def get_path(data: Mapping[str, Any], path: str) -> Any:
    current: Any = data
    for key in path.split("."):
        if not isinstance(current, Mapping) or key not in current:
            return None
        current = current[key]
    return current


def first_present(data: Mapping[str, Any], paths: Sequence[str]) -> Any:
    for path in paths:
        value = get_path(data, path)
        if value is not None and value != "":
            return value
    return None


def normalize_scalar(value: Any) -> str:
    if value is None or value == "":
        return NOT_DOCUMENTED
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return str(value).replace("\r\n", "\n").replace("\r", "\n")


def normalize_joined(value: Any) -> str:
    if value is None or value == "":
        return NOT_DOCUMENTED
    if isinstance(value, list):
        if not value:
            return "none documented"
        items: list[str] = []
        for item in value:
            if isinstance(item, Mapping):
                items.append(
                    json.dumps(
                        item,
                        ensure_ascii=False,
                        sort_keys=True,
                        separators=(",", ":"),
                    )
                )
            else:
                items.append(str(item))
        return " | ".join(items)
    if isinstance(value, Mapping):
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return str(value).replace("\r\n", "\n").replace("\r", "\n")


def normalized_field(data: Mapping[str, Any], *paths: str, joined: bool = False) -> str:
    value = first_present(data, paths)
    return normalize_joined(value) if joined else normalize_scalar(value)


def recorded_deck_ir(data: Mapping[str, Any]) -> str:
    """Return only an explicit Deck IR field or an IR item named as DeckIR."""
    explicit = get_path(data, "architecture_review.deck_ir")
    if explicit is not None and explicit != "":
        return normalize_scalar(explicit)
    intermediate = first_present(
        data,
        (
            "architecture.intermediate_representation",
            "architecture_review.intermediate_representation",
        ),
    )
    values = intermediate if isinstance(intermediate, list) else [intermediate]
    matches = [
        str(value)
        for value in values
        if value is not None and "deckir" in str(value).lower().replace(" ", "")
    ]
    return " | ".join(matches) if matches else NOT_DOCUMENTED


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def identify_slug(data: Mapping[str, Any], source_path: Path) -> str:
    recorded = first_present(data, ("slug", "audit_slug"))
    filename_slug = source_path.name.removesuffix("_audit.json")
    if recorded is not None and str(recorded) != filename_slug:
        raise ValueError(
            f"{source_path.name}: recorded slug {recorded!r} does not match filename "
            f"slug {filename_slug!r}"
        )
    return str(recorded) if recorded is not None else filename_slug


def extract_project_label(data: Mapping[str, Any], slug: str) -> str:
    recorded = first_present(data, ("project", "scope.repository"))
    if isinstance(recorded, Mapping):
        owner = recorded.get("owner")
        name = recorded.get("name")
        if owner and name:
            return f"{owner}/{name}"
        return EXPECTED_PROJECTS[slug]
    target = first_present(data, ("scope.target",))
    if recorded is None:
        return EXPECTED_PROJECTS[slug]
    label = str(recorded)
    if target and slug in {
        "minimax_ai_skills_pptx_generator",
        "anthropics_skills_pptx",
    }:
        label = f"{label}:{target}"
    return label


def extract_hash_records(data: Mapping[str, Any]) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    mapped = data.get("sha256")
    if isinstance(mapped, Mapping):
        for path, digest in sorted(mapped.items(), key=lambda item: str(item[0]).lower()):
            records.append(
                {
                    "path": str(path),
                    "sha256": normalize_scalar(digest),
                    "role": NOT_DOCUMENTED,
                }
            )
    listed = data.get("file_hashes_sha256")
    if isinstance(listed, list):
        for item in listed:
            if not isinstance(item, Mapping):
                continue
            records.append(
                {
                    "path": normalize_scalar(item.get("path")),
                    "sha256": normalize_scalar(item.get("sha256")),
                    "role": normalize_scalar(item.get("role")),
                }
            )
    key_listed = data.get("key_file_hashes")
    if isinstance(key_listed, list):
        for item in key_listed:
            if not isinstance(item, Mapping):
                continue
            records.append(
                {
                    "path": normalize_scalar(item.get("path")),
                    "sha256": normalize_scalar(item.get("sha256")),
                    "role": normalize_scalar(item.get("role")),
                }
            )
    unique: dict[tuple[str, str], dict[str, str]] = {}
    for record in records:
        unique[(record["path"], record["sha256"])] = record
    return sorted(
        unique.values(),
        key=lambda item: (item["path"].lower(), item["sha256"]),
    )


def critical_hash_summary(data: Mapping[str, Any]) -> tuple[str, str]:
    records = extract_hash_records(data)
    if not records:
        return "0", NOT_DOCUMENTED
    summary = " | ".join(
        f"{record['path']}={record['sha256']}"
        + (
            f" [{record['role']}]"
            if record["role"] != NOT_DOCUMENTED
            else ""
        )
        for record in records
    )
    return str(len(records)), summary


def dependency_records(data: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    direct = data.get("dependencies")
    if isinstance(direct, list):
        return [item for item in direct if isinstance(item, Mapping)]
    nested = get_path(data, "dependency_review.declared_dependencies")
    if isinstance(nested, list):
        return [item for item in nested if isinstance(item, Mapping)]
    python_dependencies = get_path(data, "technical_security.python_dependencies")
    node_dependencies = get_path(data, "technical_security.node_dependencies")
    records: list[Mapping[str, Any]] = []
    if isinstance(python_dependencies, list):
        records.extend(
            {
                "name": str(item),
                "ecosystem": "Python requirement declaration",
                "version": NOT_DOCUMENTED,
                "required_for": NOT_DOCUMENTED,
                "installation_scope_in_guidance": normalized_field(
                    data, "technical_security.installation_behavior"
                ),
                "locked": normalized_field(
                    data, "technical_security.dependency_locking"
                ),
            }
            for item in python_dependencies
        )
    if node_dependencies is not None:
        records.append(
            {
                "name": normalize_scalar(node_dependencies),
                "ecosystem": "Node dependency declaration",
                "version": NOT_DOCUMENTED,
                "required_for": NOT_DOCUMENTED,
                "installation_scope_in_guidance": normalized_field(
                    data, "technical_security.installation_behavior"
                ),
                "locked": normalized_field(
                    data, "technical_security.dependency_locking"
                ),
            }
        )
    if records:
        return records
    return []


def repository_status(data: Mapping[str, Any]) -> str:
    return normalized_field(data, "repository.git_status", "repository_state.git_status")


def audit_execution_flag(data: Mapping[str, Any]) -> str:
    return normalized_field(
        data,
        "execution_constraints_observed.upstream_script_executed",
        "isolation_controls.upstream_generator_or_bridge_executed",
        "retrieval.scripts_executed",
    )


def install_execution_flag(data: Mapping[str, Any]) -> str:
    return normalized_field(
        data,
        "execution_constraints_observed.upstream_setup_or_install_executed",
        "isolation_controls.upstream_setup_or_install_executed",
        "retrieval.install_scripts_executed",
    )


def registry_row(
    data: Mapping[str, Any], slug: str, source_path: Path, raw_dir: Path
) -> dict[str, str]:
    hash_count, hashes = critical_hash_summary(data)
    dependencies = dependency_records(data)
    return {
        "audit_slug": slug,
        "source_project": extract_project_label(data, slug),
        "target_scope": normalized_field(data, "scope.target"),
        "repository_url": normalized_field(
            data,
            "repository.requested_url",
            "repository.url",
            "repository.remote_url",
        ),
        "source_commit": normalized_field(
            data, "repository.head_commit", "repository.pinned_commit"
        ),
        "commit_date": normalized_field(
            data,
            "repository.head_commit_date",
            "repository.commit_time",
            "repository_state.commit_time",
        ),
        "retrieval_date": normalized_field(
            data,
            "repository.retrieval_completed_at",
            "repository.retrieved_at",
            "repository.retrieval_started_at",
            "retrieved_at",
            "retrieval.recorded_at",
            "retrieval.retrieval_date",
            "audit_generated_at",
            "audited_at",
        ),
        "repository_status": repository_status(data),
        "commit_signature_status": normalized_field(
            data, "repository.commit_signature_status"
        ),
        "retrieval_method": normalized_field(
            data,
            "repository.retrieval_method",
            "repository.checkout.mode",
            "retrieval.clone_mode",
        ),
        "checkout_path": normalized_field(
            data,
            "repository.clone_path",
            "repository.checkout.directory",
            "project.local_source_path",
        ),
        "audit_type": normalized_field(
            data, "audit_type", "scope.review_type", "audit_scope"
        ),
        "dynamic_validation_status": normalized_field(
            data,
            "audit_conclusion.dynamic_validation_status",
            "assessment.dynamic_validation_status",
        ),
        "upstream_code_executed_during_audit": audit_execution_flag(data),
        "upstream_install_executed_during_audit": install_execution_flag(data),
        "dependencies_installed_during_audit": normalized_field(
            data,
            "execution_constraints_observed.dependency_installed",
            "isolation_controls.dependencies_installed",
            "retrieval.dependencies_installed",
        ),
        "network_access_record": normalized_field(
            data,
            "repository.network_access_performed_by_audit",
            "repository.network_access",
            "retrieval.network_access",
            joined=True,
        ),
        "documented_upstream_external_network_behavior": normalized_field(
            data,
            "technical_security_review.external_network_behavior",
            "technical_security.external_network",
            joined=True,
        ),
        "external_api_required": normalized_field(
            data,
            "technical_security_review.external_api_required_by_target",
            "scientific_review.external_api_required",
        ),
        "api_key_handling": normalized_field(
            data,
            "technical_security_review.api_key_handling",
            "technical_security_review.api_key_handling_in_target",
            "technical_security.api_key_handling",
        ),
        "local_service_behavior": normalized_field(
            data,
            "technical_security_review.local_service_behavior",
            "technical_security_review.local_service",
            "technical_security.local_services",
        ),
        "user_directory_behavior": normalized_field(
            data,
            "technical_security_review.user_directory_modification",
            "technical_security_review.user_directory_risk",
            "technical_security.user_directory_modification",
        ),
        "automatic_git_behavior": normalized_field(
            data,
            "technical_security_review.automatic_git_behavior",
            "technical_security_review.git_operations_in_target",
            "technical_security.git_behavior",
        ),
        "input_or_output_overwrite_behavior": normalized_field(
            data,
            "technical_security_review.output_overwrite_behavior",
            "technical_security_review.input_overwrite_risks",
            "technical_security_review.input_overwrite_risk",
            "technical_security.input_upload_or_exfiltration_risk",
            joined=True,
        ),
        "declared_dependency_count": str(len(dependencies)),
        "critical_hash_count": hash_count,
        "critical_file_hashes": hashes,
        "raw_audit_file": source_path.relative_to(raw_dir.parent.parent.parent).as_posix(),
        "raw_audit_sha256": sha256_file(source_path),
    }


def license_row(data: Mapping[str, Any], slug: str) -> dict[str, str]:
    return {
        "audit_slug": slug,
        "source_project": extract_project_label(data, slug),
        "source_commit": normalized_field(
            data, "repository.head_commit", "repository.pinned_commit"
        ),
        "repository_license": normalized_field(
            data,
            "license_review.root_license_detected",
            "license.repository_license",
            "license.repository_wide_license",
            "license.declared_license",
        ),
        "target_license": normalized_field(
            data,
            "license.target_license",
            "license.target_frontmatter_license",
            "license_review.skill_front_matter_license_statement",
        ),
        "license_conflict": normalized_field(
            data, "license_review.conflict"
        ),
        "compatibility_assessment": normalized_field(
            data,
            "license_review.compatibility_assessment",
            "license.canonical_adoption_implication",
            "license.direct_copy_conclusion",
            "license.compatibility_assessment",
        ),
        "direct_copy_legally_indicated": normalized_field(
            data, "license.direct_copy_legally_indicated"
        ),
        "current_direct_copy_status": normalized_field(
            data,
            "direct_copy_eligibility.status",
            "assessment.copy_or_derivative_code",
        ),
        "attribution_or_copy_conditions": normalized_field(
            data,
            "license_review.attribution_requirement_if_reuse_is_later_authorized",
            "license.direct_copy_conditions",
            "license.express_restrictions",
            "direct_copy_eligibility.eligible_scope",
            joined=True,
        ),
        "allowed_current_use": normalized_field(
            data,
            "license.allowed_use_for_this_project",
            "direct_copy_eligibility.acceptable_current_use",
            "direct_copy_eligibility.eligible_scope",
            joined=True,
        ),
        "prohibited_current_use": normalized_field(
            data,
            "direct_copy_eligibility.prohibited_current_use",
            "direct_copy_eligibility.not_eligible_without_separate_approval",
            "assessment.do_not_adopt_as_is",
            joined=True,
        ),
        "third_party_asset_provenance": normalized_field(
            data,
            "license_review.third_party_asset_provenance",
            "license.license_boundary",
            "license.qualification",
        ),
        "license_file_hash": normalized_field(
            data,
            "license.target_license_sha256",
            "license.root_license_sha256",
            "license.license_sha256",
        ),
    }


def capability_row(data: Mapping[str, Any], slug: str) -> dict[str, str]:
    return {
        "audit_slug": slug,
        "source_project": extract_project_label(data, slug),
        "source_commit": normalized_field(
            data, "repository.head_commit", "repository.pinned_commit"
        ),
        "routing_or_actions": normalized_field(
            data,
            "entry_and_trigger_review.routing",
            "architecture.routes",
            "architecture.routing",
            "skill_entry.top_level_routes",
            joined=True,
        ),
        "four_route_exclusivity": normalized_field(
            data,
            "entry_and_trigger_review.top_level_route_exclusivity",
            "skill_entry.routing_rule",
        ),
        "input_contract": normalized_field(
            data, "architecture_review.input_contract", "architecture.input_contract"
        ),
        "output_contract": normalized_field(
            data, "architecture_review.output_contract", "architecture.output_contract"
        ),
        "intermediate_representation": normalized_field(
            data,
            "architecture_review.intermediate_representation",
            "architecture.intermediate_representation",
            joined=True,
        ),
        "deck_ir": recorded_deck_ir(data),
        "generation_backend": normalized_field(
            data,
            "architecture_review.generation_backend",
            "architecture.generation_backend",
            joined=True,
        ),
        "reference_template_handling": normalized_field(
            data,
            "architecture_review.reference_template_handling",
            "architecture.template_handling",
        ),
        "reference_style_extraction": normalized_field(
            data, "visual_review.reference_ppt_style_extraction"
            , "visual_capabilities.reference_style_extraction"
        ),
        "native_template_fill": normalized_field(
            data,
            "architecture_review.native_template_fill",
            "architecture.reference_template_handling",
        ),
        "masters_and_layouts": normalized_field(
            data, "visual_review.masters_and_layouts"
            , "visual_capabilities.masters_layouts_placeholders"
        ),
        "stable_slide_id": normalized_field(
            data,
            "architecture_review.stable_slide_id",
            "architecture.incremental_capability",
            "architecture.incremental_update",
        ),
        "incremental_update": normalized_field(
            data,
            "architecture_review.incremental_update",
            "architecture.incremental_update",
            "architecture.incremental_capability",
        ),
        "single_slide_update": normalized_field(
            data, "architecture_review.single_slide_update"
        ),
        "rendering": normalized_field(
            data, "architecture_review.rendering", "architecture.rendering"
        ),
        "failure_recovery": normalized_field(
            data,
            "architecture_review.failure_recovery",
            "architecture.failure_recovery",
            joined=True,
        ),
        "claim_source_mapping": normalized_field(
            data, "scientific_review.claim_source_mapping"
            , "scientific_integrity.claim_source_mapping"
        ),
        "no_fabrication_rule": normalized_field(
            data,
            "scientific_review.explicit_no_fabrication_rule",
            "scientific_review.fabrication_prohibition",
            "scientific_integrity.fact_inference_boundary",
        ),
        "fact_inference_separation": normalized_field(
            data,
            "scientific_review.fact_summary_inference_confirmation_separation",
            "scientific_review.fact_inference_suggestion_separation",
            "scientific_integrity.facts_vs_inference_vs_suggestion",
        ),
        "causal_language_guard": normalized_field(
            data, "scientific_review.causal_language_guard"
            , "scientific_integrity.causal_and_mechanism_safeguards"
        ),
        "computational_or_mechanism_guard": normalized_field(
            data,
            "scientific_review.computational_inference_guard",
            "scientific_review.mechanism_evidence_guard",
            "scientific_integrity.causal_and_mechanism_safeguards",
        ),
        "input_hash_protection": normalized_field(
            data, "scientific_review.input_hash_protection"
        ),
        "citation_support": normalized_field(
            data,
            "scientific_review.citation_and_footnote_support",
            "scientific_review.citation_support",
            "scientific_integrity.citation_and_footnote_support",
        ),
        "limitations_slide_support": normalized_field(
            data,
            "scientific_review.limitations_slide_support",
            "scientific_review.limitations_slide_required",
            "scientific_integrity.limitations_page_support",
        ),
        "action_title": normalized_field(
            data, "visual_review.action_title", "visual_capabilities.action_titles"
        ),
        "density_control": normalized_field(
            data, "visual_review.density_control", "visual_capabilities.density_control"
        ),
        "formula_pipeline": normalized_field(
            data, "visual_review.formula_pipeline", "visual_capabilities.formula_support"
        ),
        "scientific_diagram_pipeline": normalized_field(
            data,
            "visual_review.scientific_diagram_pipeline",
            "visual_capabilities.diagram_support",
        ),
        "svg_or_drawingml_backend": normalized_field(
            data,
            "visual_review.svg_or_drawingml_backend",
            "visual_capabilities.svg_and_drawingml",
        ),
        "ooxml_qa": normalized_field(
            data,
            "ooxml_qa_review",
            "visual_capabilities.ooxml_qa",
            "architecture.qa",
            "architecture_review.qa",
            joined=True,
        ),
        "editable_object_depth": normalized_field(
            data,
            "visual_review.editable_depth",
            "visual_review.editable_object_depth",
            "visual_capabilities.editable_depth",
        ),
        "visual_strengths": normalized_field(
            data, "visual_review.strengths", joined=True
        ),
        "visual_weaknesses": normalized_field(
            data,
            "visual_review.weaknesses",
            "capability_assessment.missing_or_insufficient_for_v2",
            joined=True,
        ),
        "scientific_risks": normalized_field(
            data,
            "scientific_review.scientific_risks",
            "scientific_review.false_reference_risk",
            "scientific_integrity.fact_inference_risk",
            "scientific_integrity.false_reference_risk",
            joined=True,
        ),
        "security_risk": normalized_field(
            data,
            "technical_security_review.overall_risk",
            "technical_security_review.supply_chain_risk",
            "risk_and_rejection_reasons.overall_risk",
            "technical_security.supply_chain_assessment",
        ),
        "candidate_or_high_value_capabilities": normalized_field(
            data,
            "audit_conclusion.highest_value_capabilities",
            "assessment.candidate_capabilities",
            "assessment.candidate_behavioral_requirements",
            "capability_assessment.highest_value_candidates",
            joined=True,
        ),
        "whole_project_recommendation": normalized_field(
            data,
            "audit_conclusion.recommended_adoption_mode_for_later_governance",
            "assessment.whole_skill_recommendation",
            "direct_copy_eligibility.preferred_adoption_mode",
        ),
        "dynamic_validation_status": normalized_field(
            data,
            "audit_conclusion.dynamic_validation_status",
            "assessment.dynamic_validation_status",
        ),
    }


def dependency_rows(data: Mapping[str, Any], slug: str) -> list[dict[str, str]]:
    records = dependency_records(data)
    shared = {
        "audit_slug": slug,
        "source_project": extract_project_label(data, slug),
        "source_commit": normalized_field(
            data, "repository.head_commit", "repository.pinned_commit"
        ),
        "package_manifest_present": normalized_field(
            data,
            "dependency_review.package_manifest_present",
            "target_inventory.package_manifest_present",
        ),
        "lockfile_present": normalized_field(
            data,
            "dependency_review.lockfile_present",
            "target_inventory.dependency_lock_present",
        ),
        "versions_pinned_overall": normalized_field(
            data,
            "dependency_review.versions_pinned",
            "technical_security_review.dependency_versions_locked",
            "technical_security.dependency_locking",
        ),
        "dependency_or_supply_chain_risk": normalized_field(
            data,
            "dependency_review.dependency_risk",
            "technical_security_review.supply_chain_risk",
            "technical_security_review.supply_chain_risks",
            "technical_security.supply_chain_assessment",
            joined=True,
        ),
    }
    if not records:
        return [
            {
                **shared,
                "dependency_name": NOT_DOCUMENTED,
                "ecosystem_or_type": NOT_DOCUMENTED,
                "version": NOT_DOCUMENTED,
                "required": NOT_DOCUMENTED,
                "required_for": NOT_DOCUMENTED,
                "installation_scope_or_behavior": NOT_DOCUMENTED,
                "locked": NOT_DOCUMENTED,
            }
        ]
    rows: list[dict[str, str]] = []
    for dependency in records:
        rows.append(
            {
                **shared,
                "dependency_name": normalize_scalar(dependency.get("name")),
                "ecosystem_or_type": normalize_scalar(
                    dependency.get("ecosystem", dependency.get("type"))
                ),
                "version": normalize_scalar(dependency.get("version")),
                "required": normalize_scalar(dependency.get("required")),
                "required_for": normalize_scalar(
                    dependency.get(
                        "required_for",
                        dependency.get("purpose"),
                    )
                ),
                "installation_scope_or_behavior": normalize_scalar(
                    dependency.get(
                        "installation_scope_in_guidance",
                        dependency.get(
                            "install_scope_in_docs",
                            dependency.get("install_behavior"),
                        ),
                    )
                ),
                "locked": normalize_scalar(dependency.get("locked")),
            }
        )
    return rows


def markdown_safe(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ")


def bullet_values(value: Any) -> list[str]:
    if value is None or value == "":
        return [NOT_DOCUMENTED]
    if isinstance(value, list):
        if not value:
            return ["none documented"]
        return [normalize_scalar(item) for item in value]
    return [normalize_scalar(value)]


def security_review_markdown(
    audits: Sequence[tuple[str, Mapping[str, Any], Path]]
) -> str:
    lines = [
        "# Upstream Security Review",
        "",
        "This report is generated only from the six isolated static audit records. "
        "No upstream setup, installer, generator, bridge, API, or local service was "
        "executed by this aggregation step.",
        "",
        "| Project | Commit | Audit execution | Install execution | Overall risk |",
        "|---|---|---:|---:|---|",
    ]
    for slug, data, _ in audits:
        lines.append(
            "| "
            + " | ".join(
                markdown_safe(value)
                for value in (
                    extract_project_label(data, slug),
                    normalized_field(
                        data, "repository.head_commit", "repository.pinned_commit"
                    ),
                    audit_execution_flag(data),
                    install_execution_flag(data),
                    normalized_field(
                        data,
                        "technical_security_review.overall_risk",
                        "technical_security_review.supply_chain_risk",
                        "risk_and_rejection_reasons.overall_risk",
                        "technical_security.supply_chain_assessment",
                    ),
                )
            )
            + " |"
        )
    lines.extend(["", "## Project findings", ""])
    for slug, data, _ in audits:
        lines.extend(
            [
                f"### {extract_project_label(data, slug)}",
                "",
                f"- Pinned commit: `{normalized_field(data, 'repository.head_commit', 'repository.pinned_commit')}`",
                f"- Repository state: {repository_status(data)}",
                f"- Network access recorded for retrieval: {normalized_field(data, 'repository.network_access_performed_by_audit', 'repository.network_access', 'retrieval.network_access', joined=True)}",
                f"- Documented upstream external-network behavior: {normalized_field(data, 'technical_security_review.external_network_behavior', 'technical_security.external_network', joined=True)}",
                f"- Documented installation behavior: {normalized_field(data, 'technical_security_review.installation_guidance', 'technical_security_review.documented_auto_install_behavior', 'technical_security.installation_behavior', joined=True)}",
                f"- API/key handling: {normalized_field(data, 'technical_security_review.api_key_handling', 'technical_security_review.api_key_handling_in_target', 'technical_security.api_key_handling')}",
                f"- External API requirement: {normalized_field(data, 'technical_security_review.external_api_required_by_target', 'scientific_review.external_api_required')}",
                f"- Local service behavior: {normalized_field(data, 'technical_security_review.local_service_behavior', 'technical_security_review.local_service', 'technical_security.local_services')}",
                f"- User-directory behavior: {normalized_field(data, 'technical_security_review.user_directory_modification', 'technical_security_review.user_directory_risk', 'technical_security.user_directory_modification')}",
                f"- Temporary-directory behavior: {normalized_field(data, 'technical_security_review.temporary_directory_behavior', 'technical_security.temporary_directories')}",
                f"- Automatic Git behavior: {normalized_field(data, 'technical_security_review.automatic_git_behavior', 'technical_security_review.git_operations_in_target', 'technical_security.git_behavior')}",
                f"- Overwrite or exfiltration risk: {normalized_field(data, 'technical_security_review.output_overwrite_behavior', 'technical_security_review.input_overwrite_risks', 'technical_security_review.input_overwrite_risk', 'technical_security.input_upload_or_exfiltration_risk', joined=True)}",
                "- Supply-chain risks:",
            ]
        )
        for item in bullet_values(
            first_present(
                data,
                (
                    "technical_security_review.supply_chain_risks",
                    "technical_security_review.supply_chain_risk",
                    "dependency_review.dependency_risk",
                    "technical_security.supply_chain_assessment",
                ),
            )
        ):
            lines.append(f"  - {item}")
        lines.extend(["- Recorded positive controls:"])
        for item in bullet_values(
            first_present(data, ("technical_security_review.positive_controls",))
        ):
            lines.append(f"  - {item}")
        lines.append("")
    lines.extend(
        [
            "## Aggregation boundary",
            "",
            "- `not documented` means the raw audit did not provide a directly mapped field.",
            "- Capability claims remain static-review observations until isolated dynamic benchmarks are run.",
            "- License classification is an engineering control, not legal advice.",
            "- No repository was promoted, installed, merged, or made canonical by this report.",
            "",
        ]
    )
    return "\n".join(lines)


def feature_notes_markdown(
    audits: Sequence[tuple[str, Mapping[str, Any], Path]]
) -> str:
    lines = [
        "# Upstream Feature Notes",
        "",
        "These notes record candidate behaviors and limitations from static review. "
        "They are not evidence that a capability passed runtime, scientific, visual, "
        "or promotion-gate validation.",
        "",
    ]
    for slug, data, _ in audits:
        lines.extend(
            [
                f"## {extract_project_label(data, slug)}",
                "",
                f"- Commit: `{normalized_field(data, 'repository.head_commit', 'repository.pinned_commit')}`",
                f"- Architecture/routes: {normalized_field(data, 'entry_and_trigger_review.routing', 'architecture.routes', 'architecture.routing', 'skill_entry.top_level_routes', joined=True)}",
                f"- Generation backend: {normalized_field(data, 'architecture_review.generation_backend', 'architecture.generation_backend', joined=True)}",
                f"- Reference/template behavior: {normalized_field(data, 'architecture_review.reference_template_handling', 'architecture.template_handling')}",
                f"- Incremental-update behavior: {normalized_field(data, 'architecture_review.incremental_update', 'architecture.incremental_update')}",
                "- Highest-value or candidate capabilities:",
            ]
        )
        for item in bullet_values(
            first_present(
                data,
                (
                    "audit_conclusion.highest_value_capabilities",
                    "assessment.candidate_capabilities",
                    "assessment.candidate_behavioral_requirements",
                    "capability_assessment.highest_value_candidates",
                ),
            )
        ):
            lines.append(f"  - {item}")
        lines.append("- Visual strengths:")
        for item in bullet_values(
            first_present(data, ("visual_review.strengths",))
        ):
            lines.append(f"  - {item}")
        lines.append("- Visual weaknesses:")
        for item in bullet_values(
            first_present(data, ("visual_review.weaknesses",))
        ):
            lines.append(f"  - {item}")
        lines.extend(
            [
                f"- Scientific traceability: claim-source mapping = "
                f"{normalized_field(data, 'scientific_review.claim_source_mapping')}; "
                f"input-hash protection = "
                f"{normalized_field(data, 'scientific_review.input_hash_protection')}.",
                f"- Static-audit recommendation: {normalized_field(data, 'audit_conclusion.recommended_adoption_mode_for_later_governance', 'assessment.whole_skill_recommendation', 'direct_copy_eligibility.preferred_adoption_mode')}",
                "",
            ]
        )
    return "\n".join(lines)


def rejection_reasons_markdown(
    audits: Sequence[tuple[str, Mapping[str, Any], Path]]
) -> str:
    lines = [
        "# Upstream Rejection and Deferral Reasons",
        "",
        "A repository-level rejection or deferral does not reject every high-level "
        "behavior. It prevents direct canonical adoption until the listed scientific, "
        "license, security, dependency, and validation gates are satisfied.",
        "",
    ]
    for slug, data, _ in audits:
        lines.extend(
            [
                f"## {extract_project_label(data, slug)}",
                "",
                f"- Commit: `{normalized_field(data, 'repository.head_commit', 'repository.pinned_commit')}`",
                f"- Current recommendation: {normalized_field(data, 'audit_conclusion.recommended_adoption_mode_for_later_governance', 'assessment.whole_skill_recommendation', 'direct_copy_eligibility.preferred_adoption_mode')}",
                "- Rejection/deferral reasons:",
            ]
        )
        reasons = first_present(
            data,
            (
                "rejection_or_deferral_reasons",
                "assessment.do_not_adopt_as_is",
                "risk_and_rejection_reasons.reasons",
            ),
        )
        for item in bullet_values(reasons):
            lines.append(f"  - {item}")
        lines.append("- Conditions or required gates before reconsideration:")
        conditions = first_present(
            data,
            (
                "direct_copy_eligibility.conditions_to_reconsider",
                "assessment.required_promotion_tests",
                "assessment.implementation_constraint",
                "direct_copy_eligibility.eligible_scope",
            ),
        )
        for item in bullet_values(conditions):
            lines.append(f"  - {item}")
        lines.append("")
    lines.extend(
        [
            "## Capabilities excluded from direct adoption by this review",
            "",
            "- Automatic or global dependency installation.",
            "- User-home or global Skill installation.",
            "- Unapproved network research, third-party API relay, or remote asset retrieval.",
            "- In-place rewrite of formal PPTX without backup and atomic replacement.",
            "- Any scientific-content generation that bypasses the canonical evidence registry.",
            "- Copying proprietary or license-conflicted code, prompts, assets, or schemas.",
            "",
        ]
    )
    return "\n".join(lines)


def atomic_write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        newline="",
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=path.parent,
        delete=False,
    ) as handle:
        handle.write(content)
        temporary_path = Path(handle.name)
    os.replace(temporary_path, path)


def atomic_write_csv(path: Path, rows: Sequence[Mapping[str, str]]) -> None:
    if not rows:
        raise ValueError(f"Cannot write empty CSV: {path}")
    fieldnames = list(rows[0].keys())
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8-sig",
        newline="",
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=path.parent,
        delete=False,
    ) as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="raise")
        writer.writeheader()
        writer.writerows(rows)
        temporary_path = Path(handle.name)
    os.replace(temporary_path, path)


def validate_csv(
    path: Path,
    expected_project_count: int,
    exactly_one_row_per_project: bool,
) -> None:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)
    if not reader.fieldnames or "audit_slug" not in reader.fieldnames:
        raise ValueError(f"{path}: missing header or audit_slug column")
    slugs = {row["audit_slug"] for row in rows}
    if len(slugs) != expected_project_count:
        raise ValueError(
            f"{path}: expected {expected_project_count} projects, found {len(slugs)}"
        )
    if exactly_one_row_per_project and len(rows) != expected_project_count:
        raise ValueError(
            f"{path}: expected one row per project, found {len(rows)} rows"
        )


def load_audits(raw_dir: Path) -> list[tuple[str, Mapping[str, Any], Path]]:
    source_files = sorted(raw_dir.glob("*_audit.json"), key=lambda path: path.name.lower())
    audits_by_slug: dict[str, tuple[Mapping[str, Any], Path]] = {}
    for source_path in source_files:
        with source_path.open("r", encoding="utf-8-sig") as handle:
            data = json.load(handle)
        if not isinstance(data, Mapping):
            raise ValueError(f"{source_path}: top-level JSON must be an object")
        slug = identify_slug(data, source_path)
        if slug not in EXPECTED_PROJECTS:
            raise ValueError(f"{source_path.name}: unexpected project slug {slug!r}")
        if slug in audits_by_slug:
            raise ValueError(f"Duplicate audit record for {slug}")
        commit = first_present(data, ("repository.head_commit", "repository.pinned_commit"))
        retrieval = first_present(
            data,
            (
                "repository.retrieval_completed_at",
                "repository.retrieved_at",
                "repository.retrieval_started_at",
                "retrieved_at",
                "retrieval.recorded_at",
                "retrieval.retrieval_date",
            ),
        )
        if not commit:
            raise ValueError(f"{source_path.name}: fixed commit SHA is missing")
        if not retrieval:
            raise ValueError(f"{source_path.name}: retrieval date is missing")
        audits_by_slug[slug] = (data, source_path)

    missing = [slug for slug in EXPECTED_PROJECTS if slug not in audits_by_slug]
    extra_count = len(source_files) - len(audits_by_slug)
    if missing or extra_count:
        details = f"missing={','.join(missing) if missing else 'none'}"
        if extra_count:
            details += f"; duplicate_or_unaccepted_files={extra_count}"
        raise RuntimeError(
            "Phase 1 aggregation requires all six distinct requested projects; " + details
        )

    return [
        (slug, audits_by_slug[slug][0], audits_by_slug[slug][1])
        for slug in EXPECTED_PROJECTS
    ]


def main() -> int:
    args = parse_args()
    project_root = args.project_root.resolve()
    raw_dir = (
        args.raw_dir.resolve()
        if args.raw_dir
        else project_root / "audit" / "upstream_review" / "raw"
    )
    output_dir = (
        args.output_dir.resolve()
        if args.output_dir
        else project_root / "audit" / "upstream_review"
    )
    if not raw_dir.is_dir():
        print(f"ERROR: raw audit directory not found: {raw_dir}", file=sys.stderr)
        return 2
    try:
        audits = load_audits(raw_dir)
        registry_rows = [
            registry_row(data, slug, path, raw_dir)
            for slug, data, path in audits
        ]
        capability_rows = [
            capability_row(data, slug) for slug, data, _ in audits
        ]
        license_rows = [license_row(data, slug) for slug, data, _ in audits]
        dependencies: list[dict[str, str]] = []
        for slug, data, _ in audits:
            dependencies.extend(dependency_rows(data, slug))

        output_paths = {
            "registry": output_dir / "upstream_registry.csv",
            "capability": output_dir / "upstream_capability_matrix.csv",
            "license": output_dir / "upstream_license_matrix.csv",
            "dependency": output_dir / "upstream_dependency_matrix.csv",
            "security": output_dir / "upstream_security_review.md",
            "features": output_dir / "upstream_feature_notes.md",
            "rejections": output_dir / "upstream_rejection_reasons.md",
        }
        atomic_write_csv(output_paths["registry"], registry_rows)
        atomic_write_csv(output_paths["capability"], capability_rows)
        atomic_write_csv(output_paths["license"], license_rows)
        atomic_write_csv(output_paths["dependency"], dependencies)
        atomic_write_text(output_paths["security"], security_review_markdown(audits))
        atomic_write_text(output_paths["features"], feature_notes_markdown(audits))
        atomic_write_text(output_paths["rejections"], rejection_reasons_markdown(audits))

        validate_csv(output_paths["registry"], 6, exactly_one_row_per_project=True)
        validate_csv(output_paths["capability"], 6, exactly_one_row_per_project=True)
        validate_csv(output_paths["license"], 6, exactly_one_row_per_project=True)
        validate_csv(output_paths["dependency"], 6, exactly_one_row_per_project=False)

        print("PHASE1_AGGREGATION=PASS")
        print("PROJECTS=6")
        print(f"DEPENDENCY_ROWS={len(dependencies)}")
        for output_path in output_paths.values():
            print(f"OUTPUT={output_path}")
        return 0
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError) as error:
        print(f"PHASE1_AGGREGATION=BLOCKED", file=sys.stderr)
        print(f"ERROR={error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
