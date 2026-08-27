from __future__ import annotations

from pathlib import Path
from typing import Any

from .chart_data import discover_chart_contracts
from .deck_ir import build_deck_ir, to_backend_spec, to_storyboard_rows
from .layout_contract import (
    LayoutContract,
    plan_slide_geometry,
    validate_planned_geometry,
)
from .inventory import is_scientific_source
from .scientific_visuals import TEMPLATE_IDS, build_visual_spec
from .source_display import short_source_label
from .utils import load_yaml_compatible, write_csv, write_json
from .visual_router import route_chart_contracts, route_unstructured_claim


STORYBOARD_FIELDS = [
    "slide_id",
    "claim_ids",
    "slide_title",
    "slide_purpose",
    "single_key_message",
    "source_ids",
    "proposed_layout",
    "layout_family",
    "visual_type",
    "short_source_label",
    "citation_requirement",
    "speaker_note_summary",
    "prohibited_overstatement",
    "confidence",
    "manual_review_required",
]


def _claim_evidence_strength(claim: dict[str, str]) -> str:
    raw = claim.get("evidence_type", "").strip().lower()
    normalized = raw.replace("-", "_").replace(" ", "_")
    aliases = {
        "reported_source_text": "reported_fact",
        "reported": "reported_fact",
        "fact": "reported_fact",
        "association": "observational_association",
        "observational": "observational_association",
        "prediction": "computational_inference",
        "enrichment": "computational_inference",
        "machine_learning": "computational_inference",
        "molecular_docking": "computational_inference",
        "molecular_dynamics": "computational_inference",
        "mechanism_hypothesis": "proposed_mechanism",
        "hypothesis": "hypothesis_only",
    }
    allowed = {
        "demonstrated_mechanism",
        "supported_interpretation",
        "reported_fact",
        "observational_association",
        "computational_inference",
        "proposed_mechanism",
        "hypothesis_only",
        "unknown",
    }
    normalized = aliases.get(normalized, normalized)
    return normalized if normalized in allowed else "unknown"


def _id_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [item.strip() for item in value.split(";") if item.strip()]
    return [str(item).strip() for item in value if str(item).strip()]


def _load_layout_governance(repo_root: Path) -> tuple[LayoutContract, dict, dict]:
    package_root = Path(__file__).resolve().parents[2]
    config_root = repo_root / "config"
    if not (config_root / "layout_contract.yaml").is_file():
        config_root = package_root / "config"
    layout = LayoutContract.from_files(
        config_root / "layout_contract.yaml",
        config_root / "safe_zones.yaml",
    )
    footer = load_yaml_compatible(config_root / "footer_policy.yaml")
    visual_templates = load_yaml_compatible(
        config_root / "visual_templates.yaml"
    )
    return layout, footer, visual_templates


def _append_scientific_visual(
    planned_slides: list[dict[str, Any]],
    brief: dict,
    all_source_ids: list[str],
    template_layout_id: str | None,
) -> None:
    requested = str(brief.get("scientific_visual_template", "")).strip()
    if not requested:
        return
    if requested not in TEMPLATE_IDS:
        raise ValueError("Unknown scientific_visual_template: " + requested)
    kwargs = (
        {
            "evidence_status": str(
                brief.get("mechanism_evidence_status", "hypothesis_only")
            )
        }
        if requested == "mechanism_hypothesis"
        else {}
    )
    diagram = build_visual_spec(requested, **kwargs)
    planned_slides.append(
        {
            "logical_slide_key": f"scientific-visual:{requested}",
            "slide_role": "study_design",
            "section_id": "methods",
            "slide_title": str(
                brief.get(
                    "scientific_visual_title",
                    requested.replace("_", " ").title(),
                )
            ),
            "slide_purpose": "Summarize a source-bounded scientific structure",
            "single_key_message": str(
                brief.get(
                    "scientific_visual_message",
                    "The diagram summarizes source-defined structure only.",
                )
            ),
            "source_ids": all_source_ids,
            "claim_ids": [],
            "proposed_layout": "scientific_diagram",
            "layout_family": "native_flow",
            "visual_type": "scientific_diagram",
            "diagram_spec": diagram,
            "visual_assets": [
                {
                    "asset_id": "VIS-" + diagram["content_hash"][:16],
                    "asset_type": "powerpoint_shapes_or_local_svg",
                    "content_hash": diagram["content_hash"],
                }
            ],
            "editable_object_requirements": [
                "native PowerPoint shapes preferred"
            ],
            "citation_requirement": "source-bound",
            "speaker_note_summary": (
                "Treat the visual as a structural summary. Do not upgrade "
                "interpretation or hypothesis into demonstrated mechanism."
            ),
            "confidence": "medium",
            "manual_review_required": "yes",
            "template_layout_id": template_layout_id,
        }
    )


def _attach_layout_governance(
    planned_slides: list[dict[str, Any]],
    manifest: list[dict[str, str]],
    layout: LayoutContract,
    footer_policy: dict,
    staging_root: Path,
    *,
    audit_presentation: bool,
) -> None:
    rows: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    for slide in planned_slides:
        source_ids = _id_list(slide.get("source_ids"))
        slide["source_ids"] = source_ids
        slide["short_source_label"] = short_source_label(
            source_ids,
            manifest,
            footer_policy,
            audit_presentation=audit_presentation,
        )
        slide["layout_family"] = str(
            slide.get(
                "layout_family",
                slide.get("visual_type", slide.get("proposed_layout", "content")),
            )
        )
        geometry = plan_slide_geometry(slide, layout)
        slide["planned_geometry"] = geometry
        issues = validate_planned_geometry(geometry, layout)
        for item in geometry:
            rows.append(
                {
                    "logical_slide_key": slide["logical_slide_key"],
                    "object_id": item["object_id"],
                    "kind": item["kind"],
                    "role": item["role"],
                    "zone": item["zone"],
                    **item["bounds"],
                    "status": "PASS",
                }
            )
        for issue in issues:
            failures.append(
                {
                    "logical_slide_key": slide["logical_slide_key"],
                    **issue,
                }
            )
    write_csv(
        staging_root / "planned_geometry_manifest.csv",
        [
            "logical_slide_key",
            "object_id",
            "kind",
            "role",
            "zone",
            "x",
            "y",
            "w",
            "h",
            "status",
        ],
        rows,
    )
    write_json(
        staging_root / "planned_geometry_qa.json",
        {
            "schema_version": "2.1",
            "status": "PASS" if not failures else "FAIL",
            "issues": failures,
            "layout_contract": layout.serializable(),
        },
    )
    if failures:
        preview = "; ".join(
            f"{item.get('logical_slide_key')}:{item.get('code')}"
            for item in failures[:8]
        )
        raise ValueError("Planned geometry QA failed: " + preview)


def create_storyboard(
    brief: dict,
    claims: list[dict[str, str]],
    manifest: list[dict[str, str]],
    input_root: Path,
    staging_root: Path,
    *,
    existing_ir: dict | None = None,
    narrative_mode: str | None = None,
    template_layout_id: str | None = None,
    style_profile: dict | None = None,
    repo_root: Path | None = None,
) -> list[dict[str, str]]:
    project = brief["project_name"]
    repo = (repo_root or Path(__file__).resolve().parents[2]).resolve()
    layout, footer_policy, visual_templates = _load_layout_governance(repo)
    chinese = str(brief.get("language", "")).lower() in {
        "zh",
        "zh-cn",
        "chinese",
        "中文",
        "简体中文",
    }
    scientific_manifest = [
        row
        for row in manifest
        if is_scientific_source(row)
    ]
    all_source_ids = [row["source_id"] for row in scientific_manifest]
    planned_slides: list[dict[str, Any]] = [
        {
            "logical_slide_key": "cover",
            "slide_role": "cover",
            "section_id": "opening",
            "slide_title": project,
            "slide_purpose": "Introduce the project and presentation objective",
            "single_key_message": brief["presentation_objective"],
            "source_ids": all_source_ids,
            "proposed_layout": "minimal_cover",
            "layout_family": "cover",
            "visual_type": "typography",
            "citation_requirement": "none unless title uses source-specific facts",
            "speaker_note_summary": (
                "State scope and that all reported content is source-bound."
            ),
            "confidence": (
                "high"
                if brief["presentation_objective"] != "INFORMATION_REQUIRED"
                else "low"
            ),
            "manual_review_required": "yes",
            "template_layout_id": template_layout_id,
        },
        {
            "logical_slide_key": "evidence-boundary",
            "slide_role": "evidence_boundary",
            "section_id": "opening",
            "slide_title": (
                "输入材料界定本次汇报的证据边界"
                if chinese
                else "The source package defines the presentation boundary"
            ),
            "slide_purpose": (
                "在解释前声明证据边界"
                if chinese
                else "State the evidence boundary before interpretation"
            ),
            "single_key_message": (
                f"已登记 {len(scientific_manifest)} 个科研来源文件；完整标识与哈希保留在审计产物。"
                if chinese
                else f"{len(scientific_manifest)} scientific sources are registered; full identifiers remain in audit artifacts."
            ),
            "source_ids": all_source_ids,
            "proposed_layout": "source_registry",
            "layout_family": "source_registry",
            "visual_type": "source_registry",
            "citation_requirement": "source manifest",
            "speaker_note_summary": (
                "Explain that unprovided information is not inferred. "
                "The full source IDs, paths, and hashes are retained in notes and audit files."
            ),
            "confidence": "high",
            "manual_review_required": "no",
            "template_layout_id": template_layout_id,
            "source_records": [
                {
                    "file_name": row["file_name"],
                    "file_type": row["file_type"],
                    "source_id": row["source_id"],
                    "relative_path": row["relative_path"],
                }
                for row in scientific_manifest
            ],
        },
    ]
    _append_scientific_visual(
        planned_slides, brief, all_source_ids, template_layout_id
    )

    contracts = discover_chart_contracts(input_root, manifest)
    write_json(
        staging_root / "chart_data_contracts.json",
        {"schema_version": "2.1", "contracts": contracts},
    )
    routed_chart_slides, consumed = route_chart_contracts(claims, contracts)
    chart_by_source = {
        slide["source_ids"][0]: slide for slide in routed_chart_slides
    }
    emitted_sources: set[str] = set()
    for claim in claims:
        source_id = claim["source_id"]
        if source_id in chart_by_source:
            if source_id not in emitted_sources:
                planned_slides.append(chart_by_source[source_id])
                emitted_sources.add(source_id)
            continue
        if claim["claim_id"] in consumed:
            continue
        route = route_unstructured_claim(claim)
        visual_type = route.pop("visual_type")
        layout_family = route.pop("layout_family")
        planned_slides.append(
            {
                "logical_slide_key": (
                    f"result:{source_id}:{claim['source_location']}:"
                    f"{claim['claim_id']}"
                ),
                "slide_role": "result",
                "section_id": "results",
                "slide_title": claim["allowed_wording"],
                "slide_purpose": "Present one source-bound result",
                "single_key_message": claim["claim_text"],
                "source_ids": [source_id],
                "claim_ids": [claim["claim_id"]],
                "proposed_layout": layout_family,
                "layout_family": layout_family,
                "visual_type": visual_type,
                "visual_asset_path": claim.get("visual_asset_path", ""),
                "citation_requirement": "required",
                "speaker_note_summary": (
                    f"Source path/location: {claim['source_location']}. "
                    f"Claim: {claim['claim_id']}. "
                    f"Avoid: {claim['prohibited_overstatement']}"
                ),
                "confidence": claim["confidence"],
                "manual_review_required": "yes",
                "evidence_strength": _claim_evidence_strength(claim),
                "prohibited_overstatement": claim["prohibited_overstatement"],
                "template_layout_id": template_layout_id,
                **route,
            }
        )

    if not claims:
        planned_slides.append(
            {
                "logical_slide_key": "results-information-required",
                "slide_role": "result",
                "section_id": "results",
                "slide_title": "Results require source material",
                "slide_purpose": "Expose the scientific blocker",
                "single_key_message": "INFORMATION_REQUIRED",
                "source_ids": [],
                "proposed_layout": "takeaway",
                "layout_family": "takeaway",
                "visual_type": "takeaway",
                "takeaways": ["INFORMATION_REQUIRED"],
                "citation_requirement": "none",
                "speaker_note_summary": (
                    "Do not present results until source-bound claims are supplied."
                ),
                "confidence": "low",
                "manual_review_required": "yes",
                "template_layout_id": template_layout_id,
            }
        )

    planned_slides.extend(
        [
            {
                "logical_slide_key": "interpretation-guardrails",
                "slide_role": "limitations",
                "section_id": "interpretation",
                "slide_title": (
                    "所有解释均须保留原始证据强度"
                    if chinese
                    else "Interpretation remains bounded by the supplied evidence"
                ),
                "slide_purpose": (
                    "防止因果化或机制升级"
                    if chinese
                    else "Prevent causal or mechanistic overstatement"
                ),
                "single_key_message": (
                    "观察性关联不能升级为因果结论，机制假设不能表述为机制证明。"
                    if chinese
                    else "Associations and computational inferences must retain their original evidence strength."
                ),
                "source_ids": all_source_ids,
                "proposed_layout": "comparison",
                "layout_family": "comparison",
                "visual_type": "comparison",
                "citation_requirement": "source-bound",
                "speaker_note_summary": (
                    "Discuss limitations and uncertainty explicitly."
                ),
                "confidence": "high",
                "manual_review_required": "yes",
                "template_layout_id": template_layout_id,
            },
            {
                "logical_slide_key": "conclusion-author-gate",
                "slide_role": "conclusion",
                "section_id": "conclusion",
                "slide_title": (
                    "人工科学审核是最终交付门槛"
                    if chinese
                    else "Human scientific review is the final release gate"
                ),
                "slide_purpose": (
                    "以辅助使用边界收束汇报"
                    if chinese
                    else "Close with the assisted-use boundary"
                ),
                "single_key_message": brief.get(
                    "key_message", "INFORMATION_REQUIRED"
                ),
                "source_ids": list(
                    dict.fromkeys(claim["source_id"] for claim in claims)
                ),
                "proposed_layout": "takeaway",
                "layout_family": "takeaway",
                "visual_type": "takeaway",
                "takeaways": [
                    brief.get("key_message", "INFORMATION_REQUIRED"),
                    "最终科学批准由用户完成。",
                ],
                "citation_requirement": "required for scientific conclusion",
                "speaker_note_summary": (
                    "Confirm every claim, number, citation, and omission before external use."
                ),
                "confidence": "medium" if claims else "low",
                "manual_review_required": "yes",
                "template_layout_id": template_layout_id,
            },
        ]
    )

    _attach_layout_governance(
        planned_slides,
        manifest,
        layout,
        footer_policy,
        staging_root,
        audit_presentation=bool(brief.get("audit_presentation", False)),
    )
    deck_ir = build_deck_ir(
        brief,
        planned_slides,
        existing_ir=existing_ir,
        narrative_mode=narrative_mode,
    )
    write_json(staging_root / "deck_ir.json", deck_ir)
    slides = to_storyboard_rows(deck_ir)

    outline = ["# Deck Outline", "", f"Project: {project}", ""]
    for slide in slides:
        outline.extend(
            [
                f"## {slide['slide_id']} - {slide['slide_title']}",
                "",
                f"- Purpose: {slide['slide_purpose']}",
                f"- Key message: {slide['single_key_message']}",
                f"- Sources: {slide['source_ids'] or 'INFORMATION_REQUIRED'}",
                f"- Layout: {slide['proposed_layout']}",
                "",
            ]
        )
    (staging_root / "deck_outline.md").write_text(
        "\n".join(outline), encoding="utf-8"
    )
    write_csv(staging_root / "storyboard.csv", STORYBOARD_FIELDS, slides)

    backend_spec = to_backend_spec(deck_ir, brief, "restrained_biomedical")
    backend_spec["layout_contract"] = layout.serializable()
    backend_spec["footer_policy"] = footer_policy
    backend_spec["visual_templates"] = visual_templates
    if style_profile is not None:
        backend_spec["style_profile"] = style_profile
    write_json(staging_root / "build_spec.json", backend_spec)
    return slides


def write_speaker_notes(
    staging_root: Path,
    slides: list[dict[str, str]],
) -> None:
    lines = ["# Speaker Notes", ""]
    for slide in slides:
        lines.extend(
            [
                f"## {slide['slide_id']} - {slide['slide_title']}",
                "",
                slide["speaker_note_summary"],
                "",
                "[Claims]",
                *(
                    f"- {claim}"
                    for claim in str(slide.get("claim_ids", "")).split(";")
                    if claim
                ),
                "",
                "[Wording-Boundary]",
                str(
                    slide.get(
                        "prohibited_overstatement",
                        "Do not exceed the registered evidence strength.",
                    )
                ),
                "",
                "[Sources]",
                *(
                    f"- {source}"
                    for source in slide["source_ids"].split(";")
                    if source
                ),
                "",
            ]
        )
    (staging_root / "speaker_notes.md").write_text(
        "\n".join(lines), encoding="utf-8"
    )
