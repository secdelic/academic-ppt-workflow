from __future__ import annotations

import csv
from pathlib import Path

from .utils import write_csv, write_json


STORYBOARD_FIELDS = [
    "slide_id",
    "slide_title",
    "slide_purpose",
    "single_key_message",
    "source_ids",
    "proposed_layout",
    "visual_type",
    "citation_requirement",
    "speaker_note_summary",
    "confidence",
    "manual_review_required",
]


def _find_chart_data(input_root: Path, manifest: list[dict[str, str]]) -> tuple[list[str], list[float], str, str] | None:
    for row in manifest:
        if row["file_type"] not in {"csv", "tsv"}:
            continue
        delimiter = "\t" if row["file_type"] == "tsv" else ","
        path = input_root / row["relative_path"]
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            records = list(csv.DictReader(handle, delimiter=delimiter))
        if not records:
            continue
        headers = {h.lower(): h for h in (records[0].keys() if records else [])}
        if "category" not in headers or "value" not in headers:
            continue
        categories: list[str] = []
        values: list[float] = []
        try:
            for record in records[:8]:
                categories.append(record[headers["category"]])
                values.append(float(record[headers["value"]]))
        except (TypeError, ValueError):
            continue
        unit_header = headers.get("unit")
        unit = records[0].get(unit_header, "") if unit_header else ""
        return categories, values, row["source_id"], unit
    return None


def create_storyboard(
    brief: dict,
    claims: list[dict[str, str]],
    manifest: list[dict[str, str]],
    input_root: Path,
    staging_root: Path,
) -> list[dict[str, str]]:
    project = brief["project_name"]
    source_ids = ";".join(row["source_id"] for row in manifest)
    slides: list[dict[str, str]] = [
        {
            "slide_title": project,
            "slide_purpose": "Introduce the project and presentation objective",
            "single_key_message": brief["presentation_objective"],
            "source_ids": source_ids,
            "proposed_layout": "minimal_cover",
            "visual_type": "typography",
            "citation_requirement": "none unless title uses source-specific facts",
            "speaker_note_summary": "State scope and that all reported content is source-bound.",
            "confidence": "high" if brief["presentation_objective"] != "INFORMATION_REQUIRED" else "low",
            "manual_review_required": "yes",
        },
        {
            "slide_title": "The source package defines the presentation boundary",
            "slide_purpose": "State the evidence boundary before interpretation",
            "single_key_message": f"{len(manifest)} source file(s) were registered with SHA-256 provenance.",
            "source_ids": source_ids,
            "proposed_layout": "text_plus_source_list",
            "visual_type": "source_map",
            "citation_requirement": "source manifest",
            "speaker_note_summary": "Explain that unprovided information is not inferred.",
            "confidence": "high",
            "manual_review_required": "no",
        },
    ]

    max_claim_slides = max(1, int(brief.get("target_slide_count", 10)) - 5)
    chart = _find_chart_data(input_root, manifest)
    for index, claim in enumerate(claims[:max_claim_slides]):
        slides.append(
            {
                "slide_title": claim["allowed_wording"],
                "slide_purpose": "Present one source-bound result",
                "single_key_message": claim["claim_text"],
                "source_ids": claim["source_id"],
                "proposed_layout": "result_with_chart" if index == 0 and chart else "result_statement",
                "visual_type": "editable_bar_chart" if index == 0 and chart else "evidence_statement",
                "citation_requirement": "required",
                "speaker_note_summary": f"Source: {claim['source_location']}. Avoid: {claim['prohibited_overstatement']}",
                "confidence": claim["confidence"],
                "manual_review_required": "yes",
            }
        )

    if not claims:
        slides.append(
            {
                "slide_title": "Results require source material",
                "slide_purpose": "Expose the scientific blocker",
                "single_key_message": "INFORMATION_REQUIRED",
                "source_ids": "",
                "proposed_layout": "neutral_notice",
                "visual_type": "text",
                "citation_requirement": "none",
                "speaker_note_summary": "Do not present results until source-bound claims are supplied.",
                "confidence": "low",
                "manual_review_required": "yes",
            }
        )

    slides.extend(
        [
            {
                "slide_title": "Interpretation remains bounded by the supplied evidence",
                "slide_purpose": "Prevent causal or mechanistic overstatement",
                "single_key_message": "Associations and computational inferences must retain their original evidence strength.",
                "source_ids": source_ids,
                "proposed_layout": "two_column_guardrails",
                "visual_type": "comparison",
                "citation_requirement": "source-bound",
                "speaker_note_summary": "Discuss limitations and uncertainty explicitly.",
                "confidence": "high",
                "manual_review_required": "yes",
            },
            {
                "slide_title": "Human scientific review is the final release gate",
                "slide_purpose": "Close with the assisted-use boundary",
                "single_key_message": brief.get("key_message", "INFORMATION_REQUIRED"),
                "source_ids": ";".join(dict.fromkeys(c["source_id"] for c in claims)),
                "proposed_layout": "conclusion",
                "visual_type": "summary",
                "citation_requirement": "required for scientific conclusion",
                "speaker_note_summary": "Confirm every claim, number, citation, and omission before external use.",
                "confidence": "medium" if claims else "low",
                "manual_review_required": "yes",
            },
        ]
    )

    for index, slide in enumerate(slides, start=1):
        slide["slide_id"] = f"SLD-{index:03d}"

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
    (staging_root / "deck_outline.md").write_text("\n".join(outline), encoding="utf-8")
    write_csv(staging_root / "storyboard.csv", STORYBOARD_FIELDS, slides)

    chart_spec = None
    if chart:
        chart_spec = {
            "categories": chart[0],
            "values": chart[1],
            "source_id": chart[2],
            "unit": chart[3],
        }
    spec_slides = []
    for slide in slides:
        sources = [value for value in slide["source_ids"].split(";") if value]
        item = dict(slide)
        item["source_ids"] = sources
        if slide["visual_type"] == "source_map":
            item["source_records"] = [
                {
                    "source_id": row["source_id"],
                    "file_name": row["file_name"],
                    "file_type": row["file_type"],
                }
                for row in manifest[:6]
            ]
        if slide["visual_type"] == "editable_bar_chart":
            item["chart_data"] = chart_spec
        spec_slides.append(item)
    write_json(
        staging_root / "build_spec.json",
        {
            "brief": brief,
            "slides": spec_slides,
            "design_profile": "restrained_biomedical",
        },
    )
    return slides


def write_speaker_notes(staging_root: Path, slides: list[dict[str, str]]) -> None:
    lines = ["# Speaker Notes", ""]
    for slide in slides:
        lines.extend(
            [
                f"## {slide['slide_id']} - {slide['slide_title']}",
                "",
                slide["speaker_note_summary"],
                "",
                "[Sources]",
                *(f"- {source}" for source in slide["source_ids"].split(";") if source),
                "",
            ]
        )
    (staging_root / "speaker_notes.md").write_text("\n".join(lines), encoding="utf-8")
