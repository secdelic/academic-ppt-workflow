from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from .utils import write_csv


DENSITY_FIELDS = [
    "slide_id",
    "slide_role",
    "layout_family",
    "content_coverage",
    "text_area_ratio",
    "visual_area_ratio",
    "whitespace_ratio",
    "object_count",
    "text_character_count",
    "visual_anchor_count",
    "bottom_zone_occupancy",
    "status",
    "reason",
]


def _area(bounds: Mapping[str, Any]) -> float:
    return max(0.0, float(bounds.get("w", 0))) * max(
        0.0, float(bounds.get("h", 0))
    )


def analyze_visual_density(
    deck_ir: Mapping[str, Any],
    *,
    slide_width: float = 13.333,
    slide_height: float = 7.5,
    preferred_min: float = 0.28,
    preferred_max: float = 0.72,
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[str],
]:
    slide_area = slide_width * slide_height
    density_rows: list[dict[str, Any]] = []
    repetition_rows: list[dict[str, Any]] = []
    low_rows: list[dict[str, Any]] = []
    overloaded_rows: list[dict[str, Any]] = []
    issues: list[str] = []
    previous_layout = ""
    repetition_count = 0
    text_only_streak = 0
    for slide in deck_ir.get("slides", []):
        objects = list(slide.get("planned_geometry", []))
        content_objects = [
            item
            for item in objects
            if item.get("role") not in {"title", "footer", "page_number"}
        ]
        text_objects = [
            item for item in content_objects if item.get("kind") == "text"
        ]
        visual_objects = [
            item
            for item in content_objects
            if item.get("kind") in {"chart", "picture", "diagram", "table"}
        ]
        content_area = min(slide_area, sum(_area(item.get("bounds", {})) for item in content_objects))
        text_area = min(slide_area, sum(_area(item.get("bounds", {})) for item in text_objects))
        visual_area = min(slide_area, sum(_area(item.get("bounds", {})) for item in visual_objects))
        coverage = content_area / slide_area
        text_ratio = text_area / slide_area
        visual_ratio = visual_area / slide_area
        whitespace = max(0.0, 1.0 - coverage)
        chars = sum(len(str(item.get("text", ""))) for item in text_objects)
        bottom_occupancy = sum(
            _area(item.get("bounds", {}))
            for item in content_objects
            if float(item.get("bounds", {}).get("y", 0))
            + float(item.get("bounds", {}).get("h", 0))
            > slide_height * 0.82
        ) / slide_area
        role = str(slide.get("slide_role", "content"))
        layout = str(
            slide.get("layout_family", slide.get("visual_type", "unknown"))
        )
        exempt_low = role in {"cover", "section_divider", "conclusion"}
        reasons: list[str] = []
        if coverage < preferred_min and not exempt_low:
            reasons.append("coverage_below_preferred_min")
        if coverage > preferred_max:
            reasons.append("coverage_above_preferred_max")
        if not visual_objects and chars < 120 and not exempt_low:
            reasons.append("low_information_text_only")
        if chars > 480:
            reasons.append("text_character_budget_exceeded")
        status = "PASS" if not reasons else "FAIL"
        row = {
            "slide_id": slide.get("slide_id", ""),
            "slide_role": role,
            "layout_family": layout,
            "content_coverage": f"{coverage:.4f}",
            "text_area_ratio": f"{text_ratio:.4f}",
            "visual_area_ratio": f"{visual_ratio:.4f}",
            "whitespace_ratio": f"{whitespace:.4f}",
            "object_count": len(content_objects),
            "text_character_count": chars,
            "visual_anchor_count": len(visual_objects),
            "bottom_zone_occupancy": f"{bottom_occupancy:.4f}",
            "status": status,
            "reason": ";".join(reasons),
        }
        density_rows.append(row)
        if "low_information_text_only" in reasons:
            low_rows.append(row)
        if any(
            reason in reasons
            for reason in (
                "coverage_above_preferred_max",
                "text_character_budget_exceeded",
            )
        ):
            overloaded_rows.append(row)

        if layout == previous_layout:
            repetition_count += 1
        else:
            previous_layout = layout
            repetition_count = 1
        repetition_status = "PASS" if repetition_count <= 2 else "FAIL"
        repetition_rows.append(
            {
                "slide_id": slide.get("slide_id", ""),
                "layout_family": layout,
                "consecutive_count": repetition_count,
                "status": repetition_status,
            }
        )
        if repetition_status == "FAIL":
            issues.append(
                f"{slide.get('slide_id')}: layout {layout} repeats "
                f"{repetition_count} consecutive times"
            )
        is_text_only = not visual_objects
        text_only_streak = text_only_streak + 1 if is_text_only else 0
        if text_only_streak > 2 and not exempt_low:
            issues.append(
                f"{slide.get('slide_id')}: more than two consecutive text-only slides"
            )
        if status == "FAIL":
            issues.extend(
                f"{slide.get('slide_id')}: density {reason}" for reason in reasons
            )
    return density_rows, repetition_rows, low_rows, overloaded_rows, issues


def write_density_reports(
    output_dir,
    deck_ir: Mapping[str, Any],
) -> list[str]:
    density, repetition, low, overloaded, issues = analyze_visual_density(deck_ir)
    write_csv(output_dir / "visual_density_report.csv", DENSITY_FIELDS, density)
    write_csv(
        output_dir / "layout_repetition_report.csv",
        ["slide_id", "layout_family", "consecutive_count", "status"],
        repetition,
    )
    write_csv(output_dir / "low_information_slide_report.csv", DENSITY_FIELDS, low)
    write_csv(output_dir / "overloaded_slide_report.csv", DENSITY_FIELDS, overloaded)
    return issues


__all__ = [
    "DENSITY_FIELDS",
    "analyze_visual_density",
    "write_density_reports",
]
