from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Mapping

from .text_measurement import choose_font_size
from .utils import load_yaml_compatible


@dataclass(frozen=True)
class Rect:
    x: float
    y: float
    w: float
    h: float

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    def to_dict(self) -> dict[str, float]:
        return asdict(self)


def _ratio_rect(
    zone: Mapping[str, Any],
    width: float,
    height: float,
) -> Rect:
    x_min = float(zone["x_min"])
    x_max = float(zone["x_max"])
    y_min = float(zone["y_min"])
    y_max = float(zone["y_max"])
    return Rect(
        x=x_min * width,
        y=y_min * height,
        w=(x_max - x_min) * width,
        h=(y_max - y_min) * height,
    )


def _intersection(a: Rect, b: Rect) -> Rect | None:
    left = max(a.x, b.x)
    top = max(a.y, b.y)
    right = min(a.right, b.right)
    bottom = min(a.bottom, b.bottom)
    if right <= left or bottom <= top:
        return None
    return Rect(left, top, right - left, bottom - top)


def _contains(outer: Rect, inner: Rect, tolerance: float = 0.002) -> bool:
    return (
        inner.x >= outer.x - tolerance
        and inner.y >= outer.y - tolerance
        and inner.right <= outer.right + tolerance
        and inner.bottom <= outer.bottom + tolerance
    )


class LayoutContract:
    def __init__(
        self,
        contract: Mapping[str, Any],
        safe_zones: Mapping[str, Any],
        *,
        aspect_ratio: str = "16:9",
    ) -> None:
        profile = contract["aspect_ratios"][aspect_ratio]
        self.aspect_ratio = aspect_ratio
        self.width = float(profile["width_in"])
        self.height = float(profile["height_in"])
        self.zones = {
            name: _ratio_rect(zone, self.width, self.height)
            for name, zone in profile["zones"].items()
        }
        self.zones["slide"] = Rect(0.0, 0.0, self.width, self.height)
        self.minimum_gutter = float(contract.get("minimum_gutter_in", 0.12))
        self.minimum_font_pt = dict(contract.get("minimum_font_pt", {}))
        self.text_policy = dict(contract.get("text_policy", {}))
        self.safe_zones = dict(safe_zones)

    @classmethod
    def from_files(
        cls,
        contract_path: Path,
        safe_zones_path: Path,
        *,
        aspect_ratio: str = "16:9",
    ) -> "LayoutContract":
        return cls(
            load_yaml_compatible(contract_path),
            load_yaml_compatible(safe_zones_path),
            aspect_ratio=aspect_ratio,
        )

    def serializable(self) -> dict[str, Any]:
        return {
            "schema_version": "2.1",
            "aspect_ratio": self.aspect_ratio,
            "slide": {"width_in": self.width, "height_in": self.height},
            "zones": {name: rect.to_dict() for name, rect in self.zones.items()},
            "minimum_gutter_in": self.minimum_gutter,
            "minimum_font_pt": self.minimum_font_pt,
            "text_policy": self.text_policy,
            "safe_zones": self.safe_zones,
        }


def _object(
    object_id: str,
    kind: str,
    role: str,
    zone: str,
    rect: Rect,
    *,
    text: str = "",
    font_size_pt: float = 0.0,
    bold: bool = False,
    allowed_overlap_with: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "object_id": object_id,
        "kind": kind,
        "role": role,
        "zone": zone,
        "bounds": rect.to_dict(),
        "text": text,
        "font_size_pt": font_size_pt,
        "bold": bold,
        "allowed_overlap_with": allowed_overlap_with or [],
    }


def plan_slide_geometry(
    slide: Mapping[str, Any],
    contract: LayoutContract,
) -> list[dict[str, Any]]:
    """Return the complete layout frame map consumed by every backend layout."""

    width, height = contract.width, contract.height
    title_zone = contract.zones["title"]
    content_zone = contract.zones["content"]
    source_zone = contract.zones["source_label"]
    page_zone = contract.zones["page_number"]
    slide_id = str(slide.get("slide_id", slide.get("logical_slide_key", "slide")))
    title = str(slide.get("slide_title", ""))
    message = str(slide.get("single_key_message", slide.get("key_message", "")))
    visual = str(slide.get("visual_type", "evidence_statement"))
    objects: list[dict[str, Any]] = []
    if str(slide.get("slide_role", "")) == "cover":
        objects.extend(
            [
                _object(
                    f"{slide_id}:cover-title",
                    "text",
                    "title",
                    "slide",
                    Rect(0.83, 1.35, 10.85, 2.0),
                    text=title,
                    font_size_pt=50,
                    bold=True,
                ),
                _object(
                    f"{slide_id}:cover-message",
                    "text",
                    "body",
                    "slide",
                    Rect(0.83, 3.78, 9.45, 1.15),
                    text=message,
                    font_size_pt=21,
                ),
            ]
        )
        return objects

    objects.append(
        _object(
            f"{slide_id}:title",
            "text",
            "title",
            "title",
            Rect(title_zone.x, title_zone.y, title_zone.w, title_zone.h),
            text=title,
            font_size_pt=32,
            bold=True,
        )
    )
    short_source = str(slide.get("short_source_label", "来源：输入材料"))
    objects.extend(
        [
            _object(
                f"{slide_id}:source",
                "text",
                "footer",
                "source_label",
                source_zone,
                text=short_source,
                font_size_pt=10,
            ),
            _object(
                f"{slide_id}:page",
                "text",
                "page_number",
                "page_number",
                page_zone,
                text="99",
                font_size_pt=11,
            ),
        ]
    )

    x, y, w, h = (
        content_zone.x,
        content_zone.y,
        content_zone.w,
        content_zone.h,
    )
    if visual in {
        "event_rate_chart",
        "editable_bar_chart",
        "missingness_chart",
        "line_chart",
        "source_figure",
    }:
        objects.extend(
            [
                _object(
                    f"{slide_id}:message",
                    "text",
                    "body",
                    "content",
                    Rect(x, y + 0.12, w * 0.31, h * 0.70),
                    text=message,
                    font_size_pt=22,
                    bold=True,
                ),
                _object(
                    f"{slide_id}:visual",
                    "picture" if visual == "source_figure" else "chart",
                    "visual",
                    "content",
                    Rect(x + w * 0.36, y + 0.02, w * 0.64, h * 0.88),
                ),
                _object(
                    f"{slide_id}:metadata",
                    "text",
                    "chart_label",
                    "content",
                    Rect(x, y + h * 0.76, w * 0.31, h * 0.16),
                    text=str(slide.get("chart_caption", "")),
                    font_size_pt=15,
                ),
            ]
        )
    elif visual in {"forest_plot", "subgroup_forest_plot"}:
        objects.extend(
            [
                _object(
                    f"{slide_id}:message",
                    "text",
                    "body",
                    "content",
                    Rect(x, y + 0.02, w, h * 0.18),
                    text=message,
                    font_size_pt=20,
                    bold=True,
                ),
                _object(
                    f"{slide_id}:visual",
                    "chart",
                    "visual",
                    "content",
                    Rect(x, y + h * 0.24, w, h * 0.68),
                ),
            ]
        )
    elif visual == "timeline":
        objects.extend(
            [
                _object(
                    f"{slide_id}:message",
                    "text",
                    "body",
                    "content",
                    Rect(x, y + 0.04, w, h * 0.18),
                    text=message,
                    font_size_pt=21,
                    bold=True,
                ),
                _object(
                    f"{slide_id}:timeline",
                    "diagram",
                    "visual",
                    "content",
                    Rect(x + 0.2, y + h * 0.30, w - 0.4, h * 0.48),
                ),
            ]
        )
    elif visual in {"matrix_2x2", "comparison", "problem_gap_objective", "takeaway"}:
        objects.extend(
            [
                _object(
                    f"{slide_id}:message",
                    "text",
                    "body",
                    "content",
                    Rect(x, y + 0.01, w, h * 0.22),
                    text=message,
                    font_size_pt=20,
                    bold=True,
                ),
                _object(
                    f"{slide_id}:structure",
                    "diagram",
                    "visual",
                    "content",
                    Rect(x, y + h * 0.28, w, h * 0.64),
                ),
            ]
        )
    elif visual in {"audit_table", "source_registry", "native_flow"}:
        objects.append(
            _object(
                f"{slide_id}:structure",
                "table" if visual != "native_flow" else "diagram",
                "visual",
                "content",
                Rect(x, y + 0.03, w, h * 0.90),
            )
        )
    else:
        objects.append(
            _object(
                f"{slide_id}:message",
                "text",
                "body",
                "content",
                Rect(x + 0.35, y + 0.35, w - 0.70, h - 0.70),
                text=message,
                font_size_pt=25,
                bold=True,
            )
        )
    return objects


def validate_planned_geometry(
    objects: list[Mapping[str, Any]],
    contract: LayoutContract,
) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    rects: dict[str, Rect] = {}
    for item in objects:
        bounds = item.get("bounds", {})
        rect = Rect(
            float(bounds.get("x", 0)),
            float(bounds.get("y", 0)),
            float(bounds.get("w", 0)),
            float(bounds.get("h", 0)),
        )
        object_id = str(item.get("object_id", "UNKNOWN"))
        rects[object_id] = rect
        zone_name = str(item.get("zone", "content"))
        zone = contract.zones.get(zone_name)
        if zone is None:
            issues.append(
                {"code": "UNKNOWN_ZONE", "object_id": object_id, "zone": zone_name}
            )
        elif not _contains(zone, rect):
            issues.append(
                {
                    "code": "ZONE_VIOLATION",
                    "object_id": object_id,
                    "zone": zone_name,
                    "bounds": rect.to_dict(),
                }
            )
        if (
            rect.x < 0
            or rect.y < 0
            or rect.right > contract.width
            or rect.bottom > contract.height
        ):
            issues.append(
                {
                    "code": "OFF_SLIDE",
                    "object_id": object_id,
                    "bounds": rect.to_dict(),
                }
            )
        if str(item.get("role")) not in {"footer", "page_number"}:
            if _intersection(rect, contract.zones["footer_exclusion"]):
                issues.append(
                    {
                        "code": "FOOTER_INTRUSION",
                        "object_id": object_id,
                        "bounds": rect.to_dict(),
                    }
                )
        if item.get("kind") == "text" and str(item.get("text", "")):
            role = str(item.get("role", "body"))
            preferred = float(item.get("font_size_pt", 18) or 18)
            metrics = choose_font_size(
                str(item.get("text", "")),
                role=role,
                preferred_font_size_pt=preferred,
                box_width_pt=rect.w * 72,
                box_height_pt=rect.h * 72,
                bold=bool(item.get("bold", False)),
                maximum_light_shrink_percent=float(
                    contract.text_policy.get("maximum_light_shrink_percent", 8)
                ),
            )
            if metrics.overflow:
                issues.append(
                    {
                        "code": "ESTIMATED_TEXT_OVERFLOW",
                        "object_id": object_id,
                        "estimated_height_pt": round(metrics.estimated_height_pt, 2),
                        "box_height_pt": round(rect.h * 72, 2),
                        "minimum_font_pt": metrics.font_size_pt,
                    }
                )
            minimum = float(contract.minimum_font_pt.get(role, 18))
            if metrics.font_size_pt < minimum:
                issues.append(
                    {
                        "code": "FONT_BELOW_MINIMUM",
                        "object_id": object_id,
                        "font_size_pt": metrics.font_size_pt,
                        "minimum_font_pt": minimum,
                    }
                )

    text_items = [
        item
        for item in objects
        if item.get("kind") == "text"
        and str(item.get("role")) not in {"footer", "page_number"}
    ]
    for index, first in enumerate(text_items):
        first_id = str(first.get("object_id"))
        allowed = set(first.get("allowed_overlap_with", []))
        for second in text_items[index + 1 :]:
            second_id = str(second.get("object_id"))
            if second_id in allowed or first_id in set(
                second.get("allowed_overlap_with", [])
            ):
                continue
            intersection = _intersection(rects[first_id], rects[second_id])
            if intersection:
                issues.append(
                    {
                        "code": "UNDECLARED_TEXT_OVERLAP",
                        "object_a": first_id,
                        "object_b": second_id,
                        "intersection": intersection.to_dict(),
                    }
                )
    return issues


__all__ = [
    "LayoutContract",
    "Rect",
    "plan_slide_geometry",
    "validate_planned_geometry",
]
