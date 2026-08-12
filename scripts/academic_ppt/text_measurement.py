from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Iterable


CJK_RE = re.compile(
    r"[\u2e80-\u2eff\u3000-\u303f\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"
)


@dataclass(frozen=True)
class TextMetrics:
    estimated_lines: int
    estimated_height_pt: float
    estimated_width_pt: float
    font_size_pt: float
    line_spacing_ratio: float
    overflow: bool


def resolve_font(
    preferred: str,
    fallbacks: Iterable[str],
    installed_fonts: Iterable[str],
) -> str:
    installed = {font.casefold(): font for font in installed_fonts}
    for candidate in [preferred, *fallbacks]:
        match = installed.get(str(candidate).casefold())
        if match:
            return match
    return str(preferred or next(iter(fallbacks), "Arial"))


def character_units(text: str, *, bold: bool = False) -> float:
    units = 0.0
    for character in text:
        if character == "\t":
            units += 2.0
        elif character.isspace():
            units += 0.33
        elif CJK_RE.fullmatch(character):
            units += 1.0
        elif character in "MW@#%&":
            units += 0.9
        elif character in "ilI.,:;'|!":
            units += 0.28
        else:
            units += 0.55
    return units * (1.04 if bold else 1.0)


def estimate_text(
    text: str,
    *,
    font_size_pt: float,
    box_width_pt: float,
    box_height_pt: float,
    bold: bool = False,
    line_spacing_ratio: float = 1.18,
    margin_left_pt: float = 4.0,
    margin_right_pt: float = 4.0,
    margin_top_pt: float = 2.0,
    margin_bottom_pt: float = 2.0,
    bullet_indent_pt: float = 0.0,
) -> TextMetrics:
    usable_width = max(
        1.0,
        box_width_pt - margin_left_pt - margin_right_pt - bullet_indent_pt,
    )
    average_glyph_width = font_size_pt * 0.94
    paragraphs = text.splitlines() or [""]
    estimated_lines = 0
    maximum_line_units = 0.0
    for paragraph in paragraphs:
        units = character_units(paragraph, bold=bold)
        maximum_line_units = max(maximum_line_units, units)
        estimated_lines += max(
            1,
            int(math.ceil(units * average_glyph_width / usable_width)),
        )
    line_height = font_size_pt * line_spacing_ratio
    estimated_height = (
        estimated_lines * line_height + margin_top_pt + margin_bottom_pt
    )
    estimated_width = min(
        usable_width,
        maximum_line_units * average_glyph_width,
    )
    return TextMetrics(
        estimated_lines=estimated_lines,
        estimated_height_pt=estimated_height,
        estimated_width_pt=estimated_width,
        font_size_pt=font_size_pt,
        line_spacing_ratio=line_spacing_ratio,
        overflow=estimated_height > box_height_pt + 0.5,
    )


def minimum_readable_font(role: str) -> float:
    return {
        "title": 28.0,
        "body": 18.0,
        "chart_label": 15.0,
        "footer": 10.0,
    }.get(role, 18.0)


def choose_font_size(
    text: str,
    *,
    role: str,
    preferred_font_size_pt: float,
    box_width_pt: float,
    box_height_pt: float,
    bold: bool = False,
    maximum_light_shrink_percent: float = 8.0,
) -> TextMetrics:
    minimum = max(
        minimum_readable_font(role),
        preferred_font_size_pt * (1.0 - maximum_light_shrink_percent / 100.0),
    )
    candidates = [preferred_font_size_pt]
    current = preferred_font_size_pt - 1.0
    while current >= minimum:
        candidates.append(current)
        current -= 1.0
    last = estimate_text(
        text,
        font_size_pt=candidates[-1],
        box_width_pt=box_width_pt,
        box_height_pt=box_height_pt,
        bold=bold,
    )
    for candidate in candidates:
        metrics = estimate_text(
            text,
            font_size_pt=candidate,
            box_width_pt=box_width_pt,
            box_height_pt=box_height_pt,
            bold=bold,
        )
        if not metrics.overflow:
            return metrics
        last = metrics
    return last


__all__ = [
    "TextMetrics",
    "character_units",
    "choose_font_size",
    "estimate_text",
    "minimum_readable_font",
    "resolve_font",
]
