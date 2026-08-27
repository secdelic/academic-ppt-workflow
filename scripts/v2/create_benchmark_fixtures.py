#!/usr/bin/env python3
"""Create deterministic, local-only synthetic fixtures for the v2 benchmark.

The generated material is deliberately fictional. It contains aggregate
records only, uses fixture-only reference tokens, and must never be described
as real-world validation.
"""

from __future__ import annotations

import argparse
import csv
import gc
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterable, Sequence

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches as PptxInches
from pptx.util import Pt as PptxPt


FIXTURE_VERSION = "2.0.0"
FIXED_TIMESTAMP = "2026-07-28T00:00:00Z"
FIXED_ZIP_DATETIME = (2026, 7, 28, 0, 0, 0)
SYNTHETIC_MARKER = "SYNTHETIC_FIXTURE_ONLY"
REAL_WORLD_STATUS = "REAL_WORLD_VALIDATION_PENDING"
PRIVATE_SENTINEL = "PRIVATE_REFERENCE_SENTINEL_DO_NOT_COPY_7F3E91B2"
PRIVATE_SENTINEL_SHA256 = hashlib.sha256(PRIVATE_SENTINEL.encode("utf-8")).hexdigest()

USABLE_WIDTH_DXA = 9360
TABLE_INDENT_DXA = 120
CELL_MARGIN_DXA = {"top": 80, "start": 120, "bottom": 80, "end": 120}
INK = RGBColor(0x0B, 0x25, 0x45)
BLUE = RGBColor(0x2E, 0x74, 0xB5)
DARK_BLUE = RGBColor(0x1F, 0x4D, 0x78)
MUTED = RGBColor(0x5A, 0x67, 0x75)
TEAL = RGBColor(0x1E, 0xA6, 0xA8)
LIGHT_BLUE = "E8EEF5"
LIGHT_NOTE = "F4F6F9"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def _write_csv(path: Path, fieldnames: Sequence[str], rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def _set_cell_shading(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_margins(cell: Any) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, width in CELL_MARGIN_DXA.items():
        tag = f"w:{edge}"
        node = tc_mar.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            tc_mar.append(node)
        node.set(qn("w:w"), str(width))
        node.set(qn("w:type"), "dxa")


def _set_table_geometry(table: Any, widths_dxa: Sequence[int]) -> None:
    if sum(widths_dxa) != USABLE_WIDTH_DXA:
        raise ValueError(f"Table widths must total {USABLE_WIDTH_DXA} DXA: {widths_dxa}")
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(USABLE_WIDTH_DXA))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(TABLE_INDENT_DXA))
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_layout = tbl_pr.find(qn("w:tblLayout"))
    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)

    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            width = int(widths_dxa[idx])
            cell.width = Inches(width / 1440)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            _set_cell_margins(cell)


def _set_run_font(
    run: Any,
    *,
    name: str = "Calibri",
    size: float | None = None,
    color: RGBColor | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def _set_style_font(style: Any, name: str, size: float, color: RGBColor) -> None:
    style.font.name = name
    style._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    style._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    style.font.size = Pt(size)
    style.font.color.rgb = color


def _configure_document(doc: Document, running_label: str) -> None:
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.right_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    _set_style_font(normal, "Calibri", 11, INK)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    title = styles["Title"]
    _set_style_font(title, "Calibri", 28, INK)
    title.font.bold = True
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(8)
    title.paragraph_format.keep_with_next = True

    subtitle = styles["Subtitle"]
    _set_style_font(subtitle, "Calibri", 13, MUTED)
    subtitle.paragraph_format.space_before = Pt(0)
    subtitle.paragraph_format.space_after = Pt(16)
    subtitle.paragraph_format.keep_with_next = True

    heading_tokens = {
        "Heading 1": (16, BLUE, 18, 10),
        "Heading 2": (13, BLUE, 14, 7),
        "Heading 3": (12, DARK_BLUE, 10, 5),
    }
    for style_name, (size, color, before, after) in heading_tokens.items():
        style = styles[style_name]
        _set_style_font(style, "Calibri", size, color)
        style.font.bold = True
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    if "Synthetic Kicker" not in [s.name for s in styles]:
        kicker = styles.add_style("Synthetic Kicker", 1)
    else:
        kicker = styles["Synthetic Kicker"]
    _set_style_font(kicker, "Calibri", 9, TEAL)
    kicker.font.bold = True
    kicker.paragraph_format.space_after = Pt(5)
    kicker.paragraph_format.keep_with_next = True

    if "Synthetic Caption" not in [s.name for s in styles]:
        caption = styles.add_style("Synthetic Caption", 1)
    else:
        caption = styles["Synthetic Caption"]
    _set_style_font(caption, "Calibri", 9, MUTED)
    caption.font.italic = True
    caption.paragraph_format.space_before = Pt(4)
    caption.paragraph_format.space_after = Pt(8)
    caption.paragraph_format.keep_with_next = True

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(0)
    run = hp.add_run(f"{running_label}  |  {SYNTHETIC_MARKER}")
    _set_run_font(run, size=8.5, color=MUTED, bold=True)

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    run = fp.add_run("Local benchmark fixture - not clinical evidence  |  Page ")
    _set_run_font(run, size=8, color=MUTED)
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = " PAGE "
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char_begin)
    run._r.append(instr_text)
    run._r.append(fld_char_end)

    props = doc.core_properties
    props.title = running_label
    props.subject = "Deterministic synthetic benchmark input"
    props.author = "Academic PPT Workflow Fixture Generator"
    props.keywords = f"{SYNTHETIC_MARKER}; aggregate-only; local-only"
    props.comments = f"{REAL_WORLD_STATUS}; no patient-level data"
    props.last_modified_by = "Academic PPT Workflow Fixture Generator"


def _add_title_block(doc: Document, kicker: str, title: str, subtitle: str, case_id: str) -> None:
    doc.add_paragraph(kicker.upper(), style="Synthetic Kicker")
    doc.add_paragraph(title, style="Title")
    doc.add_paragraph(subtitle, style="Subtitle")
    metadata = doc.add_table(rows=3, cols=2)
    metadata.style = "Table Grid"
    _set_table_geometry(metadata, [2700, 6660])
    rows = [
        ("Fixture", case_id),
        ("Evidence status", f"{SYNTHETIC_MARKER}; {REAL_WORLD_STATUS}"),
        ("Privacy", "Aggregate-level records only; no patient-level or institutional data"),
    ]
    for row, (label, value) in zip(metadata.rows, rows):
        row.cells[0].text = label
        row.cells[1].text = value
        _set_cell_shading(row.cells[0], LIGHT_BLUE)
        row.cells[0].paragraphs[0].runs[0].bold = True
    doc.add_paragraph(
        "Scientific-use boundary: every number and conclusion below exists only to test "
        "the workflow. It is not a real study, publication, protocol, or clinical claim.",
        style="Synthetic Caption",
    )


def _add_table(
    doc: Document,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    widths_dxa: Sequence[int],
    caption: str | None = None,
) -> Any:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    _set_table_geometry(table, widths_dxa)
    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = header
        _set_cell_shading(cell, LIGHT_BLUE)
        for run in cell.paragraphs[0].runs:
            _set_run_font(run, size=9.5, color=INK, bold=True)
    for values in rows:
        row = table.add_row()
        for idx, value in enumerate(values):
            row.cells[idx].text = str(value)
            for run in row.cells[idx].paragraphs[0].runs:
                _set_run_font(run, size=9.5, color=INK)
    _set_table_geometry(table, widths_dxa)
    if caption:
        doc.add_paragraph(caption, style="Synthetic Caption")
    return table


def _add_figure(doc: Document, image_path: Path, caption: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.keep_with_next = True
    paragraph.add_run().add_picture(str(image_path), width=Inches(6.2))
    doc.add_paragraph(caption, style="Synthetic Caption")


def _normalize_office_zip_bytes(payload: bytes) -> bytes:
    """Normalize an Office ZIP payload, including embedded workbooks."""
    with zipfile.ZipFile(io.BytesIO(payload), "r") as source:
        entries: list[tuple[zipfile.ZipInfo, bytes]] = []
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename.startswith("ppt/embeddings/") and info.filename.lower().endswith(
                (".xlsx", ".xlsm")
            ):
                data = _normalize_office_zip_bytes(data)
            if info.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                text = re.sub(
                    r"(<dcterms:created[^>]*>).*?(</dcterms:created>)",
                    rf"\g<1>{FIXED_TIMESTAMP}\g<2>",
                    text,
                )
                text = re.sub(
                    r"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)",
                    rf"\g<1>{FIXED_TIMESTAMP}\g<2>",
                    text,
                )
                data = text.encode("utf-8")
            entries.append((info, data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for old_info, data in sorted(entries, key=lambda item: item[0].filename):
            info = zipfile.ZipInfo(old_info.filename, FIXED_ZIP_DATETIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 0
            info.external_attr = old_info.external_attr
            info.flag_bits = old_info.flag_bits
            info.comment = old_info.comment
            target.writestr(info, data)
    return buffer.getvalue()


def _normalize_office_package(path: Path) -> None:
    """Normalize ZIP metadata and core timestamps for byte-stable fixtures."""
    # Some Windows Office filter drivers deny atomic replacement while allowing
    # an in-place rewrite of the generator-owned file.
    path.write_bytes(_normalize_office_zip_bytes(path.read_bytes()))


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    windows_dir = os.environ.get("WINDIR")
    candidates = []
    if windows_dir:
        font_dir = Path(windows_dir) / "Fonts"
        candidates.extend(
            [
                font_dir / ("arialbd.ttf" if bold else "arial.ttf"),
                font_dir / ("calibrib.ttf" if bold else "calibri.ttf"),
            ]
        )
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _draw_title(draw: ImageDraw.ImageDraw, title: str, subtitle: str) -> None:
    draw.text((80, 58), title, font=_font(48, True), fill="#0B2545")
    draw.text((80, 120), subtitle, font=_font(24), fill="#5A6775")
    draw.rounded_rectangle((80, 168, 1520, 174), radius=3, fill="#1EA6A8")


def _save_png(image: Image.Image, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path, format="PNG", optimize=False, compress_level=9)


def _case1_flow_image(path: Path) -> None:
    image = Image.new("RGB", (1600, 900), "#F7FAFC")
    draw = ImageDraw.Draw(image)
    _draw_title(draw, "Synthetic trial flow", "All counts are fixture-only aggregate values")
    boxes = [
        (520, 220, 1080, 315, "Screened\nn = 310", "#DCE9F4"),
        (520, 370, 1080, 465, "Randomized\nn = 240", "#CFECEC"),
        (120, 565, 720, 675, "Protocolized pathway\nn = 120", "#E8EEF5"),
        (880, 565, 1480, 675, "Comparator pathway\nn = 120", "#E8EEF5"),
        (120, 735, 720, 835, "Primary analysis\nn = 120", "#FFFFFF"),
        (880, 735, 1480, 835, "Primary analysis\nn = 120", "#FFFFFF"),
    ]
    for x1, y1, x2, y2, label, fill in boxes:
        draw.rounded_rectangle((x1, y1, x2, y2), radius=16, fill=fill, outline="#2E74B5", width=3)
        bbox = draw.multiline_textbbox((0, 0), label, font=_font(28, True), spacing=5, align="center")
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.multiline_text(((x1 + x2 - tw) / 2, (y1 + y2 - th) / 2), label, font=_font(28, True), fill="#0B2545", spacing=5, align="center")
    connectors = [
        ((800, 315), (800, 370)),
        ((800, 465), (420, 565)),
        ((800, 465), (1180, 565)),
        ((420, 675), (420, 735)),
        ((1180, 675), (1180, 735)),
    ]
    for start, end in connectors:
        draw.line((*start, *end), fill="#1EA6A8", width=5)
    draw.text((1100, 287), "Excluded n = 70", font=_font(22), fill="#7A5A00")
    _save_png(image, path)


def _case1_effect_image(path: Path) -> None:
    image = Image.new("RGB", (1600, 900), "#FFFFFF")
    draw = ImageDraw.Draw(image)
    _draw_title(draw, "Synthetic effect summary", "Risk ratio axis; square size is illustrative")
    chart_left, chart_right = 620, 1450
    null_x = 1035
    draw.line((chart_left, 250, chart_left, 790), fill="#A9B4C0", width=2)
    draw.line((chart_right, 250, chart_right, 790), fill="#A9B4C0", width=2)
    draw.line((null_x, 230, null_x, 805), fill="#9B1C1C", width=3)
    tick_values = [(0.5, chart_left), (1.0, null_x), (1.5, 1242), (2.0, chart_right)]
    for value, x in tick_values:
        draw.line((x, 790, x, 805), fill="#5A6775", width=2)
        draw.text((x - 22, 815), f"{value:.1f}", font=_font(20), fill="#5A6775")
    rows = [
        ("Day-7 immobility", 0.63, 0.42, 0.93, 345),
        ("Escalation of support", 0.78, 0.49, 1.24, 515),
        ("Protocol deviation", 0.91, 0.58, 1.42, 685),
    ]
    scale = (chart_right - chart_left) / 1.5
    for label, estimate, low, high, y in rows:
        draw.text((80, y - 18), label, font=_font(27, True), fill="#0B2545")
        x_est = chart_left + (estimate - 0.5) * scale
        x_low = chart_left + (low - 0.5) * scale
        x_high = chart_left + (high - 0.5) * scale
        draw.line((x_low, y, x_high, y), fill="#2E74B5", width=6)
        draw.line((x_low, y - 12, x_low, y + 12), fill="#2E74B5", width=3)
        draw.line((x_high, y - 12, x_high, y + 12), fill="#2E74B5", width=3)
        draw.rectangle((x_est - 11, y - 11, x_est + 11, y + 11), fill="#1EA6A8", outline="#0B2545")
        draw.text((1220, y - 18), f"{estimate:.2f} ({low:.2f}-{high:.2f})", font=_font(23), fill="#0B2545")
    draw.text((710, 205), "Favours protocolized", font=_font(20), fill="#1F4D78")
    draw.text((1110, 205), "Favours comparator", font=_font(20), fill="#1F4D78")
    _save_png(image, path)


def _case2_flow_image(path: Path) -> None:
    image = Image.new("RGB", (1600, 900), "#F8FAFC")
    draw = ImageDraw.Draw(image)
    _draw_title(draw, "Synthetic cohort architecture", "Planned counts, not observed recruitment")
    nodes = [
        (100, 300, 430, 465, "Screening frame\nn = 420", "#DCE9F4"),
        (635, 210, 965, 375, "Discovery cohort\nn = 200", "#CFECEC"),
        (635, 525, 965, 690, "Validation cohort\nn = 100", "#E8EEF5"),
        (1170, 210, 1500, 375, "Multi-omics subset\nn = 80", "#FFF4D6"),
        (1170, 525, 1500, 690, "Clinical endpoint set\nn = 300", "#F4F6F9"),
    ]
    connectors = [
        ((430, 382), (635, 292)),
        ((430, 382), (635, 607)),
        ((965, 292), (1170, 292)),
        ((965, 607), (1170, 607)),
    ]
    for start, end in connectors:
        draw.line((*start, *end), fill="#1EA6A8", width=6)
    for x1, y1, x2, y2, label, fill in nodes:
        draw.rounded_rectangle((x1, y1, x2, y2), radius=18, fill=fill, outline="#2E74B5", width=3)
        bbox = draw.multiline_textbbox((0, 0), label, font=_font(28, True), spacing=7, align="center")
        draw.multiline_text(
            ((x1 + x2 - (bbox[2] - bbox[0])) / 2, (y1 + y2 - (bbox[3] - bbox[1])) / 2),
            label,
            font=_font(28, True),
            fill="#0B2545",
            spacing=7,
            align="center",
        )
    draw.text((100, 770), "No patient-level rows are included in this benchmark fixture.", font=_font(25, True), fill="#9B1C1C")
    _save_png(image, path)


def _case2_timeline_image(path: Path) -> None:
    image = Image.new("RGB", (1600, 900), "#FFFFFF")
    draw = ImageDraw.Draw(image)
    _draw_title(draw, "Synthetic project timeline", "Illustrative milestones for workflow testing")
    stages = [
        ("Protocol lock", "M0", 170),
        ("Pilot QC", "M3", 425),
        ("Discovery freeze", "M8", 720),
        ("Validation freeze", "M14", 1030),
        ("Readout", "M18", 1360),
    ]
    y = 480
    draw.line((170, y, 1360, y), fill="#2E74B5", width=8)
    for idx, (label, month, x) in enumerate(stages):
        draw.ellipse((x - 24, y - 24, x + 24, y + 24), fill="#1EA6A8", outline="#0B2545", width=3)
        label_y = 330 if idx % 2 == 0 else 565
        draw.text((x - 70, label_y), label, font=_font(24, True), fill="#0B2545")
        draw.text((x - 28, label_y + 42), month, font=_font(22), fill="#5A6775")
        draw.line((x, y - 24 if idx % 2 == 0 else y + 24, x, label_y + 70 if idx % 2 == 0 else label_y - 12), fill="#A9B4C0", width=3)
    draw.rounded_rectangle((140, 730, 1460, 820), radius=12, fill="#F4F6F9", outline="#D0D8E0", width=2)
    draw.text((180, 757), "Stop gate: any unresolved cohort, exposure, outcome, or missing-data definition blocks promotion.", font=_font(24, True), fill="#7A5A00")
    _save_png(image, path)


def _case2_workflow_image(path: Path) -> None:
    image = Image.new("RGB", (1600, 900), "#F8FAFC")
    draw = ImageDraw.Draw(image)
    _draw_title(draw, "Synthetic multi-omics workflow", "All findings remain computational inference until independently validated")
    labels = [
        ("Clinical aggregates", 100, "#DCE9F4"),
        ("Transcript matrix", 395, "#CFECEC"),
        ("Protein panel", 690, "#FFF4D6"),
        ("Locked integration", 985, "#E8EEF5"),
        ("Validation report", 1280, "#F4F6F9"),
    ]
    for idx, (label, x, fill) in enumerate(labels):
        if idx < len(labels) - 1:
            draw.line((x + 220, 475, labels[idx + 1][1] - 20, 475), fill="#1EA6A8", width=7)
        draw.rounded_rectangle((x, 360, x + 220, 590), radius=20, fill=fill, outline="#2E74B5", width=3)
        words = label.split(" ")
        text = "\n".join(words)
        bbox = draw.multiline_textbbox((0, 0), text, font=_font(26, True), spacing=7, align="center")
        draw.multiline_text(
            (x + 110 - (bbox[2] - bbox[0]) / 2, 475 - (bbox[3] - bbox[1]) / 2),
            text,
            font=_font(26, True),
            fill="#0B2545",
            spacing=7,
            align="center",
        )
    draw.text((100, 710), "Mechanism status: hypothesis only", font=_font(26, True), fill="#9B1C1C")
    draw.text((100, 755), "No experimental or clinical efficacy conclusion is permitted.", font=_font(24), fill="#5A6775")
    _save_png(image, path)


def _create_case1(root: Path) -> None:
    case = root / "case1_paper_report"
    docs = case / "documents"
    figures = case / "figures"
    references = case / "references"
    data = case / "data"
    for folder in (docs, figures, references, data):
        folder.mkdir(parents=True, exist_ok=True)

    flow = figures / "figure_01_trial_flow.png"
    effects = figures / "figure_02_effect_summary.png"
    _case1_flow_image(flow)
    _case1_effect_image(effects)

    _write_csv(
        data / "aggregate_results.csv",
        [
            "record_scope",
            "outcome",
            "group",
            "n",
            "events",
            "risk_percent",
            "effect_measure",
            "estimate",
            "ci_low",
            "ci_high",
            "p_value",
            "synthetic_status",
        ],
        [
            {
                "record_scope": "aggregate_only",
                "outcome": "day_7_immobility",
                "group": "protocolized_pathway",
                "n": 120,
                "events": 30,
                "risk_percent": "25.0",
                "effect_measure": "risk_ratio",
                "estimate": "0.63",
                "ci_low": "0.42",
                "ci_high": "0.93",
                "p_value": "0.020",
                "synthetic_status": SYNTHETIC_MARKER,
            },
            {
                "record_scope": "aggregate_only",
                "outcome": "day_7_immobility",
                "group": "comparator_pathway",
                "n": 120,
                "events": 48,
                "risk_percent": "40.0",
                "effect_measure": "reference",
                "estimate": "",
                "ci_low": "",
                "ci_high": "",
                "p_value": "",
                "synthetic_status": SYNTHETIC_MARKER,
            },
            {
                "record_scope": "aggregate_only",
                "outcome": "support_escalation",
                "group": "combined_effect",
                "n": 240,
                "events": "",
                "risk_percent": "",
                "effect_measure": "risk_ratio",
                "estimate": "0.78",
                "ci_low": "0.49",
                "ci_high": "1.24",
                "p_value": "0.300",
                "synthetic_status": SYNTHETIC_MARKER,
            },
            {
                "record_scope": "aggregate_only",
                "outcome": "protocol_deviation",
                "group": "combined_effect",
                "n": 240,
                "events": "",
                "risk_percent": "",
                "effect_measure": "risk_ratio",
                "estimate": "0.91",
                "ci_low": "0.58",
                "ci_high": "1.42",
                "p_value": "0.670",
                "synthetic_status": SYNTHETIC_MARKER,
            },
        ],
    )
    references.joinpath("fixture_references.md").write_text(
        "# Fixture-only reference tokens\n\n"
        "These tokens are deliberately not bibliographic citations and do not identify real publications.\n\n"
        "- `SYNTHETIC_REF_CASE1_001`: fixture-only background rationale.\n"
        "- `SYNTHETIC_REF_CASE1_002`: fixture-only trial-reporting convention.\n"
        "- `SYNTHETIC_REF_CASE1_003`: fixture-only statistical-method note.\n",
        encoding="utf-8",
    )

    doc = Document()
    _configure_document(doc, "CASE 1 - Synthetic paper report")
    _add_title_block(
        doc,
        "Paper report benchmark",
        "A protocolized mobility pathway reduced a synthetic day-7 endpoint",
        "Fictional parallel-group trial report for content, number, citation, and narrative testing",
        "CASE_1_PAPER_REPORT",
    )
    doc.add_heading("The synthetic clinical question is explicitly bounded", level=1)
    doc.add_paragraph(
        "In this fictional trial, 240 aggregate participants were assigned 1:1 to a "
        "protocolized mobility pathway or a comparator pathway. The prespecified primary "
        "endpoint was day-7 immobility. All counts and effects are synthetic inputs, not "
        "observations about real people. [SYNTHETIC_REF_CASE1_001]"
    )
    doc.add_heading("Methods preserve estimand and denominator definitions", level=1)
    _add_table(
        doc,
        ["Field", "Synthetic specification"],
        [
            ("Design", "Parallel-group randomized fixture; allocation ratio 1:1"),
            ("Population", "Aggregate simulated cohort; n = 240"),
            ("Exposure", "Protocolized mobility pathway vs comparator pathway"),
            ("Primary outcome", "Day-7 immobility; binary endpoint"),
            ("Primary estimand", "Risk ratio using all 240 randomized aggregate records"),
            ("Uncertainty", "Two-sided 95% confidence interval; fixture p value"),
            ("Missing data", "No missing primary endpoint in the synthetic fixture"),
        ],
        [2700, 6660],
        "Table 1. Prespecified synthetic analysis contract. Source: this fixture document.",
    )
    doc.add_heading("The primary synthetic endpoint differs between pathways", level=1)
    _add_table(
        doc,
        ["Outcome", "Protocolized", "Comparator", "Effect (95% CI)", "Fixture p"],
        [
            ("Day-7 immobility", "30/120 (25.0%)", "48/120 (40.0%)", "RR 0.63 (0.42-0.93)", "0.020"),
            ("Support escalation", "Aggregate only", "Aggregate only", "RR 0.78 (0.49-1.24)", "0.300"),
            ("Protocol deviation", "Aggregate only", "Aggregate only", "RR 0.91 (0.58-1.42)", "0.670"),
        ],
        [2480, 1680, 1680, 2280, 1240],
        "Table 2. Fixture-only aggregate results. Exact values also appear in data/aggregate_results.csv.",
    )
    _add_figure(
        doc,
        flow,
        "Figure 1. Synthetic trial flow. Counts are aggregate fixture values; source: data/aggregate_results.csv and this document.",
    )
    _add_figure(
        doc,
        effects,
        "Figure 2. Synthetic effect summary. Confidence intervals and p values are fixture-only test data.",
    )
    doc.add_heading("Interpretation is limited to the synthetic randomized contrast", level=1)
    doc.add_paragraph(
        "Within the fixture, the protocolized pathway is associated with a lower risk of "
        "the primary endpoint. The secondary estimates are imprecise and do not support a "
        "claim of benefit. The deck generator must preserve these distinctions and must not "
        "generalize the fixture to clinical practice."
    )
    doc.add_heading("Limitations must remain visible", level=1)
    for item in [
        "The material is entirely synthetic and cannot establish clinical effectiveness.",
        "Only aggregate values are supplied; no patient-level reanalysis is possible.",
        "Fixture reference tokens are not publications and must not become formatted citations.",
        "The mechanism of any effect is not evaluated.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_heading("Fixture source register", level=1)
    doc.add_paragraph(
        "Claims C1-001 through C1-007 map to this DOCX, data/aggregate_results.csv, "
        "and the two PNG figures. Reference tokens map only to references/fixture_references.md."
    )
    docx_path = docs / "synthetic_trial_report.docx"
    doc.save(docx_path)
    del doc
    gc.collect()
    _normalize_office_package(docx_path)

    _write_json(
        case / "source_metadata.json",
        {
            "case_id": "CASE_1_PAPER_REPORT",
            "fixture_version": FIXTURE_VERSION,
            "created_at": FIXED_TIMESTAMP,
            "synthetic_status": SYNTHETIC_MARKER,
            "real_world_status": REAL_WORLD_STATUS,
            "patient_level_data": False,
            "network_used": False,
            "study_identity": "Fictional; no real trial, authors, institutions, or participants",
            "reference_policy": "Tokens are fixture-only and are not publications",
            "claims": [
                {"claim_id": "C1-001", "text": "240 aggregate participants were randomized 1:1.", "sources": ["documents/synthetic_trial_report.docx"]},
                {"claim_id": "C1-002", "text": "Primary aggregate risks were 25.0% and 40.0%.", "sources": ["data/aggregate_results.csv"]},
                {"claim_id": "C1-003", "text": "Fixture RR 0.63, 95% CI 0.42-0.93, p=0.020.", "sources": ["data/aggregate_results.csv"]},
                {"claim_id": "C1-004", "text": "Secondary estimates are imprecise.", "sources": ["data/aggregate_results.csv", "documents/synthetic_trial_report.docx"]},
                {"claim_id": "C1-005", "text": "No mechanism was evaluated.", "sources": ["documents/synthetic_trial_report.docx"]},
            ],
        },
    )


def _create_case2(root: Path) -> None:
    case = root / "case2_defense_project"
    docs = case / "documents"
    figures = case / "figures"
    data = case / "data"
    for folder in (docs, figures, data):
        folder.mkdir(parents=True, exist_ok=True)

    cohort = figures / "figure_01_cohort_architecture.png"
    timeline = figures / "figure_02_project_timeline.png"
    workflow = figures / "figure_03_multiomics_workflow.png"
    _case2_flow_image(cohort)
    _case2_timeline_image(timeline)
    _case2_workflow_image(workflow)

    _write_csv(
        data / "aggregate_project_metrics.csv",
        ["record_scope", "domain", "metric", "value", "unit", "analysis_role", "synthetic_status"],
        [
            {"record_scope": "aggregate_only", "domain": "recruitment", "metric": "screening_frame", "value": 420, "unit": "records", "analysis_role": "planned", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "domain": "recruitment", "metric": "target_total", "value": 300, "unit": "participants", "analysis_role": "planned", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "domain": "cohort", "metric": "discovery_target", "value": 200, "unit": "participants", "analysis_role": "planned", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "domain": "cohort", "metric": "validation_target", "value": 100, "unit": "participants", "analysis_role": "planned", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "domain": "omics", "metric": "multiomics_subset", "value": 80, "unit": "participants", "analysis_role": "planned", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "domain": "quality", "metric": "synthetic_pilot_complete_cases", "value": 46, "unit": "of 50", "analysis_role": "pilot_fixture", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "domain": "quality", "metric": "synthetic_pilot_missingness", "value": 8, "unit": "percent", "analysis_role": "pilot_fixture", "synthetic_status": SYNTHETIC_MARKER},
        ],
    )
    case.joinpath("protocol.md").write_text(
        "# Synthetic translational cohort protocol\n\n"
        f"Status: `{SYNTHETIC_MARKER}` / `{REAL_WORLD_STATUS}`\n\n"
        "## Cohort definition\n\n"
        "Fictional adults meeting a synthetic acute-illness screening rule within 24 hours "
        "of a hypothetical index time. No real eligibility logic or database fields are used.\n\n"
        "## Exposure definition\n\n"
        "A pre-index composite biomarker score computed only from synthetic baseline inputs.\n\n"
        "## Outcome definition\n\n"
        "A synthetic 28-day organ-support-free-day endpoint. The protocol forbids using "
        "post-outcome variables as baseline covariates.\n\n"
        "## Analysis plan\n\n"
        "Discovery and validation cohorts are locked before modeling. Missingness is summarized "
        "before imputation; model reference groups are declared explicitly; computational "
        "multi-omics findings remain computational inference.\n\n"
        "## Stop gates\n\n"
        "Cohort, exposure, outcome, time-window, missing-data, or validation definitions that "
        "remain unresolved block scientific promotion.\n",
        encoding="utf-8",
    )

    doc = Document()
    _configure_document(doc, "CASE 2 - Synthetic defense project")
    _add_title_block(
        doc,
        "Thesis and project benchmark",
        "A locked validation design separates discovery from confirmation",
        "Fictional translational cohort proposal for cross-file integration and methods storytelling",
        "CASE_2_DEFENSE_PROJECT",
    )
    doc.add_heading("The project asks a validation-focused scientific question", level=1)
    doc.add_paragraph(
        "The fictional project tests whether a pre-index composite biomarker score provides "
        "incremental prognostic information for a synthetic 28-day endpoint. The workflow "
        "must retain the distinction between prediction and causal explanation."
    )
    doc.add_heading("Definitions are fixed before model development", level=1)
    _add_table(
        doc,
        ["Definition", "Locked synthetic specification"],
        [
            ("Cohort", "Synthetic acute-illness screen within 24 hours of index time"),
            ("Exposure", "Pre-index composite biomarker score"),
            ("Outcome", "Synthetic 28-day organ-support-free-day endpoint"),
            ("Baseline", "Variables measured no later than index time"),
            ("Discovery", "n = 200 planned aggregate target"),
            ("Validation", "n = 100 planned aggregate target"),
            ("Missing data", "Describe first; multiple imputation only after mechanism review"),
            ("Validation", "Locked model specification; no outcome-guided feature changes"),
        ],
        [2700, 6660],
        "Table 1. Definitions are fixture-only and not a deployable clinical protocol.",
    )
    _add_figure(
        doc,
        cohort,
        "Figure 1. Planned synthetic cohort architecture. Counts are planned aggregate targets.",
    )
    doc.add_heading("The analysis plan distinguishes development from validation", level=1)
    _add_table(
        doc,
        ["Stage", "Purpose", "Permitted output", "Blocked overstatement"],
        [
            ("Discovery", "Feature reduction and model development", "Candidate predictive signature", "Confirmed clinical utility"),
            ("Internal validation", "Optimism and calibration assessment", "Corrected performance estimate", "External validity"),
            ("Locked validation", "Independent synthetic test set", "Validation-set performance", "Causal mechanism"),
            ("Multi-omics", "Integrative computational inference", "Supported interpretation", "Experimental proof"),
        ],
        [1800, 2700, 2640, 2220],
        "Table 2. Scientific language guardrails for the synthetic project.",
    )
    _add_figure(
        doc,
        workflow,
        "Figure 2. Synthetic multi-omics workflow. Mechanism status is hypothesis only.",
    )
    _add_figure(
        doc,
        timeline,
        "Figure 3. Synthetic project timeline. Milestones are illustrative, not commitments.",
    )
    doc.add_heading("Aggregate pilot values test cross-file reconciliation", level=1)
    _add_table(
        doc,
        ["Metric", "Value", "Interpretation boundary"],
        [
            ("Synthetic pilot complete cases", "46 of 50", "Fixture data-quality check only"),
            ("Synthetic pilot missingness", "8%", "No imputation result is supplied"),
            ("Planned multi-omics subset", "n = 80", "Computational inference only"),
        ],
        [3600, 1800, 3960],
        "Table 3. Values reconcile to data/aggregate_project_metrics.csv.",
    )
    doc.add_heading("Risks and limitations remain first-class content", level=1)
    for item in [
        "No real recruitment, assay, outcome, or performance result is available.",
        "The validation strategy is a test contract, not evidence of successful validation.",
        "Selection bias, temporal leakage, missingness, and batch effects require later empirical checks.",
        "Any multi-omics mechanism statement must remain computational inference or hypothesis only.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    docx_path = docs / "synthetic_project_plan.docx"
    doc.save(docx_path)
    del doc
    gc.collect()
    _normalize_office_package(docx_path)

    _write_json(
        case / "source_metadata.json",
        {
            "case_id": "CASE_2_DEFENSE_PROJECT",
            "fixture_version": FIXTURE_VERSION,
            "created_at": FIXED_TIMESTAMP,
            "synthetic_status": SYNTHETIC_MARKER,
            "real_world_status": REAL_WORLD_STATUS,
            "patient_level_data": False,
            "network_used": False,
            "study_identity": "Fictional translational project",
            "claims": [
                {"claim_id": "C2-001", "text": "Planned total target is n=300.", "sources": ["data/aggregate_project_metrics.csv", "documents/synthetic_project_plan.docx"]},
                {"claim_id": "C2-002", "text": "Discovery and validation targets are n=200 and n=100.", "sources": ["data/aggregate_project_metrics.csv"]},
                {"claim_id": "C2-003", "text": "The planned multi-omics subset is n=80.", "sources": ["data/aggregate_project_metrics.csv"]},
                {"claim_id": "C2-004", "text": "The synthetic pilot has 46 of 50 complete cases.", "sources": ["data/aggregate_project_metrics.csv"]},
                {"claim_id": "C2-005", "text": "Multi-omics findings are computational inference.", "sources": ["protocol.md", "documents/synthetic_project_plan.docx"]},
            ],
        },
    )


def _reference_pptx_js(module_path: Path, output_path: Path) -> str:
    module_json = json.dumps(str(module_path))
    output_json = json.dumps(str(output_path))
    sentinel_json = json.dumps(PRIVATE_SENTINEL)
    return f"""import {{ createRequire }} from 'node:module';
const require = createRequire(import.meta.url);
const PptxGenJS = require({module_json});
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Academic PPT Workflow Fixture Generator';
pptx.company = 'Synthetic local benchmark';
pptx.subject = 'Style-only reference; private content must not be copied';
pptx.title = 'Synthetic Clinical Style Reference';
pptx.lang = 'en-US';
pptx.category = '{SYNTHETIC_MARKER}';
pptx.comments = '{REAL_WORLD_STATUS}';
pptx.theme = {{
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US'
}};
pptx.defineSlideMaster({{
  title: 'SYNTHETIC_TITLE',
  background: {{ color: '0D2538' }},
  objects: [
    {{ rect: {{ x:0, y:0, w:13.333, h:0.16, fill:{{color:'1EA6A8'}}, line:{{color:'1EA6A8'}} }} }},
    {{ text: {{ text:'SYNTHETIC STYLE REFERENCE', options:{{x:0.7,y:6.95,w:4.6,h:0.2,fontFace:'Aptos',fontSize:10,color:'8FC9CB',margin:0,bold:true}} }} }},
    {{ placeholder: {{ text:'Presentation title', options:{{name:'title',type:'title',x:0.8,y:1.45,w:11.7,h:1.4,fontFace:'Aptos Display',fontSize:52,bold:true,color:'FFFFFF',margin:0,breakLine:false}} }} }},
    {{ placeholder: {{ text:'Subtitle', options:{{name:'subtitle',type:'body',x:0.82,y:3.15,w:10.8,h:0.65,fontFace:'Aptos',fontSize:24,color:'C9D6E1',margin:0}} }} }}
  ],
  slideNumber: {{ x:12.2,y:6.9,w:0.4,h:0.2,fontFace:'Aptos',fontSize:10,color:'8FC9CB',align:'right',margin:0 }}
}});
pptx.defineSlideMaster({{
  title: 'SYNTHETIC_CONTENT',
  background: {{ color: 'F7FAFC' }},
  objects: [
    {{ rect: {{ x:0, y:0, w:13.333, h:0.12, fill:{{color:'1EA6A8'}}, line:{{color:'1EA6A8'}} }} }},
    {{ line: {{ x:0.72,y:1.17,w:11.9,h:0,line:{{color:'D4DEE7',width:1}} }} }},
    {{ text: {{ text:'SYNTHETIC / PRIVATE REFERENCE', options:{{x:0.72,y:6.94,w:4.3,h:0.18,fontFace:'Aptos',fontSize:9,color:'687888',margin:0}} }} }},
    {{ placeholder: {{ text:'Action title', options:{{name:'title',type:'title',x:0.72,y:0.38,w:11.7,h:0.62,fontFace:'Aptos Display',fontSize:34,bold:true,color:'0D2538',margin:0}} }} }},
    {{ placeholder: {{ text:'Body content', options:{{name:'body',type:'body',x:0.82,y:1.45,w:11.65,h:4.95,fontFace:'Aptos',fontSize:20,color:'243B53',margin:0.06,breakLine:false}} }} }}
  ],
  slideNumber: {{ x:12.2,y:6.9,w:0.4,h:0.2,fontFace:'Aptos',fontSize:9,color:'687888',align:'right',margin:0 }}
}});
pptx.defineSlideMaster({{
  title: 'SYNTHETIC_SECTION',
  background: {{ color: 'E9F2F5' }},
  objects: [
    {{ rect: {{ x:0.72,y:1.35,w:0.18,h:3.6,fill:{{color:'1EA6A8'}},line:{{color:'1EA6A8'}} }} }},
    {{ placeholder: {{ text:'Section title', options:{{name:'title',type:'title',x:1.25,y:2.0,w:10.8,h:1.2,fontFace:'Aptos Display',fontSize:46,bold:true,color:'0D2538',margin:0}} }} }},
    {{ placeholder: {{ text:'Section descriptor', options:{{name:'body',type:'body',x:1.28,y:3.45,w:9.6,h:0.75,fontFace:'Aptos',fontSize:24,color:'496A7A',margin:0}} }} }}
  ],
  slideNumber: {{ x:12.2,y:6.9,w:0.4,h:0.2,fontFace:'Aptos',fontSize:9,color:'687888',align:'right',margin:0 }}
}});
let slide = pptx.addSlide('SYNTHETIC_TITLE');
slide.addText('Synthetic clinical research style', {{ placeholder:'title' }});
slide.addText('Visual reference only - content is protected', {{ placeholder:'subtitle' }});
slide.addNotes(`{SYNTHETIC_MARKER}\\n[Sources]\\nSynthetic local fixture only.`);

slide = pptx.addSlide('SYNTHETIC_SECTION');
slide.addText('Evidence before interpretation', {{ placeholder:'title' }});
slide.addText('A restrained visual system for scientific reasoning', {{ placeholder:'body' }});
slide.addNotes(`{SYNTHETIC_MARKER}\\n[Sources]\\nSynthetic local fixture only.`);

slide = pptx.addSlide('SYNTHETIC_CONTENT');
slide.addText('The synthetic primary measure changed across periods', {{ placeholder:'title' }});
slide.addChart(pptx.ChartType.bar, [
  {{ name:'Synthetic rate', labels:['Baseline','Follow-up'], values:[18,12] }}
], {{
  x:0.82,y:1.55,w:7.15,h:4.7,
  catAxisLabelFontFace:'Aptos',catAxisLabelFontSize:16,
  valAxisLabelFontFace:'Aptos',valAxisLabelFontSize:14,
  chartColors:['1EA6A8'],showLegend:false,showTitle:false,showValue:true,
  showCatName:false,showValAxisTitle:false,showCatAxisTitle:false,
  showValue:true,showBorder:false,showGridLines:true,
  valGridLine:{{color:'D4DEE7',width:1}},showValue:true
}});
slide.addText('18%', {{x:8.55,y:2.0,w:2.1,h:0.65,fontFace:'Aptos Display',fontSize:36,bold:true,color:'0D2538',margin:0}});
slide.addText('Synthetic baseline', {{x:8.55,y:2.7,w:2.5,h:0.35,fontFace:'Aptos',fontSize:17,color:'687888',margin:0}});
slide.addText('12%', {{x:8.55,y:3.75,w:2.1,h:0.65,fontFace:'Aptos Display',fontSize:36,bold:true,color:'1EA6A8',margin:0}});
slide.addText('Synthetic follow-up', {{x:8.55,y:4.45,w:2.8,h:0.35,fontFace:'Aptos',fontSize:17,color:'687888',margin:0}});
slide.addNotes(`{SYNTHETIC_MARKER}\\n[Sources]\\nSynthetic values embedded in this protected reference.`);

slide = pptx.addSlide('SYNTHETIC_CONTENT');
slide.addText('Protected reference content must never cross the style boundary', {{ placeholder:'title' }});
slide.addText({sentinel_json}, {{x:0.85,y:1.65,w:11.4,h:0.62,fontFace:'Aptos',fontSize:22,bold:true,color:'9B1C1C',margin:0}});
slide.addText('The benchmark passes only when downstream style-only and template-fill outputs omit the sentinel while retaining compatible visual tokens.', {{
  x:0.85,y:2.75,w:10.7,h:1.3,fontFace:'Aptos',fontSize:24,color:'243B53',margin:0.04,breakLine:false
}});
slide.addText('Do not copy private text, notes, media, or values.', {{x:0.85,y:4.65,w:10.2,h:0.55,fontFace:'Aptos Display',fontSize:28,bold:true,color:'1EA6A8',margin:0}});
slide.addNotes(`${{{sentinel_json}}}\\n{SYNTHETIC_MARKER}\\n[Sources]\\nProtected local fixture content.`);
await pptx.writeFile({{ fileName: {output_json}, compression: true }});
"""


def _find_node() -> Path:
    env_path = os.environ.get("ACADEMIC_PPT_NODE")
    if env_path and Path(env_path).is_file():
        return Path(env_path)
    bundled = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "node"
        / "bin"
        / "node.exe"
    )
    if bundled.is_file():
        return bundled
    discovered = shutil.which("node")
    if discovered:
        return Path(discovered)
    raise RuntimeError("Node.js is required to create the editable reference PPTX.")


def _find_pptxgenjs_module(node_modules: Path | None = None) -> Path:
    candidates = []
    if node_modules:
        candidates.append(node_modules / "pptxgenjs")
    env_modules = os.environ.get("NODE_PATH")
    if env_modules:
        candidates.extend(Path(part) / "pptxgenjs" for part in env_modules.split(os.pathsep) if part)
    candidates.append(
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "node"
        / "node_modules"
        / "pptxgenjs"
    )
    for candidate in candidates:
        if candidate.joinpath("package.json").is_file():
            return candidate
    raise RuntimeError("Project-approved PptxGenJS dependency was not found; no install was attempted.")


def _create_reference_pptx(path: Path) -> None:
    """Create the protected synthetic style reference without a child runtime.

    The fixture deliberately uses only editable native PowerPoint objects.  It
    carries a private sentinel so style-only and template-fill privacy tests
    can prove that source content is not copied into generated decks.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    generated_path = path.with_name(f"{path.stem}.generated.pptx")
    presentation = Presentation()
    presentation.slide_width = PptxInches(13.333333)
    presentation.slide_height = PptxInches(7.5)

    ink = "0B2545"
    blue = "2E74B5"
    teal = "1EA6A8"
    pale = "E8EEF5"

    def add_title(slide: Any, title: str, subtitle: str = "") -> None:
        title_box = slide.shapes.add_textbox(
            PptxInches(0.72), PptxInches(0.55), PptxInches(11.9), PptxInches(0.78)
        )
        title_frame = title_box.text_frame
        title_frame.clear()
        title_paragraph = title_frame.paragraphs[0]
        title_paragraph.text = title
        title_paragraph.font.name = "Aptos Display"
        title_paragraph.font.size = PptxPt(30)
        title_paragraph.font.bold = True
        title_paragraph.font.color.rgb = PptxRGBColor.from_string(ink)
        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                PptxInches(0.75), PptxInches(1.42), PptxInches(11.7), PptxInches(0.55)
            )
            paragraph = subtitle_box.text_frame.paragraphs[0]
            paragraph.text = subtitle
            paragraph.font.name = "Aptos"
            paragraph.font.size = PptxPt(16)
            paragraph.font.color.rgb = PptxRGBColor.from_string(blue)

    def add_footer(slide: Any, page_number: int) -> None:
        footer = slide.shapes.add_textbox(
            PptxInches(0.75), PptxInches(7.05), PptxInches(11.8), PptxInches(0.22)
        )
        paragraph = footer.text_frame.paragraphs[0]
        paragraph.text = f"SYNTHETIC PROTECTED STYLE REFERENCE  |  {page_number:02d}"
        paragraph.alignment = PP_ALIGN.RIGHT
        paragraph.font.name = "Aptos"
        paragraph.font.size = PptxPt(9)
        paragraph.font.color.rgb = PptxRGBColor.from_string(blue)

    blank_layout = presentation.slide_layouts[6]

    slide = presentation.slides.add_slide(blank_layout)
    accent = slide.shapes.add_shape(
        1, PptxInches(0), PptxInches(0), PptxInches(0.24), PptxInches(7.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = PptxRGBColor.from_string(teal)
    accent.line.fill.background()
    add_title(slide, "Protected reference: editorial design system")
    sentinel_box = slide.shapes.add_textbox(
        PptxInches(0.78), PptxInches(2.25), PptxInches(11.6), PptxInches(1.1)
    )
    sentinel_paragraph = sentinel_box.text_frame.paragraphs[0]
    sentinel_paragraph.text = PRIVATE_SENTINEL
    sentinel_paragraph.font.name = "Aptos"
    sentinel_paragraph.font.size = PptxPt(18)
    sentinel_paragraph.font.color.rgb = PptxRGBColor.from_string(blue)
    add_footer(slide, 1)

    slide = presentation.slides.add_slide(blank_layout)
    add_title(slide, "One conclusion anchors each analytical page", "REFERENCE CONTENT — NOT EVIDENCE")
    for index, (heading, body) in enumerate(
        [
            ("Claim", "Use an action title with bounded wording."),
            ("Evidence", "Place one dominant editable exhibit."),
            ("Boundary", "Keep limitations and source notes visible."),
        ]
    ):
        left = 0.75 + index * 4.15
        card = slide.shapes.add_shape(
            5, PptxInches(left), PptxInches(2.15), PptxInches(3.65), PptxInches(2.55)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = PptxRGBColor.from_string(pale)
        card.line.color.rgb = PptxRGBColor.from_string(blue)
        box = slide.shapes.add_textbox(
            PptxInches(left + 0.25), PptxInches(2.45), PptxInches(3.15), PptxInches(1.65)
        )
        frame = box.text_frame
        frame.clear()
        p0 = frame.paragraphs[0]
        p0.text = heading
        p0.font.name = "Aptos Display"
        p0.font.size = PptxPt(22)
        p0.font.bold = True
        p0.font.color.rgb = PptxRGBColor.from_string(ink)
        p1 = frame.add_paragraph()
        p1.text = body
        p1.font.name = "Aptos"
        p1.font.size = PptxPt(15)
        p1.font.color.rgb = PptxRGBColor.from_string(ink)
    add_footer(slide, 2)

    slide = presentation.slides.add_slide(blank_layout)
    add_title(slide, "Editable charts use a restrained, color-safe palette", "REFERENCE VALUES — DO NOT COPY")
    chart_data = ChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("Reference series", (28, 42, 35))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        PptxInches(0.9),
        PptxInches(2.0),
        PptxInches(7.15),
        PptxInches(4.25),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.value_axis.has_major_gridlines = True
    series = chart.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = PptxRGBColor.from_string(teal)
    note = slide.shapes.add_textbox(
        PptxInches(8.45), PptxInches(2.1), PptxInches(3.75), PptxInches(3.6)
    )
    note_frame = note.text_frame
    note_frame.clear()
    for index, text in enumerate(
        ["Editable native chart", "Direct labels preferred", "Muted gridlines", "Short source note"]
    ):
        paragraph = note_frame.paragraphs[0] if index == 0 else note_frame.add_paragraph()
        paragraph.text = text
        paragraph.font.name = "Aptos"
        paragraph.font.size = PptxPt(18)
        paragraph.font.color.rgb = PptxRGBColor.from_string(ink)
    add_footer(slide, 3)

    slide = presentation.slides.add_slide(blank_layout)
    add_title(slide, "Layouts preserve rhythm across different page roles", "REFERENCE CONTENT — NOT EVIDENCE")
    roles = [("QUESTION", teal), ("METHOD", blue), ("RESULT", ink), ("LIMITATION", "7A4EAB")]
    for index, (role, color) in enumerate(roles):
        top = 1.75 + index * 1.15
        marker = slide.shapes.add_shape(
            1, PptxInches(0.9), PptxInches(top), PptxInches(0.18), PptxInches(0.7)
        )
        marker.fill.solid()
        marker.fill.fore_color.rgb = PptxRGBColor.from_string(color)
        marker.line.fill.background()
        label = slide.shapes.add_textbox(
            PptxInches(1.35), PptxInches(top), PptxInches(2.2), PptxInches(0.65)
        )
        paragraph = label.text_frame.paragraphs[0]
        paragraph.text = role
        paragraph.font.name = "Aptos"
        paragraph.font.size = PptxPt(18)
        paragraph.font.bold = True
        paragraph.font.color.rgb = PptxRGBColor.from_string(ink)
        guide = slide.shapes.add_shape(
            1, PptxInches(3.65), PptxInches(top), PptxInches(8.1), PptxInches(0.7)
        )
        guide.fill.solid()
        guide.fill.fore_color.rgb = PptxRGBColor.from_string(pale)
        guide.line.color.rgb = PptxRGBColor.from_string(blue)
    add_footer(slide, 4)

    presentation.save(generated_path)
    _normalize_office_package(generated_path)
    path.write_bytes(generated_path.read_bytes())
    generated_path.unlink()


def _create_case3(root: Path) -> None:
    case = root / "case3_reference_template"
    docs = case / "new_content" / "documents"
    data = case / "new_content" / "data"
    reference = case / "reference"
    for folder in (docs, data, reference):
        folder.mkdir(parents=True, exist_ok=True)

    reference_pptx = reference / "protected_style_reference.pptx"
    _create_reference_pptx(reference_pptx)

    _write_csv(
        data / "aggregate_period_results.csv",
        ["record_scope", "period", "n", "synthetic_event_count", "synthetic_event_percent", "synthetic_status"],
        [
            {"record_scope": "aggregate_only", "period": "baseline", "n": 200, "synthetic_event_count": 36, "synthetic_event_percent": "18.0", "synthetic_status": SYNTHETIC_MARKER},
            {"record_scope": "aggregate_only", "period": "follow_up", "n": 200, "synthetic_event_count": 24, "synthetic_event_percent": "12.0", "synthetic_status": SYNTHETIC_MARKER},
        ],
    )
    doc = Document()
    _configure_document(doc, "CASE 3 - Synthetic template inheritance")
    _add_title_block(
        doc,
        "Reference inheritance benchmark",
        "A synthetic event rate declined after a fictional implementation period",
        "New content must adopt style without copying protected reference content",
        "CASE_3_REFERENCE_TEMPLATE",
    )
    doc.add_heading("New content is scientifically separate from the reference deck", level=1)
    doc.add_paragraph(
        "The reference PPTX is supplied only to test master, layout, typography, color, "
        "grid, and placeholder extraction. Its private slide text, notes, chart values, "
        "and media are not evidence for this new synthetic study."
    )
    doc.add_heading("Aggregate synthetic results are fully traceable", level=1)
    _add_table(
        doc,
        ["Period", "Aggregate n", "Synthetic events", "Synthetic event rate"],
        [
            ("Baseline", "200", "36", "18.0%"),
            ("Follow-up", "200", "24", "12.0%"),
        ],
        [2340, 1920, 2460, 2640],
        "Table 1. Fixture-only aggregate period results; source: data/aggregate_period_results.csv.",
    )
    doc.add_heading("Interpretation remains non-causal", level=1)
    doc.add_paragraph(
        "The before-after contrast is descriptive. The workflow must not describe the "
        "difference as an intervention effect because no concurrent comparator, adjustment, "
        "or causal design is supplied."
    )
    doc.add_heading("Template inheritance acceptance checks", level=1)
    for item in [
        "Use reference dimensions, theme fonts, colors, margins, and page-role patterns.",
        "Use native placeholders when the template-fill route is selected.",
        "Do not copy reference text, notes, chart values, media, or protected content.",
        "Keep all new claims mapped only to this DOCX and aggregate CSV.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    docx_path = docs / "synthetic_new_content.docx"
    doc.save(docx_path)
    del doc
    gc.collect()
    _normalize_office_package(docx_path)

    _write_json(
        case / "source_metadata.json",
        {
            "case_id": "CASE_3_REFERENCE_TEMPLATE",
            "fixture_version": FIXTURE_VERSION,
            "created_at": FIXED_TIMESTAMP,
            "synthetic_status": SYNTHETIC_MARKER,
            "real_world_status": REAL_WORLD_STATUS,
            "patient_level_data": False,
            "network_used": False,
            "reference_mode_default": "style-only",
            "protected_reference": "reference/protected_style_reference.pptx",
            "private_sentinel_sha256": PRIVATE_SENTINEL_SHA256,
            "sentinel_allowed_locations": ["reference/protected_style_reference.pptx"],
            "sentinel_forbidden_output_routes": ["style-only", "template-fill"],
            "claims": [
                {"claim_id": "C3-001", "text": "Baseline synthetic event rate is 18.0%.", "sources": ["new_content/data/aggregate_period_results.csv"]},
                {"claim_id": "C3-002", "text": "Follow-up synthetic event rate is 12.0%.", "sources": ["new_content/data/aggregate_period_results.csv"]},
                {"claim_id": "C3-003", "text": "The contrast is descriptive and non-causal.", "sources": ["new_content/documents/synthetic_new_content.docx"]},
            ],
        },
    )


def _scan_sentinel(root: Path) -> dict[str, list[str]]:
    present: list[str] = []
    absent: list[str] = []
    sentinel_bytes = PRIVATE_SENTINEL.encode("utf-8")
    for path in sorted(
        p for p in root.rglob("*") if p.is_file() and "_build" not in p.relative_to(root).parts
    ):
        rel = path.relative_to(root).as_posix()
        found = False
        if zipfile.is_zipfile(path):
            with zipfile.ZipFile(path, "r") as archive:
                found = any(sentinel_bytes in archive.read(name) for name in archive.namelist())
        else:
            found = sentinel_bytes in path.read_bytes()
        (present if found else absent).append(rel)
    return {"present": present, "absent": absent}


def _verify_docx_structure(path: Path) -> dict[str, Any]:
    required_parts = {"word/document.xml", "word/styles.xml", "word/settings.xml", "word/numbering.xml"}
    with zipfile.ZipFile(path, "r") as archive:
        names = set(archive.namelist())
        document_xml = archive.read("word/document.xml").decode("utf-8")
        styles_xml = archive.read("word/styles.xml").decode("utf-8")
        has_geometry = (
            'w:tblW w:w="9360" w:type="dxa"' in document_xml
            or 'w:tblW w:type="dxa" w:w="9360"' in document_xml
        )
        has_ind = 'w:tblInd w:w="120" w:type="dxa"' in document_xml or 'w:tblInd w:type="dxa" w:w="120"' in document_xml
        return {
            "path": path.as_posix(),
            "valid_zip": True,
            "required_parts_present": sorted(required_parts.intersection(names)),
            "all_required_parts_present": required_parts.issubset(names),
            "named_styles_present": all(token in styles_xml for token in ["Synthetic Kicker", "Synthetic Caption"]),
            "fixed_table_width_9360_present": has_geometry,
            "table_indent_120_present": has_ind,
            "synthetic_marker_present": SYNTHETIC_MARKER in document_xml,
        }


def _verify_pptx_structure(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path, "r") as archive:
        names = set(archive.namelist())
        slide_names = sorted(name for name in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", name))
        master_names = sorted(name for name in names if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", name))
        layout_names = sorted(name for name in names if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", name))
        combined = b"\n".join(archive.read(name) for name in slide_names + sorted(n for n in names if n.startswith("ppt/notesSlides/") and n.endswith(".xml")))
        return {
            "path": path.as_posix(),
            "valid_zip": True,
            "slide_count": len(slide_names),
            "master_count": len(master_names),
            "layout_count": len(layout_names),
            "private_sentinel_present": PRIVATE_SENTINEL.encode("utf-8") in combined,
            "editable_chart_present": any(name.startswith("ppt/charts/chart") and name.endswith(".xml") for name in names),
            "theme_present": "ppt/theme/theme1.xml" in names,
        }


def _find_render_docx() -> Path | None:
    candidate = (
        Path.home()
        / ".codex"
        / "plugins"
        / "cache"
        / "openai-primary-runtime"
        / "documents"
        / "26.727.11326"
        / "skills"
        / "documents"
        / "render_docx.py"
    )
    return candidate if candidate.is_file() else None


def _render_docx_files(root: Path, python_executable: Path, render_script: Path) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for docx in sorted(root.rglob("*.docx")):
        rel_stem = docx.relative_to(root).with_suffix("")
        output_dir = root / "_verification" / "docx_render" / rel_stem
        output_dir.mkdir(parents=True, exist_ok=True)
        command = [
            str(python_executable),
            str(render_script),
            str(docx),
            "--output_dir",
            str(output_dir),
            "--emit_pdf",
        ]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=180,
                check=False,
            )
            pages = sorted(output_dir.glob("page-*.png"))
            pdf = output_dir / f"{docx.stem}.pdf"
            results.append(
                {
                    "docx": docx.relative_to(root).as_posix(),
                    "attempted": True,
                    "exit_code": result.returncode,
                    "png_page_count": len(pages),
                    "pdf_exists": pdf.is_file() and pdf.stat().st_size > 0,
                    "render_passed": result.returncode == 0 and len(pages) > 0,
                    "stdout_tail": result.stdout[-800:],
                    "stderr_tail": result.stderr[-800:],
                }
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            results.append(
                {
                    "docx": docx.relative_to(root).as_posix(),
                    "attempted": True,
                    "render_passed": False,
                    "error": str(exc),
                }
            )
    return results


def _write_fixture_manifest(root: Path, verification: dict[str, Any]) -> None:
    role_by_suffix = {
        ".docx": "scientific_document",
        ".csv": "aggregate_data",
        ".png": "synthetic_figure",
        ".pptx": "protected_style_reference",
        ".md": "protocol_or_reference",
        ".json": "audit_metadata",
        ".pdf": "render_qa_intermediate",
    }
    files = []
    for path in sorted(
        p
        for p in root.rglob("*")
        if p.is_file()
        and p.name != "fixture_manifest.json"
        and "_build" not in p.relative_to(root).parts
    ):
        rel = path.relative_to(root).as_posix()
        files.append(
            {
                "relative_path": rel,
                "sha256": _sha256(path),
                "bytes": path.stat().st_size,
                "role": role_by_suffix.get(path.suffix.lower(), "generated_fixture"),
                "synthetic_status": SYNTHETIC_MARKER,
                "contains_private_sentinel": rel == "case3_reference_template/reference/protected_style_reference.pptx",
            }
        )
    _write_json(
        root / "fixture_manifest.json",
        {
            "fixture_version": FIXTURE_VERSION,
            "created_at": FIXED_TIMESTAMP,
            "generator": "scripts/v2/create_benchmark_fixtures.py",
            "synthetic_status": SYNTHETIC_MARKER,
            "real_world_status": REAL_WORLD_STATUS,
            "patient_level_data": False,
            "network_used": False,
            "external_install_performed": False,
            "private_sentinel_sha256": PRIVATE_SENTINEL_SHA256,
            "files": files,
            "verification": verification,
        },
    )


def create_all_fixtures(
    output_root: Path,
    *,
    node_path: Path | None = None,
    node_modules: Path | None = None,
    verify_docx_render: bool = False,
    python_executable: Path | None = None,
    render_docx_script: Path | None = None,
) -> dict[str, Any]:
    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    _create_case1(output_root)
    _create_case2(output_root)
    _create_case3(output_root)

    docx_structure = [_verify_docx_structure(path) for path in sorted(output_root.rglob("*.docx"))]
    pptx_structure = _verify_pptx_structure(
        output_root / "case3_reference_template" / "reference" / "protected_style_reference.pptx"
    )
    sentinel_scan = _scan_sentinel(output_root)
    expected_sentinel_path = "case3_reference_template/reference/protected_style_reference.pptx"
    sentinel_passed = sentinel_scan["present"] == [expected_sentinel_path]

    render_results: list[dict[str, Any]] = []
    if verify_docx_render:
        render_script = render_docx_script or _find_render_docx()
        if render_script is None:
            render_results = [
                {
                    "attempted": False,
                    "render_passed": False,
                    "reason": "Packaged render_docx.py was not found.",
                }
            ]
        else:
            render_results = _render_docx_files(
                output_root,
                python_executable or Path(sys.executable),
                render_script,
            )

    verification = {
        "structural_verification": {
            "docx": docx_structure,
            "pptx": pptx_structure,
            "sentinel_scan": {
                "passed": sentinel_passed,
                "expected_only_location": expected_sentinel_path,
                "present_locations": sentinel_scan["present"],
            },
        },
        "docx_render_verification": {
            "requested": verify_docx_render,
            "results": render_results,
            "visual_inspection_required": verify_docx_render,
        },
    }
    _write_json(output_root / "_verification" / "verification.json", verification)
    _write_fixture_manifest(output_root, verification)
    return verification


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-root",
        type=Path,
        default=Path("staging/v2_benchmark/fixtures"),
        help="Destination for generated fixtures (default: staging/v2_benchmark/fixtures).",
    )
    parser.add_argument("--node", type=Path, default=None, help="Explicit local Node.js executable.")
    parser.add_argument(
        "--node-modules",
        type=Path,
        default=None,
        help="Explicit project-approved Node module directory containing pptxgenjs.",
    )
    parser.add_argument(
        "--verify-docx-render",
        action="store_true",
        help="Render DOCX files with the packaged local renderer and record results.",
    )
    parser.add_argument("--python", type=Path, default=None, help="Python executable for render_docx.py.")
    parser.add_argument("--render-docx-script", type=Path, default=None, help="Explicit packaged render_docx.py.")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    verification = create_all_fixtures(
        args.output_root,
        node_path=args.node,
        node_modules=args.node_modules,
        verify_docx_render=args.verify_docx_render,
        python_executable=args.python,
        render_docx_script=args.render_docx_script,
    )
    structure = verification["structural_verification"]
    if not all(item["all_required_parts_present"] for item in structure["docx"]):
        print("Fixture generation failed DOCX structural verification.", file=sys.stderr)
        return 2
    if not structure["pptx"]["private_sentinel_present"]:
        print("Fixture generation failed: reference sentinel was not embedded.", file=sys.stderr)
        return 2
    if not structure["sentinel_scan"]["passed"]:
        print("Fixture generation failed: private sentinel escaped its allowed reference file.", file=sys.stderr)
        return 2
    print(
        json.dumps(
            {
                "status": "SYNTHETIC_FIXTURES_READY",
                "output_root": str(args.output_root.resolve()),
                "real_world_status": REAL_WORLD_STATUS,
                "docx_count": len(structure["docx"]),
                "reference_pptx_slides": structure["pptx"]["slide_count"],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
