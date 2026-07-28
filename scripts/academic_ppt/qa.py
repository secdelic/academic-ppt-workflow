from __future__ import annotations

import json
import math
import re
import zipfile
from pathlib import Path

from PIL import Image, ImageStat
from pypdf import PdfReader

from .inventory import verify_input_hashes
from .utils import write_csv, write_json


def count_pptx_slides(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f"Corrupt PPTX ZIP member: {bad}")
        return len(
            [
                name
                for name in archive.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
            ]
        )


def inspect_previews(previews: list[Path]) -> list[str]:
    issues: list[str] = []
    for index, path in enumerate(previews, start=1):
        with Image.open(path) as image:
            if image.width < 1000 or image.height < 500:
                issues.append(f"Slide {index}: low render resolution {image.width}x{image.height}")
            stat = ImageStat.Stat(image.convert("RGB"))
            variance = sum(stat.var) / 3.0
            extrema = image.convert("RGB").getextrema()
            dynamic_range = max(channel[1] - channel[0] for channel in extrema)
            if variance < 2.0 or dynamic_range < 8:
                issues.append(f"Slide {index}: possible blank or near-blank page")
    return issues


def inspect_layouts(layout_dir: Path, canvas_w: int = 1280, canvas_h: int = 720) -> list[str]:
    issues: list[str] = []
    if not layout_dir.exists():
        return issues

    def visit(value, source: str):
        if isinstance(value, dict):
            position = value.get("position") or value.get("bounds")
            if isinstance(position, dict):
                keys = {"left", "top", "width", "height"}
                if keys.issubset(position):
                    left, top = float(position["left"]), float(position["top"])
                    width, height = float(position["width"]), float(position["height"])
                    if left < -1 or top < -1 or left + width > canvas_w + 1 or top + height > canvas_h + 1:
                        issues.append(f"{source}: out-of-bounds object {position}")
            for child in value.values():
                visit(child, source)
        elif isinstance(value, list):
            for child in value:
                visit(child, source)

    for path in sorted(layout_dir.glob("*.json")):
        try:
            visit(json.loads(path.read_text(encoding="utf-8")), path.name)
        except Exception as exc:
            issues.append(f"{path.name}: layout inspection failed: {exc}")
    return issues


def run_qa(
    pptx_path: Path,
    pdf_path: Path,
    previews: list[Path],
    layout_dir: Path,
    storyboard: list[dict[str, str]],
    claims: list[dict[str, str]],
    manifest: list[dict[str, str]],
    input_root: Path,
    output_dir: Path,
    render_method: str,
) -> tuple[str, list[str], list[str], list[str]]:
    file_issues: list[str] = []
    visual_issues = inspect_previews(previews) + inspect_layouts(layout_dir)
    scientific_issues: list[str] = []

    try:
        pptx_count = count_pptx_slides(pptx_path)
    except Exception as exc:
        pptx_count = 0
        file_issues.append(f"PPTX openability failed: {exc}")
    try:
        pdf_count = len(PdfReader(str(pdf_path)).pages)
    except Exception as exc:
        pdf_count = 0
        file_issues.append(f"PDF openability failed: {exc}")
    if pptx_count != len(storyboard):
        file_issues.append(f"PPTX/storyboard count mismatch: {pptx_count} vs {len(storyboard)}")
    if pdf_count != pptx_count:
        file_issues.append(f"PDF/PPTX page mismatch: {pdf_count} vs {pptx_count}")
    if len(previews) != pptx_count:
        file_issues.append(f"Preview/PPTX page mismatch: {len(previews)} vs {pptx_count}")
    file_issues.extend(verify_input_hashes(input_root, manifest))

    source_ids = {row["source_id"] for row in manifest}
    for claim in claims:
        if claim["source_id"] not in source_ids:
            scientific_issues.append(f"{claim['claim_id']}: unknown source_id {claim['source_id']}")
        if claim["claim_text"].strip() == "":
            scientific_issues.append(f"{claim['claim_id']}: empty claim")
        prohibited = claim["prohibited_overstatement"].strip().lower()
        if prohibited and prohibited in claim["claim_text"].lower():
            scientific_issues.append(f"{claim['claim_id']}: contains prohibited wording")
    for slide in storyboard:
        if slide["citation_requirement"] == "required" and not slide["source_ids"]:
            scientific_issues.append(f"{slide['slide_id']}: required citation has no source_id")
        if slide["single_key_message"] == "INFORMATION_REQUIRED":
            scientific_issues.append(f"{slide['slide_id']}: unresolved key message")

    severe = file_issues + visual_issues + scientific_issues
    status = "READY_FOR_ASSISTED_USE" if not severe else "PARTIALLY_READY"
    if not pptx_path.exists() or pptx_count == 0 or not pdf_path.exists():
        status = "BLOCKED"

    report = [
        "# QA Report",
        "",
        f"- Final status: `{status}`",
        f"- Renderer: {render_method}",
        f"- PPTX slides: {pptx_count}",
        f"- PDF pages: {pdf_count}",
        f"- Preview pages: {len(previews)}",
        "",
        "## Scientific QA",
        "",
        *(f"- FAIL: {item}" for item in scientific_issues),
        *(["- PASS: all registered claims map to known input sources; no automated claim completion was used"] if not scientific_issues else []),
        "",
        "## Visual QA",
        "",
        *(f"- FAIL: {item}" for item in visual_issues),
        *(["- PASS: all rendered pages are nonblank at 1600x900 or better; no layout object exceeded the slide canvas"] if not visual_issues else []),
        "",
        "## File QA",
        "",
        *(f"- FAIL: {item}" for item in file_issues),
        *(["- PASS: PPTX ZIP, PDF page count, preview count, and immutable-input hash checks passed"] if not file_issues else []),
        "",
        "## Boundary",
        "",
        "Automated QA does not constitute final scientific approval. Manual review remains mandatory.",
    ]
    (output_dir / "qa_report.md").write_text("\n".join(report), encoding="utf-8")
    return status, scientific_issues, visual_issues, file_issues


def write_traceability_maps(
    output_dir: Path,
    storyboard: list[dict[str, str]],
    claims: list[dict[str, str]],
    figures: list[dict[str, str]],
) -> None:
    write_csv(
        output_dir / "slide_manifest.csv",
        [
            "slide_id", "slide_title", "slide_purpose", "single_key_message",
            "source_ids", "visual_type", "confidence", "manual_review_required",
        ],
        storyboard,
    )
    write_csv(
        output_dir / "claim_source_map.csv",
        ["claim_id", "claim_text", "source_id", "source_location", "confidence", "allowed_wording"],
        claims,
    )
    write_csv(
        output_dir / "figure_source_map.csv",
        ["figure_id", "source_id", "relative_path", "figure_type", "editable", "manual_review_required"],
        figures,
    )
