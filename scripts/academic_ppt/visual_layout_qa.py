from __future__ import annotations

import json
import shutil
import subprocess
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from .utils import sha256_file


class VisualLayoutQAError(RuntimeError):
    """Raised when the read-only PowerPoint geometry audit cannot run safely."""


def inspect_powerpoint_text_layout(
    *,
    pptx_path: Path,
    output_json: Path,
    script_path: Path,
    timeout_seconds: int = 120,
    slide_indexes: list[int] | tuple[int, ...] | None = None,
    preview_dir: Path | None = None,
    output_pdf: Path | None = None,
) -> tuple[dict[str, Any], str]:
    pptx = pptx_path.resolve()
    output = output_json.resolve()
    script = script_path.resolve()
    if not pptx.is_file() or pptx.suffix.lower() != ".pptx":
        raise VisualLayoutQAError(f"Missing PPTX for layout QA: {pptx}")
    if not script.is_file() or script.suffix.lower() != ".ps1":
        raise VisualLayoutQAError(f"Missing layout QA script: {script}")
    if output.suffix.lower() != ".json":
        raise VisualLayoutQAError("Layout QA output must be JSON")
    if output.exists():
        raise VisualLayoutQAError(f"Refusing to overwrite layout QA output: {output}")
    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if not powershell:
        raise VisualLayoutQAError("PowerShell is unavailable for PowerPoint layout QA")
    before = {
        pptx: sha256_file(pptx),
        script: sha256_file(script),
    }
    command = [
        powershell,
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-InputPptx",
        str(pptx),
        "-OutputJson",
        str(output),
    ]
    if slide_indexes is not None:
        indexes = [int(value) for value in slide_indexes]
        if not indexes or any(value < 1 for value in indexes):
            raise VisualLayoutQAError(
                "slide_indexes must contain positive slide indexes"
            )
        if len(indexes) != len(set(indexes)):
            raise VisualLayoutQAError("slide_indexes contains duplicates")
        command.extend(["-SlideIndexes", ",".join(str(value) for value in indexes)])
    preview = preview_dir.resolve() if preview_dir is not None else None
    if preview is not None:
        if preview.exists() and any(preview.iterdir()):
            raise VisualLayoutQAError(
                f"Refusing to write previews into non-empty directory: {preview}"
            )
        command.extend(["-PreviewDir", str(preview)])
    pdf = output_pdf.resolve() if output_pdf is not None else None
    if pdf is not None:
        if pdf.suffix.lower() != ".pdf":
            raise VisualLayoutQAError("output_pdf must use the .pdf extension")
        if pdf.exists():
            raise VisualLayoutQAError(f"Refusing to overwrite PDF: {pdf}")
        command.extend(["-OutputPdf", str(pdf)])
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        raise VisualLayoutQAError(
            f"PowerPoint layout QA timed out after {timeout_seconds}s"
        ) from exc
    changed = [
        str(path)
        for path, digest in before.items()
        if not path.is_file() or sha256_file(path) != digest
    ]
    if changed:
        raise VisualLayoutQAError(
            "Layout QA modified a read-only input: " + ", ".join(changed)
        )
    log = (completed.stdout + "\n" + completed.stderr).strip()
    if completed.returncode != 0 or not output.is_file():
        raise VisualLayoutQAError(
            f"PowerPoint layout QA failed ({completed.returncode}): {log}"
        )
    try:
        report = json.loads(output.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise VisualLayoutQAError(f"Invalid layout QA JSON: {output}") from exc
    if not isinstance(report, dict):
        raise VisualLayoutQAError("Layout QA JSON must be a mapping")
    if report.get("source_pptx_sha256", "").lower() != before[pptx].lower():
        raise VisualLayoutQAError("Layout QA source hash does not match the PPTX")
    if slide_indexes is not None:
        returned = sorted(
            int(item.get("slide_index", 0))
            for item in report.get("slides", [])
            if isinstance(item, Mapping)
        )
        if returned != sorted(indexes):
            raise VisualLayoutQAError(
                "Layout QA returned a different slide scope than requested"
            )
    if preview is not None:
        preview_files = sorted(preview.glob("slide_*.png"))
        expected_preview_count = len(indexes) if slide_indexes is not None else int(
            report.get("slide_count", 0)
        )
        if len(preview_files) != expected_preview_count:
            raise VisualLayoutQAError(
                "PowerPoint preview count differs from requested geometry scope"
            )
    if pdf is not None and (not pdf.is_file() or pdf.stat().st_size == 0):
        raise VisualLayoutQAError("PowerPoint did not produce the requested PDF")
    return report, log


def _number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _intersection_ratio(a: Mapping[str, Any], b: Mapping[str, Any]) -> float:
    ax = _number(a.get("bound_left_pt")) or _number(a.get("left_pt"))
    ay = _number(a.get("bound_top_pt")) or _number(a.get("top_pt"))
    aw = _number(a.get("bound_width_pt")) or _number(a.get("width_pt"))
    ah = _number(a.get("bound_height_pt")) or _number(a.get("height_pt"))
    bx = _number(b.get("bound_left_pt")) or _number(b.get("left_pt"))
    by = _number(b.get("bound_top_pt")) or _number(b.get("top_pt"))
    bw = _number(b.get("bound_width_pt")) or _number(b.get("width_pt"))
    bh = _number(b.get("bound_height_pt")) or _number(b.get("height_pt"))
    if None in {ax, ay, aw, ah, bx, by, bw, bh}:
        return 0.0
    assert ax is not None and ay is not None and aw is not None and ah is not None
    assert bx is not None and by is not None and bw is not None and bh is not None
    if aw <= 0 or ah <= 0 or bw <= 0 or bh <= 0:
        return 0.0
    left = max(ax, bx)
    top = max(ay, by)
    right = min(ax + aw, bx + bw)
    bottom = min(ay + ah, by + bh)
    if right <= left or bottom <= top:
        return 0.0
    intersection = (right - left) * (bottom - top)
    return intersection / min(aw * ah, bw * bh)


def inspect_text_geometry(report: Mapping[str, Any]) -> list[str]:
    """Detect text overflow and competing text boxes without returning text."""

    issues: list[str] = []
    if report.get("status") != "PASS":
        return ["PowerPoint text-layout manifest did not report PASS"]
    slide_width = _number(report.get("slide_width_pt")) or 960.0
    slide_height = _number(report.get("slide_height_pt")) or 540.0
    footer_start = slide_height * 0.885
    for slide in report.get("slides", []):
        if not isinstance(slide, Mapping):
            continue
        slide_index = slide.get("slide_index", "?")
        shapes = [
            item
            for item in slide.get("text_shapes", [])
            if isinstance(item, Mapping) and int(item.get("text_length", 0) or 0) > 0
        ]
        for shape in shapes:
            height = _number(shape.get("height_pt"))
            width = _number(shape.get("width_pt"))
            bound_height = _number(shape.get("bound_height_pt"))
            bound_width = _number(shape.get("bound_width_pt"))
            rotation = abs(_number(shape.get("rotation_degrees")) or 0.0)
            if bool(shape.get("overflowing", False)):
                issues.append(
                    f"Slide {slide_index}: PowerPoint TextFrame2 reports overflowing text"
                )
            if rotation < 1.0 and height and bound_height and bound_height > height * 1.04 + 1:
                issues.append(
                    f"Slide {slide_index}: text may overflow vertically "
                    f"({bound_height:.1f}pt content in {height:.1f}pt box)"
                )
            if rotation < 1.0 and width and bound_width and bound_width > width * 1.04 + 1:
                issues.append(
                    f"Slide {slide_index}: text may overflow horizontally "
                    f"({bound_width:.1f}pt content in {width:.1f}pt box)"
                )
            bound_top = _number(shape.get("bound_top_pt")) or _number(
                shape.get("top_pt")
            )
            if (
                not shape.get("footer_like")
                and bound_top is not None
                and bound_height is not None
                and bound_top + bound_height > footer_start
            ):
                issues.append(
                    f"Slide {slide_index}: body text intrudes into footer exclusion zone"
                )
        for first_index, first in enumerate(shapes):
            for second in shapes[first_index + 1 :]:
                if first.get("footer_like") and second.get("footer_like"):
                    continue
                ratio = _intersection_ratio(first, second)
                first_footer = bool(first.get("footer_like"))
                second_footer = bool(second.get("footer_like"))
                if ratio > 0 and first_footer != second_footer:
                    issues.append(
                        f"Slide {slide_index}: body text overlaps footer or page-number text"
                    )
                elif ratio >= 0.04:
                    issues.append(
                        f"Slide {slide_index}: text boxes overlap "
                        f"({ratio:.0%} of the smaller box)"
                    )
        for shape in slide.get("shapes", []):
            if not isinstance(shape, Mapping):
                continue
            left = _number(shape.get("left_pt"))
            top = _number(shape.get("top_pt"))
            width = _number(shape.get("width_pt"))
            height = _number(shape.get("height_pt"))
            if None in {left, top, width, height}:
                continue
            assert left is not None and top is not None
            assert width is not None and height is not None
            if (
                left < -0.5
                or top < -0.5
                or left + width > slide_width + 0.5
                or top + height > slide_height + 0.5
            ):
                issues.append(
                    f"Slide {slide_index}: shape extends outside the slide canvas"
                )
            role = str(shape.get("shape_role", "content"))
            object_index = shape.get("object_index")
            text_match = next(
                (
                    item
                    for item in shapes
                    if item.get("object_index") == object_index
                ),
                None,
            )
            if isinstance(text_match, Mapping) and bool(
                text_match.get("footer_like")
            ):
                role = "footer"
            if (
                role not in {"footer", "page_number"}
                and top + height > footer_start
                and top < slide_height
            ):
                issues.append(
                    f"Slide {slide_index}: {role} shape intrudes into footer exclusion zone"
                )
    return issues


__all__ = [
    "VisualLayoutQAError",
    "inspect_powerpoint_text_layout",
    "inspect_text_geometry",
]
