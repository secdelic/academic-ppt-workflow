from __future__ import annotations

import argparse
import csv
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageChops, ImageFilter, ImageStat


def edge_ratio(image: Image.Image) -> float:
    edges = image.convert("L").filter(ImageFilter.FIND_EDGES)
    values = list(edges.getdata())
    return sum(value > 24 for value in values) / max(1, len(values))


def normalized_difference(first: Image.Image, second: Image.Image) -> float:
    left = first.convert("RGB").resize((320, 180))
    right = second.convert("RGB").resize((320, 180))
    stat = ImageStat.Stat(ImageChops.difference(left, right))
    return sum(stat.mean) / (3 * 255)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--powerpoint-preview", type=Path, required=True)
    parser.add_argument("--libreoffice-pdf", type=Path, required=True)
    parser.add_argument("--output-preview", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--csv-report", type=Path, required=True)
    args = parser.parse_args()
    powerpoints = sorted(args.powerpoint_preview.glob("slide_*.png"))
    args.output_preview.mkdir(parents=True, exist_ok=True)
    document = pdfium.PdfDocument(str(args.libreoffice_pdf))
    rows = []
    severe = []
    warnings = []
    for index in range(len(document)):
        page = document[index]
        bitmap = page.render(scale=2.25)
        image = bitmap.to_pil().convert("RGB")
        output = args.output_preview / f"slide_{index + 1:03d}.png"
        image.save(output)
        variance = sum(ImageStat.Stat(image).var) / 3
        footer_top = int(image.height * 0.885)
        footer = image.crop((0, footer_top, image.width, image.height))
        footer_edges = edge_ratio(footer)
        diff = (
            normalized_difference(Image.open(powerpoints[index]), image)
            if index < len(powerpoints)
            else 1.0
        )
        page_issues = []
        if variance < 2.0:
            page_issues.append("blank_or_near_blank")
            severe.append(f"LibreOffice slide {index + 1} is blank or near blank")
        if footer_edges > 0.085:
            page_issues.append("footer_linework_dense")
            severe.append(
                f"LibreOffice slide {index + 1} has dense footer-region linework"
            )
        if diff > 0.24:
            warnings.append(
                f"Slide {index + 1} perceptual difference is {diff:.1%}"
            )
        rows.append(
            {
                "slide_index": index + 1,
                "powerpoint_png": (
                    str(powerpoints[index]) if index < len(powerpoints) else ""
                ),
                "libreoffice_png": str(output),
                "libreoffice_variance": f"{variance:.4f}",
                "footer_edge_ratio": f"{footer_edges:.6f}",
                "perceptual_difference": f"{diff:.6f}",
                "status": "FAIL" if page_issues else "PASS",
                "issues": ";".join(page_issues),
            }
        )
    if len(powerpoints) != len(document):
        severe.append(
            f"Page-count mismatch: PowerPoint={len(powerpoints)}, "
            f"LibreOffice={len(document)}"
        )
    args.csv_report.parent.mkdir(parents=True, exist_ok=True)
    with args.csv_report.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]) if rows else [])
        if rows:
            writer.writeheader()
            writer.writerows(rows)
    status = "PASS" if not severe else "FAIL"
    args.report.write_text(
        "\n".join(
            [
                "# Cross-renderer Comparison",
                "",
                f"- Status: `{status}`",
                f"- PowerPoint pages: {len(powerpoints)}",
                f"- LibreOffice pages: {len(document)}",
                "- PowerPoint geometry authority: COM TextFrame2 actual bounds",
                "- LibreOffice compatibility authority: PDF page count, nonblank "
                "render, footer-region linework and perceptual comparison",
                "",
                "## Severe issues",
                "",
                *(f"- FAIL: {item}" for item in severe),
                *(["- PASS: no blank pages, page loss, or severe footer intrusion"] if not severe else []),
                "",
                "## Compatibility warnings",
                "",
                *(f"- WARN: {item}" for item in warnings),
                *(["- None"] if not warnings else []),
                "",
                "Minor renderer-specific font and chart-label differences are "
                "warnings unless they cause clipping, overlap, missing pages, "
                "or missing graphics.",
            ]
        ),
        encoding="utf-8",
    )
    print(f"STATUS={status}")
    print(f"POWERPOINT_PAGES={len(powerpoints)}")
    print(f"LIBREOFFICE_PAGES={len(document)}")
    if severe:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
