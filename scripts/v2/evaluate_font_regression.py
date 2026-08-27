from __future__ import annotations

import argparse
import json
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    manifest = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    rows = []
    failures = []
    for slide in manifest.get("slides", []):
        title = next(
            (
                shape
                for shape in slide.get("text_shapes", [])
                if shape.get("shape_role") == "title"
            ),
            {},
        )
        body = next(
            (
                shape
                for shape in slide.get("text_shapes", [])
                if shape.get("shape_role") == "content"
            ),
            {},
        )
        for shape in (title, body):
            height = float(shape.get("height_pt") or 0)
            bound_height = float(shape.get("bound_height_pt") or 0)
            if bool(shape.get("overflowing", False)) or (
                height > 0 and bound_height > height * 1.04 + 1
            ):
                failures.append(
                    f"slide {slide.get('slide_index')} {shape.get('shape_role')}"
                )
        rows.append(
            {
                "slide": slide.get("slide_index"),
                "font": title.get("font_name", "UNKNOWN"),
                "title_font_pt": title.get("font_size_pt"),
                "title_box_height_pt": title.get("height_pt"),
                "title_bound_height_pt": title.get("bound_height_pt"),
                "body_font_pt": body.get("font_size_pt"),
                "body_box_height_pt": body.get("height_pt"),
                "body_bound_height_pt": body.get("bound_height_pt"),
                "overflow": bool(title.get("overflowing", False))
                or bool(body.get("overflowing", False)),
            }
        )
    args.report.write_text(
        "\n".join(
            [
                "# Font Fallback Regression",
                "",
                f"- Status: `{'PASS' if not failures else 'FAIL'}`",
                "- Authority: PowerPoint COM TextFrame2 actual text bounds",
                "- Cases: Aptos, Arial, Microsoft YaHei, Noto Sans CJK SC, "
                "deliberately missing preferred font",
                "",
                "| Slide | Requested/reported font | Title pt | Title bound/box pt | "
                "Body pt | Body bound/box pt | Overflow |",
                "|---:|---|---:|---:|---:|---:|---|",
                *[
                    f"| {row['slide']} | {row['font']} | {row['title_font_pt']} | "
                    f"{float(row['title_bound_height_pt']):.1f}/"
                    f"{float(row['title_box_height_pt']):.1f} | "
                    f"{row['body_font_pt']} | "
                    f"{float(row['body_bound_height_pt']):.1f}/"
                    f"{float(row['body_box_height_pt']):.1f} | "
                    f"{'FAIL' if row['overflow'] else 'PASS'} |"
                    for row in rows
                ],
                "",
                *(
                    [f"- FAIL: {item}" for item in failures]
                    if failures
                    else [
                        "- PASS: PowerPoint reported no overflowing title or body "
                        "text in any requested or substituted-font case."
                    ]
                ),
            ]
        ),
        encoding="utf-8",
    )
    print(f"STATUS={'PASS' if not failures else 'FAIL'}")
    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
