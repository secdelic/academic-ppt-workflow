from __future__ import annotations

import argparse
import csv
import json
import shutil
import zipfile
from pathlib import Path
from typing import Any


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, fields: list[str], rows: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--audit-dir", type=Path, required=True)
    parser.add_argument("--previous-deck-ir", type=Path)
    args = parser.parse_args()
    output = args.output_dir.resolve()
    audit = args.audit_dir.resolve()
    audit.mkdir(parents=True, exist_ok=True)

    layout = json.loads(
        (output / "powerpoint_layout_manifest.json").read_text(
            encoding="utf-8-sig"
        )
    )
    slide_height = float(layout.get("slide_height_pt", 540))
    footer_start = slide_height * 0.885
    footer_rows = []
    text_rows = []
    for slide in layout.get("slides", []):
        slide_index = slide.get("slide_index")
        slide_id = slide.get("slide_id", "")
        intrusions = 0
        for shape in slide.get("shapes", []):
            top = float(shape.get("top_pt") or 0)
            height = float(shape.get("height_pt") or 0)
            role = str(shape.get("shape_role", "content"))
            if role not in {"footer", "page_number"} and top + height > footer_start:
                intrusions += 1
        footer_rows.append(
            {
                "slide_index": slide_index,
                "slide_id": slide_id,
                "footer_start_pt": f"{footer_start:.3f}",
                "prohibited_intrusion_count": intrusions,
                "status": "PASS" if intrusions == 0 else "FAIL",
            }
        )
        for shape in slide.get("text_shapes", []):
            height = float(shape.get("height_pt") or 0)
            width = float(shape.get("width_pt") or 0)
            bound_height = float(shape.get("bound_height_pt") or 0)
            bound_width = float(shape.get("bound_width_pt") or 0)
            vertical_overflow = bool(shape.get("overflowing", False)) or (
                height > 0 and bound_height > height * 1.04 + 1
            )
            horizontal_overflow = (
                width > 0 and bound_width > width * 1.04 + 1
            )
            text_rows.append(
                {
                    "slide_index": slide_index,
                    "slide_id": slide_id,
                    "shape_role": shape.get("shape_role", ""),
                    "text_length": shape.get("text_length", 0),
                    "box_width_pt": f"{width:.3f}",
                    "box_height_pt": f"{height:.3f}",
                    "bound_width_pt": f"{bound_width:.3f}",
                    "bound_height_pt": f"{bound_height:.3f}",
                    "textframe2_overflowing": shape.get("overflowing", False),
                    "vertical_overflow": vertical_overflow,
                    "horizontal_overflow": horizontal_overflow,
                    "status": (
                        "FAIL"
                        if vertical_overflow or horizontal_overflow
                        else "PASS"
                    ),
                }
            )
    write_csv(
        audit / "footer_intrusion_report.csv",
        [
            "slide_index",
            "slide_id",
            "footer_start_pt",
            "prohibited_intrusion_count",
            "status",
        ],
        footer_rows,
    )
    write_csv(
        audit / "text_bounds_report.csv",
        [
            "slide_index",
            "slide_id",
            "shape_role",
            "text_length",
            "box_width_pt",
            "box_height_pt",
            "bound_width_pt",
            "bound_height_pt",
            "textframe2_overflowing",
            "vertical_overflow",
            "horizontal_overflow",
            "status",
        ],
        text_rows,
    )
    for name in (
        "visual_density_report.csv",
        "layout_repetition_report.csv",
        "low_information_slide_report.csv",
        "overloaded_slide_report.csv",
    ):
        shutil.copy2(output / name, audit / name)

    contracts = json.loads(
        (output / "chart_data_contracts.json").read_text(encoding="utf-8")
    ).get("contracts", [])
    chart_rows = []
    for contract in contracts:
        denominators = list(contract.get("denominator", []))
        total = contract.get("sample_size_total")
        calculated = sum(float(value) for value in denominators) if denominators else None
        sample_ok = (
            calculated is None
            or (total is not None and abs(float(total) - calculated) < 1e-9)
        )
        category_ok = len(contract.get("categories", [])) == len(
            contract.get("values", [])
        )
        caption_text = ""
        for slide in json.loads(
            (output / "deck_ir.json").read_text(encoding="utf-8")
        ).get("slides", []):
            if slide.get("chart_data", {}).get("chart_id") == contract.get(
                "chart_id"
            ):
                caption_text = str(slide.get("chart_caption", ""))
                break
        legacy_error = "sample size: not provided" in caption_text.lower()
        chart_rows.append(
            {
                "chart_id": contract.get("chart_id"),
                "source_id": contract.get("source_id"),
                "content_type": contract.get("content_type"),
                "category_value_lengths_match": category_ok,
                "denominator_sum": "" if calculated is None else calculated,
                "sample_size_total": "" if total is None else total,
                "sample_size_consistent": sample_ok,
                "caption": caption_text,
                "legacy_not_provided_error": legacy_error,
                "status": (
                    "PASS"
                    if category_ok and sample_ok and not legacy_error
                    else "FAIL"
                ),
            }
        )
    write_csv(
        audit / "chart_metadata_validation.csv",
        [
            "chart_id",
            "source_id",
            "content_type",
            "category_value_lengths_match",
            "denominator_sum",
            "sample_size_total",
            "sample_size_consistent",
            "caption",
            "legacy_not_provided_error",
            "status",
        ],
        chart_rows,
    )

    claim_rows = read_csv(output / "claim_source_map.csv")
    source_rows = read_csv(output / "source_manifest.csv")
    known_sources = {row["source_id"] for row in source_rows}
    claims_complete = all(
        row.get("claim_id")
        and row.get("source_id") in known_sources
        and row.get("source_location")
        for row in claim_rows
    )
    searchable = "\n".join(
        [
            (output / "deck_outline.md").read_text(encoding="utf-8"),
            (output / "speaker_notes.md").read_text(encoding="utf-8"),
            (output / "deck_ir.json").read_text(encoding="utf-8"),
        ]
    )
    wrong_values = [
        token for token in ("33/58", "56.9%", "OR 3.41") if token in searchable
    ]
    required_values = [
        token
        for token in ("34/58", "58.6%", "1.42", "2.06", "3.61")
        if token in searchable
    ]
    forbidden = [
        token
        for token in (
            " caused ",
            " proved ",
            "confirmed mechanism",
            "independent causal effect",
        )
        if token.lower() in searchable.lower()
    ]
    unresolved = [
        token
        for token in ("ETHICS_ID_REQUIRED", "REF_PENDING_04")
        if token in searchable
    ]
    with zipfile.ZipFile(output / "SYN_CARDIO_AKI_SHADOW_VALIDATION.pptx") as zf:
        bad_zip_member = zf.testzip()
    current_ir = json.loads(
        (output / "deck_ir.json").read_text(encoding="utf-8")
    )
    stable_rows = []
    if args.previous_deck_ir and args.previous_deck_ir.is_file():
        previous_ir = json.loads(args.previous_deck_ir.read_text(encoding="utf-8"))
        previous_by_key = {
            slide["logical_slide_key"]: slide["slide_id"]
            for slide in previous_ir.get("slides", [])
        }
        for slide in current_ir.get("slides", []):
            key = slide.get("logical_slide_key")
            if key in previous_by_key:
                stable_rows.append(previous_by_key[key] == slide.get("slide_id"))
    scientific_pass = (
        claims_complete
        and len(claim_rows) >= 7
        and not wrong_values
        and len(required_values) == 5
        and not forbidden
        and len(unresolved) == 2
        and bad_zip_member is None
        and all(stable_rows)
    )
    (audit / "scientific_regression_report.md").write_text(
        "\n".join(
            [
                "# Scientific Regression Report",
                "",
                f"- Status: `{'PASS' if scientific_pass else 'FAIL'}`",
                f"- Formal claims: {len(claim_rows)}",
                f"- Claim-source completeness: {'PASS' if claims_complete else 'FAIL'}",
                f"- Canonical required values present: {', '.join(required_values)}",
                f"- Non-canonical values found: {', '.join(wrong_values) or 'none'}",
                f"- Prohibited causal/mechanistic wording found: {', '.join(forbidden) or 'none'}",
                f"- Unresolved markers retained: {', '.join(unresolved)}",
                f"- Stable semantic IDs for unchanged logical slides: "
                f"{'PASS' if all(stable_rows) else 'FAIL'} "
                f"({sum(stable_rows)}/{len(stable_rows)})",
                "- Frozen models were consumed from registered aggregate results; "
                "the workflow did not fit or refit a statistical model.",
                "- Final scientific approval remains manual.",
            ]
        ),
        encoding="utf-8",
    )
    print(f"FOOTER_FAILURES={sum(row['status'] == 'FAIL' for row in footer_rows)}")
    print(f"TEXT_BOUND_FAILURES={sum(row['status'] == 'FAIL' for row in text_rows)}")
    print(f"CHART_METADATA_FAILURES={sum(row['status'] == 'FAIL' for row in chart_rows)}")
    print(f"SCIENTIFIC_REGRESSION={'PASS' if scientific_pass else 'FAIL'}")


if __name__ == "__main__":
    main()
