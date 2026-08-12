from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import random
import statistics
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageStat


def read_csv(path: Path) -> list[dict[str, str]]:
    return list(csv.DictReader(path.open("r", encoding="utf-8-sig", newline="")))


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not fields:
        fields = list(rows[0]) if rows else ["status"]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def intersection(a: dict[str, float], b: dict[str, float]) -> float:
    left = max(a["left"], b["left"])
    top = max(a["top"], b["top"])
    right = min(a["left"] + a["width"], b["left"] + b["width"])
    bottom = min(a["top"] + a["height"], b["top"] + b["height"])
    return max(0.0, right - left) * max(0.0, bottom - top)


def image_metrics(path: Path) -> dict[str, float]:
    image = Image.open(path).convert("RGB").resize((400, 225))
    pixels = list(image.getdata())
    nonwhite = sum(1 for r, g, b in pixels if min(r, g, b) < 244)
    dark = sum(1 for r, g, b in pixels if max(r, g, b) < 80)
    return {
        "content_coverage": round(nonwhite / len(pixels), 4),
        "dark_ratio": round(dark / len(pixels), 4),
        "blank_area_ratio": round(1 - nonwhite / len(pixels), 4),
    }


def rms_difference(first: Path, second: Path) -> float:
    a = Image.open(first).convert("RGB").resize((400, 225))
    b = Image.open(second).convert("RGB").resize((400, 225))
    stat = ImageStat.Stat(ImageChops.difference(a, b))
    return round(math.sqrt(sum(value * value for value in stat.rms) / 3), 3)


def validate_ooxml(pptx: Path) -> dict[str, Any]:
    errors: list[str] = []
    with zipfile.ZipFile(pptx) as archive:
        names = set(archive.namelist())
        required = {"[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
        for name in required - names:
            errors.append(f"missing:{name}")
        for name in names:
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(archive.read(name))
                except ET.ParseError as exc:
                    errors.append(f"malformed:{name}:{exc}")
        slides = sorted(name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
        for slide in slides:
            rel = slide.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            if rel not in names:
                errors.append(f"missing_slide_rels:{rel}")
        missing_targets = []
        for rel_name in (name for name in names if name.endswith(".rels")):
            try:
                root = ET.fromstring(archive.read(rel_name))
            except ET.ParseError:
                continue
            base = Path(rel_name).parent.parent if Path(rel_name).parent.name == "_rels" else Path(rel_name).parent
            for rel in root:
                if rel.attrib.get("TargetMode") == "External":
                    continue
                target = rel.attrib.get("Target")
                if not target:
                    continue
                candidate = (base / target).as_posix()
                parts = []
                for part in candidate.split("/"):
                    if part == "..":
                        if parts:
                            parts.pop()
                    elif part not in {"", "."}:
                        parts.append(part)
                normalized = "/".join(parts)
                if normalized not in names:
                    missing_targets.append(f"{rel_name}->{target}")
        errors.extend(f"missing_target:{item}" for item in missing_targets)
    return {"slide_count": len(slides), "critical_errors": len(errors), "errors": errors}


def contact_sheet(image_paths: list[Path], output: Path, label: str) -> None:
    thumbs = []
    for slide_index, path in enumerate(image_paths):
        image = Image.open(path).convert("RGB")
        image.thumbnail((320, 180))
        if slide_index == 0:
            # Remove the benchmark arm token from the rendered cover while
            # preserving the project title and synthetic warning.
            ImageDraw.Draw(image).rectangle(
                (18, 68, 305, 108), fill=(23, 50, 77)
            )
        canvas = Image.new("RGB", (330, 208), "white")
        canvas.paste(image, ((330 - image.width) // 2, 8))
        ImageDraw.Draw(canvas).text((10, 188), path.stem, fill=(35, 55, 75))
        thumbs.append(canvas)
    cols = 3
    rows = math.ceil(len(thumbs) / cols)
    sheet = Image.new("RGB", (cols * 330, rows * 208 + 42), (238, 244, 247))
    ImageDraw.Draw(sheet).text((16, 12), label, fill=(23, 50, 77))
    for index, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((index % cols) * 330, 42 + (index // cols) * 208))
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--staging-root", type=Path, required=True)
    parser.add_argument("--benchmark-output", type=Path, required=True)
    parser.add_argument("--audit-output", type=Path, required=True)
    parser.add_argument("--input-hashes-before", type=Path, required=True)
    parser.add_argument("--suite", type=Path, required=True)
    args = parser.parse_args()

    deck_rows = read_csv(args.run_root / "deck_index.csv")
    pp_rows = {(r["project_key"], r["arm"]): r for r in load_json(args.run_root / "powerpoint_render_index.json")}
    lo_rows = {(r["project_key"], r["arm"]): r for r in load_json(args.run_root / "libreoffice_render_index.json")}
    slide_rows: list[dict[str, Any]] = []
    deck_results: list[dict[str, Any]] = []
    ooxml_rows: list[dict[str, Any]] = []
    source_leaks: list[dict[str, Any]] = []
    style_rows: list[dict[str, Any]] = []
    projects = sorted({r["project_key"] for r in deck_rows})

    project_source_ids: dict[str, set[str]] = {}
    for project in projects:
        manifest_path = args.staging_root / project / "source_manifest.csv"
        manifest = read_csv(manifest_path)
        project_source_ids[project] = {row["source_id"] for row in manifest}
        gt_entries = [row for row in manifest if "expected_ground_truth" in row["relative_path"]]
        style_status = load_json(args.staging_root / project / "style_status.json")
        style_rows.append({
            "project_key": project,
            "status": style_status["status"],
            "pptx_or_potx_count": style_status["pptx_or_potx_count"],
            "cache_used": style_status["cache_used"],
            "ground_truth_manifest_entries": len(gt_entries),
        })
        for other in projects:
            if other == project:
                continue
            overlap = project_source_ids.get(other, set()) & project_source_ids[project]
            if overlap:
                source_leaks.append({"project_a": project, "project_b": other, "source_ids": "|".join(sorted(overlap))})

    for deck in deck_rows:
        key = (deck["project_key"], deck["arm"])
        pp = pp_rows[key]
        lo = lo_rows[key]
        geometry = load_json(Path(pp["geometry"]))
        pptx = Path(deck["pptx_path"])
        ooxml = validate_ooxml(pptx)
        ooxml_rows.append({
            "project_key": deck["project_key"],
            "arm": deck["arm"],
            "slide_count": ooxml["slide_count"],
            "critical_errors": ooxml["critical_errors"],
            "errors": "|".join(ooxml["errors"]),
        })
        severe_overlap = footer_intrusion = clipping = off_slide = 0
        pp_images = sorted(Path(pp["preview"]).glob("slide_*.png"))
        lo_images = sorted(Path(lo["preview"]).glob("slide_*.png"))
        for slide_data, pp_image, lo_image in zip(geometry["slides"], pp_images, lo_images, strict=False):
            slide_height = float(slide_data["height"])
            footer_top = slide_height * 0.952
            texts = []
            this_footer = this_clip = this_overlap = this_off = 0
            for shape in slide_data["shapes"]:
                if shape.get("off_slide"):
                    this_off += 1
                if shape.get("overflowing"):
                    this_clip += 1
                if shape.get("has_text") and shape.get("bound_top") is not None:
                    bound = {
                        "left": float(shape["bound_left"]),
                        "top": float(shape["bound_top"]),
                        "width": float(shape["bound_width"]),
                        "height": float(shape["bound_height"]),
                        "text": str(shape.get("text", "")),
                    }
                    if bound["top"] < footer_top and bound["top"] + bound["height"] > footer_top:
                        this_footer += 1
                    texts.append(bound)
            for i, first in enumerate(texts):
                for second in texts[i + 1:]:
                    area = intersection(first, second)
                    if area > 20 and first["text"].strip() and second["text"].strip():
                        this_overlap += 1
            metrics = image_metrics(pp_image)
            diff = rms_difference(pp_image, lo_image)
            slide_rows.append({
                "project_key": deck["project_key"],
                "arm": deck["arm"],
                "slide_number": slide_data["slide_number"],
                "footer_intrusion": this_footer,
                "text_overlap_pairs": this_overlap,
                "text_clipping": this_clip,
                "off_slide_shapes": this_off,
                "content_coverage": metrics["content_coverage"],
                "blank_area_ratio": metrics["blank_area_ratio"],
                "renderer_rms_difference": diff,
                "powerpoint_png": str(pp_image),
                "libreoffice_png": str(lo_image),
            })
            footer_intrusion += this_footer
            clipping += this_clip
            severe_overlap += this_overlap
            off_slide += this_off
        deck_results.append({
            "project_key": deck["project_key"],
            "arm": deck["arm"],
            "slide_count": deck["slide_count"],
            "visual_count": deck["visual_count"],
            "semantic_critical_issues": deck["semantic_critical_issues"],
            "adapter_fallback_count": deck["fallback_count"],
            "powerpoint_open": pp["powerpoint_open"],
            "powerpoint_page_match": int(pp["pptx_slide_count"]) == int(pp["png_count"]) == int(deck["slide_count"]),
            "libreoffice_page_match": int(lo["pdf_page_count"]) == int(deck["slide_count"]),
            "footer_intrusion": footer_intrusion,
            "severe_text_overlap_pairs": severe_overlap,
            "text_clipping": clipping,
            "off_slide_shapes": off_slide,
            "ooxml_critical_errors": ooxml["critical_errors"],
            "generation_seconds": deck["generation_seconds"],
            "libreoffice_seconds": lo["runtime_seconds"],
            "output_size_bytes": pptx.stat().st_size,
        })

    args.benchmark_output.mkdir(parents=True, exist_ok=True)
    args.audit_output.mkdir(parents=True, exist_ok=True)
    write_csv(args.benchmark_output / "project_results.csv", deck_results)
    write_csv(args.benchmark_output / "slide_level_results.csv", slide_rows)
    write_csv(args.audit_output / "ooxml_validation.csv", ooxml_rows)
    write_csv(args.audit_output / "style_source_audit.csv", style_rows)
    write_csv(args.audit_output / "cross_project_leakage.csv", source_leaks)

    arm_agg = []
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in deck_results:
        grouped[row["arm"]].append(row)
    for arm, rows in sorted(grouped.items()):
        arm_agg.append({
            "arm": arm,
            "deck_count": len(rows),
            "semantic_critical_issues": sum(int(r["semantic_critical_issues"]) for r in rows),
            "severe_overlap": sum(int(r["severe_text_overlap_pairs"]) for r in rows),
            "clipping": sum(int(r["text_clipping"]) for r in rows),
            "footer_intrusion": sum(int(r["footer_intrusion"]) for r in rows),
            "ooxml_critical_errors": sum(int(r["ooxml_critical_errors"]) for r in rows),
            "powerpoint_open_rate": sum(str(r["powerpoint_open"]).lower() == "true" for r in rows) / len(rows),
            "powerpoint_page_match_rate": sum(bool(r["powerpoint_page_match"]) for r in rows) / len(rows),
            "libreoffice_page_match_rate": sum(bool(r["libreoffice_page_match"]) for r in rows) / len(rows),
            "mean_generation_seconds": round(statistics.mean(float(r["generation_seconds"]) for r in rows), 3),
            "mean_libreoffice_seconds": round(statistics.mean(float(r["libreoffice_seconds"]) for r in rows), 3),
            "adapter_fallback_count": sum(int(r["adapter_fallback_count"]) for r in rows),
        })
    write_csv(args.benchmark_output / "arm_results.csv", arm_agg)

    runtime_rows = [{
        "project_key": row["project_key"],
        "arm": row["arm"],
        "generation_seconds": row["generation_seconds"],
        "libreoffice_seconds": row["libreoffice_seconds"],
        "output_size_bytes": row["output_size_bytes"],
    } for row in deck_results]
    write_csv(args.benchmark_output / "runtime_results.csv", runtime_rows)

    # Anonymous contact sheets. Mapping is deliberately separate from review form.
    randomizer = random.Random(22077)
    arm_codes = ["Deck A", "Deck B", "Deck C", "Deck D"]
    mapping: dict[str, dict[str, str]] = {}
    review_rows = []
    for project in projects:
        shuffled = list("ABCD")
        randomizer.shuffle(shuffled)
        mapping[project] = {arm: arm_codes[index] for index, arm in enumerate(shuffled)}
        for arm in "ABCD":
            pp = pp_rows[(project, arm)]
            code = mapping[project][arm]
            contact = args.benchmark_output / "contact_sheets" / f"{project}_{code.replace(' ', '_')}.png"
            contact_sheet(sorted(Path(pp["preview"]).glob("slide_*.png")), contact, f"{project} · {code}")
            review_rows.append({
                "project_key": project,
                "blind_code": code,
                "contact_sheet": str(contact),
                "scientific_clarity_20": "",
                "chart_semantic_correctness_20": "",
                "visual_hierarchy_15": "",
                "composition_15": "",
                "chart_professionalism_10": "",
                "consistency_10": "",
                "editability_5": "",
                "presentation_readiness_5": "",
                "total_100": "",
                "reviewer_notes": "",
            })
    (args.audit_output / "blind_code_map.json").write_text(json.dumps(mapping, indent=2), encoding="utf-8")
    write_csv(args.benchmark_output / "visual_blind_review_form.csv", review_rows)
    write_csv(args.benchmark_output / "manual_repair_log.csv", [], [
        "project_key", "arm", "slide_id", "reason", "repair_count", "status"
    ])

    # Input integrity: compare all non-ground-truth suite files with frozen pre-run hashes.
    before = {row["relative_path"]: row for row in read_csv(args.input_hashes_before)}
    after_rows = []
    for rel, old in sorted(before.items()):
        path = args.suite / Path(rel)
        digest = hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else ""
        after_rows.append({
            "relative_path": rel,
            "absolute_path": str(path),
            "sha256_before": old["sha256"],
            "sha256_after": digest,
            "unchanged": digest == old["sha256"],
        })
    write_csv(args.audit_output / "suite_input_hashes_after.csv", after_rows)
    integrity_ok = all(row["unchanged"] for row in after_rows)
    (args.audit_output / "input_integrity_report.md").write_text(
        "# Input integrity\n\n"
        f"- Compared files: {len(after_rows)}\n"
        f"- Unchanged: {sum(bool(r['unchanged']) for r in after_rows)}\n"
        f"- Result: {'PASS' if integrity_ok else 'FAIL'}\n"
        "- `expected_ground_truth/` was excluded from generation-time hashing and discovery.\n",
        encoding="utf-8",
    )
    (args.run_root / "initial_qa_complete.json").write_text(
        json.dumps({
            "completed_at_utc": datetime.now(timezone.utc).isoformat(),
            "ground_truth_content_read": False,
            "powerpoint_decks": len(pp_rows),
            "libreoffice_decks": len(lo_rows),
            "input_integrity": integrity_ok,
        }, indent=2),
        encoding="utf-8",
    )
    print(f"ANALYZED_DECKS={len(deck_results)}")
    print(f"INPUT_INTEGRITY={integrity_ok}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
