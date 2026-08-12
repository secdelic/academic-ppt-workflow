from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import subprocess
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

from PIL import Image, ImageDraw


REPO = Path(__file__).resolve().parents[2]

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

ALLOWED_LATIN_TERMS = (
    "Target Trial Emulation", "Target Trial", "Clone-Censor-Weight",
    "Clone–Censor–Weight", "pseudobulk", "UMAP", "GRADE", "CellChat",
    "TTE", "AKI", "ICU", "SIMD", "QC", "DAG", "SMD", "CCW", "FDR",
    "NES", "OR", "HR", "RR", "RD", "CI", "I2", "I²", "SOFA",
    "Control", "log2FC", "time zero", "Day", "Meta", "PRISMA",
)
MOJIBAKE_RE = re.compile(
    r"锛|銆|鈥|鐨|鍚堟垚|鍏宠仈|缁撴灉|鐮旂┒|涓嶅緱|"
    r"鍒嗘瀽|棰勬祴|鏂规|瀹氫箟|鏁堝簲|闄愬埗|鍙В|"
    r"(?:[\u4e00-\u9fff])[?](?:[\u4e00-\u9fff])"
)
ABNORMAL_SPACING_RE = re.compile(r"(?<![A-Za-z])(?:[A-Za-z]\s+){2,}[A-Za-z](?![A-Za-z])")
GENERIC_TITLE_RE = re.compile(
    r"^(?:background|methods?|results?|discussion|conclusion|overview|"
    r"背景|方法|结果|讨论|结论|概览|研究设计|分析结果)\s*\d*$",
    re.I,
)
NUMERIC_RE = re.compile(r"[-+]?\d[\d,]*(?:\.\d+)?(?:e[-+]?\d+)?", re.I)

VISUAL_FIELD_ROLES = (
    "x_field", "y_field", "group_field", "facet_field", "color_field",
    "label_field", "ci_low_field", "ci_high_field", "time_field",
    "edge_source_field", "edge_target_field", "end_time_field",
)
RESERVED_FIELDS = {"document_text", "registered_figure"}
DERIVED_FIELD_BASES = {
    "position_hours": {"window"},
    "probability_score": {"probability"},
    "impact_score": {"impact"},
    "delta": {"percentage"},
    "bin_low": {"stabilized_weight"},
    "bin_high": {"stabilized_weight"},
    "bin": {"stabilized_weight"},
    "count": {"stabilized_weight"},
}

EXPECTED_SOURCE_RULES: dict[str, list[tuple[str, str, str]]] = {
    "target_trial": [
        ("love_plot", "03_balance_diagnostics.csv", "Love plot"),
        ("forest_plot:sensitivity_analysis", "05_sensitivity_results.csv", "sensitivity analysis"),
        ("absolute_risk_comparison", "04_effect_estimates.csv", "absolute effect summary"),
        ("estimand_comparison", "04_effect_estimates.csv", "ratio effect summary"),
        ("specification_matrix", "protocol", "target-trial specification"),
    ],
    "single_cell": [
        ("qc_funnel", "06_qc_summary.csv", "QC"),
        ("grouped_composition", "02_celltype_composition.csv", "cell composition"),
        ("faceted_volcano", "03_pseudobulk_deg.csv", "pseudobulk DEG"),
        ("diverging_enrichment", "04_pathway_enrichment.csv", "pathway enrichment"),
        ("communication_network", "05_cellchat_interactions.csv", "cell communication"),
    ],
    "meta_analysis": [
        ("forest_plot:primary_meta_analysis", "01_study_level_data.csv", "study forest"),
        ("pooled_evidence_panel", "02_pooled_effects.csv", "pooled estimate"),
        ("risk_of_bias_matrix", "03_risk_of_bias.csv", "risk of bias"),
        ("subgroup_forest", "04_subgroup_meta.csv", "subgroup analysis"),
        ("grade_summary", "05_grade_summary.csv", "GRADE"),
    ],
    "protocol": [
        ("eligibility_split", "01_eligibility_criteria.csv", "eligibility"),
        ("assessment_timeline", "02_schedule_of_assessments.csv", "schedule"),
        ("sample_size_waterfall", "03_sample_size_assumptions.csv", "sample size"),
        ("variable_definition_matrix", "04_variable_dictionary.csv", "variable definitions"),
        ("risk_matrix", "05_risk_register.csv", "risk register"),
        ("timeline_gantt", "06_project_milestones.csv", "milestones"),
    ],
}

CSV_FIELDS: dict[str, list[str]] = {
    "source_binding_accuracy.csv": [
        "project_key", "slide_id", "visual_id", "check_type", "source_id",
        "source_file", "source_location", "fields_used", "row_filter",
        "aggregation", "canonical_status", "expected_source", "passed", "details",
    ],
    "language_compliance.csv": [
        "project_key", "slide_id", "declared_language", "cjk_characters",
        "non_allowlisted_latin_characters", "zh_cn_coverage", "mojibake_detected",
        "abnormal_spacing_detected", "title_has_cjk", "passed",
    ],
    "narrative_coverage_report.csv": [
        "project_key", "must_include_item", "expected_claim", "expected_visual",
        "expected_row_count", "target_slide_count", "rendered_slide",
        "appendix_location", "omitted", "omission_reason",
        "manual_decision_required", "passed",
    ],
    "omission_manifest.csv": [
        "project_key", "slide_id", "visual_id", "item_type", "expected_count",
        "rendered_count", "omitted_rows", "reason", "silent", "passed",
    ],
    "chart_semantic_qa.csv": [
        "project_key", "slide_id", "visual_id", "visual_type", "check_id",
        "severity", "passed", "details",
    ],
    "layout_family_report.csv": [
        "project_key", "slide_count", "layout_family_count", "layout_families",
        "card_grid_count", "card_grid_ratio", "text_only_count", "text_only_ratio",
        "max_consecutive_same_layout", "generic_title_count", "generic_title_rate",
        "result_slide_without_hero", "passed",
    ],
    "typography_report.csv": [
        "project_key", "slide_number", "slide_id", "title_min_pt", "body_min_pt",
        "chart_detail_min_pt", "footer_min_pt", "title_under_30",
        "body_under_18_unjustified", "chart_under_15", "card_detail_under_16",
        "footer_outside_10_11", "declared_typefaces", "passed",
    ],
    "perceptual_visual_qa.csv": [
        "project_key", "slide_number", "slide_id", "content_coverage",
        "whitespace_ratio", "foreground_centroid_x", "foreground_centroid_y",
        "grid_occupancy", "lower_half_emptiness", "largest_visual_anchor_ratio",
        "perceptual_density_score", "optical_balance_score", "hero_anchor_score",
        "chart_readability_score", "text_overlap_pairs", "clipping",
        "footer_intrusion", "off_slide_shapes", "gross_visual_failure", "warnings",
    ],
    "duplicated_state_inventory.csv": [
        "artifact", "authority", "derivation", "independent_inference_allowed", "status",
    ],
    "module_responsibility_matrix.csv": [
        "module", "responsibility", "input_contract", "output_contract", "authority",
    ],
    "input_integrity.csv": [
        "project_key", "source_file", "sha256_before", "sha256_after", "unchanged",
        "registered", "ground_truth_excluded", "passed",
    ],
    "ground_truth_claim_comparison.csv": [
        "project_key", "claim_id", "identity_recalled", "value_recalled",
        "canonical_source_match", "classification",
    ],
    "ground_truth_conflict_comparison.csv": [
        "project_key", "conflict_id", "recalled", "classification",
    ],
    "ground_truth_unresolved_comparison.csv": [
        "project_key", "item_id", "marker", "recalled", "classification",
    ],
    "geometry_qa.csv": [
        "project_key", "slide_number", "slide_id", "text_overlap_pairs",
        "clipping", "footer_intrusion", "off_slide_shapes", "powerpoint_png",
        "libreoffice_png", "passed",
    ],
    "contact_sheet_code_map.csv": ["anonymous_deck_code", "project_key", "contact_sheet"],
    "generic_title_report.csv": ["project_key", "slide_id", "title", "generic", "reason"],
}


class BlockedError(RuntimeError):
    pass


def as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "pass", "passed"}


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise BlockedError(f"Required CSV is missing: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, records: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = fields or (list(records[0]) if records else CSV_FIELDS.get(path.name, ["status"]))
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)


def read_json(path: Path) -> Any:
    if not path.is_file():
        raise BlockedError(f"Required JSON is missing: {path}")
    return json.loads(path.read_text(encoding="utf-8-sig"))


def index_records(path: Path) -> list[dict[str, Any]]:
    value = read_json(path)
    if isinstance(value, list):
        return value
    if isinstance(value, dict):
        for key in ("results", "decks", "records"):
            if isinstance(value.get(key), list):
                return value[key]
    raise BlockedError(f"Unsupported render-index shape: {path}")


def index_by_project(records: Iterable[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for record in records:
        project = str(record.get("project_key", ""))
        arm = str(record.get("arm", "V24"))
        result[(project, arm)] = record
    return result


def lookup_index(
    index: dict[tuple[str, str], dict[str, Any]], project: str, arm: str
) -> dict[str, Any]:
    match = index.get((project, arm)) or index.get((project, "V24"))
    if match is None:
        candidates = [value for (key, _), value in index.items() if key == project]
        if len(candidates) == 1:
            return candidates[0]
        raise BlockedError(f"Render index has no unambiguous entry for {project}/{arm}")
    return match


def flatten_strings(value: Any) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, dict):
        return [item for child in value.values() for item in flatten_strings(child)]
    if isinstance(value, list):
        return [item for child in value for item in flatten_strings(child)]
    return []


def nested_record_fields(value: Any) -> set[str]:
    """Return keys declared by record-like payload objects at any nesting level."""
    fields: set[str] = set()
    if isinstance(value, dict):
        fields.update(value)
        for child in value.values():
            fields.update(nested_record_fields(child))
    elif isinstance(value, list):
        for child in value:
            fields.update(nested_record_fields(child))
    return fields


def audience_payload_strings(value: Any, *, structured_rows: bool = False) -> list[str]:
    """Collect audience narrative while excluding raw scientific entity tables.

    Values inside ``payload.rows`` remain source-bound chart data. Latin study,
    gene, pathway, visit, and variable names in those rows are professional
    entities rather than untranslated prose, so they are excluded from the
    zh-CN narrative denominator. Axis labels, callouts, warnings, headings, and
    other renderer-facing prose remain in scope.
    """
    if isinstance(value, str):
        return [] if structured_rows else [value]
    if isinstance(value, dict):
        result: list[str] = []
        for key, child in value.items():
            # These fields are renderer/control metadata or duplicated
            # structured evidence, not prose exposed to an audience.  Titles,
            # callouts, warnings, axis labels, takeaways and node labels remain
            # in scope.  Raw evidence rows are assessed through semantic and
            # source-lineage QA rather than the localization denominator.
            if (
                key in {
                    "image_path", "source_file", "source_id", "label_field",
                    "sort", "scale", "top_label_policy", "fixed", "random",
                    "prediction", "pooled", "claim_id", "status", "id",
                    "role", "source", "target", "sequential",
                }
                or key.endswith("_field")
            ):
                continue
            result.extend(audience_payload_strings(child, structured_rows=structured_rows or key == "rows"))
        return result
    if isinstance(value, list):
        return [
            item
            for child in value
            for item in audience_payload_strings(child, structured_rows=structured_rows)
        ]
    return []


def numeric_tokens(value: str) -> list[float]:
    return [float(item.replace(",", "")) for item in NUMERIC_RE.findall(value or "")]


def normalized_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def expected_value_matches(expected: str, observed: str) -> bool:
    if normalized_text(expected) in {"none", "derived from canonical table"}:
        return True
    expected_numbers = numeric_tokens(expected)
    observed_numbers = numeric_tokens(observed)
    if expected_numbers:
        return all(
            any(abs(left - right) <= max(1e-8, abs(left) * 1e-6) for right in observed_numbers)
            for left in expected_numbers
        )
    return normalized_text(expected) in normalized_text(observed)


def canonical_source_matches(expected_source: str, binding: dict[str, Any]) -> bool:
    """Match either an exact source filename or a declared source category.

    Ground-truth contracts normally name the canonical file.  Some fixtures
    intentionally name a human-readable category such as ``protocol
    materials`` instead.  Category matching remains fail-closed: after
    removing generic descriptor words, every remaining token must occur in the
    registered file/path/location metadata.  No claim values are consulted.
    """
    expected = str(expected_source or "").strip()
    if not expected:
        return False
    if Path(str(binding.get("source_file", ""))).name.casefold() == expected.casefold():
        return True
    generic_words = {"material", "materials", "source", "sources", "document", "documents", "file", "files"}
    tokens = [
        token for token in re.findall(r"[a-z0-9]+", expected.casefold())
        if token not in generic_words
    ]
    if not tokens:
        return False
    actual = " ".join(str(binding.get(key, "")) for key in (
        "source_file", "source_location", "short_label_zh", "canonical_status",
    )).casefold()
    return all(token in actual for token in tokens)


def text_intersection(first: dict[str, Any], second: dict[str, Any]) -> tuple[float, float]:
    left = max(float(first["bound_left"]), float(second["bound_left"]))
    top = max(float(first["bound_top"]), float(second["bound_top"]))
    right = min(
        float(first["bound_left"]) + float(first["bound_width"]),
        float(second["bound_left"]) + float(second["bound_width"]),
    )
    bottom = min(
        float(first["bound_top"]) + float(first["bound_height"]),
        float(second["bound_top"]) + float(second["bound_height"]),
    )
    area = max(0.0, right - left) * max(0.0, bottom - top)
    first_area = max(1.0, float(first["bound_width"]) * float(first["bound_height"]))
    second_area = max(1.0, float(second["bound_width"]) * float(second["bound_height"]))
    return area, area / min(first_area, second_area)


def is_footer_or_page_number(shape: dict[str, Any]) -> bool:
    name = str(shape.get("name", "")).casefold()
    return "footer" in name or "slide number" in name or "页码" in name


def contact_sheet(images: list[Path], output: Path, code: str) -> None:
    if not images:
        raise BlockedError(f"No PowerPoint PNGs available for contact sheet: {output}")
    thumbs: list[Image.Image] = []
    for index, image_path in enumerate(images, 1):
        with Image.open(image_path) as source:
            image = source.convert("RGB")
        image.thumbnail((390, 220))
        canvas = Image.new("RGB", (404, 250), "white")
        canvas.paste(image, ((404 - image.width) // 2, 6))
        ImageDraw.Draw(canvas).text((12, 229), f"S{index:03d}", fill=(36, 55, 76))
        thumbs.append(canvas)
    cols = 3
    rows = math.ceil(len(thumbs) / cols)
    sheet = Image.new("RGB", (cols * 404, rows * 250 + 42), (238, 244, 247))
    ImageDraw.Draw(sheet).text((16, 14), f"Deck {code} - blinded visual review", fill=(23, 50, 77))
    for index, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((index % cols) * 404, 42 + (index // cols) * 250))
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)


def dominant_foreground_metrics(image_path: Path) -> dict[str, float]:
    with Image.open(image_path) as source:
        image = source.convert("RGB").resize((400, 225))
    quantized = image.quantize(colors=24).convert("RGB")
    counts = Counter(quantized.getdata())
    background = counts.most_common(1)[0][0]
    pixels = list(image.getdata())
    width, height = image.size
    coordinates: list[tuple[int, int]] = []
    grid_counts = [0] * 72
    lower_count = 0
    for index, pixel in enumerate(pixels):
        distance = math.sqrt(sum((int(pixel[channel]) - int(background[channel])) ** 2 for channel in range(3)))
        if distance <= 24:
            continue
        x, y = index % width, index // width
        coordinates.append((x, y))
        cell = min(11, x * 12 // width) + 12 * min(5, y * 6 // height)
        grid_counts[cell] += 1
        if y >= height // 2:
            lower_count += 1
    total = width * height
    coverage = len(coordinates) / total
    if coordinates:
        centroid_x = sum(x for x, _ in coordinates) / len(coordinates) / width
        centroid_y = sum(y for _, y in coordinates) / len(coordinates) / height
    else:
        centroid_x = centroid_y = 0.5
    cell_area = total / 72
    occupied = sum(count / cell_area >= 0.015 for count in grid_counts) / 72
    lower_emptiness = 1 - lower_count / (total / 2)
    density_score = max(0.0, 100.0 - abs(coverage - 0.24) / 0.24 * 70.0)
    centre_distance = math.sqrt(((centroid_x - 0.50) / 0.50) ** 2 + ((centroid_y - 0.52) / 0.52) ** 2)
    balance_score = max(0.0, 100.0 - centre_distance * 85.0)
    return {
        "content_coverage": coverage,
        "whitespace_ratio": 1 - coverage,
        "foreground_centroid_x": centroid_x,
        "foreground_centroid_y": centroid_y,
        "grid_occupancy": occupied,
        "lower_half_emptiness": lower_emptiness,
        "perceptual_density_score": density_score,
        "optical_balance_score": balance_score,
    }


def pptx_typography(pptx: Path) -> tuple[dict[int, list[dict[str, Any]]], set[str]]:
    slides: dict[int, list[dict[str, Any]]] = defaultdict(list)
    typefaces: set[str] = set()
    with zipfile.ZipFile(pptx) as archive:
        for member in archive.namelist():
            if member.startswith("ppt/theme/") and member.endswith(".xml"):
                root = ET.fromstring(archive.read(member))
                for element in root.iter():
                    face = element.attrib.get("typeface")
                    if face:
                        typefaces.add(face)
        slide_members = [
            member for member in archive.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", member)
        ]
        slide_members.sort(key=lambda item: int(re.search(r"(\d+)", Path(item).stem).group(1)))
        for slide_no, member in enumerate(slide_members, 1):
            root = ET.fromstring(archive.read(member))
            for shape in root.findall(".//p:sp", NS):
                c_nv_pr = shape.find("./p:nvSpPr/p:cNvPr", NS)
                name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
                text = "".join(node.text or "" for node in shape.findall(".//a:t", NS))
                sizes: list[float] = []
                for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
                    for prop in shape.findall(f".//{tag}", NS):
                        if prop.attrib.get("sz"):
                            sizes.append(int(prop.attrib["sz"]) / 100)
                        face = prop.attrib.get("typeface")
                        if face:
                            typefaces.add(face)
                if text:
                    slides[slide_no].append({"name": name, "text": text, "sizes": sizes})
    return slides, typefaces


def classify_typography(
    project: str, pptx: Path, slide_specs: list[dict[str, Any]]
) -> tuple[list[dict[str, Any]], set[str]]:
    shapes_by_slide, typefaces = pptx_typography(pptx)
    records: list[dict[str, Any]] = []
    for index, spec in enumerate(slide_specs, 1):
        title_sizes: list[float] = []
        body_sizes: list[float] = []
        chart_detail_sizes: list[float] = []
        footer_sizes: list[float] = []
        title_under = body_under = chart_under = detail_under = footer_bad = 0
        for shape in shapes_by_slide.get(index, []):
            name = shape["name"].casefold()
            shape_text = normalized_text(shape.get("text", ""))
            sizes = shape["sizes"]
            if not sizes:
                continue
            if shape_text == normalized_text(spec.get("title", "")) or ":title" in name or "deck-title" in name:
                title_sizes.extend(sizes)
                title_under += sum(size < 30 for size in sizes)
            elif ":footer" in name or shape_text.startswith(("来源：", "来源:", "source:")):
                footer_sizes.extend(sizes)
                footer_bad += sum(not (9.8 <= size <= 11.2) for size in sizes)
            elif (
                ":section" in name
                or "slide number" in name
                or shape_text in {"问题", "方法", "结果", "局限性", "结论", "计划", "附录"}
            ):
                continue
            elif all(size <= 16.1 for size in sizes):
                chart_detail_sizes.extend(sizes)
                chart_under += sum(size < 15 for size in sizes)
                # 15 pt is a chart label; 16 pt is a justified card/detail label.
                detail_under += sum(15 <= size < 16 for size in sizes if "chart" not in name)
            else:
                body_sizes.extend(sizes)
                body_under += sum(size < 18 for size in sizes)
        passed = not any((title_under, body_under, chart_under, footer_bad))
        records.append({
            "project_key": project,
            "slide_number": index,
            "slide_id": spec["slide_id"],
            "title_min_pt": min(title_sizes) if title_sizes else "",
            "body_min_pt": min(body_sizes) if body_sizes else "",
            "chart_detail_min_pt": min(chart_detail_sizes) if chart_detail_sizes else "",
            "footer_min_pt": min(footer_sizes) if footer_sizes else "",
            "title_under_30": title_under,
            "body_under_18_unjustified": body_under,
            "chart_under_15": chart_under,
            "card_detail_under_16": detail_under,
            "footer_outside_10_11": footer_bad,
            "declared_typefaces": "|".join(sorted(typefaces)),
            "passed": passed,
        })
    return records, typefaces


def language_metrics(project: str, slide: dict[str, Any]) -> dict[str, Any]:
    strings = [slide.get("title", ""), slide.get("narrative_purpose", ""), slide.get("key_message", "")]
    for visual in slide.get("visual_specs", []):
        strings.extend(audience_payload_strings(visual.get("payload", {})))
    text = " ".join(strings)
    scrubbed = text
    for term in sorted(ALLOWED_LATIN_TERMS, key=len, reverse=True):
        scrubbed = re.sub(re.escape(term), "", scrubbed, flags=re.I)
    cjk = len(re.findall(r"[\u3400-\u9fff]", scrubbed))
    latin = len(re.findall(r"[A-Za-z]", scrubbed))
    coverage = cjk / max(1, cjk + latin)
    mojibake = bool(MOJIBAKE_RE.search(text) or "�" in text)
    abnormal_spacing = bool(ABNORMAL_SPACING_RE.search(text))
    title_has_cjk = bool(re.search(r"[\u3400-\u9fff]", slide.get("title", "")))
    passed = (
        slide.get("language") == "zh-CN"
        and coverage >= 0.90
        and not mojibake
        and not abnormal_spacing
        and title_has_cjk
    )
    return {
        "project_key": project,
        "slide_id": slide["slide_id"],
        "declared_language": slide.get("language", ""),
        "cjk_characters": cjk,
        "non_allowlisted_latin_characters": latin,
        "zh_cn_coverage": round(coverage, 4),
        "mojibake_detected": mojibake,
        "abnormal_spacing_detected": abnormal_spacing,
        "title_has_cjk": title_has_cjk,
        "passed": passed,
    }


def source_headers(source_path: Path) -> set[str]:
    if source_path.suffix.lower() != ".csv":
        return set()
    with source_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        return set(next(reader, []))


def binding_checks(
    project: str,
    kind: str,
    project_root: Path,
    graph: dict[str, Any],
    output: Path,
    geometry: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    source_registry = {item["source_id"]: item for item in graph["source_registry"]}
    header_cache: dict[str, set[str]] = {}
    rows_out: list[dict[str, Any]] = []
    footer_rows: list[dict[str, Any]] = []
    source_map = read_csv(output / "source_binding_map.csv")
    claim_map = read_csv(output / "claim_source_map.csv")
    slide_manifest = read_csv(output / "slide_manifest.csv")
    source_map_counter = Counter(
        (
            row["slide_id"], row["visual_ids"], row["source_id"], row["source_file"], row["source_location"],
            row["fields_used"], row["row_filter"], row["aggregation"], row["canonical_status"],
        )
        for row in source_map
    )
    expected_map_counter: Counter[tuple[str, ...]] = Counter()
    expected_claim_counter: Counter[tuple[str, ...]] = Counter()
    expected_slide_counter: Counter[tuple[str, ...]] = Counter()
    claim_map_counter = Counter(
        (
            row["claim_id"], row["slide_id"], row["visual_ids"], row["claim_text"],
            row["expected_value"], row["source_id"], row["source_file"],
            row["source_location"], row["fields_used"], row["wording_boundary"],
            row["canonical_status"],
        )
        for row in claim_map
    )
    slide_manifest_counter = Counter(
        (
            row["slide_number"], row["slide_id"], row["slide_role"], row["title"],
            row["language"], row["layout_family"], row["design_profile"],
            row["claim_ids"], row["visual_ids"], row["appendix_status"], row["review_flags"],
        )
        for row in slide_manifest
    )
    notes = (output / "speaker_notes.md").read_text(encoding="utf-8-sig")
    geometry_by_slide = {int(item["slide_number"]): item for item in geometry["slides"]}

    for slide_number, slide in enumerate(graph["slide_specs"], 1):
        visual_ids = "|".join(visual["visual_id"] for visual in slide.get("visual_specs", []))
        def visual_ids_for_binding(binding: dict[str, Any]) -> str:
            matches: list[str] = []
            expected_fields = set(binding.get("fields_used", []))
            for visual in slide.get("visual_specs", []):
                for visual_binding in visual.get("source_bindings", []):
                    if visual_binding.get("source_id") != binding.get("source_id"):
                        continue
                    actual_fields = set(visual_binding.get("fields_used", []))
                    if (not expected_fields and not actual_fields) or expected_fields <= actual_fields:
                        matches.append(visual["visual_id"])
                        break
            return "|".join(matches)
        expected_slide_counter[(
            str(slide_number), slide["slide_id"], slide["slide_role"], slide["title"],
            slide["language"], slide["layout_family"], slide["design_profile"],
            "|".join(item["claim_id"] for item in slide.get("claim_bindings", [])),
            visual_ids, slide["appendix_status"], "|".join(slide.get("review_flags", [])),
        )] += 1
        for claim in slide.get("claim_bindings", []):
            for binding in claim.get("source_bindings", []):
                expected_claim_counter[(
                    claim["claim_id"], slide["slide_id"], visual_ids_for_binding(binding), claim["claim_text"],
                    claim["expected_value"], binding["source_id"], binding["source_file"],
                    binding["source_location"], "|".join(binding["fields_used"]),
                    claim["wording_boundary"], claim["canonical_status"],
                )] += 1
        for binding in slide.get("source_bindings", []):
            fields_joined = "|".join(binding.get("fields_used", []))
            expected_map_counter[(
                slide["slide_id"], visual_ids_for_binding(binding), binding["source_id"], binding["source_file"],
                binding["source_location"], fields_joined, binding["row_filter"],
                binding["aggregation"], binding["canonical_status"],
            )] += 1
            registered = source_registry.get(binding.get("source_id"))
            source_match = bool(registered and registered.get("source_file") == binding.get("source_file"))
            path = project_root / binding.get("source_file", "")
            if str(path) not in header_cache:
                header_cache[str(path)] = source_headers(path) if path.is_file() else set()
            headers = header_cache[str(path)]
            fields = set(binding.get("fields_used", []))
            field_valid = bool(path.is_file()) and (
                bool(fields & RESERVED_FIELDS)
                or not headers
                or fields <= headers
            )
            passed = all([
                source_match, path.is_file(), field_valid,
                bool(binding.get("source_location")), bool(binding.get("row_filter")),
                binding.get("aggregation") is not None,
                bool(binding.get("canonical_status")),
            ])
            rows_out.append({
                "project_key": project,
                "slide_id": slide["slide_id"],
                "visual_id": "",
                "check_type": "slide_binding",
                "source_id": binding.get("source_id", ""),
                "source_file": binding.get("source_file", ""),
                "source_location": binding.get("source_location", ""),
                "fields_used": fields_joined,
                "row_filter": binding.get("row_filter", ""),
                "aggregation": binding.get("aggregation", ""),
                "canonical_status": binding.get("canonical_status", ""),
                "expected_source": "",
                "passed": passed,
                "details": "binding resolves to registered source and declared fields" if passed else "source, path, field, or binding metadata mismatch",
            })

        expected_labels = []
        for binding in slide.get("source_bindings", []):
            label = binding.get("short_label_zh", "输入材料")
            if label not in expected_labels:
                expected_labels.append(label)
        rendered_slide = geometry_by_slide.get(slide_number, {})
        rendered_height = float(rendered_slide.get("height", geometry.get("slide_height", 540)))
        footer_shapes = []
        for shape in rendered_slide.get("shapes", []):
            name = str(shape.get("name", "")).casefold()
            text_value = str(shape.get("text", "")).strip()
            top = shape.get("bound_top") if shape.get("bound_top") is not None else shape.get("top")
            is_source_text = normalized_text(text_value).startswith(("来源：", "来源:", "source:"))
            is_footer_position = (
                shape.get("has_text") and top is not None
                and float(top) >= rendered_height * (7.16 / 7.5)
                and "slide number" not in name
            )
            if "footer" in name or is_source_text or is_footer_position:
                footer_shapes.append(shape)
        footer_text = " ".join(str(shape.get("text", "")) for shape in footer_shapes)
        visible_labels = expected_labels[:3]
        footer_match = (not visible_labels and not footer_text) or all(label in footer_text for label in visible_labels)
        section_match = re.search(
            rf"(?ms)^##\s+{re.escape(slide['slide_id'])}\b.*?(?=^##\s+|\Z)",
            notes,
        )
        notes_section = section_match.group(0) if section_match else ""
        notes_match = bool(section_match) and all(
            binding["source_id"] in notes_section for binding in slide.get("source_bindings", [])
        )
        footer_rows.append({
            "project_key": project,
            "slide_id": slide["slide_id"],
            "expected_short_labels": "|".join(visible_labels),
            "rendered_footer": footer_text,
            "footer_match": footer_match,
            "notes_match": notes_match,
            "passed": footer_match and notes_match,
        })

        for visual in slide.get("visual_specs", []):
            bound_fields = {field for binding in visual["source_bindings"] for field in binding.get("fields_used", [])}
            payload_rows = visual.get("payload", {}).get("rows", [])
            payload_fields = set(payload_rows[0]) if payload_rows and isinstance(payload_rows[0], dict) else set()
            payload_fields.update(nested_record_fields(visual.get("payload", {})))
            for role in VISUAL_FIELD_ROLES:
                field = visual.get(role)
                if not field:
                    continue
                derived_bases = DERIVED_FIELD_BASES.get(str(field), set())
                payload_declares_field = field in payload_fields or (
                    field == "bin" and {"bin_low", "bin_high"} <= payload_fields
                )
                lineage = field in bound_fields or (
                    payload_declares_field
                    and any(binding.get("aggregation", "") != "none" for binding in visual["source_bindings"])
                    and (
                        (bool(derived_bases) and derived_bases <= bound_fields)
                        or re.search(
                            r"derivation|semantic|synthesis|matrix|histogram|node/edge|time-position",
                            str(visual.get("aggregation", "")),
                            re.I,
                        ) is not None
                    )
                )
                rows_out.append({
                    "project_key": project,
                    "slide_id": slide["slide_id"],
                    "visual_id": visual["visual_id"],
                    "check_type": f"field_lineage:{role}",
                    "source_id": "|".join(binding["source_id"] for binding in visual["source_bindings"]),
                    "source_file": "|".join(binding["source_file"] for binding in visual["source_bindings"]),
                    "source_location": "|".join(binding["source_location"] for binding in visual["source_bindings"]),
                    "fields_used": "|".join(sorted(bound_fields)),
                    "row_filter": visual.get("row_filter", ""),
                    "aggregation": visual.get("aggregation", ""),
                    "canonical_status": "|".join(sorted({binding["canonical_status"] for binding in visual["source_bindings"]})),
                    "expected_source": "",
                    "passed": lineage,
                    "details": f"{role}={field}" + (" is explicitly bound" if lineage else " has no declared field-level lineage"),
                })

    flattened_visuals = [
        visual for slide in graph["slide_specs"] for visual in slide.get("visual_specs", [])
    ]
    for selector, expected_source, label in EXPECTED_SOURCE_RULES[kind]:
        visual_type, _, scientific_role = selector.partition(":")
        matched = [
            visual for visual in flattened_visuals
            if visual.get("visual_type") == visual_type
            and (not scientific_role or visual.get("scientific_role") == scientific_role)
        ]
        if not matched:
            rows_out.append({
                "project_key": project, "slide_id": "", "visual_id": "",
                "check_type": "regression_expected_source", "source_id": "", "source_file": "",
                "source_location": "", "fields_used": "", "row_filter": "", "aggregation": "",
                "canonical_status": "", "expected_source": expected_source, "passed": False,
                "details": f"No VisualSpec found for {label}",
            })
            continue
        for visual in matched:
            files = [Path(binding["source_file"]).name for binding in visual["source_bindings"]]
            if expected_source.endswith(".csv"):
                matched_source = expected_source in files
            else:
                matched_source = any(expected_source.casefold() in item.casefold() for item in files)
            rows_out.append({
                "project_key": project, "slide_id": visual.get("slide_id", ""),
                "visual_id": visual["visual_id"], "check_type": "regression_expected_source",
                "source_id": "|".join(binding["source_id"] for binding in visual["source_bindings"]),
                "source_file": "|".join(binding["source_file"] for binding in visual["source_bindings"]),
                "source_location": "|".join(binding["source_location"] for binding in visual["source_bindings"]),
                "fields_used": "|".join(sorted({field for binding in visual["source_bindings"] for field in binding["fields_used"]})),
                "row_filter": visual.get("row_filter", ""), "aggregation": visual.get("aggregation", ""),
                "canonical_status": "|".join(sorted({binding["canonical_status"] for binding in visual["source_bindings"]})),
                "expected_source": expected_source, "passed": matched_source,
                "details": f"{label} must bind {expected_source}",
            })

    map_match = expected_map_counter == source_map_counter
    rows_out.append({
        "project_key": project, "slide_id": "ALL", "visual_id": "",
        "check_type": "derived_source_map_identity", "source_id": "", "source_file": "",
        "source_location": "", "fields_used": "", "row_filter": "", "aggregation": "",
        "canonical_status": "", "expected_source": "SlideSpec", "passed": map_match,
        "details": "source_binding_map.csv exactly matches SlideSpec-derived bindings" if map_match else "source map differs from SlideSpec",
    })
    claim_map_match = expected_claim_counter == claim_map_counter
    rows_out.append({
        "project_key": project, "slide_id": "ALL", "visual_id": "",
        "check_type": "derived_claim_source_map_identity", "source_id": "", "source_file": "",
        "source_location": "", "fields_used": "", "row_filter": "", "aggregation": "",
        "canonical_status": "", "expected_source": "SlideSpec.claim_bindings", "passed": claim_map_match,
        "details": "claim_source_map.csv exactly matches SlideSpec-derived claim bindings" if claim_map_match else "claim-source map differs from SlideSpec",
    })
    slide_manifest_match = expected_slide_counter == slide_manifest_counter
    rows_out.append({
        "project_key": project, "slide_id": "ALL", "visual_id": "",
        "check_type": "derived_slide_manifest_identity", "source_id": "", "source_file": "",
        "source_location": "", "fields_used": "", "row_filter": "", "aggregation": "",
        "canonical_status": "", "expected_source": "SlideSpec", "passed": slide_manifest_match,
        "details": "slide_manifest.csv exactly matches SlideSpec" if slide_manifest_match else "slide manifest differs from SlideSpec",
    })
    notes_claims_match = all(
        claim["claim_id"] in notes
        for slide in graph["slide_specs"] for claim in slide.get("claim_bindings", [])
    )
    notes_review_match = all(
        item["conflict_id"] in notes
        for slide in graph["slide_specs"] for item in slide.get("conflict_bindings", [])
    ) and all(
        item["marker"] in notes
        for slide in graph["slide_specs"] for item in slide.get("unresolved_bindings", [])
    )
    rows_out.append({
        "project_key": project, "slide_id": "ALL", "visual_id": "",
        "check_type": "derived_notes_identity", "source_id": "", "source_file": "",
        "source_location": "", "fields_used": "", "row_filter": "", "aggregation": "",
        "canonical_status": "", "expected_source": "SlideSpec bindings",
        "passed": notes_claims_match and notes_review_match,
        "details": "speaker notes retain all claim, conflict and unresolved bindings",
    })
    return rows_out, footer_rows


def add_semantic(
    records: list[dict[str, Any]], project: str, slide_id: str, visual: dict[str, Any],
    check_id: str, passed: bool, details: str, severity: str = "error",
) -> None:
    records.append({
        "project_key": project, "slide_id": slide_id, "visual_id": visual.get("visual_id", ""),
        "visual_type": visual.get("visual_type", ""), "check_id": check_id,
        "severity": severity, "passed": passed, "details": details,
    })


def float_or_none(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def semantic_checks(project: str, kind: str, slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for slide in slides:
        for visual in slide.get("visual_specs", []):
            visual_type = visual["visual_type"]
            payload = visual.get("payload", {})
            data_rows = payload.get("rows", [])
            expected = int(visual.get("expected_row_count", 0))
            rendered_value = visual.get("rendered_row_count")
            rendered = int(rendered_value) if rendered_value is not None else -1
            omitted = visual.get("omitted_rows", [])
            add_semantic(records, project, slide["slide_id"], visual, "expected_vs_rendered_rows",
                         expected == rendered or bool(omitted), f"expected={expected}; rendered={rendered}; omitted={len(omitted)}")

            if visual_type in {"forest_plot", "subgroup_forest"}:
                required = all(visual.get(field) not in {None, ""} for field in ("x_field", "ci_low_field", "ci_high_field", "label_field", "reference_value"))
                scale = payload.get("scale", "linear")
                values_ok = True
                for row in data_rows:
                    estimate = float_or_none(row.get(visual["x_field"]))
                    low = float_or_none(row.get(visual["ci_low_field"]))
                    high = float_or_none(row.get(visual["ci_high_field"]))
                    if estimate is None or low is None or high is None or not (low < estimate < high):
                        values_ok = False
                    if scale == "log" and any(value is not None and value <= 0 for value in (low, estimate, high)):
                        values_ok = False
                axis_ok = bool(payload.get("axis_label")) and visual.get("unit") not in {None, ""}
                add_semantic(records, project, slide["slide_id"], visual, "forest_axis_reference_ci",
                             required and values_ok and axis_ok, f"scale={scale}; rows={len(data_rows)}; axis={payload.get('axis_label', '')}")
                if visual_type == "subgroup_forest":
                    interaction_ok = all(row.get("interaction_p") not in {None, ""} for row in data_rows)
                    add_semantic(records, project, slide["slide_id"], visual, "subgroup_interaction_p",
                                 interaction_ok, "interaction P must be retained for every subgroup row")

            if visual_type == "love_plot":
                threshold = float_or_none(payload.get("threshold"))
                ordered = [float_or_none(row.get("absolute_smd_unweighted")) for row in data_rows]
                ordered_ok = all(value is not None for value in ordered) and ordered == sorted(ordered, reverse=True)
                legend_fields = {visual.get("x_field"), visual.get("y_field")} == {"absolute_smd_unweighted", "absolute_smd_weighted"}
                add_semantic(records, project, slide["slide_id"], visual, "love_plot_contract",
                             threshold is not None and ordered_ok and legend_fields,
                             f"threshold={threshold}; covariates={len(data_rows)}; before/after fields={legend_fields}")

            if visual_type == "grouped_composition":
                labels = {str(row.get(visual.get("label_field", ""), "")) for row in data_rows}
                groups = {str(row.get(visual.get("group_field", ""), "")) for row in data_rows}
                full_pairs = len(data_rows) == len(labels) * len(groups) and len(groups) >= 2
                add_semantic(records, project, slide["slide_id"], visual, "grouped_composition_complete",
                             full_pairs and bool(visual.get("y_field")), f"entities={len(labels)}; groups={len(groups)}; rows={len(data_rows)}")

            if visual_type == "diverging_enrichment":
                facets = {str(row.get(visual.get("facet_field", ""), "")) for row in data_rows}
                scores = [float_or_none(row.get(visual.get("x_field", ""))) for row in data_rows]
                sign_ok = any(value is not None and value < 0 for value in scores) and any(value is not None and value > 0 for value in scores)
                fdr_ok = all(row.get("FDR") not in {None, ""} for row in data_rows)
                add_semantic(records, project, slide["slide_id"], visual, "diverging_enrichment_facet_sign",
                             len(facets) >= 2 and sign_ok and fdr_ok and visual.get("reference_value") == 0,
                             f"facets={sorted(facets)}; zero_line={visual.get('reference_value')}; FDR={fdr_ok}")

            if visual_type == "communication_network":
                edges = {
                    (row.get(visual.get("edge_source_field", "")), row.get(visual.get("edge_target_field", "")), row.get(visual.get("label_field", "")))
                    for row in data_rows
                }
                passed = visual.get("sequential") is False and len(edges) == len(data_rows) and all(all(item) for item in edges)
                add_semantic(records, project, slide["slide_id"], visual, "independent_network_not_sequence",
                             passed, f"sequential={visual.get('sequential')}; independent_edges={len(edges)}")

            if visual_type == "risk_of_bias_matrix":
                studies = {row.get("study_id") or row.get("study_label") for row in data_rows}
                domains = {row.get("domain") for row in data_rows}
                complete = len(data_rows) == len(studies) * len(domains) and len(studies) > 1 and len(domains) > 1
                add_semantic(records, project, slide["slide_id"], visual, "true_rob_matrix",
                             complete, f"studies={len(studies)}; domains={len(domains)}; cells={len(data_rows)}")

            if visual_type == "dag":
                nodes = {node.get("id"): node for node in payload.get("nodes", [])}
                edges = {(edge.get("source"), edge.get("target")) for edge in payload.get("edges", [])}
                exposure = {key for key, node in nodes.items() if node.get("role") == "exposure"}
                outcome = {key for key, node in nodes.items() if node.get("role") == "outcome"}
                confounders = {key for key, node in nodes.items() if node.get("role") == "confounder"}
                confounding_ok = all(
                    any((confounder, item) in edges for item in exposure)
                    and any((confounder, item) in edges for item in outcome)
                    for confounder in confounders
                )
                exposure_outcome = any((left, right) in edges for left in exposure for right in outcome)
                passed = bool(confounders and exposure and outcome and confounding_ok and exposure_outcome and payload.get("adjustment_set") and visual.get("sequential") is False)
                add_semantic(records, project, slide["slide_id"], visual, "true_dag",
                             passed, f"nodes={len(nodes)}; edges={len(edges)}; confounders={len(confounders)}; sequential={visual.get('sequential')}")

            if visual_type == "risk_matrix":
                x_field, y_field = visual.get("x_field"), visual.get("y_field")
                coordinates = {(row.get(x_field), row.get(y_field)) for row in data_rows}
                passed = bool(x_field and y_field and x_field != y_field and coordinates and all(all(item) for item in coordinates))
                add_semantic(records, project, slide["slide_id"], visual, "true_risk_matrix",
                             passed, f"x={x_field}; y={y_field}; placed_items={len(data_rows)}")

            if visual_type == "timeline_gantt":
                start = visual.get("time_field")
                end = visual.get("end_time_field")
                dates_ok = bool(start and end) and all(row.get(start) and row.get(end) for row in data_rows)
                spans = {(row.get(start), row.get(end)) for row in data_rows}
                add_semantic(records, project, slide["slide_id"], visual, "date_scaled_gantt",
                             dates_ok and len(spans) > 1, f"time={start}; end={end}; distinct_spans={len(spans)}")

            if visual_type in {"source_figure", "umap_figure"}:
                callout = normalized_text(payload.get("callout"))
                add_semantic(records, project, slide["slide_id"], visual, "source_figure_callout",
                             bool(callout), "registered figure must have an explanatory callout")

            if visual_type == "pooled_evidence_panel":
                prediction = payload.get("prediction", {})
                estimand = normalized_text(prediction.get("estimand"))
                add_semantic(records, project, slide["slide_id"], visual, "prediction_interval_not_ci",
                             "prediction" in estimand, f"prediction estimand label={prediction.get('estimand', '')}")

    titles = [normalized_text(slide.get("title")) for slide in slides]
    duplicate_titles = {title for title, count in Counter(titles).items() if title and count > 1}
    dummy = {"visual_id": "", "visual_type": "slide"}
    add_semantic(records, project, "ALL", dummy, "duplicate_slide_semantics", not duplicate_titles,
                 "no duplicate titles" if not duplicate_titles else f"duplicate titles={sorted(duplicate_titles)}")

    if kind == "protocol":
        forbidden_visuals = {"absolute_risk_comparison", "estimand_comparison", "forest_plot", "subgroup_forest", "grouped_composition"}
        result_roles = {"result", "primary_result", "supportive_result", "observed_result"}
        prohibited_phrases = re.compile(r"observed result|recruitment completed|事件率为|显著增加|显著降低|观察到", re.I)
        all_text = " ".join(flatten_strings(slides))
        fabricated = any(
            slide.get("slide_role") == "result"
            or any(visual.get("visual_type") in forbidden_visuals or visual.get("scientific_role") in result_roles for visual in slide.get("visual_specs", []))
            for slide in slides
        ) or bool(prohibited_phrases.search(all_text))
        add_semantic(records, project, "ALL", dummy, "protocol_no_results", not fabricated,
                     "protocol contains planning definitions only" if not fabricated else "result-like content detected")
    return records


def omission_checks(project: str, slides: list[dict[str, Any]], coverage_rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for slide in slides:
        for visual in slide.get("visual_specs", []):
            expected = int(visual.get("expected_row_count", 0))
            rendered_value = visual.get("rendered_row_count")
            rendered = int(rendered_value) if rendered_value is not None else -1
            omitted = visual.get("omitted_rows", [])
            reason = visual.get("omission_reason", "")
            mismatch = expected != rendered
            silent = mismatch and not (omitted and reason)
            if mismatch or omitted:
                records.append({
                    "project_key": project, "slide_id": slide["slide_id"], "visual_id": visual["visual_id"],
                    "item_type": "visual_rows", "expected_count": expected, "rendered_count": rendered,
                    "omitted_rows": json.dumps(omitted, ensure_ascii=False), "reason": reason,
                    "silent": silent, "passed": not silent,
                })
    for item in coverage_rows:
        if as_bool(item.get("omitted")):
            records.append({
                "project_key": project, "slide_id": item.get("rendered_slide", ""), "visual_id": "",
                "item_type": "must_include", "expected_count": 1, "rendered_count": 0,
                "omitted_rows": item.get("must_include_item", ""), "reason": item.get("omission_reason", ""),
                "silent": not bool(item.get("omission_reason")), "passed": False,
            })
    return records


def generic_title(title: str, role: str) -> tuple[bool, str]:
    stripped = title.strip()
    if role in {"cover", "appendix", "section_divider"}:
        return False, "role exemption"
    if GENERIC_TITLE_RE.fullmatch(stripped):
        return True, "generic section label"
    meaningful = len(re.findall(r"[\u3400-\u9fffA-Za-z0-9]", stripped))
    if meaningful < 8:
        return True, "too short to express a conclusion or specific purpose"
    return False, "specific title"


def layout_metrics(project: str, slides: list[dict[str, Any]]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    families = [slide["layout_family"] for slide in slides]
    max_consecutive = 0
    current = 0
    previous = None
    for family in families:
        current = current + 1 if family == previous else 1
        previous = family
        max_consecutive = max(max_consecutive, current)
    card_grid_count = sum(as_bool(slide.get("card_grid")) for slide in slides)
    text_only_count = sum(
        as_bool(slide.get("text_only"))
        and slide.get("slide_role") not in {"cover", "appendix", "section_divider"}
        for slide in slides
    )
    title_rows: list[dict[str, Any]] = []
    generic_count = 0
    for slide in slides:
        generic, reason = generic_title(slide["title"], slide["slide_role"])
        generic_count += generic
        title_rows.append({
            "project_key": project, "slide_id": slide["slide_id"], "title": slide["title"],
            "generic": generic, "reason": reason,
        })
    hero_visual_types = {
        "absolute_risk_comparison", "estimand_comparison", "forest_plot",
        "subgroup_forest", "grouped_composition", "faceted_volcano",
        "diverging_enrichment", "communication_network", "pooled_evidence_panel",
        "source_figure", "umap_figure",
    }
    result_without_hero = sum(
        slide.get("slide_role") == "result"
        and not any(visual.get("visual_type") in hero_visual_types for visual in slide.get("visual_specs", []))
        for slide in slides
    )
    count = len(slides)
    record = {
        "project_key": project,
        "slide_count": count,
        "layout_family_count": len(set(families)),
        "layout_families": "|".join(sorted(set(families))),
        "card_grid_count": card_grid_count,
        "card_grid_ratio": round(card_grid_count / max(1, count), 4),
        "text_only_count": text_only_count,
        "text_only_ratio": round(text_only_count / max(1, count), 4),
        "max_consecutive_same_layout": max_consecutive,
        "generic_title_count": generic_count,
        "generic_title_rate": round(generic_count / max(1, count), 4),
        "result_slide_without_hero": result_without_hero,
        "passed": (
            len(set(families)) >= 5 and card_grid_count / max(1, count) <= 0.30
            and text_only_count / max(1, count) <= 0.15 and max_consecutive <= 2
            and generic_count / max(1, count) < 0.15 and result_without_hero == 0
        ),
    }
    return record, title_rows


def actual_geometry(
    project: str,
    geometry: dict[str, Any],
    slides: list[dict[str, Any]],
    pp_images: list[Path],
    lo_images: list[Path],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    geometry_rows: list[dict[str, Any]] = []
    visual_rows: list[dict[str, Any]] = []
    for index, slide_geometry in enumerate(geometry.get("slides", []), 1):
        spec = slides[index - 1] if index <= len(slides) else {"slide_id": "", "slide_role": ""}
        shapes = slide_geometry.get("shapes", [])
        text_shapes = [
            shape for shape in shapes
            if shape.get("has_text") and shape.get("bound_left") is not None
        ]
        overlap_pairs = 0
        for first_index, first in enumerate(text_shapes):
            for second in text_shapes[first_index + 1:]:
                area, ratio = text_intersection(first, second)
                if area > 20 and ratio > 0.04:
                    overlap_pairs += 1
        clipping = sum(as_bool(shape.get("overflowing")) for shape in text_shapes)
        off_slide = sum(as_bool(shape.get("off_slide")) for shape in shapes)
        slide_width = float(slide_geometry.get("width", geometry.get("slide_width", 960)))
        slide_height = float(slide_geometry.get("height", geometry.get("slide_height", 540)))
        footer_y = slide_height * (7.16 / 7.5)
        footer_intrusion = 0
        for shape in shapes:
            if is_footer_or_page_number(shape):
                continue
            top = shape.get("bound_top") if shape.get("bound_top") is not None else shape.get("top")
            height = shape.get("bound_height") if shape.get("bound_height") is not None else shape.get("height")
            left = shape.get("left")
            width = shape.get("width")
            if top is None or height is None or left is None or width is None:
                continue
            top_value = float(top)
            height_value = float(height)
            left_value = float(left)
            width_value = float(width)
            # A canvas-sized, non-text object is a slide background, not body
            # content intruding into the footer.  This is especially relevant
            # for cover slides, whose full-slide color rectangle necessarily
            # spans the footer coordinate.  Smaller shapes, text, charts,
            # groups, tables and pictures retain the strict zero-tolerance
            # footer gate.
            is_full_slide_background = (
                not shape.get("has_text")
                and left_value <= 2
                and top_value <= 2
                and width_value >= slide_width * 0.95
                and height_value >= slide_height * 0.95
            )
            if is_full_slide_background:
                continue
            if top_value < footer_y and top_value + height_value > footer_y + 1:
                footer_intrusion += 1
        pp_image = pp_images[index - 1] if index <= len(pp_images) else None
        lo_image = lo_images[index - 1] if index <= len(lo_images) else None
        passed = not any((overlap_pairs, clipping, footer_intrusion, off_slide))
        geometry_rows.append({
            "project_key": project, "slide_number": index, "slide_id": spec.get("slide_id", ""),
            "text_overlap_pairs": overlap_pairs, "clipping": clipping,
            "footer_intrusion": footer_intrusion, "off_slide_shapes": off_slide,
            "powerpoint_png": str(pp_image or ""), "libreoffice_png": str(lo_image or ""),
            "passed": passed,
        })
        if pp_image is None:
            continue
        metrics = dominant_foreground_metrics(pp_image)
        width = slide_width
        content_shapes = []
        for shape in shapes:
            name = str(shape.get("name", "")).casefold()
            if any(token in name for token in ("title", "footer", "slide number", "section")):
                continue
            area = float(shape.get("width", 0)) * float(shape.get("height", 0))
            if 0 < area < width * slide_height * 0.85 and float(shape.get("top", 0)) < footer_y:
                content_shapes.append(area)
        anchor_ratio = max(content_shapes, default=0.0) / max(1.0, width * slide_height)
        is_hero_required = spec.get("slide_role") == "result"
        hero_score = min(100.0, anchor_ratio / (0.18 if is_hero_required else 0.12) * 100.0)
        chart_types = {
            "forest_plot", "subgroup_forest", "love_plot", "grouped_composition",
            "faceted_volcano", "diverging_enrichment", "risk_of_bias_matrix",
            "risk_matrix", "timeline_gantt", "assessment_timeline", "weight_histogram",
        }
        is_chart = any(visual.get("visual_type") in chart_types for visual in spec.get("visual_specs", []))
        readability = 100.0
        readability -= min(55.0, overlap_pairs * 25 + clipping * 30 + footer_intrusion * 30)
        if is_chart and metrics["content_coverage"] < 0.05:
            readability -= 30
        if is_chart and anchor_ratio < 0.08:
            readability -= 20
        warnings = []
        if spec.get("slide_role") not in {"cover", "section_divider", "appendix"} and metrics["content_coverage"] < 0.025:
            warnings.append("large_blank_region")
        if metrics["content_coverage"] > 0.82:
            warnings.append("unexpectedly_dense")
        if is_hero_required and anchor_ratio < 0.08:
            warnings.append("weak_hero_anchor")
        gross = bool(not passed or "large_blank_region" in warnings or "unexpectedly_dense" in warnings)
        visual_rows.append({
            "project_key": project, "slide_number": index, "slide_id": spec.get("slide_id", ""),
            **{key: round(value, 4) for key, value in metrics.items()},
            "largest_visual_anchor_ratio": round(anchor_ratio, 4),
            "hero_anchor_score": round(hero_score, 1),
            "chart_readability_score": round(max(0.0, readability), 1),
            "text_overlap_pairs": overlap_pairs, "clipping": clipping,
            "footer_intrusion": footer_intrusion, "off_slide_shapes": off_slide,
            "gross_visual_failure": gross, "warnings": "|".join(warnings),
        })
    return geometry_rows, visual_rows


def discover_inputs(project_root: Path) -> set[str]:
    candidates = [project_root / "README_PROJECT.md", project_root / "brief" / "presentation_brief.yaml"]
    input_root = project_root / "input"
    if input_root.is_dir():
        candidates.extend(sorted(input_root.rglob("*")))
    return {
        path.relative_to(project_root).as_posix()
        for path in candidates
        if path.is_file() and "expected_ground_truth" not in path.parts
    }


def input_integrity(project: str, project_root: Path, source_manifest: list[dict[str, str]]) -> list[dict[str, Any]]:
    registered_files = {row["source_file"] for row in source_manifest}
    discovered = discover_inputs(project_root)
    records: list[dict[str, Any]] = []
    by_path = {row["source_file"]: row for row in source_manifest}
    for relative in sorted(discovered | registered_files):
        row = by_path.get(relative)
        path = project_root / relative
        digest = hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else ""
        before = row.get("sha256", "") if row else ""
        gt_excluded = "expected_ground_truth" not in Path(relative).parts
        unchanged = bool(row and path.is_file() and digest == before)
        registered = relative in registered_files
        records.append({
            "project_key": project, "source_file": relative, "sha256_before": before,
            "sha256_after": digest, "unchanged": unchanged, "registered": registered,
            "ground_truth_excluded": gt_excluded,
            "passed": unchanged and registered and gt_excluded,
        })
    return records


def ground_truth_compare(
    project: str, output: Path, gt: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], int]:
    graph = read_json(output / "canonical_object_graph.json")
    claims = {item["claim_id"]: item for item in graph["canonical_evidence_graph"]["claims"]}
    expected_claims = read_csv(gt / "expected_claims.csv")
    expected_ids = {item["claim_id"] for item in expected_claims}
    claim_rows: list[dict[str, Any]] = []
    for expected in expected_claims:
        observed = claims.get(expected["claim_id"])
        source_match = bool(observed and any(
            canonical_source_matches(expected["canonical_source"], binding)
            for binding in observed.get("source_bindings", [])
        ))
        value_match = bool(observed and expected_value_matches(expected["expected_value"], observed.get("expected_value", "")))
        identity = observed is not None
        claim_rows.append({
            "project_key": project, "claim_id": expected["claim_id"],
            "identity_recalled": identity, "value_recalled": value_match,
            "canonical_source_match": source_match,
            "classification": "true_positive" if identity and value_match and source_match else "false_negative",
        })
    invented = len(set(claims) - expected_ids)

    conflicts = graph["canonical_evidence_graph"]["conflicts"]
    conflict_rows = []
    for expected in read_csv(gt / "expected_conflicts.csv"):
        expected_numbers = numeric_tokens(expected["conflicting_value"] + " " + expected["canonical_value"])
        found = any(
            all(any(abs(left - right) < 1e-8 for right in numeric_tokens(item["conflicting_value"] + " " + item["canonical_value"])) for left in expected_numbers)
            if expected_numbers else (
                normalized_text(expected["conflicting_value"]) == normalized_text(item["conflicting_value"])
                and normalized_text(expected["canonical_value"]) == normalized_text(item["canonical_value"])
            )
            for item in conflicts
        )
        conflict_rows.append({
            "project_key": project, "conflict_id": expected["conflict_id"], "recalled": found,
            "classification": "true_positive" if found else "false_negative",
        })

    unresolved = graph["canonical_evidence_graph"]["unresolved"]
    unresolved_rows = []
    for expected in read_csv(gt / "expected_unresolved_items.csv"):
        found = any(item["marker"] == expected["marker"] for item in unresolved)
        unresolved_rows.append({
            "project_key": project, "item_id": expected["item_id"], "marker": expected["marker"],
            "recalled": found, "classification": "true_positive" if found else "false_negative",
        })
    return claim_rows, conflict_rows, unresolved_rows, invented


def baseline_runtime() -> tuple[float | None, str]:
    candidates = sorted((REPO / "benchmark" / "v2_3" / "runs").glob("*/deck_index.csv"), reverse=True)
    for path in candidates:
        records = read_csv(path)
        values = [float(row["generation_seconds"]) for row in records if row.get("arm") == "B2"]
        if len(values) >= 4:
            return sum(values) / len(values), str(path)
    project_results = REPO / "benchmark" / "v2_3" / "project_results.csv"
    if project_results.is_file():
        values = [float(row["generation_seconds"]) for row in read_csv(project_results)]
        if values:
            return sum(values) / len(values), str(project_results)
    return None, ""


def review_form(
    output: Path, code: str,
) -> tuple[dict[str, Any], bool, float | None]:
    path = output / "manual_visual_review_form.csv"
    score_fields = [
        "scientific_clarity_20", "visual_semantic_correctness_20", "hierarchy_15",
        "composition_15", "chart_professionalism_10", "consistency_10",
        "readability_5", "presentation_readiness_5",
    ]
    maxima = [20, 20, 15, 15, 10, 10, 5, 5]
    if path.is_file():
        existing = read_csv(path)
        row = existing[0] if existing else {}
    else:
        row = {}
    aggregate = {
        "anonymous_deck_code": code,
        **{field: row.get(field, "") for field in score_fields},
        "total_100": row.get("total_100", ""),
        "reviewer": row.get("reviewer", ""),
        "review_date": row.get("review_date", ""),
        "comments": row.get("comments", ""),
    }
    scores: list[float] = []
    valid = True
    for field, maximum in zip(score_fields, maxima, strict=True):
        value = float_or_none(aggregate[field])
        if value is None or not (0 <= value <= maximum):
            valid = False
        else:
            scores.append(value)
    completed = valid and bool(aggregate["reviewer"]) and bool(aggregate["review_date"])
    total = sum(scores) if completed else None
    if total is not None:
        aggregate["total_100"] = round(total, 1)
    # Preserve existing human values; only materialize a missing form or anonymous code.
    write_csv(path, [aggregate], list(aggregate))
    return aggregate, completed, total


def run_tests(python: Path, audit: Path) -> tuple[bool, str]:
    try:
        completed = subprocess.run(
            [str(python), "-m", "unittest", "discover", "-s", "tests", "-v"],
            cwd=REPO, capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=420,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        log = f"Test runner failed: {exc}\n"
        (audit / "regression_test_raw.log").write_text(log, encoding="utf-8")
        return False, log
    log = completed.stdout + "\n" + completed.stderr
    (audit / "regression_test_raw.log").write_text(log, encoding="utf-8")
    return completed.returncode == 0, log


def write_architecture_reports(audit: Path) -> None:
    architecture_path = audit / "current_architecture_graph.md"
    if not architecture_path.exists():
        architecture_path.write_text(
        "# v2.4 actual architecture graph\n\n"
        "```text\n"
        "discover_project (SourceRegistry)\n"
        "  -> canonical_object_graph.canonical_evidence_graph\n"
        "  -> story_graph\n"
        "  -> slide_specs[]\n"
        "       -> visual_specs[] + field-level source_bindings\n"
        "       -> slide_manifest.csv / source_binding_map.csv / claim_source_map.csv\n"
        "       -> speaker_notes.md / source footer / review checklist\n"
        "  -> native PptxGenJS renderer\n"
        "  -> PowerPoint COM + LibreOffice render indexes\n"
        "  -> v2.4 finalizer (geometry, lineage, semantic, coverage, visual and scientific QA)\n"
        "  -> human visual review gate\n"
        "```\n\n"
        "`canonical_object_graph.json` is the sole project-state authority. Derived audience, notes, audit, and QA artifacts are compared back to its SlideSpec/VisualSpec objects.\n",
            encoding="utf-8",
        )
    lineage_path = audit / "data_lineage_map.md"
    if not lineage_path.exists():
        lineage_path.write_text(
        "# Data lineage map\n\n"
        "| Consumer | Authoritative path | Re-inference permitted |\n"
        "|---|---|---|\n"
        "| Audience title and key message | StoryGraph -> SlideSpec | No |\n"
        "| Chart rows and roles | CanonicalEvidenceGraph -> VisualSpec.payload / field roles | No |\n"
        "| Chart/source caption | VisualSpec.source_bindings.short_label_zh | No |\n"
        "| Speaker notes | SlideSpec claim/source/conflict/unresolved bindings | No |\n"
        "| claim_source_map.csv | SlideSpec.claim_bindings | No |\n"
        "| source_binding_map.csv | SlideSpec.source_bindings | No |\n"
        "| QA row coverage | VisualSpec expected/rendered/omitted rows | No |\n"
        "| Human review | rendered PNG contact sheet plus the same manifests | No |\n",
            encoding="utf-8",
        )
    duplicate_rows = [
        {"artifact": "canonical_object_graph.json", "authority": "canonical", "derivation": "SourceRegistry + CanonicalEvidenceGraph + StoryGraph + SlideSpec + VisualSpec", "independent_inference_allowed": "false", "status": "single authority"},
        {"artifact": "slide_spec.json", "authority": "derived view", "derivation": "canonical_object_graph.slide_specs", "independent_inference_allowed": "false", "status": "identity checked"},
        {"artifact": "visual_spec.json", "authority": "derived view", "derivation": "flatten(canonical_object_graph.slide_specs[].visual_specs)", "independent_inference_allowed": "false", "status": "identity checked"},
        {"artifact": "source_binding_map.csv", "authority": "derived", "derivation": "SlideSpec.source_bindings", "independent_inference_allowed": "false", "status": "counter identity checked"},
        {"artifact": "claim_source_map.csv", "authority": "derived", "derivation": "SlideSpec.claim_bindings", "independent_inference_allowed": "false", "status": "identity checked"},
        {"artifact": "speaker_notes.md", "authority": "derived", "derivation": "SlideSpec bindings", "independent_inference_allowed": "false", "status": "source and claim presence checked"},
        {"artifact": "source footer", "authority": "derived", "derivation": "SlideSpec.source_bindings.short_label_zh", "independent_inference_allowed": "false", "status": "PowerPoint text checked"},
    ]
    duplicated_path = audit / "duplicated_state_inventory.csv"
    if not duplicated_path.exists():
        write_csv(duplicated_path, duplicate_rows)
    module_rows = [
        {"module": "scripts/v2_4/model.py", "responsibility": "source discovery, canonical identities, binding and schema validation", "input_contract": "isolated project root", "output_contract": "SourceRegistry and validated object fragments", "authority": "canonical construction"},
        {"module": "scripts/v2_4/run_v2_4.py", "responsibility": "evidence graph, story, SlideSpec/VisualSpec and derived manifests", "input_contract": "SourceRegistry + brief", "output_contract": "canonical_object_graph.json and PPTX render request", "authority": "canonical orchestration"},
        {"module": "scripts/v2_4/render_v2_4_deck.mjs", "responsibility": "design system and semantic chart rendering", "input_contract": "canonical_object_graph.json", "output_contract": "editable PPTX + object manifest", "authority": "render only"},
        {"module": "scripts/v2_4/finalize_v2_4.py", "responsibility": "post-generation scientific, lineage, geometry, visual and human gate", "input_contract": "run root + render indexes + isolated suite", "output_contract": "audit/v2_4 and benchmark/v2_4 QA artifacts", "authority": "validation only"},
    ]
    module_path = audit / "module_responsibility_matrix.csv"
    if not module_path.exists():
        write_csv(module_path, module_rows)
    root_cause_path = audit / "root_cause_analysis.md"
    if not root_cause_path.exists():
        root_cause_path.write_text(
        "# Root cause analysis\n\n"
        "v2.3 repaired evidence identity but still allowed presentation behavior to be distributed across the planner, visual-specific payloads, footer, notes, and renderer. That duplication explains stale or repeated source labels, loss of brief language, silent row truncation, cross-entity titles, sequential rendering of independent interactions, card fallback for matrices, false DAG-like flows, and incomplete slide-budget coverage.\n\n"
        "v2.4 removes those independent inference points: the canonical object graph owns claim and source identity; SlideSpec owns narrative, language, layout and coverage; VisualSpec owns field roles, entity/facet scope, row accounting and chart semantics. The renderer consumes these objects and the finalizer verifies every derived artifact against them.\n",
            encoding="utf-8",
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Finalize Academic PPT Workflow v2.4 QA and human-review gate")
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--suite", type=Path, required=True)
    parser.add_argument("--audit-root", type=Path, required=True)
    parser.add_argument("--benchmark-root", type=Path, required=True)
    parser.add_argument("--python", type=Path, default=Path(sys.executable))
    args = parser.parse_args()

    run_root = args.run_root.resolve()
    suite = args.suite.resolve()
    audit = args.audit_root.resolve()
    benchmark = args.benchmark_root.resolve()
    audit.mkdir(parents=True, exist_ok=True)
    benchmark.mkdir(parents=True, exist_ok=True)

    try:
        generation_complete = read_json(run_root / "generation_complete.json")
        if as_bool(generation_complete.get("ground_truth_used_for_generation")):
            raise BlockedError("generation_complete.json reports ground-truth use during generation")
        deck_index = read_csv(run_root / "deck_index.csv")
        if not deck_index:
            raise BlockedError("deck_index.csv is empty")
        pp_index = index_by_project(index_records(run_root / "powerpoint_render_index.json"))
        lo_index = index_by_project(index_records(run_root / "libreoffice_render_index.json"))
    except BlockedError as exc:
        (audit / "final_status.md").write_text(
            f"# v2.4 final status\n\n`BLOCKED`\n\n- Reason: {exc}\n",
            encoding="utf-8",
        )
        print("STATUS=BLOCKED")
        print(f"REASON={exc}")
        return 2

    write_architecture_reports(audit)

    binding_rows: list[dict[str, Any]] = []
    footer_rows: list[dict[str, Any]] = []
    language_rows: list[dict[str, Any]] = []
    coverage_rows_all: list[dict[str, Any]] = []
    omission_rows: list[dict[str, Any]] = []
    semantic_rows: list[dict[str, Any]] = []
    layout_rows: list[dict[str, Any]] = []
    title_rows: list[dict[str, Any]] = []
    typography_rows: list[dict[str, Any]] = []
    perceptual_rows: list[dict[str, Any]] = []
    geometry_rows: list[dict[str, Any]] = []
    integrity_rows: list[dict[str, Any]] = []
    claim_gt_rows: list[dict[str, Any]] = []
    conflict_gt_rows: list[dict[str, Any]] = []
    unresolved_gt_rows: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []
    contact_map: list[dict[str, Any]] = []
    project_summaries: list[dict[str, Any]] = []
    invented_claims = 0
    typefaces_all: set[str] = set()

    # Formal comparison starts only after generation and both render indexes
    # exist.  Preserve any earlier process warning instead of claiming this is
    # necessarily the first human/tool read.
    gt_access_time = datetime.now(timezone.utc).isoformat()
    prelog = audit / "ground_truth_access_prelog.md"
    prior_access = prelog.read_text(encoding="utf-8") if prelog.is_file() else "No earlier access was recorded."
    (audit / "ground_truth_access_log.md").write_text(
        "# Ground-truth access log\n\n"
        f"- Formal comparison started UTC: {gt_access_time}\n"
        "- Formal stage: POST_GENERATION_GROUND_TRUTH_EVALUATION\n"
        "- Preconditions: generation_complete.json present; PowerPoint and LibreOffice render indexes present.\n"
        "- Ground truth is excluded from source discovery, manifests, planning inputs and renderer inputs.\n\n"
        "## Earlier access record\n\n" + prior_access + "\n",
        encoding="utf-8",
    )

    try:
        for deck in deck_index:
            project = deck["project_key"]
            arm = deck.get("arm", "V24")
            kind = deck["kind"]
            output = Path(deck["output_dir"]).resolve()
            project_root = suite / project
            if not project_root.is_dir():
                raise BlockedError(f"Isolated project root is missing: {project_root}")
            graph = read_json(output / "canonical_object_graph.json")
            slides = graph["slide_specs"]
            visual_specs = [visual for slide in slides for visual in slide.get("visual_specs", [])]
            if read_json(output / "slide_spec.json") != slides or read_json(output / "visual_spec.json") != visual_specs:
                raise BlockedError(f"Derived SlideSpec/VisualSpec identity mismatch: {project}")
            if graph.get("backend") != "native_pptxgenjs" or as_bool(graph.get("external_skill_used")):
                raise BlockedError(f"Unapproved backend or external Skill recorded: {project}")

            pptx = Path(deck["pptx_path"])
            object_manifest = read_json(pptx.with_suffix(".object_manifest.json"))
            object_by_id = {item["visual_id"]: item for item in object_manifest}
            declared_rendered_counts = {
                visual["visual_id"]: visual.get("rendered_row_count") for visual in visual_specs
            }
            # The renderer's object manifest is the independent authority for
            # actual row/mark coverage.  Downstream omission and semantic QA
            # consume that value rather than trusting a planner-side default.
            for visual in visual_specs:
                artifact = object_by_id.get(visual["visual_id"])
                if artifact is not None:
                    visual["rendered_row_count"] = artifact.get("rendered_row_count")

            pp = lookup_index(pp_index, project, arm)
            lo = lookup_index(lo_index, project, arm)
            geometry = read_json(Path(pp["geometry"]))
            pp_images = sorted(Path(pp["preview"]).glob("slide_*.png"))
            lo_images = sorted(Path(lo["preview"]).glob("slide_*.png"))
            slide_count = int(deck["slide_count"])
            if not (len(slides) == slide_count == len(pp_images) == len(lo_images)):
                raise BlockedError(
                    f"Rendered page assets are incomplete for {project}: "
                    f"spec={len(slides)}, deck={slide_count}, PP={len(pp_images)}, LO={len(lo_images)}"
                )

            code = hashlib.sha256(f"{run_root.name}|{project}".encode("utf-8")).hexdigest()[:6].upper()
            sheet = benchmark / "contact_sheets" / f"Deck_{code}.png"
            contact_sheet(pp_images, sheet, code)
            contact_sheet(pp_images, output / "contact_sheet.png", code)
            contact_map.append({"anonymous_deck_code": code, "project_key": project, "contact_sheet": str(sheet)})
            review, review_complete, review_total = review_form(output, code)
            review_rows.append(review)

            project_binding, project_footer = binding_checks(project, kind, project_root, graph, output, geometry)
            binding_rows.extend(project_binding)
            footer_rows.extend(project_footer)
            language_rows.extend(language_metrics(project, slide) for slide in slides)

            project_coverage = read_csv(output / "narrative_coverage_report.csv")
            for item in project_coverage:
                item = {"project_key": project, **item}
                item["passed"] = not as_bool(item.get("omitted"))
                coverage_rows_all.append(item)
            omission_rows.extend(omission_checks(project, slides, project_coverage))
            semantic_rows.extend(semantic_checks(project, kind, slides))
            layout_record, project_titles = layout_metrics(project, slides)
            layout_rows.append(layout_record)
            title_rows.extend(project_titles)

            project_typography, typefaces = classify_typography(project, pptx, slides)
            typography_rows.extend(project_typography)
            typefaces_all.update(typefaces)

            project_geometry, project_perceptual = actual_geometry(project, geometry, slides, pp_images, lo_images)
            geometry_rows.extend(project_geometry)
            perceptual_rows.extend(project_perceptual)

            manifest = read_csv(output / "source_manifest.csv")
            integrity_rows.extend(input_integrity(project, project_root, manifest))

            gt = project_root / "expected_ground_truth"
            if not gt.is_dir() or gt.is_relative_to(project_root / "input"):
                raise BlockedError(f"Ground truth is missing or not isolated from input: {project}")
            claims, conflicts, unresolved, invented = ground_truth_compare(project, output, gt)
            claim_gt_rows.extend(claims)
            conflict_gt_rows.extend(conflicts)
            unresolved_gt_rows.extend(unresolved)
            invented_claims += invented

            artifact_match = all(
                visual["visual_id"] in object_by_id
                and object_by_id[visual["visual_id"]].get("data_contract_id") == visual.get("data_contract_id")
                and object_by_id[visual["visual_id"]].get("source_bindings") == visual.get("source_bindings")
                and int(object_by_id[visual["visual_id"]].get("rendered_row_count", -1)) == int(visual.get("rendered_row_count", -2))
                and set(object_by_id[visual["visual_id"]].get("rendered_row_keys", [])) == set(visual.get("expected_row_keys", []))
                and set(visual.get("rendered_row_keys", [])) == set(visual.get("expected_row_keys", []))
                and declared_rendered_counts.get(visual["visual_id"]) == object_by_id[visual["visual_id"]].get("rendered_row_count")
                and int(object_by_id[visual["visual_id"]].get("created_object_count", 0)) > 0
                for visual in visual_specs
            ) and len(object_by_id) == len(visual_specs)
            for visual in visual_specs:
                artifact = object_by_id.get(visual["visual_id"], {})
                created = int(artifact.get("created_object_count", 0) or 0)
                semantic_rows.append({
                    "project_key": project,
                    "slide_id": visual.get("slide_id", artifact.get("slide_id", "")),
                    "visual_id": visual["visual_id"],
                    "visual_type": visual.get("visual_type", ""),
                    "check_id": "renderer_created_objects",
                    "severity": "error",
                    "passed": created > 0,
                    "details": (
                        f"created_object_count={created}; rendered_row_count="
                        f"{artifact.get('rendered_row_count', 'missing')} (object manifest authority)"
                    ),
                })
            semantic_rows.append({
                "project_key": project, "slide_id": "ALL", "visual_id": "ALL",
                "visual_type": "object_manifest", "check_id": "native_chart_data_binding",
                "severity": "error", "passed": artifact_match,
                "details": "object manifest matches VisualSpec data contracts/sources/rows and records created_object_count > 0",
            })

            pp_open = as_bool(pp.get("powerpoint_open"))
            pp_match = int(pp.get("pptx_slide_count", -1)) == slide_count == int(pp.get("png_count", -2))
            lo_match = int(lo.get("pdf_page_count", -1)) == slide_count and int(lo.get("exit_code", -1)) == 0
            project_summaries.append({
                "project_key": project, "kind": kind, "slide_count": slide_count,
                "target_slide_count": int(deck["target_slide_count"]),
                "generation_seconds": float(deck["generation_seconds"]),
                "powerpoint_open": pp_open, "powerpoint_page_match": pp_match,
                "libreoffice_page_match": lo_match, "layout_pass": layout_record["passed"],
                "human_review_complete": review_complete, "human_score": review_total if review_total is not None else "",
            })

            write_csv(output / "visual_qa.csv", [row for row in project_perceptual])
            write_csv(output / "source_binding_accuracy.csv", project_binding)
            write_csv(output / "narrative_coverage_report.csv", [row for row in coverage_rows_all if row["project_key"] == project])
    except BlockedError as exc:
        (audit / "final_status.md").write_text(
            f"# v2.4 final status\n\n`BLOCKED`\n\n- Reason: {exc}\n",
            encoding="utf-8",
        )
        print("STATUS=BLOCKED")
        print(f"REASON={exc}")
        return 2

    write_csv(audit / "source_binding_accuracy.csv", binding_rows)
    write_csv(audit / "source_footer_regression.csv", footer_rows)
    write_csv(audit / "language_compliance.csv", language_rows)
    write_csv(audit / "narrative_coverage_report.csv", coverage_rows_all)
    write_csv(audit / "omission_manifest.csv", omission_rows)
    write_csv(audit / "chart_semantic_qa.csv", semantic_rows)
    write_csv(audit / "layout_family_report.csv", layout_rows)
    write_csv(audit / "generic_title_report.csv", title_rows)
    write_csv(audit / "typography_report.csv", typography_rows)
    write_csv(audit / "perceptual_visual_qa.csv", perceptual_rows)
    write_csv(audit / "geometry_qa.csv", geometry_rows)
    write_csv(audit / "input_integrity.csv", integrity_rows)
    write_csv(audit / "ground_truth_claim_comparison.csv", claim_gt_rows)
    write_csv(audit / "ground_truth_conflict_comparison.csv", conflict_gt_rows)
    write_csv(audit / "ground_truth_unresolved_comparison.csv", unresolved_gt_rows)
    write_csv(audit / "contact_sheet_code_map.csv", contact_map)
    write_csv(benchmark / "visual_blind_review_form.csv", review_rows, list(review_rows[0]) if review_rows else None)
    write_csv(benchmark / "project_results.csv", project_summaries)

    slide_budget_lines = ["# Slide budget report", ""]
    slide_budget_pass = True
    for item in project_summaries:
        ratio = item["slide_count"] / max(1, item["target_slide_count"])
        passed = ratio >= 0.80
        slide_budget_pass &= passed
        slide_budget_lines.extend([
            f"## {item['project_key']}", "",
            f"- Target: {item['target_slide_count']}",
            f"- Rendered: {item['slide_count']}",
            f"- Ratio: {ratio:.1%}",
            f"- Gate: {'PASS' if passed else 'FAIL — compression rationale required'}", "",
        ])
    (audit / "slide_budget_report.md").write_text("\n".join(slide_budget_lines), encoding="utf-8")

    binding_accuracy = sum(as_bool(row["passed"]) for row in binding_rows) / max(1, len(binding_rows))
    field_rows = [row for row in binding_rows if str(row["check_type"]).startswith("field_lineage:")]
    field_coverage = sum(as_bool(row["passed"]) for row in field_rows) / max(1, len(field_rows))
    footer_pass = all(as_bool(row["passed"]) for row in footer_rows)
    language_by_project: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in language_rows:
        language_by_project[row["project_key"]].append(row)
    language_project_scores = {
        project: sum(row["cjk_characters"] for row in records)
        / max(1, sum(row["cjk_characters"] + row["non_allowlisted_latin_characters"] for row in records))
        for project, records in language_by_project.items()
    }
    language_pass = all(score >= .90 for score in language_project_scores.values()) and all(
        row["declared_language"] == "zh-CN"
        and as_bool(row["title_has_cjk"])
        and not as_bool(row["mojibake_detected"])
        and not as_bool(row["abnormal_spacing_detected"])
        for row in language_rows
    )
    coverage_pass = all(as_bool(row["passed"]) for row in coverage_rows_all)
    silent_truncation = sum(as_bool(row["silent"]) for row in omission_rows)
    semantic_failures = [row for row in semantic_rows if row["severity"] == "error" and not as_bool(row["passed"])]
    layout_pass = all(as_bool(row["passed"]) for row in layout_rows)
    typography_pass = all(as_bool(row["passed"]) for row in typography_rows)
    geometry_pass = all(as_bool(row["passed"]) for row in geometry_rows)
    gross_visual_pass = not any(as_bool(row["gross_visual_failure"]) for row in perceptual_rows)
    pp_pass = all(as_bool(row["powerpoint_open"]) and as_bool(row["powerpoint_page_match"]) for row in project_summaries)
    lo_pass = all(as_bool(row["libreoffice_page_match"]) for row in project_summaries)
    integrity_pass = all(as_bool(row["passed"]) for row in integrity_rows)
    claim_recall = sum(row["classification"] == "true_positive" for row in claim_gt_rows) / max(1, len(claim_gt_rows))
    conflict_recall = sum(as_bool(row["recalled"]) for row in conflict_gt_rows) / max(1, len(conflict_gt_rows))
    unresolved_recall = sum(as_bool(row["recalled"]) for row in unresolved_gt_rows) / max(1, len(unresolved_gt_rows))

    baseline_mean, baseline_source = baseline_runtime()
    current_mean = sum(row["generation_seconds"] for row in project_summaries) / max(1, len(project_summaries))
    runtime_ratio = current_mean / baseline_mean if baseline_mean and baseline_mean > 0 else math.inf
    runtime_pass = runtime_ratio <= 1.5
    tests_pass, _ = run_tests(args.python, audit)

    all_reviewed = all(as_bool(row["human_review_complete"]) for row in project_summaries)
    human_scores_pass = all(
        row["human_score"] != "" and float(row["human_score"]) >= 75
        for row in project_summaries
    ) if all_reviewed else False
    completed_human_failure = any(
        as_bool(row["human_review_complete"])
        and row["human_score"] != ""
        and float(row["human_score"]) < 75
        for row in project_summaries
    )

    (audit / "source_footer_regression.md").write_text(
        "# Source footer regression\n\n"
        f"- Slide footers checked: {len(footer_rows)}\n"
        f"- Footer/notes derived-binding matches: {sum(as_bool(row['passed']) for row in footer_rows)}/{len(footer_rows)}\n"
        f"- Gate: {'PASS' if footer_pass else 'FAIL'}\n\n"
        "Audience-facing footers are checked against the first three unique SlideSpec short-source labels; full source identifiers remain in notes and audit maps.\n",
        encoding="utf-8",
    )
    (audit / "font_fallback_report.md").write_text(
        "# Font fallback report\n\n"
        f"- OOXML-declared typefaces: {', '.join(sorted(typefaces_all)) or 'theme inherited'}\n"
        "- Preferred East Asian face: Microsoft YaHei\n"
        f"- PowerPoint open/geometry gate: {'PASS' if pp_pass and geometry_pass else 'FAIL'}\n"
        f"- LibreOffice page-parity gate: {'PASS' if lo_pass else 'FAIL'}\n"
        "- Font substitution is treated as a compatibility warning unless actual PowerPoint text bounds overflow or minimum-size checks fail. Font files were not copied or distributed.\n",
        encoding="utf-8",
    )
    (audit / "design_system_audit.md").write_text(
        "# Presentation Design System audit\n\n"
        f"- Decks with at least five layout families and ratio/repetition gates passed: {sum(as_bool(row['passed']) for row in layout_rows)}/{len(layout_rows)}\n"
        f"- Card-grid ratio gate (<=30%): {'PASS' if all(float(row['card_grid_ratio']) <= .30 for row in layout_rows) else 'FAIL'}\n"
        f"- Text-only ratio gate (<=15%): {'PASS' if all(float(row['text_only_ratio']) <= .15 for row in layout_rows) else 'FAIL'}\n"
        f"- Consecutive identical composition gate (<=2): {'PASS' if all(int(row['max_consecutive_same_layout']) <= 2 for row in layout_rows) else 'FAIL'}\n"
        "- Shared master families: cover, section_divider, hero_insight, chart_led, figure_with_callout, split_screen, comparison_matrix, process_branch, evidence_ladder, risk_heatmap, conclusion_synthesis, appendix_audit.\n"
        "- Domain profiles remain differentiated while sharing typography and academic safety rules.\n",
        encoding="utf-8",
    )
    (audit / "scientific_regression_report.md").write_text(
        "# Scientific regression report\n\n"
        f"- Claim/value/canonical-source recall: {claim_recall:.1%}\n"
        f"- Conflict recall: {conflict_recall:.1%}\n"
        f"- Unresolved recall: {unresolved_recall:.1%}\n"
        f"- Invented canonical claims: {invented_claims}\n"
        f"- Protocol result fabrication failures: {sum(row['check_id'] == 'protocol_no_results' and not as_bool(row['passed']) for row in semantic_rows)}\n"
        f"- Input integrity: {sum(as_bool(row['passed']) for row in integrity_rows)}/{len(integrity_rows)}\n"
        f"- Scientific regression gate: {'PASS' if claim_recall == conflict_recall == unresolved_recall == 1 and invented_claims == 0 and integrity_pass else 'FAIL'}\n",
        encoding="utf-8",
    )
    baseline_line = (
        f"- v2.3 B2 mean generation time: {baseline_mean:.3f} s/deck\n"
        if baseline_mean is not None else "- v2.3 B2 baseline: unavailable\n"
    )
    runtime_text = (
        "# Runtime report\n\n"
        f"- v2.4 mean generation time: {current_mean:.3f} s/deck\n"
        + baseline_line
        + f"- Baseline source: {baseline_source or 'not found'}\n"
        + f"- Ratio: {runtime_ratio:.2f}x\n"
        + f"- Gate (<=1.5x): {'PASS' if runtime_pass else 'FAIL'}\n"
    )
    (audit / "runtime_report.md").write_text(runtime_text, encoding="utf-8")
    (audit / "input_integrity_report.md").write_text(
        "# Input integrity report\n\n"
        f"- Registered source paths checked: {len(integrity_rows)}\n"
        f"- Unchanged, registered, and ground-truth-isolated: {sum(as_bool(row['passed']) for row in integrity_rows)}/{len(integrity_rows)}\n"
        f"- Gate: {'PASS' if integrity_pass else 'FAIL'}\n",
        encoding="utf-8",
    )
    (audit / "human_visual_review_status.md").write_text(
        "# Human visual review status\n\n"
        f"- Completed forms: {sum(as_bool(row['human_review_complete']) for row in project_summaries)}/{len(project_summaries)}\n"
        f"- All completed scores >=75/100: {'YES' if human_scores_pass else 'NO / NOT YET ASSESSED'}\n"
        "- Automatic perceptual metrics are screening aids and do not replace the human review.\n",
        encoding="utf-8",
    )

    automated_gates = {
        "slide-level source binding accuracy = 100%": binding_accuracy == 1,
        "field-level lineage coverage = 100%": field_coverage == 1,
        "source footer and notes identity": footer_pass,
        "zh-CN language compliance >=90%": language_pass,
        "mandatory coverage = 100%": coverage_pass,
        "controlled-row silent truncation = 0": silent_truncation == 0,
        "semantic chart grammar": not semantic_failures,
        "slide budget >=80% of target": slide_budget_pass,
        "layout diversity/card/text/repetition": layout_pass,
        "minimum typography": typography_pass,
        "PowerPoint geometry": geometry_pass,
        "perceptual gross-failure screen": gross_visual_pass,
        "PowerPoint open and page parity": pp_pass,
        "LibreOffice page parity": lo_pass,
        "claim/conflict/unresolved recall = 100%": claim_recall == conflict_recall == unresolved_recall == 1,
        "invented claims = 0": invented_claims == 0,
        "input hashes and paths unchanged": integrity_pass,
        "native PptxGenJS only": all(row.get("kind") for row in project_summaries) and not as_bool(generation_complete.get("external_skill_used")) and generation_complete.get("backend") == "native_pptxgenjs",
        "runtime <=1.5x v2.3 B2": runtime_pass,
        "regression tests": tests_pass,
    }
    automated_pass = all(automated_gates.values())
    critical_scientific = not all([
        binding_accuracy == 1, field_coverage == 1, coverage_pass,
        not semantic_failures, claim_recall == 1, conflict_recall == 1,
        unresolved_recall == 1, invented_claims == 0, integrity_pass,
    ])
    if automated_pass and all_reviewed and human_scores_pass:
        status = "WORKFLOW_PRESENTATION_DESIGN_REMEDIATION_PASSED"
    elif automated_pass and not all_reviewed and not completed_human_failure:
        status = "READY_FOR_HUMAN_VISUAL_REVIEW"
    elif critical_scientific:
        status = "WORKFLOW_PRESENTATION_DESIGN_REMEDIATION_FAILED"
    else:
        status = "WORKFLOW_PRESENTATION_DESIGN_REMEDIATION_PARTIAL"

    gate_lines = [f"| {name} | {'PASS' if passed else 'FAIL'} |" for name, passed in automated_gates.items()]
    human_lines = [
        f"| {row['project_key']} | {'complete' if row['human_review_complete'] else 'pending'} | {row['human_score'] or 'not assessed'} |"
        for row in project_summaries
    ]
    (audit / "final_status.md").write_text(
        "# Academic PPT Workflow v2.4 final status\n\n"
        f"`{status}`\n\n"
        "## Automated gates\n\n| Gate | Result |\n|---|---|\n"
        + "\n".join(gate_lines)
        + "\n\n## Human visual gate\n\n| Project | Review | Score / 100 |\n|---|---|---:|\n"
        + "\n".join(human_lines)
        + "\n\nThe status cannot be promoted to `WORKFLOW_PRESENTATION_DESIGN_REMEDIATION_PASSED` until every human form is complete and every deck scores at least 75/100.\n\n"
        "- External Skill introduced: 0\n"
        "- Production backend: native_pptxgenjs\n"
        "- v2 overall state: PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING\n",
        encoding="utf-8",
    )

    print(f"STATUS={status}")
    print(f"SOURCE_BINDING_ACCURACY={binding_accuracy:.4f}")
    print(f"FIELD_LINEAGE_COVERAGE={field_coverage:.4f}")
    print(f"CLAIM_RECALL={claim_recall:.4f}")
    print(f"CONFLICT_RECALL={conflict_recall:.4f}")
    print(f"UNRESOLVED_RECALL={unresolved_recall:.4f}")
    print(f"HUMAN_REVIEW_COMPLETED={sum(as_bool(row['human_review_complete']) for row in project_summaries)}/{len(project_summaries)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
