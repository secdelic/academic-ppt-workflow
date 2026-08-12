from __future__ import annotations

import argparse
import csv
import hashlib
import importlib.util
import json
import math
import re
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


REPO = Path(__file__).resolve().parents[2]
V24_FINALIZER = REPO / "scripts" / "v2_4" / "finalize_v2_4.py"

SCORE_FIELDS = (
    ("scientific_clarity_20", 20.0),
    ("visual_semantic_correctness_20", 20.0),
    ("hierarchy_15", 15.0),
    ("composition_15", 15.0),
    ("chart_professionalism_10", 10.0),
    ("consistency_8", 8.0),
    ("readability_5", 5.0),
    ("visual_richness_5", 5.0),
    ("presentation_readiness_2", 2.0),
)

LEGACY_V24_SCORE_FIELDS = (
    ("scientific_clarity_20", 20.0),
    ("visual_semantic_correctness_20", 20.0),
    ("hierarchy_15", 15.0),
    ("composition_15", 15.0),
    ("chart_professionalism_10", 10.0),
    ("consistency_10", 10.0),
    ("readability_5", 5.0),
    ("presentation_readiness_5", 5.0),
)

VISUAL_SCORE_FIELDS = [
    "pair_code", "anonymous_deck", *[name for name, _ in SCORE_FIELDS],
    "total_100", "reviewer", "review_date", "comments",
]

SOURCE_FIGURE_TYPES = {"source_figure", "umap_figure"}

MOJIBAKE_MARKERS = (
    "\ufffd", "锟斤拷", "Ã", "Â", "â€", "鈥", "馃", "闂?", "鎴?",
)

SENSITIVE_CHART_RE = re.compile(
    r"(?:(?:patient|participant|subject|person|record|medical_record)[ _-]?"
    r"(?:id|number|name)|\bmrn\b|姓名|身份证|住院号|病案号)",
    re.I,
)
SOURCE_ID_RE = re.compile(r"\bSRC-[A-F0-9]{8,}\b", re.I)
SHA_RE = re.compile(r"\b[a-f0-9]{64}\b", re.I)
LOCAL_PATH_RE = re.compile(r"(?:[A-Za-z]:\\|(?:input|staging|output)/)", re.I)


class BlockedError(RuntimeError):
    """A missing or unsafe prerequisite that prevents a valid promotion audit."""


def _load_v24() -> Any:
    if not V24_FINALIZER.is_file():
        raise BlockedError(f"v2.4 finalizer is missing: {V24_FINALIZER}")
    spec = importlib.util.spec_from_file_location("academic_ppt_v24_finalize", V24_FINALIZER)
    if spec is None or spec.loader is None:
        raise BlockedError(f"Cannot import v2.4 finalizer: {V24_FINALIZER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().casefold() in {"1", "true", "yes", "pass", "passed"}


def number(value: Any) -> float | None:
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return None
    return parsed if math.isfinite(parsed) else None


def read_json(path: Path) -> Any:
    if not path.is_file():
        raise BlockedError(f"Required JSON is missing: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise BlockedError(f"Cannot read JSON {path}: {exc}") from exc


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise BlockedError(f"Required CSV is missing: {path}")
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))
    except (OSError, UnicodeError, csv.Error) as exc:
        raise BlockedError(f"Cannot read CSV {path}: {exc}") from exc


def write_csv(path: Path, records: list[dict[str, Any]], fields: Iterable[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    field_list = list(fields)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=field_list, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)


def canonical(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def optional_mapping_contract(value: Any) -> Any:
    """Treat an omitted optional mapping and an emitted empty mapping alike.

    The renderer serializes non-applicable presentation contracts (for example,
    native-chart privacy on an editable-shape visual) as ``{}``, while the
    canonical VisualSpec leaves them as ``null``.  Those representations are
    semantically identical only when the mapping is empty; non-empty mappings
    and malformed non-mapping values remain subject to exact comparison.
    """
    return {} if value is None or value == {} else value


def path_from_row(row: dict[str, Any], key: str, root: Path) -> Path:
    value = Path(str(row.get(key, "")))
    if not str(value):
        raise BlockedError(f"deck_index row lacks {key}: {row.get('project_key', 'unknown')}")
    return value.resolve() if value.is_absolute() else (root / value).resolve()


def records_from_index(path: Path) -> list[dict[str, Any]]:
    value = read_json(path)
    if isinstance(value, list):
        return value
    if isinstance(value, dict):
        for key in ("results", "decks", "records"):
            if isinstance(value.get(key), list):
                return value[key]
    raise BlockedError(f"Render index does not contain records: {path}")


def unique_index(records: list[dict[str, Any]], label: str) -> dict[str, dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in records:
        grouped[str(row.get("project_key", ""))].append(row)
    duplicate = [key for key, values in grouped.items() if not key or len(values) != 1]
    if duplicate:
        raise BlockedError(f"{label} must contain exactly one row per project: {duplicate}")
    return {key: values[0] for key, values in grouped.items()}


def discover_pngs(path: Path) -> list[Path]:
    if not path.is_dir():
        return []
    images = sorted(path.glob("slide_*.png"))
    return images or sorted(path.glob("slide-*.png"))


def flatten_visuals(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [visual for slide in slides for visual in slide.get("visual_specs", [])]


def find_object_manifest(output: Path, pptx: Path) -> Path:
    exact = pptx.with_suffix(".object_manifest.json")
    if exact.is_file():
        return exact
    candidates = sorted(output.glob("*.object_manifest.json"))
    if len(candidates) != 1:
        raise BlockedError(f"Expected one object manifest in {output}; found {len(candidates)}")
    return candidates[0]


def load_run(root: Path, *, expected_version: str) -> dict[str, Any]:
    generation = read_json(root / "generation_complete.json")
    if as_bool(generation.get("ground_truth_used_for_generation")):
        raise BlockedError(f"{expected_version} reports ground-truth use during generation")
    if as_bool(generation.get("external_skill_used")):
        raise BlockedError(f"{expected_version} reports an external Skill in the production path")
    if generation.get("backend") != "native_pptxgenjs":
        raise BlockedError(
            f"{expected_version} backend is {generation.get('backend')!r}, not native_pptxgenjs"
        )
    deck_rows = read_csv(root / "deck_index.csv")
    decks = unique_index(deck_rows, f"{expected_version} deck index")
    pp = unique_index(records_from_index(root / "powerpoint_render_index.json"), f"{expected_version} PowerPoint index")
    lo = unique_index(records_from_index(root / "libreoffice_render_index.json"), f"{expected_version} LibreOffice index")
    if set(decks) != set(pp) or set(decks) != set(lo):
        raise BlockedError(
            f"{expected_version} deck/PowerPoint/LibreOffice project sets differ: "
            f"deck={sorted(decks)}, PP={sorted(pp)}, LO={sorted(lo)}"
        )
    if len(decks) != 4:
        raise BlockedError(f"{expected_version} requires four isolated projects; found {len(decks)}")
    return {"root": root, "generation": generation, "decks": decks, "pp": pp, "lo": lo}


def load_project(run: dict[str, Any], project: str) -> dict[str, Any]:
    deck = run["decks"][project]
    output = path_from_row(deck, "output_dir", run["root"])
    pptx = path_from_row(deck, "pptx_path", run["root"])
    if not output.is_dir() or not pptx.is_file() or pptx.stat().st_size == 0:
        raise BlockedError(f"Missing or empty deck output for {project}: {pptx}")
    graph = read_json(output / "canonical_object_graph.json")
    slides = graph.get("slide_specs")
    if not isinstance(slides, list) or not slides:
        raise BlockedError(f"Canonical graph lacks SlideSpec[] for {project}")
    visuals = flatten_visuals(slides)
    manifest_path = find_object_manifest(output, pptx)
    manifest = read_json(manifest_path)
    if not isinstance(manifest, list):
        raise BlockedError(f"Object manifest must be an array: {manifest_path}")
    visual_ids = [str(item.get("visual_id", "")) for item in visuals]
    manifest_ids = [str(item.get("visual_id", "")) for item in manifest]
    if Counter(visual_ids) != Counter(manifest_ids):
        raise BlockedError(f"VisualSpec/object-manifest identity mismatch for {project}")
    pp = run["pp"][project]
    lo = run["lo"][project]
    geometry = read_json(Path(str(pp.get("geometry", ""))))
    pp_images = discover_pngs(Path(str(pp.get("preview", ""))))
    lo_images = discover_pngs(Path(str(lo.get("preview", ""))))
    slide_count = int(deck.get("slide_count", len(slides)))
    if len(slides) != slide_count or len(pp_images) != slide_count or len(lo_images) != slide_count:
        raise BlockedError(
            f"Incomplete render set for {project}: specs={len(slides)}, index={slide_count}, "
            f"PowerPoint PNG={len(pp_images)}, LibreOffice PNG={len(lo_images)}"
        )
    return {
        "deck": deck, "output": output, "pptx": pptx, "graph": graph,
        "slides": slides, "visuals": visuals, "manifest": manifest,
        "manifest_path": manifest_path, "manifest_by_visual": {
            str(item["visual_id"]): item for item in manifest
        }, "pp": pp, "lo": lo, "geometry": geometry,
        "pp_images": pp_images, "lo_images": lo_images, "slide_count": slide_count,
    }


def _visible_text(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).casefold()


def _shape_box(shape: dict[str, Any], *, actual_text: bool = False) -> tuple[float, float, float, float] | None:
    top_key = "bound_top" if actual_text and shape.get("bound_top") is not None else "top"
    height_key = "bound_height" if actual_text and shape.get("bound_height") is not None else "height"
    values = (shape.get("left"), shape.get(top_key), shape.get("width"), shape.get(height_key))
    if any(value is None for value in values):
        return None
    try:
        return tuple(float(value) for value in values)  # type: ignore[return-value]
    except (TypeError, ValueError):
        return None


def _registered_footer_labels(spec: dict[str, Any]) -> list[str]:
    return [
        _visible_text(binding.get("short_label_zh"))
        for binding in spec.get("source_bindings", [])
        if _visible_text(binding.get("short_label_zh"))
    ]


def is_v25_legal_footer_shape(
    shape: dict[str, Any], spec: dict[str, Any], slide_number: int,
    slide_width: float, slide_height: float,
) -> bool:
    """Recognize only v2.5 footer objects wholly inside the footer zone.

    Geometry is expressed either in inches (OOXML) or points/pixels (COM); the
    ratio-based test works for both. A body object crossing into the zone is
    intentionally never exempted, even if it happens to contain source-like
    words.
    """
    box = _shape_box(shape)
    if box is None:
        return False
    left, top, width, height = box
    footer_top = slide_height * (7.12 / 7.5)
    tolerance = max(0.01, slide_height * 0.002)
    if top < footer_top - tolerance or top + height > slide_height + tolerance:
        return False
    text = _visible_text(shape.get("text"))
    if not shape.get("has_text") and not text:
        return (
            left <= tolerance and width >= slide_width * 0.95
            and height <= slide_height * 0.08
        )
    if height > slide_height * 0.08:
        return False
    name = str(shape.get("name", "")).casefold()
    if text in {"来源：", "来源:", str(slide_number), "未登记"}:
        return True
    if any(token in name for token in ("footer-prefix", "footer-label", "page-number")):
        return True
    labels = _registered_footer_labels(spec)
    return bool(
        text and labels and (
            any(label in text or text in label for label in labels)
            or (re.search(r"等\d+项$", text) and any(label in text for label in labels))
        )
    )


def v25_footer_intrusion_count(
    slide_geometry: dict[str, Any], spec: dict[str, Any], slide_number: int,
    geometry: dict[str, Any],
) -> int:
    shapes = slide_geometry.get("shapes", [])
    slide_width = float(slide_geometry.get("width", geometry.get("slide_width", 960)))
    slide_height = float(slide_geometry.get("height", geometry.get("slide_height", 540)))
    footer_top = slide_height * (7.12 / 7.5)
    tolerance = max(0.5, slide_height * 0.002)
    intrusions = 0
    for shape in shapes:
        if is_v25_legal_footer_shape(shape, spec, slide_number, slide_width, slide_height):
            continue
        box = _shape_box(shape, actual_text=bool(shape.get("has_text")))
        if box is None:
            continue
        left, top, width, height = box
        is_full_slide_background = (
            not shape.get("has_text") and left <= 2 and top <= 2
            and width >= slide_width * 0.95 and height >= slide_height * 0.95
        )
        if is_full_slide_background:
            continue
        if top + height > footer_top + tolerance:
            intrusions += 1
    return intrusions


def actual_geometry_v25(
    v24: Any, project: str, geometry: dict[str, Any], slides: list[dict[str, Any]],
    pp_images: list[Path], lo_images: list[Path],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rows, perceptual = v24.actual_geometry(project, geometry, slides, pp_images, lo_images)
    for index, row in enumerate(rows, 1):
        if index > len(geometry.get("slides", [])):
            continue
        spec = slides[index - 1] if index <= len(slides) else {}
        intrusion = v25_footer_intrusion_count(geometry["slides"][index - 1], spec, index, geometry)
        row["footer_intrusion"] = intrusion
        row["passed"] = not any((
            int(row.get("text_overlap_pairs", 0) or 0),
            int(row.get("clipping", 0) or 0), intrusion,
            int(row.get("off_slide_shapes", 0) or 0),
        ))
    return rows, perceptual


def pptx_typography_with_geometry(
    pptx: Path, ns: dict[str, str],
) -> tuple[dict[int, list[dict[str, Any]]], set[str]]:
    slides: dict[int, list[dict[str, Any]]] = defaultdict(list)
    typefaces: set[str] = set()
    with zipfile.ZipFile(pptx) as archive:
        members = [name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)]
        members.sort(key=lambda item: int(re.search(r"(\d+)", Path(item).stem).group(1)))
        for slide_number, member in enumerate(members, 1):
            root = ET.fromstring(archive.read(member))
            for shape in root.findall(".//p:sp", ns):
                c_nv_pr = shape.find("./p:nvSpPr/p:cNvPr", ns)
                name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
                text = "".join(node.text or "" for node in shape.findall(".//a:t", ns))
                sizes: list[float] = []
                for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
                    for prop in shape.findall(f".//{tag}", ns):
                        if prop.attrib.get("sz"):
                            sizes.append(int(prop.attrib["sz"]) / 100)
                        face = prop.attrib.get("typeface")
                        if face:
                            typefaces.add(face)
                xfrm = shape.find("./p:spPr/a:xfrm", ns)
                off = xfrm.find("./a:off", ns) if xfrm is not None else None
                ext = xfrm.find("./a:ext", ns) if xfrm is not None else None
                slides[slide_number].append({
                    "name": name, "text": text, "sizes": sizes,
                    "has_text": bool(text),
                    "left": int(off.attrib["x"]) / 914400 if off is not None else None,
                    "top": int(off.attrib["y"]) / 914400 if off is not None else None,
                    "width": int(ext.attrib["cx"]) / 914400 if ext is not None else None,
                    "height": int(ext.attrib["cy"]) / 914400 if ext is not None else None,
                })
    return slides, typefaces


def v25_typography_role(shape: dict[str, Any], spec: dict[str, Any], slide_number: int) -> str:
    name = str(shape.get("name", "")).casefold()
    text = _visible_text(shape.get("text"))
    sizes = shape.get("sizes", [])
    if text == _visible_text(spec.get("title")) or ":title" in name or "deck-title" in name:
        return "title"
    if is_v25_legal_footer_shape(shape, spec, slide_number, 13.333, 7.5):
        return "footer"
    top = number(shape.get("top"))
    if top is not None and top <= 1.40 and (
        ":section" in name or text in {_visible_text(item) for item in ("问题", "方法", "结果", "局限性", "结论", "计划", "附录")}
    ):
        return "section"
    return "chart_detail" if sizes and all(float(size) <= 16.1 for size in sizes) else "body"


def classify_typography_v25(
    project: str, pptx: Path, slide_specs: list[dict[str, Any]], v24: Any,
) -> tuple[list[dict[str, Any]], set[str]]:
    shapes_by_slide, typefaces = pptx_typography_with_geometry(pptx, v24.NS)
    records: list[dict[str, Any]] = []
    for index, spec in enumerate(slide_specs, 1):
        title_sizes: list[float] = []
        body_sizes: list[float] = []
        chart_sizes: list[float] = []
        footer_sizes: list[float] = []
        title_under = body_under = chart_under = detail_under = footer_bad = 0
        for shape in shapes_by_slide.get(index, []):
            sizes = shape.get("sizes", [])
            if not sizes or not shape.get("text"):
                continue
            role = v25_typography_role(shape, spec, index)
            if role == "title":
                title_sizes.extend(sizes); title_under += sum(size < 30 for size in sizes)
            elif role == "footer":
                footer_sizes.extend(sizes); footer_bad += sum(not (9.8 <= size <= 11.2) for size in sizes)
            elif role == "section":
                continue
            elif role == "chart_detail":
                chart_sizes.extend(sizes); chart_under += sum(size < 15 for size in sizes)
                detail_under += sum(15 <= size < 16 for size in sizes if "chart" not in str(shape.get("name", "")).casefold())
            else:
                body_sizes.extend(sizes); body_under += sum(size < 18 for size in sizes)
        passed = not any((title_under, body_under, chart_under, footer_bad))
        records.append({
            "project_key": project, "slide_number": index, "slide_id": spec["slide_id"],
            "title_min_pt": min(title_sizes) if title_sizes else "",
            "body_min_pt": min(body_sizes) if body_sizes else "",
            "chart_detail_min_pt": min(chart_sizes) if chart_sizes else "",
            "footer_min_pt": min(footer_sizes) if footer_sizes else "",
            "title_under_30": title_under, "body_under_18_unjustified": body_under,
            "chart_under_15": chart_under, "card_detail_under_16": detail_under,
            "footer_outside_10_11": footer_bad,
            "declared_typefaces": "|".join(sorted(typefaces)), "passed": passed,
        })
    return records, typefaces


def spec_value(spec: dict[str, Any], key: str, default: Any = None) -> Any:
    for container in (
        spec,
        spec.get("presentation", {}),
        spec.get("art_direction", {}),
        spec.get("presentation", {}).get("art_direction", {})
        if isinstance(spec.get("presentation"), dict) else {},
    ):
        if isinstance(container, dict) and key in container:
            return container[key]
    return default


def graph_art_direction(graph: dict[str, Any]) -> dict[str, Any]:
    for value in (
        graph.get("art_direction_spec"), graph.get("art_direction"),
        graph.get("presentation", {}).get("art_direction")
        if isinstance(graph.get("presentation"), dict) else None,
    ):
        if isinstance(value, dict) and value:
            return value
    return {}


def rhythm_entries(graph: dict[str, Any]) -> list[dict[str, Any]]:
    value = graph.get("deck_rhythm_plan", [])
    if isinstance(value, dict):
        for key in ("slides", "entries", "plan"):
            if isinstance(value.get(key), list):
                return value[key]
        return []
    return value if isinstance(value, list) else []


def consecutive_max(values: list[str]) -> int:
    maximum = 0
    current = 0
    previous: str | None = None
    for value in values:
        current = current + 1 if value == previous else 1
        previous = value
        maximum = max(maximum, current)
    return maximum


def stable_code(*parts: str, length: int = 8) -> str:
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()[:length].upper()


def normalized_bindings(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    fields = (
        "source_id", "source_file", "source_location", "fields_used",
        "row_filter", "aggregation", "canonical_status",
    )
    normalized = [{key: item.get(key) for key in fields} for item in items]
    return sorted(normalized, key=canonical)


def scientific_contract(graph: dict[str, Any]) -> dict[str, Any]:
    evidence = graph.get("canonical_evidence_graph", {})

    def claims() -> list[dict[str, Any]]:
        records = []
        for item in evidence.get("claims", []):
            records.append({
                "claim_id": item.get("claim_id"),
                "claim_text": item.get("claim_text"),
                "expected_value": item.get("expected_value"),
                "wording_boundary": item.get("wording_boundary"),
                "canonical_status": item.get("canonical_status"),
                "source_bindings": normalized_bindings(item.get("source_bindings", [])),
            })
        return sorted(records, key=lambda row: str(row.get("claim_id", "")))

    def conflicts() -> list[dict[str, Any]]:
        fields = (
            "conflict_id", "source_id", "source_file", "source_location",
            "conflicting_value", "canonical_value", "resolution",
            "manual_review_required",
        )
        return sorted(
            [{key: item.get(key) for key in fields} for item in evidence.get("conflicts", [])],
            key=lambda row: str(row.get("conflict_id", "")),
        )

    def unresolved() -> list[dict[str, Any]]:
        fields = (
            "item_id", "marker", "status", "source_id", "source_file",
            "source_location", "required_action", "manual_review_required",
        )
        return sorted(
            [{key: item.get(key) for key in fields} for item in evidence.get("unresolved", [])],
            key=lambda row: str(row.get("item_id", "")),
        )

    source_registry = sorted([
        {
            "source_id": item.get("source_id"), "source_file": item.get("source_file"),
            "source_type": item.get("source_type"), "source_role": item.get("source_role"),
            "sha256": item.get("sha256"), "size_bytes": item.get("size_bytes"),
            "parsed": item.get("parsed"),
            "ground_truth_excluded": item.get("ground_truth_excluded"),
        }
        for item in graph.get("source_registry", [])
    ], key=lambda row: str(row.get("source_id", "")))
    story = graph.get("story_graph", {})
    slides = graph.get("slide_specs", [])
    return {
        "source_registry": source_registry,
        "claims": claims(),
        "conflicts": conflicts(),
        "unresolved": unresolved(),
        "must_include": story.get("must_include", graph.get("brief", {}).get("must_include", [])),
        "must_exclude": story.get("must_exclude", graph.get("brief", {}).get("must_exclude", [])),
        "language": story.get("language", graph.get("brief", {}).get("language")),
        "slide_ids": sorted(str(slide.get("slide_id", "")) for slide in slides),
        "claim_ids_by_slide": sorted(
            (str(slide.get("slide_id", "")), sorted(
                str(claim.get("claim_id", "")) for claim in slide.get("claim_bindings", [])
            )) for slide in slides
        ),
        "backend": graph.get("backend"),
        "external_skill_used": graph.get("external_skill_used"),
        "ground_truth_used_for_generation": graph.get("ground_truth_used_for_generation"),
    }


def compare_scientific_contract(
    project: str, baseline: dict[str, Any], candidate: dict[str, Any],
) -> list[dict[str, Any]]:
    left = scientific_contract(baseline["graph"])
    right = scientific_contract(candidate["graph"])
    checks = (
        ("source_registry_and_hashes_frozen", "source_registry"),
        ("canonical_claim_registry_frozen", "claims"),
        ("conflict_registry_frozen", "conflicts"),
        ("unresolved_registry_frozen", "unresolved"),
        ("must_include_frozen", "must_include"),
        ("must_exclude_frozen", "must_exclude"),
        ("language_contract_frozen", "language"),
        ("stable_semantic_slide_ids", "slide_ids"),
        ("claim_slide_bindings_frozen", "claim_ids_by_slide"),
        ("native_backend_frozen", "backend"),
        ("no_external_skill", "external_skill_used"),
        ("no_ground_truth_generation", "ground_truth_used_for_generation"),
    )
    rows = []
    for check_id, field in checks:
        passed = canonical(left.get(field)) == canonical(right.get(field))
        if field == "backend":
            passed = passed and right.get(field) == "native_pptxgenjs"
        if field in {"external_skill_used", "ground_truth_used_for_generation"}:
            passed = passed and not as_bool(right.get(field))
        rows.append({
            "project_key": project, "check_id": check_id, "passed": passed,
            "baseline_digest": hashlib.sha256(canonical(left.get(field)).encode("utf-8")).hexdigest(),
            "candidate_digest": hashlib.sha256(canonical(right.get(field)).encode("utf-8")).hexdigest(),
            "details": "frozen contract preserved" if passed else f"candidate {field} differs from v2.4",
        })
    return rows


def validate_object_contract(
    project: str, candidate: dict[str, Any],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    by_id = candidate["manifest_by_visual"]
    slide_by_id = {str(slide.get("slide_id", "")): slide for slide in candidate["slides"]}
    for visual in candidate["visuals"]:
        visual_id = str(visual.get("visual_id", ""))
        artifact = by_id.get(visual_id, {})
        expected_keys = set(map(str, visual.get("expected_row_keys", [])))
        rendered_keys = set(map(str, artifact.get("rendered_row_keys", [])))
        source_match = canonical(artifact.get("source_bindings", [])) == canonical(visual.get("source_bindings", []))
        row_match = (
            expected_keys == rendered_keys
            and int(artifact.get("rendered_row_count", -1)) == int(visual.get("expected_row_count", -2))
            and int(artifact.get("created_object_count", 0) or 0) > 0
        )
        data_match = artifact.get("data_contract_id") == visual.get("data_contract_id")
        version_match = str(artifact.get("renderer_version", "")) == "2.5"
        slide = slide_by_id.get(str(visual.get("slide_id", "")), {})
        def annotation_contract(items: Any) -> list[dict[str, Any]]:
            if not isinstance(items, list):
                return []
            return [
                {
                    "annotation_type": item.get("annotation_type", "direct_annotation"),
                    "priority": item.get("priority", item.get("role", "secondary")),
                    "text": item.get("text", ""),
                }
                for item in items if isinstance(item, dict)
            ]

        privacy_match = (
            canonical(optional_mapping_contract(artifact.get("native_chart_privacy")))
            == canonical(optional_mapping_contract(visual.get("native_chart_privacy")))
        )
        if visual.get("render_mode") == "native_chart":
            privacy_match = privacy_match and bool(artifact.get("native_chart_privacy")) and bool(
                visual.get("native_chart_privacy")
            )
        presentation_checks = {
            "render_mode": artifact.get("render_mode") == visual.get("render_mode"),
            "render_visual_type": artifact.get("render_visual_type") == visual.get("render_visual_type"),
            "layout_variant": artifact.get("layout_variant") == slide.get("layout_variant"),
            "background_variant": artifact.get("background_variant") == slide.get("background_variant"),
            "annotation_specs": (
                canonical(annotation_contract(artifact.get("annotation_specs", [])))
                == canonical(annotation_contract(visual.get("annotation_specs", [])))
            ),
            "figure_rebuild_decision": (
                artifact.get("figure_rebuild_decision") == visual.get("figure_rebuild_decision")
            ),
            "native_chart_privacy": privacy_match,
        }
        presentation_match = all(presentation_checks.values())
        failed_presentation = ",".join(
            key for key, check_passed in presentation_checks.items() if not check_passed
        ) or "none"
        passed = source_match and row_match and data_match and version_match and presentation_match
        rows.append({
            "project_key": project, "visual_id": visual_id,
            "check_id": "v2_5_object_manifest_contract", "passed": passed,
            "details": (
                f"renderer={artifact.get('renderer_version')}; data={data_match}; "
                f"source={source_match}; rows={row_match}; presentation={presentation_match}; "
                f"presentation_failures={failed_presentation}; "
                f"objects={artifact.get('created_object_count')}"
            ),
        })
    return rows


def language_and_unicode_rows(
    v24: Any, project: str, candidate: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    language_rows: list[dict[str, Any]] = []
    unicode_rows: list[dict[str, Any]] = []
    pptx_strings: dict[int, str] = defaultdict(str)
    malformed_xml = 0
    try:
        with zipfile.ZipFile(candidate["pptx"]) as package:
            slide_names = sorted(
                (name for name in package.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
                key=lambda name: int(re.search(r"\d+", Path(name).stem).group()),
            )
            for index, name in enumerate(slide_names, 1):
                try:
                    root = ET.fromstring(package.read(name))
                except ET.ParseError:
                    malformed_xml += 1
                    continue
                pptx_strings[index] = " ".join(
                    node.text or "" for node in root.iter()
                    if node.tag.endswith("}t")
                )
    except (OSError, zipfile.BadZipFile) as exc:
        raise BlockedError(f"Cannot inspect PPTX Unicode content for {project}: {exc}") from exc

    def structured_entity_terms(slide: dict[str, Any]) -> set[str]:
        """Return exact source-bound row labels that may remain professional terms.

        v2.4 intentionally excludes structured evidence rows from the zh-CN
        prose denominator.  v2.5 measures the rendered PPTX, so the same
        exclusion must be applied to rendered study names, genes, pathways and
        similar source-bound entity labels.  Narrative sentences and node
        labels are not excluded and therefore still require localization.
        """
        terms: set[str] = set()

        def visit(value: Any, *, in_rows: bool = False) -> None:
            if isinstance(value, str):
                normalized = value.strip()
                if in_rows and normalized:
                    terms.add(normalized)
                return
            if isinstance(value, dict):
                for key, child in value.items():
                    visit(child, in_rows=in_rows or key == "rows")
                return
            if isinstance(value, list):
                for child in value:
                    visit(child, in_rows=in_rows)

        for visual in slide.get("visual_specs", []):
            visit(visual.get("payload", {}))
        return terms

    def rendered_language_metrics(slide: dict[str, Any], rendered: str) -> dict[str, Any]:
        # Registry markers are governed audit identifiers, not untranslated
        # prose. They must remain verbatim so they can be searched and traced.
        scrubbed = re.sub(r"\[[A-Z0-9_:-]+\]", "", rendered)
        scrubbed = re.sub(r"\bINFORMATION_REQUIRED\b", "", scrubbed)
        allowed_terms = set(v24.ALLOWED_LATIN_TERMS) | {
            "KDIGO", "TAPSE", "CI", "FDR", "NES", "log2FC", "pp", "h",
            "LV", "RV", "PA", "IVC", "MAP", "CVP", "RRT", "n=",
        } | structured_entity_terms(slide)
        for term in sorted(allowed_terms, key=len, reverse=True):
            scrubbed = re.sub(re.escape(term), "", scrubbed, flags=re.I)
        cjk = len(re.findall(r"[\u3400-\u9fff]", scrubbed))
        latin = len(re.findall(r"[A-Za-z]", scrubbed))
        coverage = cjk / max(1, cjk + latin)
        mojibake = bool(v24.MOJIBAKE_RE.search(rendered) or "\ufffd" in rendered)
        abnormal_spacing = bool(v24.ABNORMAL_SPACING_RE.search(rendered))
        title_has_cjk = bool(re.search(r"[\u3400-\u9fff]", str(slide.get("title", ""))))
        passed = (
            slide.get("language") == "zh-CN"
            and not mojibake
            and not abnormal_spacing
            and title_has_cjk
            and bool(rendered.strip())
        )
        return {
            "project_key": project,
            "slide_id": slide.get("slide_id", ""),
            "declared_language": slide.get("language", ""),
            "cjk_characters": cjk,
            "non_allowlisted_latin_characters": latin,
            "zh_cn_coverage": round(coverage, 4),
            "mojibake_detected": mojibake,
            "abnormal_spacing_detected": abnormal_spacing,
            "title_has_cjk": title_has_cjk,
            # Per-slide language rows expose coverage for review, while the
            # contractual >=90% threshold is applied at deck level below.
            "passed": passed,
        }

    for index, slide in enumerate(candidate["slides"], 1):
        rendered = pptx_strings.get(index, "")
        language = rendered_language_metrics(slide, rendered)
        language_rows.append(language)
        audience = " ".join([
            str(slide.get("title", "")), str(slide.get("narrative_purpose", "")),
            str(slide.get("key_message", "")),
            *[
                text
                for visual in slide.get("visual_specs", [])
                for text in v24.audience_payload_strings(visual.get("payload", {}))
            ],
        ])
        combined = audience + " " + rendered
        markers = sorted({marker for marker in MOJIBAKE_MARKERS if marker in combined})
        abnormal = bool(v24.ABNORMAL_SPACING_RE.search(combined))
        replacement = combined.count("\ufffd")
        passed = not markers and not abnormal and replacement == 0 and malformed_xml == 0
        unicode_rows.append({
            "project_key": project, "slide_number": index,
            "slide_id": slide.get("slide_id", ""),
            "replacement_character_count": replacement,
            "mojibake_markers": "|".join(markers),
            "abnormal_spacing": abnormal,
            "malformed_slide_xml_count": malformed_xml,
            "pptx_text_present": bool(rendered.strip()),
            "zh_cn_coverage": language.get("zh_cn_coverage", 0),
            "passed": passed,
        })
    return language_rows, unicode_rows


def source_hash_rows(
    v24: Any, project: str, project_root: Path,
    baseline: dict[str, Any], candidate: dict[str, Any],
) -> list[dict[str, Any]]:
    baseline_sources = {
        item.get("source_file"): item for item in baseline["graph"].get("source_registry", [])
    }
    candidate_sources = {
        item.get("source_file"): item for item in candidate["graph"].get("source_registry", [])
    }
    manifest = read_csv(candidate["output"] / "source_manifest.csv")
    integrity = v24.input_integrity(project, project_root, manifest)
    rows = []
    for row in integrity:
        relative = row["source_file"]
        before = baseline_sources.get(relative, {})
        after = candidate_sources.get(relative, {})
        cross_version = bool(before and after and before.get("sha256") == after.get("sha256"))
        rows.append({
            **row,
            "v2_4_sha256": before.get("sha256", ""),
            "v2_5_sha256": after.get("sha256", ""),
            "cross_version_unchanged": cross_version,
            "passed": as_bool(row.get("passed")) and cross_version,
        })
    return rows


def run_hash_manifest_rows(run_root: Path) -> list[dict[str, Any]]:
    before_rows = read_csv(run_root / "input_hashes_before.csv")
    after_rows = read_csv(run_root / "input_hashes_after.csv")
    before = {row.get("path", ""): row.get("sha256", "").casefold() for row in before_rows}
    after = {row.get("path", ""): row.get("sha256", "").casefold() for row in after_rows}
    if "" in before or "" in after:
        raise BlockedError("v2.5 input hash manifest contains an empty path")
    rows: list[dict[str, Any]] = []
    for path_value in sorted(set(before) | set(after)):
        path = Path(path_value)
        actual = hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else ""
        passed = (
            path_value in before and path_value in after
            and before[path_value] == after[path_value] == actual.casefold()
        )
        rows.append({
            "path": path_value, "sha256_before": before.get(path_value, ""),
            "sha256_after": after.get(path_value, ""), "sha256_current": actual,
            "path_unchanged": path_value in before and path_value in after,
            "passed": passed,
        })
    integrity_marker = read_json(run_root / "input_integrity.json")
    if not as_bool(integrity_marker.get("unchanged")) or integrity_marker.get("errors"):
        for row in rows:
            row["passed"] = False
    return rows


def optional_render_contract(candidate: dict[str, Any]) -> dict[str, dict[str, Any]]:
    exact = candidate["pptx"].with_suffix(".render_contract.json")
    if not exact.is_file():
        candidates = sorted(candidate["output"].glob("*.render_contract.json"))
        exact = candidates[0] if len(candidates) == 1 else exact
    if not exact.is_file():
        return {}
    value = read_json(exact)
    if isinstance(value, dict):
        for key in ("slides", "records", "slide_contracts"):
            if isinstance(value.get(key), list):
                value = value[key]
                break
    if not isinstance(value, list):
        return {}
    return {str(item.get("slide_id", "")): item for item in value if item.get("slide_id")}


def explicit_hero_ids(graph: dict[str, Any]) -> tuple[set[str], set[str]]:
    art = graph_art_direction(graph)
    slide_ids: set[str] = set()
    visual_ids: set[str] = set()
    for key in ("hero_slide_id", "primary_hero_slide_id"):
        if art.get(key):
            slide_ids.add(str(art[key]))
    for key in ("hero_slide_ids", "primary_hero_slide_ids"):
        value = art.get(key, [])
        if isinstance(value, list):
            slide_ids.update(map(str, value))
    for key in ("hero_visual_id", "primary_hero_visual_id"):
        if art.get(key):
            visual_ids.add(str(art[key]))
    for key in ("hero_visual_ids", "primary_hero_visual_ids"):
        value = art.get(key, [])
        if isinstance(value, list):
            visual_ids.update(map(str, value))
    return slide_ids, visual_ids


def art_direction_rows(
    project: str, candidate: dict[str, Any],
) -> tuple[
    list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]],
    list[dict[str, Any]], list[dict[str, Any]],
]:
    graph = candidate["graph"]
    slides = candidate["slides"]
    art = graph_art_direction(graph)
    rhythm = rhythm_entries(graph)
    rhythm_by_slide = {str(item.get("slide_id", "")): item for item in rhythm}
    render_by_slide = optional_render_contract(candidate)
    manifest_by_slide: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in candidate["manifest"]:
        manifest_by_slide[str(item.get("slide_id", ""))].append(item)
    explicit_slides, explicit_visuals = explicit_hero_ids(graph)

    variant_rows: list[dict[str, Any]] = []
    hero_rows: list[dict[str, Any]] = []
    annotation_rows: list[dict[str, Any]] = []
    sequence_variants: list[str] = []
    sequence_families: list[str] = []
    sequence_backgrounds: list[str] = []
    source_figures = 0
    native_charts = 0

    for index, slide in enumerate(slides, 1):
        slide_id = str(slide.get("slide_id", ""))
        rhythm_item = rhythm_by_slide.get(slide_id, {})
        render_item = render_by_slide.get(slide_id, {})
        artifacts = manifest_by_slide.get(slide_id, [])
        artifact_variant = next((str(item.get("layout_variant")) for item in artifacts if item.get("layout_variant")), "")
        artifact_background = next((str(item.get("background_variant")) for item in artifacts if item.get("background_variant")), "")
        variant = str(
            artifact_variant
            or render_item.get("layout_variant")
            or spec_value(slide, "layout_variant", "")
            or rhythm_item.get("layout_variant", "")
        )
        background = str(
            artifact_background
            or render_item.get("background_variant")
            or spec_value(slide, "background_variant", "")
            or rhythm_item.get("background_variant", "")
        )
        family = str(slide.get("layout_family", ""))
        sequence_variants.append(variant)
        sequence_families.append(family)
        sequence_backgrounds.append(background)
        rhythm_present = slide_id in rhythm_by_slide
        plan_variant = str(rhythm_item.get("layout_variant", variant))
        plan_background = str(rhythm_item.get("background_variant", background))
        background_is_unregistered_raster = bool(
            re.search(r"(?:image|photo|raster|bitmap)", background, re.I)
            and not slide.get("source_bindings")
        )
        variant_pass = bool(variant) and bool(background) and rhythm_present
        if plan_variant:
            variant_pass = variant_pass and variant == plan_variant
        if plan_background:
            variant_pass = variant_pass and background == plan_background
        variant_pass = variant_pass and not background_is_unregistered_raster
        variant_rows.append({
            "project_key": project, "slide_number": index, "slide_id": slide_id,
            "slide_role": slide.get("slide_role", ""), "layout_family": family,
            "layout_variant": variant, "planned_layout_variant": plan_variant,
            "background_variant": background,
            "planned_background_variant": plan_background,
            "visual_intensity": rhythm_item.get("visual_intensity", spec_value(slide, "visual_intensity", "")),
            "density_target": rhythm_item.get("density_target", spec_value(slide, "density_target", "")),
            "rhythm_entry_present": rhythm_present,
            "unregistered_raster_background": background_is_unregistered_raster,
            "passed": variant_pass,
        })

        visual_ids = {str(item.get("visual_id", "")) for item in slide.get("visual_specs", [])}
        result_like = slide.get("slide_role") == "result"
        hero_contract = slide.get("hero_visual", {}) if isinstance(slide.get("hero_visual"), dict) else {}
        hero_required = (
            slide_id in explicit_slides
            or bool(visual_ids & explicit_visuals)
            or as_bool(rhythm_item.get("is_hero"))
            or str(rhythm_item.get("hero_priority", "")).casefold() == "required"
            or as_bool(hero_contract.get("enabled"))
        )
        hero_candidates = []
        for artifact in artifacts:
            share = number(artifact.get("hero_area_ratio", artifact.get("hero_share", artifact.get("hero_ratio"))))
            if share is not None:
                hero_candidates.append((share, artifact))
        render_share = number(render_item.get("hero_area_ratio", render_item.get("hero_share", render_item.get("hero_ratio"))))
        if render_share is not None:
            hero_candidates.append((render_share, render_item))
        hero_share, hero_artifact = max(hero_candidates, key=lambda pair: pair[0]) if hero_candidates else (None, {})
        if hero_required:
            bounds = hero_artifact.get("hero_bounds") if isinstance(hero_artifact, dict) else None
            planned_ratio = number(hero_contract.get("target_area_ratio", rhythm_item.get("hero_area_ratio")))
            hero_pass = (
                hero_share is not None and 0.35 <= hero_share <= 0.65
                and planned_ratio is not None and 0.35 <= planned_ratio <= 0.65
                and abs(hero_share - planned_ratio) <= 0.08
                and int(hero_contract.get("max_primary_anchors", 0) or 0) == 1
                and int(hero_contract.get("max_auxiliary_annotations", 0) or 0) <= 2
            )
            if isinstance(bounds, dict):
                hero_pass = hero_pass and all(number(bounds.get(key)) is not None for key in ("x", "y", "w", "h"))
            hero_rows.append({
                "project_key": project, "slide_number": index, "slide_id": slide_id,
                "slide_role": slide.get("slide_role", ""),
                "explicit_hero": slide_id in explicit_slides or bool(visual_ids & explicit_visuals),
                "result_like": result_like,
                "hero_area_ratio": "" if hero_share is None else round(hero_share, 4),
                "planned_hero_area_ratio": "" if planned_ratio is None else round(planned_ratio, 4),
                "hero_bounds": canonical(bounds) if isinstance(bounds, dict) else "",
                "minimum_share": 0.35, "maximum_share": 0.65,
                "passed": hero_pass,
            })

        for visual in slide.get("visual_specs", []):
            visual_id = str(visual.get("visual_id", ""))
            artifact = candidate["manifest_by_visual"].get(visual_id, {})
            visual_type = str(visual.get("visual_type", ""))
            native_charts += int(as_bool(artifact.get("native_chart")) or artifact.get("render_mode") == "native_chart")
            source_figures += int(visual_type in SOURCE_FIGURE_TYPES)
            planned_annotations = spec_value(visual, "annotation_specs", spec_value(visual, "annotations", []))
            planned_count = len(planned_annotations) if isinstance(planned_annotations, list) else int(bool(planned_annotations))
            planned_primary = 0
            planned_secondary = 0
            planned_priority_valid = isinstance(planned_annotations, list)
            if isinstance(planned_annotations, list):
                for annotation in planned_annotations:
                    if not isinstance(annotation, dict):
                        planned_priority_valid = False
                        continue
                    priority = str(
                        annotation.get("priority", annotation.get("role", "secondary"))
                    ).strip().casefold()
                    if priority == "primary":
                        planned_primary += 1
                    elif priority == "secondary":
                        planned_secondary += 1
                    else:
                        planned_priority_valid = False
            actual_count = int(artifact.get("annotation_count", 0) or 0)
            primary = int(artifact.get("annotation_primary_count", 0) or 0)
            secondary = int(artifact.get("annotation_secondary_count", 0) or 0)
            # AnnotationSpec is an explicit presentation contract.  A chart is
            # allowed to have no callout; the global rule is a maximum of one
            # primary and one secondary annotation, not a mandatory primary
            # annotation on every chart/result.  Source-figure callouts remain
            # independently enforced by source_figure_rows().
            annotation_required = planned_count > 0
            annotation_pass = (
                planned_priority_valid
                and (not annotation_required or actual_count >= 1)
                and actual_count <= 2 and primary <= 1 and secondary <= 1
                and actual_count == planned_count
                and primary == planned_primary
                and secondary == planned_secondary
            )
            annotation_rows.append({
                "project_key": project, "slide_id": slide_id, "visual_id": visual_id,
                "visual_type": visual_type, "annotation_required": annotation_required,
                "planned_annotation_count": planned_count,
                "annotation_count": actual_count,
                "primary_annotation_count": primary,
                "secondary_annotation_count": secondary,
                "source_bound": bool(visual.get("source_bindings")),
                "passed": annotation_pass and bool(visual.get("source_bindings")),
            })

    plan_ids = {str(item.get("slide_id", "")) for item in rhythm}
    slide_ids = {str(slide.get("slide_id", "")) for slide in slides}
    max_same_family = consecutive_max(sequence_families)
    max_same_variant = consecutive_max(sequence_variants)
    max_same_silhouette = consecutive_max([
        f"{family}|{variant}" for family, variant in zip(sequence_families, sequence_variants, strict=True)
    ])
    card_count = sum(as_bool(slide.get("card_grid")) for slide in slides)
    text_only = sum(
        as_bool(slide.get("text_only")) and slide.get("slide_role") not in {"cover", "appendix", "section_divider"}
        for slide in slides
    )
    count = len(slides)
    rhythm_semantics_match = all(
        str(slide.get("background_variant", "")) == str(rhythm_by_slide.get(str(slide.get("slide_id", "")), {}).get("background_variant", ""))
        and str(slide.get("visual_intensity", "")) == str(rhythm_by_slide.get(str(slide.get("slide_id", "")), {}).get("visual_intensity", ""))
        and str(slide.get("density_target", "")) == str(rhythm_by_slide.get(str(slide.get("slide_id", "")), {}).get("density_target", ""))
        and as_bool(slide.get("hero_visual", {}).get("enabled")) == as_bool(rhythm_by_slide.get(str(slide.get("slide_id", "")), {}).get("is_hero"))
        for slide in slides
    )
    rhythm_pass = bool(art) and plan_ids == slide_ids and rhythm_semantics_match and max_same_silhouette <= 2
    rhythm_rows = [{
        "project_key": project, "slide_count": count,
        "art_direction_present": bool(art), "rhythm_plan_present": bool(rhythm),
        "rhythm_slide_coverage": round(len(plan_ids & slide_ids) / max(1, len(slide_ids)), 4),
        "layout_family_count": len(set(sequence_families)),
        "layout_variant_count": len(set(sequence_variants) - {""}),
        "background_variant_count": len(set(sequence_backgrounds) - {""}),
        "max_consecutive_same_family": max_same_family,
        "max_consecutive_same_variant": max_same_variant,
        "max_consecutive_same_silhouette": max_same_silhouette,
        "rhythm_semantics_match": rhythm_semantics_match,
        "passed": rhythm_pass,
    }]
    diversity_pass = (
        len(set(sequence_families)) >= 5
        and len(set(sequence_variants) - {""}) >= 4
        and len(set(sequence_backgrounds) - {""}) >= 2
        and card_count / max(1, count) <= 0.30
        and text_only / max(1, count) <= 0.15
        and max_same_silhouette <= 2
    )
    diversity_rows = [{
        "project_key": project, "slide_count": count,
        "layout_family_count": len(set(sequence_families)),
        "layout_variant_count": len(set(sequence_variants) - {""}),
        "background_variant_count": len(set(sequence_backgrounds) - {""}),
        "visual_type_count": len({str(visual.get("visual_type", "")) for visual in candidate["visuals"]}),
        "card_grid_count": card_count, "card_grid_ratio": round(card_count / max(1, count), 4),
        "text_only_count": text_only, "text_only_ratio": round(text_only / max(1, count), 4),
        "max_consecutive_same_silhouette": max_same_silhouette,
        "hero_slide_count": len([row for row in hero_rows if row["project_key"] == project]),
        "native_chart_count": native_charts, "source_figure_count": source_figures,
        "passed": diversity_pass,
    }]
    return rhythm_rows, variant_rows, hero_rows, annotation_rows, diversity_rows


def embedded_workbook_scan(package: zipfile.ZipFile) -> tuple[int, int, list[str]]:
    workbooks = [name for name in package.namelist() if name.startswith("ppt/embeddings/") and name.lower().endswith(".xlsx")]
    max_rows = 0
    strings: list[str] = []
    for workbook in workbooks:
        try:
            from io import BytesIO
            with zipfile.ZipFile(BytesIO(package.read(workbook))) as nested:
                for name in nested.namelist():
                    if name == "xl/sharedStrings.xml" or name.startswith("xl/worksheets/") and name.endswith(".xml"):
                        payload = nested.read(name)
                        try:
                            root = ET.fromstring(payload)
                            strings.extend(node.text or "" for node in root.iter() if node.text)
                            if name.startswith("xl/worksheets/"):
                                rows = [
                                    int(match.group(1)) for node in root.iter()
                                    if node.tag.endswith("}row")
                                    and (match := re.match(r"(\d+)", str(node.attrib.get("r", ""))))
                                ]
                                max_rows = max(max_rows, max(rows, default=0))
                        except ET.ParseError:
                            strings.append("MALFORMED_EMBEDDED_WORKBOOK_XML")
        except (KeyError, OSError, zipfile.BadZipFile):
            strings.append("UNREADABLE_EMBEDDED_WORKBOOK")
    return len(workbooks), max_rows, strings


def native_chart_inventory(project: str, candidate: dict[str, Any]) -> list[dict[str, Any]]:
    manifest_native = [
        item for item in candidate["manifest"]
        if as_bool(item.get("native_chart")) or item.get("render_mode") == "native_chart"
    ]
    try:
        with zipfile.ZipFile(candidate["pptx"]) as package:
            chart_xml = [name for name in package.namelist() if re.fullmatch(r"ppt/charts/chart\d+\.xml", name)]
            chart_rels = [
                name for name in package.namelist()
                if name.startswith("ppt/charts/_rels/") and name.endswith(".rels")
            ]
            external = 0
            chart_text: list[str] = []
            malformed = 0
            for name in [*chart_xml, *chart_rels]:
                payload = package.read(name)
                try:
                    root = ET.fromstring(payload)
                    chart_text.extend(node.text or "" for node in root.iter() if node.text)
                    external += sum(
                        str(node.attrib.get("TargetMode", "")).casefold() == "external"
                        for node in root.iter()
                    )
                except ET.ParseError:
                    malformed += 1
            workbook_count, max_rows, workbook_text = embedded_workbook_scan(package)
    except (OSError, zipfile.BadZipFile) as exc:
        raise BlockedError(f"Cannot inspect native charts in {candidate['pptx']}: {exc}") from exc
    inspected_text = " ".join([*chart_text, *workbook_text])
    suspicious = sorted(set(match.group(0) for match in SENSITIVE_CHART_RE.finditer(inspected_text)))
    internal_leak = bool(SOURCE_ID_RE.search(inspected_text) or SHA_RE.search(inspected_text) or LOCAL_PATH_RE.search(inspected_text))
    package_pass = external == 0 and malformed == 0 and not suspicious and not internal_leak and max_rows <= 250
    rows: list[dict[str, Any]] = []
    for item in manifest_native:
        visual = next(
            (visual for visual in candidate["visuals"] if visual.get("visual_id") == item.get("visual_id")),
            {},
        )
        privacy = item.get("native_chart_privacy") or visual.get("native_chart_privacy") or {}
        source_bound = bool(item.get("source_bindings"))
        data_scope = str(privacy.get("data_scope", ""))
        aggregate_source = data_scope in {"aggregate_visual_spec_only", "aggregate_study_level"}
        workbook_contract = item.get("chart_workbook_source") == "visual_spec_aggregate_rows"
        privacy_flags_pass = all(
            privacy.get(key) is False
            for key in (
                "patient_level_data_embedded", "cell_level_data_embedded",
                "absolute_source_path_embedded", "unrelated_fields_embedded",
            )
        )
        embedded_contract_rows = int(privacy.get("embedded_row_count", -1) or 0)
        editable = str(item.get("editability", "")).casefold() in {"full", "native", "editable"}
        series = int(item.get("chart_series_count", 0) or 0)
        categories = int(item.get("chart_category_count", 0) or 0)
        chart_shape_metadata_pass = (series > 0 and categories > 0) or (
            len(chart_xml) > 0 and embedded_contract_rows > 0
        )
        passed = (
            package_pass and source_bound and aggregate_source and privacy_flags_pass
            and workbook_contract and editable and embedded_contract_rows > 0
            and chart_shape_metadata_pass
        )
        rows.append({
            "project_key": project, "slide_id": item.get("slide_id", ""),
            "visual_id": item.get("visual_id", ""),
            "chart_type": item.get("chart_type", item.get("render_visual_type", visual.get("render_visual_type", ""))),
            "chart_series_count": series, "chart_category_count": categories,
            "chart_workbook_source": item.get("chart_workbook_source", data_scope),
            "chart_workbook_contract_pass": workbook_contract,
            "native_chart_data_scope": data_scope,
            "native_chart_contract_row_count": embedded_contract_rows,
            "privacy_flags_pass": privacy_flags_pass,
            "editability": item.get("editability", ""), "source_bound": source_bound,
            "package_chart_count": len(chart_xml), "embedded_workbook_count": workbook_count,
            "embedded_max_row": max_rows, "external_relationship_count": external,
            "sensitive_token_hits": "|".join(suspicious), "internal_audit_token_leak": internal_leak,
            "malformed_chart_xml": malformed, "passed": passed,
        })
    if not manifest_native:
        rows.append({
            "project_key": project, "slide_id": "", "visual_id": "", "chart_type": "",
            "chart_series_count": 0, "chart_category_count": 0, "chart_workbook_source": "",
            "chart_workbook_contract_pass": True,
            "native_chart_data_scope": "", "native_chart_contract_row_count": 0,
            "privacy_flags_pass": package_pass,
            "editability": "", "source_bound": False, "package_chart_count": len(chart_xml),
            "embedded_workbook_count": workbook_count, "embedded_max_row": max_rows,
            "external_relationship_count": external, "sensitive_token_hits": "|".join(suspicious),
            "internal_audit_token_leak": internal_leak, "malformed_chart_xml": malformed,
            "passed": package_pass and len(chart_xml) == 0,
        })
    return rows


def source_figure_rows(project: str, project_root: Path, candidate: dict[str, Any]) -> list[dict[str, Any]]:
    registry = {item.get("source_id"): item for item in candidate["graph"].get("source_registry", [])}
    rows: list[dict[str, Any]] = []
    for visual in candidate["visuals"]:
        if visual.get("visual_type") not in SOURCE_FIGURE_TYPES:
            continue
        artifact = candidate["manifest_by_visual"].get(str(visual.get("visual_id", "")), {})
        bindings = visual.get("source_bindings", [])
        figure_bindings = [
            binding for binding in bindings
            if str(registry.get(binding.get("source_id"), {}).get("source_role", "")).casefold() == "figure"
            or Path(str(binding.get("source_file", ""))).suffix.casefold() in {".png", ".jpg", ".jpeg", ".svg"}
        ]
        asset_value = (
            artifact.get("asset_path") or visual.get("asset_path")
            or visual.get("payload", {}).get("asset_path")
            or visual.get("payload", {}).get("figure_path")
            or visual.get("payload", {}).get("image_path")
        )
        asset_path = Path(str(asset_value)) if asset_value else Path()
        if asset_value and not asset_path.is_absolute():
            asset_path = (project_root / asset_path).resolve()
        source_hash_match = False
        asset_hash_match = False
        source_file = ""
        if figure_bindings:
            source_file = str(figure_bindings[0].get("source_file", ""))
            registered = registry.get(figure_bindings[0].get("source_id"), {})
            expected_path = (project_root / source_file).resolve()
            if expected_path.is_file():
                source_hash_match = hashlib.sha256(expected_path.read_bytes()).hexdigest() == registered.get("sha256")
            if not asset_value:
                asset_path = expected_path
            if asset_path.is_file() and expected_path.is_file():
                asset_hash_match = (
                    hashlib.sha256(asset_path.read_bytes()).hexdigest()
                    == hashlib.sha256(expected_path.read_bytes()).hexdigest()
                )
        decision = str(
            artifact.get("figure_rebuild_decision")
            or visual.get("figure_rebuild_decision") or ""
        )
        render_visual_type = str(
            artifact.get("render_visual_type")
            or visual.get("render_visual_type") or visual.get("visual_type", "")
        )
        render_mode = str(artifact.get("render_mode") or visual.get("render_mode") or "")
        treatment = str(artifact.get("source_figure_treatment", spec_value(visual, "source_figure_treatment", "")))
        callout = int(artifact.get("annotation_primary_count", 0) or 0) >= 1
        redraw = decision == "REDRAW_FROM_STRUCTURED_DATA"
        preserve = decision in {"PRESERVE_WITH_CALLOUT", "PRESERVE_SOURCE_FIGURE"}
        if redraw:
            privacy = artifact.get("native_chart_privacy") or visual.get("native_chart_privacy") or {}
            privacy_pass = all(
                privacy.get(key) is False
                for key in (
                    "patient_level_data_embedded", "cell_level_data_embedded",
                    "absolute_source_path_embedded", "unrelated_fields_embedded",
                )
            )
            passed = (
                bool(figure_bindings) and source_hash_match
                and render_mode == "native_chart"
                and render_visual_type not in SOURCE_FIGURE_TYPES
                and privacy_pass
            )
        elif preserve:
            privacy_pass = True
            passed = (
                bool(figure_bindings) and asset_path.is_file() and source_hash_match
                and asset_hash_match and bool(treatment) and callout
                and render_mode in {"source_figure", "svg"}
            )
        else:
            privacy_pass = False
            passed = False
        rows.append({
            "project_key": project, "slide_id": visual.get("slide_id", ""),
            "visual_id": visual.get("visual_id", ""), "visual_type": visual.get("visual_type", ""),
            "source_file": source_file, "asset_path": str(asset_path),
            "source_binding_present": bool(figure_bindings), "source_hash_match": source_hash_match,
            "asset_hash_match": asset_hash_match,
            "figure_rebuild_decision": decision,
            "render_visual_type": render_visual_type,
            "render_mode": render_mode,
            "source_figure_treatment": treatment,
            "native_callout_present": callout,
            "native_chart_privacy_pass": privacy_pass,
            "editability": artifact.get("editability", ""),
            "passed": passed,
        })
    return rows


def build_blinded_materials(
    v24: Any,
    baseline_run: dict[str, Any],
    candidate_run: dict[str, Any],
    projects: dict[str, tuple[dict[str, Any], dict[str, Any]]],
    audit: Path,
    benchmark: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    contact_root = benchmark / "contact_sheets"
    blind_root = benchmark / "blinded_ab_review"
    contact_root.mkdir(parents=True, exist_ok=True)
    blind_root.mkdir(parents=True, exist_ok=True)
    public_rows: list[dict[str, Any]] = []
    key_rows: list[dict[str, Any]] = []
    for project in sorted(projects):
        baseline, candidate = projects[project]
        pair_code = f"PAIR-{stable_code(baseline_run['root'].name, candidate_run['root'].name, project, length=6)}"
        candidate_first = int(stable_code(pair_code, "assignment", length=2), 16) % 2 == 0
        assignment = [
            ("Deck A", "v2.5", candidate) if candidate_first else ("Deck A", "v2.4", baseline),
            ("Deck B", "v2.4", baseline) if candidate_first else ("Deck B", "v2.5", candidate),
        ]
        public: dict[str, Any] = {"pair_code": pair_code}
        for deck_label, version, item in assignment:
            filename = f"{pair_code}_{deck_label.replace(' ', '_')}.png"
            destination = blind_root / filename
            v24.contact_sheet(item["pp_images"], destination, f"{pair_code} · {deck_label}")
            mirror = contact_root / filename
            v24.contact_sheet(item["pp_images"], mirror, f"{pair_code} · {deck_label}")
            key_rows.append({
                "pair_code": pair_code, "anonymous_deck": deck_label,
                "project_key": project, "version": version,
                "run_root": str(item["output"].parent), "output_dir": str(item["output"]),
                "contact_sheet": str(destination),
            })
            public[f"{deck_label.casefold().replace(' ', '_')}_contact_sheet"] = str(destination)
        public_rows.append(public)

    write_csv(
        blind_root / "review_index.csv", public_rows,
        ["pair_code", "deck_a_contact_sheet", "deck_b_contact_sheet"],
    )
    write_csv(
        audit / "blinding_key.csv", key_rows,
        ["pair_code", "anonymous_deck", "project_key", "version", "run_root", "output_dir", "contact_sheet"],
    )
    return public_rows, key_rows


def preserve_blank_score_form(path: Path, key_rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    expected_keys = [
        (str(item["pair_code"]), str(item["anonymous_deck"])) for item in key_rows
    ]
    if path.is_file():
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            existing_fields = reader.fieldnames or []
            existing = list(reader)
        missing_fields = [field for field in VISUAL_SCORE_FIELDS if field not in existing_fields]
        if missing_fields:
            raise BlockedError(
                f"Existing human score form has an incompatible schema; no rewrite was performed: {missing_fields}"
            )
    else:
        existing = []
    by_key: dict[tuple[str, str], dict[str, str]] = {}
    for row in existing:
        key = (row.get("pair_code", ""), row.get("anonymous_deck", ""))
        if not all(key) or key in by_key:
            raise BlockedError(f"Invalid or duplicate human score row in {path}: {key}")
        by_key[key] = row
    if path.is_file():
        if set(by_key) != set(expected_keys):
            raise BlockedError(
                "Existing human score form does not match the current blind key; no rewrite was performed"
            )
        return existing

    output_rows: list[dict[str, str]] = []
    for key in expected_keys:
        # Never infer or materialize a score.  Existing human-entered strings
        # are not present on first creation; every score cell stays empty.
        output_rows.append({field: "" for field in VISUAL_SCORE_FIELDS} | {
            "pair_code": key[0], "anonymous_deck": key[1],
        })
    write_csv(path, output_rows, VISUAL_SCORE_FIELDS)
    return output_rows


def completed_score(
    row: dict[str, Any],
    score_fields: tuple[tuple[str, float], ...] = SCORE_FIELDS,
) -> dict[str, Any] | None:
    values: dict[str, float] = {}
    for field, maximum in score_fields:
        parsed = number(row.get(field))
        if parsed is None or not (0 <= parsed <= maximum):
            return None
        values[field] = parsed
    reviewer = str(row.get("reviewer", "")).strip()
    review_date = str(row.get("review_date", "")).strip()
    if not reviewer or not review_date:
        return None
    return {**values, "total_100": sum(values.values()), "reviewer": reviewer, "review_date": review_date}


def per_deck_human_score(output: Path) -> dict[str, Any] | None:
    path = output / "manual_visual_review_form.csv"
    if not path.is_file():
        return None
    rows = read_csv(path)
    if len(rows) != 1:
        return None
    return completed_score(rows[0]) or completed_score(rows[0], LEGACY_V24_SCORE_FIELDS)


def collect_human_scores(
    score_rows: list[dict[str, str]],
    key_rows: list[dict[str, Any]],
    projects: dict[str, tuple[dict[str, Any], dict[str, Any]]],
) -> dict[tuple[str, str], dict[str, Any]]:
    key_map = {
        (str(item["pair_code"]), str(item["anonymous_deck"])):
        (str(item["project_key"]), str(item["version"]))
        for item in key_rows
    }
    scores: dict[tuple[str, str], dict[str, Any]] = {}
    for row in score_rows:
        mapped = key_map.get((row.get("pair_code", ""), row.get("anonymous_deck", "")))
        completed = completed_score(row)
        if mapped and completed:
            scores[mapped] = completed
    # Existing deck-level forms are accepted only when all dimensions,
    # reviewer, and review_date were genuinely filled.  Values are never
    # copied into the blinded score form or synthesized by this finalizer.
    for project, (baseline, candidate) in projects.items():
        for version, item in (("v2.4", baseline), ("v2.5", candidate)):
            if (project, version) not in scores:
                completed = per_deck_human_score(item["output"])
                if completed:
                    scores[(project, version)] = completed
    return scores


def human_gate(
    projects: Iterable[str], scores: dict[tuple[str, str], dict[str, Any]],
) -> tuple[bool, bool, list[dict[str, Any]], dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    improvements: list[float] = []
    all_complete = True
    for project in sorted(projects):
        baseline = scores.get((project, "v2.4"))
        candidate = scores.get((project, "v2.5"))
        complete = baseline is not None and candidate is not None
        all_complete &= complete
        improvement: float | None = None
        noninferior = False
        clarity_noninferior = False
        semantic_noninferior = False
        candidate_threshold = False
        if complete:
            improvement = candidate["total_100"] - baseline["total_100"]
            improvements.append(improvement)
            noninferior = candidate["total_100"] >= baseline["total_100"]
            clarity_noninferior = candidate["scientific_clarity_20"] >= baseline["scientific_clarity_20"]
            semantic_noninferior = candidate["visual_semantic_correctness_20"] >= baseline["visual_semantic_correctness_20"]
            candidate_threshold = candidate["total_100"] >= 80
        rows.append({
            "project_key": project, "v2_4_review_complete": baseline is not None,
            "v2_5_review_complete": candidate is not None,
            "v2_4_total": "" if baseline is None else baseline["total_100"],
            "v2_5_total": "" if candidate is None else candidate["total_100"],
            "improvement": "" if improvement is None else improvement,
            "v2_5_at_least_80": candidate_threshold,
            "total_noninferior": noninferior,
            "scientific_clarity_noninferior": clarity_noninferior,
            "visual_semantics_noninferior": semantic_noninferior,
        })
    mean_improvement = sum(improvements) / len(improvements) if improvements else None
    passed = bool(
        all_complete
        and mean_improvement is not None and mean_improvement >= 5
        and all(
            as_bool(row["v2_5_at_least_80"])
            and as_bool(row["total_noninferior"])
            and as_bool(row["scientific_clarity_noninferior"])
            and as_bool(row["visual_semantics_noninferior"])
            for row in rows
        )
    )
    summary = {
        "all_complete": all_complete,
        "mean_improvement": mean_improvement,
        "all_v2_5_at_least_80": all_complete and all(as_bool(row["v2_5_at_least_80"]) for row in rows),
        "all_noninferior": all_complete and all(as_bool(row["total_noninferior"]) for row in rows),
        "clarity_noninferior": all_complete and all(as_bool(row["scientific_clarity_noninferior"]) for row in rows),
        "semantic_noninferior": all_complete and all(as_bool(row["visual_semantics_noninferior"]) for row in rows),
    }
    return all_complete, passed, rows, summary


def markdown_gate_table(gates: dict[str, bool]) -> str:
    lines = ["| Gate | Result |", "|---|---|"]
    lines.extend(f"| {name} | {'PASS' if passed else 'FAIL'} |" for name, passed in gates.items())
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Finalize Academic PPT Workflow v2.5 art-direction remediation by "
            "comparing a frozen v2.4 run with a rendered v2.5 run. Human scores "
            "are read only when every dimension, reviewer, and date are present."
        )
    )
    parser.add_argument(
        "--v2-4-run-root", "--baseline-run-root", dest="v2_4_run_root",
        type=Path, required=True, help="Frozen v2.4 B2/regression run root",
    )
    parser.add_argument(
        "--v2-5-run-root", "--candidate-run-root", dest="v2_5_run_root",
        type=Path, required=True, help="Rendered v2.5 candidate run root",
    )
    parser.add_argument("--suite", type=Path, required=True, help="Isolated four-project synthetic suite root")
    parser.add_argument("--audit-root", type=Path, required=True, help="Destination audit/v2_5 directory")
    parser.add_argument("--benchmark-root", type=Path, required=True, help="Destination benchmark/v2_5 directory")
    parser.add_argument("--python", type=Path, default=Path(sys.executable), help="Python runtime used for regression tests")
    args = parser.parse_args()

    baseline_root = args.v2_4_run_root.resolve()
    candidate_root = args.v2_5_run_root.resolve()
    suite = args.suite.resolve()
    audit = args.audit_root.resolve()
    benchmark = args.benchmark_root.resolve()
    audit.mkdir(parents=True, exist_ok=True)
    benchmark.mkdir(parents=True, exist_ok=True)

    try:
        v24 = _load_v24()
        baseline_run = load_run(baseline_root, expected_version="v2.4")
        candidate_run = load_run(candidate_root, expected_version="v2.5")
        if set(baseline_run["decks"]) != set(candidate_run["decks"]):
            raise BlockedError(
                "v2.4/v2.5 project sets differ: "
                f"v2.4={sorted(baseline_run['decks'])}, v2.5={sorted(candidate_run['decks'])}"
            )
        project_pairs: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {}
        for project in sorted(candidate_run["decks"]):
            project_root = suite / project
            if not project_root.is_dir():
                raise BlockedError(f"Isolated suite project is missing: {project_root}")
            project_pairs[project] = (
                load_project(baseline_run, project),
                load_project(candidate_run, project),
            )
    except Exception as exc:
        (audit / "final_status.md").write_text(
            "# Academic PPT Workflow v2.5 final status\n\n"
            "`BLOCKED`\n\n"
            f"- Reason: {type(exc).__name__}: {exc}\n"
            "- Validation material: synthetic regression suite; real-world validation was not performed.\n",
            encoding="utf-8",
        )
        print("STATUS=BLOCKED")
        print(f"REASON={type(exc).__name__}: {exc}")
        return 2

    science_rows: list[dict[str, Any]] = []
    object_rows: list[dict[str, Any]] = []
    binding_rows: list[dict[str, Any]] = []
    footer_rows: list[dict[str, Any]] = []
    language_rows: list[dict[str, Any]] = []
    unicode_rows: list[dict[str, Any]] = []
    coverage_rows: list[dict[str, Any]] = []
    omission_rows: list[dict[str, Any]] = []
    semantic_rows: list[dict[str, Any]] = []
    rhythm_rows: list[dict[str, Any]] = []
    variant_rows: list[dict[str, Any]] = []
    hero_rows: list[dict[str, Any]] = []
    annotation_rows: list[dict[str, Any]] = []
    native_rows: list[dict[str, Any]] = []
    source_figure_report: list[dict[str, Any]] = []
    diversity_rows: list[dict[str, Any]] = []
    typography_rows: list[dict[str, Any]] = []
    geometry_rows: list[dict[str, Any]] = []
    perceptual_rows: list[dict[str, Any]] = []
    integrity_rows: list[dict[str, Any]] = []
    claim_gt_rows: list[dict[str, Any]] = []
    conflict_gt_rows: list[dict[str, Any]] = []
    unresolved_gt_rows: list[dict[str, Any]] = []
    render_rows: list[dict[str, Any]] = []
    runtime_rows: list[dict[str, Any]] = []
    run_hash_rows: list[dict[str, Any]] = []
    invented_claims = 0

    # The formal ground-truth stage starts only after both run roots and both
    # render indexes have been loaded and all four output sets were verified.
    # Preserve earlier-access evidence; do not claim this is the first read.
    prelog = audit / "ground_truth_access_prelog.md"
    prior_access = prelog.read_text(encoding="utf-8") if prelog.is_file() else "No earlier access was recorded in this audit root."
    (audit / "ground_truth_access_log.md").write_text(
        "# Ground-truth access log\n\n"
        f"- Formal comparison started UTC: {datetime.now(timezone.utc).isoformat()}\n"
        "- Stage: POST_GENERATION_GROUND_TRUTH_EVALUATION\n"
        "- Preconditions: v2.4 and v2.5 generation markers, PowerPoint indexes, LibreOffice indexes, PPTX files, and PNG sets were present.\n"
        "- This log does not claim that formal comparison was the first process or human read.\n\n"
        "## Earlier access record\n\n" + prior_access + "\n",
        encoding="utf-8",
    )

    try:
        for project, (baseline, candidate) in project_pairs.items():
            kind = str(candidate["deck"].get("kind", candidate["graph"].get("kind", "")))
            project_root = suite / project
            science_rows.extend(compare_scientific_contract(project, baseline, candidate))
            project_object_rows = validate_object_contract(project, candidate)
            object_rows.extend(project_object_rows)

            # Object-manifest row evidence is the rendered authority.
            for visual in candidate["visuals"]:
                artifact = candidate["manifest_by_visual"][str(visual["visual_id"])]
                visual["rendered_row_count"] = artifact.get("rendered_row_count")
                visual["rendered_row_keys"] = artifact.get("rendered_row_keys", [])

            project_binding, project_footer = v24.binding_checks(
                project, kind, project_root, candidate["graph"], candidate["output"], candidate["geometry"]
            )
            binding_rows.extend(project_binding)
            footer_rows.extend(project_footer)

            project_language, project_unicode = language_and_unicode_rows(v24, project, candidate)
            language_rows.extend(project_language)
            unicode_rows.extend(project_unicode)

            project_coverage = read_csv(candidate["output"] / "narrative_coverage_report.csv")
            for item in project_coverage:
                record: dict[str, Any] = {"project_key": project, **item}
                record["passed"] = not as_bool(item.get("omitted")) and bool(item.get("rendered_slide"))
                coverage_rows.append(record)
            omission_rows.extend(v24.omission_checks(project, candidate["slides"], project_coverage))
            project_semantic = v24.semantic_checks(project, kind, candidate["slides"])
            semantic_rows.extend(project_semantic)

            project_rhythm, project_variants, project_heroes, project_annotations, project_diversity = art_direction_rows(project, candidate)
            rhythm_rows.extend(project_rhythm)
            variant_rows.extend(project_variants)
            hero_rows.extend(project_heroes)
            annotation_rows.extend(project_annotations)
            diversity_rows.extend(project_diversity)
            native_rows.extend(native_chart_inventory(project, candidate))
            source_figure_report.extend(source_figure_rows(project, project_root, candidate))

            project_typography, _ = classify_typography_v25(project, candidate["pptx"], candidate["slides"], v24)
            typography_rows.extend(project_typography)
            project_geometry, project_perceptual = actual_geometry_v25(
                v24, project, candidate["geometry"], candidate["slides"], candidate["pp_images"], candidate["lo_images"]
            )
            geometry_rows.extend(project_geometry)
            perceptual_rows.extend(project_perceptual)

            integrity_rows.extend(source_hash_rows(v24, project, project_root, baseline, candidate))

            gt = project_root / "expected_ground_truth"
            if not gt.is_dir() or gt.is_relative_to(project_root / "input"):
                raise BlockedError(f"Ground truth is absent or not isolated for {project}")
            claims, conflicts, unresolved, invented = v24.ground_truth_compare(project, candidate["output"], gt)
            claim_gt_rows.extend(claims)
            conflict_gt_rows.extend(conflicts)
            unresolved_gt_rows.extend(unresolved)
            invented_claims += invented

            pp = candidate["pp"]
            lo = candidate["lo"]
            slide_count = candidate["slide_count"]
            pp_pass = (
                as_bool(pp.get("powerpoint_open"))
                and int(pp.get("pptx_slide_count", -1)) == slide_count
                and int(pp.get("png_count", -1)) == slide_count
                and as_bool(pp.get("pdf_exists", True))
            )
            lo_pass = int(lo.get("exit_code", -1)) == 0 and int(lo.get("pdf_page_count", -1)) == slide_count
            render_rows.append({
                "project_key": project, "slide_count": slide_count,
                "powerpoint_open": as_bool(pp.get("powerpoint_open")),
                "powerpoint_slide_count": pp.get("pptx_slide_count", ""),
                "powerpoint_png_count": pp.get("png_count", ""),
                "powerpoint_pdf_exists": as_bool(pp.get("pdf_exists", True)),
                "powerpoint_passed": pp_pass,
                "libreoffice_exit_code": lo.get("exit_code", ""),
                "libreoffice_page_count": lo.get("pdf_page_count", ""),
                "libreoffice_passed": lo_pass,
            })

            baseline_seconds = number(baseline["deck"].get("generation_seconds"))
            candidate_seconds = number(candidate["deck"].get("generation_seconds"))
            ratio = candidate_seconds / baseline_seconds if baseline_seconds and candidate_seconds is not None else math.inf
            runtime_rows.append({
                "project_key": project,
                "v2_4_generation_seconds": "" if baseline_seconds is None else baseline_seconds,
                "v2_5_generation_seconds": "" if candidate_seconds is None else candidate_seconds,
                "runtime_ratio": ratio, "limit_ratio": 1.5,
                "passed": ratio <= 1.5,
            })
    except Exception as exc:
        (audit / "final_status.md").write_text(
            "# Academic PPT Workflow v2.5 final status\n\n"
            "`BLOCKED`\n\n"
            f"- Reason: {type(exc).__name__}: {exc}\n"
            "- Validation material: synthetic regression suite; real-world validation was not performed.\n",
            encoding="utf-8",
        )
        print("STATUS=BLOCKED")
        print(f"REASON={type(exc).__name__}: {exc}")
        return 2

    try:
        run_hash_rows = run_hash_manifest_rows(candidate_root)
    except Exception as exc:
        (audit / "final_status.md").write_text(
            "# Academic PPT Workflow v2.5 final status\n\n"
            "`BLOCKED`\n\n"
            f"- Reason: input-hash evidence unavailable: {type(exc).__name__}: {exc}\n"
            "- Validation material: synthetic regression suite; real-world validation was not performed.\n",
            encoding="utf-8",
        )
        print("STATUS=BLOCKED")
        print(f"REASON=input-hash evidence unavailable: {type(exc).__name__}: {exc}")
        return 2

    # Detailed evidence is written before gate aggregation so failures remain
    # auditable instead of being reduced to one final status string.
    write_csv(audit / "frozen_scientific_contract.csv", science_rows,
              ["project_key", "check_id", "passed", "baseline_digest", "candidate_digest", "details"])
    write_csv(audit / "object_manifest_contract.csv", object_rows,
              ["project_key", "visual_id", "check_id", "passed", "details"])
    write_csv(audit / "source_binding_accuracy.csv", binding_rows, v24.CSV_FIELDS["source_binding_accuracy.csv"])
    write_csv(audit / "source_footer_regression.csv", footer_rows,
              ["project_key", "slide_id", "expected_short_labels", "rendered_footer", "footer_match", "notes_match", "passed"])
    write_csv(audit / "language_compliance.csv", language_rows, v24.CSV_FIELDS["language_compliance.csv"])
    write_csv(audit / "unicode_typography_report.csv", unicode_rows,
              ["project_key", "slide_number", "slide_id", "replacement_character_count", "mojibake_markers", "abnormal_spacing", "malformed_slide_xml_count", "pptx_text_present", "zh_cn_coverage", "passed"])
    write_csv(audit / "narrative_coverage_report.csv", coverage_rows, v24.CSV_FIELDS["narrative_coverage_report.csv"])
    write_csv(audit / "omission_manifest.csv", omission_rows, v24.CSV_FIELDS["omission_manifest.csv"])
    write_csv(audit / "chart_semantic_qa.csv", semantic_rows, v24.CSV_FIELDS["chart_semantic_qa.csv"])
    write_csv(audit / "deck_rhythm_report.csv", rhythm_rows,
              ["project_key", "slide_count", "art_direction_present", "rhythm_plan_present", "rhythm_slide_coverage", "layout_family_count", "layout_variant_count", "background_variant_count", "max_consecutive_same_family", "max_consecutive_same_variant", "max_consecutive_same_silhouette", "rhythm_semantics_match", "passed"])
    write_csv(audit / "visual_variant_report.csv", variant_rows,
              ["project_key", "slide_number", "slide_id", "slide_role", "layout_family", "layout_variant", "planned_layout_variant", "background_variant", "planned_background_variant", "visual_intensity", "density_target", "rhythm_entry_present", "unregistered_raster_background", "passed"])
    write_csv(audit / "hero_visual_report.csv", hero_rows,
              ["project_key", "slide_number", "slide_id", "slide_role", "explicit_hero", "result_like", "hero_area_ratio", "planned_hero_area_ratio", "hero_bounds", "minimum_share", "maximum_share", "passed"])
    write_csv(audit / "annotation_report.csv", annotation_rows,
              ["project_key", "slide_id", "visual_id", "visual_type", "annotation_required", "planned_annotation_count", "annotation_count", "primary_annotation_count", "secondary_annotation_count", "source_bound", "passed"])
    write_csv(audit / "native_chart_inventory.csv", native_rows,
              ["project_key", "slide_id", "visual_id", "chart_type", "chart_series_count", "chart_category_count", "chart_workbook_source", "chart_workbook_contract_pass", "native_chart_data_scope", "native_chart_contract_row_count", "privacy_flags_pass", "editability", "source_bound", "package_chart_count", "embedded_workbook_count", "embedded_max_row", "external_relationship_count", "sensitive_token_hits", "internal_audit_token_leak", "malformed_chart_xml", "passed"])
    write_csv(audit / "source_figure_rebuild_report.csv", source_figure_report,
              ["project_key", "slide_id", "visual_id", "visual_type", "source_file", "asset_path", "source_binding_present", "source_hash_match", "asset_hash_match", "figure_rebuild_decision", "render_visual_type", "render_mode", "source_figure_treatment", "native_callout_present", "native_chart_privacy_pass", "editability", "passed"])
    write_csv(audit / "visual_diversity_report.csv", diversity_rows,
              ["project_key", "slide_count", "layout_family_count", "layout_variant_count", "background_variant_count", "visual_type_count", "card_grid_count", "card_grid_ratio", "text_only_count", "text_only_ratio", "max_consecutive_same_silhouette", "hero_slide_count", "native_chart_count", "source_figure_count", "passed"])
    write_csv(audit / "typography_report.csv", typography_rows, v24.CSV_FIELDS["typography_report.csv"])
    write_csv(audit / "render_geometry_report.csv", geometry_rows, v24.CSV_FIELDS["geometry_qa.csv"])
    write_csv(audit / "perceptual_visual_qa.csv", perceptual_rows, v24.CSV_FIELDS["perceptual_visual_qa.csv"])
    write_csv(audit / "input_integrity.csv", integrity_rows,
              ["project_key", "source_file", "sha256_before", "sha256_after", "v2_4_sha256", "v2_5_sha256", "unchanged", "registered", "ground_truth_excluded", "cross_version_unchanged", "passed"])
    write_csv(audit / "run_input_hash_integrity.csv", run_hash_rows,
              ["path", "sha256_before", "sha256_after", "sha256_current", "path_unchanged", "passed"])
    write_csv(audit / "ground_truth_claim_comparison.csv", claim_gt_rows, v24.CSV_FIELDS["ground_truth_claim_comparison.csv"])
    write_csv(audit / "ground_truth_conflict_comparison.csv", conflict_gt_rows, v24.CSV_FIELDS["ground_truth_conflict_comparison.csv"])
    write_csv(audit / "ground_truth_unresolved_comparison.csv", unresolved_gt_rows, v24.CSV_FIELDS["ground_truth_unresolved_comparison.csv"])
    write_csv(audit / "render_compatibility.csv", render_rows,
              ["project_key", "slide_count", "powerpoint_open", "powerpoint_slide_count", "powerpoint_png_count", "powerpoint_pdf_exists", "powerpoint_passed", "libreoffice_exit_code", "libreoffice_page_count", "libreoffice_passed"])
    write_csv(audit / "runtime_comparison.csv", runtime_rows,
              ["project_key", "v2_4_generation_seconds", "v2_5_generation_seconds", "runtime_ratio", "limit_ratio", "passed"])

    _, blind_key = build_blinded_materials(v24, baseline_run, candidate_run, project_pairs, audit, benchmark)
    score_rows = preserve_blank_score_form(benchmark / "visual_scores.csv", blind_key)
    human_scores = collect_human_scores(score_rows, blind_key, project_pairs)
    human_complete, human_pass, human_rows, human_summary = human_gate(project_pairs, human_scores)
    write_csv(audit / "human_visual_score_gate.csv", human_rows,
              ["project_key", "v2_4_review_complete", "v2_5_review_complete", "v2_4_total", "v2_5_total", "improvement", "v2_5_at_least_80", "total_noninferior", "scientific_clarity_noninferior", "visual_semantics_noninferior"])

    tests_pass, _ = v24.run_tests(args.python.resolve(), audit)

    binding_pass = all(as_bool(row.get("passed")) for row in binding_rows)
    field_rows = [row for row in binding_rows if str(row.get("check_type", "")).startswith("field_lineage:")]
    field_pass = bool(field_rows) and all(as_bool(row.get("passed")) for row in field_rows)
    language_by_project: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in language_rows:
        language_by_project[str(row["project_key"])].append(row)
    language_scores = {
        project: sum(int(row["cjk_characters"]) for row in rows)
        / max(1, sum(int(row["cjk_characters"]) + int(row["non_allowlisted_latin_characters"]) for row in rows))
        for project, rows in language_by_project.items()
    }
    language_pass = (
        len(language_scores) == 4
        and all(score >= 0.90 for score in language_scores.values())
        and all(as_bool(row.get("passed")) for row in language_rows)
    )
    coverage_pass = bool(coverage_rows) and all(as_bool(row.get("passed")) for row in coverage_rows)
    no_silent_omission = not any(as_bool(row.get("silent")) for row in omission_rows)
    semantic_failures = [row for row in semantic_rows if row.get("severity") == "error" and not as_bool(row.get("passed"))]
    claim_recall = sum(row.get("classification") == "true_positive" for row in claim_gt_rows) / max(1, len(claim_gt_rows))
    conflict_recall = sum(as_bool(row.get("recalled")) for row in conflict_gt_rows) / max(1, len(conflict_gt_rows))
    unresolved_recall = sum(as_bool(row.get("recalled")) for row in unresolved_gt_rows) / max(1, len(unresolved_gt_rows))
    geometry_pass = bool(geometry_rows) and all(as_bool(row.get("passed")) for row in geometry_rows)
    pp_pass = all(as_bool(row.get("powerpoint_passed")) for row in render_rows)
    lo_pass = all(as_bool(row.get("libreoffice_passed")) for row in render_rows)
    runtime_mean_baseline = sum(number(row.get("v2_4_generation_seconds")) or 0 for row in runtime_rows) / max(1, len(runtime_rows))
    runtime_mean_candidate = sum(number(row.get("v2_5_generation_seconds")) or 0 for row in runtime_rows) / max(1, len(runtime_rows))
    runtime_ratio = runtime_mean_candidate / runtime_mean_baseline if runtime_mean_baseline > 0 else math.inf
    runtime_pass = runtime_ratio <= 1.5
    actual_native_charts = [row for row in native_rows if row.get("visual_id")]

    automated_gates = {
        "frozen scientific contract": all(as_bool(row.get("passed")) for row in science_rows),
        "renderer object and row contract": all(as_bool(row.get("passed")) for row in object_rows),
        "slide-level source/claim bindings = 100%": binding_pass,
        "field-level lineage = 100%": field_pass,
        "source footer and notes identity": all(as_bool(row.get("passed")) for row in footer_rows),
        "zh-CN language compliance >=90%": language_pass,
        "Unicode and typography encoding": all(as_bool(row.get("passed")) for row in unicode_rows),
        "must_include coverage = 100%": coverage_pass,
        "silent truncation = 0": no_silent_omission,
        "chart scientific semantics": not semantic_failures,
        "hero visual contract": bool(hero_rows) and all(as_bool(row.get("passed")) for row in hero_rows),
        "deck rhythm contract": all(as_bool(row.get("passed")) for row in rhythm_rows),
        "visual variants and background treatments": all(as_bool(row.get("passed")) for row in variant_rows),
        "source-bound annotations": all(as_bool(row.get("passed")) for row in annotation_rows),
        # A deck with no planned native chart is a valid hybrid-rendering
        # outcome.  Its package-level sentinel row must still prove that no
        # orphan chart parts, external links, or sensitive workbook payloads
        # are present.  Native charts, when present, retain every row-level
        # aggregate/privacy/editability gate above.
        "native chart aggregate/privacy contract": bool(native_rows) and all(as_bool(row.get("passed")) for row in native_rows),
        "source figure and callout contract": bool(source_figure_report) and all(as_bool(row.get("passed")) for row in source_figure_report),
        "visual diversity": all(as_bool(row.get("passed")) for row in diversity_rows),
        "minimum typography": all(as_bool(row.get("passed")) for row in typography_rows),
        "severe overlap/clipping/footer intrusion = 0": geometry_pass,
        "PowerPoint open/page parity = 100%": pp_pass,
        "LibreOffice page parity": lo_pass,
        "claim/conflict/unresolved recall = 100%": claim_recall == conflict_recall == unresolved_recall == 1.0,
        "invented claims = 0": invented_claims == 0,
        "input hashes unchanged": (
            all(as_bool(row.get("passed")) for row in integrity_rows)
            and bool(run_hash_rows) and all(as_bool(row.get("passed")) for row in run_hash_rows)
        ),
        "runtime <=1.5x v2.4": runtime_pass,
        "native PptxGenJS only": (
            candidate_run["generation"].get("backend") == "native_pptxgenjs"
            and not as_bool(candidate_run["generation"].get("external_skill_used"))
        ),
        "regression tests": tests_pass,
    }
    automated_pass = all(automated_gates.values())
    critical_scientific_gates = (
        automated_gates["frozen scientific contract"],
        automated_gates["renderer object and row contract"],
        binding_pass, field_pass, coverage_pass, no_silent_omission,
        not semantic_failures,
        claim_recall == 1.0, conflict_recall == 1.0, unresolved_recall == 1.0,
        invented_claims == 0,
        automated_gates["input hashes unchanged"],
        automated_gates["native PptxGenJS only"],
    )
    critical_scientific_regression = not all(critical_scientific_gates)
    completed_human_failure = any(
        as_bool(row.get("v2_4_review_complete"))
        and as_bool(row.get("v2_5_review_complete"))
        and not all([
            as_bool(row.get("v2_5_at_least_80")),
            as_bool(row.get("total_noninferior")),
            as_bool(row.get("scientific_clarity_noninferior")),
            as_bool(row.get("visual_semantics_noninferior")),
        ])
        for row in human_rows
    )

    if automated_pass and human_complete and human_pass:
        status = "WORKFLOW_ART_DIRECTION_REMEDIATION_PASSED"
    elif automated_pass and not human_complete and not completed_human_failure:
        status = "READY_FOR_HUMAN_VISUAL_REVIEW"
    elif critical_scientific_regression:
        status = "WORKFLOW_ART_DIRECTION_REMEDIATION_FAILED"
    else:
        status = "WORKFLOW_ART_DIRECTION_REMEDIATION_PARTIAL"

    (audit / "art_direction_audit.md").write_text(
        "# v2.5 art-direction audit\n\n"
        f"- Candidate decks: {len(project_pairs)} isolated synthetic projects\n"
        f"- Art-direction specs present: {sum(as_bool(row['art_direction_present']) for row in rhythm_rows)}/{len(rhythm_rows)}\n"
        f"- Deck rhythm plans present: {sum(as_bool(row['rhythm_plan_present']) for row in rhythm_rows)}/{len(rhythm_rows)}\n"
        f"- Hero gates passed: {sum(as_bool(row['passed']) for row in hero_rows)}/{len(hero_rows)}\n"
        f"- Visual variant/background gates passed: {sum(as_bool(row['passed']) for row in variant_rows)}/{len(variant_rows)}\n"
        f"- Annotation gates passed: {sum(as_bool(row['passed']) for row in annotation_rows)}/{len(annotation_rows)}\n"
        f"- Native charts audited: {len(actual_native_charts)}\n"
        f"- Source figures audited: {len(source_figure_report)}\n"
        "- External Skill introduced: 0\n"
        "- Production backend: native_pptxgenjs\n\n"
        "The audit evaluates evidence-bound art direction, not decorative richness. Synthetic success is not real-world scientific validation.\n",
        encoding="utf-8",
    )
    (audit / "scientific_regression_report.md").write_text(
        "# v2.5 scientific regression report\n\n"
        f"- Critical scientific regression: {1 if critical_scientific_regression else 0}\n"
        f"- Frozen contract checks: {sum(as_bool(row['passed']) for row in science_rows)}/{len(science_rows)}\n"
        f"- Source/claim binding checks: {sum(as_bool(row['passed']) for row in binding_rows)}/{len(binding_rows)}\n"
        f"- Field lineage checks: {sum(as_bool(row['passed']) for row in field_rows)}/{len(field_rows)}\n"
        f"- Claim recall: {claim_recall:.1%}\n"
        f"- Conflict recall: {conflict_recall:.1%}\n"
        f"- Unresolved recall: {unresolved_recall:.1%}\n"
        f"- Invented canonical claims: {invented_claims}\n"
        f"- Semantic error failures: {len(semantic_failures)}\n"
        "- Dataset classification: synthetic regression suite only; not real research materials.\n",
        encoding="utf-8",
    )
    (audit / "runtime_report.md").write_text(
        "# v2.5 runtime report\n\n"
        f"- v2.4 mean generation time: {runtime_mean_baseline:.3f} s/deck\n"
        f"- v2.5 mean generation time: {runtime_mean_candidate:.3f} s/deck\n"
        f"- Ratio: {runtime_ratio:.2f}x\n"
        f"- Gate <=1.5x: {'PASS' if runtime_pass else 'FAIL'}\n",
        encoding="utf-8",
    )
    mean_improvement_text = (
        "not assessed" if human_summary["mean_improvement"] is None
        else f"{human_summary['mean_improvement']:.1f} points"
    )
    (audit / "human_visual_review_status.md").write_text(
        "# Human A/B visual review status\n\n"
        f"- Complete paired reviews: {sum(row['v2_4_review_complete'] and row['v2_5_review_complete'] for row in human_rows)}/4\n"
        f"- All v2.5 totals >=80: {human_summary['all_v2_5_at_least_80']}\n"
        f"- Every v2.5 total non-inferior to v2.4: {human_summary['all_noninferior']}\n"
        f"- Mean improvement: {mean_improvement_text}\n"
        f"- Scientific clarity non-inferior: {human_summary['clarity_noninferior']}\n"
        f"- Visual semantic correctness non-inferior: {human_summary['semantic_noninferior']}\n"
        "- No score is generated or inferred by the finalizer. Blank cells remain blank.\n",
        encoding="utf-8",
    )

    final_text = (
        "# Academic PPT Workflow v2.5 final status\n\n"
        f"`{status}`\n\n"
        "## Automated promotion gates\n\n"
        + markdown_gate_table(automated_gates)
        + "\n\n## Human visual gate\n\n"
        f"- Review complete: {'YES' if human_complete else 'NO'}\n"
        f"- Promotion thresholds passed: {'YES' if human_pass else 'NO / NOT YET ASSESSED'}\n"
        "- Required: every v2.5 deck >=80; no deck below v2.4; mean improvement >=5; scientific clarity and visual semantic correctness non-inferior.\n\n"
        "## Governance boundary\n\n"
        "- External Skill introduced: 0\n"
        "- Production backend: native_pptxgenjs\n"
        "- Validation material: synthetic regression suite only\n"
        "- Real-world validation: pending\n"
        "- v2 overall state: PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING\n"
    )
    if not human_complete and automated_pass:
        final_text += "\nAutomated gates passed, but the workflow cannot be marked PASSED until genuine human A/B scoring is complete.\n"
    (audit / "final_status.md").write_text(final_text, encoding="utf-8")

    print(f"STATUS={status}")
    print(f"CRITICAL_SCIENTIFIC_REGRESSION={int(critical_scientific_regression)}")
    print(f"CLAIM_RECALL={claim_recall:.4f}")
    print(f"CONFLICT_RECALL={conflict_recall:.4f}")
    print(f"UNRESOLVED_RECALL={unresolved_recall:.4f}")
    print(f"HUMAN_REVIEW_COMPLETE={int(human_complete)}")
    print(f"RUNTIME_RATIO={runtime_ratio:.4f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
