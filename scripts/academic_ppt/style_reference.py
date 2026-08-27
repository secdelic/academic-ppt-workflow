"""Privacy-preserving PowerPoint reference-style inspection.

This module intentionally implements only read-only OPC/OOXML inspection.  It
does not extract an archive to disk, follow external relationships, execute
embedded content, or return slide/notes/alternative text, media bytes, or chart
values.  The resulting profile is therefore suitable for the workflow's
``style-only`` reference mode and is also the safe lower layer for the other
declared reference modes.

The implementation is clean-room and uses only Python's standard library.  If
Pillow is already available it is used only to read intrinsic image dimensions;
image data is never returned or written.
"""

from __future__ import annotations

import hashlib
import io
import json
import math
import posixpath
import re
import urllib.parse
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

REL_NS = NS["pr"]
REL_TYPE_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
IMAGE_REL_TYPE = REL_TYPE_BASE + "image"
SLIDE_LAYOUT_REL_TYPE = REL_TYPE_BASE + "slideLayout"
SLIDE_MASTER_REL_TYPE = REL_TYPE_BASE + "slideMaster"
THEME_REL_TYPE = REL_TYPE_BASE + "theme"

ALLOWED_REFERENCE_MODES = (
    "style-only",
    "template-fill",
    "content-reference",
    "protected-reference",
)

MAX_XML_MEMBER_BYTES = 32 * 1024 * 1024
MAX_IMAGE_MEMBER_BYTES = 256 * 1024 * 1024
MAX_PACKAGE_ENTRIES = 50_000
MAX_TOTAL_UNCOMPRESSED_BYTES = 2 * 1024 * 1024 * 1024


class ReferenceStyleError(ValueError):
    """Raised when a reference package cannot be inspected safely."""


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_zip_names(package: zipfile.ZipFile) -> set[str]:
    infos = package.infolist()
    if len(infos) > MAX_PACKAGE_ENTRIES:
        raise ReferenceStyleError(
            f"Reference package has too many entries ({len(infos)}); inspection refused"
        )
    total = sum(max(0, item.file_size) for item in infos)
    if total > MAX_TOTAL_UNCOMPRESSED_BYTES:
        raise ReferenceStyleError(
            f"Reference package expands to {total} bytes; inspection refused"
        )
    names: set[str] = set()
    for info in infos:
        name = info.filename.replace("\\", "/")
        normalized = posixpath.normpath(name).lstrip("/")
        if (
            not name
            or name.startswith("/")
            or normalized == ".."
            or normalized.startswith("../")
        ):
            raise ReferenceStyleError(f"Unsafe OPC member path: {name!r}")
        names.add(normalized)
    return names


def _read_xml(
    package: zipfile.ZipFile,
    names: set[str],
    member: str,
    *,
    required: bool = False,
) -> ET.Element | None:
    member = posixpath.normpath(member).lstrip("/")
    if member not in names:
        if required:
            raise ReferenceStyleError(f"Required OOXML part is missing: {member}")
        return None
    info = package.getinfo(member)
    if info.file_size > MAX_XML_MEMBER_BYTES:
        raise ReferenceStyleError(f"XML part is too large to inspect safely: {member}")
    try:
        return ET.fromstring(package.read(member))
    except (ET.ParseError, KeyError, RuntimeError, zipfile.BadZipFile) as exc:
        raise ReferenceStyleError(f"Invalid XML in OOXML part {member}: {exc}") from exc


def _relationships_part(source_part: str) -> str:
    if not source_part:
        return "_rels/.rels"
    directory, filename = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", filename + ".rels")


def _source_part_for_relationships(relationships_part: str) -> str | None:
    if relationships_part == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(relationships_part)
    if not directory.endswith("/_rels") or not filename.endswith(".rels"):
        return None
    return posixpath.join(directory[: -len("/_rels")], filename[: -len(".rels")])


def _resolve_internal_target(source_part: str, target: str) -> str | None:
    """Resolve an OPC relationship target without reading outside the ZIP."""

    parsed = urllib.parse.urlsplit(target)
    if parsed.scheme or parsed.netloc:
        return None
    path = urllib.parse.unquote(parsed.path).replace("\\", "/")
    if path.startswith("/"):
        candidate = posixpath.normpath(path.lstrip("/"))
    else:
        candidate = posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), path)
        )
    if candidate in {"", ".", ".."} or candidate.startswith("../"):
        return None
    return candidate.lstrip("/")


def _relationship_rows(
    package: zipfile.ZipFile,
    names: set[str],
    source_part: str,
) -> list[dict[str, str]]:
    relationships_part = _relationships_part(source_part)
    root = _read_xml(package, names, relationships_part)
    if root is None:
        return []
    rows: list[dict[str, str]] = []
    for element in root.findall(f"{{{REL_NS}}}Relationship"):
        relationship_id = element.get("Id", "")
        relationship_type = element.get("Type", "")
        target = element.get("Target", "")
        target_mode = element.get("TargetMode", "")
        external = target_mode.lower() == "external"
        resolved = None if external else _resolve_internal_target(source_part, target)
        rows.append(
            {
                "id": relationship_id,
                "type": relationship_type,
                "target_mode": "External" if external else "Internal",
                "target": target,
                "resolved_target": resolved or "",
            }
        )
    return rows


def _redacted_external_warning(
    source_part: str, relationship: dict[str, str]
) -> dict[str, str]:
    target = relationship.get("target", "")
    parsed = urllib.parse.urlsplit(target)
    return {
        "source_part": source_part or "[package]",
        "relationship_id": relationship.get("id", ""),
        "relationship_type": relationship.get("type", ""),
        "target_mode": "External",
        "target_scheme": parsed.scheme.lower() or "unspecified",
        "target_sha256": hashlib.sha256(target.encode("utf-8")).hexdigest(),
        "action": "NOT_FOLLOWED",
    }


def _all_external_relationship_warnings(
    package: zipfile.ZipFile, names: set[str]
) -> list[dict[str, str]]:
    warnings: list[dict[str, str]] = []
    for member in sorted(name for name in names if name.endswith(".rels")):
        source = _source_part_for_relationships(member)
        if source is None:
            continue
        for relationship in _relationship_rows(package, names, source):
            if relationship["target_mode"] == "External":
                warnings.append(_redacted_external_warning(source, relationship))
    return warnings


def _int_or_unknown(raw: str | None) -> int | str:
    if raw is None:
        return "UNKNOWN"
    try:
        return int(raw)
    except (TypeError, ValueError):
        return "UNKNOWN"


def _rounded_ratio(width: int | str, height: int | str) -> float | str:
    if not isinstance(width, int) or not isinstance(height, int) or height <= 0:
        return "UNKNOWN"
    return round(width / height, 4)


def _geometry(element: ET.Element) -> dict[str, int | float | str]:
    transform = element.find(".//a:xfrm", NS)
    if transform is None:
        return {
            "x_emu": "UNKNOWN",
            "y_emu": "UNKNOWN",
            "width_emu": "UNKNOWN",
            "height_emu": "UNKNOWN",
            "aspect_ratio": "UNKNOWN",
        }
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)
    x = _int_or_unknown(offset.get("x") if offset is not None else None)
    y = _int_or_unknown(offset.get("y") if offset is not None else None)
    width = _int_or_unknown(extent.get("cx") if extent is not None else None)
    height = _int_or_unknown(extent.get("cy") if extent is not None else None)
    return {
        "x_emu": x,
        "y_emu": y,
        "width_emu": width,
        "height_emu": height,
        "aspect_ratio": _rounded_ratio(width, height),
    }


def _text_margins(shape: ET.Element) -> dict[str, int | str]:
    body = shape.find("p:txBody/a:bodyPr", NS)
    if body is None:
        return {
            "left_emu": "INHERITED_OR_UNKNOWN",
            "right_emu": "INHERITED_OR_UNKNOWN",
            "top_emu": "INHERITED_OR_UNKNOWN",
            "bottom_emu": "INHERITED_OR_UNKNOWN",
        }
    return {
        "left_emu": _int_or_unknown(body.get("lIns")),
        "right_emu": _int_or_unknown(body.get("rIns")),
        "top_emu": _int_or_unknown(body.get("tIns")),
        "bottom_emu": _int_or_unknown(body.get("bIns")),
    }


def _placeholder_inventory(root: ET.Element | None) -> list[dict[str, Any]]:
    if root is None:
        return []
    placeholders: list[dict[str, Any]] = []
    for shape in root.findall(".//p:sp", NS):
        placeholder = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if placeholder is None:
            continue
        kind = placeholder.get("type", "body")
        placeholders.append(
            {
                "placeholder_type": kind,
                "placeholder_index": placeholder.get("idx", "INHERITED_OR_UNKNOWN"),
                "placeholder_size": placeholder.get("sz", "full"),
                "orientation": placeholder.get("orient", "horz"),
                "geometry": _geometry(shape),
                "margins": _text_margins(shape),
            }
        )
    return placeholders


def _footer_rules(root: ET.Element | None) -> dict[str, Any]:
    if root is None:
        return {
            "date_time": "UNKNOWN",
            "footer": "UNKNOWN",
            "slide_number": "UNKNOWN",
            "placeholder_types_present": [],
        }
    header_footer = root.find(".//p:hf", NS)
    placeholder_types = sorted(
        {
            placeholder.get("type", "body")
            for placeholder in root.findall(".//p:ph", NS)
            if placeholder.get("type") in {"dt", "ftr", "sldNum", "hdr"}
        }
    )

    def setting(attribute: str, placeholder_type: str) -> bool | str:
        if header_footer is not None and header_footer.get(attribute) is not None:
            return header_footer.get(attribute) not in {"0", "false", "False"}
        if placeholder_type in placeholder_types:
            return True
        return "INHERITED_OR_UNKNOWN"

    return {
        "date_time": setting("dt", "dt"),
        "footer": setting("ftr", "ftr"),
        "slide_number": setting("sldNum", "sldNum"),
        "header": setting("hdr", "hdr"),
        "placeholder_types_present": placeholder_types,
    }


def _layout_role(layout_type: str) -> str:
    role_map = {
        "title": "title",
        "ctrTitle": "title",
        "titleOnly": "title_only",
        "secHead": "section",
        "obj": "title_and_content",
        "tx": "title_and_content",
        "twoObj": "two_content",
        "twoTxTwoObj": "comparison",
        "twoObjAndTx": "comparison",
        "objAndTwoObj": "comparison",
        "objTx": "content_with_caption",
        "txAndObj": "content_with_caption",
        "picTx": "picture_with_caption",
        "blank": "blank",
    }
    return role_map.get(layout_type, "custom_or_unknown")


def _theme_fonts(theme_root: ET.Element | None) -> dict[str, Any]:
    if theme_root is None:
        return {"major": "UNKNOWN", "minor": "UNKNOWN"}

    def font_group(name: str) -> dict[str, Any] | str:
        group = theme_root.find(f".//a:fontScheme/a:{name}Font", NS)
        if group is None:
            return "UNKNOWN"
        result: dict[str, Any] = {
            "latin": "UNKNOWN",
            "east_asian": "UNKNOWN",
            "complex_script": "UNKNOWN",
            "script_overrides": {},
        }
        mappings = {
            "latin": "latin",
            "ea": "east_asian",
            "cs": "complex_script",
        }
        for tag, key in mappings.items():
            element = group.find(f"a:{tag}", NS)
            if element is not None and element.get("typeface"):
                result[key] = element.get("typeface")
        overrides: dict[str, str] = {}
        for element in group.findall("a:font", NS):
            script = element.get("script")
            typeface = element.get("typeface")
            if script and typeface:
                overrides[script] = typeface
        result["script_overrides"] = dict(sorted(overrides.items()))
        return result

    return {"major": font_group("major"), "minor": font_group("minor")}


def _color_token(element: ET.Element) -> dict[str, str] | None:
    local_name = element.tag.rsplit("}", 1)[-1]
    if local_name == "srgbClr" and element.get("val"):
        return {"type": "srgb", "value": element.get("val", "").upper()}
    if local_name == "sysClr":
        value = element.get("lastClr") or element.get("val")
        if value:
            return {"type": "system", "value": value.upper()}
    if local_name == "schemeClr" and element.get("val"):
        return {"type": "scheme", "value": element.get("val", "")}
    if local_name == "prstClr" and element.get("val"):
        return {"type": "preset", "value": element.get("val", "")}
    if local_name == "scrgbClr":
        channels = [element.get(channel) for channel in ("r", "g", "b")]
        if all(value is not None for value in channels):
            return {"type": "scrgb", "value": ",".join(channels)}  # type: ignore[arg-type]
    return None


def _theme_colors(theme_root: ET.Element | None) -> list[dict[str, str]]:
    if theme_root is None:
        return []
    scheme = theme_root.find(".//a:clrScheme", NS)
    if scheme is None:
        return []
    colors: list[dict[str, str]] = []
    for slot in list(scheme):
        slot_name = slot.tag.rsplit("}", 1)[-1]
        for child in list(slot):
            token = _color_token(child)
            if token is not None:
                colors.append({"slot": slot_name, **token})
                break
    return colors


def _chart_color_tokens(
    package: zipfile.ZipFile, names: set[str]
) -> dict[str, Any]:
    """Read chart style/color *tokens*, never chart data/value parts."""

    counts: Counter[tuple[str, str]] = Counter()
    inspected_chart_parts = 0
    inspected_style_parts = 0
    for member in sorted(names):
        is_chart = bool(re.fullmatch(r"ppt/charts/chart\d+\.xml", member))
        is_style = bool(
            re.fullmatch(r"ppt/charts/(?:style|colors)\d+\.xml", member)
        )
        if not (is_chart or is_style):
            continue
        root = _read_xml(package, names, member)
        if root is None:
            continue
        if is_chart:
            inspected_chart_parts += 1
        else:
            inspected_style_parts += 1
        for element in root.iter():
            token = _color_token(element)
            if token is not None:
                counts[(token["type"], token["value"])] += 1
    return {
        "inspected_chart_part_count": inspected_chart_parts,
        "inspected_style_part_count": inspected_style_parts,
        "tokens": [
            {"type": kind, "value": value, "occurrences": count}
            for (kind, value), count in sorted(counts.items())
        ],
        "chart_values_returned": False,
    }


def _table_styles(
    package: zipfile.ZipFile, names: set[str]
) -> dict[str, Any]:
    root = _read_xml(package, names, "ppt/tableStyles.xml")
    if root is None:
        return {"default_style_id": "UNKNOWN", "style_ids": []}
    style_ids = sorted(
        {
            style_id
            for element in root.findall(".//a:tblStyle", NS)
            if (style_id := element.get("styleId"))
        }
    )
    return {
        "default_style_id": root.get("def", "UNKNOWN"),
        "style_ids": style_ids,
    }


def _image_dimensions(data: bytes) -> tuple[int, int] | None:
    try:
        from PIL import Image  # type: ignore

        with Image.open(io.BytesIO(data)) as image:
            width, height = image.size
            return int(width), int(height)
    except (ImportError, OSError, ValueError):
        return None


def _image_aspect_patterns(
    package: zipfile.ZipFile, names: set[str]
) -> tuple[list[dict[str, Any]], list[str]]:
    """Return aggregate, non-content image geometry patterns."""

    counts: Counter[tuple[str, str, str]] = Counter()
    warnings: list[str] = []
    media_dimensions: dict[str, tuple[int, int] | None] = {}
    for slide_part in sorted(
        name for name in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
    ):
        slide_root = _read_xml(package, names, slide_part)
        if slide_root is None:
            continue
        relationships = {
            row["id"]: row
            for row in _relationship_rows(package, names, slide_part)
            if row["target_mode"] == "Internal"
        }
        for picture in slide_root.findall(".//p:pic", NS):
            shape_geometry = _geometry(picture)
            shape_ratio = shape_geometry["aspect_ratio"]
            blip = picture.find(".//a:blip", NS)
            relationship_id = (
                blip.get(f"{{{NS['r']}}}embed") if blip is not None else None
            )
            target = ""
            if relationship_id and relationship_id in relationships:
                relationship = relationships[relationship_id]
                if relationship["type"] == IMAGE_REL_TYPE:
                    target = relationship["resolved_target"]
            intrinsic_ratio: float | str = "UNKNOWN"
            if target and target in names:
                if target not in media_dimensions:
                    info = package.getinfo(target)
                    if info.file_size > MAX_IMAGE_MEMBER_BYTES:
                        media_dimensions[target] = None
                        warnings.append(
                            "An image exceeded the safe dimension-inspection limit; "
                            "its intrinsic ratio is UNKNOWN."
                        )
                    else:
                        media_dimensions[target] = _image_dimensions(package.read(target))
                dimensions = media_dimensions[target]
                if dimensions and dimensions[1] > 0:
                    intrinsic_ratio = round(dimensions[0] / dimensions[1], 4)
            source_rectangle = picture.find(".//a:srcRect", NS)
            if source_rectangle is not None and any(
                source_rectangle.get(side) not in {None, "0"}
                for side in ("l", "t", "r", "b")
            ):
                fit = "cropped"
            elif (
                isinstance(shape_ratio, float)
                and isinstance(intrinsic_ratio, float)
                and not math.isclose(shape_ratio, intrinsic_ratio, rel_tol=0.02)
            ):
                fit = "stretched_or_clipped"
            else:
                fit = "uncropped_or_unknown"
            counts[(str(shape_ratio), str(intrinsic_ratio), fit)] += 1
    patterns: list[dict[str, Any]] = []
    for (shape_ratio, intrinsic_ratio, fit), count in sorted(counts.items()):
        patterns.append(
            {
                "shape_aspect_ratio": (
                    float(shape_ratio) if shape_ratio != "UNKNOWN" else "UNKNOWN"
                ),
                "intrinsic_aspect_ratio": (
                    float(intrinsic_ratio)
                    if intrinsic_ratio != "UNKNOWN"
                    else "UNKNOWN"
                ),
                "fit_pattern": fit,
                "occurrences": count,
            }
        )
    return patterns, warnings


def _resolve_first_relationship(
    relationships: Iterable[dict[str, str]], relationship_type: str
) -> str | None:
    for row in relationships:
        if (
            row["type"] == relationship_type
            and row["target_mode"] == "Internal"
            and row["resolved_target"]
        ):
            return row["resolved_target"]
    return None


def _part_profile(
    package: zipfile.ZipFile,
    names: set[str],
    part: str,
    *,
    kind: str,
) -> dict[str, Any]:
    root = _read_xml(package, names, part, required=True)
    assert root is not None
    relationships = _relationship_rows(package, names, part)
    placeholders = _placeholder_inventory(root)
    if kind == "layout":
        layout_type = root.get("type", "custom_or_unknown")
        role = _layout_role(layout_type)
    else:
        layout_type = "NOT_APPLICABLE"
        role = kind
    result: dict[str, Any] = {
        "part": part,
        "part_sha256": hashlib.sha256(package.read(part)).hexdigest(),
        "kind": kind,
        "layout_type": layout_type,
        "common_page_role": role,
        "show_master_shapes": root.get("showMasterSp", "INHERITED_OR_UNKNOWN"),
        "preserve": root.get("preserve", "INHERITED_OR_UNKNOWN"),
        "placeholders": placeholders,
        "footer_and_page_number_rules": _footer_rules(root),
    }
    if kind == "layout":
        result["master_part"] = (
            _resolve_first_relationship(relationships, SLIDE_MASTER_REL_TYPE)
            or "UNKNOWN"
        )
    if kind == "master":
        result["theme_part"] = (
            _resolve_first_relationship(relationships, THEME_REL_TYPE) or "UNKNOWN"
        )
        result["layout_parts"] = sorted(
            {
                row["resolved_target"]
                for row in relationships
                if row["type"] == SLIDE_LAYOUT_REL_TYPE
                and row["target_mode"] == "Internal"
                and row["resolved_target"]
            }
        )
    return result


def _presentation_dimensions(root: ET.Element) -> dict[str, Any]:
    slide_size = root.find("p:sldSz", NS)
    if slide_size is None:
        return {
            "width_emu": "UNKNOWN",
            "height_emu": "UNKNOWN",
            "aspect_ratio": "UNKNOWN",
            "declared_type": "UNKNOWN",
        }
    width = _int_or_unknown(slide_size.get("cx"))
    height = _int_or_unknown(slide_size.get("cy"))
    return {
        "width_emu": width,
        "height_emu": height,
        "aspect_ratio": _rounded_ratio(width, height),
        "declared_type": slide_size.get("type", "custom_or_unknown"),
    }


def _content_reference_manifest(
    package: zipfile.ZipFile, names: set[str]
) -> dict[str, Any]:
    """Return only hashes/counts; never return user content."""

    categories = {
        "slides": r"ppt/slides/slide\d+\.xml",
        "notes": r"ppt/notesSlides/notesSlide\d+\.xml",
        "charts": r"ppt/charts/chart\d+\.xml",
        "media": r"ppt/media/.+",
    }
    result: dict[str, Any] = {}
    for category, pattern in categories.items():
        members = sorted(name for name in names if re.fullmatch(pattern, name))
        result[category] = {
            "part_count": len(members),
            "aggregate_sha256": hashlib.sha256(
                "".join(
                    hashlib.sha256(package.read(member)).hexdigest()
                    for member in members
                ).encode("ascii")
            ).hexdigest(),
        }
    result["content_returned"] = False
    return result


def parse_reference_style(
    path: str | Path,
    mode: str = "style-only",
) -> dict[str, Any]:
    """Inspect a ``.pptx``/``.potx`` reference without modifying or exposing it.

    The four supported modes are policy declarations.  All four use the same
    privacy-preserving parser; ``template-fill`` additionally exposes a
    placeholder/layout map and ``content-reference`` exposes only aggregate
    content hashes/counts.  No mode in this module returns private content.
    """

    source = Path(path)
    normalized_mode = mode.strip().lower()
    if normalized_mode not in ALLOWED_REFERENCE_MODES:
        raise ReferenceStyleError(
            f"Unsupported reference mode {mode!r}; expected one of "
            + ", ".join(ALLOWED_REFERENCE_MODES)
        )
    if source.suffix.lower() not in {".pptx", ".potx"}:
        raise ReferenceStyleError("Reference style input must be .pptx or .potx")
    if not source.is_file():
        raise ReferenceStyleError(f"Reference file does not exist: {source}")

    before_hash = _sha256_file(source)
    try:
        package = zipfile.ZipFile(source, mode="r")
    except (OSError, zipfile.BadZipFile) as exc:
        raise ReferenceStyleError(f"Reference is not a readable OOXML ZIP: {exc}") from exc

    with package:
        try:
            names = _safe_zip_names(package)
            presentation = _read_xml(
                package, names, "ppt/presentation.xml", required=True
            )
            assert presentation is not None
            _read_xml(package, names, "[Content_Types].xml", required=True)
            external_warnings = _all_external_relationship_warnings(package, names)

            theme_parts = sorted(
                name
                for name in names
                if re.fullmatch(r"ppt/theme/theme\d+\.xml", name)
            )
            themes: list[dict[str, Any]] = []
            for theme_part in theme_parts:
                theme_root = _read_xml(package, names, theme_part, required=True)
                themes.append(
                    {
                        "part": theme_part,
                        "part_sha256": hashlib.sha256(
                            package.read(theme_part)
                        ).hexdigest(),
                        "fonts": _theme_fonts(theme_root),
                        "colors": _theme_colors(theme_root),
                    }
                )

            master_parts = sorted(
                name
                for name in names
                if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", name)
            )
            layout_parts = sorted(
                name
                for name in names
                if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", name)
            )
            masters = [
                _part_profile(package, names, part, kind="master")
                for part in master_parts
            ]
            layouts = [
                _part_profile(package, names, part, kind="layout")
                for part in layout_parts
            ]
            roles = Counter(layout["common_page_role"] for layout in layouts)

            title_body_positions: list[dict[str, Any]] = []
            for layout in layouts:
                for placeholder in layout["placeholders"]:
                    placeholder_type = placeholder["placeholder_type"]
                    if placeholder_type in {
                        "title",
                        "ctrTitle",
                        "body",
                        "obj",
                        "subTitle",
                    }:
                        title_body_positions.append(
                            {
                                "layout_part": layout["part"],
                                "role": layout["common_page_role"],
                                "placeholder_type": placeholder_type,
                                "geometry": placeholder["geometry"],
                                "margins": placeholder["margins"],
                            }
                        )

            image_patterns, image_warnings = _image_aspect_patterns(
                package, names
            )
            profile: dict[str, Any] = {
                "schema_version": "2.0",
                "reference_mode": normalized_mode,
                "source": {
                    "extension": source.suffix.lower(),
                    "size_bytes": source.stat().st_size,
                    "sha256": before_hash,
                    "read_only_verified": "PENDING_POST_READ_CHECK",
                },
                "privacy_contract": {
                    "slide_text_returned": False,
                    "notes_text_returned": False,
                    "alternative_text_returned": False,
                    "media_bytes_returned": False,
                    "chart_values_returned": False,
                    "external_relationships_followed": False,
                    "archive_extracted_to_disk": False,
                },
                "slide_dimensions": _presentation_dimensions(presentation),
                "theme_fonts": (
                    themes[0]["fonts"] if themes else "UNKNOWN"
                ),
                "theme_colors": (
                    themes[0]["colors"] if themes else []
                ),
                "themes": themes,
                "slide_masters": masters,
                "slide_layouts": layouts,
                "placeholder_and_geometry_summary": {
                    "master_placeholder_count": sum(
                        len(master["placeholders"]) for master in masters
                    ),
                    "layout_placeholder_count": sum(
                        len(layout["placeholders"]) for layout in layouts
                    ),
                    "title_and_body_positions": title_body_positions,
                },
                "common_page_roles": [
                    {"role": role, "layout_count": count}
                    for role, count in sorted(roles.items())
                ],
                "chart_color_tokens": _chart_color_tokens(package, names),
                "table_styles": _table_styles(package, names),
                "footer_and_page_number_rules": {
                    "show_special_placeholders_on_title_slide": presentation.get(
                        "showSpecialPlsOnTitleSld", "INHERITED_OR_UNKNOWN"
                    ),
                    "masters": [
                        {
                            "part": master["part"],
                            "rules": master["footer_and_page_number_rules"],
                        }
                        for master in masters
                    ],
                    "layouts": [
                        {
                            "part": layout["part"],
                            "rules": layout["footer_and_page_number_rules"],
                        }
                        for layout in layouts
                    ],
                },
                "image_aspect_ratio_patterns": image_patterns,
                "external_relationship_warnings": external_warnings,
                "warnings": image_warnings,
                "unknown_policy": (
                    "UNKNOWN and INHERITED_OR_UNKNOWN are preserved; values are "
                    "never guessed."
                ),
            }
            if normalized_mode == "template-fill":
                profile["native_template_map"] = {
                    "masters": [
                        {
                            "part": master["part"],
                            "layout_parts": master["layout_parts"],
                        }
                        for master in masters
                    ],
                    "layouts": [
                        {
                            "part": layout["part"],
                            "master_part": layout["master_part"],
                            "role": layout["common_page_role"],
                            "placeholders": layout["placeholders"],
                        }
                        for layout in layouts
                    ],
                }
            if normalized_mode == "content-reference":
                profile["content_reference_manifest"] = _content_reference_manifest(
                    package, names
                )
            if normalized_mode == "protected-reference":
                profile["protection"] = {
                    "content_use": "PROHIBITED",
                    "asset_reuse": "PROHIBITED",
                    "style_metadata_only": True,
                }
        except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
            raise ReferenceStyleError(
                f"Reference package could not be inspected safely: {exc}"
            ) from exc

    after_hash = _sha256_file(source)
    if after_hash != before_hash:
        raise ReferenceStyleError(
            "Reference file changed during read-only inspection; result discarded"
        )
    profile["source"]["read_only_verified"] = True
    profile["source"]["sha256_after"] = after_hash
    return profile


def reference_style_import(
    path: str | Path, mode: str = "style-only"
) -> dict[str, Any]:
    """Canonical public alias used by the v2 workflow."""

    return parse_reference_style(path, mode=mode)


def extract_reference_style(
    path: str | Path, mode: str = "style-only"
) -> dict[str, Any]:
    """Compatibility alias for callers that use extraction terminology."""

    return parse_reference_style(path, mode=mode)


def write_reference_style_profile(
    path: str | Path,
    output_path: str | Path,
    mode: str = "style-only",
) -> dict[str, Any]:
    """Write a new JSON profile without overwriting an existing file."""

    profile = parse_reference_style(path, mode=mode)
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("x", encoding="utf-8") as handle:
        json.dump(profile, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return profile


def write_style_profile(
    path: str | Path,
    output_path: str | Path,
    mode: str = "style-only",
) -> dict[str, Any]:
    """Concise workflow-facing alias for :func:`write_reference_style_profile`."""

    return write_reference_style_profile(path, output_path, mode=mode)


__all__ = [
    "ALLOWED_REFERENCE_MODES",
    "ReferenceStyleError",
    "extract_reference_style",
    "parse_reference_style",
    "reference_style_import",
    "write_reference_style_profile",
    "write_style_profile",
]
