"""Canonical KEEP presentation and editable-semantic manifests.

The manifest is a preservation contract, not a second presentation IR.  It
reads the rendered OOXML package, normalizes only evidence-backed PowerPoint
serialization behavior, and preserves every user-visible/editable property.
"""

from __future__ import annotations

import hashlib
import io
import json
import posixpath
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable, Mapping

from lxml import etree

from .utils import load_yaml_compatible


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

_OBJECT_NAMES = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart"}
_MEDIA_REL_TYPES = {"image", "audio", "video", "media"}
_COL_ID_URI = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"
_ROW_ID_URI = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}"
_ALLOWLIST_IDS = {
    "xml_namespace_prefix",
    "xml_attribute_order",
    "xml_serialization_whitespace",
    "relationship_id_same_target",
    "empty_effect_list",
    "table_column_id_extension",
    "table_row_id_extension",
    "table_cell_default_vertical_margins",
    "shape_native_id_powerpoint_normalization",
    "media_part_renaming_same_binary",
    "registered_slide_number_placeholder_value",
}


class KeepPreservationError(RuntimeError):
    """Raised when a KEEP semantic manifest cannot fail closed."""


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _canonical_hash(value: Any) -> str:
    return _sha256_bytes(
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    )


def _default_allowlist_path() -> Path:
    return Path(__file__).resolve().parents[2] / "config" / "keep_normalization_allowlist.yaml"


@lru_cache(maxsize=4)
def load_keep_normalization_allowlist(path: str | Path | None = None) -> dict[str, Any]:
    source = Path(path).resolve() if path else _default_allowlist_path()
    if not source.is_file():
        raise KeepPreservationError(f"KEEP normalization allowlist is missing: {source}")
    payload = load_yaml_compatible(source)
    rules = payload.get("rules")
    if not isinstance(rules, list):
        raise KeepPreservationError("KEEP normalization allowlist rules must be a list")
    ids: set[str] = set()
    for index, raw in enumerate(rules, 1):
        if not isinstance(raw, Mapping):
            raise KeepPreservationError(f"KEEP normalization rule {index} is not an object")
        rule_id = str(raw.get("id", "")).strip()
        field = str(raw.get("field", "")).strip()
        if not rule_id or rule_id in ids:
            raise KeepPreservationError("KEEP normalization rule ids must be unique")
        if "*" in field or "all extlst" in field.lower() or "unknown" in field.lower():
            raise KeepPreservationError(f"Broad KEEP normalization rule is prohibited: {rule_id}")
        required = ("reason", "before_after_evidence", "powerpoint_behavior", "semantic_impact")
        if any(not str(raw.get(key, "")).strip() for key in required):
            raise KeepPreservationError(f"KEEP normalization rule is incomplete: {rule_id}")
        if str(raw.get("semantic_impact", "")).strip().upper() != "NONE":
            raise KeepPreservationError(f"KEEP normalization may ignore only NONE impact: {rule_id}")
        ids.add(rule_id)
    if ids != _ALLOWLIST_IDS:
        raise KeepPreservationError(
            "KEEP normalization allowlist/code mismatch: "
            f"missing={sorted(_ALLOWLIST_IDS - ids)}, unexpected={sorted(ids - _ALLOWLIST_IDS)}"
        )
    return {"path": str(source), "rules": [dict(row) for row in rules], "rule_ids": sorted(ids)}


def _local_name(node: etree._Element) -> str:
    return etree.QName(node).localname


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _rels_member(source_part: str) -> str:
    folder, name = posixpath.split(source_part)
    return posixpath.join(folder, "_rels", name + ".rels")


def _relationship_map(
    package: zipfile.ZipFile, source_part: str
) -> dict[str, dict[str, str]]:
    member = _rels_member(source_part)
    if member not in package.namelist():
        return {}
    root = etree.fromstring(package.read(member))
    result: dict[str, dict[str, str]] = {}
    for row in root:
        rel_id = str(row.get("Id", ""))
        rel_type = str(row.get("Type", "")).rsplit("/", 1)[-1]
        target = str(row.get("Target", "")).replace("\\", "/")
        mode = str(row.get("TargetMode", "Internal"))
        resolved = target if mode == "External" else _resolve_target(source_part, target)
        result[rel_id] = {
            "id": rel_id,
            "type": rel_type,
            "target_mode": mode,
            "target": target,
            "resolved_target": resolved,
        }
    return result


def _relationship_token(
    package: zipfile.ZipFile, relationship: Mapping[str, str]
) -> str:
    mode = relationship.get("target_mode", "Internal")
    rel_type = relationship.get("type", "UNKNOWN")
    target = relationship.get("resolved_target", "")
    if mode == "External":
        digest = _sha256_bytes(target.encode("utf-8"))
        return f"{rel_type}|EXTERNAL|{digest}"
    digest = ""
    if rel_type in _MEDIA_REL_TYPES and target in package.namelist():
        digest = _sha256_bytes(package.read(target))
        return f"{rel_type}|MEDIA_SHA256|{digest}"
    return f"{rel_type}|{target}|{digest}"


def _strip_whitespace(node: etree._Element) -> None:
    for row in node.iter():
        row.text = None if not (row.text or "").strip() else (row.text or "").strip()
        row.tail = None if not (row.tail or "").strip() else (row.tail or "").strip()


def _remove_exact_extension(node: etree._Element, *, parent_name: str, uri: str) -> None:
    for ext in list(node.xpath(f".//a:{parent_name}/a:extLst/a:ext[@uri=$uri]", namespaces=NS, uri=uri)):
        ext_list = ext.getparent()
        if ext_list is None:
            continue
        ext_list.remove(ext)
        if len(ext_list) == 0 and not ext_list.attrib and not (ext_list.text or "").strip():
            parent = ext_list.getparent()
            if parent is not None:
                parent.remove(ext_list)


def _shape_reference_map(shape_tree: etree._Element | None) -> dict[str, str]:
    result: dict[str, str] = {}
    if shape_tree is None:
        return result
    ordinal = 0
    for child in shape_tree:
        if _local_name(child) not in _OBJECT_NAMES:
            continue
        ordinal += 1
        identity = child.find(".//p:cNvPr", namespaces=NS)
        if identity is None:
            continue
        raw_id = str(identity.get("id", ""))
        name = str(identity.get("name", ""))
        if raw_id and raw_id not in result:
            result[raw_id] = f"shape:{ordinal}:{name}"
    return result


def _registered_dynamic_placeholder(element: etree._Element) -> str:
    placeholder = element.find(".//p:ph", namespaces=NS)
    if placeholder is None:
        return ""
    placeholder_type = str(placeholder.get("type", ""))
    return "SLIDE_NUMBER" if placeholder_type == "sldNum" else ""


def _normalized_xml(
    package: zipfile.ZipFile,
    element: etree._Element | None,
    *,
    relationship_by_id: Mapping[str, Mapping[str, str]] | None = None,
    shape_reference_by_id: Mapping[str, str] | None = None,
) -> bytes:
    load_keep_normalization_allowlist()
    if element is None:
        return b""
    clone = etree.fromstring(etree.tostring(element))
    relationship_by_id = relationship_by_id or {}
    shape_reference_by_id = shape_reference_by_id or {}

    dynamic_role = _registered_dynamic_placeholder(clone)
    if dynamic_role == "SLIDE_NUMBER":
        for text_node in clone.xpath(".//a:t", namespaces=NS):
            text_node.text = "__DYNAMIC_SLIDE_NUMBER__"

    _remove_exact_extension(clone, parent_name="gridCol", uri=_COL_ID_URI)
    _remove_exact_extension(clone, parent_name="tr", uri=_ROW_ID_URI)
    for effect in list(clone.xpath(".//a:effectLst[not(*) and not(normalize-space())]", namespaces=NS)):
        parent = effect.getparent()
        if parent is not None:
            parent.remove(effect)
    cell_property_nodes = clone.xpath("self::a:tcPr | .//a:tcPr", namespaces=NS)
    for cell_properties in cell_property_nodes:
        if cell_properties.get("marT") is None:
            cell_properties.set("marT", "45720")
        if cell_properties.get("marB") is None:
            cell_properties.set("marB", "45720")
    for identity in clone.xpath(".//p:cNvPr", namespaces=NS):
        identity.attrib.pop("id", None)
    for row in clone.iter():
        for key in list(row.attrib):
            local = etree.QName(key).localname
            value = str(row.attrib[key])
            namespace = etree.QName(key).namespace
            if namespace == NS["r"] and value in relationship_by_id:
                row.attrib[key] = _relationship_token(package, relationship_by_id[value])
            elif local in {"spid", "id"} and _local_name(row) in {"spTgt", "stCxn", "endCxn"}:
                row.attrib[key] = shape_reference_by_id.get(value, f"UNRESOLVED_SHAPE:{value}")
    _strip_whitespace(clone)
    return etree.tostring(clone, method="c14n", with_comments=False)


def _hash_normalized(
    package: zipfile.ZipFile,
    element: etree._Element | None,
    *,
    relationship_by_id: Mapping[str, Mapping[str, str]] | None = None,
    shape_reference_by_id: Mapping[str, str] | None = None,
) -> str:
    return _sha256_bytes(
        _normalized_xml(
            package,
            element,
            relationship_by_id=relationship_by_id,
            shape_reference_by_id=shape_reference_by_id,
        )
    )


def _xfrm_manifest(shape: etree._Element) -> dict[str, Any]:
    transform = shape.find("./p:xfrm", namespaces=NS)
    if transform is None:
        transform = shape.find("./p:spPr/a:xfrm", namespaces=NS)
    if transform is None:
        transform = shape.find("./p:grpSpPr/a:xfrm", namespaces=NS)
    if transform is None:
        return {"x": None, "y": None, "cx": None, "cy": None, "rotation": 0, "flip_h": "0", "flip_v": "0"}
    offset = transform.find("./a:off", namespaces=NS)
    extent = transform.find("./a:ext", namespaces=NS)
    return {
        "x": int(offset.get("x", "0")) if offset is not None else None,
        "y": int(offset.get("y", "0")) if offset is not None else None,
        "cx": int(extent.get("cx", "0")) if extent is not None else None,
        "cy": int(extent.get("cy", "0")) if extent is not None else None,
        "rotation": int(transform.get("rot", "0")),
        "flip_h": str(transform.get("flipH", "0")),
        "flip_v": str(transform.get("flipV", "0")),
    }


def _visible_text(element: etree._Element) -> dict[str, Any]:
    paragraphs: list[str] = []
    for paragraph in element.xpath(".//a:p", namespaces=NS):
        paragraphs.append("".join(paragraph.xpath(".//a:t/text()", namespaces=NS)))
    dynamic_role = _registered_dynamic_placeholder(element)
    visible = "__DYNAMIC_SLIDE_NUMBER__" if dynamic_role == "SLIDE_NUMBER" else "\n".join(paragraphs)
    return {"visible_text": visible, "paragraph_count": len(paragraphs)}


def _text_formatting_hash(package: zipfile.ZipFile, element: etree._Element) -> str:
    rows = []
    for row in element.xpath(
        ".//a:bodyPr | .//a:lstStyle | .//a:pPr | .//a:rPr | .//a:endParaRPr",
        namespaces=NS,
    ):
        rows.append(_normalized_xml(package, row).decode("utf-8"))
    return _canonical_hash(rows)


def _property_manifest(package: zipfile.ZipFile, shape: etree._Element) -> dict[str, str]:
    properties = shape.find("./p:spPr", namespaces=NS)
    if properties is None:
        properties = shape.find("./p:grpSpPr", namespaces=NS)
    fill = None
    if properties is not None:
        for child in properties:
            if _local_name(child) in {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}:
                fill = child
                break
    line = properties.find("./a:ln", namespaces=NS) if properties is not None else None
    return {
        "fill_sha256": _hash_normalized(package, fill),
        "line_sha256": _hash_normalized(package, line),
    }


def _hyperlink_action_manifest(
    package: zipfile.ZipFile,
    shape: etree._Element,
    relationship_by_id: Mapping[str, Mapping[str, str]],
) -> list[dict[str, str]]:
    rows = []
    for link in shape.xpath(".//a:hlinkClick | .//a:hlinkHover", namespaces=NS):
        rel_id = str(link.get(f"{{{NS['r']}}}id", ""))
        rows.append(
            {
                "kind": _local_name(link),
                "target": (
                    _relationship_token(package, relationship_by_id[rel_id])
                    if rel_id in relationship_by_id
                    else ""
                ),
                "action": str(link.get("action", "")),
            }
        )
    return rows


def _table_manifest(package: zipfile.ZipFile, shape: etree._Element) -> dict[str, Any] | None:
    table = shape.find(".//a:tbl", namespaces=NS)
    if table is None:
        return None
    columns = [int(row.get("w", "0")) for row in table.xpath("./a:tblGrid/a:gridCol", namespaces=NS)]
    row_manifests = []
    for row_index, row in enumerate(table.xpath("./a:tr", namespaces=NS), 1):
        cells = []
        for column_index, cell in enumerate(row.xpath("./a:tc", namespaces=NS), 1):
            cell_properties = cell.find("./a:tcPr", namespaces=NS)
            margins = {
                "left": int(cell_properties.get("marL", "91440")) if cell_properties is not None else 91440,
                "right": int(cell_properties.get("marR", "91440")) if cell_properties is not None else 91440,
                "top": int(cell_properties.get("marT", "45720")) if cell_properties is not None else 45720,
                "bottom": int(cell_properties.get("marB", "45720")) if cell_properties is not None else 45720,
            }
            borders = {}
            for side in ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"):
                border = cell_properties.find(f"./a:{side}", namespaces=NS) if cell_properties is not None else None
                borders[side] = _hash_normalized(package, border)
            fill = None
            if cell_properties is not None:
                for child in cell_properties:
                    if _local_name(child) in {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}:
                        fill = child
                        break
            text = _visible_text(cell)
            cells.append(
                {
                    "row": row_index,
                    "column": column_index,
                    **text,
                    "merge": {
                        "grid_span": str(cell.get("gridSpan", "1")),
                        "row_span": str(cell.get("rowSpan", "1")),
                        "h_merge": str(cell.get("hMerge", "0")),
                        "v_merge": str(cell.get("vMerge", "0")),
                    },
                    "margins_emu": margins,
                    "fill_sha256": _hash_normalized(package, fill),
                    "borders": borders,
                    "text_formatting_sha256": _text_formatting_hash(package, cell),
                    "cell_properties_sha256": _hash_normalized(package, cell_properties),
                }
            )
        row_manifests.append({"row": row_index, "height_emu": int(row.get("h", "0")), "cells": cells})
    return {
        "rows": len(row_manifests),
        "columns": len(columns),
        "column_widths_emu": columns,
        "row_manifests": row_manifests,
        "table_properties_sha256": _hash_normalized(package, table.find("./a:tblPr", namespaces=NS)),
    }


def _workbook_semantic_hash(payload: bytes) -> str:
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as workbook:
            rows = []
            for member in sorted(workbook.namelist()):
                if member in {"xl/workbook.xml", "xl/sharedStrings.xml", "xl/styles.xml"} or member.startswith("xl/worksheets/"):
                    if member.endswith(".xml"):
                        root = etree.fromstring(workbook.read(member))
                        _strip_whitespace(root)
                        rows.append((member, etree.tostring(root, method="c14n", with_comments=False).decode("utf-8")))
            return _canonical_hash(rows)
    except (zipfile.BadZipFile, etree.XMLSyntaxError):
        return _sha256_bytes(payload)


def _chart_manifest(
    package: zipfile.ZipFile,
    shape: etree._Element,
    relationship_by_id: Mapping[str, Mapping[str, str]],
) -> dict[str, Any] | None:
    chart_reference = shape.find(".//c:chart", namespaces=NS)
    if chart_reference is None:
        return None
    rel_id = str(chart_reference.get(f"{{{NS['r']}}}id", ""))
    relation = relationship_by_id.get(rel_id)
    if not relation:
        return {"status": "UNRESOLVED_CHART_RELATIONSHIP", "relationship_id": rel_id}
    chart_part = relation["resolved_target"]
    if chart_part not in package.namelist():
        return {"status": "MISSING_CHART_PART", "chart_part": chart_part}
    chart_root = etree.fromstring(package.read(chart_part))
    chart_rels = _relationship_map(package, chart_part)
    chart_types = [
        _local_name(row)
        for row in chart_root.xpath(".//c:plotArea/*", namespaces=NS)
        if _local_name(row).endswith("Chart")
    ]
    series = []
    for item in chart_root.xpath(".//c:ser", namespaces=NS):
        series.append(
            {
                "formulas": [str(value) for value in item.xpath(".//c:f/text()", namespaces=NS)],
                "values": [str(value) for value in item.xpath(".//c:v/text()", namespaces=NS)],
                "semantic_sha256": _hash_normalized(package, item, relationship_by_id=chart_rels),
            }
        )
    workbook = next((row for row in chart_rels.values() if row["type"] == "package"), None)
    workbook_hash = ""
    if workbook and workbook["resolved_target"] in package.namelist():
        workbook_hash = _workbook_semantic_hash(package.read(workbook["resolved_target"]))
    return {
        "status": "OK",
        "chart_part": chart_part,
        "chart_types": chart_types,
        "series": series,
        "workbook_semantic_sha256": workbook_hash,
        "chart_semantic_sha256": _hash_normalized(package, chart_root, relationship_by_id=chart_rels),
    }


def _picture_manifest(
    package: zipfile.ZipFile,
    shape: etree._Element,
    relationship_by_id: Mapping[str, Mapping[str, str]],
) -> dict[str, Any] | None:
    if _local_name(shape) != "pic":
        return None
    blip = shape.find(".//a:blip", namespaces=NS)
    rel_id = ""
    if blip is not None:
        rel_id = str(blip.get(f"{{{NS['r']}}}embed", "") or blip.get(f"{{{NS['r']}}}link", ""))
    relation = relationship_by_id.get(rel_id)
    target = relation["resolved_target"] if relation else ""
    source_rect = shape.find(".//a:srcRect", namespaces=NS)
    return {
        "media_target_semantics": (
            _relationship_token(package, relation) if relation else "MISSING"
        ),
        "media_sha256": _sha256_bytes(package.read(target)) if target in package.namelist() else "",
        "crop": dict(sorted(source_rect.attrib.items())) if source_rect is not None else {},
    }


def _shape_manifest(
    package: zipfile.ZipFile,
    shape: etree._Element,
    *,
    logical_index: str,
    z_order: int,
    relationship_by_id: Mapping[str, Mapping[str, str]],
    shape_reference_by_id: Mapping[str, str],
) -> dict[str, Any]:
    identity = shape.find(".//p:cNvPr", namespaces=NS)
    name = str(identity.get("name", "")) if identity is not None else ""
    result: dict[str, Any] = {
        "stable_logical_index": logical_index,
        "shape_type": _local_name(shape),
        "name": name,
        "z_order": z_order,
        "bbox_emu": _xfrm_manifest(shape),
        **_visible_text(shape),
        "text_formatting_sha256": _text_formatting_hash(package, shape),
        **_property_manifest(package, shape),
        "hyperlinks_actions": _hyperlink_action_manifest(package, shape, relationship_by_id),
        "registered_dynamic_placeholder": _registered_dynamic_placeholder(shape),
        "normalized_shape_sha256": _hash_normalized(
            package,
            shape,
            relationship_by_id=relationship_by_id,
            shape_reference_by_id=shape_reference_by_id,
        ),
    }
    picture = _picture_manifest(package, shape, relationship_by_id)
    table = _table_manifest(package, shape)
    chart = _chart_manifest(package, shape, relationship_by_id)
    if picture is not None:
        result["picture"] = picture
    if table is not None:
        result["table"] = table
    if chart is not None:
        result["chart"] = chart
    if _local_name(shape) == "grpSp":
        children = []
        child_z = 0
        for child in shape:
            if _local_name(child) not in _OBJECT_NAMES:
                continue
            child_z += 1
            children.append(
                _shape_manifest(
                    package,
                    child,
                    logical_index=f"{logical_index}.{child_z}",
                    z_order=child_z,
                    relationship_by_id=relationship_by_id,
                    shape_reference_by_id=shape_reference_by_id,
                )
            )
        result["children"] = children
    return result


def _presentation_rows(package: zipfile.ZipFile) -> list[dict[str, Any]]:
    root = etree.fromstring(package.read("ppt/presentation.xml"))
    relationships = _relationship_map(package, "ppt/presentation.xml")
    rows = []
    for order, slide in enumerate(root.xpath(".//p:sldId", namespaces=NS), 1):
        rel_id = str(slide.get(f"{{{NS['r']}}}id", ""))
        relation = relationships.get(rel_id)
        if not relation:
            raise KeepPreservationError(f"Unresolved presentation slide relationship: {rel_id}")
        rows.append(
            {
                "slide_order": order,
                "powerpoint_slide_id": str(slide.get("id", "")),
                "slide_part": relation["resolved_target"],
            }
        )
    return rows


def _manifest_from_package(
    package: zipfile.ZipFile,
    row: Mapping[str, Any],
    *,
    protected_slide_designation: bool,
) -> dict[str, Any]:
    slide_part = str(row["slide_part"])
    slide_root = etree.fromstring(package.read(slide_part))
    shape_tree = slide_root.find(".//p:spTree", namespaces=NS)
    relationships = _relationship_map(package, slide_part)
    layout = next((value for value in relationships.values() if value["type"] == "slideLayout"), None)
    notes = next((value for value in relationships.values() if value["type"] == "notesSlide"), None)
    layout_target = layout["resolved_target"] if layout else ""
    layout_relationships = _relationship_map(package, layout_target) if layout_target else {}
    master = next((value for value in layout_relationships.values() if value["type"] == "slideMaster"), None)
    master_target = master["resolved_target"] if master else ""
    media = []
    media_part_names = []
    for relation in relationships.values():
        if relation["type"] not in _MEDIA_REL_TYPES:
            continue
        target = relation["resolved_target"]
        media.append(
            {
                "type": relation["type"],
                "target_mode": relation["target_mode"],
                "semantic_target": _relationship_token(package, relation),
            }
        )
        media_part_names.append(target)
    media.sort(key=lambda value: (value["type"], value["target_mode"], value["semantic_target"]))
    media_part_names.sort()
    presentation_identity = {
        "powerpoint_slide_id": str(row["powerpoint_slide_id"]),
        "slide_order": int(row["slide_order"]),
        "slide_part": slide_part,
        "layout_relationship_target": layout_target,
        "master_relationship_target": master_target,
        "media_relationship_targets": media,
        "media_part_names_diagnostic": media_part_names,
        "notes_role": "NOTES_SLIDE" if notes else "NONE",
        "notes_relationship_target": notes["resolved_target"] if notes else "",
        "protected_slide_designation": bool(protected_slide_designation),
    }
    shape_reference_by_id = _shape_reference_map(shape_tree)
    shapes = []
    z_order = 0
    if shape_tree is not None:
        for child in shape_tree:
            if _local_name(child) not in _OBJECT_NAMES:
                continue
            z_order += 1
            shapes.append(
                _shape_manifest(
                    package,
                    child,
                    logical_index=str(z_order),
                    z_order=z_order,
                    relationship_by_id=relationships,
                    shape_reference_by_id=shape_reference_by_id,
                )
            )
    background = slide_root.find("./p:cSld/p:bg", namespaces=NS)
    timing = slide_root.find("./p:timing", namespaces=NS)
    transition = slide_root.find("./p:transition", namespaces=NS)
    editable_identity = {
        "background_sha256": _hash_normalized(package, background),
        "shapes": shapes,
        "timing_sha256": _hash_normalized(
            package,
            timing,
            relationship_by_id=relationships,
            shape_reference_by_id=shape_reference_by_id,
        ),
        "transition_sha256": _hash_normalized(package, transition, relationship_by_id=relationships),
    }
    semantic = {
        "schema_version": "academic-ppt-slide-semantic-manifest/1",
        "presentation_identity": presentation_identity,
        **editable_identity,
    }
    # Slide order, part names and relationship targets belong to the separate
    # PRESENTATION_IDENTITY gate.  Insertions may legitimately shift a KEEP
    # slide's order without changing its editable contents.
    semantic["semantic_fingerprint_sha256"] = _canonical_hash(editable_identity)
    return semantic


def build_deck_semantic_manifests(
    path: Path,
    *,
    protected_by_order: Mapping[int, bool] | None = None,
) -> list[dict[str, Any]]:
    """Build canonical KEEP manifests for every slide in presentation order."""

    load_keep_normalization_allowlist()
    protected_by_order = protected_by_order or {}
    try:
        with zipfile.ZipFile(path, "r") as package:
            return [
                _manifest_from_package(
                    package,
                    row,
                    protected_slide_designation=bool(protected_by_order.get(int(row["slide_order"]), False)),
                )
                for row in _presentation_rows(package)
            ]
    except (OSError, KeyError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        raise KeepPreservationError(f"Cannot build KEEP semantic manifest: {exc}") from exc


def build_slide_semantic_manifest(
    path: Path,
    slide_index: int,
    *,
    protected_slide_designation: bool = False,
) -> dict[str, Any]:
    """Return the one canonical semantic manifest for a slide.

    This is a normalized inspection of the existing OOXML package; it is not a
    planning/rendering IR and cannot generate or mutate presentation content.
    """

    if slide_index < 1:
        raise KeepPreservationError("slide_index must be positive")
    manifests = build_deck_semantic_manifests(
        path,
        protected_by_order={slide_index: protected_slide_designation},
    )
    if slide_index > len(manifests):
        raise KeepPreservationError("slide_index exceeds deck length")
    return manifests[slide_index - 1]


def semantic_fingerprint(manifest: Mapping[str, Any]) -> str:
    value = {
        key: data
        for key, data in manifest.items()
        if key not in {"schema_version", "presentation_identity", "semantic_fingerprint_sha256"}
    }
    return _canonical_hash(value)


__all__ = [
    "KeepPreservationError",
    "build_deck_semantic_manifests",
    "build_slide_semantic_manifest",
    "load_keep_normalization_allowlist",
    "semantic_fingerprint",
]
