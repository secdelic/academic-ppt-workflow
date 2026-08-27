"""Structural and safety QA for PowerPoint OPC/OOXML packages.

The inspector is intentionally read-only.  It does not execute macros, ActiveX,
OLE packages, linked data, or external relationships.  Text is used internally
for hashes, near-empty detection, and the required ``[Sources]`` notes marker,
but raw slide or notes text is not emitted in the report.
"""

from __future__ import annotations

import hashlib
import json
import posixpath
import re
import urllib.parse
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable, Mapping
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

REL_NS = NS["pr"]
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    "officeDocument"
)
REL_SLIDE = REL_BASE + "slide"
REL_LAYOUT = REL_BASE + "slideLayout"
REL_MASTER = REL_BASE + "slideMaster"
REL_THEME = REL_BASE + "theme"
REL_NOTES = REL_BASE + "notesSlide"
REL_IMAGE = REL_BASE + "image"
REL_CHART = REL_BASE + "chart"

REQUIRED_PARTS = (
    "[Content_Types].xml",
    "_rels/.rels",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
)

PRESENTATION_CONTENT_TYPES = {
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
}
SLIDE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
)

MAX_XML_MEMBER_BYTES = 64 * 1024 * 1024
MAX_PACKAGE_ENTRIES = 50_000
MAX_TOTAL_UNCOMPRESSED_BYTES = 2 * 1024 * 1024 * 1024


class OOXMLQAError(ValueError):
    """Raised when an OPC package is unreadable or structurally unsafe to parse."""


def _issue(
    code: str,
    message: str,
    *,
    part: str | None = None,
    relationship_id: str | None = None,
    slide_part: str | None = None,
) -> dict[str, Any]:
    result: dict[str, Any] = {"code": code, "message": message}
    if part is not None:
        result["part"] = part
    if relationship_id is not None:
        result["relationship_id"] = relationship_id
    if slide_part is not None:
        result["slide_part"] = slide_part
    return result


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _sha256_member(package: zipfile.ZipFile, member: str) -> str:
    digest = hashlib.sha256()
    with package.open(member, "r") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_names(package: zipfile.ZipFile) -> set[str]:
    infos = package.infolist()
    if len(infos) > MAX_PACKAGE_ENTRIES:
        raise OOXMLQAError(
            f"Package has too many entries ({len(infos)}); QA refused"
        )
    total = sum(max(0, item.file_size) for item in infos)
    if total > MAX_TOTAL_UNCOMPRESSED_BYTES:
        raise OOXMLQAError(f"Package expands to {total} bytes; QA refused")
    names: set[str] = set()
    for info in infos:
        name = info.filename.replace("\\", "/")
        normalized = posixpath.normpath(name).lstrip("/")
        if (
            not name
            or name.startswith("/")
            or normalized in {"", ".", ".."}
            or normalized.startswith("../")
        ):
            raise OOXMLQAError(f"Unsafe OPC member path: {name!r}")
        if normalized in names:
            raise OOXMLQAError(f"Duplicate normalized OPC member path: {normalized}")
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
            raise OOXMLQAError(f"Required OOXML part is missing: {member}")
        return None
    info = package.getinfo(member)
    if info.file_size > MAX_XML_MEMBER_BYTES:
        raise OOXMLQAError(f"XML part is too large to inspect safely: {member}")
    try:
        return ET.fromstring(package.read(member))
    except (ET.ParseError, KeyError, RuntimeError, zipfile.BadZipFile) as exc:
        raise OOXMLQAError(f"Invalid XML in {member}: {exc}") from exc


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


def _resolve_target(source_part: str, target: str) -> str | None:
    parsed = urllib.parse.urlsplit(target)
    if parsed.scheme or parsed.netloc:
        return None
    path = urllib.parse.unquote(parsed.path).replace("\\", "/")
    if path.startswith("/"):
        resolved = posixpath.normpath(path.lstrip("/"))
    else:
        resolved = posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), path)
        )
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        return None
    return resolved.lstrip("/")


def _parse_relationships(
    package: zipfile.ZipFile,
    names: set[str],
    source_part: str,
) -> list[dict[str, str]]:
    rels_part = _relationships_part(source_part)
    root = _read_xml(package, names, rels_part)
    if root is None:
        return []
    rows: list[dict[str, str]] = []
    for element in root.findall(f"{{{REL_NS}}}Relationship"):
        relationship_id = element.get("Id", "")
        target = element.get("Target", "")
        target_mode = element.get("TargetMode", "")
        external = target_mode.lower() == "external"
        rows.append(
            {
                "source_part": source_part,
                "relationships_part": rels_part,
                "id": relationship_id,
                "type": element.get("Type", ""),
                "target_mode": "External" if external else "Internal",
                "target": target,
                "resolved_target": (
                    "" if external else (_resolve_target(source_part, target) or "")
                ),
            }
        )
    return rows


def _redacted_external(row: Mapping[str, str]) -> dict[str, str]:
    target = row.get("target", "")
    parsed = urllib.parse.urlsplit(target)
    return {
        "source_part": row.get("source_part", "") or "[package]",
        "relationship_id": row.get("id", ""),
        "relationship_type": row.get("type", ""),
        "target_scheme": parsed.scheme.lower() or "unspecified",
        "target_sha256": hashlib.sha256(target.encode("utf-8")).hexdigest(),
        "followed": "NO",
    }


def _parse_content_types(
    package: zipfile.ZipFile, names: set[str]
) -> dict[str, Any]:
    root = _read_xml(package, names, "[Content_Types].xml", required=True)
    assert root is not None
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    duplicate_defaults: list[str] = []
    duplicate_overrides: list[str] = []
    for element in root:
        local = element.tag.rsplit("}", 1)[-1]
        if local == "Default":
            extension = element.get("Extension", "").lower()
            content_type = element.get("ContentType", "")
            if extension in defaults:
                duplicate_defaults.append(extension)
            defaults[extension] = content_type
        elif local == "Override":
            part = element.get("PartName", "").lstrip("/")
            content_type = element.get("ContentType", "")
            if part in overrides:
                duplicate_overrides.append(part)
            overrides[part] = content_type
    return {
        "defaults": defaults,
        "overrides": overrides,
        "duplicate_default_extensions": sorted(set(duplicate_defaults)),
        "duplicate_override_parts": sorted(set(duplicate_overrides)),
    }


def _content_type_for_part(
    part: str, content_types: Mapping[str, Any]
) -> str | None:
    overrides = content_types.get("overrides", {})
    if part in overrides:
        return overrides[part]
    extension = part.rsplit(".", 1)[-1].lower() if "." in part else ""
    return content_types.get("defaults", {}).get(extension)


def _relationship_category(relationship_type: str, target: str) -> str:
    suffix = relationship_type.rsplit("/", 1)[-1].lower()
    normalized_target = target.lower()
    if suffix == "image" or normalized_target.startswith("ppt/media/"):
        return "media"
    if suffix in {"chart", "chartuserShapes".lower()} or normalized_target.startswith(
        "ppt/charts/"
    ):
        return "chart"
    if (
        suffix in {"oleobject", "package"}
        or normalized_target.startswith("ppt/embeddings/")
    ):
        return "embedded_or_ole"
    if suffix in {
        "externallink",
        "externaldata",
        "querytable",
        "connection",
        "datalink",
    }:
        return "external_data"
    if suffix in {"audio", "video", "media"}:
        return "media"
    return "other"


def _canonical_xml(data: bytes) -> bytes:
    try:
        canonical = ET.canonicalize(
            xml_data=data.decode("utf-8-sig"),
            with_comments=False,
            strip_text=False,
            rewrite_prefixes=True,
        )
        return canonical.encode("utf-8")
    except (AttributeError, ET.ParseError, UnicodeDecodeError, ValueError):
        root = ET.fromstring(data)

        def normalized(element: ET.Element) -> Any:
            return [
                element.tag,
                sorted(element.attrib.items()),
                element.text or "",
                [normalized(child) for child in list(element)],
                element.tail or "",
            ]

        return json.dumps(
            normalized(root), ensure_ascii=False, separators=(",", ":")
        ).encode("utf-8")


def _element_hash(element: ET.Element) -> str:
    return hashlib.sha256(_canonical_xml(ET.tostring(element))).hexdigest()


def _geometry(element: ET.Element) -> dict[str, int | None]:
    transform = element.find(".//a:xfrm", NS)
    if transform is None:
        return {"x_emu": None, "y_emu": None, "width_emu": None, "height_emu": None}
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)

    def number(item: ET.Element | None, key: str) -> int | None:
        if item is None or item.get(key) is None:
            return None
        try:
            return int(item.get(key, ""))
        except ValueError:
            return None

    return {
        "x_emu": number(offset, "x"),
        "y_emu": number(offset, "y"),
        "width_emu": number(extent, "cx"),
        "height_emu": number(extent, "cy"),
    }


def _object_kind(element: ET.Element) -> str:
    local = element.tag.rsplit("}", 1)[-1]
    return {
        "sp": "shape",
        "pic": "picture",
        "graphicFrame": "graphic_frame",
        "grpSp": "group",
        "cxnSp": "connector",
        "contentPart": "content_part",
    }.get(local, local)


def _relationship_ids_in_element(element: ET.Element) -> list[str]:
    ids: set[str] = set()
    relationship_namespace = "{" + NS["r"] + "}"
    for descendant in element.iter():
        for key, value in descendant.attrib.items():
            if key.startswith(relationship_namespace) and value:
                ids.add(value)
    return sorted(ids)


def _slide_object_manifest(
    slide_root: ET.Element,
    relationship_by_id: Mapping[str, Mapping[str, str]],
    package_names: set[str],
) -> dict[str, Any]:
    objects: list[dict[str, Any]] = []
    object_tags = {
        f"{{{NS['p']}}}sp",
        f"{{{NS['p']}}}pic",
        f"{{{NS['p']}}}graphicFrame",
        f"{{{NS['p']}}}grpSp",
        f"{{{NS['p']}}}cxnSp",
        f"{{{NS['p']}}}contentPart",
    }
    # Only objects directly in a shape tree are counted as top-level objects.
    shape_tree = slide_root.find("p:cSld/p:spTree", NS)
    candidates = list(shape_tree) if shape_tree is not None else []
    for ordinal, element in enumerate(
        (item for item in candidates if item.tag in object_tags), start=1
    ):
        placeholder = element.find(".//p:nvPr/p:ph", NS)
        native_id_element = element.find(".//p:cNvPr", NS)
        related_assets: list[dict[str, Any]] = []
        for relationship_id in _relationship_ids_in_element(element):
            relationship = relationship_by_id.get(relationship_id)
            if relationship is None:
                related_assets.append(
                    {
                        "relationship_id": relationship_id,
                        "relationship_type": "UNKNOWN",
                        "target_part": "UNRESOLVED",
                        "exists": False,
                    }
                )
            elif relationship["target_mode"] == "External":
                related_assets.append(
                    {
                        "relationship_id": relationship_id,
                        "relationship_type": relationship["type"],
                        "target_part": "EXTERNAL_REDACTED",
                        "target_sha256": hashlib.sha256(
                            relationship["target"].encode("utf-8")
                        ).hexdigest(),
                        "exists": "NOT_APPLICABLE",
                    }
                )
            else:
                related_assets.append(
                    {
                        "relationship_id": relationship_id,
                        "relationship_type": relationship["type"],
                        "target_part": relationship["resolved_target"]
                        or "UNRESOLVED",
                        "exists": relationship["resolved_target"] in package_names,
                    }
                )
        objects.append(
            {
                "object_ordinal": ordinal,
                "native_object_id": (
                    native_id_element.get("id", "UNKNOWN")
                    if native_id_element is not None
                    else "UNKNOWN"
                ),
                "kind": _object_kind(element),
                "placeholder_type": (
                    placeholder.get("type", "body")
                    if placeholder is not None
                    else "NOT_A_PLACEHOLDER"
                ),
                "geometry": _geometry(element),
                "object_sha256": _element_hash(element),
                "related_assets": related_assets,
            }
        )

    text_values = [element.text or "" for element in slide_root.findall(".//a:t", NS)]
    text_character_count = len("".join(text_values).strip())
    counts = Counter(item["kind"] for item in objects)
    table_count = len(slide_root.findall(".//a:tbl", NS))
    chart_count = len(slide_root.findall(".//c:chart", NS))
    ole_count = len(slide_root.findall(".//p:oleObj", NS))
    media_relationship_count = sum(
        1
        for relationship in relationship_by_id.values()
        if _relationship_category(
            relationship["type"], relationship.get("resolved_target", "")
        )
        in {"media", "embedded_or_ole"}
    )
    result = {
        "object_count": len(objects),
        "object_kind_counts": dict(sorted(counts.items())),
        "table_count": table_count,
        "chart_count": chart_count,
        "ole_object_count": ole_count,
        "media_relationship_count": media_relationship_count,
        "text_character_count": text_character_count,
        "text_run_count": len(text_values),
        "text_sha256": hashlib.sha256(
            "\n".join(text_values).encode("utf-8")
        ).hexdigest(),
        "objects": objects,
    }
    result["manifest_sha256"] = hashlib.sha256(
        json.dumps(result, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()
    return result


def _canonical_slide_hash(
    package: zipfile.ZipFile,
    names: set[str],
    slide_part: str,
    relationships: Iterable[Mapping[str, str]],
) -> str:
    digest = hashlib.sha256()
    digest.update(_canonical_xml(package.read(slide_part)))
    related_fingerprints: list[tuple[str, str, str, str]] = []
    for relationship in relationships:
        relationship_type = relationship.get("type", "")
        # Notes content is deliberately excluded from the rendered slide hash.
        if relationship_type == REL_NOTES:
            continue
        if relationship.get("target_mode") == "External":
            target_fingerprint = hashlib.sha256(
                relationship.get("target", "").encode("utf-8")
            ).hexdigest()
            related_fingerprints.append(
                (
                    relationship.get("id", ""),
                    relationship_type,
                    "EXTERNAL",
                    target_fingerprint,
                )
            )
        else:
            target = relationship.get("resolved_target", "")
            if target and target in names:
                target_fingerprint = _sha256_member(package, target)
            else:
                target_fingerprint = "MISSING"
            related_fingerprints.append(
                (
                    relationship.get("id", ""),
                    relationship_type,
                    "INTERNAL",
                    target_fingerprint,
                )
            )
    digest.update(
        json.dumps(
            sorted(related_fingerprints), separators=(",", ":"), ensure_ascii=True
        ).encode("ascii")
    )
    return digest.hexdigest()


def _source_marker_check(
    package: zipfile.ZipFile,
    names: set[str],
    slide_part: str,
    relationships: Iterable[Mapping[str, str]],
    known_source_ids: set[str] | None = None,
) -> dict[str, Any]:
    notes_rows = [
        row
        for row in relationships
        if row.get("type") == REL_NOTES and row.get("target_mode") == "Internal"
    ]
    if not notes_rows:
        return {
            "slide_part": slide_part,
            "notes_part": "MISSING",
            "has_sources_marker": False,
            "source_id_token_count": 0,
            "recognized_source_id_count": 0,
            "unknown_source_id_count": 0,
        }
    notes_part = notes_rows[0].get("resolved_target", "")
    notes_root = _read_xml(package, names, notes_part)
    if notes_root is None:
        return {
            "slide_part": slide_part,
            "notes_part": notes_part or "UNRESOLVED",
            "has_sources_marker": False,
            "source_id_token_count": 0,
            "recognized_source_id_count": 0,
            "unknown_source_id_count": 0,
        }
    text_runs = [element.text or "" for element in notes_root.findall(".//a:t", NS)]
    joined = "\n".join(text_runs)
    compact = "".join(text_runs)
    marker = "[Sources]" in joined or "[Sources]" in compact
    source_ids = set(
        re.findall(
            r"\bSRC-[A-Za-z0-9](?:[A-Za-z0-9_.-]*[A-Za-z0-9])?",
            joined,
        )
    )
    if known_source_ids is None:
        recognized_count: int | str = "NOT_CHECKED"
        unknown_count: int | str = "NOT_CHECKED"
    else:
        recognized_count = len(source_ids & known_source_ids)
        unknown_count = len(source_ids - known_source_ids)
    return {
        "slide_part": slide_part,
        "notes_part": notes_part,
        "has_sources_marker": marker,
        "source_id_token_count": len(source_ids),
        "recognized_source_id_count": recognized_count,
        "unknown_source_id_count": unknown_count,
        "notes_text_returned": False,
    }


def _relationship_chain_target(
    rows: Iterable[Mapping[str, str]], relationship_type: str
) -> tuple[str | None, int]:
    matches = [
        row
        for row in rows
        if row.get("type") == relationship_type
        and row.get("target_mode") == "Internal"
        and row.get("resolved_target")
    ]
    return (matches[0].get("resolved_target") if matches else None, len(matches))


def inspect_pptx_ooxml(
    path: str | Path,
    *,
    known_source_ids: Iterable[str] | None = None,
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]] | None = None,
    previous_identity_map: Iterable[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Run structural OOXML QA and return a machine-readable audit report."""

    source = Path(path)
    if source.suffix.lower() not in {".pptx", ".potx", ".pptm", ".potm"}:
        raise OOXMLQAError("OOXML QA input must be PPTX, POTX, PPTM, or POTM")
    if not source.is_file():
        raise OOXMLQAError(f"OOXML package does not exist: {source}")
    before_hash = _sha256_file(source)
    try:
        package = zipfile.ZipFile(source, mode="r")
    except (OSError, zipfile.BadZipFile) as exc:
        raise OOXMLQAError(f"Input is not a readable OOXML ZIP: {exc}") from exc

    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    if known_source_ids is not None:
        normalized_known_source_ids = {
            str(source_id) for source_id in known_source_ids
        }
    elif deck_ir is not None:
        normalized_known_source_ids = {
            str(source_id)
            for slide in _ir_slide_list(deck_ir)
            for source_id in slide.get("source_ids", [])
        }
    else:
        normalized_known_source_ids = None
    with package:
        names = _safe_names(package)
        missing_required = [part for part in REQUIRED_PARTS if part not in names]
        if missing_required:
            raise OOXMLQAError(
                "Required OPC parts are missing: " + ", ".join(missing_required)
            )
        presentation = _read_xml(
            package, names, "ppt/presentation.xml", required=True
        )
        assert presentation is not None
        content_types = _parse_content_types(package, names)
        if content_types["duplicate_default_extensions"]:
            errors.append(
                _issue(
                    "DUPLICATE_CONTENT_TYPE_DEFAULT",
                    "Duplicate Default content-type extensions were declared.",
                    part="[Content_Types].xml",
                )
            )
        if content_types["duplicate_override_parts"]:
            errors.append(
                _issue(
                    "DUPLICATE_CONTENT_TYPE_OVERRIDE",
                    "Duplicate Override content-type parts were declared.",
                    part="[Content_Types].xml",
                )
            )
        presentation_type = _content_type_for_part(
            "ppt/presentation.xml", content_types
        )
        if presentation_type not in PRESENTATION_CONTENT_TYPES:
            errors.append(
                _issue(
                    "INVALID_PRESENTATION_CONTENT_TYPE",
                    "The presentation part does not have a recognized presentation content type.",
                    part="ppt/presentation.xml",
                )
            )

        relationship_rows: dict[str, list[dict[str, str]]] = {}
        all_relationships: list[dict[str, str]] = []
        for rels_part in sorted(name for name in names if name.endswith(".rels")):
            source_part = _source_part_for_relationships(rels_part)
            if source_part is None:
                warnings.append(
                    _issue(
                        "UNRECOGNIZED_RELATIONSHIPS_PART",
                        "A .rels member is not in a recognized OPC relationship location.",
                        part=rels_part,
                    )
                )
                continue
            if source_part and source_part not in names:
                errors.append(
                    _issue(
                        "RELATIONSHIP_SOURCE_MISSING",
                        "A relationships part has no corresponding source part.",
                        part=rels_part,
                    )
                )
            rows = _parse_relationships(package, names, source_part)
            relationship_rows[source_part] = rows
            all_relationships.extend(rows)
            ids = [row["id"] for row in rows]
            for duplicate_id, count in Counter(ids).items():
                if not duplicate_id or count > 1:
                    errors.append(
                        _issue(
                            "DUPLICATE_OR_EMPTY_RELATIONSHIP_ID",
                            "Relationship IDs must be non-empty and unique within a part.",
                            part=rels_part,
                            relationship_id=duplicate_id or "[empty]",
                        )
                    )
            for row in rows:
                if row["target_mode"] == "External":
                    continue
                if not row["resolved_target"]:
                    errors.append(
                        _issue(
                            "UNSAFE_RELATIONSHIP_TARGET",
                            "An internal relationship target resolves outside the OPC package.",
                            part=rels_part,
                            relationship_id=row["id"],
                        )
                    )
                elif row["resolved_target"] not in names:
                    errors.append(
                        _issue(
                            "DANGLING_RELATIONSHIP",
                            "An internal relationship target is missing.",
                            part=rels_part,
                            relationship_id=row["id"],
                        )
                    )

        root_office_relationships = [
            row
            for row in relationship_rows.get("", [])
            if row["type"] == REL_OFFICE_DOCUMENT
            and row["target_mode"] == "Internal"
            and row["resolved_target"] == "ppt/presentation.xml"
        ]
        if len(root_office_relationships) != 1:
            errors.append(
                _issue(
                    "INVALID_ROOT_OFFICE_DOCUMENT_RELATIONSHIP",
                    "The package must contain exactly one internal officeDocument relationship to ppt/presentation.xml.",
                    part="_rels/.rels",
                )
            )

        external_relationships = [
            _redacted_external(row)
            for row in all_relationships
            if row["target_mode"] == "External"
        ]
        for row in all_relationships:
            if row["target_mode"] != "External":
                continue
            category = _relationship_category(row["type"], "")
            issue = _issue(
                "EXTERNAL_RELATIONSHIP_NOT_FOLLOWED",
                "An external relationship was detected and was not followed.",
                part=row["relationships_part"],
                relationship_id=row["id"],
            )
            if category == "external_data":
                errors.append(issue)
            else:
                warnings.append(issue)

        presentation_relationships = {
            row["id"]: row
            for row in relationship_rows.get("ppt/presentation.xml", [])
        }
        slide_order: list[dict[str, Any]] = []
        seen_slide_ids: set[str] = set()
        seen_relationship_ids: set[str] = set()
        seen_targets: set[str] = set()
        slide_list = presentation.find("p:sldIdLst", NS)
        if slide_list is None:
            errors.append(
                _issue(
                    "MISSING_SLIDE_ID_LIST",
                    "ppt/presentation.xml has no p:sldIdLst.",
                    part="ppt/presentation.xml",
                )
            )
            slide_id_elements: list[ET.Element] = []
        else:
            slide_id_elements = slide_list.findall("p:sldId", NS)
        for index, element in enumerate(slide_id_elements, start=1):
            numeric_id = element.get("id", "")
            relationship_id = element.get(f"{{{NS['r']}}}id", "")
            relationship = presentation_relationships.get(relationship_id)
            slide_part = (
                relationship["resolved_target"]
                if relationship is not None
                and relationship["type"] == REL_SLIDE
                and relationship["target_mode"] == "Internal"
                else ""
            )
            if not numeric_id or numeric_id in seen_slide_ids:
                errors.append(
                    _issue(
                        "DUPLICATE_OR_EMPTY_PRESENTATION_SLIDE_ID",
                        "p:sldId/@id values must be non-empty and unique.",
                        part="ppt/presentation.xml",
                        relationship_id=relationship_id,
                    )
                )
            seen_slide_ids.add(numeric_id)
            if not relationship_id or relationship_id in seen_relationship_ids:
                errors.append(
                    _issue(
                        "DUPLICATE_OR_EMPTY_SLIDE_RELATIONSHIP_ID",
                        "p:sldId/@r:id values must be non-empty and unique.",
                        part="ppt/presentation.xml",
                        relationship_id=relationship_id or "[empty]",
                    )
                )
            seen_relationship_ids.add(relationship_id)
            if not slide_part:
                errors.append(
                    _issue(
                        "SLIDE_ORDER_RELATIONSHIP_INVALID",
                        "A p:sldId entry does not resolve to an internal slide part.",
                        part="ppt/presentation.xml",
                        relationship_id=relationship_id,
                    )
                )
            elif slide_part in seen_targets:
                errors.append(
                    _issue(
                        "DUPLICATE_SLIDE_TARGET",
                        "Multiple p:sldId entries resolve to the same slide part.",
                        part="ppt/presentation.xml",
                        relationship_id=relationship_id,
                    )
                )
            seen_targets.add(slide_part)
            slide_order.append(
                {
                    "slide_index": index,
                    "presentation_slide_id": numeric_id,
                    "relationship_id": relationship_id,
                    "slide_part": slide_part or "UNRESOLVED",
                }
            )

        archive_slide_parts = sorted(
            name
            for name in names
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
        )
        unreferenced_slides = sorted(set(archive_slide_parts) - seen_targets)
        if unreferenced_slides:
            warnings.append(
                _issue(
                    "UNREFERENCED_SLIDE_PART",
                    f"{len(unreferenced_slides)} slide part(s) are not in presentation order.",
                    part="ppt/presentation.xml",
                )
            )

        for order_row in slide_order:
            slide_part = order_row["slide_part"]
            if slide_part == "UNRESOLVED" or slide_part not in names:
                continue
            if _content_type_for_part(slide_part, content_types) != SLIDE_CONTENT_TYPE:
                errors.append(
                    _issue(
                        "INVALID_SLIDE_CONTENT_TYPE",
                        "A referenced slide part lacks the expected slide content type.",
                        part=slide_part,
                    )
                )

        slide_chains: list[dict[str, Any]] = []
        notes_source_checks: list[dict[str, Any]] = []
        object_manifest: list[dict[str, Any]] = []
        canonical_hashes: list[dict[str, Any]] = []
        near_empty_pages: list[dict[str, Any]] = []
        for order_row in slide_order:
            slide_part = order_row["slide_part"]
            if slide_part == "UNRESOLVED" or slide_part not in names:
                continue
            slide_root = _read_xml(package, names, slide_part, required=True)
            assert slide_root is not None
            slide_relationships = relationship_rows.get(slide_part, [])
            layout_part, layout_count = _relationship_chain_target(
                slide_relationships, REL_LAYOUT
            )
            if layout_count != 1 or not layout_part or layout_part not in names:
                errors.append(
                    _issue(
                        "INVALID_SLIDE_LAYOUT_CHAIN",
                        "Each slide must resolve to exactly one existing slide layout.",
                        slide_part=slide_part,
                    )
                )
            layout_relationships = (
                relationship_rows.get(layout_part, []) if layout_part else []
            )
            master_part, master_count = _relationship_chain_target(
                layout_relationships, REL_MASTER
            )
            if master_count != 1 or not master_part or master_part not in names:
                errors.append(
                    _issue(
                        "INVALID_LAYOUT_MASTER_CHAIN",
                        "Each used layout must resolve to exactly one existing slide master.",
                        slide_part=slide_part,
                    )
                )
            master_relationships = (
                relationship_rows.get(master_part, []) if master_part else []
            )
            theme_part, theme_count = _relationship_chain_target(
                master_relationships, REL_THEME
            )
            if theme_count != 1 or not theme_part or theme_part not in names:
                errors.append(
                    _issue(
                        "INVALID_MASTER_THEME_CHAIN",
                        "Each used master must resolve to exactly one existing theme.",
                        slide_part=slide_part,
                    )
                )
            slide_chains.append(
                {
                    "slide_index": order_row["slide_index"],
                    "slide_part": slide_part,
                    "layout_part": layout_part or "UNRESOLVED",
                    "master_part": master_part or "UNRESOLVED",
                    "theme_part": theme_part or "UNRESOLVED",
                    "chain_valid": bool(
                        layout_count == 1
                        and layout_part in names
                        and master_count == 1
                        and master_part in names
                        and theme_count == 1
                        and theme_part in names
                    ),
                }
            )

            source_check = _source_marker_check(
                package,
                names,
                slide_part,
                slide_relationships,
                normalized_known_source_ids,
            )
            notes_source_checks.append(source_check)
            if not source_check["has_sources_marker"]:
                errors.append(
                    _issue(
                        "MISSING_NOTES_SOURCES_MARKER",
                        "Slide notes are missing the required [Sources] marker.",
                        slide_part=slide_part,
                    )
                )
            if (
                normalized_known_source_ids is not None
                and source_check["unknown_source_id_count"]
            ):
                errors.append(
                    _issue(
                        "UNKNOWN_NOTES_SOURCE_ID",
                        "Slide notes contain source IDs that are not in the supplied source registry.",
                        slide_part=slide_part,
                    )
                )

            relationship_by_id = {
                row["id"]: row for row in slide_relationships if row["id"]
            }
            manifest = _slide_object_manifest(
                slide_root, relationship_by_id, names
            )
            manifest_row = {
                "slide_index": order_row["slide_index"],
                "slide_part": slide_part,
                **manifest,
            }
            object_manifest.append(manifest_row)
            canonical_hash = _canonical_slide_hash(
                package, names, slide_part, slide_relationships
            )
            canonical_hashes.append(
                {
                    "slide_index": order_row["slide_index"],
                    "slide_part": slide_part,
                    "canonical_sha256": canonical_hash,
                    "object_manifest_sha256": manifest["manifest_sha256"],
                }
            )
            semantic_objects = (
                manifest["object_kind_counts"].get("picture", 0)
                + manifest["chart_count"]
                + manifest["table_count"]
                + manifest["ole_object_count"]
                + manifest["media_relationship_count"]
            )
            if manifest["text_character_count"] < 4 and semantic_objects == 0:
                near_empty_pages.append(
                    {
                        "slide_index": order_row["slide_index"],
                        "slide_part": slide_part,
                        "text_character_count": manifest["text_character_count"],
                        "semantic_object_count": semantic_objects,
                    }
                )

        if near_empty_pages:
            warnings.append(
                _issue(
                    "NEAR_EMPTY_SLIDES",
                    f"{len(near_empty_pages)} slide(s) appear near-empty and require review.",
                )
            )

        duplicate_groups: list[dict[str, Any]] = []
        by_hash: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in canonical_hashes:
            by_hash[row["canonical_sha256"]].append(row)
        for digest, rows in sorted(by_hash.items()):
            if len(rows) > 1:
                duplicate_groups.append(
                    {
                        "canonical_sha256": digest,
                        "slide_indexes": [row["slide_index"] for row in rows],
                        "slide_parts": [row["slide_part"] for row in rows],
                    }
                )
        if duplicate_groups:
            warnings.append(
                _issue(
                    "DUPLICATE_SLIDES",
                    f"{len(duplicate_groups)} exact canonical duplicate group(s) were detected.",
                )
            )

        asset_relationships: list[dict[str, Any]] = []
        for row in all_relationships:
            target_for_category = (
                row["resolved_target"]
                if row["target_mode"] == "Internal"
                else row["target"]
            )
            category = _relationship_category(row["type"], target_for_category)
            if category not in {"media", "chart", "embedded_or_ole", "external_data"}:
                continue
            asset_relationships.append(
                {
                    "source_part": row["source_part"] or "[package]",
                    "relationship_id": row["id"],
                    "relationship_type": row["type"],
                    "category": category,
                    "target_mode": row["target_mode"],
                    "target_part": (
                        row["resolved_target"]
                        if row["target_mode"] == "Internal"
                        else "EXTERNAL_REDACTED"
                    ),
                    "target_sha256": (
                        hashlib.sha256(row["target"].encode("utf-8")).hexdigest()
                        if row["target_mode"] == "External"
                        else (
                            _sha256_member(package, row["resolved_target"])
                            if row["resolved_target"] in names
                            else "MISSING"
                        )
                    ),
                    "exists": (
                        row["resolved_target"] in names
                        if row["target_mode"] == "Internal"
                        else "NOT_APPLICABLE"
                    ),
                }
            )

        content_type_values = set(content_types["overrides"].values()) | set(
            content_types["defaults"].values()
        )
        macro_parts = sorted(
            name
            for name in names
            if name.lower().endswith("vbaproject.bin")
            or "vbaproject" in name.lower()
        )
        active_x_parts = sorted(
            name for name in names if name.lower().startswith("ppt/activex/")
        )
        embedded_parts = sorted(
            name for name in names if name.lower().startswith("ppt/embeddings/")
        )
        macro_content_type = any("macroenabled" in value.lower() for value in content_type_values)
        ole_relationships = [
            row
            for row in all_relationships
            if _relationship_category(
                row["type"], row.get("resolved_target", "")
            )
            == "embedded_or_ole"
        ]
        external_data_relationships = [
            row
            for row in all_relationships
            if _relationship_category(row["type"], "") == "external_data"
        ]
        security_flags = {
            "macro_enabled_content_type": macro_content_type,
            "macro_parts": macro_parts,
            "activex_parts": active_x_parts,
            "embedded_or_ole_parts": embedded_parts,
            "ole_relationship_count": len(ole_relationships),
            "external_relationship_count": len(external_relationships),
            "external_data_relationship_count": len(external_data_relationships),
            "executed_embedded_content": False,
            "followed_external_relationships": False,
        }
        if macro_content_type or macro_parts:
            errors.append(
                _issue(
                    "MACRO_CONTENT_PRESENT",
                    "Macro-enabled content is present and was not executed.",
                )
            )
        if active_x_parts:
            errors.append(
                _issue(
                    "ACTIVEX_CONTENT_PRESENT",
                    "ActiveX content is present and was not executed.",
                )
            )
        if embedded_parts or ole_relationships:
            warnings.append(
                _issue(
                    "EMBEDDED_OR_OLE_CONTENT_PRESENT",
                    "Embedded/OLE content is present and was not opened.",
                )
            )

        report: dict[str, Any] = {
            "schema_version": "2.0",
            "source": {
                "extension": source.suffix.lower(),
                "size_bytes": source.stat().st_size,
                "sha256": before_hash,
                "read_only_verified": "PENDING_POST_READ_CHECK",
            },
            "required_opc_parts": {
                "required": list(REQUIRED_PARTS),
                "missing": missing_required,
            },
            "content_types": {
                "presentation": presentation_type or "MISSING",
                "default_count": len(content_types["defaults"]),
                "override_count": len(content_types["overrides"]),
                "duplicate_default_extensions": content_types[
                    "duplicate_default_extensions"
                ],
                "duplicate_override_parts": content_types[
                    "duplicate_override_parts"
                ],
            },
            "relationship_summary": {
                "relationship_count": len(all_relationships),
                "internal_count": sum(
                    row["target_mode"] == "Internal" for row in all_relationships
                ),
                "external_count": len(external_relationships),
                "dangling_count": sum(
                    issue["code"] == "DANGLING_RELATIONSHIP" for issue in errors
                ),
            },
            "external_relationships": external_relationships,
            "slide_order": slide_order,
            "slide_relationship_chains": slide_chains,
            "notes_source_checks": notes_source_checks,
            "asset_relationships": asset_relationships,
            "security_flags": security_flags,
            "object_manifest": object_manifest,
            "canonical_slide_hashes": canonical_hashes,
            "duplicate_pages": duplicate_groups,
            "near_empty_pages": near_empty_pages,
            "unreferenced_slide_parts": unreferenced_slides,
            "errors": errors,
            "warnings": warnings,
            "privacy_contract": {
                "slide_text_returned": False,
                "notes_text_returned": False,
                "alternative_text_returned": False,
                "external_relationships_followed": False,
                "embedded_content_executed": False,
            },
        }

    after_hash = _sha256_file(source)
    if before_hash != after_hash:
        raise OOXMLQAError(
            "OOXML package changed during read-only QA; result discarded"
        )
    report["source"]["sha256_after"] = after_hash
    report["source"]["read_only_verified"] = True
    report["valid"] = not report["errors"]
    report["status"] = (
        "FAIL"
        if report["errors"]
        else ("WARN" if report["warnings"] else "PASS")
    )
    if deck_ir is not None:
        identity = validate_ir_identity_map(
            report,
            deck_ir,
            previous_identity_map=previous_identity_map,
        )
        report["ir_identity_check"] = identity
        if not identity["consistent"]:
            report["errors"].extend(identity["errors"])
            report["valid"] = False
            report["status"] = "FAIL"
        report["warnings"].extend(identity["warnings"])
        if report["valid"] and report["warnings"]:
            report["status"] = "WARN"
    report["issues"] = [
        {**item, "severity": "ERROR"} for item in report["errors"]
    ] + [{**item, "severity": "WARNING"} for item in report["warnings"]]
    return report


def _ir_slide_list(
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]],
) -> list[Mapping[str, Any]]:
    if isinstance(deck_ir, list):
        return list(deck_ir)
    if not isinstance(deck_ir, Mapping):
        raise OOXMLQAError("Deck IR must be a mapping or a list of slide mappings")
    slides = deck_ir.get("slides")
    if isinstance(slides, list):
        return slides
    deck = deck_ir.get("deck")
    if isinstance(deck, Mapping) and isinstance(deck.get("slides"), list):
        return list(deck["slides"])
    raise OOXMLQAError("Deck IR does not contain a slides list")


def build_ir_identity_map(
    qa_report: Mapping[str, Any],
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    """Bind stable IR slide IDs to OOXML slide parts without changing either."""

    ir_slides = _ir_slide_list(deck_ir)
    order_rows = list(qa_report.get("slide_order", []))
    hash_by_part = {
        row["slide_part"]: row["canonical_sha256"]
        for row in qa_report.get("canonical_slide_hashes", [])
    }
    order_by_part = {row["slide_part"]: row for row in order_rows}
    order_by_powerpoint_id = {
        str(row["presentation_slide_id"]): row
        for row in order_rows
        if row.get("presentation_slide_id") not in {None, ""}
    }
    identity_map: list[dict[str, Any]] = []
    for fallback_index, slide in enumerate(ir_slides, start=1):
        slide_id = slide.get("slide_id")
        explicit_part = slide.get("pptx_slide_part") or slide.get("ooxml_slide_part")
        explicit_powerpoint_id = slide.get("powerpoint_slide_id")
        explicit_index = slide.get("slide_index") or slide.get("slide_order")
        selected: Mapping[str, Any] | None = None
        if explicit_part and explicit_part in order_by_part:
            selected = order_by_part[str(explicit_part)]
        elif (
            explicit_powerpoint_id is not None
            and str(explicit_powerpoint_id) in order_by_powerpoint_id
        ):
            selected = order_by_powerpoint_id[str(explicit_powerpoint_id)]
        elif explicit_index is not None:
            try:
                numeric_index = int(explicit_index)
            except (TypeError, ValueError):
                numeric_index = -1
            selected = next(
                (
                    row
                    for row in order_rows
                    if int(row.get("slide_index", -1)) == numeric_index
                ),
                None,
            )
        elif fallback_index <= len(order_rows):
            selected = order_rows[fallback_index - 1]
        part = selected.get("slide_part") if selected else "UNRESOLVED"
        identity_map.append(
            {
                "slide_id": slide_id if slide_id is not None else "MISSING",
                "slide_revision": slide.get("slide_revision", "UNKNOWN"),
                "ir_content_hash": slide.get("content_hash", "UNKNOWN"),
                "slide_index": (
                    selected.get("slide_index") if selected else "UNRESOLVED"
                ),
                "powerpoint_slide_id": (
                    selected.get("presentation_slide_id")
                    if selected
                    else "UNRESOLVED"
                ),
                "pptx_slide_part": part,
                "ooxml_canonical_sha256": hash_by_part.get(
                    str(part), "UNRESOLVED"
                ),
            }
        )
    return identity_map


def validate_ir_identity_map(
    qa_report_or_path: Mapping[str, Any] | str | Path,
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]],
    *,
    previous_identity_map: Iterable[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Check slide-ID uniqueness, binding completeness, and preservation.

    ``previous_identity_map`` is optional.  When present, an unchanged
    ``ir_content_hash`` must retain the same canonical OOXML hash even if its
    page position changes.
    """

    if isinstance(qa_report_or_path, (str, Path)):
        report = inspect_pptx_ooxml(qa_report_or_path)
    else:
        report = qa_report_or_path
    ir_slides = _ir_slide_list(deck_ir)
    identity_map = build_ir_identity_map(report, ir_slides)
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    if len(identity_map) != len(report.get("slide_order", [])):
        errors.append(
            _issue(
                "IR_SLIDE_COUNT_MISMATCH",
                "Deck IR slide count does not match the presentation slide count.",
            )
        )
    slide_ids = [row["slide_id"] for row in identity_map]
    missing_count = sum(slide_id in {None, "", "MISSING"} for slide_id in slide_ids)
    if missing_count:
        errors.append(
            _issue(
                "IR_SLIDE_ID_MISSING",
                f"{missing_count} IR slide(s) do not have a stable slide_id.",
            )
        )
    duplicates = sorted(
        str(slide_id)
        for slide_id, count in Counter(slide_ids).items()
        if slide_id not in {None, "", "MISSING"} and count > 1
    )
    if duplicates:
        errors.append(
            _issue(
                "IR_SLIDE_ID_DUPLICATE",
                "Deck IR contains duplicate stable slide_id values.",
            )
        )
    unresolved = [
        row for row in identity_map if row["pptx_slide_part"] == "UNRESOLVED"
    ]
    if unresolved:
        errors.append(
            _issue(
                "IR_IDENTITY_BINDING_UNRESOLVED",
                f"{len(unresolved)} IR slide(s) could not be bound to OOXML slide parts.",
            )
        )

    if previous_identity_map is not None:
        previous_by_id = {
            row.get("slide_id"): row
            for row in previous_identity_map
            if row.get("slide_id") not in {None, "", "MISSING"}
        }
        for current in identity_map:
            previous = previous_by_id.get(current["slide_id"])
            if previous is None:
                continue
            previous_ir_hash = previous.get("ir_content_hash", "UNKNOWN")
            current_ir_hash = current.get("ir_content_hash", "UNKNOWN")
            if (
                previous_ir_hash not in {"UNKNOWN", None, ""}
                and previous_ir_hash == current_ir_hash
                and previous.get("ooxml_canonical_sha256")
                != current.get("ooxml_canonical_sha256")
            ):
                errors.append(
                    _issue(
                        "UNCHANGED_IR_SLIDE_CHANGED_IN_OOXML",
                        "A stable slide_id with unchanged IR content hash has a different canonical OOXML hash.",
                    )
                )
            if (
                previous.get("pptx_slide_part")
                != current.get("pptx_slide_part")
                and previous_ir_hash == current_ir_hash
            ):
                warnings.append(
                    _issue(
                        "SLIDE_PART_REBOUND_AFTER_REORDER",
                        "A stable slide_id moved to another OOXML slide part; content identity was checked independently.",
                    )
                )

    return {
        "consistent": not errors,
        "identity_map": identity_map,
        "errors": errors,
        "warnings": warnings,
    }


def run_ooxml_qa(
    path: str | Path,
    *,
    known_source_ids: Iterable[str] | None = None,
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]] | None = None,
    previous_identity_map: Iterable[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Workflow-facing alias."""

    return inspect_pptx_ooxml(
        path,
        known_source_ids=known_source_ids,
        deck_ir=deck_ir,
        previous_identity_map=previous_identity_map,
    )


def qa_pptx(
    path: str | Path,
    *,
    known_source_ids: Iterable[str] | None = None,
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Short compatibility alias."""

    return inspect_pptx_ooxml(
        path, known_source_ids=known_source_ids, deck_ir=deck_ir
    )


def inspect_package(
    path: str | Path,
    known_source_ids: Iterable[str] | None = None,
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Simple integration API requested by the v2 workflow runner."""

    return inspect_pptx_ooxml(
        path, known_source_ids=known_source_ids, deck_ir=deck_ir
    )


def write_ooxml_report(
    path: str | Path,
    output_path: str | Path,
    *,
    known_source_ids: Iterable[str] | None = None,
    deck_ir: Mapping[str, Any] | list[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Write a new OOXML QA JSON report without overwriting an existing file."""

    report = inspect_package(
        path, known_source_ids=known_source_ids, deck_ir=deck_ir
    )
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("x", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return report


def validate_ooxml(path: str | Path) -> bool:
    """Return whether structural QA completed without errors."""

    return bool(inspect_pptx_ooxml(path)["valid"])


def assert_ooxml_valid(path: str | Path) -> dict[str, Any]:
    """Return the report or raise when structural errors remain."""

    report = inspect_pptx_ooxml(path)
    if not report["valid"]:
        codes = sorted({item["code"] for item in report["errors"]})
        raise OOXMLQAError("OOXML QA failed: " + ", ".join(codes))
    return report


__all__ = [
    "OOXMLQAError",
    "assert_ooxml_valid",
    "build_ir_identity_map",
    "inspect_pptx_ooxml",
    "inspect_package",
    "qa_pptx",
    "run_ooxml_qa",
    "validate_ir_identity_map",
    "validate_ooxml",
    "write_ooxml_report",
]
