from __future__ import annotations

import hashlib
import json
import re
import uuid
from collections.abc import Mapping, Sequence
from enum import Enum
from pathlib import Path
from typing import Any

from .narrative import (
    NARRATIVE_MODES,
    guard_evidence_strength,
    select_narrative_mode,
)


SCHEMA_VERSION = "2.0"
_DECK_NAMESPACE = uuid.uuid5(
    uuid.NAMESPACE_URL, "https://local.invalid/academic-ppt-workflow/v2"
)
_HEX_64 = re.compile(r"^[0-9a-f]{64}$")

REQUIRED_DECK_FIELDS = frozenset(
    {
        "schema_version",
        "deck_id",
        "deck_version",
        "narrative_mode",
        "brief_hash",
        "slides",
        "canonical_hash",
    }
)
REQUIRED_SLIDE_FIELDS = frozenset(
    {
        "slide_id",
        "slide_revision",
        "slide_role",
        "narrative_mode",
        "key_message",
        "claim_ids",
        "source_ids",
        "layout_id",
        "backend",
        "template_layout_id",
        "visual_assets",
        "editable_object_requirements",
        "speaker_notes",
        "manual_review_required",
        "content_hash",
        "render_hash",
        "logical_slide_key",
        "section_id",
        "figure_ids",
        "managed_object_ids",
        "powerpoint_slide_id",
    }
)

_OPERATIONAL_SLIDE_FIELDS = frozenset(
    {
        "slide_id",
        "slide_revision",
        "content_hash",
        "render_hash",
        "managed_object_ids",
        "powerpoint_slide_id",
    }
)
_PASSTHROUGH_FIELDS = (
    "source_records",
    "chart_data",
    "diagram_spec",
    "formula_manifest",
    "evidence_strength",
    "prohibited_overstatement",
    "visual_asset_path",
    "planned_geometry",
    "layout_family",
    "short_source_label",
    "chart_caption",
    "timeline_spec",
    "matrix_spec",
    "audit_rows",
    "takeaways",
    "problem_gap_objective",
    "flow_spec",
)


def _json_ready(value: Any) -> Any:
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, Path):
        return value.as_posix()
    if isinstance(value, Mapping):
        return {
            str(key): _json_ready(item)
            for key, item in sorted(value.items(), key=lambda pair: str(pair[0]))
        }
    if isinstance(value, (set, frozenset)):
        prepared = [_json_ready(item) for item in value]
        return sorted(
            prepared,
            key=lambda item: json.dumps(
                item, ensure_ascii=False, sort_keys=True, separators=(",", ":")
            ),
        )
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        return [_json_ready(item) for item in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    raise TypeError(f"Deck IR value is not JSON-compatible: {type(value).__name__}")


def canonical_json(value: Any) -> str:
    """Return a canonical JSON representation used for audit hashes."""

    return json.dumps(
        _json_ready(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )


def canonical_json_hash(value: Any) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def deterministic_deck_id(brief: Mapping[str, Any]) -> str:
    """Derive a deck identity from stable project identity fields."""

    explicit_project_key = str(brief.get("project_id", "")).strip()
    project_name = str(brief.get("project_name", "")).strip()
    if not explicit_project_key and not project_name:
        raise ValueError("brief.project_name or brief.project_id is required")
    identity = {
        "project_key": explicit_project_key or project_name,
        "presentation_type": str(brief.get("presentation_type", "")).strip(),
    }
    return "DECK-" + str(uuid.uuid5(_DECK_NAMESPACE, canonical_json(identity))).upper()


def stable_slide_id(deck_id: str, logical_slide_key: str) -> str:
    """Create a stable slide ID that is independent of page number and order."""

    clean_key = str(logical_slide_key).strip()
    if not deck_id or not clean_key:
        raise ValueError("deck_id and logical_slide_key are required")
    slide_uuid = uuid.uuid5(_DECK_NAMESPACE, f"{deck_id}\n{clean_key}")
    return "SLD-" + slide_uuid.hex[:20].upper()


def _id_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        candidates = value.split(";")
    elif isinstance(value, Sequence) and not isinstance(
        value, (str, bytes, bytearray)
    ):
        candidates = list(value)
    else:
        raise TypeError("Identifier collections must be a list or semicolon string")
    result: list[str] = []
    for candidate in candidates:
        item = str(candidate).strip()
        if item and item not in result:
            result.append(item)
    return result


def _normalise_bool(value: Any, *, default: bool) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"yes", "true", "1"}:
            return True
        if lowered in {"no", "false", "0"}:
            return False
    raise ValueError(f"Expected a boolean value, received {value!r}")


def _slug(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")


def _derive_logical_slide_key(slide: Mapping[str, Any]) -> str:
    explicit = str(slide.get("logical_slide_key", "")).strip()
    if explicit:
        return explicit

    claim_ids = sorted(_id_list(slide.get("claim_ids", slide.get("claim_id"))))
    figure_ids = sorted(_id_list(slide.get("figure_ids", slide.get("figure_id"))))
    section_id = str(slide.get("section_id", "")).strip()
    slide_role = str(slide.get("slide_role", "")).strip()
    purpose = str(slide.get("slide_purpose", "")).strip()
    title = str(slide.get("slide_title", "")).strip()

    # No page number or legacy SLD-### value participates in this derivation.
    # Explicit logical keys remain the preferred contract because title/purpose
    # changes can legitimately represent a different semantic slide.
    semantic_anchor = {
        "section_id": section_id,
        "slide_role": slide_role,
        "claim_ids": claim_ids,
        "figure_ids": figure_ids,
        "purpose": purpose,
        "title": title,
    }
    if not any(
        (section_id, slide_role, claim_ids, figure_ids, purpose, title)
    ):
        raise ValueError(
            "A planned slide requires logical_slide_key or semantic fields "
            "(section_id, slide_role, claim_ids, figure_ids, purpose, or title)"
        )
    readable = _slug(section_id or slide_role or title or purpose)[:36] or "slide"
    return f"auto:{readable}:{canonical_json_hash(semantic_anchor)[:16]}"


def _infer_slide_role(slide: Mapping[str, Any]) -> str:
    explicit = str(slide.get("slide_role", "")).strip()
    if explicit:
        return explicit
    layout = str(slide.get("proposed_layout", slide.get("layout_id", ""))).strip()
    visual = str(slide.get("visual_type", "")).strip()
    if layout == "minimal_cover":
        return "cover"
    if visual == "source_map":
        return "evidence_boundary"
    if _id_list(slide.get("claim_ids", slide.get("claim_id"))):
        return "result"
    return "content"


def _slide_content_payload(slide: Mapping[str, Any]) -> dict[str, Any]:
    return {
        key: value
        for key, value in slide.items()
        if key not in _OPERATIONAL_SLIDE_FIELDS
    }


def compute_slide_content_hash(slide: Mapping[str, Any]) -> str:
    return canonical_json_hash(_slide_content_payload(slide))


def _normalise_planned_slide(
    deck_id: str,
    planned: Mapping[str, Any],
    deck_narrative_mode: str,
    existing: Mapping[str, Any] | None,
) -> dict[str, Any]:
    logical_key = _derive_logical_slide_key(planned)
    slide_id = stable_slide_id(deck_id, logical_key)
    slide_mode = str(planned.get("narrative_mode") or deck_narrative_mode)
    if slide_mode not in NARRATIVE_MODES:
        raise ValueError(f"Unknown slide narrative_mode: {slide_mode!r}")

    key_message = str(
        planned.get(
            "key_message",
            planned.get("single_key_message", "INFORMATION_REQUIRED"),
        )
    ).strip()
    title = str(planned.get("slide_title", "")).strip()
    evidence_strength = planned.get("evidence_strength")
    if evidence_strength:
        guard_evidence_strength(title, str(evidence_strength))
        guard_evidence_strength(key_message, str(evidence_strength))

    slide_role = _infer_slide_role(planned)
    section_id = str(planned.get("section_id") or slide_role).strip()
    source_ids = _id_list(planned.get("source_ids"))
    claim_ids = _id_list(planned.get("claim_ids", planned.get("claim_id")))
    figure_ids = _id_list(planned.get("figure_ids", planned.get("figure_id")))
    managed_object_ids = _id_list(planned.get("managed_object_ids"))
    if not managed_object_ids and existing:
        managed_object_ids = _id_list(existing.get("managed_object_ids"))

    slide: dict[str, Any] = {
        "slide_id": slide_id,
        "slide_revision": 1,
        "slide_role": slide_role,
        "narrative_mode": slide_mode,
        "key_message": key_message or "INFORMATION_REQUIRED",
        "claim_ids": claim_ids,
        "source_ids": source_ids,
        "layout_id": str(
            planned.get("layout_id", planned.get("proposed_layout", "content"))
        ).strip()
        or "content",
        "backend": str(planned.get("backend", "pptxgenjs")).strip()
        or "pptxgenjs",
        "template_layout_id": planned.get("template_layout_id"),
        "visual_assets": _json_ready(planned.get("visual_assets", [])),
        "editable_object_requirements": _json_ready(
            planned.get("editable_object_requirements", [])
        ),
        "speaker_notes": str(
            planned.get(
                "speaker_notes",
                planned.get("speaker_note_summary", ""),
            )
        ),
        "manual_review_required": _normalise_bool(
            planned.get("manual_review_required"), default=True
        ),
        "content_hash": "",
        "render_hash": "",
        "logical_slide_key": logical_key,
        "section_id": section_id,
        "figure_ids": figure_ids,
        "managed_object_ids": managed_object_ids,
        "powerpoint_slide_id": planned.get(
            "powerpoint_slide_id",
            existing.get("powerpoint_slide_id") if existing else None,
        ),
        # Current v1 storyboard fields remain available to the CSV and stable
        # PptxGenJS adapters while the canonical fields above become primary.
        "slide_title": title,
        "slide_purpose": str(planned.get("slide_purpose", "")).strip(),
        "single_key_message": key_message or "INFORMATION_REQUIRED",
        "proposed_layout": str(
            planned.get("proposed_layout", planned.get("layout_id", "content"))
        ).strip()
        or "content",
        "visual_type": str(planned.get("visual_type", "text")).strip() or "text",
        "citation_requirement": str(
            planned.get("citation_requirement", "source-bound")
        ).strip()
        or "source-bound",
        "speaker_note_summary": str(
            planned.get(
                "speaker_note_summary",
                planned.get("speaker_notes", ""),
            )
        ),
        "confidence": str(planned.get("confidence", "unknown")).strip()
        or "unknown",
    }
    for field in _PASSTHROUGH_FIELDS:
        if field in planned:
            slide[field] = _json_ready(planned[field])

    content_hash = compute_slide_content_hash(slide)
    slide["content_hash"] = content_hash

    if existing:
        try:
            previous_revision = int(existing.get("slide_revision", 1))
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Existing slide {slide_id} has an invalid slide_revision"
            ) from exc
        if previous_revision < 1:
            raise ValueError(
                f"Existing slide {slide_id} has an invalid slide_revision"
            )
        if existing.get("content_hash") == content_hash:
            slide["slide_revision"] = previous_revision
            slide["render_hash"] = str(existing.get("render_hash", ""))
        else:
            slide["slide_revision"] = previous_revision + 1

    explicit_render_hash = planned.get("render_hash")
    if explicit_render_hash is not None:
        slide["render_hash"] = str(explicit_render_hash)
    return slide


def _existing_slide_indexes(
    existing_ir: Mapping[str, Any] | None,
) -> tuple[dict[str, Mapping[str, Any]], dict[str, Mapping[str, Any]]]:
    by_id: dict[str, Mapping[str, Any]] = {}
    by_key: dict[str, Mapping[str, Any]] = {}
    if not existing_ir:
        return by_id, by_key
    slides = existing_ir.get("slides", [])
    if not isinstance(slides, list):
        raise ValueError("existing_ir.slides must be a list")
    for slide in slides:
        if not isinstance(slide, Mapping):
            raise ValueError("Every existing_ir slide must be a mapping")
        if slide.get("slide_id"):
            by_id[str(slide["slide_id"])] = slide
        if slide.get("logical_slide_key"):
            by_key[str(slide["logical_slide_key"])] = slide
    return by_id, by_key


def _deck_content_signature(ir: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "deck_id": ir["deck_id"],
        "narrative_mode": ir["narrative_mode"],
        "brief_hash": ir["brief_hash"],
        "slides": [
            {
                "slide_id": slide["slide_id"],
                "claim_ids": ";".join(slide["claim_ids"]),
                "content_hash": slide["content_hash"],
            }
            for slide in ir["slides"]
        ],
    }


def build_deck_ir(
    brief: Mapping[str, Any],
    planned_slides: Sequence[Mapping[str, Any]],
    existing_ir: Mapping[str, Any] | None = None,
    narrative_mode: str | None = None,
) -> dict[str, Any]:
    """Build a canonical Deck IR with stable semantic slide identities."""

    if not isinstance(brief, Mapping):
        raise TypeError("brief must be a mapping")
    if not isinstance(planned_slides, Sequence) or isinstance(
        planned_slides, (str, bytes, bytearray)
    ):
        raise TypeError("planned_slides must be a sequence of mappings")

    mode = select_narrative_mode(brief, explicit=narrative_mode)
    derived_deck_id = deterministic_deck_id(brief)
    if existing_ir:
        existing_deck_id = str(existing_ir.get("deck_id", "")).strip()
        if existing_deck_id and existing_deck_id != derived_deck_id:
            raise ValueError(
                "existing_ir.deck_id does not match the deterministic brief identity"
            )
    deck_id = derived_deck_id
    by_id, by_key = _existing_slide_indexes(existing_ir)

    slides: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    seen_keys: set[str] = set()
    for index, planned in enumerate(planned_slides):
        if not isinstance(planned, Mapping):
            raise TypeError(f"planned_slides[{index}] must be a mapping")
        logical_key = _derive_logical_slide_key(planned)
        predicted_id = stable_slide_id(deck_id, logical_key)
        existing = by_id.get(predicted_id) or by_key.get(logical_key)
        slide = _normalise_planned_slide(deck_id, planned, mode, existing)
        if slide["slide_id"] in seen_ids or logical_key in seen_keys:
            raise ValueError(
                f"Duplicate logical slide identity {logical_key!r}; provide unique "
                "logical_slide_key values instead of page-number-based IDs"
            )
        seen_ids.add(slide["slide_id"])
        seen_keys.add(logical_key)
        slides.append(slide)

    draft: dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "deck_id": deck_id,
        "deck_version": 1,
        "narrative_mode": mode,
        "brief_hash": canonical_json_hash(brief),
        "slides": slides,
    }
    if existing_ir:
        try:
            previous_version = int(existing_ir.get("deck_version", 1))
        except (TypeError, ValueError) as exc:
            raise ValueError("existing_ir.deck_version must be a positive integer") from exc
        if previous_version < 1:
            raise ValueError("existing_ir.deck_version must be a positive integer")
        unchanged = _deck_content_signature(draft) == _deck_content_signature(
            existing_ir
        )
        draft["deck_version"] = previous_version if unchanged else previous_version + 1

    draft["canonical_hash"] = canonical_json_hash(draft)
    validate_deck_ir(draft)
    return draft


def validate_deck_ir(ir: Mapping[str, Any]) -> None:
    """Validate structural identity, content hashes, and evidence linkage."""

    if not isinstance(ir, Mapping):
        raise TypeError("Deck IR must be a mapping")
    missing_deck = sorted(REQUIRED_DECK_FIELDS - set(ir))
    if missing_deck:
        raise ValueError("Deck IR missing fields: " + ", ".join(missing_deck))
    if ir["schema_version"] != SCHEMA_VERSION:
        raise ValueError(f"Unsupported Deck IR schema_version: {ir['schema_version']!r}")
    if ir["narrative_mode"] not in NARRATIVE_MODES:
        raise ValueError(f"Unknown deck narrative_mode: {ir['narrative_mode']!r}")
    if not isinstance(ir["deck_version"], int) or ir["deck_version"] < 1:
        raise ValueError("deck_version must be a positive integer")
    if not isinstance(ir["slides"], list):
        raise ValueError("slides must be a list")
    if not _HEX_64.fullmatch(str(ir["brief_hash"])):
        raise ValueError("brief_hash must be a SHA-256 hex digest")

    seen_ids: set[str] = set()
    seen_keys: set[str] = set()
    for index, slide in enumerate(ir["slides"]):
        if not isinstance(slide, Mapping):
            raise ValueError(f"slides[{index}] must be a mapping")
        missing_slide = sorted(REQUIRED_SLIDE_FIELDS - set(slide))
        if missing_slide:
            raise ValueError(
                f"slides[{index}] missing fields: " + ", ".join(missing_slide)
            )
        logical_key = str(slide["logical_slide_key"]).strip()
        expected_id = stable_slide_id(str(ir["deck_id"]), logical_key)
        if slide["slide_id"] != expected_id:
            raise ValueError(
                f"slides[{index}].slide_id is not stable for logical_slide_key"
            )
        if expected_id in seen_ids or logical_key in seen_keys:
            raise ValueError(f"Duplicate slide identity at slides[{index}]")
        seen_ids.add(expected_id)
        seen_keys.add(logical_key)
        if not isinstance(slide["slide_revision"], int) or slide["slide_revision"] < 1:
            raise ValueError(
                f"slides[{index}].slide_revision must be a positive integer"
            )
        if slide["narrative_mode"] not in NARRATIVE_MODES:
            raise ValueError(f"slides[{index}] has an unknown narrative_mode")
        for field in (
            "claim_ids",
            "source_ids",
            "visual_assets",
            "editable_object_requirements",
            "figure_ids",
            "managed_object_ids",
        ):
            if not isinstance(slide[field], list):
                raise ValueError(f"slides[{index}].{field} must be a list")
        if slide["claim_ids"] and not slide["source_ids"]:
            raise ValueError(
                f"slides[{index}] has claim_ids without source_ids; formal claims "
                "must remain traceable"
            )
        if not isinstance(slide["manual_review_required"], bool):
            raise ValueError(
                f"slides[{index}].manual_review_required must be boolean"
            )
        expected_content_hash = compute_slide_content_hash(slide)
        if slide["content_hash"] != expected_content_hash:
            raise ValueError(f"slides[{index}].content_hash does not match content")
        render_hash = str(slide["render_hash"])
        if render_hash and not _HEX_64.fullmatch(render_hash):
            raise ValueError(
                f"slides[{index}].render_hash must be empty or a SHA-256 digest"
            )
        powerpoint_slide_id = slide["powerpoint_slide_id"]
        if powerpoint_slide_id is not None and not isinstance(
            powerpoint_slide_id, (str, int)
        ):
            raise ValueError(
                f"slides[{index}].powerpoint_slide_id must be string, integer, or null"
            )

    canonical_payload = {key: value for key, value in ir.items() if key != "canonical_hash"}
    expected_canonical_hash = canonical_json_hash(canonical_payload)
    if ir["canonical_hash"] != expected_canonical_hash:
        raise ValueError("canonical_hash does not match the canonical Deck IR")


def to_storyboard_rows(ir: Mapping[str, Any]) -> list[dict[str, str]]:
    """Project Deck IR into the v1 storyboard CSV contract."""

    validate_deck_ir(ir)
    rows: list[dict[str, str]] = []
    for slide in ir["slides"]:
        rows.append(
            {
                "slide_id": slide["slide_id"],
                "slide_title": str(slide.get("slide_title", "")),
                "slide_purpose": str(slide.get("slide_purpose", "")),
                "single_key_message": str(slide["key_message"]),
                "source_ids": ";".join(slide["source_ids"]),
                "proposed_layout": str(slide["layout_id"]),
                "layout_family": str(
                    slide.get(
                        "layout_family",
                        slide.get("visual_type", slide["layout_id"]),
                    )
                ),
                "visual_type": str(slide.get("visual_type", "text")),
                "short_source_label": str(
                    slide.get("short_source_label", "来源：输入材料")
                ),
                "citation_requirement": str(
                    slide.get("citation_requirement", "source-bound")
                ),
                "speaker_note_summary": str(slide["speaker_notes"]),
                "prohibited_overstatement": str(
                    slide.get(
                        "prohibited_overstatement",
                        "Do not exceed the registered evidence strength.",
                    )
                ),
                "confidence": str(slide.get("confidence", "unknown")),
                "manual_review_required": (
                    "yes" if slide["manual_review_required"] else "no"
                ),
            }
        )
    return rows


def to_backend_spec(
    ir: Mapping[str, Any],
    brief: Mapping[str, Any],
    design_profile: str | Mapping[str, Any],
) -> dict[str, Any]:
    """Project Deck IR into the stable PptxGenJS/native backend contract."""

    validate_deck_ir(ir)
    if canonical_json_hash(brief) != ir["brief_hash"]:
        raise ValueError("brief does not match Deck IR brief_hash")

    slides: list[dict[str, Any]] = []
    for slide in ir["slides"]:
        item: dict[str, Any] = {
            "slide_id": slide["slide_id"],
            "slide_revision": slide["slide_revision"],
            "slide_role": slide["slide_role"],
            "logical_slide_key": slide["logical_slide_key"],
            "section_id": slide["section_id"],
            "slide_title": slide.get("slide_title", ""),
            "slide_purpose": slide.get("slide_purpose", ""),
            "single_key_message": slide["key_message"],
            "source_ids": list(slide["source_ids"]),
            "claim_ids": list(slide["claim_ids"]),
            "figure_ids": list(slide["figure_ids"]),
            "proposed_layout": slide["layout_id"],
            "layout_id": slide["layout_id"],
            "backend": slide["backend"],
            "template_layout_id": slide["template_layout_id"],
            "visual_type": slide.get("visual_type", "text"),
            "visual_assets": _json_ready(slide["visual_assets"]),
            "editable_object_requirements": _json_ready(
                slide["editable_object_requirements"]
            ),
            "citation_requirement": slide.get(
                "citation_requirement", "source-bound"
            ),
            "speaker_note_summary": slide["speaker_notes"],
            "speaker_notes": slide["speaker_notes"],
            "confidence": slide.get("confidence", "unknown"),
            "manual_review_required": (
                "yes" if slide["manual_review_required"] else "no"
            ),
            "managed_object_ids": list(slide["managed_object_ids"]),
            "powerpoint_slide_id": slide["powerpoint_slide_id"],
            "content_hash": slide["content_hash"],
            "render_hash": slide["render_hash"],
        }
        for field in _PASSTHROUGH_FIELDS:
            if field in slide:
                item[field] = _json_ready(slide[field])
        slides.append(item)

    return {
        "brief": _json_ready(brief),
        "slides": slides,
        "design_profile": _json_ready(design_profile),
        "deck_ir": {
            "schema_version": ir["schema_version"],
            "deck_id": ir["deck_id"],
            "deck_version": ir["deck_version"],
            "narrative_mode": ir["narrative_mode"],
            "canonical_hash": ir["canonical_hash"],
        },
    }
