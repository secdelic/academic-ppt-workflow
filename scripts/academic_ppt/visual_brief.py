from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping

from .utils import load_yaml_compatible


SCHEMA_VERSION = "academic-ppt-visual-brief/1"
PUBLIC_ROUTES = frozenset({"generate", "enhance", "template-fill", "template-create"})
APPROVAL_STATES = frozenset({"draft", "approved", "rejected"})
BACKGROUND_STRATEGIES = frozenset({"light", "tinted", "mixed", "dark-emphasis", "inherit"})

# Low to high.  ``resolve_visual_direction`` applies these layers in this order.
VISUAL_DIRECTION_PRECEDENCE = (
    "auto",
    "canonical_workspace",
    "existing_deck_style",
    "authorized_template",
    "approved_visual_brief",
    "user_explicit_instruction",
)

TOP_LEVEL_FIELDS = frozenset(
    {
        "schema_version",
        "visual_brief_id",
        "approval",
        "scope",
        "direction",
        "style_reference_ids",
        "approved_asset_ids",
        "workspace_hint",
    }
)
DIRECTION_FIELDS = frozenset(
    {
        "direction_name",
        "visual_tone",
        "visual_motif",
        "primary_palette",
        "secondary_palette",
        "background_strategy",
        "typography_personality",
        "hero_style",
        "figure_treatment",
        "chart_treatment",
        "annotation_style",
        "card_style",
        "divider_style",
        "density_pattern",
        "signature_components",
        "allowed_variants",
        "prohibited_styles",
    }
)
LIST_FIELDS = frozenset(
    {
        "visual_tone",
        "primary_palette",
        "secondary_palette",
        "signature_components",
        "allowed_variants",
        "prohibited_styles",
    }
)

# Scientific content belongs to the canonical scientific graph, never this file.
FORBIDDEN_SCIENTIFIC_FIELDS = frozenset(
    {
        "claim",
        "claims",
        "claim_bindings",
        "fact",
        "facts",
        "fact_bindings",
        "data",
        "dataset",
        "chart_data",
        "result",
        "results",
        "estimate",
        "ci_low",
        "ci_high",
        "p_value",
        "interaction_p",
        "numerator",
        "denominator",
        "sample_size",
        "effect",
        "outcome",
        "exposure",
        "source_id",
        "source_ids",
        "source_bindings",
        "citation",
        "citations",
        "patient",
        "diagnosis",
        "scientific_evidence",
    }
)

# These fields always remain native template authority on template-fill.
TEMPLATE_PROTECTED_FIELDS = frozenset(
    {
        "master",
        "master_id",
        "master_roles",
        "layout",
        "layout_ids",
        "theme_fonts",
        "theme_colors",
        "brand_colors",
        "logo",
        "logo_policy",
        "footer",
        "footer_policy",
        "page_number",
        "page_number_policy",
        "animation",
        "animation_policy",
    }
)


class VisualBriefError(ValueError):
    """Raised when the presentation-only visual contract is unsafe or invalid."""


def _normalise_key(value: Any) -> str:
    return str(value).strip().lower().replace("-", "_").replace(" ", "_")


def _walk_keys(value: Any, path: str = "$") -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    if isinstance(value, Mapping):
        for key, child in value.items():
            key_text = str(key)
            child_path = f"{path}.{key_text}"
            found.append((child_path, _normalise_key(key_text)))
            found.extend(_walk_keys(child, child_path))
    elif isinstance(value, (list, tuple)):
        for index, child in enumerate(value):
            found.extend(_walk_keys(child, f"{path}[{index}]"))
    return found


def _is_nonempty_text(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _validate_text_list(value: Any, path: str, *, required: bool = False) -> list[str]:
    if value is None and not required:
        return []
    if not isinstance(value, list) or (required and not value):
        return [f"{path} must be a{' non-empty' if required else ''} list"]
    errors: list[str] = []
    for index, item in enumerate(value):
        if not _is_nonempty_text(item):
            errors.append(f"{path}[{index}] must be a non-empty string")
    if len({str(item) for item in value}) != len(value):
        errors.append(f"{path} must not contain duplicates")
    return errors


def _validate_direction(
    direction: Any,
    path: str = "$.direction",
    *,
    require_complete: bool = True,
) -> list[str]:
    if not isinstance(direction, Mapping):
        return [f"{path} must be a mapping"]
    errors: list[str] = []
    unknown = sorted(set(direction) - DIRECTION_FIELDS)
    if unknown:
        errors.append(f"{path} contains unsupported field(s): {', '.join(unknown)}")
    for required in ("direction_name", "visual_motif"):
        if require_complete and not _is_nonempty_text(direction.get(required)):
            errors.append(f"{path}.{required} must be a non-empty string")
        elif required in direction and not _is_nonempty_text(direction.get(required)):
            errors.append(f"{path}.{required} must be a non-empty string")
    errors.extend(
        _validate_text_list(
            direction.get("visual_tone"),
            f"{path}.visual_tone",
            required=require_complete,
        )
    )
    for field in LIST_FIELDS - {"visual_tone"}:
        errors.extend(_validate_text_list(direction.get(field), f"{path}.{field}"))
    for palette_name in ("primary_palette", "secondary_palette"):
        palette = direction.get(palette_name, [])
        if isinstance(palette, list):
            for index, colour in enumerate(palette):
                if not isinstance(colour, str) or len(colour) != 7 or not colour.startswith("#"):
                    errors.append(f"{path}.{palette_name}[{index}] must be #RRGGBB")
                    continue
                try:
                    int(colour[1:], 16)
                except ValueError:
                    errors.append(f"{path}.{palette_name}[{index}] must be #RRGGBB")
    background = direction.get("background_strategy")
    if background is not None and background not in BACKGROUND_STRATEGIES:
        errors.append(
            f"{path}.background_strategy must be one of: "
            + ", ".join(sorted(BACKGROUND_STRATEGIES))
        )
    for field in DIRECTION_FIELDS - LIST_FIELDS - {"direction_name", "visual_motif", "background_strategy"}:
        value = direction.get(field)
        if value is not None and not _is_nonempty_text(value):
            errors.append(f"{path}.{field} must be a non-empty string")
    return errors


def validate_visual_brief(value: Mapping[str, Any] | Any) -> list[str]:
    """Return deterministic contract errors without touching scientific state."""

    if not isinstance(value, Mapping):
        return ["$ must be a mapping"]
    errors: list[str] = []
    unknown = sorted(set(value) - TOP_LEVEL_FIELDS)
    if unknown:
        errors.append("$ contains unsupported field(s): " + ", ".join(unknown))
    if value.get("schema_version") != SCHEMA_VERSION:
        errors.append(f"$.schema_version must equal {SCHEMA_VERSION!r}")
    brief_id = value.get("visual_brief_id")
    if not _is_nonempty_text(brief_id) or len(str(brief_id)) > 80:
        errors.append("$.visual_brief_id must be a non-empty string of at most 80 characters")

    approval = value.get("approval")
    if not isinstance(approval, Mapping):
        errors.append("$.approval must be a mapping")
    else:
        approval_unknown = sorted(set(approval) - {"status", "approved_by", "approved_at"})
        if approval_unknown:
            errors.append("$.approval contains unsupported field(s): " + ", ".join(approval_unknown))
        if approval.get("status") not in APPROVAL_STATES:
            errors.append("$.approval.status must be draft, approved, or rejected")
        if approval.get("status") == "approved" and not _is_nonempty_text(approval.get("approved_by")):
            errors.append("$.approval.approved_by is required when status is approved")
        for key in ("approved_by", "approved_at"):
            item = approval.get(key)
            if item is not None and not _is_nonempty_text(item):
                errors.append(f"$.approval.{key} must be a non-empty string or null")

    scope = value.get("scope")
    if scope is not None:
        if not isinstance(scope, Mapping):
            errors.append("$.scope must be a mapping")
        else:
            scope_unknown = sorted(set(scope) - {"route", "redesign_requested"})
            if scope_unknown:
                errors.append("$.scope contains unsupported field(s): " + ", ".join(scope_unknown))
            if scope.get("route") is not None and scope.get("route") not in PUBLIC_ROUTES:
                errors.append("$.scope.route is not a supported user-facing route")
            if scope.get("redesign_requested") is not None and not isinstance(scope.get("redesign_requested"), bool):
                errors.append("$.scope.redesign_requested must be boolean")

    errors.extend(_validate_direction(value.get("direction")))
    for field in ("style_reference_ids", "approved_asset_ids"):
        errors.extend(_validate_text_list(value.get(field), f"$.{field}"))
    workspace_hint = value.get("workspace_hint")
    if workspace_hint is not None and not _is_nonempty_text(workspace_hint):
        errors.append("$.workspace_hint must be a non-empty string or null")

    for path, key in _walk_keys(value):
        if key in FORBIDDEN_SCIENTIFIC_FIELDS:
            errors.append(f"{path} is scientific content and is prohibited in visual_brief")
        if key in TEMPLATE_PROTECTED_FIELDS:
            errors.append(f"{path} attempts to override a template-protected field")
    return sorted(set(errors))


def load_visual_brief(path: Path, *, require_approved: bool = False) -> dict[str, Any]:
    """Load and validate a visual brief; optionally require its human gate."""

    value = load_yaml_compatible(path)
    errors = validate_visual_brief(value)
    if errors:
        raise VisualBriefError("Invalid visual_brief: " + "; ".join(errors))
    if require_approved and value["approval"]["status"] != "approved":
        raise VisualBriefError("visual_brief is not approved by a human")
    return value


def _direction_from_source(value: Mapping[str, Any] | None) -> dict[str, Any]:
    if not isinstance(value, Mapping):
        return {}
    for field in ("direction", "style"):
        nested = value.get(field)
        if isinstance(nested, Mapping):
            return deepcopy(dict(nested))
    return deepcopy(dict(value))


def _merge_layer(target: dict[str, Any], layer: Mapping[str, Any]) -> None:
    for key, value in layer.items():
        if isinstance(value, Mapping) and isinstance(target.get(key), Mapping):
            nested = dict(target[key])
            _merge_layer(nested, value)
            target[key] = nested
        else:
            target[key] = deepcopy(value)


def _approved_direction(visual_brief: Mapping[str, Any] | None) -> dict[str, Any]:
    if visual_brief is None:
        return {}
    errors = validate_visual_brief(visual_brief)
    if errors:
        raise VisualBriefError("Invalid visual_brief: " + "; ".join(errors))
    if visual_brief["approval"]["status"] != "approved":
        return {}
    return _direction_from_source(visual_brief)


def resolve_visual_direction(
    *,
    route: str,
    user_explicit_instruction: Mapping[str, Any] | None = None,
    visual_brief: Mapping[str, Any] | None = None,
    authorized_template: Mapping[str, Any] | None = None,
    existing_deck_style: Mapping[str, Any] | None = None,
    canonical_workspace: Mapping[str, Any] | None = None,
    auto_direction: Mapping[str, Any] | None = None,
    redesign_requested: bool = False,
    changed_slide_count: int | None = None,
) -> dict[str, Any]:
    """Resolve presentation direction without creating a second Visual IR.

    The returned object is a routing/audit decision.  It does not contain or
    modify scientific evidence.  Layers are applied from low to high priority.
    Formal template master/layout/brand attributes are always re-applied last
    for ``template-fill``.
    """

    if route not in PUBLIC_ROUTES:
        raise VisualBriefError(f"Unsupported user-facing route: {route!r}")
    if changed_slide_count is not None and (not isinstance(changed_slide_count, int) or changed_slide_count < 0):
        raise VisualBriefError("changed_slide_count must be a non-negative integer or null")

    approved_direction = _approved_direction(visual_brief)
    explicit_direction = _direction_from_source(user_explicit_instruction)
    if explicit_direction:
        explicit_errors = _validate_direction(
            explicit_direction,
            "$.user_explicit_instruction",
            require_complete=False,
        )
        scientific = [
            f"{path} is scientific content and is prohibited in visual direction"
            for path, key in _walk_keys(explicit_direction, "$.user_explicit_instruction")
            if key in FORBIDDEN_SCIENTIFIC_FIELDS
        ]
        if explicit_errors or scientific:
            raise VisualBriefError("Invalid explicit visual instruction: " + "; ".join(explicit_errors + scientific))

    template_authorized = bool(
        isinstance(authorized_template, Mapping)
        and authorized_template.get("authorized", True) is True
    )
    template_style = _direction_from_source(authorized_template) if template_authorized else {}
    if template_style:
        template_style.pop("authorized", None)
        template_style.pop("protected_brand", None)

    layers = {
        "auto": _direction_from_source(auto_direction),
        "canonical_workspace": _direction_from_source(canonical_workspace),
        "existing_deck_style": _direction_from_source(existing_deck_style),
        "authorized_template": template_style,
        "approved_visual_brief": approved_direction,
        "user_explicit_instruction": explicit_direction,
    }

    small_enhance_inheritance = bool(
        route == "enhance"
        and not redesign_requested
        and (changed_slide_count is None or changed_slide_count <= 10)
        and not approved_direction
        and not explicit_direction
        and layers["existing_deck_style"]
    )

    resolved: dict[str, Any] = {}
    applied_sources: list[str] = []
    for source in VISUAL_DIRECTION_PRECEDENCE:
        layer = layers[source]
        if layer:
            _merge_layer(resolved, layer)
            applied_sources.append(source)

    primary_authority = applied_sources[-1] if applied_sources else "none"
    template_protected: dict[str, Any] = {}
    if route == "template-fill" and template_authorized:
        candidate = authorized_template.get("protected_brand", {})
        if isinstance(candidate, Mapping):
            template_protected = {
                key: deepcopy(value)
                for key, value in candidate.items()
                if _normalise_key(key) in TEMPLATE_PROTECTED_FIELDS
            }
        # Accept protected keys at top level for adapters that expose them there.
        template_protected.update(
            {
                key: deepcopy(value)
                for key, value in authorized_template.items()
                if _normalise_key(key) in TEMPLATE_PROTECTED_FIELDS
            }
        )
        _merge_layer(resolved, template_protected)

    return {
        "contract_version": "academic-ppt-visual-direction-resolution/1",
        "route": route,
        "precedence_high_to_low": list(reversed(VISUAL_DIRECTION_PRECEDENCE)),
        "applied_sources_low_to_high": applied_sources,
        "primary_authority": primary_authority,
        "resolved_direction": resolved,
        "template_brand_authority": "authorized_template" if template_protected else None,
        "template_protected_fields": sorted(template_protected),
        "existing_deck_inherited": small_enhance_inheritance,
        "art_director_required": bool(
            redesign_requested and not approved_direction and not explicit_direction
        ),
        "human_visual_brief_approved": bool(approved_direction),
        "scientific_authority": "unchanged_canonical_scientific_core",
    }


def visual_direction_to_style_profile(
    resolution: Mapping[str, Any],
) -> dict[str, Any] | None:
    """Project approved presentation tokens into the existing renderer input.

    This is deliberately a thin adapter to the established ``style_profile``
    contract. It creates no new IR and carries no claims, data, or sources.
    Formal template-fill never uses this projection because the native
    template remains authoritative there.
    """

    if resolution.get("route") in {"template-fill", "enhance"}:
        return None
    direction = resolution.get("resolved_direction", {})
    if not isinstance(direction, Mapping) or not direction:
        return None
    primary = direction.get("primary_palette", [])
    secondary = direction.get("secondary_palette", [])
    colours = [
        str(value).lstrip("#").upper()
        for value in [*primary, *secondary]
        if isinstance(value, str) and len(value.lstrip("#")) == 6
    ]

    def luminance(colour: str) -> float:
        channels = [int(colour[index : index + 2], 16) / 255 for index in (0, 2, 4)]
        linear = [item / 12.92 if item <= 0.04045 else ((item + 0.055) / 1.055) ** 2.4 for item in channels]
        return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

    dark = min(colours, key=luminance) if colours else "102A43"
    light_candidates = [item for item in colours if luminance(item) >= 0.72]
    light = max(light_candidates, key=luminance) if light_candidates else "F7FAFC"
    accents = [item for item in colours if item not in {dark, light}]
    while len(accents) < 4:
        accents.extend(
            item
            for item in ("0F766E", "627D98", "D97706", "2563EB", "70AD47")
            if item not in accents and item not in {dark, light}
        )
    tokens = {
        "lt1": light,
        "dk2": dark,
        "accent1": accents[0],
        "accent2": accents[1],
        "accent4": accents[2],
        "accent5": accents[3],
        "accent6": "70AD47",
        "lt2": "D9E2EC",
    }
    return {
        "schema_version": "2.0",
        "reference_mode": "approved-visual-brief",
        "source": {"sha256": "PRESENTATION_ONLY_CONTRACT", "read_only_verified": True},
        "theme_colors": [
            {"slot": slot, "value": value}
            for slot, value in tokens.items()
        ],
        "visual_direction_resolution": dict(resolution),
        "scientific_content_included": False,
    }


__all__ = [
    "SCHEMA_VERSION",
    "VISUAL_DIRECTION_PRECEDENCE",
    "VisualBriefError",
    "load_visual_brief",
    "resolve_visual_direction",
    "validate_visual_brief",
    "visual_direction_to_style_profile",
]
