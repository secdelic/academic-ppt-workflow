from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from types import MappingProxyType
from typing import Any, Mapping, Sequence


class RouteValidationError(ValueError):
    """Raised when a v2 request cannot be assigned to one safe top-level route."""


class TopLevelRoute(str, Enum):
    GENERATE = "generate"
    CREATE_STYLE_PROFILE = "create-style-profile"
    FILL_TEMPLATE = "fill-template"
    ENHANCE_EXISTING = "enhance-existing"


ROUTE_VALUES = frozenset(route.value for route in TopLevelRoute)
REFERENCE_MODES = frozenset(
    {"style-only", "template-fill", "content-reference", "protected-reference"}
)
UPDATE_OPTION_KEYS = ("update_slide", "update_section", "update_figure")
ALLOWED_OPTION_KEYS = frozenset(
    {
        "allow_managed_overwrite",
        "backend",
        "audit_full",
        "clinical_privacy_mode",
        "create_backup",
        "cross_renderer_validation",
        "final_delivery",
        "manual_review_approved",
        "narrative_mode",
        "preserve_user_edits",
        "reference_mode",
        "retry_slide",
        "retry_stage",
        "resume_from",
        "update_figure",
        "update_section",
        "update_slide",
        "use_network",
        "workflow_mode",
    }
)


@dataclass(frozen=True, slots=True)
class RouteRequest:
    """Generic, immutable request passed to the v2 route controller.

    The request intentionally contains no executable callbacks.  It describes
    local inputs and requested behavior; a later orchestrator decides which
    implementation to invoke only after :func:`validate_route_request` passes.
    """

    route: Any
    brief: Mapping[str, Any]
    input_paths: tuple[str | Path, ...] = ()
    reference_path: str | Path | None = None
    template_path: str | Path | None = None
    existing_deck_path: str | Path | None = None
    style_profile_path: str | Path | None = None
    output_path: str | Path | None = None
    options: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "RouteRequest":
        """Create a request from a CLI/API-style mapping.

        A caller may provide exactly one truthy route flag instead of ``route``.
        Multiple truthy flags are retained as a tuple so validation reports the
        ambiguity rather than silently choosing one.
        """

        route: Any = value.get("route")
        if route is None:
            route_flags = [
                candidate
                for candidate in ROUTE_VALUES
                if value.get(candidate) is True
                or value.get(candidate.replace("-", "_")) is True
            ]
            route = route_flags[0] if len(route_flags) == 1 else tuple(route_flags)

        raw_inputs = value.get("input_paths", value.get("sources", ()))
        if raw_inputs is None:
            inputs: tuple[str | Path, ...] = ()
        elif isinstance(raw_inputs, (str, Path)):
            inputs = (raw_inputs,)
        else:
            inputs = tuple(raw_inputs)

        return cls(
            route=route,
            brief=value.get("brief", {}),
            input_paths=inputs,
            reference_path=value.get("reference_path"),
            template_path=value.get("template_path"),
            existing_deck_path=value.get("existing_deck_path"),
            style_profile_path=value.get("style_profile_path"),
            output_path=value.get("output_path"),
            options=value.get("options", {}),
        )


def _normalise_route(value: Any) -> TopLevelRoute:
    if isinstance(value, TopLevelRoute):
        return value
    if not isinstance(value, str):
        raise RouteValidationError(
            "Exactly one scalar top-level route is required; route lists or "
            "simultaneous route flags are not allowed."
        )
    try:
        return TopLevelRoute(value.strip())
    except ValueError as exc:
        allowed = ", ".join(sorted(ROUTE_VALUES))
        raise RouteValidationError(
            f"Unknown top-level route {value!r}; expected one of: {allowed}."
        ) from exc


def _validate_suffix(value: str | Path | None, field_name: str, allowed: set[str]) -> None:
    if value is None:
        return
    suffix = Path(value).suffix.lower()
    if suffix not in allowed:
        choices = ", ".join(sorted(allowed))
        raise RouteValidationError(
            f"{field_name} must use one of {choices}; received {suffix or '<none>'}."
        )


def _require_absent(request: RouteRequest, route: TopLevelRoute, fields: Sequence[str]) -> None:
    present = [name for name in fields if getattr(request, name) is not None]
    if present:
        raise RouteValidationError(
            f"Route {route.value!r} cannot also set: {', '.join(present)}."
        )


def validate_route_request(
    request: RouteRequest | Mapping[str, Any],
) -> RouteRequest:
    """Validate and return a normalised, immutable route request.

    This function validates contracts only; it does not open or modify a file.
    File existence is deliberately checked by the intake stage so callers can
    validate a request before materialising staged copies.
    """

    if isinstance(request, Mapping):
        request = RouteRequest.from_mapping(request)
    if not isinstance(request, RouteRequest):
        raise TypeError("request must be a RouteRequest or mapping")

    route = _normalise_route(request.route)
    if not isinstance(request.brief, Mapping):
        raise RouteValidationError("brief must be a mapping")
    if not str(request.brief.get("project_name", "")).strip():
        raise RouteValidationError("brief.project_name is required")
    if not isinstance(request.options, Mapping):
        raise RouteValidationError("options must be a mapping")

    unknown_options = sorted(set(request.options) - ALLOWED_OPTION_KEYS)
    if unknown_options:
        raise RouteValidationError(
            "Unknown route option(s): " + ", ".join(unknown_options)
        )

    options = dict(request.options)
    if options.get("use_network") not in (None, False):
        raise RouteValidationError(
            "use_network is prohibited by the controlled v2 route contract"
        )
    if options.get("preserve_user_edits") not in (None, True, False):
        raise RouteValidationError("preserve_user_edits must be boolean")
    if options.get("allow_managed_overwrite") not in (None, True, False):
        raise RouteValidationError("allow_managed_overwrite must be boolean")
    if options.get("create_backup") not in (None, True, False):
        raise RouteValidationError("create_backup must be boolean")
    if options.get("backend") not in (None, "pptxgenjs", "native-template-fill"):
        raise RouteValidationError(
            "backend must be 'pptxgenjs' or 'native-template-fill'"
        )

    reference_mode = options.get("reference_mode", "style-only")
    if reference_mode not in REFERENCE_MODES:
        allowed = ", ".join(sorted(REFERENCE_MODES))
        raise RouteValidationError(
            f"reference_mode must be one of: {allowed}"
        )

    updates = [key for key in UPDATE_OPTION_KEYS if options.get(key)]
    if len(updates) > 1:
        raise RouteValidationError(
            "At most one update target may be selected: "
            + ", ".join(UPDATE_OPTION_KEYS)
        )
    for key in (*UPDATE_OPTION_KEYS, "resume_from", "narrative_mode"):
        if key in options and options[key] is not None:
            if not isinstance(options[key], str) or not options[key].strip():
                raise RouteValidationError(f"{key} must be a non-empty string")
    workflow_mode = options.get("workflow_mode")
    if workflow_mode not in (None, "full_validation", "fast_enhance"):
        raise RouteValidationError(
            "workflow_mode must be 'full_validation' or 'fast_enhance'"
        )
    for key in (
        "audit_full",
        "clinical_privacy_mode",
        "cross_renderer_validation",
        "final_delivery",
        "manual_review_approved",
    ):
        if options.get(key) not in (None, True, False):
            raise RouteValidationError(f"{key} must be boolean")
    for key in ("retry_slide", "retry_stage"):
        if options.get(key) is not None and (
            not isinstance(options[key], str) or not options[key].strip()
        ):
            raise RouteValidationError(f"{key} must be a non-empty string")
    if options.get("retry_slide") and options.get("retry_stage"):
        raise RouteValidationError(
            "retry_slide and retry_stage are mutually exclusive"
        )
    if options.get("manual_review_approved") and not options.get("final_delivery"):
        raise RouteValidationError(
            "manual_review_approved requires final_delivery=true"
        )

    _validate_suffix(request.reference_path, "reference_path", {".pptx", ".potx"})
    _validate_suffix(request.template_path, "template_path", {".pptx", ".potx"})
    _validate_suffix(request.existing_deck_path, "existing_deck_path", {".pptx"})
    _validate_suffix(request.style_profile_path, "style_profile_path", {".json"})

    if route is TopLevelRoute.GENERATE:
        _require_absent(
            request,
            route,
            ("reference_path", "template_path", "existing_deck_path"),
        )
        if updates:
            raise RouteValidationError(
                "Incremental update targets require route 'enhance-existing'"
            )
        if options.get("backend") == "native-template-fill":
            raise RouteValidationError(
                "native-template-fill backend requires route 'fill-template'"
            )
    elif route is TopLevelRoute.CREATE_STYLE_PROFILE:
        if request.reference_path is None:
            raise RouteValidationError(
                "create-style-profile requires reference_path (.pptx or .potx)"
            )
        _require_absent(
            request,
            route,
            ("template_path", "existing_deck_path", "style_profile_path"),
        )
        if updates:
            raise RouteValidationError(
                "create-style-profile cannot select an incremental update target"
            )
        if reference_mode == "template-fill":
            raise RouteValidationError(
                "reference_mode 'template-fill' requires route 'fill-template'"
            )
    elif route is TopLevelRoute.FILL_TEMPLATE:
        if request.template_path is None:
            raise RouteValidationError(
                "fill-template requires template_path (.pptx or .potx)"
            )
        _require_absent(request, route, ("reference_path", "existing_deck_path"))
        if updates:
            raise RouteValidationError(
                "fill-template cannot select an incremental update target"
            )
    elif route is TopLevelRoute.ENHANCE_EXISTING:
        if request.existing_deck_path is None:
            raise RouteValidationError(
                "enhance-existing requires existing_deck_path (.pptx)"
            )
        _require_absent(request, route, ("reference_path", "template_path"))
        if options.get("preserve_user_edits") is False and not options.get(
            "allow_managed_overwrite"
        ):
            raise RouteValidationError(
                "Disabling preserve_user_edits requires allow_managed_overwrite=true"
            )
    has_retry = bool(options.get("retry_slide") or options.get("retry_stage"))
    if has_retry and route is not TopLevelRoute.ENHANCE_EXISTING:
        raise RouteValidationError(
            "retry_slide and retry_stage require route 'enhance-existing'"
        )
    if has_retry and workflow_mode == "full_validation":
        raise RouteValidationError(
            "retry_slide and retry_stage are fast_enhance recovery controls"
        )
    if workflow_mode == "fast_enhance" and route is not TopLevelRoute.ENHANCE_EXISTING:
        raise RouteValidationError(
            "workflow_mode 'fast_enhance' is restricted to route 'enhance-existing'"
        )

    # A fresh mapping prevents later caller mutation from changing an already
    # validated request.  The route itself is normalised to its canonical value.
    return RouteRequest(
        route=route.value,
        brief=MappingProxyType(dict(request.brief)),
        input_paths=tuple(request.input_paths),
        reference_path=request.reference_path,
        template_path=request.template_path,
        existing_deck_path=request.existing_deck_path,
        style_profile_path=request.style_profile_path,
        output_path=request.output_path,
        options=MappingProxyType(options),
    )
