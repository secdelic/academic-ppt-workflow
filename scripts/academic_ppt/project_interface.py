from __future__ import annotations

import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping

from .style_reference import parse_reference_style
from .utils import load_yaml_compatible, output_timestamp, safe_slug, write_json
from .visual_brief import (
    load_visual_brief,
    resolve_visual_direction,
    visual_direction_to_style_profile,
)


PUBLIC_ROUTES = ("generate", "enhance", "template-fill", "template-create")
PUBLIC_QUALITIES = ("quick", "validated", "full")
_INTERNAL_OUTPUT_NAMES = frozenset(
    {
        "operation_plan.json",
        "changed_candidate_spec.json",
        "changed_candidate_manifest.json",
        "checkpoint.json",
        "run_state.json",
        "slidespec.json",
        "visualspec.json",
    }
)


class ProjectInterfaceError(ValueError):
    """Raised when the user-facing project or output contract is invalid."""


@dataclass(frozen=True, slots=True)
class ProjectPaths:
    root: Path
    input_root: Path
    brief_root: Path
    approved_assets_root: Path
    output_root: Path
    audit_root: Path
    cache_root: Path
    private_root: Path

    @classmethod
    def from_root(cls, root: str | Path) -> "ProjectPaths":
        resolved = Path(root).expanduser().resolve()
        return cls(
            root=resolved,
            input_root=resolved / "input",
            brief_root=resolved / "brief",
            approved_assets_root=resolved / "approved_assets",
            output_root=resolved / "output",
            audit_root=resolved / "audit",
            cache_root=resolved / "cache",
            private_root=resolved / "private",
        )

    def presentation_brief(self) -> Path:
        return self.brief_root / "presentation_brief.yaml"

    def visual_brief(self) -> Path:
        return self.brief_root / "visual_brief.yaml"


@dataclass(frozen=True, slots=True)
class RunOutputPaths:
    project: ProjectPaths
    run_id: str
    route: str
    quality: str
    run_root: Path
    draft: Path
    final: Path
    preview: Path
    audit: Path
    engine_output: Path
    engine_audit: Path
    staging: Path

    def public_contract(self) -> dict[str, Any]:
        """Return only concepts that a daily user needs to understand."""

        if self.quality == "quick":
            primary = (
                ["draft/style_profile.json"]
                if self.route == "template-create"
                else ["draft/deck.pptx"]
            )
            if self.route == "enhance":
                primary.append("audit/changed_slide_qa.md")
        else:
            primary = (
                ["draft/style_profile.json"]
                if self.route == "template-create"
                else [
                    "final/deck.pptx",
                    "final/deck.pdf",
                    "preview/contact_sheet.png",
                    "audit/scientific_qa.md",
                    "audit/visual_qa.md",
                ]
            )
        return {
            "schema_version": "academic-ppt-user-output/1",
            "route": self.route,
            "quality": self.quality,
            "run_id": self.run_id,
            "public_directories": ["draft", "final", "preview", "audit"],
            "required_artifacts": primary,
            "internal_workflow_objects_exposed": False,
        }


def _normalise_public_route(value: str) -> str:
    aliases = {
        "create-style-profile": "template-create",
        "fill-template": "template-fill",
        "enhance-existing": "enhance",
    }
    route = aliases.get(str(value), str(value))
    if route not in PUBLIC_ROUTES:
        raise ProjectInterfaceError(f"Unknown route: {value!r}")
    return route


def _normalise_quality(value: str) -> str:
    quality = str(value)
    if quality not in PUBLIC_QUALITIES:
        raise ProjectInterfaceError(f"Unknown quality: {value!r}")
    return quality


def resolve_project_reference(
    value: str | Path, *, workspace_home: str | Path | None = None
) -> ProjectPaths:
    """Resolve an explicit path or a project id beneath PPT_WORKSPACE_HOME."""

    raw = Path(value).expanduser()
    if raw.is_absolute() or len(raw.parts) > 1 or raw.exists():
        return ProjectPaths.from_root(raw)
    home_raw = workspace_home or os.environ.get("PPT_WORKSPACE_HOME")
    if not home_raw:
        raise ProjectInterfaceError(
            "A project id requires PPT_WORKSPACE_HOME; otherwise pass a project path"
        )
    return ProjectPaths.from_root(Path(home_raw).expanduser() / "projects" / raw)


def apply_project_brief_defaults(args: Any) -> Any:
    """Use route/quality from the project brief unless the CLI set them."""

    project = resolve_project_reference(getattr(args, "project_root"))
    raw_brief = str(getattr(args, "brief", "") or "")
    brief_path = (
        project.presentation_brief()
        if not raw_brief or raw_brief == "brief/presentation_brief.yaml"
        else Path(raw_brief).expanduser()
    )
    if not brief_path.is_absolute():
        brief_path = (project.root / brief_path).resolve()
    if not brief_path.is_file():
        raise ProjectInterfaceError(f"Presentation brief is missing: {brief_path}")
    brief = load_yaml_compatible(brief_path)
    if not bool(getattr(args, "route_explicit", False)) and brief.get("route"):
        args.route = str(brief["route"])
    if not bool(getattr(args, "quality_explicit", False)) and brief.get("quality"):
        args.quality = str(brief["quality"])
    args.brief = str(brief_path)
    return args


def normalise_presentation_brief(value: Mapping[str, Any]) -> dict[str, Any]:
    """Flatten the two user-facing brief sections for the frozen engine.

    The returned mapping preserves every original user field and adds only the
    legacy aliases already consumed by the stable scientific core. It does not
    infer or alter scientific content.
    """

    result = dict(value)
    presentation = value.get("presentation", {})
    content = value.get("content_contract", {})
    privacy = value.get("privacy", {})
    inputs = value.get("inputs", {})
    if isinstance(presentation, Mapping):
        aliases = {
            "type": "presentation_type",
            "objective": "presentation_objective",
            "audience": "audience",
            "language": "language",
            "duration_minutes": "duration_minutes",
            "target_slide_count": "target_slide_count",
            "aspect_ratio": "output_aspect_ratio",
        }
        for source, target in aliases.items():
            if target not in result and presentation.get(source) is not None:
                result[target] = presentation[source]
    if isinstance(content, Mapping):
        for source, target in (
            ("must_include", "must_include"),
            ("must_not_claim", "must_exclude"),
            ("key_message", "key_message"),
            ("speaker_notes_required", "speaker_notes_required"),
            ("appendix_required", "appendix_required"),
        ):
            if target not in result and content.get(source) is not None:
                result[target] = content[source]
    if isinstance(privacy, Mapping):
        result.setdefault(
            "clinical_privacy_mode", privacy.get("mode") == "clinical_private"
        )
    if isinstance(inputs, Mapping):
        if inputs.get("template_path"):
            result.setdefault("template_path", inputs["template_path"])
        if inputs.get("existing_deck_path"):
            result.setdefault("existing_deck_path", inputs["existing_deck_path"])
    return result


def prepare_run_output(
    project: ProjectPaths,
    *,
    route: str,
    quality: str,
    run_id: str | None = None,
    create: bool = True,
) -> RunOutputPaths:
    route = _normalise_public_route(route)
    quality = _normalise_quality(quality)
    identifier = str(run_id or output_timestamp()).strip()
    if not identifier or safe_slug(identifier) != identifier:
        raise ProjectInterfaceError("run_id must be a safe, non-empty identifier")
    run_root = project.output_root / identifier
    result = RunOutputPaths(
        project=project,
        run_id=identifier,
        route=route,
        quality=quality,
        run_root=run_root,
        draft=run_root / "draft",
        final=run_root / "final",
        preview=run_root / "preview",
        audit=project.audit_root / identifier,
        engine_output=project.private_root / "runtime_output" / identifier,
        engine_audit=project.private_root / "runtime_audit" / identifier,
        staging=project.private_root / "staging",
    )
    if create:
        if run_root.exists() or result.audit.exists():
            raise ProjectInterfaceError(
                f"Run output already exists; choose a new run_id: {identifier}"
            )
        for directory in (
            result.draft,
            result.final,
            result.preview,
            result.audit,
            result.engine_output,
            result.engine_audit,
            result.staging,
        ):
            directory.mkdir(parents=True, exist_ok=True)
        write_json(result.audit / "artifact_contract.json", result.public_contract())
    return result


def validate_approved_assets(
    project: ProjectPaths, registry_path: str | Path | None = None
) -> list[dict[str, Any]]:
    """Validate non-evidence assets without registering them as sources.

    The registry is optional.  Every registered path must remain inside the
    project ``approved_assets`` directory and explicitly state that it is not
    scientific evidence.
    """

    if registry_path:
        registry = Path(registry_path).expanduser().resolve()
    else:
        canonical_registry = project.approved_assets_root / "asset_registry.yaml"
        legacy_registry = project.approved_assets_root / "registry.yaml"
        registry = canonical_registry if canonical_registry.exists() else legacy_registry
    if not registry.exists():
        return []
    document = load_yaml_compatible(registry)
    rows = document.get("assets", [])
    if not isinstance(rows, list):
        raise ProjectInterfaceError("approved_assets registry.assets must be a list")
    required = {
        "asset_id",
        "role",
        "source",
        "scientific_evidence",
        "approved_by",
        "approved_at",
    }
    root = project.approved_assets_root.resolve()
    validated: list[dict[str, Any]] = []
    identifiers: set[str] = set()
    for index, raw in enumerate(rows, 1):
        if not isinstance(raw, Mapping):
            raise ProjectInterfaceError(f"approved asset #{index} must be a mapping")
        missing = sorted(required - set(raw))
        if missing:
            raise ProjectInterfaceError(
                f"approved asset #{index} is missing: {', '.join(missing)}"
            )
        asset_id = str(raw["asset_id"]).strip()
        if not asset_id or asset_id in identifiers:
            raise ProjectInterfaceError("approved asset_id must be non-empty and unique")
        identifiers.add(asset_id)
        if raw["scientific_evidence"] is not False:
            raise ProjectInterfaceError(
                f"approved asset {asset_id!r} must set scientific_evidence=false"
            )
        role = str(raw["role"]).strip()
        source = str(raw["source"]).strip()
        approved_by = str(raw["approved_by"]).strip()
        approved_at = str(raw["approved_at"]).strip()
        if not role or not source or not approved_by or not approved_at:
            raise ProjectInterfaceError(
                f"approved asset {asset_id!r} requires role, source, approved_by, and approved_at"
            )
        relative_file = raw.get("file", raw.get("path"))
        if not isinstance(relative_file, str) or not relative_file.strip():
            raise ProjectInterfaceError(
                f"approved asset {asset_id!r} requires a non-empty file path"
            )
        relative_path = Path(relative_file)
        if relative_path.parts and relative_path.parts[0].casefold() == "approved_assets":
            relative_path = Path(*relative_path.parts[1:])
        asset_path = (root / relative_path).resolve()
        try:
            asset_path.relative_to(root)
        except ValueError as exc:
            raise ProjectInterfaceError(
                f"approved asset {asset_id!r} escapes approved_assets"
            ) from exc
        if not asset_path.is_file():
            raise ProjectInterfaceError(
                f"approved asset {asset_id!r} does not exist: {raw['path']}"
            )
        validated.append(
            {
                "asset_id": asset_id,
                "path": asset_path,
                "role": role,
                "source": source,
                "scientific_evidence": False,
                "approved_by": approved_by,
                "approved_at": approved_at,
            }
        )
    return validated


def configure_project_args(args: Any) -> tuple[Any, RunOutputPaths]:
    """Configure the frozen engine from one explicit user project root."""

    project = resolve_project_reference(getattr(args, "project_root"))
    if not project.root.is_dir():
        raise ProjectInterfaceError(f"Project root does not exist: {project.root}")
    default_brief = project.presentation_brief()
    raw_brief = str(getattr(args, "brief", "") or "")
    if not raw_brief or raw_brief == "brief/presentation_brief.yaml":
        args.brief = str(default_brief)
    brief_path = Path(args.brief).expanduser()
    if not brief_path.is_absolute():
        brief_path = (project.root / brief_path).resolve()
        args.brief = str(brief_path)
    if not brief_path.is_file():
        raise ProjectInterfaceError(f"Presentation brief is missing: {brief_path}")
    brief = normalise_presentation_brief(load_yaml_compatible(brief_path))
    normalised_brief_path = project.private_root / "presentation_brief.normalised.yaml"
    write_json(normalised_brief_path, brief)
    args.brief = str(normalised_brief_path)

    route = _normalise_public_route(getattr(args, "public_route", args.route))
    quality = _normalise_quality(getattr(args, "quality", "validated"))
    run = prepare_run_output(
        project,
        route=route,
        quality=quality,
        run_id=getattr(args, "run_id", None),
    )
    args.run_id = run.run_id
    args.input_root = str(project.input_root)
    args.staging_root = str(run.staging)
    args.output_root = str(run.engine_output)
    args.audit_root = str(run.engine_audit)
    args.project_cache_root = str(project.cache_root)

    def resolve_declared(keys: Iterable[str]) -> str | None:
        for key in keys:
            value = brief.get(key)
            if value is None and isinstance(brief.get("inputs"), Mapping):
                value = brief["inputs"].get(key)
            if value:
                path = Path(str(value)).expanduser()
                return str(path.resolve() if path.is_absolute() else (project.root / path).resolve())
        return None

    if route == "enhance" and not getattr(args, "existing_pptx", None):
        args.existing_pptx = resolve_declared(("existing_deck", "existing_deck_path", "existing_pptx"))
    if route in {"template-fill", "template-create"} and not getattr(
        args, "reference_pptx", None
    ):
        args.reference_pptx = resolve_declared(("template", "template_path", "template_pptx"))

    # Presentation direction is resolved before the frozen engine runs. The
    # decision is an audit/style adapter only and never enters Evidence.
    active_visual_brief = None
    if project.visual_brief().is_file():
        active_visual_brief = load_visual_brief(
            project.visual_brief(), require_approved=True
        )
    redesign_requested = bool(
        brief.get("redesign_requested")
        or (isinstance(brief.get("visual_direction"), Mapping) and brief["visual_direction"].get("redesign_requested"))
    )
    existing_style = None
    if route == "enhance" and getattr(args, "existing_pptx", None):
        existing_style = parse_reference_style(
            Path(args.existing_pptx), mode="protected-reference"
        )
    template_style = None
    if route in {"template-fill", "template-create"} and getattr(
        args, "reference_pptx", None
    ):
        template_style = parse_reference_style(
            Path(args.reference_pptx), mode=(
                "template-fill" if route == "template-fill" else "style-only"
            )
        )
        template_style = {
            "authorized": True,
            "style": template_style,
            "protected_brand": {
                "master_roles": "preserve",
                "layout_ids": "preserve",
                "theme_fonts": "preserve",
                "theme_colors": "preserve",
                "logo_policy": "preserve",
                "footer_policy": "preserve",
                "page_number_policy": "preserve",
                "animation_policy": "preserve",
            },
        }
    resolution = resolve_visual_direction(
        route=route,
        user_explicit_instruction=(
            brief.get("explicit_visual_instruction")
            if isinstance(brief.get("explicit_visual_instruction"), Mapping)
            else None
        ),
        visual_brief=active_visual_brief,
        authorized_template=template_style,
        existing_deck_style=existing_style,
        canonical_workspace=(
            brief.get("canonical_workspace")
            if isinstance(brief.get("canonical_workspace"), Mapping)
            else None
        ),
        auto_direction={"direction_name": "canonical-auto"},
        redesign_requested=redesign_requested,
    )
    write_json(run.audit / "visual_direction_resolution.json", resolution)
    if route == "generate":
        projected = visual_direction_to_style_profile(resolution)
        if projected is not None:
            style_path = project.private_root / "resolved_visual_style_profile.json"
            write_json(style_path, projected)
            args.style_profile = str(style_path)

    # This validation is deliberately not returned to the evidence pipeline.
    validate_approved_assets(project)
    return args, run


def _copy_one(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise ProjectInterfaceError(f"Refusing to overwrite run artifact: {destination}")
    shutil.copy2(source, destination)


def _find_first(root: Path, patterns: Iterable[str]) -> Path | None:
    for pattern in patterns:
        candidates = sorted(path for path in root.rglob(pattern) if path.is_file())
        if candidates:
            return candidates[0]
    return None


def _qa_section(source: Path, heading: str) -> str:
    text = source.read_text(encoding="utf-8-sig")
    marker = f"## {heading}"
    if marker not in text:
        return text
    tail = text.split(marker, 1)[1]
    body = tail.split("\n## ", 1)[0].strip()
    return f"# {heading}\n\n{body}\n"


def publish_standard_outputs(engine_output: Path, run: RunOutputPaths) -> Path:
    """Copy user deliverables from a completed engine run into the stable tree."""

    engine_output = Path(engine_output).resolve()
    if not engine_output.is_dir():
        raise ProjectInterfaceError(f"Engine output is missing: {engine_output}")
    pptx = _find_first(engine_output, ("updated.pptx", "*.pptx"))
    pdf = _find_first(engine_output, ("updated.pdf", "*.pdf"))
    contact = _find_first(engine_output, ("contact_sheet.png",))
    changed_previews = sorted({
        path.resolve()
        for pattern in ("changed_preview/*.png", "slide_*.png")
        for path in engine_output.rglob(pattern)
        if path.is_file()
    })
    qa = _find_first(engine_output, ("qa_report.md",))
    summary = _find_first(engine_output, ("execution_summary.md",))
    changed_qa = _find_first(engine_output, ("changed_slide_qa.md",))

    if run.route == "template-create":
        profile = _find_first(engine_output, ("reference_style_profile.json",))
        if profile is None:
            raise ProjectInterfaceError("template-create produced no style profile")
        _copy_one(profile, run.draft / "style_profile.json")
        if summary is not None:
            _copy_one(summary, run.audit / "execution_summary.md")
        return run.run_root
    if pptx is not None:
        target = run.draft / "deck.pptx" if run.quality == "quick" else run.final / "deck.pptx"
        _copy_one(pptx, target)
    if run.quality != "quick":
        if pdf is not None:
            _copy_one(pdf, run.final / "deck.pdf")
        if contact is not None:
            _copy_one(contact, run.preview / "contact_sheet.png")
        elif changed_previews:
            from PIL import Image, ImageOps, ImageDraw

            previews: list[Image.Image] = []
            for path in changed_previews:
                with Image.open(path) as source:
                    previews.append(ImageOps.contain(source.convert("RGB"), (640, 360)))
            columns = 2
            rows = (len(previews) + columns - 1) // columns
            sheet = Image.new("RGB", (columns * 640, rows * 360), "white")
            for index, preview in enumerate(previews):
                x = (index % columns) * 640
                y = (index // columns) * 360
                sheet.paste(preview, (x, y))
                ImageDraw.Draw(sheet).rectangle(
                    (x, y, x + preview.width - 1, y + preview.height - 1),
                    outline="#D9E2EC",
                    width=1,
                )
            sheet.save(run.preview / "contact_sheet.png")
        if qa is not None:
            (run.audit / "scientific_qa.md").write_text(
                _qa_section(qa, "Scientific QA"), encoding="utf-8"
            )
            (run.audit / "visual_qa.md").write_text(
                _qa_section(qa, "Visual QA"), encoding="utf-8"
            )
        elif changed_qa is not None:
            changed_body = changed_qa.read_text(encoding="utf-8-sig").strip()
            (run.audit / "scientific_qa.md").write_text(
                "# Scientific QA\n\n"
                "Changed-slide source/scientific contract:\n\n"
                + changed_body
                + "\n",
                encoding="utf-8",
            )
            (run.audit / "visual_qa.md").write_text(
                "# Visual QA\n\n"
                "Changed-slide geometry/typography/visual contract:\n\n"
                + changed_body
                + "\n",
                encoding="utf-8",
            )
    if summary is not None:
        _copy_one(summary, run.audit / "execution_summary.md")
    if changed_qa is not None:
        _copy_one(changed_qa, run.audit / "changed_slide_qa.md")
    required = run.public_contract()["required_artifacts"]
    result = validate_public_output_tree(run)
    if required and result["status"] != "PASS":
        raise ProjectInterfaceError(
            "Engine completed but public output contract failed: "
            + ", ".join(result["missing"])
        )
    return run.run_root


def validate_public_output_tree(run: RunOutputPaths) -> dict[str, Any]:
    contract = run.public_contract()
    roots = {
        "draft": run.draft,
        "final": run.final,
        "preview": run.preview,
        "audit": run.audit,
    }
    present = {
        path.relative_to(run.run_root).as_posix()
        for key, root in roots.items()
        if key != "audit"
        for path in root.rglob("*")
        if path.is_file()
    }
    present.update(
        f"audit/{path.relative_to(run.audit).as_posix()}"
        for path in run.audit.rglob("*")
        if path.is_file()
    )
    required = set(contract["required_artifacts"])
    hidden = sorted(
        name for name in present if Path(name).name.casefold() in _INTERNAL_OUTPUT_NAMES
    )
    missing = sorted(required - present)
    return {
        "status": "PASS" if not missing and not hidden else "FAIL",
        "missing": missing,
        "internal_artifacts_exposed": hidden,
        "present": sorted(present),
    }


__all__ = [
    "PUBLIC_QUALITIES",
    "PUBLIC_ROUTES",
    "ProjectInterfaceError",
    "ProjectPaths",
    "RunOutputPaths",
    "apply_project_brief_defaults",
    "configure_project_args",
    "prepare_run_output",
    "normalise_presentation_brief",
    "publish_standard_outputs",
    "resolve_project_reference",
    "validate_approved_assets",
    "validate_public_output_tree",
]
