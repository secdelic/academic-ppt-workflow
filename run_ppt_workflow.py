from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "scripts"))

from academic_ppt.runner import STAGES, execute  # noqa: E402


PUBLIC_ROUTE_MAP = {
    "generate": "generate",
    "template-create": "create-style-profile",
    "template-fill": "fill-template",
    "enhance": "enhance-existing",
}
LEGACY_ROUTE_MAP = {
    "create-style-profile": "create-style-profile",
    "fill-template": "fill-template",
    "enhance-existing": "enhance-existing",
}
QUALITY_VALUES = ("quick", "validated", "full")


class _StoreExplicit(argparse.Action):
    """Store a value while remembering that it came from the public CLI."""

    def __call__(self, parser, namespace, values, option_string=None):  # type: ignore[no-untyped-def]
        setattr(namespace, self.dest, values)
        setattr(namespace, f"{self.dest}_explicit", True)


def normalise_public_request(args: argparse.Namespace) -> argparse.Namespace:
    """Translate the four public routes without exposing internal contracts.

    This is intentionally a thin compatibility layer.  It does not create a
    second workflow authority or change any scientific, evidence, rendering,
    or QA implementation.
    """

    try:
        internal_route = {**PUBLIC_ROUTE_MAP, **LEGACY_ROUTE_MAP}[str(args.route)]
    except KeyError as exc:
        raise ValueError(f"Unknown public route: {args.route!r}") from exc
    quality = str(getattr(args, "quality", "validated") or "validated")
    if quality not in QUALITY_VALUES:
        raise ValueError(f"Unknown quality mode: {quality!r}")

    args.public_route = next(
        (name for name, value in PUBLIC_ROUTE_MAP.items() if value == internal_route),
        str(args.route),
    )
    args.route = internal_route
    if getattr(args, "workflow_mode", None) is not None:
        if args.workflow_mode == "fast_enhance" and internal_route != "enhance-existing":
            raise ValueError("fast_enhance is restricted to the enhance route")
    elif quality == "full":
        args.workflow_mode = "full_validation"
    elif internal_route == "enhance-existing":
        args.workflow_mode = "fast_enhance"
        if quality == "validated":
            args.final_delivery = True
    else:
        # Generate/template routes currently retain the validated full path;
        # no reduced generator or template engine is introduced in v2.5.3.
        args.workflow_mode = "full_validation"
    return args


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Source-bound, render-verified academic PPT workflow"
    )
    parser.add_argument(
        "--brief",
        default="brief/presentation_brief.yaml",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--project",
        "--project-root",
        dest="project_root",
        help=(
            "Standard project path, or a project id beneath "
            "PPT_WORKSPACE_HOME/projects"
        ),
    )
    parser.add_argument(
        "--route",
        choices=sorted({**PUBLIC_ROUTE_MAP, **LEGACY_ROUTE_MAP}),
        metavar="{generate,template-create,template-fill,enhance}",
        default="generate",
        action=_StoreExplicit,
        help="User workflow: generate, template-create, template-fill, or enhance",
    )
    parser.add_argument(
        "--quality",
        choices=QUALITY_VALUES,
        default="validated",
        action=_StoreExplicit,
        help="Validation depth exposed to users: quick, validated, or full",
    )
    parser.add_argument("--repo-root", help=argparse.SUPPRESS)
    parser.add_argument("--input-root", help=argparse.SUPPRESS)
    parser.add_argument("--staging-root", help=argparse.SUPPRESS)
    parser.add_argument("--output-root", help=argparse.SUPPRESS)
    parser.add_argument("--audit-root", help=argparse.SUPPRESS)
    parser.add_argument("--run-id", help=argparse.SUPPRESS)
    parser.add_argument("--backend", choices=["pptxgenjs", "artifact_tool"], help=argparse.SUPPRESS)
    parser.add_argument("--node", help=argparse.SUPPRESS)
    parser.add_argument("--node-modules", help=argparse.SUPPRESS)
    parser.add_argument("--soffice", help=argparse.SUPPRESS)
    parser.add_argument("--pdftoppm", help=argparse.SUPPRESS)
    parser.add_argument(
        "--from-stage",
        "--resume-from",
        dest="from_stage",
        choices=STAGES,
        default="preflight",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--resume", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--workflow-mode",
        "--mode",
        dest="workflow_mode",
        choices=["full_validation", "fast_enhance"],
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--final-delivery",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--cross-renderer-validation",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--audit-full",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--retry-slide", help=argparse.SUPPRESS)
    parser.add_argument("--retry-stage", help=argparse.SUPPRESS)
    parser.add_argument(
        "--operation-plan",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--changed-candidate-pptx",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--project-cache-root",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--clinical-privacy-mode",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--export-pdf",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--manual-review-approved",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--reference-pptx", help=argparse.SUPPRESS)
    parser.add_argument(
        "--template",
        dest="reference_pptx",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--ppt",
        dest="existing_pptx",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--input",
        dest="input_root",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--auto-approve-content",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--style-profile", help=argparse.SUPPRESS)
    parser.add_argument(
        "--reference-mode",
        choices=["style-only", "template-fill", "content-reference", "protected-reference"],
        default="style-only",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--existing-pptx", dest="existing_pptx", help=argparse.SUPPRESS)
    parser.add_argument("--previous-deck-ir", help=argparse.SUPPRESS)
    parser.add_argument("--previous-object-manifest", help=argparse.SUPPRESS)
    parser.add_argument("--update-slide", help=argparse.SUPPRESS)
    parser.add_argument("--update-section", help=argparse.SUPPRESS)
    parser.add_argument("--update-figure", help=argparse.SUPPRESS)
    parser.add_argument("--allow-managed-overwrite", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--template-layout-id", help=argparse.SUPPRESS)
    parser.add_argument(
        "--narrative-mode",
        choices=[
            "scientific_problem",
            "conclusion_first",
            "thesis_defense",
            "journal_club",
            "clinical_protocol",
            "methods_teaching",
            "neutral_briefing",
        ],
        help=argparse.SUPPRESS,
    )
    parser.set_defaults(route_explicit=False, quality_explicit=False)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        if args.project_root:
            from academic_ppt.project_interface import apply_project_brief_defaults

            args = apply_project_brief_defaults(args)
        args = normalise_public_request(args)
        public_run = None
        if args.project_root:
            from academic_ppt.project_interface import configure_project_args

            args, public_run = configure_project_args(args)
        output = execute(args)
        if public_run is not None:
            from academic_ppt.project_interface import publish_standard_outputs

            output = publish_standard_outputs(output, public_run)
    except Exception as exc:
        print(f"BLOCKED: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2
    print(f"OUTPUT_DIR={output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
