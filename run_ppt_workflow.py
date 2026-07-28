from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "scripts"))

from academic_ppt.runner import STAGES, execute  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Source-bound, render-verified academic PPT workflow"
    )
    parser.add_argument("--brief", default="brief/presentation_brief.yaml")
    parser.add_argument("--repo-root")
    parser.add_argument("--input-root")
    parser.add_argument("--staging-root")
    parser.add_argument("--output-root")
    parser.add_argument("--run-id", help="Unique timestamp-like run identifier")
    parser.add_argument("--backend", choices=["pptxgenjs", "artifact_tool"])
    parser.add_argument("--node")
    parser.add_argument("--node-modules")
    parser.add_argument("--soffice")
    parser.add_argument("--pdftoppm")
    parser.add_argument("--from-stage", choices=STAGES, default="preflight")
    parser.add_argument("--resume", action="store_true")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        output = execute(args)
    except Exception as exc:
        print(f"BLOCKED: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2
    print(f"OUTPUT_DIR={output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
