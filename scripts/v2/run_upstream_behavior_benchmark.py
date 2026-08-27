#!/usr/bin/env python3
"""Run the isolated, local-only Arm B behavior benchmark.

This adapter is not a production backend.  It consumes an already source-bound
canonical build specification, executes no upstream code, performs no network
access, and writes all artifacts to a new isolated directory.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from scripts.academic_ppt.ooxml_qa import inspect_package
from scripts.academic_ppt.rendering import create_contact_sheet, render_pptx
from scripts.academic_ppt.visual_layout_qa import (
    inspect_powerpoint_text_layout,
    inspect_text_geometry,
)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _write_json(path: Path, payload: object) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--build-spec", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--node", type=Path, required=True)
    parser.add_argument("--node-modules", type=Path, required=True)
    parser.add_argument("--repo-root", type=Path, default=Path.cwd())
    args = parser.parse_args()

    build_spec = args.build_spec.resolve()
    output_dir = args.output_dir.resolve()
    repo_root = args.repo_root.resolve()
    if not build_spec.is_file():
        raise SystemExit(f"Missing build specification: {build_spec}")
    if output_dir.exists():
        raise SystemExit(f"Refusing to overwrite benchmark output: {output_dir}")
    output_dir.mkdir(parents=True)

    spec = json.loads(build_spec.read_text(encoding="utf-8-sig"))
    project = str(spec.get("brief", {}).get("project_name", "arm_b_benchmark"))
    pptx = output_dir / f"{project}_arm_b.pptx"
    pdf = output_dir / f"{project}_arm_b.pdf"
    preview_dir = output_dir / "preview"
    audit_dir = output_dir / "audit"
    audit_dir.mkdir()

    env = os.environ.copy()
    env["NODE_PATH"] = str(args.node_modules.resolve())
    command = [
        str(args.node.resolve()),
        str((repo_root / "scripts/v2/generate_upstream_behavior_arm.mjs").resolve()),
        str(build_spec),
        str(pptx),
        str(audit_dir),
    ]
    started = time.perf_counter()
    generated = subprocess.run(
        command,
        cwd=str(repo_root),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
        check=False,
    )
    generation_seconds = time.perf_counter() - started
    (audit_dir / "generation_stdout.txt").write_text(generated.stdout, encoding="utf-8")
    (audit_dir / "generation_stderr.txt").write_text(generated.stderr, encoding="utf-8")
    if generated.returncode != 0 or not pptx.is_file() or pptx.stat().st_size == 0:
        raise SystemExit(f"Arm B generation failed with exit code {generated.returncode}")

    render_started = time.perf_counter()
    renderer, previews, render_log = render_pptx(
        pptx,
        pdf,
        preview_dir,
        repo_root / "scripts",
        timeout=120,
    )
    render_seconds = time.perf_counter() - render_started
    create_contact_sheet(previews, preview_dir / "contact_sheet.png")
    (audit_dir / "render_log.txt").write_text(render_log, encoding="utf-8")

    source_ids = {
        str(source_id)
        for slide in spec.get("slides", [])
        for source_id in slide.get("source_ids", [])
    }
    ooxml = inspect_package(pptx, known_source_ids=source_ids)
    _write_json(output_dir / "ooxml_qa.json", ooxml)
    layout_report, layout_log = inspect_powerpoint_text_layout(
        pptx_path=pptx,
        output_json=output_dir / "powerpoint_text_layout.json",
        script_path=repo_root / "scripts/inspect_pptx_layout.ps1",
        timeout_seconds=120,
    )
    geometry_issues = inspect_text_geometry(layout_report)
    (audit_dir / "layout_log.txt").write_text(layout_log, encoding="utf-8")

    summary = {
        "arm": "B",
        "status": (
            "PASS"
            if ooxml.get("valid") and not geometry_issues and len(previews) == len(spec.get("slides", []))
            else "FAIL"
        ),
        "scope": "ISOLATED_HIGH_LEVEL_BEHAVIOR_BENCHMARK",
        "upstream_code_executed": False,
        "network_used": False,
        "third_party_api_used": False,
        "build_spec": str(build_spec),
        "build_spec_sha256": _sha256(build_spec),
        "pptx_sha256": _sha256(pptx),
        "pdf_sha256": _sha256(pdf),
        "slide_count_spec": len(spec.get("slides", [])),
        "slide_count_png": len(previews),
        "renderer": renderer,
        "ooxml_valid": bool(ooxml.get("valid")),
        "ooxml_error_count": len(ooxml.get("errors", [])),
        "ooxml_warning_count": len(ooxml.get("warnings", [])),
        "geometry_issues": geometry_issues,
        "generation_seconds": round(generation_seconds, 3),
        "render_seconds": round(render_seconds, 3),
        "total_seconds": round(time.perf_counter() - started, 3),
    }
    _write_json(output_dir / "benchmark_run_summary.json", summary)
    print(json.dumps(summary, ensure_ascii=False))
    return 0 if summary["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
