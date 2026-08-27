from __future__ import annotations

import csv
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
AUDIT = ROOT / "audit" / "v2_2"
SUITE = ROOT / "input" / "academic_ppt_regression_suite_v2_1"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def command(*args: str) -> str:
    return subprocess.run(
        list(args),
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    ).stdout.strip()


def write_csv(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    AUDIT.mkdir(parents=True, exist_ok=True)
    include_roots = [
        ROOT / "config",
        ROOT / "scripts",
        ROOT / "tests",
        ROOT / "regression",
    ]
    explicit = [
        ROOT / "run_ppt_workflow.py",
        ROOT / "AGENTS.md",
        ROOT / "package.json",
        ROOT / "requirements-lock.txt",
        ROOT / "audit" / "v2_1" / "final_status.md",
        ROOT / "audit" / "v2_1" / "regression_test_results.md",
        ROOT / "audit" / "v2_1" / "scientific_regression_report.md",
        ROOT / "audit" / "v2_1" / "cross_renderer_comparison.md",
        ROOT / "audit" / "v2_1" / "input_integrity_report.md",
    ]
    paths = []
    for root in include_roots:
        if root.exists():
            paths.extend(path for path in root.rglob("*") if path.is_file())
    paths.extend(path for path in explicit if path.is_file())
    unique = sorted(set(paths))
    rows = [
        {
            "relative_path": path.relative_to(ROOT).as_posix(),
            "size_bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }
        for path in unique
    ]
    write_csv(
        AUDIT / "baseline_file_hashes.csv",
        ["relative_path", "size_bytes", "sha256"],
        rows,
    )

    non_ground_truth = []
    for path in sorted(SUITE.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(SUITE).as_posix()
        if "/expected_ground_truth/" in f"/{relative}":
            continue
        non_ground_truth.append(
            {
                "relative_path": relative,
                "size_bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )
    write_csv(
        AUDIT / "suite_input_hashes_before.csv",
        ["relative_path", "size_bytes", "sha256"],
        non_ground_truth,
    )

    final_output = (
        ROOT
        / "output"
        / "SYN_CARDIO_AKI_SHADOW_VALIDATION"
        / "20260730_122932_v21_final"
    )
    key_output = {}
    for name in (
        "SYN_CARDIO_AKI_SHADOW_VALIDATION.pptx",
        "SYN_CARDIO_AKI_SHADOW_VALIDATION.pdf",
        "deck_ir.json",
        "claim_source_map.csv",
        "powerpoint_layout_manifest.json",
    ):
        path = final_output / name
        if path.is_file():
            key_output[name] = {
                "path": str(path),
                "size_bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
    manifest = {
        "schema_version": "2.2",
        "created_at": datetime.now(timezone.utc).astimezone().isoformat(),
        "branch_before_v2_2": "workflow/academic-ppt-v2-1-layout-qa",
        "branch": command("git", "branch", "--show-current"),
        "head": command("git", "rev-parse", "HEAD"),
        "tags": command("git", "tag", "--list").splitlines(),
        "working_tree_dirty": bool(command("git", "status", "--short")),
        "working_tree_snapshot_file_count": len(rows),
        "dependency_locks": {
            path.name: {
                "path": str(path),
                "sha256": sha256_file(path),
            }
            for path in (ROOT / "package.json", ROOT / "requirements-lock.txt")
            if path.is_file()
        },
        "v2_1_regression": "77/77 PASS",
        "v2_1_synthetic_deck": {
            "slides": 15,
            "powerpoint_pages": 15,
            "libreoffice_pages": 15,
            "status": "WORKFLOW_LAYOUT_QA_REMEDIATION_PASSED",
            "files": key_output,
        },
        "v2_overall_status": "PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING",
        "suite": {
            "path": str(SUITE),
            "non_ground_truth_files_hashed": len(non_ground_truth),
            "ground_truth_content_read_before_generation": False,
        },
        "network_access": False,
        "external_services": False,
        "global_installation": False,
    }
    (AUDIT / "baseline_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    runtime = {
        "python": sys.version,
        "platform": platform.platform(),
        "node": command(os.environ.get("PPT_NODE", "node"), "--version"),
        "powerpoint_com": "validated in v2.1 and re-probed during v2.2",
        "libreoffice": command(
            os.environ.get("PPT_SOFFICE", "soffice.com"), "--version"
        ),
        "graphviz": command(os.environ.get("PPT_GRAPHVIZ", "dot"), "-V"),
        "dependency_lock_hashes": manifest["dependency_locks"],
    }
    (AUDIT / "baseline_runtime_manifest.json").write_text(
        json.dumps(runtime, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (AUDIT / "baseline_test_results.md").write_text(
        """# v2.1 Frozen Baseline

- Workflow status: `WORKFLOW_LAYOUT_QA_REMEDIATION_PASSED`
- Existing regression suite: `77/77 PASS`
- Synthetic deck: 15 PPTX slides, 15 PowerPoint PDF pages, 15 PNG pages
- LibreOffice compatibility: 15/15 pages
- Scientific regression: PASS
- Input integrity: PASS
- Overall v2 status remains `PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING`

The Git HEAD still points to the v1 baseline because v2/v2.1 were intentionally
left uncommitted. Therefore this baseline is frozen by the complete
`baseline_file_hashes.csv` worktree snapshot, not by a misleading tag.
""",
        encoding="utf-8",
    )
    (AUDIT / "working_tree_status.txt").write_text(
        command("git", "status", "--short", "--untracked-files=all") + "\n",
        encoding="utf-8",
    )
    print(f"BASELINE_FILES={len(rows)}")
    print(f"SUITE_NON_GROUND_TRUTH_FILES={len(non_ground_truth)}")


if __name__ == "__main__":
    main()
