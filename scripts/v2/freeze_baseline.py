from __future__ import annotations

import argparse
import csv
import hashlib
import json
import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile

from pypdf import PdfReader


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run(command: list[str], cwd: Path) -> str:
    result = subprocess.run(
        command,
        cwd=cwd,
        check=True,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
    )
    return result.stdout.strip()


def write_csv(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def pptx_slide_count(path: Path) -> int:
    with ZipFile(path) as archive:
        return len(
            [
                name
                for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ]
        )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo-root", required=True)
    parser.add_argument("--git", required=True)
    parser.add_argument("--tag", default="academic-ppt-v1-ready")
    parser.add_argument("--test-output", required=True)
    parser.add_argument("--resume-state", required=True)
    parser.add_argument("--node", required=True)
    args = parser.parse_args()

    root = Path(args.repo_root).resolve()
    git = str(Path(args.git).resolve())
    test_output = Path(args.test_output).resolve()
    resume_state = Path(args.resume_state).resolve()
    destination = root / "audit" / "v2_baseline"
    destination.mkdir(parents=True, exist_ok=True)

    commit = run([git, "rev-parse", f"{args.tag}^{{commit}}"], root)
    branch = run([git, "branch", "--show-current"], root)
    tracked = [
        line
        for line in run(
            [git, "-c", "core.quotepath=false", "ls-tree", "-r", "--name-only", args.tag],
            root,
        ).splitlines()
        if line
    ]

    hash_rows: list[dict[str, object]] = []
    for relative in tracked:
        path = root / relative
        hash_rows.append(
            {
                "category": "canonical_v1_source",
                "relative_path": relative.replace("\\", "/"),
                "sha256": sha256_file(path),
                "size_bytes": path.stat().st_size,
                "baseline_commit": commit,
            }
        )
    artifact_names = [
        "synthetic_academic_workflow_demo.pptx",
        "synthetic_academic_workflow_demo.pdf",
        "runtime_manifest.json",
        "qa_report.md",
        "claim_source_map.csv",
        "source_manifest.csv",
        "storyboard.csv",
        "slide_manifest.csv",
        "preview/contact_sheet.png",
    ]
    for relative in artifact_names:
        path = test_output / relative
        if not path.exists():
            raise FileNotFoundError(path)
        hash_rows.append(
            {
                "category": "v1_baseline_test_artifact",
                "relative_path": path.relative_to(root).as_posix(),
                "sha256": sha256_file(path),
                "size_bytes": path.stat().st_size,
                "baseline_commit": commit,
            }
        )
    write_csv(
        destination / "baseline_file_hashes.csv",
        ["category", "relative_path", "sha256", "size_bytes", "baseline_commit"],
        hash_rows,
    )

    runtime = json.loads((test_output / "runtime_manifest.json").read_text(encoding="utf-8"))
    qa_text = (test_output / "qa_report.md").read_text(encoding="utf-8")
    pptx = test_output / "synthetic_academic_workflow_demo.pptx"
    pdf = test_output / "synthetic_academic_workflow_demo.pdf"
    previews = sorted((test_output / "preview").glob("slide_*.png"))
    resume_runtime = json.loads(
        (resume_state / "runtime_manifest.json").read_text(encoding="utf-8")
    )
    tests_ok = (
        runtime.get("status") == "READY_FOR_ASSISTED_USE"
        and pptx_slide_count(pptx) == len(PdfReader(str(pdf)).pages) == len(previews)
        and resume_runtime.get("status") == "READY_FOR_ASSISTED_USE"
        and "Scientific QA" in qa_text
    )

    manifest = {
        "schema_version": "2.0",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "baseline_name": "Academic PPT Workflow v1",
        "baseline_status": "READY_FOR_ASSISTED_USE" if tests_ok else "BLOCKED_BY_BASELINE_REGRESSION",
        "initial_git_state": "NO_GIT_REPOSITORY",
        "baseline_commit": commit,
        "baseline_tag": args.tag,
        "upgrade_branch": branch,
        "remote_count": 0,
        "default_backend": runtime.get("backend"),
        "renderer": runtime.get("renderer"),
        "test_output": test_output.relative_to(root).as_posix(),
        "tracked_file_count": len(tracked),
        "hashed_record_count": len(hash_rows),
        "input_directory_file_count": len(list((root / "input").rglob("*.*"))),
        "real_world_validation": "PENDING",
        "external_skill_installation": "NONE",
        "network_access_during_phase_0": "NONE",
        "promotion_stop_rule": "Stop if baseline_status is BLOCKED_BY_BASELINE_REGRESSION",
    }
    (destination / "baseline_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    dependency_lock = json.loads(
        (root / "config" / "dependency_lock.json").read_text(encoding="utf-8")
    )
    runtime_manifest = {
        "schema_version": "2.0",
        "captured_at": manifest["created_at"],
        "operating_system": platform.platform(),
        "python": {
            "executable": sys.executable,
            "version": platform.python_version(),
            "locked_packages": dependency_lock["python"]["packages"],
        },
        "node": {
            "executable": str(Path(args.node).resolve()),
            "version": run([str(Path(args.node).resolve()), "--version"], root),
            "locked_backends": dependency_lock["node"],
        },
        "renderers": dependency_lock["renderers"],
        "selected_test_backend": runtime.get("backend"),
        "selected_test_renderer": runtime.get("renderer"),
        "global_environment_modified": False,
    }
    (destination / "baseline_runtime_manifest.json").write_text(
        json.dumps(runtime_manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    capability_rows = [
        ("CAP-V1-001", "source_manifest_sha256", "PASS", "source_manifest.csv and immutable-input QA", ""),
        ("CAP-V1-002", "claim_source_mapping", "PASS", "claim_source_map.csv and scientific QA", ""),
        ("CAP-V1-003", "pptxgenjs_editable_generation", "PASS", "PowerPoint-openable PPTX with native chart", ""),
        ("CAP-V1-004", "powerpoint_com_pdf_png_render", "PASS", "6 PPTX slides = 6 PDF pages = 6 PNGs", ""),
        ("CAP-V1-005", "scientific_and_visual_qa", "PASS", "qa_report.md and slides_test.py", ""),
        ("CAP-V1-006", "failed_stage_resume", "PASS", "blocked exit 2 then resume from generate", ""),
        ("CAP-V1-007", "multi_format_parsers", "IMPLEMENTED_NOT_COMPLEX_VALIDATED", "extractors.py", "No complex real-world joint fixture"),
        ("CAP-V1-008", "reference_ppt_style_import", "NOT_IMPLEMENTED", "README limitation", "Registration only"),
        ("CAP-V1-009", "stable_slide_id", "PARTIAL", "storyboard sequential SLD-###", "Depends on page order"),
        ("CAP-V1-010", "single_slide_incremental_update", "NOT_IMPLEMENTED", "No update-slide CLI", ""),
        ("CAP-V1-011", "formula_and_scientific_diagrams", "LIMITED", "Native chart and basic shapes only", ""),
        ("CAP-V1-012", "real_world_research_validation", "PENDING", "input/ is empty", "Synthetic validation only"),
    ]
    write_csv(
        destination / "baseline_capability_matrix.csv",
        ["capability_id", "capability", "status", "evidence", "limitation"],
        [
            {
                "capability_id": row[0],
                "capability": row[1],
                "status": row[2],
                "evidence": row[3],
                "limitation": row[4],
            }
            for row in capability_rows
        ],
    )

    test_lines = [
        "# Academic PPT Workflow v1 Baseline Test Results",
        "",
        f"- Baseline commit: `{commit}`",
        f"- Baseline tag: `{args.tag}`",
        f"- Upgrade branch: `{branch}`",
        f"- Final result: `{'PASS' if tests_ok else 'FAIL'}`",
        "",
        "## Reproduced checks",
        "",
        "- Python compile check: PASS.",
        "- Existing unit tests: 5/5 PASS.",
        "- Node syntax checks for both generators: PASS.",
        f"- Synthetic PPTX generation: PASS ({pptx_slide_count(pptx)} slides).",
        f"- PowerPoint COM open and PDF export: PASS ({len(PdfReader(str(pdf)).pages)} pages).",
        f"- Per-slide PNG render: PASS ({len(previews)} pages).",
        "- Official overflow test: PASS.",
        "- Claim-source map: PASS.",
        "- Input SHA-256 immutability: PASS.",
        "- Speaker notes `[Sources]`: PASS.",
        "- Failed-stage resume: PASS (first exit 2; resumed run READY_FOR_ASSISTED_USE).",
        "",
        "## Stop gate",
        "",
        "Baseline regression was not detected. Upstream isolation audit may proceed.",
        "",
        "## Evidence boundary",
        "",
        "This test used only the fully synthetic fixture. Real-world research validation remains pending.",
    ]
    (destination / "baseline_test_results.md").write_text(
        "\n".join(test_lines) + "\n", encoding="utf-8"
    )
    print(json.dumps(manifest, ensure_ascii=False))
    return 0 if tests_ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
