#!/usr/bin/env python3
"""Fail closed when tests depend on developer-local or historical run state."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path
from typing import Iterable


SCHEMA_VERSION = "academic-ppt-ci-self-containment/1"
TEXT_SUFFIXES = {".py", ".md", ".json", ".yaml", ".yml", ".csv", ".ps1", ".mjs", ".cjs"}
CLASSIFICATIONS = {
    "TRACKED_FIXTURE",
    "GENERATED_TEMP",
    "DOCUMENTATION_ONLY",
    "FORBIDDEN_PRIVATE_DEPENDENCY",
}
PRIVATE_ABSOLUTE_PATH = re.compile(
    r"(?i)(?<![A-Za-z0-9])(?:[A-Z]:[\\/]"
    r"|\\\\[A-Za-z0-9][A-Za-z0-9.-]*[\\/][A-Za-z0-9$_.-]+(?:[\\/]|$))"
)
FORBIDDEN_RESOURCE = re.compile(
    r"(?i)(?:staging(?:[\"']|[\\/])"
    r"|audit(?:[\"']\s*)?[\\/]v\d"
    r"|benchmark(?:[\"']\s*)?[\\/]v\d"
    r"|output(?:[\"']\s*)?[\\/]"
    r"|\.cache(?:[\"']\s*)?[\\/]project_state"
    r"|config(?:[\"']\s*)?[\\/]local\.yaml)"
)
HISTORICAL_RUN_ID = re.compile(r"\b20\d{6}_\d{6}(?:_[A-Za-z0-9_-]+)?\b")
POLICY_MARKERS = ("self-containment: documentation-only", "check-ignore")
GENERATED_MARKERS = (
    "tmp_path", "temporary", "runner_temp", "workspace.", "root /", "root\\",
    "run_id", "synthetic", "generated-temp", "test_unused",
)


def _git_index_paths(repo_root: Path) -> set[str]:
    completed = subprocess.run(
        ["git", "ls-files", "--cached"],
        cwd=repo_root,
        text=True,
        capture_output=True,
        check=True,
    )
    return {line.strip().replace("\\", "/") for line in completed.stdout.splitlines() if line.strip()}


def _iter_text_files(repo_root: Path) -> Iterable[Path]:
    roots = (repo_root / "tests", repo_root / ".github" / "workflows", repo_root / "scripts" / "ci")
    seen: set[Path] = set()
    for root in roots:
        if not root.exists():
            continue
        for path in sorted(root.rglob("*")):
            if path.is_file() and path.suffix.lower() in TEXT_SUFFIXES and path not in seen:
                seen.add(path)
                yield path


def _classification(
    line: str,
    *,
    relative_path: str,
    private_absolute: bool,
    historical_run: bool,
) -> str:
    lowered = line.casefold()
    if relative_path.endswith(".md") or relative_path == "scripts/ci/verify_test_self_containment.py":
        return "DOCUMENTATION_ONLY"
    if any(marker in lowered for marker in POLICY_MARKERS):
        return "DOCUMENTATION_ONLY"
    if not private_absolute and any(marker in lowered for marker in GENERATED_MARKERS):
        return "GENERATED_TEMP"
    if historical_run and "run_id" in lowered:
        return "GENERATED_TEMP"
    return "FORBIDDEN_PRIVATE_DEPENDENCY"


def scan_repository(repo_root: Path) -> dict[str, object]:
    repo_root = repo_root.resolve()
    tracked = _git_index_paths(repo_root)
    findings: list[dict[str, object]] = []

    for path in _iter_text_files(repo_root):
        relative = path.relative_to(repo_root).as_posix()
        text = path.read_text(encoding="utf-8", errors="replace")
        for line_number, line in enumerate(text.splitlines(), start=1):
            private_match = PRIVATE_ABSOLUTE_PATH.search(line)
            resource_match = FORBIDDEN_RESOURCE.search(line)
            run_match = HISTORICAL_RUN_ID.search(line)
            if not (private_match or resource_match or run_match):
                continue
            classification = _classification(
                line,
                relative_path=relative,
                private_absolute=bool(private_match),
                historical_run=bool(run_match),
            )
            findings.append(
                {
                    "path": relative,
                    "line": line_number,
                    "kind": (
                        "PRIVATE_ABSOLUTE_PATH" if private_match else
                        "HISTORICAL_RUN_ID" if run_match else
                        "LOCAL_OR_HISTORICAL_RESOURCE"
                    ),
                    "classification": classification,
                    "detail": "redacted" if private_match else (resource_match or run_match).group(0),
                }
            )

    fixture_root = repo_root / "tests" / "fixtures" / "synthetic" / "ci_contracts"
    fixture_rows: list[dict[str, object]] = []
    if fixture_root.is_dir():
        for path in sorted(p for p in fixture_root.rglob("*") if p.is_file()):
            relative = path.relative_to(repo_root).as_posix()
            is_tracked = relative in tracked
            fixture_rows.append(
                {"path": relative, "classification": "TRACKED_FIXTURE" if is_tracked else "FORBIDDEN_PRIVATE_DEPENDENCY", "git_tracked": is_tracked}
            )
            if not is_tracked:
                findings.append(
                    {
                        "path": relative,
                        "line": 0,
                        "kind": "UNTRACKED_STATIC_FIXTURE",
                        "classification": "FORBIDDEN_PRIVATE_DEPENDENCY",
                        "detail": "static fixture is absent from the Git index",
                    }
                )

    invalid = [row for row in findings if row["classification"] not in CLASSIFICATIONS]
    if invalid:
        raise RuntimeError("scanner produced an invalid dependency classification")
    forbidden = [row for row in findings if row["classification"] == "FORBIDDEN_PRIVATE_DEPENDENCY"]
    counts = {name: sum(1 for row in findings if row["classification"] == name) for name in sorted(CLASSIFICATIONS)}
    return {
        "schema_version": SCHEMA_VERSION,
        "status": "PASS" if not forbidden else "FAIL",
        "scanned_roots": ["tests/", ".github/workflows/", "scripts/ci/"],
        "findings": findings,
        "static_fixtures": fixture_rows,
        "classification_counts": counts,
        "forbidden_dependencies": len(forbidden),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", default=str(Path(__file__).resolve().parents[2]))
    parser.add_argument("--output", default="ci_self_containment_report.json")
    args = parser.parse_args()
    report = scan_repository(Path(args.repo_root))
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": report["status"], "forbidden_dependencies": report["forbidden_dependencies"], "output": str(output)}))
    return 0 if report["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
