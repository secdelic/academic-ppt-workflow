from __future__ import annotations

import argparse
import csv
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    if not fields:
        fields = list(rows[0]) if rows else ["status"]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--suite", type=Path, required=True)
    parser.add_argument("--staging-root", type=Path, required=True)
    parser.add_argument("--audit-root", type=Path, required=True)
    args = parser.parse_args()

    generation = args.run_root / "generation_complete.json"
    initial_qa = args.run_root / "initial_qa_complete.json"
    if not generation.is_file() or not initial_qa.is_file():
        raise RuntimeError("Generation and initial QA must complete before ground-truth access")
    if json.loads(generation.read_text(encoding="utf-8"))["ground_truth_content_read"]:
        raise RuntimeError("Generation marker unexpectedly says ground truth was read")
    if json.loads(initial_qa.read_text(encoding="utf-8"))["ground_truth_content_read"]:
        raise RuntimeError("Initial QA marker unexpectedly says ground truth was read")

    access_time = datetime.now(timezone.utc).isoformat()
    access_log = {
        "stage": "POST_GENERATION_GROUND_TRUTH_EVALUATION",
        "first_content_read_at_utc": access_time,
        "generation_complete_marker": str(generation.resolve()),
        "initial_qa_complete_marker": str(initial_qa.resolve()),
    }
    args.audit_root.mkdir(parents=True, exist_ok=True)
    (args.audit_root / "ground_truth_access_log.json").write_text(
        json.dumps(access_log, indent=2), encoding="utf-8"
    )

    comparison_rows: list[dict[str, Any]] = []
    expected_summary: list[dict[str, Any]] = []
    for project in sorted(path for path in args.suite.iterdir() if path.is_dir()):
        gt_dir = project / "expected_ground_truth"
        if not gt_dir.is_dir():
            continue
        generated_claims = []
        for claim_path in sorted((args.run_root / project.name).glob("arm_*/claim_source_map.csv")):
            generated_claims.extend(read_csv(claim_path))
        generated_blob = "\n".join(
            f"{row.get('claim_id','')} {row.get('claim_text','')} {row.get('source_ids','')} {row.get('wording_boundary','')}"
            for row in generated_claims
        ).lower()
        expected_files = sorted(path for path in gt_dir.rglob("*") if path.is_file())
        for path in expected_files:
            rel = path.relative_to(project).as_posix()
            text = path.read_text(encoding="utf-8-sig", errors="replace")
            expected_summary.append({
                "project_key": project.name,
                "relative_path": rel,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "size_bytes": path.stat().st_size,
                "first_read_stage": "POST_GENERATION_GROUND_TRUTH_EVALUATION",
            })
            if path.suffix.lower() == ".csv":
                rows = read_csv(path)
                headers = list(rows[0]) if rows else []
                expected_summary[-1]["headers"] = "|".join(headers)
                for index, row in enumerate(rows, start=1):
                    tokens = [
                        str(value).strip().lower()
                        for key, value in row.items()
                        if value not in {None, ""} and any(
                            marker in key.lower()
                            for marker in ("claim", "visual", "source", "expected", "status", "rule")
                        )
                    ]
                    token_hits = [token for token in tokens if len(token) >= 4 and token in generated_blob]
                    comparison_rows.append({
                        "project_key": project.name,
                        "ground_truth_file": rel,
                        "row_number": index,
                        "expected_fields": "|".join(tokens),
                        "matched_tokens": "|".join(token_hits),
                        "classification": (
                            "true_positive" if token_hits else "not_assessed"
                        ),
                        "note": (
                            "Exact token present in generated claim/source audit."
                            if token_hits else
                            "Ground-truth row requires manual semantic review; no automated exact-token assertion."
                        ),
                    })
            else:
                expected_summary[-1]["headers"] = ""

    write_csv(args.audit_root / "ground_truth_file_manifest.csv", expected_summary)
    write_csv(args.audit_root / "ground_truth_comparison.csv", comparison_rows)
    (args.audit_root / "ground_truth_evaluation.md").write_text(
        "# Post-generation ground-truth evaluation\n\n"
        f"- First content read: `{access_time}`\n"
        f"- Stage: `POST_GENERATION_GROUND_TRUTH_EVALUATION`\n"
        f"- Expected files read: {len(expected_summary)}\n"
        f"- Automated exact-token rows: {sum(r['classification'] == 'true_positive' for r in comparison_rows)}\n"
        f"- Not assessed automatically: {sum(r['classification'] == 'not_assessed' for r in comparison_rows)}\n"
        "- No expected-ground-truth file entered source discovery, Deck IR, Visual IR, claims, generation, or initial QA.\n"
        "- `not_assessed` rows require human semantic comparison and are not counted as passes or failures.\n",
        encoding="utf-8",
    )
    print(f"GROUND_TRUTH_FILES={len(expected_summary)}")
    print(f"FIRST_READ={access_time}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
