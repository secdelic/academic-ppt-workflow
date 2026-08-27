from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path
from typing import Any


def rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write(path: Path, data: list[dict[str, Any]], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(data)


def number_tokens(value: str) -> list[float]:
    return [float(token) for token in re.findall(r"(?<![A-Za-z])[-+]?\d+(?:\.\d+)?(?:e[-+]?\d+)?", value, re.I)]


def visual_numbers(value: Any) -> list[float]:
    result: list[float] = []
    if isinstance(value, dict):
        for item in value.values():
            result.extend(visual_numbers(item))
    elif isinstance(value, list):
        for item in value:
            result.extend(visual_numbers(item))
    elif isinstance(value, (int, float)):
        result.append(float(value))
    return result


def numeric_match(expected: str, observed: list[float]) -> bool:
    tokens = number_tokens(expected)
    if not tokens:
        return False
    for token in tokens:
        alternatives = {token}
        if "%" in expected or "percentage point" in expected.lower():
            alternatives.add(token / 100)
        if not any(any(abs(value - alt) <= max(1e-8, abs(alt) * 1e-6) for alt in alternatives) for value in observed):
            return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--suite", type=Path, required=True)
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--staging-root", type=Path, required=True)
    parser.add_argument("--audit-root", type=Path, required=True)
    args = parser.parse_args()
    if not (args.audit_root / "ground_truth_access_log.json").is_file():
        raise RuntimeError("Ground truth access stage was not recorded")

    claim_results = []
    conflict_results = []
    unresolved_results = []
    for project in sorted(path for path in args.suite.iterdir() if path.is_dir()):
        gt = project / "expected_ground_truth"
        if not gt.is_dir():
            continue
        visual_path = args.run_root / project.name / "arm_D" / "visual_ir.json"
        claim_map_path = args.run_root / project.name / "arm_D" / "claim_source_map.csv"
        manifest_path = args.staging_root / project.name / "source_manifest.csv"
        visuals = json.loads(visual_path.read_text(encoding="utf-8"))
        claim_map = rows(claim_map_path)
        manifest = rows(manifest_path)
        observed_numbers = visual_numbers(visuals)
        observed_text = json.dumps(visuals, ensure_ascii=False).lower()
        claim_blob = json.dumps(claim_map, ensure_ascii=False).lower()
        sources_by_name = {
            Path(row["relative_path"]).name: row["source_id"] for row in manifest
        }
        visual_sources = {sid for visual in visuals for sid in visual["source_ids"]}

        for expected in rows(gt / "expected_claims.csv"):
            canonical_name = Path(expected["canonical_source"]).name
            sid = sources_by_name.get(canonical_name)
            source_present = bool(sid and sid in visual_sources)
            exact_id = expected["claim_id"].lower() in claim_blob
            value = expected["expected_value"]
            value_match = numeric_match(value, observed_numbers)
            if not number_tokens(value):
                terms = [
                    token.lower()
                    for token in re.findall(r"[A-Za-z][A-Za-z_-]{2,}", value)
                    if token.lower() not in {"from", "table", "canonical", "none"}
                ]
                value_match = bool(terms and all(term in observed_text for term in terms))
            if source_present and value_match and exact_id:
                classification = "true_positive"
            elif source_present and value_match:
                classification = "partially_correct"
            else:
                classification = "false_negative"
            claim_results.append({
                "project_key": project.name,
                "claim_id": expected["claim_id"],
                "canonical_source": canonical_name,
                "source_present_in_visual_ir": source_present,
                "expected_value_represented": value_match,
                "canonical_claim_id_in_claim_map": exact_id,
                "classification": classification,
                "note": (
                    "Value and source are represented, but canonical claim identity is absent."
                    if classification == "partially_correct"
                    else "Formal claim recall is incomplete." if classification == "false_negative"
                    else "Formal claim, value, and source are all represented."
                ),
            })
        for expected in rows(gt / "expected_conflicts.csv"):
            marker = expected["conflicting_value"].lower()
            found = marker in claim_blob or marker in observed_text
            conflict_results.append({
                "project_key": project.name,
                "conflict_id": expected["conflict_id"],
                "detected": False,
                "classification": "false_negative",
                "note": "The visual benchmark runner did not perform DOCX/PDF conflict extraction.",
            })
        for expected in rows(gt / "expected_unresolved_items.csv"):
            marker = expected["marker"]
            found = marker.lower() in claim_blob
            unresolved_results.append({
                "project_key": project.name,
                "item_id": expected["item_id"],
                "marker": marker,
                "preserved": found,
                "classification": "true_positive" if found else "false_negative",
                "note": (
                    "Marker remains in generated audit."
                    if found else
                    "The visual benchmark runner did not propagate unresolved markers."
                ),
            })

    write(args.audit_root / "claim_recall_report.csv", claim_results, [
        "project_key", "claim_id", "canonical_source",
        "source_present_in_visual_ir", "expected_value_represented",
        "canonical_claim_id_in_claim_map", "classification", "note",
    ])
    write(args.audit_root / "conflict_detection_report.csv", conflict_results, [
        "project_key", "conflict_id", "detected", "classification", "note",
    ])
    write(args.audit_root / "unresolved_item_report.csv", unresolved_results, [
        "project_key", "item_id", "marker", "preserved", "classification", "note",
    ])
    summary = {
        "claims_total": len(claim_results),
        "claims_true_positive": sum(r["classification"] == "true_positive" for r in claim_results),
        "claims_partially_correct": sum(r["classification"] == "partially_correct" for r in claim_results),
        "claims_false_negative": sum(r["classification"] == "false_negative" for r in claim_results),
        "conflicts_detected": sum(bool(r["detected"]) for r in conflict_results),
        "conflicts_total": len(conflict_results),
        "unresolved_preserved": sum(bool(r["preserved"]) for r in unresolved_results),
        "unresolved_total": len(unresolved_results),
    }
    (args.audit_root / "ground_truth_detailed_summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
