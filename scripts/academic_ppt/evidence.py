from __future__ import annotations

import re
from pathlib import Path

from .utils import write_csv


EVIDENCE_FIELDS = [
    "claim_id",
    "claim_text",
    "source_id",
    "source_location",
    "evidence_type",
    "confidence",
    "allowed_wording",
    "prohibited_overstatement",
]

CONTENT_SECTIONS = {
    "研究背景": ["background", "背景"],
    "临床或科学问题": ["clinical question", "scientific question", "研究问题"],
    "知识空白": ["knowledge gap", "gap", "知识空白"],
    "研究目的": ["objective", "aim", "目的"],
    "假设": ["hypothesis", "假设"],
    "研究设计": ["study design", "design", "研究设计"],
    "数据来源": ["data source", "dataset", "数据来源"],
    "纳入排除标准": ["inclusion", "exclusion", "纳入", "排除"],
    "暴露或干预": ["exposure", "intervention", "暴露", "干预"],
    "结局": ["outcome", "endpoint", "结局"],
    "统计方法": ["statistical", "analysis", "统计"],
    "主要结果": ["primary result", "main result", "主要结果"],
    "次要结果": ["secondary result", "次要结果"],
    "敏感性分析": ["sensitivity", "敏感性"],
    "亚组分析": ["subgroup", "亚组"],
    "机制解释": ["mechanism", "机制"],
    "研究优势": ["strength", "优势"],
    "局限性": ["limitation", "局限"],
    "结论": ["conclusion", "结论"],
    "参考文献": ["reference", "doi", "参考文献"],
}


def _source_lines(extracted: dict[str, str]):
    for source_id, text in extracted.items():
        for number, line in enumerate(text.splitlines(), start=1):
            line = line.strip()
            if line:
                yield source_id, number, line


def build_evidence(
    extracted: dict[str, str],
    manifest: list[dict[str, str]],
    staging_root: Path,
    brief: dict,
) -> tuple[list[dict[str, str]], list[str]]:
    claims: list[dict[str, str]] = []
    content_hits: dict[str, list[str]] = {section: [] for section in CONTENT_SECTIONS}
    unresolved: list[str] = []

    for source_id, line_number, line in _source_lines(extracted):
        lower = line.lower()
        for section, keywords in CONTENT_SECTIONS.items():
            if any(keyword.lower() in lower for keyword in keywords):
                content_hits[section].append(f"{source_id}:L{line_number} - {line[:240]}")
        if line.startswith("CLAIM|"):
            parts = line.split("|")
            if len(parts) < 7:
                unresolved.append(
                    f"Malformed CLAIM record at {source_id}:L{line_number}; expected 7 fields after CLAIM"
                )
                continue
            claim_text, source_location, evidence_type, confidence, allowed, prohibited = parts[1:7]
            claims.append(
                {
                    "claim_id": f"CLM-{len(claims) + 1:04d}",
                    "claim_text": claim_text.strip(),
                    "source_id": source_id,
                    "source_location": source_location.strip() or f"line {line_number}",
                    "evidence_type": evidence_type.strip() or "reported_source_text",
                    "confidence": confidence.strip() or "low",
                    "allowed_wording": allowed.strip() or claim_text.strip(),
                    "prohibited_overstatement": prohibited.strip() or "Any wording beyond the source statement",
                }
            )

    inventory_lines = ["# Content Inventory", "", "All entries are source excerpts or locations; no missing content is inferred.", ""]
    for section in list(CONTENT_SECTIONS) + ["待确认事项"]:
        inventory_lines.extend([f"## {section}", ""])
        hits = content_hits.get(section, [])
        if hits:
            inventory_lines.extend(f"- {hit}" for hit in hits[:20])
        else:
            inventory_lines.append("- INFORMATION_REQUIRED")
        inventory_lines.append("")
    (staging_root / "content_inventory.md").write_text(
        "\n".join(inventory_lines), encoding="utf-8"
    )
    write_csv(staging_root / "evidence_inventory.csv", EVIDENCE_FIELDS, claims)

    figures = [
        {
            "figure_id": f"FIG-{i:04d}",
            "source_id": row["source_id"],
            "relative_path": row["relative_path"],
            "figure_type": row["file_type"],
            "editable": "no" if row["file_type"] in {"png", "jpg", "jpeg"} else "conditional",
            "manual_review_required": "yes",
        }
        for i, row in enumerate(
            [r for r in manifest if r["file_type"] in {"png", "jpg", "jpeg", "svg"}],
            start=1,
        )
    ]
    write_csv(
        staging_root / "figure_inventory.csv",
        ["figure_id", "source_id", "relative_path", "figure_type", "editable", "manual_review_required"],
        figures,
    )
    tables = [
        {
            "table_id": f"TAB-{i:04d}",
            "source_id": row["source_id"],
            "relative_path": row["relative_path"],
            "table_type": row["file_type"],
            "manual_review_required": "yes",
        }
        for i, row in enumerate(
            [r for r in manifest if r["file_type"] in {"csv", "tsv", "xlsx"}],
            start=1,
        )
    ]
    write_csv(
        staging_root / "table_inventory.csv",
        ["table_id", "source_id", "relative_path", "table_type", "manual_review_required"],
        tables,
    )

    for key, value in brief.items():
        if value == "INFORMATION_REQUIRED":
            unresolved.append(f"Brief field `{key}` is INFORMATION_REQUIRED")
    if not claims:
        unresolved.append("No source-bound claims were registered; result/conclusion slides cannot be assertive")
    for row in manifest:
        if row["parsed_successfully"] != "yes":
            unresolved.append(f"Source parsing failed: {row['relative_path']} - {row['parse_warning']}")
        elif row["parse_warning"]:
            unresolved.append(f"Source warning: {row['relative_path']} - {row['parse_warning']}")
        if row["possible_sensitive_information"] == "yes":
            unresolved.append(f"Possible sensitive information requires manual review: {row['relative_path']}")
    return claims, unresolved
