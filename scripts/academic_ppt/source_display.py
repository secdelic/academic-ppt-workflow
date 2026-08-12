from __future__ import annotations

import re
from collections.abc import Mapping
from typing import Any


def _compile_rules(policy: Mapping[str, Any]) -> list[tuple[re.Pattern[str], str]]:
    compiled: list[tuple[re.Pattern[str], str]] = []
    for rule in policy.get("short_label_rules", []):
        pattern = str(rule.get("path_pattern", "")).strip()
        label = str(rule.get("label", "")).strip()
        if pattern and label:
            compiled.append((re.compile(pattern), label))
    return compiled


def short_source_label(
    source_ids: list[str],
    manifest: list[dict[str, str]],
    policy: Mapping[str, Any],
    *,
    audit_presentation: bool = False,
) -> str:
    if not source_ids:
        return "来源：INFORMATION_REQUIRED"
    by_id = {row["source_id"]: row for row in manifest}
    if audit_presentation:
        maximum = int(policy.get("level_1", {}).get("maximum_source_labels", 2))
        shown = source_ids[:maximum]
        suffix = f" 等 {len(source_ids)} 项" if len(source_ids) > maximum else ""
        return "来源：" + "；".join(shown) + suffix

    labels: list[str] = []
    rules = _compile_rules(policy)
    fallback = str(policy.get("fallback_label", "输入材料"))
    for source_id in source_ids:
        row = by_id.get(source_id, {})
        searchable = " ".join(
            [
                str(row.get("relative_path", "")),
                str(row.get("file_name", "")),
                str(row.get("file_type", "")),
            ]
        )
        label = next(
            (candidate for pattern, candidate in rules if pattern.search(searchable)),
            fallback,
        )
        if label not in labels:
            labels.append(label)
    maximum = int(policy.get("level_1", {}).get("maximum_source_labels", 2))
    displayed = labels[:maximum]
    suffix = f" 等 {len(labels)} 类输入" if len(labels) > maximum else ""
    text = "来源：" + "；".join(displayed) + suffix
    limit = int(policy.get("level_1", {}).get("maximum_characters", 72))
    return text if len(text) <= limit else text[: max(1, limit - 1)] + "…"


__all__ = ["short_source_label"]
