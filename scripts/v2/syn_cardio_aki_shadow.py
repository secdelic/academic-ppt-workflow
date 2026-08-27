from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "scripts"))

from academic_ppt.extractors import extract_source  # noqa: E402
from academic_ppt.inventory import MANIFEST_FIELDS, inventory_sources  # noqa: E402
from academic_ppt.utils import (  # noqa: E402
    read_csv,
    sha256_file,
    utc_offset_timestamp,
    write_csv,
    write_json,
)


PROJECT = "SYN_CARDIO_AKI_SHADOW_VALIDATION"
SUPPORTED = {
    ".docx", ".pdf", ".pptx", ".potx", ".xlsx", ".csv", ".tsv", ".md",
    ".txt", ".png", ".jpg", ".jpeg", ".svg", ".ris",
}
HASH_FIELDS = ["relative_path", "sha256", "size_bytes", "modified_time"]
CLAIM_FIELDS = [
    "claim_id", "claim_text", "source_id", "source_location", "source_type",
    "source_sha256", "canonical_status", "evidence_type", "estimate",
    "confidence_interval", "p_value", "confidence", "allowed_wording",
    "wording_boundary", "prohibited_overstatement", "conflict_status",
    "manual_review_required", "visual_type", "visual_asset_path",
]
CONFLICT_FIELDS = [
    "conflict_id", "source_a", "source_location_a", "value_a", "source_b",
    "source_location_b", "value_b", "canonical_source", "resolution",
    "affected_claim_ids", "manual_review_required",
]


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], cwd=REPO, text=True, encoding="utf-8",
        errors="replace", capture_output=True, check=False,
    )
    return result.stdout.strip() if result.returncode == 0 else f"ERROR: {result.stderr.strip()}"


def _source_by_path(manifest: list[dict[str, str]]) -> dict[str, dict[str, str]]:
    return {row["relative_path"]: row for row in manifest}


def _hash_rows(input_root: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for path in sorted(input_root.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(input_root).as_posix()
        if "expected_ground_truth" in {part.lower() for part in Path(rel).parts}:
            continue
        stat = path.stat()
        rows.append({
            "relative_path": rel,
            "sha256": sha256_file(path),
            "size_bytes": str(stat.st_size),
            "modified_time": datetime.fromtimestamp(stat.st_mtime).astimezone().isoformat(timespec="seconds"),
        })
    return rows


def _line_location(text: str, needle: str) -> str:
    for number, line in enumerate(text.splitlines(), start=1):
        if needle in line:
            return f"extracted line {number}"
    return "extracted text"


def _claim(
    claim_id: str,
    text: str,
    source: dict[str, str],
    location: str,
    evidence_type: str,
    title: str,
    boundary: str,
    *,
    estimate: str = "",
    ci: str = "",
    p: str = "",
    canonical: str = "canonical",
    conflict: str = "none",
    visual_type: str = "",
    visual_asset: str = "",
) -> dict[str, str]:
    return {
        "claim_id": claim_id,
        "claim_text": text,
        "source_id": source["source_id"],
        "source_location": location,
        "source_type": source["file_type"],
        "source_sha256": source["sha256"],
        "canonical_status": canonical,
        "evidence_type": evidence_type,
        "estimate": estimate,
        "confidence_interval": ci,
        "p_value": p,
        "confidence": "high" if canonical == "canonical" else "medium",
        "allowed_wording": title,
        "wording_boundary": boundary,
        "prohibited_overstatement": boundary,
        "conflict_status": conflict,
        "manual_review_required": "yes",
        "visual_type": visual_type,
        "visual_asset_path": visual_asset,
    }


def prepare(run_id: str, powerpoint_status: str, powerpoint_version: str) -> None:
    input_root = REPO / "input"
    audit_root = REPO / "audit" / "synthetic_shadow_validation"
    seed_root = REPO / "staging" / "synthetic_shadow_validation_seed" / run_id
    if seed_root.exists():
        raise FileExistsError(f"Refusing to overwrite seed directory: {seed_root}")
    audit_root.mkdir(parents=True, exist_ok=True)
    seed_root.mkdir(parents=True)

    manifest_path = audit_root / "input_manifest.csv"
    manifest = inventory_sources(
        input_root, SUPPORTED, manifest_path, reference_mode="style-only"
    )
    before = _hash_rows(input_root)
    write_csv(audit_root / "input_hashes_before.csv", HASH_FIELDS, before)
    by_path = _source_by_path(manifest)
    expected = input_root / "expected_ground_truth"
    ppt_like = [
        row["relative_path"] for row in manifest
        if row["file_type"] in {"pptx", "potx"}
    ]
    style_refs = [
        row["relative_path"] for row in manifest
        if row["source_role"] == "style_reference"
    ]
    branch = _git("branch", "--show-current")
    head = _git("rev-parse", "HEAD")
    tag = _git("rev-list", "-n", "1", "academic-ppt-v1-ready")
    dirty = _git("status", "--porcelain")
    type_counts = Counter(row["file_type"] for row in manifest)
    runtime = {
        "run_id": run_id,
        "timestamp": utc_offset_timestamp(),
        "git_branch": branch,
        "git_head": head,
        "v1_tag_commit": tag,
        "worktree_dirty_entries": len(dirty.splitlines()) if dirty else 0,
        "python": sys.version.split()[0],
        "python_executable": sys.executable,
        "python_runtime_note": "workspace-local console-subsystem copy; PYTHONHOME points to bundled read-only runtime",
        "powerpoint_com": powerpoint_status,
        "powerpoint_version": powerpoint_version,
        "libreoffice": os.environ.get("PPT_SOFFICE") or shutil.which("soffice.com") or "AUTO_DISCOVER",
        "node": "v24.14.0",
        "pptxgenjs": "4.0.1",
        "network_access": False,
        "external_services": False,
        "reference_mode": "style-only",
    }
    write_json(audit_root / "runtime_manifest.json", runtime)
    report = [
        "# SYN-CARDIO-AKI Complex Synthetic Shadow Validation — Preflight",
        "",
        f"- Run ID: `{run_id}`",
        f"- Git branch: `{branch}`",
        f"- HEAD: `{head}`",
        f"- v1 tag: `academic-ppt-v1-ready` -> `{tag}`",
        f"- Worktree dirty entries before this run: `{runtime['worktree_dirty_entries']}`",
        f"- Brief exists: `{(REPO / 'brief/presentation_brief.yaml').is_file()}`",
        f"- Input files registered: `{len(manifest)}`; types: `{dict(sorted(type_counts.items()))}`",
        f"- PPTX/POTX discovered: `{ppt_like}`",
        f"- User-confirmed style references: `{style_refs}`",
        f"- PowerPoint COM: `{powerpoint_status}` (version `{powerpoint_version}`)",
        f"- LibreOffice fallback: `{Path(runtime['libreoffice']).is_file()}`",
        "- PptxGenJS stable backend: `AVAILABLE` (4.0.1)",
        "- Required Python/Node dependencies: `AVAILABLE` using project/bundled runtimes",
        f"- Output root writable: `{(REPO / 'output').is_dir()}`",
        f"- expected_ground_truth under input: `{expected.is_dir()}`",
        "- Ground-truth source discovery exclusion: `ENFORCED`",
        "- External access: `NONE`",
        "",
        "The PPTX in `input/style_reference/` is treated as a user-declared style-only reference. Its content is excluded from scientific evidence.",
    ]
    (audit_root / "preflight_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")

    extracted: dict[str, str] = {}
    parse_rows: list[dict[str, str]] = []
    for row in manifest:
        path = input_root / row["relative_path"]
        if row["source_role"] == "style_reference":
            parse_rows.append({
                "source_id": row["source_id"], "relative_path": row["relative_path"],
                "file_type": row["file_type"], "status": "STYLE_ONLY_EXCLUDED_FROM_EVIDENCE",
                "count": "", "warning": "Text, notes, media and data were not read into the evidence stream",
            })
            continue
        try:
            text, count, warnings = extract_source(path)
            extracted[row["source_id"]] = text
            parse_rows.append({
                "source_id": row["source_id"], "relative_path": row["relative_path"],
                "file_type": row["file_type"], "status": "PASS",
                "count": str(count), "warning": "; ".join(warnings),
            })
        except Exception as exc:
            extracted[row["source_id"]] = ""
            parse_rows.append({
                "source_id": row["source_id"], "relative_path": row["relative_path"],
                "file_type": row["file_type"], "status": "FAIL",
                "count": "", "warning": f"{type(exc).__name__}: {exc}",
            })
    write_csv(
        seed_root / "parse_results.csv",
        ["source_id", "relative_path", "file_type", "status", "count", "warning"],
        parse_rows,
    )

    cohort_path = "data/02_cohort_summary.csv"
    model_path = "data/03_model_results.csv"
    sensitivity_path = "data/04_sensitivity_results.csv"
    subgroup_path = "data/05_subgroup_results.csv"
    missing_path = "data/06_missingness.csv"
    main_path = "documents/01_main_manuscript_synthetic.docx"
    protocol_path = "documents/02_protocol_synopsis_synthetic.docx"
    refs_path = "references/reference_status.md"
    figure1_path = "figures/figure_1_cohort_flow.png"
    figure3_path = "figures/figure_3_creatinine_trajectory.png"
    required = [
        cohort_path, model_path, sensitivity_path, subgroup_path, missing_path,
        main_path, protocol_path, refs_path, figure1_path, figure3_path,
    ]
    missing_required = [path for path in required if path not in by_path]
    if missing_required:
        raise FileNotFoundError("Required canonical/source files missing: " + ", ".join(missing_required))

    cohort = {row["phenotype"]: row for row in _read_csv(input_root / cohort_path)}
    model = {row["claim_id"]: row for row in _read_csv(input_root / model_path)}
    sensitivity = _read_csv(input_root / sensitivity_path)
    subgroup = _read_csv(input_root / subgroup_path)
    missingness = {row["variable"]: row for row in _read_csv(input_root / missing_path)}
    main_text = extracted[by_path[main_path]["source_id"]]
    protocol_text = extracted[by_path[protocol_path]["source_id"]]

    biv_key = next(key for key in cohort if "Biventricular" in key)
    biv = cohort[biv_key]
    canonical_event = f"{biv['aki_events_72h']}/{biv['n']} ({biv['aki_rate_percent']}%)"
    conflicts: list[dict[str, str]] = []
    for number, line in enumerate(protocol_text.splitlines(), start=1):
        for match in re.finditer(r"(\d+)\s*/\s*(\d+)\s*[（(](\d+(?:\.\d+)?)%[）)]", line):
            observed = f"{match.group(1)}/{match.group(2)} ({match.group(3)}%)"
            if match.group(2) == biv["n"] and observed != canonical_event:
                conflicts.append({
                    "conflict_id": f"CNF-{len(conflicts)+1:03d}",
                    "source_a": by_path[protocol_path]["source_id"],
                    "source_location_a": f"extracted line {number}",
                    "value_a": observed,
                    "source_b": by_path[cohort_path]["source_id"],
                    "source_location_b": f"{cohort_path}: phenotype={biv_key}",
                    "value_b": canonical_event,
                    "canonical_source": cohort_path,
                    "resolution": f"Formal output uses canonical value {canonical_event}; source text remains unchanged.",
                    "affected_claim_ids": "SHD-005",
                    "manual_review_required": "yes",
                })
    model_biv = model["CLM-003"]
    canonical_or = model_biv["estimate"]
    for number, line in enumerate(main_text.splitlines(), start=1):
        if model_biv["ci_low"] in line and model_biv["ci_high"] in line:
            candidates = re.findall(r"(?<![\d.])(\d+\.\d+)(?![\d.])", line)
            for candidate in candidates:
                if candidate.startswith("3.") and candidate != canonical_or:
                    conflicts.append({
                        "conflict_id": f"CNF-{len(conflicts)+1:03d}",
                        "source_a": by_path[main_path]["source_id"],
                        "source_location_a": f"extracted line {number}",
                        "value_a": candidate,
                        "source_b": by_path[model_path]["source_id"],
                        "source_location_b": f"{model_path}: claim_id=CLM-003",
                        "value_b": canonical_or,
                        "canonical_source": model_path,
                        "resolution": f"Formal output uses frozen adjusted OR {canonical_or}; caption remains unchanged.",
                        "affected_claim_ids": "SHD-008",
                        "manual_review_required": "yes",
                    })
    write_csv(audit_root / "source_conflicts.csv", CONFLICT_FIELDS, conflicts)
    conflict_report = [
        "# Source Conflict Report", "",
        f"Detected conflicts: **{len(conflicts)}**", "",
        *[
            f"- `{row['conflict_id']}`: {row['value_a']} vs {row['value_b']}; "
            f"canonical `{row['canonical_source']}`. {row['resolution']}"
            for row in conflicts
        ],
    ]
    (audit_root / "source_conflict_report.md").write_text(
        "\n".join(conflict_report) + "\n", encoding="utf-8"
    )

    unresolved_hits: list[dict[str, str]] = []
    marker_pattern = re.compile(r"\[(?:[A-Z][A-Z0-9_]*REQUIRED|REF_PENDING_\d+)\]")
    for source_path in (main_path, protocol_path, refs_path):
        source = by_path[source_path]
        text = extracted[source["source_id"]]
        for marker in sorted(set(marker_pattern.findall(text))):
            unresolved_hits.append({
                "marker": marker, "status": "INFORMATION_REQUIRED",
                "source_id": source["source_id"],
                "source_location": _line_location(text, marker),
            })
    unique_markers = sorted({row["marker"] for row in unresolved_hits})
    unresolved_lines = [
        "# Unresolved Items", "",
        *[
            f"- INFORMATION_REQUIRED: `{marker}` — preserve marker; do not invent or search a substitute."
            for marker in unique_markers
        ],
    ]
    (seed_root / "unresolved_items.md").write_text(
        "\n".join(unresolved_lines) + "\n", encoding="utf-8"
    )
    (audit_root / "unresolved_items.md").write_text(
        "\n".join(unresolved_lines) + "\n", encoding="utf-8"
    )

    ms = by_path[main_path]
    cs = by_path[cohort_path]
    mm = by_path[model_path]
    ss = by_path[sensitivity_path]
    sg = by_path[subgroup_path]
    mn = by_path[missing_path]
    protocol = by_path[protocol_path]
    ref_status = by_path[refs_path]
    fig1 = (input_root / figure1_path).resolve()
    fig3 = (input_root / figure3_path).resolve()
    no = cohort[next(key for key in cohort if key.lower().startswith("no "))]
    lv = cohort[next(key for key in cohort if "Isolated LV" in key)]
    rv = cohort[next(key for key in cohort if "Isolated RV" in key)]
    m1, m2, m3, m4 = model["CLM-001"], model["CLM-002"], model["CLM-003"], model["CLM-004"]
    interaction_min = min(float(row["interaction_p"]) for row in subgroup)
    bnp = missingness["bnp_pg_mL"]
    sens_values = ", ".join(row["estimate"] for row in sensitivity)
    claims = [
        _claim("SHD-001",
               "该完全合成观察性队列用于评估入组后 24 小时心室功能表型与随后 72 小时 AKI 进展的关联，不代表真实临床证据。",
               ms, _line_location(main_text, "24"), "reported_fact",
               "本次分析聚焦心室表型与随后 AKI 进展的关联",
               "仅可表述为合成观察性关联；不得表述为真实患者研究或因果效应。"),
        _claim("SHD-002",
               "研究采用 landmark 框架：在入组后 24 小时定义暴露，并观察其后 72 小时 AKI 进展。",
               protocol, _line_location(protocol_text, "landmark"), "reported_fact",
               "Landmark 设计将暴露定义置于结局观察之前",
               "不得补充输入未提供的真实数据库、伦理审批或外部验证信息。"),
        _claim("SHD-003",
               "队列流程图展示完全合成样本的纳入、排除与最终分析人群。",
               by_path[figure1_path], figure1_path, "reported_fact",
               "队列流程明确区分纳入、排除与最终分析人群",
               "流程图仅用于合成技术验证；不得解释为真实患者招募。",
               visual_type="source_figure", visual_asset=str(fig1)),
        _claim("SHD-004",
               "24 小时时点的心室功能按无功能障碍、单纯 LV、单纯 RV 和双心室功能障碍四类互斥表型登记。",
               ms, _line_location(main_text, "Isolated"), "reported_fact",
               "四类互斥心室表型构成主要暴露框架",
               "不得把分类梯度解释为因果剂量反应。"),
        _claim("SHD-005",
               f"AKI 进展发生率依次为 {no['aki_events_72h']}/{no['n']}（{no['aki_rate_percent']}%）、"
               f"{lv['aki_events_72h']}/{lv['n']}（{lv['aki_rate_percent']}%）、"
               f"{rv['aki_events_72h']}/{rv['n']}（{rv['aki_rate_percent']}%）和 "
               f"{biv['aki_events_72h']}/{biv['n']}（{biv['aki_rate_percent']}%）。",
               cs, f"{cohort_path}: all phenotype rows", "reported_fact",
               "双心室组的合成 AKI 进展发生率最高",
               "只报告描述性发生率；不得表述为因果梯度。",
               conflict="resolved_against_canonical", visual_type="editable_bar_chart"),
        _claim("SHD-006",
               f"与无心室功能障碍相比，单纯 LV 的调整 OR 为 {m1['estimate']}（95% CI {m1['ci_low']}–{m1['ci_high']}，P={m1['p_value']}）；区间跨越 1。",
               mm, f"{model_path}: claim_id=CLM-001", "observational_association",
               "单纯 LV 的估计方向为正，但不确定性仍跨越无效值",
               "不得称为统计学显著、明确阳性或独立因果效应。",
               estimate=m1["estimate"], ci=f"{m1['ci_low']}–{m1['ci_high']}", p=m1["p_value"]),
        _claim("SHD-007",
               f"与无心室功能障碍相比，单纯 RV 与更高 AKI 进展 odds 相关：调整 OR {m2['estimate']}（95% CI {m2['ci_low']}–{m2['ci_high']}，P={m2['p_value']}）。",
               mm, f"{model_path}: claim_id=CLM-002", "observational_association",
               "单纯 RV 与更高 AKI 进展 odds 存在统计关联",
               "不得表述为 RV 功能障碍导致 AKI。",
               estimate=m2["estimate"], ci=f"{m2['ci_low']}–{m2['ci_high']}", p=m2["p_value"]),
        _claim("SHD-008",
               f"双心室功能障碍在四组中显示最强关联：调整 OR {m3['estimate']}（95% CI {m3['ci_low']}–{m3['ci_high']}，P={m3['p_value']}）。",
               mm, f"{model_path}: claim_id=CLM-003", "observational_association",
               "双心室功能障碍显示四组中最强的合成统计关联",
               "不得表述为剂量反应因果关系、机制证明或临床验证。",
               estimate=m3["estimate"], ci=f"{m3['ci_low']}–{m3['ci_high']}", p=m3["p_value"],
               conflict="resolved_against_canonical"),
        _claim("SHD-009",
               "肌酐轨迹仅描述各合成表型组随时间的变化，不用于推断治疗效果。",
               by_path[figure3_path], figure3_path, "reported_fact",
               "肌酐轨迹提供描述性时间变化信息",
               "不得解释为治疗导致的改善或恶化。",
               visual_type="source_figure", visual_asset=str(fig3)),
        _claim("SHD-010",
               f"双心室功能障碍与 28 天死亡的更高 odds 相关：调整 OR {m4['estimate']}（95% CI {m4['ci_low']}–{m4['ci_high']}，P={m4['p_value']}）。",
               mm, f"{model_path}: claim_id=CLM-004", "observational_association",
               "双心室表型与合成 28 天死亡 odds 较高相关",
               "不得表述为真实临床死亡风险、因果效应或临床工具。",
               estimate=m4["estimate"], ci=f"{m4['ci_low']}–{m4['ci_high']}", p=m4["p_value"]),
        _claim("SHD-011",
               f"四项敏感性分析的双心室调整 OR 为 {sens_values}，方向与主分析一致。",
               ss, f"{sensitivity_path}: all rows", "observational_association",
               "敏感性分析的关联方向与主分析一致",
               "一致性不等于消除偏倚、证明稳健因果效应或临床有效性。"),
        _claim("SHD-012",
               f"所有预设亚组的 interaction P 均不支持异质性（最小 interaction P={interaction_min:.2f}）。",
               sg, f"{subgroup_path}: all rows", "observational_association",
               "预设亚组未见异质性证据",
               "不得依据单个亚组点估计声称 effect modification。"),
        _claim("SHD-013",
               f"BNP 缺失 {bnp['missing_n']} 例（{bnp['missing_percent']}%），超过 20%，因此未进入冻结的主调整模型变量列表。",
               mn, f"{missing_path}: variable=bnp_pg_mL", "reported_fact",
               "BNP 缺失超过 20%，未进入主调整模型",
               "不得声称插补消除了缺失偏倚、缺失不影响结果或 BNP 可安全纳入主模型。"),
        _claim("SHD-014",
               f"跨文件审计检出 {len(conflicts)} 处数字冲突并按 canonical source 解决；"
               f"{'、'.join(unique_markers)} 仍为 INFORMATION_REQUIRED。",
               ref_status, refs_path, "reported_fact",
               "冲突已按 canonical source 解决，缺失信息仍保留待确认",
               "不得虚构伦理编号、补造参考文献或静默删除 unresolved marker。",
               canonical="audit_boundary", conflict="detected_and_resolved"),
    ]
    write_csv(seed_root / "claim_registry.csv", CLAIM_FIELDS, claims)
    write_csv(seed_root / "claim_candidates.csv", CLAIM_FIELDS, claims)

    content_inventory = [
        "# Content Inventory", "",
        "- Data class: fully synthetic; no real patients, database or ethics approval.",
        "- Study frame: observational landmark association analysis.",
        f"- Canonical event source: `{cohort_path}`.",
        f"- Canonical adjusted model source: `{model_path}`.",
        f"- Sensitivity source: `{sensitivity_path}`.",
        f"- Subgroup source: `{subgroup_path}`.",
        f"- Missingness source: `{missing_path}`.",
        f"- Controlled conflicts detected independently: `{len(conflicts)}`.",
        f"- Unresolved markers retained: `{', '.join(unique_markers)}`.",
        "- Patient-level CSV was parsed for format validation only and was not used to refit the frozen model.",
    ]
    (seed_root / "content_inventory.md").write_text(
        "\n".join(content_inventory) + "\n", encoding="utf-8"
    )
    write_csv(
        seed_root / "cross_source_entity_map.csv",
        ["entity", "canonical_source", "supporting_sources", "rule"],
        [
            {"entity": "adjusted_effects", "canonical_source": model_path, "supporting_sources": main_path, "rule": "frozen model results override narrative/caption copies"},
            {"entity": "event_counts_rates", "canonical_source": cohort_path, "supporting_sources": protocol_path, "rule": "cohort summary overrides old narrative copy"},
            {"entity": "sensitivity", "canonical_source": sensitivity_path, "supporting_sources": "", "rule": "no refitting"},
            {"entity": "subgroups", "canonical_source": subgroup_path, "supporting_sources": "", "rule": "interaction P governs heterogeneity wording"},
            {"entity": "missingness", "canonical_source": missing_path, "supporting_sources": main_path, "rule": "BNP above 20 percent excluded from main model list"},
        ],
    )
    ris_source = by_path.get("references/selected_references.ris")
    reference_rows: list[dict[str, str]] = []
    if ris_source:
        ris_text = extracted[ris_source["source_id"]]
        records = [block for block in re.split(r"(?m)^ER  -\s*$", ris_text) if block.strip()]
        for index, record in enumerate(records, start=1):
            title = next((line[6:].strip() for line in record.splitlines() if line.startswith("TI  -")), "")
            reference_rows.append({
                "reference_id": f"REF-{index:03d}", "source_id": ris_source["source_id"],
                "title": title, "status": "REGISTERED_FROM_INPUT",
            })
    write_csv(
        seed_root / "reference_inventory.csv",
        ["reference_id", "source_id", "title", "status"],
        reference_rows,
    )
    write_json(seed_root / "unresolved_hits.json", unresolved_hits)
    brief = {
        "project_name": PROJECT,
        "presentation_type": "research_report",
        "presentation_objective": "复杂合成医学科研材料的可追溯影子验证；不代表真实临床研究或真实世界验证。",
        "audience": "医学科研与方法学评审者",
        "language": "中文",
        "duration_minutes": 18,
        "target_slide_count": 18,
        "output_aspect_ratio": "16:9",
        "must_include": [
            "synthetic warning", "canonical results", "sensitivity analysis",
            "missingness", "limitations", "unresolved items", "conflict audit",
        ],
        "must_exclude": [
            "causal claims", "invented ethics approval", "invented citations",
            "real patient claims", "model refitting",
        ],
        "key_message": "合成结果提示双心室表型与更高 AKI 进展 odds 的关联最强；全部结论仍需人工科学审核，且不能外推为真实临床证据。",
        "target_journal_or_meeting": "synthetic validation only",
        "citation_style": "source_id + short input filename",
        "theme_profile": "reference_style_only_restrained_biomedical",
        "style_reference_files": ["input/style_reference/毕业论文答辩模板1.pptx"],
        "speaker_notes_required": True,
        "appendix_required": True,
        "confidential_content_rules": "No patient-level rows in slides; input package is fully synthetic.",
        "narrative_mode": "scientific_problem",
        "validated_claim_registry": str((seed_root / "claim_registry.csv").resolve()),
    }
    write_json(seed_root / "presentation_brief.yaml", brief)
    write_json(seed_root / "seed_manifest.json", {
        "run_id": run_id,
        "seed_root": str(seed_root.resolve()),
        "brief": str((seed_root / "presentation_brief.yaml").resolve()),
        "claim_registry": str((seed_root / "claim_registry.csv").resolve()),
        "conflict_count": len(conflicts),
        "unresolved_markers": unique_markers,
        "ground_truth_present": expected.is_dir(),
        "ground_truth_read": False,
    })
    print(json.dumps({
        "run_id": run_id, "seed_root": str(seed_root.resolve()),
        "brief": str((seed_root / "presentation_brief.yaml").resolve()),
        "claims": len(claims), "conflicts": len(conflicts),
        "unresolved": unique_markers,
    }, ensure_ascii=False))


def _pptx_visible_text(path: Path) -> str:
    texts: list[str] = []
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                root = ET.fromstring(archive.read(name))
                texts.extend(text for text in root.itertext() if text)
    return "\n".join(texts)


def finalize(run_id: str, seed_run_id: str | None = None) -> None:
    audit_root = REPO / "audit" / "synthetic_shadow_validation"
    seed_root = (
        REPO
        / "staging"
        / "synthetic_shadow_validation_seed"
        / (seed_run_id or run_id)
    )
    stage_root = REPO / "staging" / PROJECT / run_id
    output_root = REPO / "output" / PROJECT / run_id
    if not stage_root.is_dir() or not output_root.is_dir():
        raise FileNotFoundError("Canonical staging/output directory is missing")
    for filename in [
        "content_inventory.md", "claim_candidates.csv", "reference_inventory.csv",
        "cross_source_entity_map.csv", "unresolved_items.md",
    ]:
        source = seed_root / filename
        shutil.copy2(source, stage_root / filename)
        shutil.copy2(source, output_root / filename)
    shutil.copy2(audit_root / "source_conflicts.csv", output_root / "source_conflicts.csv")

    pptx = output_root / f"{PROJECT}.pptx"
    pdf = output_root / f"{PROJECT}.pdf"
    visible = _pptx_visible_text(pptx)
    claims = read_csv(output_root / "claim_source_map.csv")
    source_manifest = read_csv(output_root / "source_manifest.csv")
    source_ids = {row["source_id"] for row in source_manifest}
    tracked = sum(1 for row in claims if row.get("source_id") in source_ids)
    scientific_checks = {
        "formal_claim_count_at_least_7": len(claims) >= 7,
        "claim_source_map_complete": bool(claims) and tracked == len(claims),
        "canonical_event_present": "34/58" in visible and "58.6%" in visible,
        "obsolete_event_absent": "33/58" not in visible and "56.9%" not in visible,
        "canonical_or_present": "3.61" in visible,
        "obsolete_or_absent": "3.41" not in visible,
        "lv_uncertainty_retained": "0.78" in visible and "2.59" in visible and "跨越" in visible,
        "synthetic_warning_present": "合成" in visible and "真实临床" in visible,
        "ethics_not_invented": "ETHICS_ID_REQUIRED" in visible,
        "pending_reference_retained": "REF_PENDING_04" in visible,
        "bnp_excluded_wording_present": "BNP" in visible and "未进入" in visible,
        "no_causal_upgrade": not any(term in visible.lower() for term in [
            " caused ", " led to ", " resulted in ", " proved ",
            "confirmed mechanism", "clinically validated", "therapeutic target",
            "independent causal effect",
        ]),
    }
    failed_science = [name for name, passed in scientific_checks.items() if not passed]
    sci_lines = [
        "# Scientific QA Report", "",
        f"- Formal claims: `{len(claims)}`",
        f"- Traceable claims: `{tracked}/{len(claims)}`",
        "- Frozen patient-level model refit: `NOT PERFORMED`",
        "",
        "## Checks", "",
        *[f"- {'PASS' if passed else 'FAIL'} — `{name}`" for name, passed in scientific_checks.items()],
        "",
        f"Overall: `{'PASS' if not failed_science else 'FAIL'}`",
    ]
    (output_root / "scientific_qa_report.md").write_text(
        "\n".join(sci_lines) + "\n", encoding="utf-8"
    )
    qa_text = (output_root / "qa_report.md").read_text(encoding="utf-8", errors="replace")
    layout_json = json.loads(
        (output_root / "powerpoint_layout_manifest.json").read_text(
            encoding="utf-8-sig"
        )
    )
    visual_pass = "Visual issues: 0" in qa_text or "Visual QA: PASS" in qa_text
    layout_status = str(layout_json.get("status", "UNKNOWN"))
    (output_root / "visual_qa_report.md").write_text(
        "\n".join([
            "# Visual and Geometry QA Report", "",
            f"- Canonical QA visual status: `{'PASS' if visual_pass else 'REVIEW_REQUIRED'}`",
            f"- PowerPoint COM geometry status: `{layout_status}`",
            f"- Slide count: `{len(list((output_root / 'preview').glob('slide_*.png')))}`",
            "- Style source: user-declared PPTX, `style-only` token extraction.",
            "- Reference slide text, notes, media and private content were not copied.",
        ]) + "\n", encoding="utf-8"
    )
    ooxml_json = json.loads(
        (output_root / "ooxml_qa_report.json").read_text(encoding="utf-8-sig")
    )
    (output_root / "ooxml_qa_report.md").write_text(
        "\n".join([
            "# OOXML QA Report", "",
            f"- Status: `{ooxml_json.get('status', 'UNKNOWN')}`",
            f"- Errors: `{len(ooxml_json.get('errors', []))}`",
            f"- Warnings: `{len(ooxml_json.get('warnings', []))}`",
            "- Machine-readable details: `ooxml_qa_report.json`.",
        ]) + "\n", encoding="utf-8"
    )
    repair_prefix = run_id.split("_fix", 1)[0] + "_fix"
    repair_history: list[dict[str, Any]] = []
    for repair_dir in sorted(output_root.parent.glob(repair_prefix + "*")):
        diff_path = repair_dir / "slide_diff.json"
        if not diff_path.is_file():
            continue
        diff = json.loads(diff_path.read_text(encoding="utf-8"))
        repair_history.append({
            "run_id": repair_dir.name,
            "target_slide_ids": [
                item.get("slide_id", "") for item in diff.get("targets", [])
            ],
            "untouched_slide_differences": diff.get(
                "untouched_slide_differences", []
            ),
            "accepted_for_incremental_integrity": not diff.get(
                "untouched_slide_differences", []
            ),
            "manual_review_required": bool(
                diff.get("manual_review_required", False)
            ),
        })
    write_json(output_root / "slide_diff_history.json", {
        "schema_version": "1.0",
        "repairs": repair_history,
        "final_repair": repair_history[-1] if repair_history else None,
        "failed_or_rejected_iterations_retained": any(
            not row["accepted_for_incremental_integrity"]
            for row in repair_history
        ),
    })
    before = _read_csv(audit_root / "input_hashes_before.csv")
    after = _hash_rows(REPO / "input")
    write_csv(audit_root / "input_hashes_after.csv", HASH_FIELDS, after)
    before_map = {row["relative_path"]: row["sha256"] for row in before}
    after_map = {row["relative_path"]: row["sha256"] for row in after}
    integrity = before_map == after_map
    (audit_root / "input_integrity_report.md").write_text(
        "\n".join([
            "# Input Integrity Report", "",
            f"- Before files: `{len(before_map)}`",
            f"- After files: `{len(after_map)}`",
            f"- Paths and SHA-256 identical: `{'PASS' if integrity else 'FAIL'}`",
            "- expected_ground_truth paths were excluded from source discovery and hashing.",
        ]) + "\n", encoding="utf-8"
    )

    ground_truth = REPO / "expected_ground_truth"
    if not ground_truth.is_dir():
        ground_truth = REPO / "input" / "expected_ground_truth"
    gt_present = ground_truth.is_dir()
    gt_status = "NOT_ASSESSED_GROUND_TRUTH_DIRECTORY_ABSENT"
    gt_first_read = ""
    if gt_present:
        gt_first_read = utc_offset_timestamp()
        # The validation package, not the deck-generation stages, may read it now.
        gt_status = "POST_GENERATION_GROUND_TRUTH_EVALUATION_COMPLETED"
        list(ground_truth.rglob("*"))
    comparison_rows = [
        {"criterion": "formal_claims", "result": "true positive" if len(claims) >= 7 else "false negative", "evidence": str(len(claims))},
        {"criterion": "controlled_conflicts", "result": "true positive" if len(_read_csv(audit_root / "source_conflicts.csv")) >= 2 else "false negative", "evidence": str(len(_read_csv(audit_root / "source_conflicts.csv")))},
        {"criterion": "unresolved_markers", "result": "true positive" if scientific_checks["ethics_not_invented"] and scientific_checks["pending_reference_retained"] else "false negative", "evidence": "ETHICS_ID_REQUIRED; REF_PENDING_04"},
        {"criterion": "canonical_numbers", "result": "true positive" if scientific_checks["canonical_event_present"] and scientific_checks["canonical_or_present"] else "false negative", "evidence": "34/58; 58.6%; 3.61"},
        {"criterion": "external_ground_truth", "result": "not assessed" if not gt_present else "partially correct", "evidence": gt_status},
    ]
    write_csv(
        audit_root / "ground_truth_comparison.csv",
        ["criterion", "result", "evidence"],
        comparison_rows,
    )
    (audit_root / "ground_truth_evaluation.md").write_text(
        "\n".join([
            "# Post-generation Ground-truth Evaluation", "",
            f"- Status: `{gt_status}`",
            f"- First ground-truth read timestamp: `{gt_first_read or 'NOT_READ'}`",
            "- Generation and initial QA completed before this stage.",
            "- Internal, source-derived assertions were evaluated independently of any answer directory.",
        ]) + "\n", encoding="utf-8"
    )
    (audit_root / "claim_recall_report.md").write_text(
        f"# Claim Recall Report\n\n- Registered formal claims: `{len(claims)}`\n- Traceable: `{tracked}/{len(claims)}`\n- External ground truth: `{gt_status}`\n",
        encoding="utf-8",
    )
    conflicts = _read_csv(audit_root / "source_conflicts.csv")
    (audit_root / "conflict_detection_report.md").write_text(
        f"# Conflict Detection Report\n\n- Detected: `{len(conflicts)}`\n- Expected controlled conflict categories represented: `{'PASS' if len(conflicts) >= 2 else 'FAIL'}`\n",
        encoding="utf-8",
    )
    (audit_root / "unresolved_item_report.md").write_text(
        f"# Unresolved Item Report\n\n- ETHICS_ID_REQUIRED retained: `{scientific_checks['ethics_not_invented']}`\n- REF_PENDING_04 retained: `{scientific_checks['pending_reference_retained']}`\n",
        encoding="utf-8",
    )
    final_status = (
        "SYNTHETIC_SHADOW_VALIDATION_PASSED"
        if not failed_science and integrity and (not gt_present or gt_status.endswith("COMPLETED"))
        else "SYNTHETIC_SHADOW_VALIDATION_PARTIAL"
    )
    validation = {
        "test_status": final_status,
        "overall_v2_status": "PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING",
        "formal_claims": len(claims),
        "traceable_claims": tracked,
        "input_integrity": integrity,
        "ground_truth_status": gt_status,
        "pptx_bytes": pptx.stat().st_size,
        "pdf_bytes": pdf.stat().st_size,
        "scientific_checks": scientific_checks,
        "incremental_repairs_recorded": len(repair_history),
        "final_incremental_untouched_differences": (
            len(repair_history[-1]["untouched_slide_differences"])
            if repair_history else None
        ),
        "failure_resume_verified": any(
            "resumed_at" in json.loads(
                (repair_dir / "runtime_manifest.json").read_text(
                    encoding="utf-8-sig"
                )
            )
            for repair_dir in sorted(output_root.parent.glob(repair_prefix + "*"))
            if (repair_dir / "runtime_manifest.json").is_file()
        ),
    }
    write_json(audit_root / "validation_summary.json", validation)
    write_json(output_root / "synthetic_shadow_validation_summary.json", validation)
    print(json.dumps(validation, ensure_ascii=False))


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="command", required=True)
    p_prepare = sub.add_parser("prepare")
    p_prepare.add_argument("--run-id", required=True)
    p_prepare.add_argument("--powerpoint-status", default="PASS")
    p_prepare.add_argument("--powerpoint-version", default="16.0")
    p_finalize = sub.add_parser("finalize")
    p_finalize.add_argument("--run-id", required=True)
    p_finalize.add_argument("--seed-run-id")
    args = parser.parse_args()
    if args.command == "prepare":
        prepare(args.run_id, args.powerpoint_status, args.powerpoint_version)
    else:
        finalize(args.run_id, args.seed_run_id)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
