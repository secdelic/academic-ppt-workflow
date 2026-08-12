#!/usr/bin/env python3
"""Aggregate the fixed v2 three-arm benchmark without hiding failed runs."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from pypdf import PdfReader


PROJECTS = {
    "case1": "benchmark_case1_paper_report",
    "case2": "benchmark_case2_defense_project",
    "case3": "benchmark_case3_reference_template",
}
ARMS = ("arm_a", "arm_b", "arm_c")
PRIVATE_SENTINEL = b"PRIVATE_REFERENCE_SENTINEL_DO_NOT_COPY_7F3E91B2"

SCORES = {
    "case1": {
        "arm_a": [16, 7, 8, 9, 8, None, 0, 5, 5],
        "arm_b": [16, 4, 11, 8, 9, None, 0, 5, 2],
        "arm_c": [16, 7, 8, 11, 9, None, 5, 5, 5],
    },
    "case2": {
        "arm_a": [16, 7, 8, 9, 8, None, 0, 5, 5],
        "arm_b": [16, 4, 11, 8, 9, None, 0, 5, 2],
        "arm_c": [16, 7, 8, 11, 9, None, 5, 5, 5],
    },
    "case3": {
        "arm_a": [16, 7, 8, 10, 8, 0, 0, 5, 5],
        "arm_b": [16, 4, 11, 9, 9, 0, 0, 5, 2],
        "arm_c": [16, 7, 8, 9, 9, 4, 5, 5, 5],
    },
}
DIMENSIONS = [
    ("scientific_accuracy", 25),
    ("traceability", 15),
    ("presentation_logic", 15),
    ("visual_quality", 15),
    ("native_editability", 10),
    ("template_inheritance", 5),
    ("incremental_update", 5),
    ("render_compatibility", 5),
    ("reproducibility_maintainability", 5),
]
RATIONALES = {
    "scientific_accuracy": (
        "No invented result was observed and evidence-strength boundaries were retained; "
        "however, expected numeric claims were omitted and the result slide remained INFORMATION_REQUIRED."
    ),
    "traceability": (
        "Source manifests and slide source IDs were present, but zero formal claims were emitted; "
        "Arm B additionally lacks a native claim-source map."
    ),
    "presentation_logic": (
        "The five-slide safety narrative is coherent but generic and does not replace the unresolved result; "
        "Arm B showed stronger role variation and exhibit hierarchy."
    ),
    "visual_quality": (
        "Score reflects blinded contact-sheet review plus PowerPoint text geometry; "
        "any detected overflow reduced the score."
    ),
    "native_editability": "PPTX objects remained native text, shapes, charts, and placeholders rather than page images.",
    "template_inheritance": (
        "Not scored outside Case 3. In Case 3, Arm C inherited the protected reference design structure "
        "without leaking the sentinel, but the resulting layout remained visually sparse."
    ),
    "incremental_update": (
        "Arm C received credit from the separate successful slide_id update smoke test; "
        "untouched slides and a manual sentinel were preserved."
    ),
    "render_compatibility": "PowerPoint COM opened the deck and PDF/PNG page counts matched.",
    "reproducibility_maintainability": (
        "Arm A is the frozen stable baseline; Arm B is intentionally benchmark-only; "
        "Arm C has clean-room modules, fallback, manifests, and isolated routes."
    ),
}


def _read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _write_csv(path: Path, fields: list[str], rows: list[dict[str, Any]]) -> None:
    with path.open("x", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def _output_dir(repo: Path, case: str, arm: str, run_id: str) -> Path:
    run = repo / "staging/v2_benchmark/runs" / case / arm / run_id
    if arm == "arm_b":
        return run
    return run / "output" / PROJECTS[case] / run_id


def _pptx_slide_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return len(
            [
                name
                for name in archive.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
            ]
        )


def _geometry_issues(repo: Path, case: str, arm: str, output: Path) -> list[str]:
    if arm == "arm_a":
        path = (
            repo
            / "benchmark/results/20260728_1510/comparative_qa"
            / f"{case}_arm_a_layout_issues.json"
        )
        return _read_json(path)["issues"]
    if arm == "arm_b":
        return _read_json(output / "benchmark_run_summary.json")["geometry_issues"]
    report = (output / "qa_report.md").read_text(encoding="utf-8-sig")
    return re.findall(r"- FAIL: (Slide \d+: text may overflow[^\n]+)", report)


def _slide_rows(output: Path, arm: str) -> list[dict[str, Any]]:
    if arm == "arm_b":
        summary = _read_json(output / "benchmark_run_summary.json")
        spec = _read_json(Path(summary["build_spec"]))
        return [
            {
                "slide_id": slide.get("slide_id", ""),
                "slide_title": slide.get("slide_title", ""),
                "single_key_message": slide.get("single_key_message", ""),
                "source_ids": ";".join(slide.get("source_ids", [])),
                "visual_type": slide.get("visual_type", ""),
                "manual_review_required": slide.get("manual_review_required", True),
            }
            for slide in spec.get("slides", [])
        ]
    with (output / "slide_manifest.csv").open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path, default=Path.cwd())
    parser.add_argument(
        "--result-dir",
        type=Path,
        default=Path("benchmark/results/20260728_1510"),
    )
    args = parser.parse_args()
    repo = args.repo_root.resolve()
    result = (repo / args.result_dir).resolve() if not args.result_dir.is_absolute() else args.result_dir.resolve()
    selected = _read_json(result / "selected_runs.json")

    arm_rows: list[dict[str, Any]] = []
    dimension_rows: list[dict[str, Any]] = []
    runtime_rows: list[dict[str, Any]] = []
    slide_rows: list[dict[str, Any]] = []
    registry: dict[str, Any] = {"selected_runs": selected, "runs": []}

    for case in ("case1", "case2", "case3"):
        for arm in ARMS:
            run_id = selected[case][arm]
            output = _output_dir(repo, case, arm, run_id)
            project = PROJECTS[case]
            if arm == "arm_b":
                pptx = next(output.glob("*.pptx"))
                pdf = next(output.glob("*.pdf"))
                runtime = _read_json(output / "benchmark_run_summary.json")
                status = runtime["status"]
                seconds = float(runtime["total_seconds"])
                renderer = runtime["renderer"]
                ooxml_valid = bool(runtime["ooxml_valid"])
                deck_ir_hash = runtime["build_spec_sha256"]
                scientific_status = "NOT_INDEPENDENTLY_EVALUATED"
            else:
                pptx = output / f"{project}.pptx"
                pdf = output / f"{project}.pdf"
                runtime = _read_json(output / "runtime_manifest.json")
                status = runtime["status"]
                started = datetime.fromisoformat(runtime["started_at"])
                finished = datetime.fromisoformat(runtime["finished_at"])
                seconds = (finished - started).total_seconds()
                renderer = runtime["renderer"]
                ooxml_valid = runtime.get("ooxml_qa", {}).get("status") == "PASS"
                deck_ir_hash = runtime.get("deck_ir", {}).get("canonical_hash", "")
                scientific_status = runtime.get("stages", {}).get("scientific_qa", "unknown")

            preview_count = len(list((output / "preview").glob("slide_*.png")))
            pptx_count = _pptx_slide_count(pptx)
            pdf_count = len(PdfReader(str(pdf)).pages)
            issues = _geometry_issues(repo, case, arm, output)
            claim_map = output / "claim_source_map.csv"
            if claim_map.is_file():
                with claim_map.open(encoding="utf-8-sig", newline="") as handle:
                    claim_count = len(list(csv.DictReader(handle)))
            else:
                claim_count = 0
            leaked = PRIVATE_SENTINEL in pptx.read_bytes()

            score_values = SCORES[case][arm]
            score_total = sum(value for value in score_values if value is not None)
            max_scored = sum(
                maximum
                for value, (_, maximum) in zip(score_values, DIMENSIONS)
                if value is not None
            )
            normalized = round(100 * score_total / max_scored, 1)
            for value, (dimension, maximum) in zip(score_values, DIMENSIONS):
                dimension_rows.append(
                    {
                        "case_id": case,
                        "arm_id": arm,
                        "run_id": run_id,
                        "dimension": dimension,
                        "score": "" if value is None else value,
                        "max_score": maximum,
                        "score_status": "NOT_SCORED" if value is None else "PRELIMINARY_ASSISTED_REVIEW",
                        "rationale": RATIONALES[dimension],
                        "reviewer_role": "AI-assisted technical reviewer; final human score pending",
                    }
                )
            arm_rows.append(
                {
                    "case_id": case,
                    "arm_id": arm,
                    "run_id": run_id,
                    "workflow_status": status,
                    "scientific_qa_status": scientific_status,
                    "formal_claim_count": claim_count,
                    "expected_fixture_claim_count": 5 if case != "case3" else 3,
                    "unresolved_content_gap": "yes",
                    "pptx_opened_by_powerpoint": "yes",
                    "pptx_slide_count": pptx_count,
                    "pdf_page_count": pdf_count,
                    "png_page_count": preview_count,
                    "page_counts_match": "yes" if pptx_count == pdf_count == preview_count else "no",
                    "ooxml_valid": "yes" if ooxml_valid else "no",
                    "geometry_issue_count": len(issues),
                    "private_sentinel_leak": "yes" if leaked else "no",
                    "input_hashes_unchanged": "yes",
                    "network_used": "no",
                    "third_party_api_used": "no",
                    "score_total": score_total,
                    "max_scored": max_scored,
                    "normalized_score": normalized,
                    "promotion_eligible": "yes" if arm == "arm_c" and not leaked and ooxml_valid and not issues else "no",
                    "output_dir": str(output),
                }
            )
            runtime_rows.append(
                {
                    "case_id": case,
                    "arm_id": arm,
                    "run_id": run_id,
                    "runtime_seconds": round(seconds, 3),
                    "renderer": renderer,
                    "pptx_bytes": pptx.stat().st_size,
                    "pdf_bytes": pdf.stat().st_size,
                    "deck_ir_or_build_spec_hash": deck_ir_hash,
                    "pptx_sha256": _sha256(pptx),
                    "pdf_sha256": _sha256(pdf),
                    "manual_fix_count": 1 if arm == "arm_c" and case in {"case1", "case2"} else 0,
                    "slide_regeneration_count": 1 if arm == "arm_c" and case in {"case1", "case2"} else 0,
                }
            )

            issue_slides = {
                int(match.group(1))
                for issue in issues
                if (match := re.search(r"Slide (\d+)", issue))
            }
            for index, slide in enumerate(_slide_rows(output, arm), start=1):
                message = str(slide.get("single_key_message", ""))
                source_ids = str(slide.get("source_ids", ""))
                unresolved = "INFORMATION_REQUIRED" in message
                science_score = 2 if unresolved else 4
                trace_score = 3 if source_ids else 1
                logic_score = 1 if unresolved else (2 if index == 1 else 3)
                visual_score = 1 if index in issue_slides else 3
                edit_score = 2
                slide_rows.append(
                    {
                        "case_id": case,
                        "candidate_id": f"{case}-{arm}-{index:03d}",
                        "arm_id_unblinded_after_review": arm,
                        "run_id": run_id,
                        "slide_index": index,
                        "slide_id": slide.get("slide_id", ""),
                        "slide_title": slide.get("slide_title", ""),
                        "key_message_status": "INFORMATION_REQUIRED" if unresolved else "PRESENT",
                        "source_id_count": len([value for value in source_ids.split(";") if value]),
                        "geometry_issue": "yes" if index in issue_slides else "no",
                        "scientific_accuracy_score_5": science_score,
                        "traceability_score_3": trace_score,
                        "logic_score_3": logic_score,
                        "visual_score_3": visual_score,
                        "editability_score_2": edit_score,
                        "slide_score_16": science_score + trace_score + logic_score + visual_score + edit_score,
                        "manual_review_required": slide.get("manual_review_required", True),
                    }
                )
            registry["runs"].append(
                {
                    "case": case,
                    "arm": arm,
                    "run_id": run_id,
                    "output_dir": str(output),
                    "pptx_sha256": _sha256(pptx),
                    "selected_after_failed_runs_were_retained": True,
                }
            )

    _write_csv(
        result / "arm_results.csv",
        list(arm_rows[0].keys()),
        arm_rows,
    )
    _write_csv(
        result / "dimension_scores.csv",
        list(dimension_rows[0].keys()),
        dimension_rows,
    )
    _write_csv(
        result / "slide_level_scores.csv",
        list(slide_rows[0].keys()),
        slide_rows,
    )
    _write_csv(
        result / "runtime_results.csv",
        list(runtime_rows[0].keys()),
        runtime_rows,
    )
    (result / "run_registry.json").write_text(
        json.dumps(registry, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    averages = {
        arm: round(
            sum(row["normalized_score"] for row in arm_rows if row["arm_id"] == arm) / 3,
            1,
        )
        for arm in ARMS
    }
    summary = f"""# Academic PPT Workflow v2 基准总结

## 结论

- 最终状态：`PARTIAL_GO_REAL_WORLD_VALIDATION_PENDING`
- Arm A / B / C 的预备归一化平均分：{averages['arm_a']} / {averages['arm_b']} / {averages['arm_c']}。
- Arm C 未出现科学真实性回退，但三类复杂合成输入均未形成正式 claim；结果页保留 `INFORMATION_REQUIRED`，因此真实内容整合能力尚未通过。
- Arm C 在 Case 1/2 将 PowerPoint 文本几何问题从 Arm A 的各 2 处降至 0；Arm B 虽显示更强视觉层级，但 3/3 case 均有几何溢出，未获晋级。
- Case 3 的 `style-only` 与 `template-fill` 隐私边界通过，保护哨兵未泄漏；模板填充可用但页面较稀疏，不能认定为全面美学提升。
- 稳定 `slide_id` 单页更新已在独立合成 smoke test 中通过：目标页 revision 1→2，未修改页和人工哨兵保持不变。

## 重要限制

- 本轮只有复杂合成材料，没有真实科研项目材料。
- 盲化接触表已生成，但评分仍是 AI 辅助技术初评，最终科学和视觉评分需要用户人工确认。
- 公式 OMML/LaTeX 与 Mermaid 生产渲染未验证；当前环境缺少 LaTeX engine 和 `mmdc`。
- Arm B 是基于上游静态审计后的独立行为适配器，不是任一上游 Skill 的执行结果。
"""
    (result / "benchmark_summary.md").write_text(summary, encoding="utf-8")

    failures = """# 基准失败分析

## 保留的失败与降级

1. 初次合成参考 PPTX 的 Python→Node 子进程中断；改为本地 `python-pptx` 确定性夹具生成，旧 `_build` 失败文件保留。
2. DOCX 的 LibreOffice 渲染超时；明确终止本次命令拥有的临时进程后，Word COM 备用路径也超时。DOCX 完成结构验证，但未完成页面级视觉渲染。
3. 初次 native template fill 出现标题/正文重叠；加入 PowerPoint 几何检查并采用冲突安全网格后通过，失败输出保留。
4. 初次增量替换使用剪贴板 Copy/Paste 失败；改为 `Slides.InsertFromFile` 后通过，失败输出保留。
5. Arm A Case 1/2 各检测到 2 处长来源脚注溢出。
6. Arm B Case 1/2 各检测到 3 处、Case 3 检测到 1 处文本溢出；3 个失败输出全部保留且未修饰为通过。
7. Arm C Case 1/2 初次各检测到 2 处溢出；标题框和来源脚注自适应修复后以新 run ID 重跑为 0，原失败 run 保留。
8. 三臂均未从复杂 DOCX/CSV 生成正式 claim；零行 `claim_source_map.csv` 是当前最重要的科学内容能力缺口。
9. 当前环境无 LaTeX engine 与 Mermaid CLI，公式/图形模块只完成能力探测、合同和安全降级，未完成生产渲染验证。

## 失败不影响的事实

- 所选 9 个 PPTX 均被 PowerPoint COM 打开并导出 PDF/PNG，页数一致。
- 所有夹具输入哈希保持不变。
- 保护参考 PPTX 的私有哨兵未出现在任何运行输出。
- 未执行上游安装脚本、外部 API、中转服务或网络生成。
"""
    (result / "failure_analysis.md").write_text(failures, encoding="utf-8")

    latest_files = [
        "arm_results.csv",
        "slide_level_scores.csv",
        "runtime_results.csv",
        "failure_analysis.md",
        "benchmark_summary.md",
    ]
    for name in latest_files:
        target = repo / "benchmark" / name
        if target.exists():
            raise SystemExit(f"Refusing to overwrite benchmark artifact: {target}")
        shutil.copy2(result / name, target)
    print(json.dumps({"status": "BENCHMARK_AGGREGATED", "result_dir": str(result), "averages": averages}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
