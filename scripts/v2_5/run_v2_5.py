from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import importlib.util
import json
import math
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def load_v25_model():
    """Load the sibling v2.5 model under a versioned module name.

    The regression suite imports several historical runners in one Python
    process.  Those runners also use the generic module name ``model``; relying
    on that shared name can therefore bind this runner to a frozen v2.4 module.
    A versioned import keeps the presentation-layer contract deterministic
    without changing any frozen scientific module.
    """
    model_path = Path(__file__).resolve().with_name("model.py")
    spec = importlib.util.spec_from_file_location("academic_ppt_v25_model", model_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load v2.5 model: {model_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules.setdefault(spec.name, module)
    spec.loader.exec_module(module)
    return module


v25 = load_v25_model()


REPO = Path(__file__).resolve().parents[2]
SENSITIVE_ROW_FIELDS = re.compile(r"(?:^|_)(?:patient|subject|participant|cell|sample|synthetic)?_?id$|raw_records?", re.I)
SENSITIVE_CATEGORY_VALUE = re.compile(
    r"^(?:patient|subject|participant|sample|mrn)[\s_:#-]*[A-Za-z0-9-]+$",
    re.I,
)
ABSOLUTE_PATH_VALUE = re.compile(r"^(?:[A-Za-z]:[\\/]|/|\\\\)")
NATIVE_FIELD_ALLOWLIST = {
    "weight_histogram": {"bin_low", "bin_high", "count"},
    "grouped_composition": {"group", "cell_type", "cell_count", "percentage"},
    "funnel_plot": {"study_label", "log_effect", "standard_error"},
}


def load_v24_runner():
    """Load v2.4 derivation helpers without creating another evidence path."""
    v24_dir = REPO / "scripts" / "v2_4"
    model_spec = importlib.util.spec_from_file_location("academic_ppt_v24_model", v24_dir / "model.py")
    if model_spec is None or model_spec.loader is None:
        raise RuntimeError("Unable to load frozen v2.4 model")
    v24_model = importlib.util.module_from_spec(model_spec)
    model_spec.loader.exec_module(v24_model)
    original_model = sys.modules.get("model")
    sys.modules["model"] = v24_model
    try:
        runner_spec = importlib.util.spec_from_file_location("academic_ppt_v24_runner", v24_dir / "run_v2_4.py")
        if runner_spec is None or runner_spec.loader is None:
            raise RuntimeError("Unable to load frozen v2.4 runner")
        runner = importlib.util.module_from_spec(runner_spec)
        runner_spec.loader.exec_module(runner)
    finally:
        if original_model is None:
            sys.modules.pop("model", None)
        else:
            sys.modules["model"] = original_model
    return runner


V24 = load_v24_runner()


def dump_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if fields is None:
        fields = list(rows[0]) if rows else []
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def project_source_path(project_root: Path, source_file: object) -> Path:
    """Resolve a registered source without allowing answer-key or path escape.

    Frozen graph hashes prove which registry was approved, but source paths are
    still treated as untrusted filesystem input.  Resolve links before the
    containment check so a junction/symlink cannot alias an excluded tree.
    """
    root = project_root.resolve()
    raw = str(source_file or "").strip()
    if not raw:
        raise RuntimeError("Registered source has an empty source_file")
    relative = Path(raw)
    if relative.is_absolute():
        raise RuntimeError(f"Registered source must be project-relative: {raw}")
    if any(part.lower() == "expected_ground_truth" for part in relative.parts):
        raise RuntimeError(f"Registered source enters expected_ground_truth: {raw}")
    candidate = (root / relative).resolve()
    try:
        resolved_relative = candidate.relative_to(root)
    except ValueError as exc:
        raise RuntimeError(f"Registered source escapes project root: {raw}") from exc
    if any(part.lower() == "expected_ground_truth" for part in resolved_relative.parts):
        raise RuntimeError(f"Registered source resolves into expected_ground_truth: {raw}")
    return candidate


def baseline_projects(baseline_root: Path, suite_root: Path) -> list[tuple[Path, Path]]:
    index_path = baseline_root / "deck_index.csv"
    if not index_path.is_file():
        raise RuntimeError(f"Missing v2.4 deck index: {index_path}")
    with index_path.open("r", encoding="utf-8-sig", newline="") as handle:
        names = [row["project_key"] for row in csv.DictReader(handle)]
    pairs: list[tuple[Path, Path]] = []
    for name in names:
        baseline = baseline_root / name
        project = suite_root / name
        if not (baseline / "canonical_object_graph.json").is_file():
            raise RuntimeError(f"Missing frozen graph for {name}")
        if not project.is_dir():
            raise RuntimeError(f"Missing suite project for {name}")
        pairs.append((baseline, project))
    if len(pairs) != 4:
        raise RuntimeError(f"Expected four isolated projects, found {len(pairs)}")
    return pairs


def source_hashes(graph: dict[str, Any], project_root: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for source in graph.get("source_registry", []):
        path = project_source_path(project_root, source.get("source_file"))
        if not path.is_file():
            raise RuntimeError(f"Registered source missing: {path}")
        actual = sha256(path)
        expected = str(source.get("sha256", "")).lower()
        if expected and actual.lower() != expected:
            raise RuntimeError(f"Frozen source hash mismatch before generation: {path}")
        result[str(path.resolve())] = actual
    return result


def project_input_hashes(project_root: Path) -> dict[str, str]:
    """Hash every project input except the isolated answer key.

    The expected_ground_truth tree is deliberately neither entered nor read at
    generation time.  All other project files, including the brief and any
    unregistered hidden input, are protected so a newly created or modified
    file cannot evade the before/after integrity gate.
    """
    project_root = project_root.resolve()
    result: dict[str, str] = {}
    for root, directories, filenames in os.walk(project_root):
        root_path = Path(root)
        safe_directories = []
        for name in directories:
            if name.lower() == "expected_ground_truth":
                continue
            resolved_directory = (root_path / name).resolve()
            try:
                relative_directory = resolved_directory.relative_to(project_root)
            except ValueError as exc:
                raise RuntimeError(f"Project input directory escapes project root: {root_path / name}") from exc
            if any(part.lower() == "expected_ground_truth" for part in relative_directory.parts):
                continue
            safe_directories.append(name)
        directories[:] = safe_directories
        for filename in filenames:
            path = root_path / filename
            if path.is_file():
                resolved_file = path.resolve()
                try:
                    relative_file = resolved_file.relative_to(project_root)
                except ValueError as exc:
                    raise RuntimeError(f"Project input file escapes project root: {path}") from exc
                if any(part.lower() == "expected_ground_truth" for part in relative_file.parts):
                    raise RuntimeError(f"Project input file aliases expected_ground_truth: {path}")
                result[str(path.resolve())] = sha256(path)
    return result


def structured_fields_by_source(slides: list[dict[str, Any]]) -> dict[str, set[str]]:
    result: dict[str, set[str]] = {}
    for slide in slides:
        for visual in slide.get("visual_specs", []):
            for binding in visual.get("source_bindings", []):
                result.setdefault(binding["source_id"], set()).update(binding.get("fields_used", []))
    return result


def source_bound_presentation_rows(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Index canonical VisualSpec rows with their exact source contract.

    This is not a data rediscovery layer: it only exposes rows already present
    in the frozen v2.4 VisualSpecs, retaining source, field, filter,
    aggregation, row-key and data-contract identity.
    """
    result: list[dict[str, Any]] = []
    for slide in slides:
        for visual in slide.get("visual_specs", []):
            rows = visual.get("payload", {}).get("rows")
            if not isinstance(rows, list) or not rows or not isinstance(rows[0], dict):
                continue
            for binding in visual.get("source_bindings", []):
                result.append({
                    "source_id": binding["source_id"],
                    "source_file": binding.get("source_file", ""),
                    "fields_used": sorted(set(binding.get("fields_used", []))),
                    "row_filter": " ".join(str(binding.get("row_filter", "")).lower().split()),
                    "aggregation": " ".join(str(binding.get("aggregation", "none")).lower().split()),
                    "data_contract_id": visual.get("data_contract_id"),
                    "source_visual_id": visual.get("visual_id"),
                    "rows": copy.deepcopy(rows),
                    "row_keys": copy.deepcopy(visual.get("expected_row_keys", [])),
                    "expected_row_count": visual.get("expected_row_count"),
                })
    return result


def add_funnel_payload(visual: dict[str, Any], candidates: list[dict[str, Any]]) -> bool:
    required = {"odds_ratio", "ci_low", "ci_high", "study_label"}
    matched: list[dict[str, Any]] = []
    for binding in visual.get("source_bindings", []):
        target_filter = " ".join(str(binding.get("row_filter", "")).lower().split())
        target_aggregation = " ".join(str(binding.get("aggregation", "none")).lower().split())
        for candidate in candidates:
            if candidate["source_id"] != binding["source_id"]:
                continue
            if candidate["row_filter"] != target_filter or candidate["aggregation"] != target_aggregation:
                continue
            if not required <= set(candidate["fields_used"]):
                continue
            if int(candidate.get("expected_row_count") or -1) != len(candidate["rows"]):
                continue
            if len(candidate.get("row_keys", [])) != len(candidate["rows"]):
                continue
            if not all(required <= set(row) for row in candidate["rows"]):
                continue
            matched.append(candidate)
    if not matched:
        return False
    # Renderer row keys are VisualSpec-local and can legitimately differ when
    # two canonical visuals present the same source rows.  Dataset identity is
    # therefore the source-bound canonical row content, not a local object key.
    unique = {v25.canonical_json_hash(item["rows"]): item for item in matched}
    if len(unique) != 1:
        raise RuntimeError(f"Ambiguous source-bound funnel transform for {visual.get('visual_id')}")
    selected = next(iter(unique.values()))
    rows = selected["rows"]
    converted = []
    aggregate_row_keys = []
    for index, row in enumerate(rows):
        estimate = float(row["odds_ratio"])
        low = float(row["ci_low"])
        high = float(row["ci_high"])
        if min(estimate, low, high) <= 0 or not low < estimate < high:
            raise RuntimeError(f"Invalid study-level estimate for funnel transform in {visual.get('visual_id')}")
        converted.append({
            "study_label": row.get("study_label") or row.get("study_id"),
            "log_effect": math.log(estimate),
            "standard_error": (math.log(high) - math.log(low)) / (2 * 1.96),
        })
        aggregate_row_keys.append(
            v25.stable(
                "AGGROW",
                selected["source_id"],
                selected["row_filter"],
                index,
                row.get("study_id") or row.get("study_label") or v25.canonical_json_hash(row),
                n=12,
            )
        )
    if not converted:
        return False
    visual["presentation_payload"] = {
        "rows": converted,
        "x_field": "log_effect",
        "y_field": "standard_error",
        "label_field": "study_label",
        "aggregation": "study-level effect and CI transformed to log(OR) and SE; no model refit",
        "funnel_semantics": {
            "effect_measure": "odds_ratio",
            "transformed_scale": "log(OR)",
            "null_reference": 0.0,
            "reference_kind": "null_effect_not_pooled",
            "guide_formula": "log(OR) = 0 ± 1.96 × SE",
            "guide_interpretation": "null-centered pseudo 95% limits; not a pooled-effect reference",
            "derivation": "odds-ratio null of 1 transformed to log scale; deterministic chart semantics",
        },
        "aggregate_row_keys": aggregate_row_keys,
        "source_transform_contract": {
            "source_id": selected["source_id"],
            "source_file": selected["source_file"],
            "source_visual_id": selected["source_visual_id"],
            "source_data_contract_id": selected["data_contract_id"],
            "fields_used": ["study_label", "odds_ratio", "ci_low", "ci_high"],
            "row_filter": selected["row_filter"],
            "aggregation": "deterministic log(OR) and SE transform; no model refit",
            "source_row_count": len(rows),
            "source_row_keys": aggregate_row_keys,
        },
    }
    visual["native_chart_privacy"] = {
        "data_scope": "aggregate_study_level",
        "embedded_row_count": len(converted),
        "embedded_fields": ["study_label", "log_effect", "standard_error"],
        "patient_level_data_embedded": False,
        "cell_level_data_embedded": False,
        "absolute_source_path_embedded": False,
        "unrelated_fields_embedded": False,
        "approved_aggregate": True,
        "validated_fail_closed": True,
    }
    return True


def validate_native_chart_rows(visual: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    render_type = str(visual.get("render_visual_type") or visual.get("visual_type"))
    if render_type == "funnel_plot":
        rows = visual.get("presentation_payload", {}).get("rows", [])
    elif render_type == "weight_histogram":
        rows = visual.get("payload", {}).get("bins", [])
    else:
        rows = visual.get("payload", {}).get("aggregate_rows", visual.get("payload", {}).get("rows", []))
    if not isinstance(rows, list) or not rows or not all(isinstance(row, dict) for row in rows):
        raise RuntimeError(f"Native chart requires non-empty aggregate rows: {visual.get('visual_id')}")
    if len(rows) > 500:
        raise RuntimeError(f"Native chart aggregate row budget exceeded: {visual.get('visual_id')} has {len(rows)} rows")
    fields = sorted({str(key) for row in rows for key in row})
    allowed = NATIVE_FIELD_ALLOWLIST.get(render_type)
    if allowed is None:
        allowed = {"category", "label", "series", "value", "x", "y", "time", "group"}
        if not visual.get("native_chart_privacy", {}).get("approved_aggregate"):
            raise RuntimeError(f"Generic native chart lacks explicit aggregate approval: {visual.get('visual_id')}")
    unexpected = set(fields) - allowed
    if unexpected:
        raise RuntimeError(f"Native chart contains fields outside the {render_type} allowlist: {sorted(unexpected)}")
    sensitive = [field for field in fields if SENSITIVE_ROW_FIELDS.search(field)]
    if sensitive:
        raise RuntimeError(f"Native chart contains row identifiers: {sensitive}")
    for row in rows:
        for value in row.values():
            if isinstance(value, (dict, list, tuple, set)) or value is None:
                raise RuntimeError(f"Native chart contains a non-scalar or null value: {visual.get('visual_id')}")
            if isinstance(value, str) and ABSOLUTE_PATH_VALUE.search(value.strip()):
                raise RuntimeError(f"Native chart contains an absolute path value: {visual.get('visual_id')}")
    validate_native_chart_value_contract(render_type, rows, visual.get("visual_id"))
    return rows, fields


def validate_native_chart_value_contract(
    render_type: str,
    rows: list[dict[str, Any]],
    visual_id: object,
) -> None:
    """Validate actual workbook values, not only their declared field names."""

    def number(row: dict[str, Any], field: str) -> float:
        value = row.get(field)
        if isinstance(value, bool):
            raise RuntimeError(f"Native chart boolean is invalid for numeric field {field}: {visual_id}")
        try:
            parsed = float(value)
        except (TypeError, ValueError) as exc:
            raise RuntimeError(f"Native chart non-numeric value in {field}: {visual_id}") from exc
        if not math.isfinite(parsed):
            raise RuntimeError(f"Native chart non-finite value in {field}: {visual_id}")
        return parsed

    def category(row: dict[str, Any], field: str) -> str:
        value = row.get(field)
        if not isinstance(value, str):
            raise RuntimeError(f"Native chart category {field} must be text: {visual_id}")
        text = value.strip()
        if not text or len(text) > 160 or any(char in text for char in "\r\n\t\x00"):
            raise RuntimeError(f"Native chart category {field} is invalid: {visual_id}")
        if ABSOLUTE_PATH_VALUE.search(text) or SENSITIVE_CATEGORY_VALUE.fullmatch(text):
            raise RuntimeError(f"Native chart category {field} may contain a local path or row identifier: {visual_id}")
        return text

    for row in rows:
        if render_type == "weight_histogram":
            low = number(row, "bin_low")
            high = number(row, "bin_high")
            count = number(row, "count")
            if high <= low or count < 0:
                raise RuntimeError(f"Native histogram bin/count contract failed: {visual_id}")
        elif render_type == "grouped_composition":
            category(row, "group")
            category(row, "cell_type")
            cell_count = number(row, "cell_count")
            percentage = number(row, "percentage")
            if cell_count < 0 or not 0 <= percentage <= 100:
                raise RuntimeError(f"Native grouped-composition range contract failed: {visual_id}")
        elif render_type == "funnel_plot":
            category(row, "study_label")
            number(row, "log_effect")
            if number(row, "standard_error") <= 0:
                raise RuntimeError(f"Native funnel standard_error must be positive: {visual_id}")


def add_native_chart_contracts(slides: list[dict[str, Any]], candidates: list[dict[str, Any]]) -> None:
    for slide in slides:
        for visual in slide.get("visual_specs", []):
            if visual.get("render_visual_type") == "funnel_plot":
                if not add_funnel_payload(visual, candidates):
                    visual["render_visual_type"] = visual["visual_type"]
                    visual["render_mode"] = "source_figure"
                    visual["figure_rebuild_decision"] = "PRESERVE_WITH_CALLOUT"
            if visual.get("render_mode") != "native_chart":
                continue
            rows, fields = validate_native_chart_rows(visual)
            existing = visual.get("native_chart_privacy", {})
            visual["native_chart_privacy"] = {
                **existing,
                "data_scope": existing.get("data_scope", "aggregate_visual_spec_only"),
                "embedded_row_count": len(rows),
                "embedded_fields": fields,
                "patient_level_data_embedded": False,
                "cell_level_data_embedded": False,
                "absolute_source_path_embedded": False,
                "unrelated_fields_embedded": False,
                "approved_aggregate": True,
                "validated_fail_closed": True,
            }


def figure_decision_rows(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for slide in slides:
        for visual in slide.get("visual_specs", []):
            if visual.get("visual_type") not in {"source_figure", "umap_figure"} and not visual.get("figure_rebuild_decision"):
                continue
            rows.append({
                "slide_id": slide["slide_id"],
                "visual_id": visual["visual_id"],
                "scientific_role": visual.get("scientific_role", ""),
                "source_ids": "|".join(binding["source_id"] for binding in visual.get("source_bindings", [])),
                "decision": visual.get("figure_rebuild_decision", ""),
                "render_visual_type": visual.get("render_visual_type", visual.get("visual_type", "")),
                "render_mode": visual.get("render_mode", ""),
                "rationale": {
                    "REDRAW_FROM_STRUCTURED_DATA": "Registered aggregate structured rows support a traceable redraw.",
                    "PRESERVE_WITH_CALLOUT": "Source-only figure retained; no OCR-derived canonical data.",
                    "PRESERVE_SOURCE_FIGURE": "Source-only scientific figure retained unchanged.",
                    "MANUAL_REVIEW_REQUIRED": "No safe automatic reconstruction path.",
                }.get(visual.get("figure_rebuild_decision"), "Not a source-figure route."),
            })
    return rows


def presentation_manifest(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for index, slide in enumerate(slides, 1):
        result.append({
            "slide_number": index,
            "slide_id": slide["slide_id"],
            "slide_role": slide["slide_role"],
            "layout_family": slide["layout_family"],
            "layout_variant": slide["layout_variant"],
            "background_variant": slide["background_variant"],
            "visual_intensity": slide["visual_intensity"],
            "density_target": slide["density_target"],
            "hero_enabled": slide["hero_visual"]["enabled"],
            "hero_area_ratio": slide["hero_visual"]["target_area_ratio"],
            "render_modes": "|".join(visual["render_mode"] for visual in slide.get("visual_specs", [])),
            "annotation_count": sum(len(visual.get("annotation_specs", [])) for visual in slide.get("visual_specs", [])),
        })
    return result


class RendererWorker:
    """Short-lived offline batch worker; not a service or scientific authority."""

    def __init__(self, node: str, node_modules: Path | None) -> None:
        env = dict(os.environ)
        if node_modules:
            env["NODE_PATH"] = str(node_modules.resolve())
        self.process = subprocess.Popen(
            [node, str(REPO / "scripts" / "v2_5" / "render_v2_5_worker.mjs")],
            cwd=REPO,
            env=env,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        self.job_counter = 0

    def render(self, graph_path: Path, pptx: Path, log_path: Path) -> None:
        if self.process.poll() is not None or not self.process.stdin or not self.process.stdout:
            raise RuntimeError("v2.5 renderer worker is unavailable")
        self.job_counter += 1
        job_id = f"render-{self.job_counter:03d}"
        payload = {"job_id": job_id, "graph_path": str(graph_path.resolve()), "output_path": str(pptx.resolve())}
        self.process.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
        self.process.stdin.flush()
        output_lines: list[str] = []
        result: dict[str, Any] | None = None
        while True:
            line = self.process.stdout.readline()
            if line == "":
                break
            output_lines.append(line)
            if line.startswith("V25_WORKER_RESULT="):
                result = json.loads(line.split("=", 1)[1])
                if result.get("job_id") == job_id:
                    break
        log_path.write_text("".join(output_lines), encoding="utf-8")
        if not result or not result.get("ok") or not pptx.is_file() or pptx.stat().st_size == 0:
            detail = (result or {}).get("error") or "worker exited before a matching result"
            raise RuntimeError(f"v2.5 renderer failed: {detail}")

    def close(self) -> None:
        if self.process.stdin and not self.process.stdin.closed:
            self.process.stdin.close()
        try:
            return_code = self.process.wait(timeout=30)
        except subprocess.TimeoutExpired:
            self.process.terminate()
            return_code = self.process.wait(timeout=10)
        if return_code != 0:
            raise RuntimeError(f"v2.5 renderer worker exited with code {return_code}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Academic PPT Workflow v2.5 presentation-layer runner")
    parser.add_argument("--baseline-run", type=Path, required=True, help="Frozen v2.4 run root")
    parser.add_argument("--batch-regression-suite", type=Path, required=True, help="Four-project isolated suite")
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--node", required=True)
    parser.add_argument("--node-modules", type=Path)
    parser.add_argument("--frozen-manifest", type=Path, default=REPO / "audit" / "v2_5" / "frozen_v2_4_manifest.json")
    args = parser.parse_args()

    baseline_root = args.baseline_run.resolve()
    suite_root = args.batch_regression_suite.resolve()
    frozen_manifest_path = args.frozen_manifest.resolve()
    if not frozen_manifest_path.is_file():
        raise RuntimeError(f"Missing pre-registered v2.4 freeze manifest: {frozen_manifest_path}")
    frozen_manifest = json.loads(frozen_manifest_path.read_text(encoding="utf-8"))
    if Path(frozen_manifest.get("baseline_run", "")).resolve() != baseline_root:
        raise RuntimeError("Requested baseline-run is not the pre-registered frozen v2.4 run")
    if frozen_manifest.get("backend") != "native_pptxgenjs" or frozen_manifest.get("external_skill_used"):
        raise RuntimeError("Frozen manifest backend or external-skill contract is invalid")
    for relative, expected_hash in frozen_manifest.get("v2_4_module_sha256", {}).items():
        path = REPO / relative
        if not path.is_file() or sha256(path).lower() != str(expected_hash).lower():
            raise RuntimeError(f"Frozen v2.4 module hash mismatch: {relative}")
    run_root = REPO / "benchmark" / "v2_5" / "runs" / args.run_id
    stage_root = REPO / "staging" / "v2_5" / args.run_id
    if run_root.exists() or stage_root.exists():
        raise RuntimeError(f"Refusing to overwrite existing v2.5 run: {args.run_id}")
    run_root.mkdir(parents=True)
    stage_root.mkdir(parents=True)

    pairs = baseline_projects(baseline_root, suite_root)
    hashes_before: dict[str, str] = {}
    deck_index: list[dict[str, Any]] = []
    renderer_worker = RendererWorker(args.node, args.node_modules)
    for baseline_dir, project_root in pairs:
        started = time.perf_counter()
        output = run_root / project_root.name
        stage = stage_root / project_root.name
        output.mkdir()
        stage.mkdir()
        baseline_graph_path = baseline_dir / "canonical_object_graph.json"
        baseline_graph = json.loads(baseline_graph_path.read_text(encoding="utf-8"))
        registered = frozen_manifest.get("projects", {}).get(project_root.name)
        if not registered:
            raise RuntimeError(f"Project is absent from the frozen v2.4 manifest: {project_root.name}")
        if sha256(baseline_graph_path).lower() != str(registered.get("canonical_graph_sha256", "")).lower():
            raise RuntimeError(f"Frozen canonical graph SHA mismatch for {project_root.name}")
        if baseline_graph.get("backend") != "native_pptxgenjs" or baseline_graph.get("external_skill_used"):
            raise RuntimeError(f"Frozen baseline backend contract failed for {project_root.name}")
        registered_hashes = source_hashes(baseline_graph, project_root)
        actual_project_hashes = project_input_hashes(project_root)
        if not set(registered_hashes) <= set(actual_project_hashes):
            raise RuntimeError(f"Registered source inventory escaped project input scope for {project_root.name}")
        hashes_before.update(actual_project_hashes)
        frozen_identity = v25.scientific_contract_identity(baseline_graph)
        if frozen_identity.lower() != str(registered.get("scientific_contract_hash", "")).lower():
            raise RuntimeError(f"Pre-registered scientific contract mismatch for {project_root.name}")

        base_slides = copy.deepcopy(baseline_graph["slide_specs"])
        all_visuals = [visual for slide in base_slides for visual in slide.get("visual_specs", [])]
        brief = copy.deepcopy(baseline_graph["brief"])
        art = v25.build_art_direction_spec(
            brief,
            presentation_type=brief.get("presentation_type"),
            narrative_mode=baseline_graph.get("story_graph", {}).get("narrative_mode"),
            domain=baseline_graph["kind"],
            slide_count=len(base_slides),
            available_visuals=[visual["visual_type"] for visual in all_visuals],
        )
        rhythm = v25.build_deck_rhythm_plan(base_slides, art)
        fields_by_source = structured_fields_by_source(base_slides)
        presentation_rows = source_bound_presentation_rows(base_slides)
        slides = v25.plan_slide_presentation(base_slides, art, rhythm, fields_by_source)
        add_native_chart_contracts(slides, presentation_rows)
        validation_errors = v25.validate_presentation_specs(slides, art, rhythm)
        if validation_errors:
            dump_json(output / "presentation_validation_errors.json", validation_errors)
            raise RuntimeError(f"Presentation contract validation failed for {project_root.name}: {validation_errors[:5]}")

        graph = copy.deepcopy(baseline_graph)
        graph.update({
            "graph_version": "2.5",
            "slide_specs": slides,
            "art_direction_spec": art,
            "deck_rhythm_plan": rhythm,
            "frozen_v24_scientific_contract_hash": frozen_identity,
            "v24_baseline_graph_path": str(baseline_graph_path.resolve()),
            "backend": "native_pptxgenjs",
            "external_skill_used": False,
            "ground_truth_used_for_generation": False,
            "presentation_layer_only": True,
        })
        current_identity = v25.scientific_contract_identity(graph)
        if current_identity != frozen_identity:
            raise RuntimeError(f"Critical scientific contract regression for {project_root.name}")

        coverage = V24.coverage_contracts(brief, slides)
        if any(item["omitted"] for item in coverage):
            raise RuntimeError(f"Frozen narrative coverage failed for {project_root.name}")
        evidence = graph["canonical_evidence_graph"]
        claims = evidence.get("claims", [])
        conflicts = evidence.get("conflicts", [])
        unresolved = evidence.get("unresolved", [])

        write_csv(output / "source_manifest.csv", graph.get("source_registry", []))
        write_csv(output / "evidence_registry.csv", claims)
        write_csv(output / "conflict_registry.csv", conflicts)
        write_csv(output / "unresolved_registry.csv", unresolved)
        dump_json(output / "canonical_object_graph.json", graph)
        dump_json(output / "slide_spec.json", slides)
        dump_json(output / "visual_spec.json", [visual for slide in slides for visual in slide.get("visual_specs", [])])
        dump_json(output / "art_direction_spec.json", art)
        dump_json(output / "deck_rhythm_plan.json", rhythm)
        dump_json(output / "story_graph.json", graph.get("story_graph", {}))
        dump_json(output / "frozen_contract_identity.json", {"v24": frozen_identity, "v25": current_identity, "identical": True})
        V24.derived_artifacts(output, slides, claims, conflicts, unresolved, coverage)
        write_csv(output / "presentation_manifest.csv", presentation_manifest(slides))
        write_csv(output / "figure_rebuild_decisions.csv", figure_decision_rows(slides), ["slide_id", "visual_id", "scientific_role", "source_ids", "decision", "render_visual_type", "render_mode", "rationale"])
        completed_reviews = v25.read_completed_human_reviews([baseline_dir / "manual_visual_review_form.csv"])
        write_csv(output / "completed_human_reviews.csv", completed_reviews, ["anonymous_deck_code", "total_100", "reviewer", "review_date", "comments", "review_path"])
        # HumanVisualFeedbackRegistry is intentionally empty until a human
        # submits slide-level feedback.  The workflow never infers or fills it.
        write_csv(
            output / "human_visual_feedback_registry.csv",
            [],
            ["slide_id", "issue_type", "severity", "comment", "recommended_action", "accepted_rejected", "reviewer", "review_date", "source_review_path"],
        )
        write_csv(
            output / "design_regression_fixture_registry.csv",
            [],
            ["fixture_id", "issue_type", "source_slide_role", "generic_contract", "accepted_rejected", "test_path"],
        )
        code = v25.stable("DECK", args.run_id, project_root.name, n=6)
        write_csv(output / "manual_visual_review_form.csv", [v25.empty_human_review_form(code)])

        stage_graph = stage / "canonical_object_graph.json"
        dump_json(stage_graph, graph)
        pptx = output / f"{project_root.name}_v2_5.pptx"
        renderer_worker.render(stage_graph, pptx, output / "generation.log")
        manifest_path = pptx.with_suffix(".object_manifest.json")
        if not manifest_path.is_file():
            raise RuntimeError(f"Missing v2.5 object manifest for {project_root.name}")
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        actual = {item["visual_id"]: item for item in manifest}
        row_validation = []
        for slide in slides:
            for visual in slide.get("visual_specs", []):
                item = actual.get(visual["visual_id"], {})
                rendered_keys = item.get("rendered_row_keys", [])
                passed = item.get("rendered_row_count") == visual["expected_row_count"] and set(rendered_keys) == set(visual["expected_row_keys"])
                row_validation.append({
                    "slide_id": slide["slide_id"],
                    "visual_id": visual["visual_id"],
                    "visual_type": visual["visual_type"],
                    "render_mode": visual["render_mode"],
                    "expected_row_count": visual["expected_row_count"],
                    "rendered_row_count": item.get("rendered_row_count"),
                    "status": "PASS" if passed else "FAIL",
                })
        write_csv(output / "rendered_row_validation.csv", row_validation)
        if any(row["status"] != "PASS" for row in row_validation):
            raise RuntimeError(f"Rendered row identity regression for {project_root.name}")

        diversity = v25.visual_diversity_metrics(slides)
        dump_json(output / "visual_diversity_metrics.json", diversity)
        deck_index.append({
            "project_key": project_root.name,
            "arm": "V25",
            "kind": graph["kind"],
            "pptx_path": str(pptx.resolve()),
            "output_dir": str(output.resolve()),
            "slide_count": len(slides),
            "target_slide_count": brief.get("target_slide_count"),
            "hero_slide_count": rhythm["hero_slide_count"],
            "background_variant_count": diversity["background_variant_count"],
            "composition_variant_count": diversity["composition_variant_count"],
            "native_chart_count": sum(visual.get("render_mode") == "native_chart" for slide in slides for visual in slide.get("visual_specs", [])),
            "scientific_contract_identity": current_identity,
            "generation_seconds": round(time.perf_counter() - started, 3),
        })

    renderer_worker.close()

    hashes_after: dict[str, str] = {}
    for baseline_dir, project_root in pairs:
        graph = json.loads((baseline_dir / "canonical_object_graph.json").read_text(encoding="utf-8"))
        # Revalidate every registered source and then compare the complete
        # non-ground-truth project inventory against the pre-generation set.
        source_hashes(graph, project_root)
        hashes_after.update(project_input_hashes(project_root))
    integrity_errors = v25.verify_hash_manifest(hashes_before, hashes_after)
    write_csv(run_root / "input_hashes_before.csv", [{"path": path, "sha256": value} for path, value in sorted(hashes_before.items())])
    write_csv(run_root / "input_hashes_after.csv", [{"path": path, "sha256": value} for path, value in sorted(hashes_after.items())])
    dump_json(run_root / "input_integrity.json", {
        "protected_file_count": len(hashes_before),
        "ground_truth_excluded_from_generation": True,
        "errors": integrity_errors,
        "unchanged": not integrity_errors,
    })
    if integrity_errors:
        raise RuntimeError(f"Input integrity failure: {integrity_errors[:3]}")
    write_csv(run_root / "deck_index.csv", deck_index)
    dump_json(run_root / "generation_complete.json", {
        "run_id": args.run_id,
        "completed_at_utc": datetime.now(timezone.utc).isoformat(),
        "project_count": len(deck_index),
        "baseline_run": str(baseline_root),
        "input_mode": "batch_regression_suite",
        "presentation_layer_only": True,
        "ground_truth_used_for_generation": False,
        "external_skill_used": False,
        "backend": "native_pptxgenjs",
    })
    print(f"PROJECTS={len(deck_index)}")
    print(f"RUN_ROOT={run_root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
