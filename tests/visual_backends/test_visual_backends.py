from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from extensions.visual_backends.registry import render_visual
from scripts.academic_ppt.visual_ir import (
    make_visual_ir,
    timeline_similarity,
    validate_visual_ir_schema,
    validate_visual_semantics,
)


REPO = Path(__file__).resolve().parents[2]


def visual(visual_type: str, **overrides):
    values = {
        "slide_key": f"fixture-{visual_type}",
        "visual_type": visual_type,
        "semantic_role": "result",
        "source_ids": ["SRC-FIXTURE0001"],
        "categories": ["A", "B", "C"],
        "wording_boundary": "Synthetic association only; do not imply causality.",
    }
    values.update(overrides)
    return make_visual_ir(**values)


def latest_rendered_run() -> Path:
    candidates = []
    for path in (REPO / "benchmark" / "v2_2" / "runs").glob("*"):
        if (
            (path / "generation_complete.json").is_file()
            and (path / "powerpoint_render_index.json").is_file()
            and (path / "libreoffice_render_index.json").is_file()
        ):
            candidates.append(path)
    if not candidates:
        raise AssertionError("No completed rendered v2.2 run")
    return sorted(candidates)[-1]


class VisualIRTests(unittest.TestCase):
    def test_visual_ir_schema(self):
        item = visual("timeline")
        self.assertEqual(validate_visual_ir_schema(item), [])

    def test_outcome_separation(self):
        item = visual(
            "forest_plot", estimate=[1.2, 1.3], ci_low=[1.0, 1.1],
            ci_high=[1.4, 1.5], series=[
                {"outcome": "mortality"}, {"outcome": "AKI"}
            ], reference_value=1,
        )
        codes = {x["code"] for x in validate_visual_semantics(item)}
        self.assertIn("OUTCOME_MIXED_WITHOUT_FACET", codes)

    def test_estimand_separation(self):
        item = visual(
            "forest_plot", estimate=[1.2, 0.1], ci_low=[1.0, 0.0],
            ci_high=[1.4, 0.2], series=[
                {"estimand": "OR"}, {"estimand": "RD"}
            ], reference_value=1,
        )
        codes = {x["code"] for x in validate_visual_semantics(item)}
        self.assertIn("ESTIMAND_MIXED_WITHOUT_LABEL", codes)

    def test_sensitivity_analysis_labels(self):
        item = visual(
            "sensitivity_plot", estimate=[0.1, 0.08], ci_low=[0.02, 0.01],
            ci_high=[0.18, 0.16], analysis_label=["Main analysis", "Complete case"],
            reference_value=0,
        )
        self.assertEqual(validate_visual_semantics(item), [])

    def test_subgroup_interaction_p(self):
        item = visual(
            "subgroup_plot", categories=["Prospective", "Retrospective"],
            estimate=[2.1, 1.9], ci_low=[1.4, 1.2], ci_high=[3.0, 2.8],
            interaction_p=[0.58, 0.58], reference_value=1,
            wording_boundary="Point estimates do not establish effect modification.",
        )
        self.assertEqual(validate_visual_semantics(item), [])

    def test_event_rate_denominator(self):
        item = visual(
            "event_rate_plot", categories=["A", "B"], estimate=[25.0, 50.0],
            numerator=[1, 2], denominator=[4, 4], sample_size=8,
            annotation_rule={"rate_decimals": 1},
        )
        self.assertEqual(validate_visual_semantics(item), [])

    def test_missingness_threshold(self):
        item = visual(
            "missingness_plot", categories=["BNP"], estimate=[24.1],
            numerator=[81], annotation_rule={"decision_threshold": 20},
            highlight_rule={"threshold_line": 20},
            wording_boundary="Missingness may still bias results.",
        )
        self.assertEqual(validate_visual_semantics(item), [])

    def test_timeline_semantic_duplicate(self):
        first = visual("timeline", categories=["Screen", "Baseline", "Follow-up"])
        second = visual("timeline", categories=["Screen", "Baseline", "Follow-up"])
        self.assertGreaterEqual(timeline_similarity(first, second), 0.82)

    def test_true_2x2_matrix(self):
        item = visual(
            "matrix_2x2", categories=["00", "01", "10", "11"],
            annotation_rule={"row_dimension": "LV", "column_dimension": "RV"},
        )
        self.assertEqual(validate_visual_semantics(item), [])

    def test_no_results_protocol(self):
        item = visual(
            "forest_plot", estimate=[1.2], ci_low=[1.0], ci_high=[1.4],
            reference_value=1,
        )
        codes = {x["code"] for x in validate_visual_semantics(item, protocol_without_results=True)}
        self.assertIn("PROTOCOL_RESULT_FORBIDDEN", codes)


class IsolationAndAdapterTests(unittest.TestCase):
    def test_no_cross_project_source_leak(self):
        run = latest_rendered_run()
        staging = REPO / "staging" / "v2_2" / run.name
        sets = []
        for manifest in sorted(staging.glob("*/source_manifest.csv")):
            with manifest.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
            sets.append({row["source_id"] for row in rows})
        self.assertEqual(len(sets), 4)
        for index, first in enumerate(sets):
            for second in sets[index + 1:]:
                self.assertFalse(first & second)

    def test_no_ground_truth_leak(self):
        run = latest_rendered_run()
        staging = REPO / "staging" / "v2_2" / run.name
        for manifest in staging.glob("*/source_manifest.csv"):
            self.assertNotIn("expected_ground_truth", manifest.read_text(encoding="utf-8-sig"))

    def test_no_style_cache_leak(self):
        run = latest_rendered_run()
        staging = REPO / "staging" / "v2_2" / run.name
        for status in staging.glob("*/style_status.json"):
            data = json.loads(status.read_text(encoding="utf-8"))
            self.assertEqual(data["status"], "NO_REFERENCE_STYLE")
            self.assertFalse(data["cache_used"])

    def test_svg_drawingml_editability(self):
        item = visual(
            "conceptual_framework", categories=["Problem", "Gap", "Objective"],
            preferred_backend="svg_drawingml_adapter",
            editability_requirement="vector_acceptable",
        )
        with tempfile.TemporaryDirectory(dir=REPO / "staging") as tmp:
            artifact = render_visual(item, Path(tmp), backend="svg_drawingml_adapter")
            self.assertFalse(artifact.fallback_used)
            self.assertTrue(any(obj.get("drawingml_translatable") for obj in artifact.object_manifest))

    def test_graphviz_text_editability(self):
        item = visual(
            "flow_diagram", categories=["Eligibility", "Clone", "Weight", "Estimate"],
            preferred_backend="graphviz_adapter",
        )
        with tempfile.TemporaryDirectory(dir=REPO / "staging") as tmp:
            artifact = render_visual(item, Path(tmp), backend="graphviz_adapter")
            self.assertFalse(artifact.fallback_used)
            nodes = [obj for obj in artifact.object_manifest if obj.get("kind") == "node"]
            self.assertEqual(len(nodes), 4)
            self.assertTrue(all(obj.get("native_shape_translation_available") for obj in nodes))

    def test_formula_fallback(self):
        item = visual(
            "formula", categories=["alpha", "power"],
            preferred_backend="latex_omml_adapter",
        )
        with tempfile.TemporaryDirectory(dir=REPO / "staging") as tmp:
            artifact = render_visual(item, Path(tmp), backend="latex_omml_adapter")
            self.assertEqual(artifact.editable_level, "native_after_injection")
            self.assertTrue(any(obj.get("kind") == "omml_equation" for obj in artifact.object_manifest))

    def test_powerpoint_render(self):
        rows = json.loads((latest_rendered_run() / "powerpoint_render_index.json").read_text(encoding="utf-8-sig"))
        self.assertEqual(len(rows), 16)
        self.assertTrue(all(row["powerpoint_open"] for row in rows))
        self.assertTrue(all(int(row["pptx_slide_count"]) == int(row["png_count"]) for row in rows))

    def test_libreoffice_render(self):
        rows = json.loads((latest_rendered_run() / "libreoffice_render_index.json").read_text(encoding="utf-8-sig"))
        self.assertEqual(len(rows), 16)
        self.assertTrue(all(int(row["exit_code"]) == 0 and int(row["pdf_page_count"]) > 0 for row in rows))

    def test_backend_fallback(self):
        item = visual("timeline", categories=["A", "B"])
        with tempfile.TemporaryDirectory(dir=REPO / "staging") as tmp:
            artifact = render_visual(item, Path(tmp), backend="unregistered")
            self.assertEqual(artifact.render_backend, "native_pptxgenjs")

    def test_input_hash_integrity(self):
        report = (REPO / "audit" / "v2_2" / "suite_input_hashes_after.csv")
        with report.open("r", encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle))
        self.assertTrue(rows)
        self.assertTrue(all(row["unchanged"].lower() == "true" for row in rows))


if __name__ == "__main__":
    unittest.main()
