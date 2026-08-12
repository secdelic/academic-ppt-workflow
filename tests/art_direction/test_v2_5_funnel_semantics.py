from __future__ import annotations

import importlib
import sys
import unittest
from pathlib import Path


REPO = Path(__file__).resolve().parents[2]
V25_DIR = REPO / "scripts" / "v2_5"
RENDERER = V25_DIR / "render_v2_5_deck.mjs"


class V25FunnelSemanticTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        sys.path.insert(0, str(V25_DIR))
        try:
            cls.runner = importlib.import_module("run_v2_5")
        finally:
            sys.path.pop(0)
        cls.renderer = RENDERER.read_text(encoding="utf-8")

    def test_source_bound_transform_declares_null_not_pooled_reference(self):
        visual = {
            "visual_id": "VIS-GENERIC-FUNNEL",
            "source_bindings": [{
                "source_id": "SRC-GENERIC-STUDIES",
                "row_filter": "all eligible studies",
                "aggregation": "none",
            }],
        }
        rows = [
            {"study_label": "Study A", "odds_ratio": 1.4, "ci_low": 1.1, "ci_high": 1.9},
            {"study_label": "Study B", "odds_ratio": 0.9, "ci_low": 0.6, "ci_high": 1.3},
        ]
        candidate = {
            "source_id": "SRC-GENERIC-STUDIES",
            "source_file": "input/data/studies.csv",
            "fields_used": ["study_label", "odds_ratio", "ci_low", "ci_high"],
            "row_filter": "all eligible studies",
            "aggregation": "none",
            "data_contract_id": "DATA-GENERIC-STUDIES",
            "source_visual_id": "VIS-GENERIC-FOREST",
            "rows": rows,
            "row_keys": ["ROW-A", "ROW-B"],
            "expected_row_count": 2,
        }

        self.assertTrue(self.runner.add_funnel_payload(visual, [candidate]))
        payload = visual["presentation_payload"]
        semantics = payload["funnel_semantics"]
        self.assertEqual(semantics["null_reference"], 0.0)
        self.assertEqual(semantics["reference_kind"], "null_effect_not_pooled")
        self.assertEqual(semantics["guide_formula"], "log(OR) = 0 ± 1.96 × SE")
        self.assertIn("not a pooled-effect reference", semantics["guide_interpretation"])
        self.assertNotIn("pooled", payload["source_transform_contract"]["fields_used"])
        self.assertEqual(visual["native_chart_privacy"]["embedded_fields"], ["study_label", "log_effect", "standard_error"])
        self.assertFalse(visual["native_chart_privacy"]["patient_level_data_embedded"])

    def test_renderer_fails_closed_and_uses_manual_plot_coordinates(self):
        contract = self.renderer.split("function funnelSemanticContract", 1)[1].split("function renderFunnelSemanticOverlay", 1)[0]
        overlay = self.renderer.split("function renderFunnelSemanticOverlay", 1)[1].split("function renderNativeChart", 1)[0]
        native = self.renderer.split("function renderNativeChart", 1)[1].split("function resolveSourceFigureTreatment", 1)[0]

        self.assertIn("requires an explicit funnel_semantics contract", contract)
        self.assertIn('reference_kind !== "null_effect_not_pooled"', contract)
        self.assertIn('fields.has(field)', contract)
        self.assertIn('chartOptions.layout = funnelPlotLayout', native)
        self.assertLess(native.index("slide.addChart"), native.index("renderFunnelSemanticOverlay"))
        self.assertIn('null-reference-log-or-zero', overlay)
        self.assertIn('null-centered-guide-left', overlay)
        self.assertIn('null-centered-guide-right', overlay)
        self.assertIn('1.96 * guideSe', overlay)
        self.assertIn('非合并效应', overlay)
        self.assertIn('pooledEffectUsed: false', overlay)
        self.assertIn('funnelSemanticOverlay: modeEvidence.funnelSemanticOverlay || null', self.renderer)
        self.assertIn('funnel_semantic_overlay: evidence.funnelSemanticOverlay || null', self.renderer)

    def test_overlay_annotations_use_measured_columns_and_gutters(self):
        overlay = self.renderer.split("function renderFunnelSemanticOverlay", 1)[1].split("function renderNativeChart", 1)[0]
        self.assertIn("const annotationGutter = 0.14", overlay)
        self.assertIn("const annotationLeft", overlay)
        self.assertIn("const annotationRight", overlay)
        self.assertIn("const guideLabelBounds", overlay)
        self.assertIn("const nullLabelBounds", overlay)
        self.assertIn("guideLabelBounds.x + guideLabelBounds.w + annotationGutter", overlay)
        self.assertIn("insufficient horizontal annotation width", overlay)
        self.assertIn("overlapping semantic annotation columns", overlay)
        self.assertIn("insufficient annotation-to-plot gutter", overlay)
        self.assertGreaterEqual(overlay.count("wrap: false"), 2)
        self.assertIn('align: "left"', overlay)
        self.assertIn('align: "right"', overlay)


if __name__ == "__main__":
    unittest.main()
