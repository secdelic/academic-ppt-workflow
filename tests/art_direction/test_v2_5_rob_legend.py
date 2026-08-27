from __future__ import annotations

import re
import unittest
from pathlib import Path


REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "scripts" / "v2_5" / "render_v2_5_deck.mjs"


def renderer_source() -> str:
    return RENDERER.read_text(encoding="utf-8")


def function_source(source: str, name: str, next_name: str) -> str:
    match = re.search(
        rf"function\s+{re.escape(name)}\s*\([^)]*\)\s*\{{(?P<body>.*?)\n\}}\n\nfunction\s+{re.escape(next_name)}\s*\(",
        source,
        flags=re.DOTALL,
    )
    if not match:
        raise AssertionError(f"Unable to locate renderer function {name}")
    return match.group("body")


class RiskOfBiasLegendRegression(unittest.TestCase):
    def test_semantic_palette_is_single_fail_closed_contract(self) -> None:
        source = renderer_source()
        self.assertIn("const RISK_OF_BIAS_SEMANTICS", source)
        self.assertIn('Low: Object.freeze({ color: "4C956C"', source)
        self.assertIn('"Some concerns": Object.freeze({ color: "E9A23B"', source)
        self.assertIn('High: Object.freeze({ color: "C74B50"', source)
        helper = function_source(source, "riskOfBiasSemantic", "observedRiskOfBiasLegend")
        self.assertIn("throw new Error", helper)
        self.assertNotIn("|| RISK_OF_BIAS_SEMANTICS", helper)

    def test_study_characteristics_legend_uses_only_observed_bound_rows(self) -> None:
        source = renderer_source()
        helper = function_source(source, "observedRiskOfBiasLegend", "audienceAxisLabel")
        self.assertIn("rows || []", helper)
        self.assertIn("row.risk_of_bias", helper)
        self.assertIn("filter(category => observed.has(category))", helper)
        body = function_source(source, "renderStudyCharacteristics", "renderPooledPanel")
        self.assertIn("observedRiskOfBiasLegend(rows, visual.visual_id)", body)
        self.assertIn("riskOfBiasSemantic(row.risk_of_bias, visual.visual_id).color", body)
        self.assertIn("visual.visual_id}:rob-legend:${item.category}", body)
        self.assertIn("riskOfBiasLegend: robLegend.map", body)
        self.assertNotRegex(body, r"risk_of_bias\s*===\s*\"High\"")

    def test_overview_and_full_matrix_share_identical_color_semantics(self) -> None:
        source = renderer_source()
        matrix = function_source(source, "renderROB", "renderGrade")
        self.assertIn("riskOfBiasSemantic(item.judgment, visual.visual_id).color", matrix)
        self.assertIn("RISK_OF_BIAS_ORDER", matrix)
        self.assertIn("...rendererEvidence", source)
        self.assertIn("riskOfBiasLegend: modeEvidence.riskOfBiasLegend || null", source)
        self.assertIn("risk_of_bias_legend: evidence.riskOfBiasLegend || null", source)


if __name__ == "__main__":
    unittest.main()
