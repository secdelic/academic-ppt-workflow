from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.visual_density import analyze_visual_density


def _slide(index: int, layout: str, kind: str = "text"):
    return {
        "slide_id": f"SLD-{index}",
        "slide_role": "result",
        "layout_family": layout,
        "planned_geometry": [
            {
                "kind": kind,
                "role": "body" if kind == "text" else "visual",
                "text": "单句结论" if kind == "text" else "",
                "bounds": {"x": 1, "y": 2, "w": 4, "h": 1},
            }
        ],
    }


class DensityBudgetTests(unittest.TestCase):
    def test_three_sparse_text_slides_fail_low_density_and_streak(self):
        deck = {"slides": [_slide(1, "takeaway"), _slide(2, "takeaway"), _slide(3, "takeaway")]}
        _density, repetition, low, _overloaded, issues = analyze_visual_density(deck)
        self.assertEqual(3, len(low))
        self.assertTrue(any("more than two" in issue for issue in issues))
        self.assertEqual("FAIL", repetition[-1]["status"])

    def test_structured_visuals_break_text_only_streak(self):
        deck = {
            "slides": [
                _slide(1, "timeline", "diagram"),
                _slide(2, "forest_plot", "chart"),
                _slide(3, "matrix_2x2", "diagram"),
            ]
        }
        _density, repetition, _low, _overloaded, issues = analyze_visual_density(deck)
        self.assertTrue(all(row["status"] == "PASS" for row in repetition))
        self.assertFalse(any("text-only" in issue for issue in issues))


if __name__ == "__main__":
    unittest.main()
