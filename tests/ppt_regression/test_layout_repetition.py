from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.visual_density import analyze_visual_density


class LayoutRepetitionTests(unittest.TestCase):
    def test_third_consecutive_layout_is_rejected(self):
        slides = []
        for index in range(3):
            slides.append(
                {
                    "slide_id": f"SLD-{index}",
                    "slide_role": "result",
                    "layout_family": "takeaway",
                    "planned_geometry": [
                        {
                            "kind": "diagram",
                            "role": "visual",
                            "bounds": {"x": 1, "y": 2, "w": 8, "h": 3},
                        }
                    ],
                }
            )
        _density, repetition, _low, _overloaded, issues = analyze_visual_density(
            {"slides": slides}
        )
        self.assertEqual("FAIL", repetition[-1]["status"])
        self.assertTrue(any("repeats 3" in issue for issue in issues))


if __name__ == "__main__":
    unittest.main()
