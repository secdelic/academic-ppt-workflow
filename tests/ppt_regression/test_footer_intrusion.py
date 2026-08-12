from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.visual_layout_qa import inspect_text_geometry


class FooterIntrusionTests(unittest.TestCase):
    def test_powerpoint_actual_bounds_detect_body_footer_overlap(self):
        report = {
            "status": "PASS",
            "slide_width_pt": 960,
            "slide_height_pt": 540,
            "slides": [
                {
                    "slide_index": 1,
                    "text_shapes": [
                        {
                            "text_length": 20,
                            "left_pt": 60,
                            "top_pt": 450,
                            "width_pt": 600,
                            "height_pt": 60,
                            "bound_left_pt": 60,
                            "bound_top_pt": 450,
                            "bound_width_pt": 600,
                            "bound_height_pt": 60,
                            "footer_like": False,
                        },
                        {
                            "text_length": 8,
                            "left_pt": 60,
                            "top_pt": 500,
                            "width_pt": 500,
                            "height_pt": 20,
                            "bound_left_pt": 60,
                            "bound_top_pt": 500,
                            "bound_width_pt": 500,
                            "bound_height_pt": 20,
                            "footer_like": True,
                        },
                    ],
                    "shapes": [],
                }
            ],
        }
        issues = inspect_text_geometry(report)
        self.assertTrue(any("footer" in issue.lower() for issue in issues), issues)

    def test_v2_fixture_preserves_expected_failure_evidence(self):
        fixture = (
            ROOT
            / "regression"
            / "fixtures"
            / "synthetic_cardio_aki_v2_0"
            / "powerpoint_layout_manifest.json"
        )
        self.assertTrue(fixture.is_file())
        report = json.loads(fixture.read_text(encoding="utf-8-sig"))
        self.assertIsInstance(report.get("slides"), list)


if __name__ == "__main__":
    unittest.main()
