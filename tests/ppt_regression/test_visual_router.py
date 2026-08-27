from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.visual_router import route_content_type


class VisualRouterTests(unittest.TestCase):
    def test_six_structured_content_types_route_to_visuals(self):
        cases = [
            (
                "effect_estimates",
                {
                    "categories": ["A", "B", "C"],
                    "values": [1.1, 1.6, 2.1],
                    "confidence_interval": [
                        {"low": 0.8, "high": 1.5},
                        {"low": 1.1, "high": 2.2},
                        {"low": 1.4, "high": 3.0},
                    ],
                    "reference_value": 1,
                },
                "forest_plot",
            ),
            (
                "categorical_event_rates",
                {
                    "categories": ["A", "B", "C", "D"],
                    "values": [10, 20, 30, 40],
                    "denominator": [25, 25, 25, 25],
                },
                "event_rate_chart",
            ),
            (
                "longitudinal_measurements",
                {"categories": ["T0", "T1", "T2"], "values": [1, 2, 3]},
                "line_chart",
            ),
            (
                "missingness",
                {"categories": ["A", "B"], "values": [4, 24]},
                "missingness_chart",
            ),
            (
                "temporal_study_design",
                {"milestones": ["入组", "landmark", "结局"]},
                "timeline",
            ),
            (
                "two_binary_dimensions",
                {
                    "row_dimension": "L",
                    "column_dimension": "R",
                    "cells": ["00", "01", "10", "11"],
                },
                "matrix_2x2",
            ),
        ]
        for content_type, payload, expected in cases:
            with self.subTest(content_type=content_type):
                visual, missing = route_content_type(content_type, payload)
                self.assertEqual(expected, visual)
                self.assertEqual([], missing)

    def test_incomplete_forest_data_falls_back_without_invention(self):
        visual, missing = route_content_type(
            "effect_estimates",
            {"categories": ["A"], "values": [1.2]},
        )
        self.assertEqual("structured_table", visual)
        self.assertIn("confidence_interval", missing)


if __name__ == "__main__":
    unittest.main()
