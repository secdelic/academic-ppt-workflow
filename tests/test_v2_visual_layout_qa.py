from __future__ import annotations

import unittest

from scripts.academic_ppt.visual_layout_qa import inspect_text_geometry


def _shape(
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    bound_width: float,
    bound_height: float,
    footer_like: bool = False,
) -> dict:
    return {
        "left_pt": left,
        "top_pt": top,
        "width_pt": width,
        "height_pt": height,
        "bound_width_pt": bound_width,
        "bound_height_pt": bound_height,
        "rotation_degrees": 0,
        "text_length": 20,
        "footer_like": footer_like,
    }


class VisualLayoutQATests(unittest.TestCase):
    def test_non_overlapping_text_passes(self) -> None:
        report = {
            "status": "PASS",
            "slides": [
                {
                    "slide_index": 1,
                    "text_shapes": [
                        _shape(
                            left=50,
                            top=40,
                            width=800,
                            height=80,
                            bound_width=500,
                            bound_height=40,
                        ),
                        _shape(
                            left=80,
                            top=180,
                            width=760,
                            height=250,
                            bound_width=600,
                            bound_height=120,
                        ),
                    ],
                }
            ],
        }
        self.assertEqual(inspect_text_geometry(report), [])

    def test_overflow_and_overlap_are_reported_without_text(self) -> None:
        report = {
            "status": "PASS",
            "slides": [
                {
                    "slide_index": 3,
                    "text_shapes": [
                        _shape(
                            left=100,
                            top=60,
                            width=600,
                            height=80,
                            bound_width=610,
                            bound_height=110,
                        ),
                        _shape(
                            left=120,
                            top=110,
                            width=600,
                            height=160,
                            bound_width=400,
                            bound_height=90,
                        ),
                    ],
                }
            ],
        }
        issues = inspect_text_geometry(report)
        self.assertTrue(any("overflow vertically" in issue for issue in issues))
        self.assertTrue(any("text boxes overlap" in issue for issue in issues))

    def test_footer_overlap_is_ignored(self) -> None:
        report = {
            "status": "PASS",
            "slides": [
                {
                    "slide_index": 1,
                    "text_shapes": [
                        _shape(
                            left=50,
                            top=490,
                            width=800,
                            height=30,
                            bound_width=700,
                            bound_height=20,
                            footer_like=True,
                        ),
                        _shape(
                            left=50,
                            top=495,
                            width=800,
                            height=30,
                            bound_width=700,
                            bound_height=20,
                            footer_like=True,
                        ),
                    ],
                }
            ],
        }
        self.assertEqual(inspect_text_geometry(report), [])


if __name__ == "__main__":
    unittest.main()
