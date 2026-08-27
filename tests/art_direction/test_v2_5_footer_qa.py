from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path


REPO = Path(__file__).resolve().parents[2]
FINALIZER = REPO / "scripts" / "v2_5" / "finalize_v2_5.py"


def load_finalizer():
    spec = importlib.util.spec_from_file_location("academic_ppt_v25_footer_qa", FINALIZER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {FINALIZER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class V25FooterQATests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qa = load_finalizer()
        cls.spec = {
            "title": "通用标题",
            "source_bindings": [{"short_label_zh": "队列汇总表"}],
        }

    def test_footer_band_wholly_inside_zone_is_allowed(self):
        band = {
            "has_text": False,
            "text": "",
            "left": 0,
            "top": 512.64,
            "width": 960,
            "height": 27.36,
        }
        self.assertTrue(self.qa.is_v25_legal_footer_shape(band, self.spec, 2, 960, 540))

    def test_body_shape_crossing_footer_remains_a_failure(self):
        body = {
            "has_text": True,
            "text": "正文不得进入页脚",
            "left": 72,
            "top": 490,
            "width": 600,
            "height": 40,
            "bound_top": 490,
            "bound_height": 40,
        }
        geometry = {"slide_width": 960, "slide_height": 540, "slides": [{"shapes": [body]}]}
        self.assertFalse(self.qa.is_v25_legal_footer_shape(body, self.spec, 1, 960, 540))
        self.assertEqual(self.qa.v25_footer_intrusion_count(geometry["slides"][0], self.spec, 1, geometry), 1)

    def test_footer_label_and_page_number_use_footer_threshold(self):
        label = {
            "has_text": True,
            "text": "队列汇总表",
            "left": 1.34,
            "top": 7.155,
            "width": 8.0,
            "height": 0.28,
            "sizes": [11.0],
        }
        page = {**label, "text": "2", "left": 12.4, "width": 0.4}
        self.assertEqual(self.qa.v25_typography_role(label, self.spec, 2), "footer")
        self.assertEqual(self.qa.v25_typography_role(page, self.spec, 2), "footer")

    def test_small_body_text_is_not_reclassified_as_footer(self):
        body = {
            "has_text": True,
            "text": "14 pt 正文负控",
            "left": 1.0,
            "top": 3.0,
            "width": 5.0,
            "height": 0.4,
            "sizes": [14.0],
        }
        self.assertEqual(self.qa.v25_typography_role(body, self.spec, 1), "chart_detail")
        self.assertFalse(self.qa.is_v25_legal_footer_shape(body, self.spec, 1, 13.333, 7.5))


if __name__ == "__main__":
    unittest.main()
