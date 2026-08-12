from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.layout_contract import (
    LayoutContract,
    plan_slide_geometry,
    validate_planned_geometry,
)


class SafeZoneTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.contract = LayoutContract.from_files(
            ROOT / "config" / "layout_contract.yaml",
            ROOT / "config" / "safe_zones.yaml",
        )

    def test_generic_layouts_stay_out_of_footer(self):
        visuals = [
            "event_rate_chart",
            "forest_plot",
            "line_chart",
            "timeline",
            "matrix_2x2",
            "comparison",
            "takeaway",
            "audit_table",
            "native_flow",
        ]
        for visual in visuals:
            with self.subTest(visual=visual):
                slide = {
                    "slide_id": f"SLD-{visual}",
                    "slide_role": "result",
                    "slide_title": "通用安全区验证标题",
                    "single_key_message": "结构化内容必须保留在正文安全区内。",
                    "visual_type": visual,
                    "chart_caption": "N=336；来源于同一 ChartDataContract",
                }
                issues = validate_planned_geometry(
                    plan_slide_geometry(slide, self.contract), self.contract
                )
                self.assertEqual([], issues)

    def test_footer_intrusion_is_a_hard_failure(self):
        objects = [
            {
                "object_id": "SLD-X:body",
                "kind": "text",
                "role": "body",
                "zone": "slide",
                "bounds": {"x": 1.0, "y": 6.3, "w": 8.0, "h": 0.8},
                "text": "必须被识别的页脚侵入",
                "font_size_pt": 18,
            }
        ]
        codes = {
            issue["code"]
            for issue in validate_planned_geometry(objects, self.contract)
        }
        self.assertIn("FOOTER_INTRUSION", codes)

    def test_no_project_or_page_specific_layout_hardcoding(self):
        modules = [
            ROOT / "scripts" / "academic_ppt" / "layout_contract.py",
            ROOT / "scripts" / "academic_ppt" / "visual_router.py",
            ROOT / "scripts" / "generate_deck_pptxgen.mjs",
        ]
        joined = "\n".join(path.read_text(encoding="utf-8") for path in modules)
        self.assertNotIn("SYN_CARDIO_AKI_SHADOW_VALIDATION", joined)
        self.assertNotRegex(joined, r"if\s*\(\s*slide_number\s*==")


if __name__ == "__main__":
    unittest.main()
