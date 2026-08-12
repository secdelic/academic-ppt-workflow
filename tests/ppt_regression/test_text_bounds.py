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
from academic_ppt.text_measurement import resolve_font


class TextBoundsTests(unittest.TestCase):
    def test_font_fallbacks_are_deterministic(self):
        installed = ["Arial", "Microsoft YaHei", "Noto Sans CJK SC"]
        self.assertEqual(
            "Arial", resolve_font("Aptos", ["Arial"], installed)
        )
        self.assertEqual(
            "Microsoft YaHei",
            resolve_font("微软雅黑", ["Microsoft YaHei", "Arial"], installed),
        )
        self.assertEqual(
            "Noto Sans CJK SC",
            resolve_font("思源黑体", ["Noto Sans CJK SC", "Arial"], installed),
        )
        self.assertEqual(
            "Arial", resolve_font("Missing Preferred", ["Arial"], installed)
        )

    def test_long_chinese_and_mixed_title_remains_in_title_zone(self):
        contract = LayoutContract.from_files(
            ROOT / "config" / "layout_contract.yaml",
            ROOT / "config" / "safe_zones.yaml",
        )
        titles = [
            "较长中文结论标题必须在标题安全区内合理换行且不侵入正文区域",
            "Long-term creatinine trajectory 与心室功能表型之间仅提示观察性关联",
        ]
        for index, title in enumerate(titles):
            slide = {
                "slide_id": f"SLD-TITLE-{index}",
                "slide_role": "result",
                "slide_title": title,
                "single_key_message": "正文消息",
                "visual_type": "takeaway",
            }
            objects = plan_slide_geometry(slide, contract)
            issues = validate_planned_geometry(objects, contract)
            self.assertFalse(
                any(issue["code"] == "ESTIMATED_TEXT_OVERFLOW" for issue in issues),
                issues,
            )
            title_object = next(
                item for item in objects if item["object_id"].endswith(":title")
            )
            self.assertGreaterEqual(title_object["font_size_pt"], 28)


if __name__ == "__main__":
    unittest.main()
