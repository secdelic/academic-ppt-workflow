from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[2]
RENDERER = ROOT / "scripts" / "v2_5" / "render_v2_5_deck.mjs"


class CjkTypographyWrapContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = RENDERER.read_text(encoding="utf-8")

    def test_numeric_cjk_units_are_protected_from_wrap(self):
        self.assertIn("小时|分钟|天|周|月|年|个百分点", self.source)
        self.assertIn("bodyProtectedRanges", self.source)

    def test_body_text_uses_token_safe_manual_wrapping(self):
        self.assertIn("function wrapAudienceBodyText", self.source)
        self.assertIn("needsTokenSafeBodyWrap(item[1]", self.source)
        self.assertIn("needsTokenSafeBodyWrap(value", self.source)
        self.assertIn("? wrapAudienceBodyText", self.source)
        self.assertIn("wrap: tokenSafeWrap ? false : undefined", self.source)

    def test_scientific_ci_and_pp_typography_is_normalized(self):
        self.assertIn('.replace(/95\\s*%\\s*CI/gi, "95% CI")', self.source)
        self.assertNotIn("95%CI", self.source)
        self.assertIn('.replace(/^95% CI\\s+/, "95% CI\\n")', self.source)

    def test_punctuation_only_lines_are_rejoined(self):
        self.assertIn("Punctuation belongs with the preceding sentence", self.source)
        self.assertIn("lines[index - 1] += lines[index]", self.source)


if __name__ == "__main__":
    unittest.main()
