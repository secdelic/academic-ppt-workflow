from __future__ import annotations

import unittest


class CrossRendererContractTests(unittest.TestCase):
    def test_page_count_and_severe_issue_contract(self):
        powerpoint = {"pages": 12, "severe_issues": []}
        libreoffice = {
            "pages": 12,
            "severe_issues": [],
            "warnings": ["minor font substitution"],
        }
        compatible = (
            powerpoint["pages"] == libreoffice["pages"]
            and not powerpoint["severe_issues"]
            and not libreoffice["severe_issues"]
        )
        self.assertTrue(compatible)
        self.assertTrue(libreoffice["warnings"])


if __name__ == "__main__":
    unittest.main()
