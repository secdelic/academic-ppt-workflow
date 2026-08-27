from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.chart_data import contract_from_records, validate_chart_contract
from academic_ppt.visual_router import chart_caption


class ChartMetadataTests(unittest.TestCase):
    def test_numerator_denominator_drive_one_shared_metadata_contract(self):
        contract = contract_from_records(
            source_id="SRC-TEST",
            source_path="input/data/cohort.csv",
            records=[
                {"group": "A", "events": "10", "n": "40", "rate_percent": "25.0"},
                {"group": "B", "events": "18", "n": "60", "rate_percent": "30.0"},
            ],
        )
        self.assertIsNotNone(contract)
        value = contract.to_dict()
        self.assertEqual(100.0, value["sample_size_total"])
        self.assertEqual({"A": 40.0, "B": 60.0}, value["sample_size_by_group"])
        self.assertEqual([], validate_chart_contract(value))
        caption = chart_caption(value)
        self.assertIn("N=100", caption)
        self.assertNotIn("not provided", caption.lower())


if __name__ == "__main__":
    unittest.main()
