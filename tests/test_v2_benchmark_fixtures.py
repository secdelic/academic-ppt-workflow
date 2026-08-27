from __future__ import annotations

import csv
import hashlib
import importlib.util
import json
import tempfile
import unittest
import zipfile
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = REPO_ROOT / "scripts" / "v2" / "create_benchmark_fixtures.py"
SPEC = importlib.util.spec_from_file_location("create_benchmark_fixtures", MODULE_PATH)
assert SPEC and SPEC.loader
fixtures_module = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(fixtures_module)


def _hash_tree(root: Path) -> dict[str, str]:
    return {
        path.relative_to(root).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in sorted(root.rglob("*"))
        if path.is_file() and path.name not in {"fixture_manifest.json", "verification.json"}
    }


class BenchmarkFixtureTests(unittest.TestCase):
    def _generate(self, root: Path) -> dict:
        return fixtures_module.create_all_fixtures(root, verify_docx_render=False)

    def test_expected_cases_and_types_are_created(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "fixtures"
            verification = self._generate(root)
            expected = [
                root / "case1_paper_report" / "documents" / "synthetic_trial_report.docx",
                root / "case1_paper_report" / "data" / "aggregate_results.csv",
                root / "case1_paper_report" / "figures" / "figure_01_trial_flow.png",
                root / "case1_paper_report" / "figures" / "figure_02_effect_summary.png",
                root / "case2_defense_project" / "documents" / "synthetic_project_plan.docx",
                root / "case2_defense_project" / "data" / "aggregate_project_metrics.csv",
                root / "case2_defense_project" / "protocol.md",
                root / "case2_defense_project" / "figures" / "figure_01_cohort_architecture.png",
                root / "case2_defense_project" / "figures" / "figure_02_project_timeline.png",
                root / "case2_defense_project" / "figures" / "figure_03_multiomics_workflow.png",
                root / "case3_reference_template" / "reference" / "protected_style_reference.pptx",
                root / "case3_reference_template" / "new_content" / "documents" / "synthetic_new_content.docx",
                root / "case3_reference_template" / "new_content" / "data" / "aggregate_period_results.csv",
                root / "fixture_manifest.json",
            ]
            self.assertTrue(all(path.is_file() and path.stat().st_size > 0 for path in expected))
            self.assertTrue(verification["structural_verification"]["sentinel_scan"]["passed"])

    def test_every_tabular_record_is_aggregate_and_synthetic(self) -> None:
        forbidden_headers = {"patient_id", "subject_id", "mrn", "name", "date_of_birth"}
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "fixtures"
            self._generate(root)
            for path in root.rglob("*.csv"):
                with path.open("r", encoding="utf-8-sig", newline="") as handle:
                    reader = csv.DictReader(handle)
                    self.assertTrue(reader.fieldnames)
                    self.assertTrue(forbidden_headers.isdisjoint({field.lower() for field in reader.fieldnames or []}))
                    for row in reader:
                        self.assertEqual(row.get("record_scope"), "aggregate_only")
                        self.assertEqual(row.get("synthetic_status"), fixtures_module.SYNTHETIC_MARKER)

    def test_private_sentinel_exists_only_in_protected_reference(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "fixtures"
            self._generate(root)
            scan = fixtures_module._scan_sentinel(root)
            self.assertEqual(
                scan["present"],
                ["case3_reference_template/reference/protected_style_reference.pptx"],
            )
            reference = root / scan["present"][0]
            with zipfile.ZipFile(reference, "r") as archive:
                payload = b"\n".join(archive.read(name) for name in archive.namelist())
            self.assertIn(fixtures_module.PRIVATE_SENTINEL.encode("utf-8"), payload)
            metadata = json.loads(
                (root / "case3_reference_template" / "source_metadata.json").read_text(encoding="utf-8")
            )
            self.assertNotIn(fixtures_module.PRIVATE_SENTINEL, json.dumps(metadata))
            self.assertEqual(metadata["private_sentinel_sha256"], fixtures_module.PRIVATE_SENTINEL_SHA256)

    def test_office_packages_have_required_editable_structure(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "fixtures"
            verification = self._generate(root)["structural_verification"]
            self.assertEqual(len(verification["docx"]), 3)
            for item in verification["docx"]:
                self.assertTrue(item["all_required_parts_present"])
                self.assertTrue(item["named_styles_present"])
                self.assertTrue(item["fixed_table_width_9360_present"])
                self.assertTrue(item["table_indent_120_present"])
                self.assertTrue(item["synthetic_marker_present"])
            pptx = verification["pptx"]
            self.assertEqual(pptx["slide_count"], 4)
            self.assertGreaterEqual(pptx["master_count"], 1)
            self.assertGreaterEqual(pptx["layout_count"], 3)
            self.assertTrue(pptx["editable_chart_present"])
            self.assertTrue(pptx["theme_present"])

    def test_generation_is_byte_stable_for_scientific_inputs(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            first = root / "first"
            second = root / "second"
            self._generate(first)
            self._generate(second)
            self.assertEqual(_hash_tree(first), _hash_tree(second))

    def test_manifest_discloses_real_world_validation_pending(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "fixtures"
            self._generate(root)
            manifest = json.loads((root / "fixture_manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["synthetic_status"], "SYNTHETIC_FIXTURE_ONLY")
            self.assertEqual(manifest["real_world_status"], "REAL_WORLD_VALIDATION_PENDING")
            self.assertFalse(manifest["patient_level_data"])
            self.assertFalse(manifest["network_used"])
            self.assertFalse(manifest["external_install_performed"])


if __name__ == "__main__":
    unittest.main()
