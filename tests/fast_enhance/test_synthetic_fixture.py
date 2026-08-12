from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fixture_factory import (  # noqa: E402
    build_fixture_definition,
    discover_node_runtime,
    generate_synthetic_pptx,
    materialize_synthetic_fixture,
)


def _tree_hashes(root: Path) -> dict[str, str]:
    return {
        path.relative_to(root).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def _pptx_slide_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(
            1
            for name in archive.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )


def _pptx_master_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(
            1
            for name in archive.namelist()
            if name.startswith("ppt/slideMasters/slideMaster") and name.endswith(".xml")
        )


def _pptx_layout_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(
            1
            for name in archive.namelist()
            if name.startswith("ppt/slideLayouts/slideLayout") and name.endswith(".xml")
        )


class SyntheticFastEnhanceFixtureTests(unittest.TestCase):
    def test_fixture_contract_is_twenty_one_plus_seven_without_page_derived_ids(self):
        fixture = build_fixture_definition()
        expected = fixture["expected"]
        self.assertEqual(expected["original_slide_count"], 21)
        self.assertEqual(expected["changed_slide_count"], 7)
        self.assertEqual(expected["final_slide_count"], 25)
        self.assertEqual(
            expected["operation_counts"],
            {"KEEP": 18, "REPLACE": 3, "INSERT_AFTER": 4},
        )
        baseline_ids = [slide["slide_id"] for slide in fixture["baseline_slides"]]
        candidate_ids = [slide["slide_id"] for slide in fixture["candidate_slides"]]
        self.assertEqual(len(baseline_ids), len(set(baseline_ids)))
        self.assertEqual(len(candidate_ids), len(set(candidate_ids)))
        for slide_id in baseline_ids + candidate_ids:
            self.assertRegex(slide_id, r"^SLD-[A-F0-9]{20}$")
            self.assertNotRegex(slide_id, r"(?:SLIDE|PAGE)[-_]?\d+$")

    def test_materialization_writes_only_text_specs_until_render_is_requested(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "fixture"
            fixture = materialize_synthetic_fixture(root)
            self.assertIsNone(fixture.baseline_pptx)
            self.assertIsNone(fixture.candidate_pptx)
            self.assertFalse(any(root.rglob("*.pptx")))
            contract = json.loads(fixture.fixture_contract.read_text(encoding="utf-8"))
            self.assertTrue(contract["synthetic_only"])
            self.assertEqual(contract["expected"]["changed_slide_count"], 7)

    def test_generation_is_runtime_only_and_does_not_modify_fixture_inputs(self):
        if discover_node_runtime(ROOT) is None:
            self.skipTest("Local Node/PptxGenJS runtime unavailable; no installation allowed")
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "fixture"
            fixture = materialize_synthetic_fixture(root)
            before_prior = _tree_hashes(fixture.prior_inputs)
            before_current = _tree_hashes(fixture.current_inputs)
            rendered = generate_synthetic_pptx(fixture)
            self.assertEqual(_tree_hashes(fixture.prior_inputs), before_prior)
            self.assertEqual(_tree_hashes(fixture.current_inputs), before_current)
            assert rendered.baseline_pptx is not None
            assert rendered.candidate_pptx is not None
            self.assertEqual(_pptx_slide_count(rendered.baseline_pptx), 21)
            self.assertEqual(_pptx_slide_count(rendered.candidate_pptx), 7)
            self.assertGreaterEqual(_pptx_master_count(rendered.baseline_pptx), 1)
            self.assertGreaterEqual(_pptx_master_count(rendered.candidate_pptx), 1)
            self.assertGreaterEqual(_pptx_layout_count(rendered.baseline_pptx), 2)
            self.assertGreaterEqual(_pptx_layout_count(rendered.candidate_pptx), 2)
            manifest = json.loads(
                rendered.generator_manifest.read_text(encoding="utf-8")  # type: ignore[union-attr]
            )
            self.assertEqual(manifest["external_service_usage"], 0)
            self.assertEqual(manifest["baseline_slide_count"], 21)
            self.assertEqual(manifest["candidate_slide_count"], 7)

    def test_fixture_contains_no_real_case_or_clinical_identity_terms(self):
        with tempfile.TemporaryDirectory() as temporary:
            fixture = materialize_synthetic_fixture(Path(temporary) / "fixture")
            body = "\n".join(
                path.read_text(encoding="utf-8")
                for path in fixture.root.rglob("*")
                if path.is_file()
            ).lower()
            prohibited = (
                "patient",
                "medical record",
                "bed number",
                "姓名",
                "患者",
                "床号",
                "住院号",
            )
            for marker in prohibited:
                self.assertNotIn(marker.lower(), body)

    def test_repository_fast_fixture_has_no_checked_in_binary_deck(self):
        fixture_dir = Path(__file__).resolve().parent
        binary_suffixes = {".pptx", ".ppt", ".potx", ".pdf", ".png", ".jpg", ".jpeg"}
        binary_files = [
            path for path in fixture_dir.rglob("*") if path.is_file() and path.suffix.lower() in binary_suffixes
        ]
        self.assertEqual(binary_files, [])

    def test_fixture_cli_runs_under_locked_python_without_private_path_output(self):
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "cli_fixture"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(Path(__file__).with_name("fixture_factory.py")),
                    "--out",
                    str(output),
                ],
                cwd=ROOT,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
                check=False,
            )
            self.assertEqual(completed.returncode, 0, completed.stderr)
            summary = json.loads(completed.stdout)
            self.assertEqual(summary["original_slide_count"], 21)
            self.assertEqual(summary["changed_slide_count"], 7)
            self.assertEqual(summary["external_service_usage"], 0)
            self.assertNotIn(str(output), completed.stdout)


if __name__ == "__main__":
    unittest.main()
