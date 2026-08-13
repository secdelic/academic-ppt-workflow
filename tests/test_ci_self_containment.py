from __future__ import annotations

import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from scripts.ci.verify_test_self_containment import scan_repository


ROOT = Path(__file__).resolve().parents[1]


def _init_repo(root: Path) -> None:
    subprocess.run(["git", "init", "-q"], cwd=root, check=True)
    (root / "tests" / "fixtures" / "synthetic" / "ci_contracts").mkdir(parents=True)
    (root / ".github" / "workflows").mkdir(parents=True)
    (root / "scripts" / "ci").mkdir(parents=True)


class SelfContainmentGateTests(unittest.TestCase):
    def test_current_candidate_has_no_forbidden_test_dependency(self) -> None:
        report = scan_repository(ROOT)
        self.assertEqual(report["forbidden_dependencies"], 0, report["findings"])
        self.assertEqual(report["status"], "PASS")

    def test_repo_staging_dependency_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            _init_repo(root)
            path = root / "tests" / "test_bad.py"
            path.write_text('target = REPO / "staging" / "run"\n', encoding="utf-8")  # self-containment: documentation-only
            subprocess.run(["git", "add", "tests/test_bad.py"], cwd=root, check=True)
            report = scan_repository(root)
            self.assertEqual(report["status"], "FAIL")
            self.assertEqual(report["forbidden_dependencies"], 1)

    def test_generated_temporary_resource_is_classified_not_forbidden(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            _init_repo(root)
            path = root / "tests" / "test_generated.py"
            path.write_text('target = root / "staging"  # generated-temp\n', encoding="utf-8")
            subprocess.run(["git", "add", "tests/test_generated.py"], cwd=root, check=True)
            report = scan_repository(root)
            self.assertEqual(report["status"], "PASS", report["findings"])
            self.assertTrue(any(row["classification"] == "GENERATED_TEMP" for row in report["findings"]))

    def test_untracked_static_fixture_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            _init_repo(root)
            fixture = root / "tests" / "fixtures" / "synthetic" / "ci_contracts" / "contract.json"
            fixture.write_text(json.dumps({"fixture_provenance": "SYNTHETIC"}), encoding="utf-8")
            report = scan_repository(root)
            self.assertEqual(report["status"], "FAIL")
            self.assertEqual(report["findings"][0]["kind"], "UNTRACKED_STATIC_FIXTURE")

    def test_private_absolute_path_fails_closed_without_echoing_value(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            _init_repo(root)
            separator = chr(92)
            private_value = separator.join(("C:", "Users", "developer", "private.dat"))
            path = root / "tests" / "test_private.py"
            path.write_text(f'value = {private_value!r}\n', encoding="utf-8")
            subprocess.run(["git", "add", "tests/test_private.py"], cwd=root, check=True)
            report = scan_repository(root)
            self.assertEqual(report["status"], "FAIL")
            encoded = json.dumps(report)
            self.assertNotIn("developer", encoded)
            self.assertIn("redacted", encoded)


if __name__ == "__main__":
    unittest.main()
