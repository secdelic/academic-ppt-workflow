from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.evidence import build_evidence
from academic_ppt.extractors import extract_source
from academic_ppt.inventory import inventory_sources
from academic_ppt.utils import load_yaml_compatible, resolve_path


class WorkflowUnitTests(unittest.TestCase):
    def test_configuration_is_json_compatible_yaml(self):
        config = load_yaml_compatible(ROOT / "config" / "workflow.yaml")
        self.assertEqual(config["schema_version"], "1.0")
        self.assertIn(".pptx", config["supported_extensions"])
        self.assertIn(".svg", config["supported_extensions"])

    def test_cli_path_has_priority(self):
        with tempfile.TemporaryDirectory() as temp:
            repo = Path(temp)
            resolved = resolve_path("cli-input", "NONEXISTENT_ENV", "configured", repo, "default")
            self.assertEqual(resolved, (repo / "cli-input").resolve())

    def test_inventory_hash_and_read_only_source(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "input" / "study.md"
            source.parent.mkdir()
            source.write_text("source text", encoding="utf-8")
            before = source.read_bytes()
            rows = inventory_sources(root / "input", {".md"}, root / "manifest.csv")
            self.assertEqual(len(rows), 1)
            self.assertEqual(len(rows[0]["sha256"]), 64)
            self.assertEqual(source.read_bytes(), before)

    def test_extract_text_and_csv(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            md = root / "x.md"
            csv_path = root / "x.csv"
            md.write_text("Background\nExact text.", encoding="utf-8")
            csv_path.write_text("category,value\nA,1\n", encoding="utf-8")
            self.assertIn("Exact text", extract_source(md)[0])
            self.assertIn("category | value", extract_source(csv_path)[0])

    def test_claim_registry_does_not_complete_missing_content(self):
        with tempfile.TemporaryDirectory() as temp:
            staging = Path(temp)
            claims, unresolved = build_evidence(
                {
                    "SRC-test": (
                        "CLAIM|Exact source claim.|section 1|reported fact|high|"
                        "Exact source claim.|Do not add causality"
                    )
                },
                [
                    {
                        "source_id": "SRC-test",
                        "file_type": "md",
                        "relative_path": "test.md",
                        "parsed_successfully": "yes",
                        "parse_warning": "",
                        "possible_sensitive_information": "no",
                    }
                ],
                staging,
                {"project_name": "test", "key_message": "INFORMATION_REQUIRED"},
            )
            self.assertEqual(len(claims), 1)
            self.assertEqual(claims[0]["claim_text"], "Exact source claim.")
            self.assertTrue(any("key_message" in item for item in unresolved))


if __name__ == "__main__":
    unittest.main()
