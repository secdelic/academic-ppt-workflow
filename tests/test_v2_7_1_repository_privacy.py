from __future__ import annotations

import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import scripts.audit_repository_content as privacy_audit


def _write_ooxml(path: Path, *, private_path: str | None = None, external: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    description = private_path or "synthetic figure"
    relationship = (
        '<Relationship Id="rId9" Type="urn:synthetic" '
        'Target="https://example.invalid/asset" TargetMode="External"/>'
        if external
        else ""
    )
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="urn:synthetic"/>',
        )
        package.writestr(
            "ppt/slides/slide1.xml",
            f'<?xml version="1.0" encoding="UTF-8"?><slide descr="{description}"/>',
        )
        package.writestr(
            "ppt/slides/_rels/slide1.xml.rels",
            f'<?xml version="1.0" encoding="UTF-8"?><Relationships>{relationship}</Relationships>',
        )


def _audit_explicit(repo: Path, paths: list[Path]) -> dict:
    relative = [path.relative_to(repo).as_posix() for path in paths]
    with patch.object(privacy_audit, "candidate_files", return_value=relative):
        return privacy_audit.audit(repo)


class RepositoryPrivacyTests(unittest.TestCase):
    def test_candidate_files_delegates_to_release_selection(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo = Path(temp_dir)
            included = repo / "scripts" / "example.py"
            excluded = repo / "input" / "private.txt"
            included.parent.mkdir(parents=True)
            excluded.parent.mkdir(parents=True)
            included.write_text("print('synthetic')\n", encoding="utf-8")
            excluded.write_text("synthetic\n", encoding="utf-8")

            self.assertEqual(privacy_audit.candidate_files(repo), ["scripts/example.py"])

    def test_json_escaped_windows_path_is_blocked_without_literal_test_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo = Path(temp_dir)
            path = repo / "config" / "synthetic.json"
            path.parent.mkdir(parents=True)
            drive = chr(67) + chr(58)
            separator = chr(92)
            private_value = separator.join((drive, "Users", "Example", "private", "asset.png"))
            path.write_text(json.dumps({"asset": private_value}), encoding="utf-8")

            result = _audit_explicit(repo, [path])

            self.assertEqual(result["status"], "BLOCKED")
            self.assertTrue(any(issue["issue"] == "absolute private path" for issue in result["issues"]))
            self.assertNotIn(private_value, json.dumps(result))

    def test_ooxml_internal_path_and_external_relationship_are_blocked(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo = Path(temp_dir)
            package = repo / "visual_workspaces" / "demo" / "synthetic_preview.pptx"
            drive = chr(68) + chr(58)
            separator = chr(92)
            private_value = separator.join((drive, "research", "private", "figure.png"))
            _write_ooxml(package, private_path=private_value, external=True)

            result = _audit_explicit(repo, [package])

            issues = {issue["issue"] for issue in result["issues"]}
            self.assertEqual(result["status"], "BLOCKED")
            self.assertIn("absolute private path", issues)
            self.assertIn("OOXML external relationship", issues)
            self.assertEqual(result["office_packages_scanned"], 1)
            self.assertEqual(result["external_relationship_count"], 1)
            self.assertNotIn(private_value, json.dumps(result))

    def test_clean_allowlisted_synthetic_ooxml_passes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo = Path(temp_dir)
            package = repo / "visual_workspaces" / "demo" / "synthetic_preview.pptx"
            _write_ooxml(package)

            result = _audit_explicit(repo, [package])

            self.assertEqual(result["status"], "PASS")
            self.assertEqual(result["binary_files_checked"], 1)
            self.assertEqual(result["office_xml_parts_scanned"], 3)

    def test_path_and_asset_policy_controls_are_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repo = Path(temp_dir)
            secret_file = repo / "config" / "settings.yaml"
            local_file = repo / "config" / "local.yaml"
            temp_file = repo / "scripts" / "draft.tmp"
            font_file = repo / "visual_workspaces" / "demo" / "font.ttf"
            unapproved_binary = repo / "docs" / "image.png"
            prohibited_file = repo / "private" / "notes.txt"
            large_file = repo / "scripts" / "large.txt"
            paths = [
                secret_file,
                local_file,
                temp_file,
                font_file,
                unapproved_binary,
                prohibited_file,
                large_file,
            ]
            for path in paths:
                path.parent.mkdir(parents=True, exist_ok=True)
            key_name = "api" + "_key"
            secret_file.write_text(key_name + ": " + chr(34) + "sk-" + "A" * 24 + chr(34), encoding="utf-8")
            local_file.write_text("synthetic: true\n", encoding="utf-8")
            temp_file.write_text("synthetic\n", encoding="utf-8")
            font_file.write_bytes(b"synthetic-font")
            unapproved_binary.write_bytes(b"not-an-authorized-image")
            prohibited_file.write_text("synthetic\n", encoding="utf-8")
            large_file.write_bytes(b"x" * (privacy_audit.MAX_RELEASE_FILE_BYTES + 1))

            result = _audit_explicit(repo, paths)

            issues = {issue["issue"] for issue in result["issues"]}
            self.assertEqual(result["status"], "BLOCKED")
            self.assertTrue(
                {
                    "secret-like value",
                    "local configuration file",
                    "temporary or backup file",
                    "embedded font file is prohibited",
                    "binary asset is outside the release allowlist",
                    "prohibited release path class",
                    "large file requires explicit authorization",
                }
                <= issues
            )


if __name__ == "__main__":
    unittest.main()
