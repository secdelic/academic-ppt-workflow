from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.deck_ir import build_deck_ir  # noqa: E402
from academic_ppt.incremental import (  # noqa: E402
    IncrementalUpdateError,
    UpdateTarget,
    backup_source_deck,
    build_update_targets,
    resolve_update_slide_ids,
    run_powerpoint_replacement,
    verify_no_manual_conflict,
)
from academic_ppt.template_fill import (  # noqa: E402
    TemplateFillError,
    run_native_template_fill,
)
from academic_ppt.utils import sha256_file  # noqa: E402


BRIEF = {
    "project_name": "Synthetic incremental contract",
    "presentation_type": "research_report",
    "language": "en-US",
}
PLANNED = [
    {
        "logical_slide_key": "opening/boundary",
        "section_id": "opening",
        "slide_role": "evidence_boundary",
        "slide_title": "Registered sources define the evidence boundary",
        "single_key_message": "Only synthetic registered material is used.",
        "source_ids": ["SRC-001"],
        "figure_ids": [],
        "manual_review_required": "no",
    },
    {
        "logical_slide_key": "results/primary",
        "section_id": "results",
        "slide_role": "primary_result",
        "slide_title": "The synthetic association is reported without causality",
        "single_key_message": "The registered synthetic estimate is retained.",
        "claim_ids": ["CLM-001"],
        "source_ids": ["SRC-001"],
        "figure_ids": ["FIG-001"],
        "manual_review_required": "yes",
    },
]


def _write_deck_ir(path: Path, ir: dict) -> None:
    path.write_text(
        json.dumps(ir, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _target(slide_id: str = "SLD-0123456789ABCDEF0123") -> UpdateTarget:
    return UpdateTarget(
        slide_id=slide_id,
        target_index=2,
        candidate_index=1,
        section_id="results",
        figure_ids=("FIG-001",),
        previous_content_hash="a" * 64,
        candidate_content_hash="b" * 64,
        candidate_revision=2,
    )


class SelectorAndTargetTests(unittest.TestCase):
    def test_selector_is_exclusive_and_whitespace_is_not_a_selector(self):
        ir = build_deck_ir(BRIEF, PLANNED)
        with self.assertRaisesRegex(IncrementalUpdateError, "Exactly one"):
            resolve_update_slide_ids(ir)
        with self.assertRaisesRegex(IncrementalUpdateError, "Exactly one"):
            resolve_update_slide_ids(
                ir, update_slide=ir["slides"][0]["slide_id"], update_section="results"
            )
        with self.assertRaisesRegex(IncrementalUpdateError, "Exactly one"):
            resolve_update_slide_ids(ir, update_slide="   ")

    def test_slide_section_and_figure_select_stable_semantic_ids(self):
        ir = build_deck_ir(BRIEF, PLANNED)
        opening_id = ir["slides"][0]["slide_id"]
        result_id = ir["slides"][1]["slide_id"]
        self.assertEqual(
            resolve_update_slide_ids(ir, update_slide=f"  {opening_id}  "),
            [opening_id],
        )
        self.assertEqual(
            resolve_update_slide_ids(ir, update_section="results"), [result_id]
        )
        self.assertEqual(
            resolve_update_slide_ids(ir, update_figure="FIG-001"), [result_id]
        )

    def test_targets_bind_by_slide_id_not_page_number(self):
        previous = build_deck_ir(BRIEF, PLANNED)
        changed = copy.deepcopy(PLANNED)
        changed[1][
            "single_key_message"
        ] = "The revised registered synthetic estimate is retained."
        candidate = build_deck_ir(
            BRIEF, list(reversed(changed)), existing_ir=previous
        )
        result_id = previous["slides"][1]["slide_id"]
        targets = build_update_targets(previous, candidate, [result_id])
        self.assertEqual(len(targets), 1)
        self.assertEqual(targets[0].target_index, 2)
        self.assertEqual(targets[0].candidate_index, 1)
        self.assertEqual(targets[0].candidate_revision, 2)
        self.assertNotEqual(
            targets[0].previous_content_hash, targets[0].candidate_content_hash
        )

    def test_target_contract_rejects_wrong_deck_duplicates_and_revision_regression(self):
        previous = build_deck_ir(BRIEF, PLANNED)
        result_id = previous["slides"][1]["slide_id"]
        different = copy.deepcopy(previous)
        different["deck_id"] = "DECK-DIFFERENT"
        with self.assertRaisesRegex(IncrementalUpdateError, "different deck_id"):
            build_update_targets(previous, different, [result_id])
        with self.assertRaisesRegex(IncrementalUpdateError, "Duplicate slide IDs"):
            build_update_targets(previous, previous, [result_id, result_id])

        regressed = copy.deepcopy(previous)
        regressed["slides"][1]["content_hash"] = "b" * 64
        regressed["slides"][1]["slide_revision"] = 1
        with self.assertRaisesRegex(IncrementalUpdateError, "must increment"):
            build_update_targets(previous, regressed, [result_id])


class BackupAndConflictTests(unittest.TestCase):
    def test_backup_is_new_hash_identical_and_source_unchanged(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            source = root / "existing.pptx"
            source.write_bytes(b"synthetic-pptx-source")
            before = sha256_file(source)
            report = backup_source_deck(source, root / "backup")
            self.assertEqual(report["source_sha256"], before)
            self.assertEqual(report["source_sha256_after"], before)
            self.assertEqual(report["backup_sha256"], before)
            self.assertEqual(sha256_file(source), before)
            self.assertTrue(Path(report["backup_path"]).is_file())
            with self.assertRaisesRegex(IncrementalUpdateError, "overwrite backup"):
                backup_source_deck(source, root / "backup")

    def test_managed_object_conflicts_are_blocked_without_explicit_permission(self):
        slide_id = "SLD-0123456789ABCDEF0123"
        prior = {
            "slides": [
                {"slide_id": slide_id, "managed_object_hash": "prior-hash"}
            ]
        }
        same = {
            "slides": [
                {"slide_id": slide_id, "managed_object_hash": "prior-hash"}
            ]
        }
        changed = {
            "slides": [
                {"slide_id": slide_id, "managed_object_hash": "manual-edit-hash"}
            ]
        }
        self.assertEqual(
            verify_no_manual_conflict(
                [slide_id], same, prior, allow_managed_overwrite=False
            ),
            [],
        )
        with self.assertRaisesRegex(IncrementalUpdateError, "MANUAL_EDIT_CONFLICT"):
            verify_no_manual_conflict(
                [slide_id], changed, prior, allow_managed_overwrite=False
            )
        warnings = verify_no_manual_conflict(
            [slide_id], changed, prior, allow_managed_overwrite=True
        )
        self.assertIn(slide_id, warnings[0])

    def test_unknown_manual_edit_state_requires_explicit_permission(self):
        slide_id = "SLD-0123456789ABCDEF0123"
        with self.assertRaisesRegex(IncrementalUpdateError, "STATE_UNKNOWN"):
            verify_no_manual_conflict(
                [slide_id], None, None, allow_managed_overwrite=False
            )
        warnings = verify_no_manual_conflict(
            [slide_id], None, None, allow_managed_overwrite=True
        )
        self.assertIn("STATE_UNKNOWN", warnings[0])


class ReplacementExecutionContractTests(unittest.TestCase):
    def test_command_uses_working_copy_and_preserves_all_inputs(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            existing = root / "existing.pptx"
            candidate = root / "candidate.pptx"
            script = root / "replace.ps1"
            output = root / "formal" / "updated.pptx"
            existing.write_bytes(b"existing-deck")
            candidate.write_bytes(b"candidate-deck")
            script.write_text("# synthetic mocked script\n", encoding="utf-8")
            hashes = {
                path: sha256_file(path) for path in (existing, candidate, script)
            }

            with patch(
                "academic_ppt.incremental.subprocess.run",
                return_value=subprocess.CompletedProcess(
                    args=[], returncode=0, stdout="UPDATED_SLIDES=1", stderr=""
                ),
            ) as mocked:
                log = run_powerpoint_replacement(
                    existing_pptx=existing,
                    candidate_pptx=candidate,
                    output_pptx=output,
                    targets=[_target()],
                    script_path=script,
                    timeout_seconds=20,
                )

            self.assertIn("UPDATED_SLIDES=1", log)
            self.assertEqual(output.read_bytes(), existing.read_bytes())
            for path, digest in hashes.items():
                self.assertEqual(sha256_file(path), digest)
            command = mocked.call_args.args[0]
            self.assertEqual(command[:6], [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
            ])
            self.assertEqual(command[command.index("-ExistingCopy") + 1], str(output.resolve()))
            self.assertEqual(command[command.index("-Candidate") + 1], str(candidate.resolve()))
            mapping = json.loads(
                (output.parent / "slide_update_mapping.json").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(mapping["targets"][0]["slide_id"], _target().slide_id)
            with self.assertRaisesRegex(IncrementalUpdateError, "overwrite update output"):
                run_powerpoint_replacement(
                    existing_pptx=existing,
                    candidate_pptx=candidate,
                    output_pptx=output,
                    targets=[_target()],
                    script_path=script,
                )

    def test_path_and_mapping_contracts_fail_before_subprocess(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            existing = root / "existing.pptx"
            candidate = root / "candidate.pptx"
            script = root / "replace.ps1"
            existing.write_bytes(b"existing")
            candidate.write_bytes(b"candidate")
            script.write_text("# mock\n", encoding="utf-8")
            duplicate_index = replace(_target(), candidate_index=2)
            duplicate_index = [
                _target(),
                replace(
                    duplicate_index,
                    slide_id="SLD-1123456789ABCDEF0123",
                    target_index=2,
                ),
            ]
            with patch("academic_ppt.incremental.subprocess.run") as mocked:
                with self.assertRaisesRegex(
                    IncrementalUpdateError, "Duplicate target slide indexes"
                ):
                    run_powerpoint_replacement(
                        existing_pptx=existing,
                        candidate_pptx=candidate,
                        output_pptx=root / "out.pptx",
                        targets=duplicate_index,
                        script_path=script,
                    )
                with self.assertRaisesRegex(
                    IncrementalUpdateError, "must be a .pptx"
                ):
                    run_powerpoint_replacement(
                        existing_pptx=existing,
                        candidate_pptx=candidate,
                        output_pptx=root / "out.pdf",
                        targets=[_target()],
                        script_path=script,
                    )
            mocked.assert_not_called()

    def test_input_mutation_by_backend_is_detected(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            existing = root / "existing.pptx"
            candidate = root / "candidate.pptx"
            script = root / "replace.ps1"
            existing.write_bytes(b"existing")
            candidate.write_bytes(b"candidate")
            script.write_text("# mock\n", encoding="utf-8")

            def mutate_candidate(*_args, **_kwargs):
                candidate.write_bytes(b"changed-by-backend")
                return subprocess.CompletedProcess(
                    args=[], returncode=0, stdout="", stderr=""
                )

            with patch(
                "academic_ppt.incremental.subprocess.run",
                side_effect=mutate_candidate,
            ):
                with self.assertRaisesRegex(IncrementalUpdateError, "Input changed"):
                    run_powerpoint_replacement(
                        existing_pptx=existing,
                        candidate_pptx=candidate,
                        output_pptx=root / "out.pptx",
                        targets=[_target()],
                        script_path=script,
                    )


class NativeTemplateContractTests(unittest.TestCase):
    def _valid_inputs(self, root: Path) -> tuple[Path, Path, Path]:
        template = root / "reference.pptx"
        deck_ir = root / "deck_ir.json"
        script = root / "fill.ps1"
        template.write_bytes(b"synthetic-template")
        _write_deck_ir(deck_ir, build_deck_ir(BRIEF, PLANNED))
        script.write_text("# synthetic mocked script\n", encoding="utf-8")
        return template, deck_ir, script

    def test_invalid_deck_ir_is_rejected_before_backend_execution(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            template = root / "reference.pptx"
            deck_ir = root / "deck_ir.json"
            script = root / "fill.ps1"
            template.write_bytes(b"synthetic-template")
            deck_ir.write_text('{"slides": []}\n', encoding="utf-8")
            script.write_text("# mock\n", encoding="utf-8")
            with patch("academic_ppt.template_fill.subprocess.run") as mocked:
                with self.assertRaisesRegex(TemplateFillError, "contract validation"):
                    run_native_template_fill(
                        template_path=template,
                        deck_ir_path=deck_ir,
                        output_pptx=root / "out.pptx",
                        script_path=script,
                    )
            mocked.assert_not_called()

    def test_mocked_native_fill_builds_safe_command_and_preserves_inputs(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            template, deck_ir, script = self._valid_inputs(root)
            output = root / "formal" / "filled.pptx"
            hashes = {
                path: sha256_file(path) for path in (template, deck_ir, script)
            }

            def create_output(command, **_kwargs):
                output_arg = Path(command[command.index("-OutputPptx") + 1])
                output_arg.write_bytes(b"mocked-native-pptx")
                return subprocess.CompletedProcess(
                    args=command,
                    returncode=0,
                    stdout="NATIVE_TEMPLATE_SLIDES=2",
                    stderr="",
                )

            with patch(
                "academic_ppt.template_fill.subprocess.run",
                side_effect=create_output,
            ) as mocked:
                log = run_native_template_fill(
                    template_path=template,
                    deck_ir_path=deck_ir,
                    output_pptx=output,
                    script_path=script,
                    timeout_seconds=20,
                )
            self.assertIn("NATIVE_TEMPLATE_SLIDES=2", log)
            self.assertEqual(output.read_bytes(), b"mocked-native-pptx")
            for path, digest in hashes.items():
                self.assertEqual(sha256_file(path), digest)
            command = mocked.call_args.args[0]
            self.assertEqual(command[:6], [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
            ])
            self.assertEqual(command[command.index("-Template") + 1], str(template.resolve()))
            self.assertEqual(command[command.index("-DeckIr") + 1], str(deck_ir.resolve()))
            self.assertNotIn("http", " ".join(command).lower())

    def test_native_fill_detects_input_mutation_and_refuses_overwrite(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            template, deck_ir, script = self._valid_inputs(root)
            output = root / "filled.pptx"

            def mutate_template(command, **_kwargs):
                template.write_bytes(b"mutated")
                Path(command[command.index("-OutputPptx") + 1]).write_bytes(b"output")
                return subprocess.CompletedProcess(
                    args=command, returncode=0, stdout="", stderr=""
                )

            with patch(
                "academic_ppt.template_fill.subprocess.run",
                side_effect=mutate_template,
            ):
                with self.assertRaisesRegex(TemplateFillError, "Input changed"):
                    run_native_template_fill(
                        template_path=template,
                        deck_ir_path=deck_ir,
                        output_pptx=output,
                        script_path=script,
                    )

            template.write_bytes(b"synthetic-template-restored")
            existing_output = root / "already_exists.pptx"
            existing_output.write_bytes(b"formal")
            with patch("academic_ppt.template_fill.subprocess.run") as mocked:
                with self.assertRaisesRegex(TemplateFillError, "Refusing to overwrite"):
                    run_native_template_fill(
                        template_path=template,
                        deck_ir_path=deck_ir,
                        output_pptx=existing_output,
                        script_path=script,
                    )
            mocked.assert_not_called()


class PowerShellStaticSafetyTests(unittest.TestCase):
    def test_scripts_parse_without_invoking_com(self):
        scripts = [
            ROOT / "scripts" / "replace_pptx_slides.ps1",
            ROOT / "scripts" / "fill_native_template.ps1",
        ]
        for script in scripts:
            escaped = str(script).replace("'", "''")
            command = (
                "$tokens=$null; $errors=$null; "
                f"[System.Management.Automation.Language.Parser]::ParseFile('{escaped}', "
                "[ref]$tokens, [ref]$errors) | Out-Null; "
                "if ($errors.Count -gt 0) { "
                "$errors | ForEach-Object { Write-Error $_.Message }; exit 1 }"
            )
            completed = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    command,
                ],
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=20,
            )
            self.assertEqual(
                completed.returncode,
                0,
                msg=f"{script.name}: {completed.stdout}\n{completed.stderr}",
            )

    def test_scripts_have_no_network_install_or_source_slide_copy_commands(self):
        replacement = (
            ROOT / "scripts" / "replace_pptx_slides.ps1"
        ).read_text(encoding="utf-8")
        native = (
            ROOT / "scripts" / "fill_native_template.ps1"
        ).read_text(encoding="utf-8")
        combined = (replacement + "\n" + native).lower()
        for prohibited in (
            "invoke-webrequest",
            "invoke-restmethod",
            "start-process",
            "install-module",
            "install-package",
            "git ",
        ):
            self.assertNotIn(prohibited, combined)
        self.assertNotIn("savecopyas", native.lower())
        self.assertIn("designs.clone", native.lower())
        self.assertIn("never source slides", native.lower())
        self.assertIn("remove-referencecontentshapes", native.lower())
        self.assertIn("reference_content_shapes_removed", native.lower())
        self.assertIn("insertfromfile", replacement.lower())
        self.assertNotIn(".slides.paste", replacement.lower())
        self.assertNotIn(".copy()", replacement.lower())


if __name__ == "__main__":
    unittest.main()
