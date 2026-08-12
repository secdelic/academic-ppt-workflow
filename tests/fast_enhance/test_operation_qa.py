from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.fast_qa import (  # noqa: E402
    run_changed_slide_qa,
    run_whole_deck_lightweight_qa,
)
from academic_ppt.incremental import (  # noqa: E402
    IncrementalUpdateError,
    build_operation_plan,
    run_powerpoint_operations,
    validate_operation_plan,
    write_operation_plan,
)

from tests.fast_enhance.fixture_factory import (  # noqa: E402
    build_fixture_definition,
    materialize_synthetic_fixture,
)


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _write_minimal_deck(
    path: Path, slide_count: int, slide_specs: list[dict] | None = None
) -> None:
    content_types = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
    ]
    root_rels = (
        f'<Relationships xmlns="{PR_NS}"><Relationship Id="rId1" Type="{R_NS}/officeDocument" '
        'Target="ppt/presentation.xml"/></Relationships>'
    )
    presentation_ids = []
    presentation_rels = []
    members: dict[str, str] = {}
    layout_xml = (
        f'<p:sldLayout xmlns:p="{P_NS}" xmlns:a="{A_NS}" type="blank">'
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>'
    )
    members["ppt/slideLayouts/slideLayout1.xml"] = layout_xml
    for index in range(1, slide_count + 1):
        content_types.append(
            f'<Override PartName="/ppt/slides/slide{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
        presentation_ids.append(f'<p:sldId id="{255 + index}" r:id="rId{index}"/>')
        presentation_rels.append(
            f'<Relationship Id="rId{index}" Type="{R_NS}/slide" Target="slides/slide{index}.xml"/>'
        )
        members[f"ppt/slides/slide{index}.xml"] = (
            f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
            f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title {index}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Synthetic slide</a:t></a:r></a:p></p:txBody></p:sp>'
            '</p:spTree></p:cSld></p:sld>'
        )
        slide_relationships = [
            f'<Relationship Id="rId1" Type="{R_NS}/slideLayout" '
            'Target="../slideLayouts/slideLayout1.xml"/>'
        ]
        if slide_specs is not None:
            spec = slide_specs[index - 1]
            content_types.append(
                f'<Override PartName="/ppt/notesSlides/notesSlide{index}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
            )
            slide_relationships.append(
                f'<Relationship Id="rId2" Type="{R_NS}/notesSlide" '
                f'Target="../notesSlides/notesSlide{index}.xml"/>'
            )
            note_text = (
                f"[Slide-ID]\n{spec['slide_id']}\n"
                f"[Slide-Revision]\n{spec['slide_revision']}\n"
                f"[Content-Hash]\n{spec['content_hash']}"
            )
            members[f"ppt/notesSlides/notesSlide{index}.xml"] = (
                f'<p:notes xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><p:spTree>'
                '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
                '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                f'<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{note_text}</a:t></a:r></a:p></p:txBody></p:sp>'
                '</p:spTree></p:cSld></p:notes>'
            )
        members[f"ppt/slides/_rels/slide{index}.xml.rels"] = (
            f'<Relationships xmlns="{PR_NS}">' + "".join(slide_relationships) + "</Relationships>"
        )
    content_types.append("</Types>")
    members["[Content_Types].xml"] = "".join(content_types)
    members["_rels/.rels"] = root_rels
    members["ppt/presentation.xml"] = (
        f'<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst>'
        + "".join(presentation_ids)
        + "</p:sldIdLst></p:presentation>"
    )
    members["ppt/_rels/presentation.xml.rels"] = (
        f'<Relationships xmlns="{PR_NS}">' + "".join(presentation_rels) + "</Relationships>"
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, value in members.items():
            package.writestr(name, value)


def _build_plan(root: Path) -> tuple[dict, Path, Path]:
    definition = build_fixture_definition()
    source = root / "baseline.pptx"
    candidate = root / "candidate.pptx"
    _write_minimal_deck(source, len(definition["baseline_slides"]))
    _write_minimal_deck(
        candidate,
        len(definition["candidate_slides"]),
        definition["candidate_slides"],
    )
    operations = []
    for row in definition["operations"]:
        normalized = dict(row)
        if normalized["action"] in {"REPLACE", "INSERT_AFTER"}:
            normalized["candidate_pptx"] = str(candidate)
            candidate_slide = definition["candidate_slides"][normalized["candidate_index"] - 1]
            normalized["slide_revision"] = candidate_slide["slide_revision"]
            normalized["content_hash"] = candidate_slide["content_hash"]
        operations.append(normalized)
    final_ids = [row["slide_id"] for row in definition["final_slides"]]
    changed_specs = [
        {
            **row,
            "source_bindings": [
                {"source_id": source_id} for source_id in row["source_ids"]
            ],
            "scientific_qa_status": "PASS",
        }
        for row in definition["candidate_slides"]
    ]
    plan = build_operation_plan(
        source_pptx=source,
        source_slide_ids=[row["slide_id"] for row in definition["baseline_slides"]],
        operations=operations,
        expected_final_order=final_ids,
        changed_slide_specs=changed_specs,
        source_impacts=[{
            "source_id": "SRC-SYN-B",
            "affected_slide_ids": [row["slide_id"] for row in changed_specs[:2]],
        }],
        manual_review_items=[{"review_id": "MR-SYN-001", "status": "pending"}],
    )
    return plan, source, candidate


class OperationPlanTests(unittest.TestCase):
    def test_keep_replace_insert_after_contract_is_explicit_and_hash_bound(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            plan, source, candidate = _build_plan(root)
            self.assertEqual(plan["expected_final_slide_count"], 25)
            self.assertEqual(len(plan["changed_slide_ids"]), 7)
            self.assertEqual(len(plan["changed_slide_specs"]), 7)
            self.assertEqual(
                {row["action"] for row in plan["operations"]},
                {"KEEP", "REPLACE", "INSERT_AFTER"},
            )
            self.assertEqual(plan["source_deck"]["path"], str(source.resolve()))
            self.assertEqual(plan["candidate_decks"][0]["path"], str(candidate.resolve()))
            self.assertEqual(validate_operation_plan(plan), plan)
            plan_path = write_operation_plan(root / "operation_plan.json", plan)
            self.assertTrue(plan_path.is_file())
            with self.assertRaisesRegex(IncrementalUpdateError, "overwrite operation plan"):
                write_operation_plan(plan_path, plan)

    def test_invalid_order_or_input_mutation_fails_closed(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            plan, source, _candidate = _build_plan(root)
            reordered = copy.deepcopy(plan)
            reordered["expected_final_order"][0:2] = reversed(reordered["expected_final_order"][0:2])
            with self.assertRaisesRegex(IncrementalUpdateError, "relative order"):
                validate_operation_plan(reordered)
            source.write_bytes(source.read_bytes() + b"changed")
            with self.assertRaisesRegex(IncrementalUpdateError, "hash"):
                validate_operation_plan(plan)

    def test_every_changed_slide_requires_new_scientifically_reviewed_spec(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            plan, _source, _candidate = _build_plan(root)

            missing = copy.deepcopy(plan)
            missing["changed_slide_specs"] = missing["changed_slide_specs"][:-1]
            with self.assertRaisesRegex(IncrementalUpdateError, "cover every"):
                validate_operation_plan(missing)

            no_review = copy.deepcopy(plan)
            no_review["changed_slide_specs"][0].pop("scientific_qa_status")
            with self.assertRaisesRegex(IncrementalUpdateError, "scientific review PASS"):
                validate_operation_plan(no_review)

            no_binding = copy.deepcopy(plan)
            no_binding["changed_slide_specs"][0]["source_bindings"] = []
            with self.assertRaisesRegex(IncrementalUpdateError, "no source_bindings"):
                validate_operation_plan(no_binding)

    def test_operation_revision_and_content_hash_must_match_changed_spec(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            plan, _source, _candidate = _build_plan(root)
            first_changed_id = plan["changed_slide_ids"][0]
            original_operation = next(
                item for item in plan["operations"] if item["slide_id"] == first_changed_id
            )

            wrong_hash = copy.deepcopy(plan)
            changed_operation = next(
                item
                for item in wrong_hash["operations"]
                if item["slide_id"] == first_changed_id
            )
            changed_operation["content_hash"] = "0" * 64
            with self.assertRaisesRegex(IncrementalUpdateError, "content_hash must match"):
                validate_operation_plan(wrong_hash)

            wrong_revision = copy.deepcopy(plan)
            changed_operation = next(
                item
                for item in wrong_revision["operations"]
                if item["slide_id"] == first_changed_id
            )
            changed_operation["slide_revision"] = int(original_operation["slide_revision"]) + 1
            with self.assertRaisesRegex(IncrementalUpdateError, "slide_revision must match"):
                validate_operation_plan(wrong_revision)

    def test_candidate_slide_notes_must_match_reviewed_slidespec(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            plan, _source, candidate = _build_plan(root)
            notes_name = "ppt/notesSlides/notesSlide1.xml"
            with zipfile.ZipFile(candidate, "r") as package:
                members = {
                    name: package.read(name) for name in package.namelist()
                }
            expected_hash = plan["changed_slide_specs"][0]["content_hash"]
            members[notes_name] = members[notes_name].replace(
                expected_hash.encode("utf-8"), b"0" * 64
            )
            with zipfile.ZipFile(candidate, "w", zipfile.ZIP_DEFLATED) as package:
                for name, payload in members.items():
                    package.writestr(name, payload)
            changed = copy.deepcopy(plan)
            candidate_row = changed["candidate_decks"][0]
            from academic_ppt.utils import sha256_file

            candidate_row["sha256"] = sha256_file(candidate)
            changed["input_integrity"]["candidate_sha256"][
                candidate_row["candidate_id"]
            ] = candidate_row["sha256"]
            with self.assertRaisesRegex(
                IncrementalUpdateError, "Candidate content_hash"
            ):
                validate_operation_plan(changed)

    def test_fixture_contains_no_patient_content(self):
        with tempfile.TemporaryDirectory() as raw:
            fixture = materialize_synthetic_fixture(Path(raw) / "fixture")
            for path in fixture.root.rglob("*"):
                if path.is_file():
                    text = path.read_text(encoding="utf-8", errors="ignore")
                    self.assertNotIn("patient", text.lower())
                    self.assertNotIn("患者", text)
                    self.assertNotIn("病例", text)


class OperationRunnerTests(unittest.TestCase):
    def test_one_safe_powershell_call_and_optional_exports(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            plan, source, _candidate = _build_plan(root)
            plan_path = write_operation_plan(root / "operation_plan.json", plan)
            script = root / "apply.ps1"
            script.write_text("# synthetic mock only\n", encoding="utf-8")
            output = root / "output" / "updated.pptx"
            preview = root / "preview"
            pdf = root / "output" / "updated.pdf"

            def mock_apply(command, **_kwargs):
                output_path = Path(command[command.index("-OutputPptx") + 1])
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(source.read_bytes())
                preview_path = Path(command[command.index("-PreviewDir") + 1])
                preview_path.mkdir(parents=True, exist_ok=True)
                for index in range(7):
                    (preview_path / f"slide_{index + 1:03d}.png").write_bytes(b"png")
                pdf_path = Path(command[command.index("-ExportPdf") + 1])
                pdf_path.write_bytes(b"pdf")
                return subprocess.CompletedProcess(
                    command,
                    0,
                    stdout=(
                        "POWERPOINT_OPEN=PASS\nAPPLIED_OPERATIONS=7\nFINAL_SLIDES=25\n"
                        f"FINAL_ORDER_SHA256={plan['expected_final_order_sha256']}\n"
                        "CHANGED_PREVIEW_COUNT=7\nPDF_EXPORT=PASS\n"
                    ),
                    stderr="",
                )

            # Mock output must structurally contain the planned final count.
            def mock_apply_with_count(command, **kwargs):
                _write_minimal_deck(Path(command[command.index("-OutputPptx") + 1]), 25)
                preview_path = Path(command[command.index("-PreviewDir") + 1])
                preview_path.mkdir(parents=True, exist_ok=True)
                for index in range(7):
                    (preview_path / f"slide_{index + 1:03d}.png").write_bytes(b"png")
                Path(command[command.index("-ExportPdf") + 1]).write_bytes(b"pdf")
                return mock_apply(command, **kwargs)._replace(args=command)

            # Avoid the helper overwriting its own synthetic output.
            def final_mock(command, **_kwargs):
                _write_minimal_deck(Path(command[command.index("-OutputPptx") + 1]), 25)
                preview_path = Path(command[command.index("-PreviewDir") + 1])
                preview_path.mkdir(parents=True, exist_ok=True)
                for index in range(7):
                    (preview_path / f"slide_{index + 1:03d}.png").write_bytes(b"png")
                Path(command[command.index("-ExportPdf") + 1]).write_bytes(b"pdf")
                return subprocess.CompletedProcess(
                    command,
                    0,
                    stdout=(
                        "POWERPOINT_OPEN=PASS\nAPPLIED_OPERATIONS=7\nFINAL_SLIDES=25\n"
                        f"FINAL_ORDER_SHA256={plan['expected_final_order_sha256']}\n"
                        "CHANGED_PREVIEW_COUNT=7\nPDF_EXPORT=PASS\n"
                    ),
                    stderr="",
                )

            with patch("academic_ppt.incremental.subprocess.run", side_effect=final_mock) as mocked:
                result = run_powerpoint_operations(
                    source_pptx=source,
                    output_pptx=output,
                    operation_plan_path=plan_path,
                    script_path=script,
                    preview_dir=preview,
                    export_pdf=pdf,
                    timeout_seconds=30,
                )
            self.assertEqual(mocked.call_count, 1)
            command = mocked.call_args.args[0]
            self.assertEqual(command[:6], [
                "powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File"
            ])
            self.assertIn("-PreviewDir", command)
            self.assertIn("-ExportPdf", command)
            self.assertEqual(result["changed_preview_count"], 7)
            self.assertTrue(result["pdf_exported"])
            self.assertTrue(result["source_unchanged"])


class FastQATests(unittest.TestCase):
    def test_changed_slide_gate_ignores_unchanged_issues_but_requires_geometry(self):
        changed = ["SLD-NEW", "SLD-REPLACED"]
        specs = [
            {"slide_id": item, "source_bindings": [{"source_id": "SRC-SYN"}]}
            for item in changed
        ]
        report = run_changed_slide_qa(
            changed_slide_ids=changed,
            slide_specs=specs,
            source_binding_issues=[
                {"slide_id": "SLD-UNCHANGED", "severity": "critical", "code": "OLD"}
            ],
            scientific_issues=[],
            powerpoint_geometry={
                "slides": [
                    {"slide_id": slide_id, "status": "PASS", "issues": []}
                    for slide_id in changed
                ]
            },
        )
        self.assertEqual(report["status"], "PASS")
        self.assertEqual(report["draft_gate"], "DRAFT_READY")
        self.assertEqual(report["ignored_unchanged_source_issues"], 1)

        failed = run_changed_slide_qa(
            changed_slide_ids=changed,
            slide_specs=specs,
            scientific_issues=[
                {"slide_id": "SLD-NEW", "severity": "critical", "code": "SYNTHETIC_SCIENCE"}
            ],
            powerpoint_geometry={"slides": [{"slide_id": "SLD-REPLACED", "status": "PASS"}]},
        )
        self.assertEqual(failed["status"], "FAIL")
        self.assertTrue(any(row["code"] == "POWERPOINT_GEOMETRY_MISSING" for row in failed["issues"]))

    def test_lightweight_qa_checks_order_contract_without_deep_or_libreoffice(self):
        with tempfile.TemporaryDirectory() as raw:
            deck = Path(raw) / "synthetic.pptx"
            _write_minimal_deck(deck, 3)
            expected_ids = ["SLD-A", "SLD-B", "SLD-C"]
            import hashlib

            order_hash = hashlib.sha256("\n".join(expected_ids).encode("utf-8")).hexdigest()
            result = run_whole_deck_lightweight_qa(
                deck,
                expected_slide_count=3,
                expected_final_order=expected_ids,
                operation_result={
                    "powerpoint_opened": True,
                    "final_slide_count": 3,
                    "final_order_sha256": order_hash,
                },
            )
            self.assertEqual(result["status"], "PASS")
            self.assertFalse(result["libreoffice_invoked"])
            self.assertFalse(result["deep_ooxml_invoked"])
            self.assertFalse(result["full_deck_textframe2_scan_invoked"])


class PowerShellSafetyTests(unittest.TestCase):
    def test_script_parses_and_has_no_network_or_libreoffice(self):
        script = ROOT / "scripts" / "apply_pptx_operations.ps1"
        escaped = str(script).replace("'", "''")
        command = (
            "$tokens=$null; $errors=$null; "
            f"[System.Management.Automation.Language.Parser]::ParseFile('{escaped}', "
            "[ref]$tokens, [ref]$errors) | Out-Null; "
            "if ($errors.Count -gt 0) { $errors | ForEach-Object { Write-Error $_.Message }; exit 1 }"
        )
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
            check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        text = script.read_text(encoding="utf-8").lower()
        self.assertEqual(text.count("new-object -comobject powerpoint.application"), 1)
        self.assertIn("insertfromfile", text)
        self.assertIn("changed_preview_count", text)
        for prohibited in (
            "invoke-webrequest",
            "invoke-restmethod",
            "install-module",
            "install-package",
            "soffice",
            "libreoffice",
            "start-process",
            "git ",
        ):
            self.assertNotIn(prohibited, text)


if __name__ == "__main__":
    unittest.main()
