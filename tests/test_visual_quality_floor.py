"""Synthetic-only tests of the process floor. No real deck is generated or mutated."""
from __future__ import annotations

import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from pypdf import PdfWriter

from scripts.academic_ppt.visual_quality_floor import (
    ART_FIELDS, CONTRACTS, OUTPUTS, ROOT, digest, evaluate, read_json, reference,
    sha256, write_outputs,
)


class QualityFloorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="synthetic-floor-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.manifest = self.fixture()

    def save(self, name, value):
        path = self.root / name
        path.write_text(json.dumps(value, indent=2)+"\n", encoding="utf8")
        return path

    def fixture(self, families=None, image_slide=None, same_structure=False):
        families = families or ["HERO_SCENIC", "CONTEXT_CONTRAST", "PROCESS_HORIZONTAL", "EDITORIAL_TAKEAWAY"]
        n = len(families)
        deck = Presentation()
        deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
        image = self.root / "synthetic-image.png"
        Image.new("RGB", (40, 40), "green").save(image)
        for i in range(n):
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            box = slide.shapes.add_textbox(Inches(1 if same_structure else .5+i), Inches(1), Inches(3), Inches(1))
            box.text = "SYNTHETIC: " + str(i)
            slide.notes_slide.notes_text_frame.text = "[Sources]\nSYNTHETIC_TEST_SOURCE (no real science)"
            if image_slide == i+1:
                slide.shapes.add_picture(str(image), Inches(8), Inches(1), Inches(2), Inches(2))
        deck_path = self.root / "synthetic.pptx"
        deck.save(deck_path)
        deck_sha = sha256(deck_path)
        evidence = {"deck_sha256": deck_sha, "content_sha256": "c"*64,
                    "geometry_pages": [{"slide_index": i+1} for i in range(n)],
                    "scientific": "PASS", "semantic": "PASS", "geometry": "PASS", "typography": "PASS"}
        evidence_path = self.save("qa.json", evidence)
        pdf = PdfWriter()
        for _ in range(n):
            pdf.add_blank_page(width=960, height=540)
        with (self.root / "synthetic.pdf").open("wb") as stream:
            pdf.write(stream)
        previews = []
        for i in range(n):
            path = self.root / f"page-{i+1}.png"
            Image.new("RGB", (80, 45), (i*20, 100, 100)).save(path)
            previews.append({**reference(path), "slide_index": i+1})
        sheet_path = self.root / "sheet.png"
        Image.new("RGB", (80, 45*n), "white").save(sheet_path)
        records, observations = [], []
        with zipfile.ZipFile(deck_path) as package:
            for i, family in enumerate(families, 1):
                sid = f"S{i:03d}"
                generic = family.startswith("GENERIC") or family == "TEXT_ONLY"
                records.append({"slide_index": i, "slide_id": sid, "page_role": "CONTEXT",
                                "semantic_mapping": {"mapping_status": "SOURCE_BOUND"},
                                "composition_family": {"observed": family, "generic_fallback": generic,
                                    "fallback_reason": "SYNTHETIC: required parallel text retained" if generic else None,
                                    "fallback_reason_provenance": "SYNTHETIC_ORIGINAL_SELECTOR" if generic else "NOT_APPLICABLE",
                                    "candidate_evaluation": [{"family_id": "PAIRED_COLUMNS", "eligible": False, "reasons": ["SYNTHETIC constraint"]}]},
                                "art_direction_execution": {"recorded": True},
                                "illustration_decision": {"slide_id": sid, "illustration_needed": False,
                                    "illustration_role": "NONE", "selected_source": "NONE", "scientific_evidence": False,
                                    "asset_ids": [], "decision_reason": "SYNTHETIC: no image needed"},
                                "visual_QA": [{"check_id": "Q02", "severity": "BLOCK", "result": "PASS"}]})
                import hashlib
                observations.append({"slide_id": sid, "slide_index": i, "observed_family": family,
                                     "slide_sha256": hashlib.sha256(package.read(f"ppt/slides/slide{i}.xml")).hexdigest()})
        observed_path = self.save("observations.json", observations)
        brief = self.save("brief.json", {"approval": {"status": "approved"}, "direction": {
            **{f: "SYNTHETIC approved direction" for f in ART_FIELDS if f not in ("palette", "approved_assets")},
            "primary_palette": ["00FF00"], "secondary_palette": []}, "approved_asset_ids": []})
        manifest = {"deck": reference(deck_path), "project_id": "SYNTHETIC_ONLY", "records": records,
                    "deck_content_binding": reference(evidence_path, "/content_sha256"),
                    "scientific_inputs": [reference(evidence_path)],
                    "observations": [{"slide_index": i+1, "evidence": reference(observed_path, f"/{i}"),
                                      "preview_sha256": previews[i]["sha256"]} for i in range(n)],
                    "qa": {k: {"status": reference(evidence_path, "/"+k), "deck_binding": reference(evidence_path, "/deck_sha256")}
                           for k in ("scientific", "semantic", "geometry", "typography")},
                    "render": {"deck_binding": reference(evidence_path, "/deck_sha256"), "previews": previews,
                               "contact_sheet": {**reference(sheet_path), "slide_indexes": list(range(1, n+1))},
                               "pdf": reference(self.root / "synthetic.pdf")},
                    "assets": [], "art_direction": {"brief": reference(brief), "fields": []}, "human_reviews": {}}
        manifest["qa"]["geometry"]["pages"] = reference(evidence_path, "/geometry_pages")
        return manifest

    def evaluate(self, manifest=None):
        return evaluate(manifest or self.manifest, mode="shadow")

    def codes(self, report):
        return {f["code"] for f in report["findings"]}

    def assert_block(self, code, manifest=None):
        report = self.evaluate(manifest)
        self.assertEqual(report["status"], "VISUAL_QUALITY_BLOCKED")
        self.assertIn(code, self.codes(report))
        return report

    def change_qa(self, field, value):
        path = self.root / "qa.json"
        data = read_json(path)
        data[field] = value
        self.save("qa.json", data)
        # Refresh all file references, keeping the changed real evidence value.
        def update(v):
            if isinstance(v, dict):
                if v.get("path") == str(path.resolve()):
                    v["sha256"] = sha256(path)
                for child in v.values():
                    update(child)
            elif isinstance(v, list):
                for child in v:
                    update(child)
        update(self.manifest)

    def test_missing_page_role(self):
        self.manifest["records"][0]["page_role"] = None
        self.assert_block("PAGE_ROLE_UNRESOLVED")

    def test_multiple_primary_roles(self):
        self.manifest["records"][0]["page_role"] = ["HERO", "CONTEXT"]
        self.assert_block("PAGE_ROLE_UNRESOLVED")

    def test_missing_illustration_decision(self):
        self.manifest["records"][0].pop("illustration_decision")
        self.assert_block("MISSING_ILLUSTRATION_DECISION")

    def test_string_boolean_is_not_decision(self):
        self.manifest["records"][0]["illustration_decision"]["illustration_needed"] = "false"
        self.assert_block("MISSING_ILLUSTRATION_DECISION")

    def test_silent_generic_fallback(self):
        self.manifest = self.fixture(["GENERIC_CARD_GRID", "HERO_SCENIC"])
        self.manifest["records"][0]["composition_family"]["fallback_reason"] = ""
        self.assert_block("SILENT_GENERIC_FALLBACK")

    def test_authorized_generic_fallback_visible(self):
        self.manifest = self.fixture(["GENERIC_CARD_GRID", "HERO_SCENIC"])
        report = self.evaluate()
        self.assertNotEqual(report["status"], "VISUAL_QUALITY_BLOCKED")
        self.assertTrue(report["slide_records"][0]["fallback_used"])
        self.assertEqual(report["slide_records"][0]["generic_fallback_status"], "EXPLICIT_JUSTIFIED")

    def test_three_consecutive_generic_cards_warn_not_block(self):
        report = self.evaluate(self.fixture(["GENERIC_CARD_3", "GENERIC_PANEL", "GENERIC_CARD_4", "HERO_SCENIC"]))
        self.assertEqual(report["summary"]["max_consecutive_generic_cards"], 3)
        self.assertIn("CONSECUTIVE_GENERIC_CARDS", self.codes(report))
        self.assertNotEqual(report["status"], "VISUAL_QUALITY_BLOCKED")

    def test_four_consecutive_generic_cards(self):
        report = self.evaluate(self.fixture(["GENERIC_CARD_GRID"]*4))
        self.assertEqual(report["summary"]["max_consecutive_generic_cards"], 4)
        self.assertIn("MACRO_COMPOSITION_REPETITION", self.codes(report))
        self.assertNotEqual(report["status"], "VISUAL_QUALITY_BLOCKED")

    def test_non_consecutive_repetition(self):
        report = self.evaluate(self.fixture(["HERO_SCENIC", "CONTEXT_CONTRAST", "HERO_SCENIC", "EDITORIAL_TAKEAWAY"]))
        self.assertEqual(report["summary"]["non_consecutive_repeated_families"]["HERO_SCENIC"], [1, 3])

    def test_same_structure_different_labels(self):
        report = self.evaluate(self.fixture(same_structure=True))
        self.assertEqual(list(report["summary"]["same_structure_different_labels"].values()), [[1, 2, 3, 4]])
        self.assertIn("CONSECUTIVE_NEAR_IDENTICAL_STRUCTURE", self.codes(report))

    def test_unauthorized_asset(self):
        report = self.assert_block("UNAUTHORIZED_PRESENTATION_ASSET", self.fixture(image_slide=1))
        self.assertEqual(report["summary"]["illustration_page_count"], 1)

    def add_approved_asset(self):
        self.manifest = self.fixture(image_slide=1)
        from scripts.academic_ppt.ooxml_qa import inspect_pptx_ooxml
        package = inspect_pptx_ooxml(self.manifest["deck"]["path"])
        rel = next(r for r in package["asset_relationships"] if r["category"] == "media")
        media = [{"part": rel["target_part"], "sha256": rel["target_sha256"]}]
        source_hash = sha256(self.root / "synthetic-image.png")
        event = {"asset_id": "SYNTHETIC_ASSET", "asset_sha256": source_hash, "status": "APPROVED",
                 "approved_by": "SYNTHETIC_TEST_REVIEWER", "approved_at": "2026-01-01T00:00:00Z",
                 "project_id": "SYNTHETIC_ONLY", "deck_content_sha256": "c"*64, "slide_ids": ["S001"],
                 "placement_role": "CONCEPT_SUPPORT", "allowed_transforms": ["NONE"],
                 "reviewed_preview_sha256": self.manifest["render"]["previews"][0]["sha256"],
                 "decision_evidence_ref": "SYNTHETIC_TEST_ONLY", "scientific_evidence": False, "presentation_only": True}
        event_path = self.save("approval.json", event)
        events = self.save("approval_events.json", [event])
        placement = {"asset_id": "SYNTHETIC_ASSET", "slide_id": "S001", "media": media, "source_sha256": source_hash}
        placement_path = self.save("placement.json", placement)
        self.manifest["assets"] = [{"asset_id": "SYNTHETIC_ASSET", "slide_id": "S001", "media": media,
                                   "source_type": "APPROVED_PROJECT_ASSET", "source_sha256": source_hash, "transform": "NONE",
                                   "authorization": reference(event_path), "approval_events": reference(events),
                                   "source_asset": reference(self.root / "synthetic-image.png"),
                                   "placement": reference(placement_path)}]
        self.manifest["records"][0]["illustration_decision"].update(
            illustration_needed=True, illustration_role="CONCEPT_SUPPORT", selected_source="APPROVED_PROJECT_ASSET",
            asset_ids=["SYNTHETIC_ASSET"], illustration_subject="Synthetic green square")
        return event, event_path

    def test_approved_asset_exact_slide(self):
        self.add_approved_asset()
        self.assertNotEqual(self.evaluate()["status"], "VISUAL_QUALITY_BLOCKED")

    def test_approved_asset_wrong_slide(self):
        event, path = self.add_approved_asset()
        event["slide_ids"] = ["S002"]
        self.save(path.name, event)
        self.manifest["assets"][0]["authorization"] = reference(path)
        events = self.save("approval_events.json", [event])
        self.manifest["assets"][0]["approval_events"] = reference(events)
        self.assert_block("UNAUTHORIZED_PRESENTATION_ASSET")

    def test_revoked_asset(self):
        event, path = self.add_approved_asset()
        event["status"] = "REVOKED"
        events = self.save("approval_events.json", [read_json(path), event])
        self.manifest["assets"][0]["approval_events"] = reference(events)
        self.assert_block("UNAUTHORIZED_PRESENTATION_ASSET")

    def test_asset_evidence_flag(self):
        event, path = self.add_approved_asset()
        event["scientific_evidence"] = True
        self.save(path.name, event)
        self.manifest["assets"][0]["authorization"] = reference(path)
        events = self.save("approval_events.json", [event])
        self.manifest["assets"][0]["approval_events"] = reference(events)
        self.assert_block("UNSAFE_SCIENTIFIC_LOOKING_ILLUSTRATION")

    def test_asset_transform_scope(self):
        self.add_approved_asset()
        self.manifest["assets"][0]["transform"] = "FLIP_LATERALITY"
        self.assert_block("UNAUTHORIZED_PRESENTATION_ASSET")

    def test_asset_source_unresolved(self):
        self.add_approved_asset()
        self.manifest["assets"][0]["source_asset"]["path"] = str(self.root / "missing.png")
        self.assert_block("UNAUTHORIZED_PRESENTATION_ASSET")

    def test_asset_hash_changed(self):
        self.add_approved_asset()
        self.manifest["assets"][0]["media"][0]["sha256"] = "0"*64
        self.assert_block("UNAUTHORIZED_PRESENTATION_ASSET")

    def test_generated_asset_without_safety_review(self):
        self.add_approved_asset()
        self.manifest["assets"][0]["source_type"] = "AI_CONCEPTUAL_CANDIDATE"
        self.assert_block("GENERATED_ASSET_SAFETY_NOT_VERIFIED")

    def test_geometry_partial_scope(self):
        self.change_qa("geometry_pages", [{"slide_index": 1}])
        self.assert_block("GEOMETRY_PAGE_COVERAGE_INCOMPLETE")

    def test_required_source_figure_cannot_be_missing(self):
        self.manifest["records"][0]["illustration_decision"].update(
            illustration_needed=True, selected_source="EXISTING_SCIENTIFIC_SOURCE_FIGURE",
            illustration_role="EVIDENCE_DISPLAY", scientific_evidence=True)
        self.assert_block("REQUIRED_SOURCE_FIGURE_MISSING")

    def approve_synthetic_human_reviews(self):
        report = self.evaluate()
        signature = {"status": "APPROVED", "reviewer": "SYNTHETIC_ONLY", "reviewed_at": "2026-01-01T00:00:00Z",
                     "deck_sha256": report["deck_sha256"], "render_set_sha256": report["render_set_sha256"],
                     "slide_ids": [r["slide_id"] for r in report["slide_records"]],
                     "scientific_approval": True, "visual_approval": True}
        path = self.save("synthetic_human_review.json", signature)
        self.manifest["human_reviews"] = {c["check_id"]: reference(path)
                                          for c in report["checks"] if c["severity"] == "HUMAN_REVIEW"}

    def test_signed_review_with_warnings_status(self):
        self.approve_synthetic_human_reviews()
        self.assertEqual(self.evaluate()["status"], "VISUAL_QUALITY_PASS_WITH_WARNINGS")

    def test_signed_review_stale_render_not_accepted(self):
        self.approve_synthetic_human_reviews()
        path = self.root / "synthetic_human_review.json"
        value = read_json(path)
        value["render_set_sha256"] = "0"*64
        self.save(path.name, value)
        self.manifest["human_reviews"] = {k: reference(path) for k in self.manifest["human_reviews"]}
        self.assertEqual(self.evaluate()["status"], "VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED")

    def test_all_evidence_and_synthetic_signatures_pass(self):
        facts = self.save("synthetic_art_receipts.json", {"consumer": "SYNTHETIC_APPLIED", "render": "SYNTHETIC_OBSERVED"})
        for field in ART_FIELDS:
            self.manifest["art_direction"]["fields"].append({"field": field,
                "slide_ids": [r["slide_id"] for r in self.manifest["records"]],
                "consumer": "SYNTHETIC_TEST_CONSUMER", "consumer_evidence": reference(facts, "/consumer"),
                "consumer_value": "SYNTHETIC_APPLIED", "render_evidence": reference(facts, "/render"),
                "render_value": "SYNTHETIC_OBSERVED", "render_deck_sha256": self.manifest["deck"]["sha256"]})
        self.approve_synthetic_human_reviews()
        self.assertEqual(self.evaluate()["status"], "VISUAL_QUALITY_PASS")

    def test_art_direction_metadata_is_not_execution(self):
        report = self.evaluate()
        self.assertIn("ART_DIRECTION_NOT_EXECUTED_OR_UNVERIFIED", self.codes(report))
        self.assertTrue(all(r["consumed"] is False and r["observed_in_render"] is False for r in report["art_direction_records"]))
        self.assertEqual(len(report["art_direction_records"]), 40)

    def test_contact_sheet_missing(self):
        self.manifest["render"].pop("contact_sheet")
        self.assert_block("CONTACT_SHEET_MISSING_OR_STALE")

    def test_human_review_only_concern(self):
        report = self.evaluate()
        self.assertEqual(report["status"], "VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED")
        self.assertTrue(report["summary"]["human_review_required"])

    def test_scientific_qa_failure(self):
        self.change_qa("scientific", "FAIL")
        self.assert_block("SCIENTIFIC_QA_FAILURE")

    def test_geometry_failure(self):
        self.change_qa("geometry", "FAIL")
        self.assert_block("GEOMETRY_QA_FAILURE")

    def test_missing_required_qa(self):
        self.manifest["qa"].pop("scientific")
        self.assert_block("SCIENTIFIC_QA_FAILURE")

    def test_scientific_source_drift(self):
        self.manifest["scientific_inputs"][0]["sha256"] = "0"*64
        self.assert_block("SCIENTIFIC_INPUT_EVIDENCE_MISSING")

    def test_severity_cannot_be_downgraded(self):
        self.manifest["records"][0]["visual_QA"] = [
            {"check_id": "Q02", "severity": "WARN", "result": "FAIL", "basis": "SYNTHETIC regression"}]
        report = self.assert_block("INHERITED_Q02")
        self.assertEqual(next(f["severity"] for f in report["findings"] if f["code"] == "INHERITED_Q02"), "BLOCK")

    def test_records_cannot_shrink_denominator(self):
        self.manifest["records"].pop()
        report = self.assert_block("SLIDE_RECORD_COVERAGE_INVALID")
        self.assertEqual(report["summary"]["slide_count"], 4)
        self.assertEqual(len(report["slide_records"]), 4)

    def test_duplicate_ids_block(self):
        self.manifest["records"][1]["slide_id"] = self.manifest["records"][0]["slide_id"]
        self.assert_block("SLIDE_RECORD_COVERAGE_INVALID")

    def test_declared_family_cannot_replace_observation(self):
        self.manifest["observations"] = []
        self.assert_block("ACTUAL_COMPOSITION_NOT_VERIFIED")

    def test_stale_observation_blocks(self):
        self.manifest["observations"][0]["preview_sha256"] = "0"*64
        self.assert_block("ACTUAL_COMPOSITION_NOT_VERIFIED")

    def test_retrospective_reason_not_original_receipt(self):
        self.manifest = self.fixture(["GENERIC_CARD_GRID", "HERO_SCENIC"])
        self.manifest["records"][0]["composition_family"]["fallback_reason_provenance"] = "RETROSPECTIVE_AUDIT_NOT_ORIGINAL_SELECTION"
        self.assert_block("ORIGINAL_FALLBACK_RECEIPT_NOT_ASSESSED")

    def test_semantic_limit_not_repaired(self):
        for status in ("NEEDS_AUTHOR_DECISION", "NO_SAFE_MAPPING"):
            self.manifest["records"][0]["semantic_mapping"]["mapping_status"] = status
            report = self.assert_block("SEMANTIC_MAPPING_UNRESOLVED")
            self.assertEqual(report["slide_records"][0]["semantic_mapping_status"], status)
            self.assertEqual(report["slide_records"][0]["illustration_decision"], "NO_ILLUSTRATION_NEEDED")
            self.assertEqual(report["slide_records"][0]["design_record"]["illustration_decision"],
                             self.manifest["records"][0]["illustration_decision"])

    def test_default_off_does_not_read_or_write(self):
        self.assertIsNone(evaluate(None))
        output = self.root / "off"
        p = subprocess.run([sys.executable, "-B", "-m", "scripts.academic_ppt.visual_quality_floor",
                            "--manifest", "DOES_NOT_EXIST.json", "--output", str(output)], cwd=ROOT, capture_output=True)
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertFalse(output.exists())
        self.assertIn(b"DEFAULT_OFF", p.stdout)

    def test_enforcement_preserves_findings_and_status(self):
        shadow = self.evaluate()
        enforced = evaluate(self.manifest, mode="enforce")
        self.assertEqual(enforced["status"], shadow["status"])
        self.assertEqual(enforced["findings"], shadow["findings"])
        self.assertTrue(enforced["production_active"])

    def test_no_input_mutation_and_eight_outputs(self):
        original = copy.deepcopy(self.manifest)
        hashes = {str(p): sha256(p) for p in self.root.iterdir() if p.is_file()}
        report = self.evaluate()
        self.assertEqual(self.manifest, original)
        for p, expected in hashes.items():
            self.assertEqual(sha256(p), expected)
        output = self.root / "report"
        write_outputs(report, output)
        self.assertEqual({p.name for p in output.iterdir()}, set(OUTPUTS))
        with self.assertRaises(FileExistsError):
            write_outputs(report, output)

    def test_outputs_cannot_escape_staging(self):
        with self.assertRaisesRegex(ValueError, "staging"):  # self-containment: documentation-only
            write_outputs(self.evaluate(), ROOT / "input/forbidden-test")

    def test_future_illustration_interface_present(self):
        for r in self.evaluate()["slide_records"]:
            for key in ("illustration_decision", "illustration_role", "illustration_subject_summary",
                        "allowed_visual_types", "prohibited_visual_types", "source_priority", "human_approval_required"):
                self.assertIn(key, r)
            self.assertFalse(r["insertion_allowed_now"])

    def test_contract_has_no_private_project_authority(self):
        authority = read_json(CONTRACTS / "authority.json")
        self.assertNotIn("source_package", authority)
        self.assertEqual(authority["authority"], "ASSISTED_VISUAL_QUALITY_CONTRACT")


if __name__ == "__main__":
    unittest.main()
