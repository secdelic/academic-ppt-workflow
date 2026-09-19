"""Synthetic metadata only: no renderer, real project data, or generated assets."""
from __future__ import annotations

import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from scripts.academic_ppt.visual_ambition import (
    ACTIONS, CHANNELS, OUTPUTS, ROOT, SUBJECTS, audit_ambition, plan, read_bundle, write_outputs,
)
from scripts.academic_ppt.visual_quality_floor import digest, reference, sha256


def fixture(n=1):
    slides, contexts = [], []
    for i in range(1, n+1):
        sid = f"S{i:03d}"
        slides.append({"slide_id": sid, "slide_index": i, "page_role": "FRAMEWORK", "composition_family": "FRAMEWORK_HUB",
                       "layout_repetition_family": "FRAMEWORK_HUB", "semantic_mapping_status": "SOURCE_BOUND",
                       "illustration_decision": "NO_ILLUSTRATION_NEEDED", "approved_asset_ids": [],
                       "design_record": {"composition_family": {"selector_inputs": {"content_density": "LOW", "max_text_block_length": 30}}}})
        contexts.append({"slide_id": sid, "narrative_function": "CORE_FRAMEWORK", "communication_need": "ROLE_DISTINCTION",
                         "illustration_potential": "HIGH", "subject_class": "ABSTRACT_MEDICAL_CONCEPT",
                         "space_status": "AVAILABLE_CANDIDATE", "reference_scene_opportunity": False,
                         "source_figure_required": False, "requested_visual_type": None,
                         "microcopy_requires_protected_rewrite": False, "repeated_wording": False,
                         "negative_space_hint": None, "evidence_refs": [{"source": "SYNTHETIC_ONLY"}]})
    return {"floor_report": {"summary": {"slide_count": n}, "status": "VISUAL_QUALITY_PASS", "findings": []},
            "floor_records": slides, "contexts": contexts, "third_party_upload_allowed": False,
            "visual_brief": {"approval": {"status": "approved"}, "direction": {"direction_name": "Editorial / Infographic / Warm-academic", "primary_palette": ["#5B7F72"], "secondary_palette": []}},
            "independent_real_projects_evaluated": 0}


class AmbitionTests(unittest.TestCase):
    def setUp(self):
        self.bundle = fixture()

    def run_plan(self):
        return plan(self.bundle, enabled=True)

    def record(self):
        return self.run_plan()["records"][0]

    def block(self, check="Q02", sid="S001"):
        self.bundle["floor_report"]["status"] = "VISUAL_QUALITY_BLOCKED"
        self.bundle["floor_report"]["findings"].append({"check_id": check, "severity": "BLOCK", "slide_id": sid,
                                                     "code": "SYNTHETIC_BLOCK", "result": "NOT_ASSESSED", "basis": "SYNTHETIC_ONLY"})

    def test_pass_floor_high_opportunity(self):
        r = self.record()
        self.assertEqual(r["quality_floor_status"], "PASS")
        self.assertEqual(r["enrichment_opportunity"], "HIGH")
        self.assertEqual(r["visual_ambition_status"], "ENRICHMENT_RECOMMENDED")

    def test_pass_floor_no_enrichment_needed(self):
        self.bundle["contexts"][0].update(communication_need="ALREADY_CLEAR", illustration_potential="NONE", subject_class="NONE")
        r = self.record()
        self.assertEqual(r["quality_floor_status"], "PASS")
        self.assertEqual(r["visual_ambition_status"], "NO_MEANINGFUL_ENRICHMENT_NEEDED")
        self.assertEqual(r["recommended_next_action"], "NO_ACTION")
        self.assertEqual(r["primary_enrichment_channels"], [])

    def test_block_floor_high_opportunity_remains_block(self):
        self.block()
        result = self.run_plan()
        self.assertEqual(result["quality_floor_status"], "VISUAL_QUALITY_BLOCKED")
        self.assertEqual(result["records"][0]["enrichment_opportunity"], "HIGH")
        self.assertEqual(result["records"][0]["execution_eligibility"], "HOLD_FLOOR_BLOCK")

    def test_deck_block_holds_other_slides(self):
        self.bundle = fixture(2)
        self.block(sid="S001")
        result = self.run_plan()
        self.assertEqual(result["records"][1]["quality_floor_status"], "PASS")
        self.assertEqual(result["handoff_records"][1]["deck_quality_floor_status"], "VISUAL_QUALITY_BLOCKED")
        self.assertEqual(result["handoff_records"][1]["execution_eligibility"], "HOLD_DECK_FLOOR_BLOCK")

    def test_no_illustration_needed_and_high_enrichment(self):
        r = self.record()
        self.assertEqual(r["illustration_decision"], "NO_ILLUSTRATION_NEEDED")
        self.assertEqual(r["illustration_enrichment_opportunity"], "HIGH")
        self.assertTrue(r["illustration_candidate_allowed"])

    def test_anchor_framework(self):
        r = self.record()
        self.assertTrue(r["is_anchor"])
        self.assertEqual(r["visual_anchor_role"], "ANCHOR")
        self.assertEqual(r["recommended_ambition_level"], "TARGET")

    def test_non_anchor_utility(self):
        self.bundle["floor_records"][0]["page_role"] = "APPENDIX"
        self.bundle["contexts"][0].update(narrative_function="UTILITY", communication_need="ALREADY_CLEAR",
                                         illustration_potential="NONE", subject_class="NONE")
        r = self.record()
        self.assertFalse(r["is_anchor"])
        self.assertEqual(r["visual_anchor_role"], "UTILITY")
        self.assertEqual(r["recommended_ambition_level"], "FLOOR")

    def test_anchor_not_selected_by_page_number(self):
        self.bundle = fixture(3)
        self.bundle["contexts"][0]["narrative_function"] = "UTILITY"
        rows = self.run_plan()["records"]
        self.assertFalse(rows[0]["is_anchor"])
        self.assertTrue(rows[2]["is_anchor"])

    def test_more_than_two_primary_channels_rejected(self):
        self.bundle["contexts"][0]["override"] = {"primary_enrichment_channels": list(CHANNELS[:3])}
        with self.assertRaisesRegex(ValueError, "INVALID_PRIMARY_ENRICHMENT_CHANNELS"):
            self.run_plan()

    def test_microcopy_unsafe(self):
        self.bundle["contexts"][0]["microcopy_requires_protected_rewrite"] = True
        r = self.record()
        self.assertEqual(r["microcopy_opportunity"], "MICROCOPY_UNSAFE")
        self.assertNotEqual(r["recommended_next_action"], "MICROCOPY_PROPOSAL")
        self.assertIsNone(r["microcopy_text"])

    def test_unsafe_microcopy_override_withheld_explicitly(self):
        self.bundle["contexts"][0].update(microcopy_requires_protected_rewrite=True,
                                         override={"recommended_next_action": "MICROCOPY_PROPOSAL"})
        r = self.record()
        self.assertEqual(r["recommended_next_action"], "ART_DIRECTION_REVIEW")
        self.assertEqual(r["action_override_withheld_by_safety_gate"], "MICROCOPY_PROPOSAL")

    def test_source_figure_required(self):
        self.bundle["contexts"][0]["source_figure_required"] = True
        r = self.record()
        self.assertEqual(r["recommended_next_action"], "SOURCE_FIGURE_REQUIRED")
        self.assertFalse(r["illustration_candidate_allowed"])
        self.assertEqual(r["allowed_visual_types"], [])
        self.assertEqual(r["composition_upgrade_candidate"]["type"], "FIGURE_LED_COMPOSITION")

    def test_ai_prohibited_medical_visual(self):
        self.bundle["contexts"][0]["requested_visual_type"] = "CT"
        r = self.record()
        self.assertEqual(r["scientific_risk"], "PROHIBITED_MEDICAL_VISUAL")
        self.assertFalse(r["illustration_candidate_allowed"])
        self.assertIsNone(r["illustration_subject_summary"])
        self.assertIn("SCIENTIFIC_VISUAL_CONFUSION_RISK", [f["code"] for f in r["ambition_findings"]])

    def test_over_decoration(self):
        self.bundle["floor_records"][0]["approved_asset_ids"] = ["SYNTHETIC_EXISTING_ASSET"]
        self.bundle["contexts"][0]["space_status"] = "TIGHT"
        codes = [f["code"] for f in self.record()["ambition_findings"]]
        self.assertIn("OVER_DECORATION_RISK", codes)

    def test_reference_grade_opportunity_not_achievement(self):
        self.bundle["contexts"][0]["reference_scene_opportunity"] = True
        result = self.run_plan()
        r = result["records"][0]
        self.assertEqual(r["enrichment_opportunity"], "REFERENCE_GRADE_OPPORTUNITY")
        self.assertEqual(r["recommended_ambition_level"], "STRETCH")
        self.assertEqual(r["ambition_level_achieved"], "NOT_ASSESSED")
        self.assertFalse(result["summary"]["reference_grade_achieved"])

    def test_scene_opportunity_does_not_force_illustration(self):
        self.bundle["contexts"][0].update(reference_scene_opportunity=True, illustration_potential="NONE", subject_class="NONE")
        r = self.record()
        self.assertEqual(r["recommended_ambition_level"], "STRETCH")
        self.assertFalse(r["illustration_candidate_allowed"])
        self.assertEqual(r["primary_enrichment_channels"], ["COMPOSITION_UPGRADE", "EDITORIAL_COMPONENTS"])

    def test_conceptual_anatomy_retains_scientific_review(self):
        self.bundle["contexts"][0]["subject_class"] = "ORGAN_SILHOUETTE"
        r = self.record()
        self.assertEqual(r["allowed_visual_types"], ["ORGAN_SILHOUETTE"])
        self.assertIn("SCIENTIFIC_VISUAL_CONFUSION_RISK", [f["code"] for f in r["ambition_findings"]])

    def test_missing_ambition_record(self):
        self.bundle["contexts"] = []
        with self.assertRaisesRegex(ValueError, "MISSING_OR_DUPLICATE_AMBITION_RECORD"):
            self.run_plan()

    def test_duplicate_ambition_record(self):
        self.bundle["contexts"].append(copy.deepcopy(self.bundle["contexts"][0]))
        with self.assertRaisesRegex(ValueError, "MISSING_OR_DUPLICATE_AMBITION_RECORD"):
            self.run_plan()

    def test_default_off_no_read_no_write(self):
        with patch("pathlib.Path.read_text", side_effect=AssertionError("Unexpected file read")), patch(
                "pathlib.Path.write_text", side_effect=AssertionError("Unexpected write")):
            self.assertIsNone(plan(None))
        with tempfile.TemporaryDirectory(prefix="ambition-off-") as folder:
            for module in ("scripts.academic_ppt.visual_ambition",):
                args = [sys.executable, "-B", "-m", module, "--output", str(Path(folder)/"must-not-exist")]
                if module.endswith("visual_ambition"):
                    args += ["--manifest", str(Path(folder)/"missing.json")]
                p = subprocess.run(args, cwd=ROOT, capture_output=True)
                self.assertEqual(p.returncode, 0, p.stderr)
                self.assertEqual(p.stdout, b"")
                self.assertEqual(list(Path(folder).iterdir()), [])

    def test_read_only_floor_inputs(self):
        original = copy.deepcopy(self.bundle)
        result = self.run_plan()
        self.assertEqual(self.bundle, original)
        self.assertEqual(result["quality_floor_report_digest"], digest(original["floor_report"]))
        self.assertEqual(result["quality_floor_records_digest"], digest(original["floor_records"]))

    def test_floor_severity_cannot_change(self):
        self.block()
        self.bundle["floor_report"]["findings"][0]["severity"] = "WARN"
        with self.assertRaisesRegex(ValueError, "VISUAL_AMBITION_REQUIRES_FLOOR_CONTRACT_CHANGE"):
            self.run_plan()

    def test_floor_inconsistent_status_rejected(self):
        self.block()
        self.bundle["floor_report"]["status"] = "VISUAL_QUALITY_PASS"
        with self.assertRaisesRegex(ValueError, "FROZEN_FLOOR_STATUS_INCONSISTENT"):
            self.run_plan()

    def test_semantic_restrictions_preserved(self):
        for restriction in ("NEEDS_AUTHOR_DECISION", "NO_SAFE_MAPPING"):
            self.bundle = fixture()
            self.block("Q03")
            self.bundle["floor_records"][0]["semantic_mapping_status"] = restriction
            self.bundle["contexts"][0]["override"] = {"recommended_next_action": "REFERENCE_GRADE_SCENE_TRIAL"}
            r = self.record()
            self.assertEqual(r["recommended_next_action"], "HUMAN_SEMANTIC_DECISION")
            self.assertEqual(r["semantic_risk"], restriction)
            self.assertFalse(r["illustration_candidate_allowed"])
            self.assertEqual(r["illustration_decision"], "NO_ILLUSTRATION_NEEDED")

    def test_pending_human_review_not_floor_pass(self):
        self.bundle["floor_report"].update(status="VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED", findings=[
            {"check_id": "Q20", "severity": "HUMAN_REVIEW", "result": "PENDING", "slide_id": None,
             "code": "SYNTHETIC_HUMAN_PENDING", "basis": "SYNTHETIC_ONLY"}])
        result = self.run_plan()
        self.assertEqual(result["records"][0]["quality_floor_status"], "HUMAN_REVIEW")
        self.assertEqual(result["summary"]["full_floor_pass_with_high_opportunity"], [])
        self.assertEqual(result["summary"]["hard_checks_pass_with_high_opportunity"], ["S001"])

    def test_failure_mode_ambition_too_low_and_anchor(self):
        self.bundle["contexts"][0]["override"] = {"recommended_ambition_level": "FLOOR", "primary_enrichment_channels": []}
        codes = {f["code"] for f in self.record()["ambition_findings"]}
        self.assertTrue({"AMBITION_TOO_LOW", "ANCHOR_SLIDE_UNDERDESIGNED", "ILLUSTRATION_OPPORTUNITY_IGNORED"} <= codes)

    def test_failure_mode_microcopy_ignored(self):
        self.bundle["floor_records"][0]["design_record"]["composition_family"]["selector_inputs"]["max_text_block_length"] = 80
        self.bundle["contexts"][0]["override"] = {"primary_enrichment_channels": ["COMPOSITION_UPGRADE"], "secondary_enrichment_channels": []}
        self.assertIn("MICROCOPY_OPPORTUNITY_IGNORED", [f["code"] for f in self.record()["ambition_findings"]])

    def test_failure_mode_reference_scene_ignored(self):
        self.bundle["contexts"][0].update(reference_scene_opportunity=True, override={"recommended_ambition_level": "TARGET"})
        self.assertIn("REFERENCE_SCENE_OPPORTUNITY_IGNORED", [f["code"] for f in self.record()["ambition_findings"]])

    def test_failure_mode_every_page_same(self):
        self.bundle = fixture(3)
        self.bundle["contexts"][1]["narrative_function"] = "MAJOR_TIMELINE"
        result = self.run_plan()
        self.assertIn("EVERY_PAGE_TREATED_THE_SAME", [f["code"] for f in result["ambition_findings"]])

    def test_ambition_findings_cannot_block_or_override_floor(self):
        self.bundle["contexts"][0]["space_status"] = "TIGHT"
        report = self.run_plan()
        self.assertTrue(all(f["severity"] in ("WARN", "HUMAN_REVIEW") and not f["floor_override_allowed"] for f in report["ambition_findings"]))
        self.assertEqual(report["quality_floor_status"], "VISUAL_QUALITY_PASS")

    def test_b2_interface_does_not_copy_raw_private_content(self):
        token = "SYNTHETIC_PRIVATE_PATIENT_IDENTIFIER_12345"
        self.bundle["floor_records"][0]["title"] = token
        self.bundle["floor_records"][0]["illustration_subject_summary"] = token
        self.bundle["visual_brief"]["workspace_hint"] = token
        handoff = self.run_plan()["handoff_records"]
        self.assertNotIn(token, json.dumps(handoff))
        self.assertIn(handoff[0]["illustration_subject_summary"], SUBJECTS.values())
        self.assertFalse(handoff[0]["is_generation_request"])
        self.assertFalse(handoff[0]["external_export_approved"])

    def test_freeform_summary_in_context_rejected(self):
        self.bundle["contexts"][0]["illustration_subject_summary"] = "SYNTHETIC_PRIVATE_CONTENT"
        with self.assertRaisesRegex(ValueError, "INVALID_AMBITION_CONTEXT_FIELDS"):
            self.run_plan()

    def test_missing_illustration_subject_is_not_invented(self):
        self.bundle["contexts"][0]["subject_class"] = "NONE"
        with self.assertRaisesRegex(ValueError, "ILLUSTRATION_SUBJECT_CLASS_REQUIRED"):
            self.run_plan()

    def test_private_content_in_role_rejected(self):
        self.bundle["floor_records"][0]["page_role"] = "SYNTHETIC_PRIVATE_CONTENT"
        with self.assertRaisesRegex(ValueError, "unresolved primary page role"):
            self.run_plan()

    def test_unsafe_slide_id_rejected(self):
        self.bundle["floor_records"][0]["slide_id"] = "PATIENT_EXAMPLE"
        self.bundle["contexts"][0]["slide_id"] = "PATIENT_EXAMPLE"
        with self.assertRaisesRegex(ValueError, "UNSAFE_HANDOFF_SLIDE_ID"):
            self.run_plan()

    def test_palette_cannot_carry_raw_text(self):
        self.bundle["visual_brief"]["direction"]["primary_palette"] = ["SYNTHETIC_PRIVATE_CONTENT"]
        with self.assertRaisesRegex(ValueError, "INVALID_HANDOFF_PALETTE"):
            self.run_plan()

    def test_unapproved_benchmark_is_not_invented(self):
        self.bundle["visual_brief"]["approval"]["status"] = "pending"
        with self.assertRaisesRegex(ValueError, "APPROVED_VISUAL_BRIEF_REQUIRED"):
            self.run_plan()

    def test_negative_space_no_unverified_free_text(self):
        self.bundle["contexts"][0]["negative_space_hint"] = "Place over a clinical measurement"
        with self.assertRaisesRegex(ValueError, "INVALID_AMBITION_CONTEXT_ENUM"):
            self.run_plan()

    def test_future_interface_fields_and_fixed_source_priority(self):
        r = self.run_plan()["handoff_records"][0]
        self.assertEqual(r["source_priority"], ["EXISTING_SCIENTIFIC_SOURCE_FIGURE", "APPROVED_PROJECT_ASSET",
            "APPROVED_DETERMINISTIC_LIBRARY", "AI_CONCEPTUAL_CANDIDATE", "NONE"])
        for key in ("slide_id", "page_role", "illustration_enrichment_opportunity", "illustration_role_candidate",
                    "illustration_subject_summary", "allowed_visual_types", "prohibited_visual_types",
                    "visual_brief_style_family", "palette", "negative_space_hint", "human_approval_required"):
            self.assertIn(key, r)
        self.assertTrue(r["human_approval_required"])
        self.assertFalse(r["generation_authorized"])

    def test_output_contract_and_no_overwrite(self):
        with tempfile.TemporaryDirectory(prefix="synthetic-ambition-") as folder:
            out = Path(folder) / "plan"
            write_outputs(self.run_plan(), out)
            self.assertEqual({p.name for p in out.iterdir()}, set(OUTPUTS))
            with self.assertRaises(FileExistsError):
                write_outputs(self.run_plan(), out)

    def test_output_path_protection(self):
        with self.assertRaisesRegex(ValueError, "staging"):  # self-containment: documentation-only
            write_outputs(self.run_plan(), ROOT / "input/forbidden-ambition-output")

    def test_manifest_reader_hash_and_privacy_evidence(self):
        with tempfile.TemporaryDirectory(prefix="synthetic-ambition-input-") as folder:
            root = Path(folder)
            def save(name, value):
                path = root / name
                path.write_text(json.dumps(value), encoding="utf8")
                return reference(path)
            floor = save("floor.json", self.bundle["floor_report"])
            records = save("records.json", self.bundle["floor_records"])
            self.bundle["contexts"][0]["evidence_refs"] = [{**records, "pointer": "/0"}]
            contexts = save("contexts.json", self.bundle["contexts"])
            brief = save("brief.json", self.bundle["visual_brief"])
            privacy = save("privacy.json", {"privacy": {"third_party_upload_allowed": False}})
            manifest = {"floor_report": floor, "floor_records": records, "contexts": contexts, "visual_brief": brief,
                        "privacy_policy_evidence": privacy, "third_party_upload_allowed": False}
            manifest_ref = save("manifest.json", manifest)
            path = Path(manifest_ref["path"])
            original_hashes = {str(p): sha256(p) for p in root.iterdir()}
            loaded, verified = read_bundle(path)
            result = plan(loaded, enabled=True)
            self.assertEqual(result["quality_floor_status"], "VISUAL_QUALITY_PASS")
            self.assertTrue(verified)
            self.assertEqual(original_hashes, {str(p): sha256(p) for p in root.iterdir()})
            manifest["third_party_upload_allowed"] = True
            save("manifest.json", manifest)
            with self.assertRaisesRegex(ValueError, "PROJECT_PRIVACY_POLICY_MISMATCH"):
                read_bundle(path)
            manifest["third_party_upload_allowed"] = False
            manifest["floor_report"]["sha256"] = "0"*64
            save("manifest.json", manifest)
            with self.assertRaisesRegex(ValueError, "FROZEN_INPUT_CHANGED"):
                read_bundle(path)

    def test_no_production_import_or_generation(self):
        for name in ("runner.py", "visual_quality_floor.py", "project_interface.py", "fast_enhance.py", "template_fill.py"):
            self.assertNotIn("visual_ambition", (ROOT/"scripts/academic_ppt"/name).read_text(encoding="utf8"))
        report = self.run_plan()
        self.assertEqual(report["generation_requests_created"], 0)
        self.assertEqual(report["images_generated"], 0)
        self.assertEqual(report["microcopy_generated_or_applied"], 0)


if __name__ == "__main__":
    unittest.main()
