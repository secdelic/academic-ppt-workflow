"""B2 safety contracts; all fixtures synthetic, no patient material."""
import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from scripts.academic_ppt import illustration_requests as ir
from tests.support import illustration_fixture as adapter


class IllustrationRequestTests(unittest.TestCase):
    def setUp(self):
        self.record = {"slide_id": "S001", "page_role": "HERO", "visual_anchor_role": "ANCHOR",
                       "recommended_ambition_level": "STRETCH", "enrichment_opportunity": "HIGH",
                       "quality_floor_status": "HUMAN_REVIEW", "deck_quality_floor_status": "VISUAL_QUALITY_BLOCKED",
                       "source_priority": copy.deepcopy(ir.SOURCE_PRIORITY), "source_figure_required": False,
                       "recommended_next_action": "REFERENCE_GRADE_SCENE_TRIAL", "illustration_decision": "NO_ILLUSTRATION_NEEDED"}
        self.context = adapter.context_for(1)
        self.brief = {"approved": True, "style_family": ir.STYLE, "palette": copy.deepcopy(ir.PALETTE)}
        self.sources = {"scientific_source_figure": False, "approved_project_asset": False, "approved_deterministic_library": False}

    def build(self):
        return ir.build_request(self.record, self.context, self.brief, self.sources, enabled=True)

    def reject_change(self, field, value):
        request = self.build(); request[field] = value
        with self.assertRaises(ValueError): ir.validate_request(request)

    def test_default_off_does_not_read_or_write(self):
        with patch.object(Path, "read_text", side_effect=AssertionError), patch.object(Path, "write_text", side_effect=AssertionError):
            self.assertIsNone(ir.build_request(None, None, None, None))


    def test_read_only_planning(self):
        before = copy.deepcopy((self.record, self.context, self.brief, self.sources))
        self.build(); self.assertEqual(before, (self.record, self.context, self.brief, self.sources))

    def test_required_request_fields_and_closed_schema(self):
        r = self.build(); self.assertEqual(set(r), set(ir.request_schema()["required"]))
        self.assertFalse(ir.request_schema()["additionalProperties"])

    def test_no_source_text_or_paths_are_projected(self):
        self.record["private_notes"] = "SYNTHETIC_SENTINEL_NOT_FOR_EXPORT"
        self.record["audit_provenance"] = [{"path": ("C:" + chr(92) + 'private/synthetic.txt')}]
        payload = json.dumps(self.build()); self.assertNotIn("SYNTHETIC_SENTINEL", payload); self.assertNotIn("private", payload)

    def test_free_text_context_rejected(self):
        self.context["subject_summary"] = "arbitrary source paragraph"
        with self.assertRaises(ValueError): self.build()

    def test_patient_identifier_in_subject_rejected(self):
        self.context["subject_class"] = "Patient SYNTHETIC-ID-001"
        with self.assertRaises(ValueError): self.build()

    def test_local_path_subject_rejected(self):
        self.context["subject_class"] = ("C:" + chr(92) + 'private/note.txt')
        with self.assertRaises(ValueError): self.build()

    def test_patient_identifier_slide_id_rejected(self):
        self.record["slide_id"] = "Patient-SYNTHETIC-001"
        with self.assertRaises(ValueError): self.build()

    def test_path_negative_space_rejected(self):
        self.context["negative_space_region"] = ("C:" + chr(92) + 'notes')
        with self.assertRaises(ValueError): self.build()

    def test_free_text_request_rejected(self):
        self.reject_change("subject_summary", "Patient with private values")

    def test_unknown_request_field_rejected(self):
        r = self.build(); r["speaker_notes"] = "private"
        with self.assertRaises(ValueError): ir.validate_request(r)

    def test_reasoning_leakage_rejected(self):
        self.reject_change("reasoning_summary", ("C:" + chr(92) + 'private/notes.txt'))

    def test_source_priority_cannot_be_reordered(self):
        self.record["source_priority"].reverse()
        with self.assertRaises(ValueError): self.build()

    def test_existing_source_figure_wins(self):
        self.sources = {k: True for k in self.sources}
        r = self.build(); self.assertEqual(r["status"], "REUSE_SOURCE_FIGURE"); self.assertFalse(r["illustration_needed"])

    def test_approved_asset_precedes_library_and_ai(self):
        self.sources["approved_project_asset"] = True; self.sources["approved_deterministic_library"] = True
        self.assertEqual(self.build()["source_selection"], "APPROVED_PROJECT_ASSET")

    def test_approved_library_precedes_ai(self):
        self.sources["approved_deterministic_library"] = True
        self.assertEqual(self.build()["source_selection"], "APPROVED_DETERMINISTIC_LIBRARY")

    def test_required_source_figure_has_no_decorative_substitute(self):
        self.context["source_figure_required"] = True
        r = self.build(); self.assertEqual(r["status"], "SOURCE_FIGURE_REQUIRED"); self.assertEqual(r["allowed_visual_types"], [])

    def test_planner_source_figure_requirement_cannot_be_removed(self):
        self.record["source_figure_required"] = True
        self.assertEqual(self.build()["status"], "SOURCE_FIGURE_REQUIRED")

    def test_prohibited_medical_visuals_cannot_be_removed(self):
        self.reject_change("prohibited_visual_types", [])

    def test_forbidden_visual_cannot_be_added(self):
        self.reject_change("allowed_visual_types", ["EEG"])

    def test_numeric_negative_control(self):
        self.context = adapter.context_for(11)
        r = self.build(); self.assertEqual(r["status"], "NO_ILLUSTRATION_NEEDED"); self.assertFalse(r["illustration_needed"])

    def test_closing_no_illustration_and_no_microcopy(self):
        self.context = adapter.context_for(15)
        r = self.build(); self.assertEqual(r["status"], "NO_ILLUSTRATION_NEEDED"); self.assertFalse(r["microcopy_applied"])

    def test_unresolved_no_visual_execution(self):
        for index in (5, 9):
            self.context = adapter.context_for(index); self.record["recommended_next_action"] = "HUMAN_SEMANTIC_DECISION"
            self.assertEqual(self.build()["status"], "NO_VISUAL_EXECUTION")

    def test_unresolved_mapping_cannot_be_reclassified(self):
        self.record["recommended_next_action"] = "HUMAN_SEMANTIC_DECISION"
        with self.assertRaises(ValueError): self.build()

    def test_semantic_boundary_must_match_subject(self):
        self.context["semantic_boundary"] = "NATIVE_TIMELINE_ONLY_FUTURE_EXPLORATION"
        with self.assertRaises(ValueError): self.build()

    def test_candidate_only_scientific_false_and_unapproved(self):
        r = self.build(); self.assertTrue(r["candidate_only"]); self.assertTrue(r["presentation_only"])
        self.assertIs(r["scientific_evidence"], False); self.assertIs(r["human_approved"], False)
        self.reject_change("scientific_evidence", True)

    def test_determinism(self):
        self.assertEqual(self.build(), self.build())

    def test_hash_stable_under_key_order(self):
        a = self.build(); b = dict(reversed(list(a.items())))
        self.assertEqual(ir.canonical_hash(a), ir.canonical_hash(b)); ir.validate_request(b)

    def test_hash_tampering_rejected(self):
        self.reject_change("request_id", "IRQ-" + "0"*24)

    def test_visual_brief_style_preserved(self):
        self.assertEqual(self.build()["palette"], self.brief["palette"])
        self.brief["style_family"] = "NEON"
        with self.assertRaises(ValueError): self.build()

    def test_palette_leakage_rejected(self):
        self.brief["palette"][0] = "patient name"
        with self.assertRaises(ValueError): self.build()

    def test_negative_space_bounds(self):
        self.context["negative_space_region"]["w"] = 2
        with self.assertRaises(ValueError): self.build()

    def test_negative_space_nan(self):
        self.context["negative_space_region"]["x"] = float("nan")
        with self.assertRaises(ValueError): self.build()

    def test_negative_space_zero_area(self):
        self.context["negative_space_region"]["w"] = 0
        with self.assertRaises(ValueError): self.build()

    def test_focal_cannot_intrude_negative_space(self):
        self.context["focal_region"] = copy.deepcopy(self.context["negative_space_region"])
        with self.assertRaises(ValueError): self.build()

    def test_floor_block_remains_and_cannot_be_overridden(self):
        self.assertEqual(self.build()["deck_quality_floor_status"], "VISUAL_QUALITY_BLOCKED")
        self.reject_change("floor_override_allowed", True)

    def test_planner_roles_and_ambition_retained(self):
        r = self.build()
        self.assertEqual((r["page_role"], r["anchor_role"], r["ambition_level"], r["enrichment_opportunity"]),
                         (self.record["page_role"], self.record["visual_anchor_role"], self.record["recommended_ambition_level"], self.record["enrichment_opportunity"]))

    def test_generation_and_production_not_authorized(self):
        self.reject_change("generation_authorized", True); self.reject_change("production_insertion_allowed", True)

    def asset_fixture(self):
        request = self.build(); sha = "a"*64
        asset = {"candidate_id": "B2-P001", "slide_id": request["slide_id"], "request_id": request["request_id"],
                 "request_hash": ir.canonical_hash(request), "asset_sha256": sha, "kind": "PLACEHOLDER", "status": "CANDIDATE",
                 "scientific_evidence": False, "presentation_only": True, "human_approved": False,
                 "generator_identity": "native_pptxgenjs_request_box", "generation_timestamp": None}
        auth = {"slide_id": request["slide_id"], "request_id": request["request_id"], "asset_sha256": sha,
                "candidate_preview_approved": True, "production_approved": False}
        return asset, request, auth, sha

    def test_asset_slide_authorization(self):
        a,r,u,h = self.asset_fixture(); ir.validate_candidate(a,r,u,file_sha256=h)
        u["slide_id"] = "S099"
        with self.assertRaises(ValueError): ir.validate_candidate(a,r,u,file_sha256=h)

    def test_asset_request_binding(self):
        a,r,u,h = self.asset_fixture(); a["request_hash"] = "0"*64
        with self.assertRaises(ValueError): ir.validate_candidate(a,r,u,file_sha256=h)

    def test_asset_hash_binding(self):
        a,r,u,h = self.asset_fixture()
        with self.assertRaises(ValueError): ir.validate_candidate(a,r,u,file_sha256="b"*64)

    def test_asset_cannot_claim_scientific_approval(self):
        a,r,u,h = self.asset_fixture(); a["human_approved"] = True
        with self.assertRaises(ValueError): ir.validate_candidate(a,r,u,file_sha256=h)

    def test_placeholder_not_mislabeled_as_generated(self):
        a,r,u,h = self.asset_fixture(); a["generator_identity"] = "invented_generator"
        with self.assertRaises(ValueError): ir.validate_candidate(a,r,u,file_sha256=h)

    def test_source_status_is_not_truthy_string(self):
        self.sources["approved_project_asset"] = "yes"
        with self.assertRaises(ValueError): self.build()



    def test_generated_candidate_requires_generation_authorization(self):
        a,r,u,h = self.asset_fixture(); a["kind"] = "GENERATED_CANDIDATE"
        a["generator_identity"] = "unapproved_test_generator"; a["generation_timestamp"] = "2026-01-01T00:00:00Z"
        with self.assertRaises(ValueError): ir.validate_candidate(a,r,u,file_sha256=h)


if __name__ == "__main__":
    unittest.main()
