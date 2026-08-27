from __future__ import annotations

import copy
import json
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.deck_ir import (  # noqa: E402
    build_deck_ir,
    canonical_json_hash,
    stable_slide_id,
    to_backend_spec,
    to_storyboard_rows,
    validate_deck_ir,
)
from academic_ppt.narrative import (  # noqa: E402
    EvidenceBoundaryError,
    NARRATIVE_MODES,
    REQUIRED_POLICY_FIELDS,
    guard_evidence_strength,
    load_narrative_modes,
    neutral_title,
    select_narrative_mode,
)
from academic_ppt.routing import (  # noqa: E402
    ROUTE_VALUES,
    RouteRequest,
    RouteValidationError,
    TopLevelRoute,
    validate_route_request,
)


BRIEF = {
    "project_name": "Synthetic source-bound study",
    "presentation_type": "research_report",
    "language": "en-US",
    "presentation_objective": "Summarise registered synthetic evidence",
}

PLANNED_SLIDES = [
    {
        "logical_slide_key": "opening/source-boundary",
        "section_id": "opening",
        "slide_role": "evidence_boundary",
        "slide_title": "The source package defines the evidence boundary",
        "slide_purpose": "Declare scope",
        "single_key_message": "Only registered evidence may enter the deck.",
        "source_ids": ["SRC-001"],
        "proposed_layout": "text_plus_source_list",
        "visual_type": "source_map",
        "citation_requirement": "source manifest",
        "speaker_note_summary": "State the evidence boundary.",
        "confidence": "high",
        "manual_review_required": "no",
    },
    {
        "logical_slide_key": "results/primary-outcome",
        "section_id": "results",
        "slide_role": "primary_result",
        "slide_title": "The registered result is reported without causal wording",
        "slide_purpose": "Present one result",
        "single_key_message": "The source reports a synthetic association.",
        "claim_ids": ["CLM-001"],
        "source_ids": ["SRC-001"],
        "figure_ids": ["FIG-001"],
        "proposed_layout": "result_with_chart",
        "visual_type": "editable_bar_chart",
        "citation_requirement": "required",
        "speaker_note_summary": "Do not infer causality.",
        "confidence": "high",
        "manual_review_required": "yes",
        "editable_object_requirements": ["chart"],
    },
]


class RouteContractTests(unittest.TestCase):
    def test_route_enum_is_exact_and_complete(self):
        expected = {
            "generate",
            "create-style-profile",
            "fill-template",
            "enhance-existing",
        }
        self.assertEqual(ROUTE_VALUES, expected)
        self.assertEqual({item.value for item in TopLevelRoute}, expected)

    def test_generate_request_is_normalised_without_network(self):
        result = validate_route_request(
            RouteRequest(
                route="generate",
                brief=BRIEF,
                input_paths=("input/documents/paper.pdf",),
                options={"backend": "pptxgenjs", "use_network": False},
            )
        )
        self.assertEqual(result.route, "generate")
        self.assertEqual(result.options["backend"], "pptxgenjs")

    def test_invalid_route_combinations_are_rejected(self):
        with self.assertRaisesRegex(RouteValidationError, "Exactly one"):
            validate_route_request(
                {
                    "generate": True,
                    "fill_template": True,
                    "brief": BRIEF,
                    "template_path": "template.pptx",
                }
            )
        with self.assertRaisesRegex(RouteValidationError, "cannot also set"):
            validate_route_request(
                RouteRequest(
                    route="generate",
                    brief=BRIEF,
                    template_path="template.pptx",
                )
            )

    def test_route_prerequisites_and_options_are_enforced(self):
        with self.assertRaisesRegex(RouteValidationError, "reference_path"):
            validate_route_request(
                RouteRequest(route="create-style-profile", brief=BRIEF)
            )
        with self.assertRaisesRegex(RouteValidationError, "At most one"):
            validate_route_request(
                RouteRequest(
                    route="enhance-existing",
                    brief=BRIEF,
                    existing_deck_path="existing.pptx",
                    options={
                        "update_slide": "SLD-A",
                        "update_section": "results",
                    },
                )
            )
        with self.assertRaisesRegex(RouteValidationError, "prohibited"):
            validate_route_request(
                RouteRequest(
                    route="generate",
                    brief=BRIEF,
                    options={"use_network": True},
                )
            )


class NarrativeContractTests(unittest.TestCase):
    def test_config_has_all_seven_modes_and_policy_fields(self):
        modes = load_narrative_modes(ROOT / "config" / "narrative_modes.yaml")
        self.assertEqual(set(modes), set(NARRATIVE_MODES))
        self.assertEqual(len(modes), 7)
        for policy in modes.values():
            self.assertTrue(REQUIRED_POLICY_FIELDS.issubset(policy))

    def test_narrative_selection_is_explicit_or_profile_based(self):
        self.assertEqual(
            select_narrative_mode(BRIEF),
            "scientific_problem",
        )
        self.assertEqual(
            select_narrative_mode(BRIEF, explicit="neutral_briefing"),
            "neutral_briefing",
        )
        with self.assertRaises(ValueError):
            select_narrative_mode(BRIEF, explicit="marketing_pitch")

    def test_neutral_title_and_evidence_guard_do_not_upgrade_evidence(self):
        self.assertEqual(
            neutral_title("Gene-set enrichment", "computational_inference"),
            "Gene-set enrichment — computational inference",
        )
        safe = "Exposure was associated with the reported outcome."
        self.assertEqual(
            guard_evidence_strength(safe, "observational_association"), safe
        )
        with self.assertRaises(EvidenceBoundaryError):
            guard_evidence_strength(
                "Exposure caused the reported outcome.",
                "observational_association",
            )
        with self.assertRaises(EvidenceBoundaryError):
            guard_evidence_strength(
                "The analysis experimentally validated clinical effectiveness.",
                "computational_inference",
            )


class DeckIRTests(unittest.TestCase):
    def test_slide_ids_are_stable_when_page_order_changes(self):
        original = build_deck_ir(BRIEF, PLANNED_SLIDES)
        reordered = build_deck_ir(
            BRIEF,
            list(reversed(PLANNED_SLIDES)),
            existing_ir=original,
        )
        original_ids = {
            slide["logical_slide_key"]: slide["slide_id"]
            for slide in original["slides"]
        }
        reordered_ids = {
            slide["logical_slide_key"]: slide["slide_id"]
            for slide in reordered["slides"]
        }
        self.assertEqual(original_ids, reordered_ids)
        self.assertEqual(
            [slide["slide_revision"] for slide in reordered["slides"]],
            [1, 1],
        )
        self.assertEqual(reordered["deck_version"], original["deck_version"] + 1)
        for logical_key, slide_id in original_ids.items():
            self.assertEqual(
                slide_id, stable_slide_id(original["deck_id"], logical_key)
            )

    def test_revision_bumps_only_when_content_hash_changes(self):
        first = build_deck_ir(BRIEF, PLANNED_SLIDES)
        unchanged = build_deck_ir(BRIEF, copy.deepcopy(PLANNED_SLIDES), first)
        self.assertEqual(unchanged["deck_version"], first["deck_version"])
        self.assertEqual(
            [slide["slide_revision"] for slide in unchanged["slides"]],
            [1, 1],
        )

        changed_plans = copy.deepcopy(PLANNED_SLIDES)
        changed_plans[1][
            "single_key_message"
        ] = "The source reports a revised synthetic association."
        changed = build_deck_ir(BRIEF, changed_plans, unchanged)
        by_key = {
            slide["logical_slide_key"]: slide for slide in changed["slides"]
        }
        self.assertEqual(
            by_key["opening/source-boundary"]["slide_revision"], 1
        )
        self.assertEqual(
            by_key["results/primary-outcome"]["slide_revision"], 2
        )
        self.assertNotEqual(
            by_key["results/primary-outcome"]["content_hash"],
            unchanged["slides"][1]["content_hash"],
        )

    def test_deck_ir_hashes_and_claim_source_boundary_are_validated(self):
        ir = build_deck_ir(BRIEF, PLANNED_SLIDES)
        validate_deck_ir(ir)
        canonical_payload = {
            key: value for key, value in ir.items() if key != "canonical_hash"
        }
        self.assertEqual(ir["canonical_hash"], canonical_json_hash(canonical_payload))

        tampered = json.loads(json.dumps(ir))
        tampered["slides"][1]["source_ids"] = []
        with self.assertRaisesRegex(ValueError, "claim_ids without source_ids"):
            validate_deck_ir(tampered)

    def test_storyboard_and_backend_adapters_preserve_stable_identity(self):
        ir = build_deck_ir(BRIEF, PLANNED_SLIDES)
        rows = to_storyboard_rows(ir)
        self.assertEqual(rows[0]["slide_id"], ir["slides"][0]["slide_id"])
        self.assertEqual(rows[1]["source_ids"], "SRC-001")
        self.assertEqual(rows[1]["manual_review_required"], "yes")

        spec = to_backend_spec(ir, BRIEF, "restrained_biomedical")
        self.assertEqual(spec["deck_ir"]["canonical_hash"], ir["canonical_hash"])
        self.assertEqual(spec["slides"][1]["claim_ids"], ["CLM-001"])
        self.assertEqual(spec["slides"][1]["source_ids"], ["SRC-001"])


if __name__ == "__main__":
    unittest.main()
