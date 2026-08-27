from __future__ import annotations

import tempfile
import unittest
from copy import deepcopy
from pathlib import Path

from scripts.academic_ppt.fast_enhance import build_cached_change_impact_graph
from scripts.academic_ppt.extractors import extract_selected
from scripts.academic_ppt.fast_production import (
    FastProductionError,
    build_changed_slide_plans,
)
from scripts.academic_ppt.inventory import (
    APPROVED_VISUAL_ASSET,
    FORMAL_TEMPLATE,
    PRESENTATION_BASELINE,
    SCIENTIFIC_SOURCE,
    STYLE_REFERENCE,
    classify_source_role,
    inventory_sources,
    scientific_delta_manifest,
)


FIXTURE_PROVENANCE = "SYNTHETIC"
FIXTURE_ID = "SECOND_DEVICE_EXISTING_DECK_LINEAGE"


def _cache() -> dict:
    return {
        "slide_specs": [
            {
                "slide_id": "SLD-SYNTHETIC-001",
                "slide_revision": 1,
                "slide_role": "content",
                "layout_family": "figure_with_callout",
            }
        ]
    }


def _brief(binding: object, *, claim_ids: list[str] | None = None) -> dict:
    return {
        "fixture_provenance": FIXTURE_PROVENANCE,
        "fixture_id": FIXTURE_ID,
        "fast_enhance": {
            "content_review_status": "CONTENT_APPROVED",
            "changed_slides": [
                {
                    "operation": "REPLACE",
                    "target_slide_id": "SLD-SYNTHETIC-001",
                    "title": "Synthetic bounded update",
                    "key_message": "Only declared synthetic content is shown.",
                    "facts": ["Synthetic lineage statement."],
                    "source_bindings": [binding],
                    "claim_ids": claim_ids or [],
                }
            ],
        },
    }


def _registry() -> list[dict]:
    return [
        {
            "source_id": "SRC-SCIENCE",
            "relative_path": "documents/supplemental_source.docx",
            "source_role": SCIENTIFIC_SOURCE,
        },
        {
            "source_id": "SRC-BASELINE",
            "relative_path": "existing/existing_deck.pptx",
            "source_role": PRESENTATION_BASELINE,
        },
        {
            "source_id": "SRC-STYLE",
            "relative_path": "style_reference/reference.pptx",
            "source_role": STYLE_REFERENCE,
        },
        {
            "source_id": "SRC-TEMPLATE",
            "relative_path": "template/formal.potx",
            "source_role": FORMAL_TEMPLATE,
        },
        {
            "source_id": "SRC-ASSET",
            "relative_path": "approved_assets/concept.svg",
            "source_role": APPROVED_VISUAL_ASSET,
        },
    ]


class SourceRoleContractTests(unittest.TestCase):
    def test_canonical_role_classifier(self) -> None:
        self.assertEqual(
            classify_source_role(
                "existing/existing_deck.pptx", route="enhance-existing"
            ),
            PRESENTATION_BASELINE,
        )
        self.assertEqual(
            classify_source_role("style_reference/reference.pptx"), STYLE_REFERENCE
        )
        self.assertEqual(
            classify_source_role("template/formal.potx"), FORMAL_TEMPLATE
        )
        self.assertEqual(
            classify_source_role("approved_assets/concept.svg"),
            APPROVED_VISUAL_ASSET,
        )
        self.assertEqual(
            classify_source_role("documents/source.docx"), SCIENTIFIC_SOURCE
        )

    def test_existing_deck_requires_existing_explicit_content_reference_mode(self) -> None:
        self.assertEqual(
            classify_source_role(
                "existing/existing_deck.pptx",
                route="enhance-existing",
                reference_mode="content-reference",
            ),
            SCIENTIFIC_SOURCE,
        )

    def test_external_workspace_inventory_keeps_deck_out_of_scientific_role(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            workspace = Path(temporary) / "workspace"
            existing = workspace / "projects" / "P1" / "input" / "existing"
            documents = workspace / "projects" / "P1" / "input" / "documents"
            existing.mkdir(parents=True)
            documents.mkdir(parents=True)
            (existing / "existing_deck.pptx").write_bytes(b"synthetic package marker")
            (documents / "supplemental_source.docx").write_bytes(b"synthetic document marker")
            rows = inventory_sources(
                workspace / "projects" / "P1" / "input",
                {".pptx", ".docx"},
                workspace / "manifest.csv",
                route="enhance-existing",
            )
            roles = {row["relative_path"]: row["source_role"] for row in rows}
            self.assertEqual(
                roles["existing/existing_deck.pptx"], PRESENTATION_BASELINE
            )
            self.assertEqual(
                roles["documents/supplemental_source.docx"], SCIENTIFIC_SOURCE
            )

    def test_presentation_baseline_content_is_not_extracted_as_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            input_root = root / "input"
            existing = input_root / "existing" / "existing_deck.pptx"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"not a real package because extraction must be skipped")
            rows = inventory_sources(
                input_root,
                {".pptx"},
                root / "manifest.csv",
                route="enhance-existing",
            )
            extracted, stats = extract_selected(
                input_root,
                root / "staging",
                rows,
                {rows[0]["source_id"]},
            )
            self.assertEqual(extracted[rows[0]["source_id"]], "")
            self.assertEqual(stats["parsed"], 1)
            self.assertIn("intentionally excluded", rows[0]["parse_warning"])

    def test_scientific_delta_excludes_presentation_baseline(self) -> None:
        delta = {
            "hash_algorithm": "sha256",
            "entries": [
                {"source_key": "existing/existing_deck.pptx", "status": "MODIFIED"},
                {"source_key": "documents/supplemental_source.docx", "status": "NEW"},
            ],
        }
        projected = scientific_delta_manifest(delta, route="enhance-existing")
        self.assertEqual(
            projected["parse_source_keys"],
            ["documents/supplemental_source.docx"],
        )
        self.assertEqual(projected["counts"]["MODIFIED"], 0)
        self.assertEqual(projected["counts"]["NEW"], 1)

    def test_change_impact_denominator_contains_only_scientific_sources(self) -> None:
        graph = build_cached_change_impact_graph(
            source_registry=_registry(),
            evidence_registry=[
                {
                    "claim_id": "CLM-SYNTHETIC",
                    "source_id": "SRC-SCIENCE",
                }
            ],
            slide_specs=[
                {
                    "slide_id": "SLD-SYNTHETIC-001",
                    "claim_ids": ["CLM-SYNTHETIC"],
                }
            ],
        )
        sources = graph.to_dict()["nodes"]["sources"]
        self.assertEqual(sources, ["documents/supplemental_source.docx"])


class ChangedSlideLineageTests(unittest.TestCase):
    def test_scientific_source_binding_passes(self) -> None:
        plans = build_changed_slide_plans(
            brief=_brief("SRC-SCIENCE"),
            cached_state=_cache(),
            source_registry=_registry(),
        )
        self.assertEqual(plans[0]["presentation_lineage"], "CHANGED_SCIENTIFIC_CONTENT")
        self.assertEqual(plans[0]["source_bindings"][0]["claim_ceiling"], "SCIENTIFIC_SOURCE_BOUND")

    def test_plain_existing_deck_binding_cannot_satisfy_scientific_content(self) -> None:
        with self.assertRaisesRegex(FastProductionError, "Presentation baseline"):
            build_changed_slide_plans(
                brief=_brief("SRC-BASELINE"),
                cached_state=_cache(),
                source_registry=_registry(),
            )

    def test_explicit_inherited_presentation_content_is_bounded_and_reviewable(self) -> None:
        plans = build_changed_slide_plans(
            brief=_brief(
                {
                    "source_id": "SRC-BASELINE",
                    "canonical_status": "INHERITED_PRESENTATION_CONTENT",
                }
            ),
            cached_state=_cache(),
            source_registry=_registry(),
        )
        plan = plans[0]
        self.assertEqual(plan["presentation_lineage"], "INHERITED_PRESENTATION_CONTENT")
        self.assertTrue(plan["source_reverification_required"])
        self.assertIn("SOURCE_REVERIFICATION_REQUIRED", plan["review_flags"])
        self.assertEqual(
            plan["source_bindings"][0]["claim_ceiling"],
            "PRESENTATION_ONLY_NOT_VERIFIED_SCIENTIFIC_EVIDENCE",
        )

    def test_claim_id_still_requires_scientific_source(self) -> None:
        brief = _brief(
            {
                "source_id": "SRC-BASELINE",
                "canonical_status": "INHERITED_PRESENTATION_CONTENT",
            },
            claim_ids=["CLM-SYNTHETIC"],
        )
        with self.assertRaisesRegex(FastProductionError, "scientific source binding"):
            build_changed_slide_plans(
                brief=brief,
                cached_state=_cache(),
                source_registry=_registry(),
                changed_claims=[
                    {
                        "claim_id": "CLM-SYNTHETIC",
                        "claim_text": "Synthetic claim.",
                        "allowed_wording": "Synthetic claim.",
                    }
                ],
            )

    def test_style_template_and_asset_never_satisfy_scientific_lineage(self) -> None:
        for source_id in ("SRC-STYLE", "SRC-TEMPLATE", "SRC-ASSET"):
            with self.subTest(source_id=source_id):
                brief = deepcopy(_brief(source_id))
                with self.assertRaisesRegex(FastProductionError, "cannot satisfy"):
                    build_changed_slide_plans(
                        brief=brief,
                        cached_state=_cache(),
                        source_registry=_registry(),
                    )


if __name__ == "__main__":
    unittest.main()
