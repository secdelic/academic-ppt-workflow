from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from scripts.academic_ppt.fast_production import (
    FastProductionError,
    build_changed_slide_plans,
    candidate_attempt_record,
    content_review_status,
    preflight_changed_slide_plans,
)


def _cache() -> dict:
    return {
        "slide_specs": [
            {
                "slide_id": "SLD-BASE-001",
                "slide_revision": 1,
                "slide_role": "content",
                "layout_family": "figure_with_callout",
            },
            {
                "slide_id": "SLD-BASE-002",
                "slide_revision": 2,
                "slide_role": "discussion",
                "layout_family": "split_screen",
            },
        ]
    }


def _registry() -> list[dict]:
    return [
        {"source_id": "SRC-CURRENT", "relative_path": "supplement.csv"},
        {"source_id": "SRC-ADD", "relative_path": "addendum.txt"},
    ]


def _brief(*, clinical: bool = False, approved: bool = True) -> dict:
    return {
        "project_name": "SYNTHETIC_PRODUCTION_CLOSURE",
        "presentation_type": "clinical_simulation" if clinical else "software_verification",
        "fast_enhance": {
            "content_review_status": "CONTENT_APPROVED" if approved else "",
            "changed_slides": [
                {
                    "operation": "REPLACE",
                    "target_slide_id": "SLD-BASE-002",
                    "title": "Bounded replacement",
                    "key_message": "Only changed evidence is summarized.",
                    "facts": ["The synthetic source reports one bounded change."],
                    "uncertainties": ["Human review remains required."],
                    "source_bindings": ["supplement.csv"],
                    "prohibited_wording": ["Do not imply a real clinical result."],
                },
                {
                    "operation": "INSERT_AFTER",
                    "logical_key": "synthetic/new",
                    "insertion_anchor": "SLD-BASE-002",
                    "title": "Bounded insertion",
                    "key_message": "The insertion is explicitly anchored.",
                    "facts": ["The synthetic addendum is local and source-bound."],
                    "source_bindings": ["SRC-ADD"],
                },
            ],
        },
    }


class ProductionPlannerTests(unittest.TestCase):
    def test_batch_planner_derives_revision_id_hash_and_bindings(self) -> None:
        plans = build_changed_slide_plans(
            brief=_brief(), cached_state=_cache(), source_registry=_registry()
        )
        self.assertEqual(len(plans), 2)
        self.assertEqual(plans[0]["slide_revision"], 3)
        self.assertTrue(plans[1]["slide_id"].startswith("SLD-FAST-"))
        self.assertRegex(plans[0]["content_hash"], r"^[0-9a-f]{64}$")
        self.assertEqual(plans[0]["source_bindings"][0]["source_id"], "SRC-CURRENT")

    def test_retry_slide_replans_only_one_declared_slide(self) -> None:
        full = build_changed_slide_plans(
            brief=_brief(), cached_state=_cache(), source_registry=_registry()
        )
        one = build_changed_slide_plans(
            brief=_brief(), cached_state=_cache(), source_registry=_registry(),
            retry_slide=full[0]["slide_id"],
        )
        self.assertEqual([row["slide_id"] for row in one], [full[0]["slide_id"]])

    def test_clinical_content_gate_cannot_be_auto_approved(self) -> None:
        self.assertEqual(
            content_review_status(
                brief=_brief(clinical=True, approved=False),
                auto_approve_content=True,
                clinical_privacy_mode=True,
            ),
            "CONTENT_REVIEW_REQUIRED",
        )

    def test_explicit_review_approval_is_preserved(self) -> None:
        self.assertEqual(
            content_review_status(
                brief=_brief(clinical=True, approved=True),
                auto_approve_content=False,
                clinical_privacy_mode=True,
            ),
            "CONTENT_APPROVED",
        )

    def test_layout_preflight_and_attempt_limit(self) -> None:
        plans = build_changed_slide_plans(
            brief=_brief(), cached_state=_cache(), source_registry=_registry()
        )
        self.assertEqual(preflight_changed_slide_plans(plans)["status"], "PASS")
        self.assertEqual(candidate_attempt_record(plans, attempt=2)["maximum_attempts"], 2)
        with self.assertRaises(FastProductionError):
            candidate_attempt_record(plans, attempt=3)

    def test_unsupported_source_or_missing_fact_fails_closed(self) -> None:
        brief = _brief()
        brief["fast_enhance"]["changed_slides"][0]["source_bindings"] = ["missing.csv"]
        with self.assertRaises(FastProductionError):
            build_changed_slide_plans(
                brief=brief, cached_state=_cache(), source_registry=_registry()
            )


if __name__ == "__main__":
    unittest.main()
