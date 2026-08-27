from __future__ import annotations

import json
import tempfile
import unittest
from copy import deepcopy
from pathlib import Path

from scripts.academic_ppt.visual_brief import (
    VisualBriefError,
    load_visual_brief,
    resolve_visual_direction,
    validate_visual_brief,
    visual_direction_to_style_profile,
)


REPO = Path(__file__).resolve().parents[1]


def approved_brief() -> dict:
    value = json.loads((REPO / "config" / "visual_brief.template.yaml").read_text(encoding="utf-8"))
    value["visual_brief_id"] = "approved-general-direction"
    value["approval"] = {
        "status": "approved",
        "approved_by": "authorized-reviewer",
        "approved_at": "2026-08-12T12:00:00+08:00",
    }
    return value


class VisualBriefContractTests(unittest.TestCase):
    def test_schema_and_template_are_json_compatible_yaml(self) -> None:
        schema = json.loads((REPO / "config" / "visual_brief.schema.yaml").read_text(encoding="utf-8"))
        template = json.loads((REPO / "config" / "visual_brief.template.yaml").read_text(encoding="utf-8"))
        self.assertEqual(schema["$id"], "academic-ppt/visual-brief/1")
        self.assertEqual(validate_visual_brief(template), [])
        self.assertFalse(set(schema["properties"]["direction"]["properties"]) & {"claim", "results", "data"})

    def test_loader_requires_human_approval_when_requested(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "visual_brief.yaml"
            path.write_text(json.dumps(approved_brief(), ensure_ascii=False), encoding="utf-8")
            loaded = load_visual_brief(path, require_approved=True)
            self.assertEqual(loaded["approval"]["status"], "approved")
            loaded["approval"]["status"] = "draft"
            path.write_text(json.dumps(loaded, ensure_ascii=False), encoding="utf-8")
            with self.assertRaisesRegex(VisualBriefError, "not approved"):
                load_visual_brief(path, require_approved=True)

    def test_scientific_fact_or_source_binding_is_rejected(self) -> None:
        for field, value in (
            ("sample_size", 450),
            ("results", [{"estimate": 1.2}]),
            ("source_bindings", ["SRC-001"]),
        ):
            brief = approved_brief()
            brief["direction"][field] = value
            errors = validate_visual_brief(brief)
            self.assertTrue(any("prohibited" in error or "unsupported" in error for error in errors), field)

    def test_fixed_precedence_is_applied_low_to_high(self) -> None:
        brief = approved_brief()
        brief["direction"]["visual_motif"] = "approved-brief"
        explicit = deepcopy(brief["direction"])
        explicit["visual_motif"] = "explicit-user"
        result = resolve_visual_direction(
            route="generate",
            user_explicit_instruction=explicit,
            visual_brief=brief,
            authorized_template={"authorized": True, "style": {"visual_motif": "template"}},
            existing_deck_style={"visual_motif": "existing"},
            canonical_workspace={"visual_motif": "workspace"},
            auto_direction={"visual_motif": "auto"},
        )
        self.assertEqual(result["resolved_direction"]["visual_motif"], "explicit-user")
        self.assertEqual(result["primary_authority"], "user_explicit_instruction")
        self.assertEqual(
            result["precedence_high_to_low"],
            [
                "user_explicit_instruction",
                "approved_visual_brief",
                "authorized_template",
                "existing_deck_style",
                "canonical_workspace",
                "auto",
            ],
        )

    def test_explicit_user_instruction_may_be_a_partial_visual_override(self) -> None:
        result = resolve_visual_direction(
            route="generate",
            user_explicit_instruction={"background_strategy": "light"},
            visual_brief=approved_brief(),
        )
        self.assertEqual(result["primary_authority"], "user_explicit_instruction")
        self.assertEqual(result["resolved_direction"]["background_strategy"], "light")
        self.assertEqual(result["resolved_direction"]["visual_motif"], approved_brief()["direction"]["visual_motif"])

    def test_unapproved_visual_brief_does_not_override_fallback(self) -> None:
        brief = approved_brief()
        brief["approval"] = {"status": "draft", "approved_by": None, "approved_at": None}
        result = resolve_visual_direction(
            route="generate",
            visual_brief=brief,
            canonical_workspace={"visual_motif": "workspace"},
            auto_direction={"visual_motif": "auto"},
        )
        self.assertEqual(result["primary_authority"], "canonical_workspace")
        self.assertFalse(result["human_visual_brief_approved"])

    def test_small_enhance_without_visual_brief_inherits_existing_deck(self) -> None:
        result = resolve_visual_direction(
            route="enhance",
            existing_deck_style={"visual_motif": "preserved-deck", "background_strategy": "inherit"},
            canonical_workspace={"visual_motif": "workspace"},
            auto_direction={"visual_motif": "auto"},
            redesign_requested=False,
            changed_slide_count=7,
        )
        self.assertEqual(result["primary_authority"], "existing_deck_style")
        self.assertEqual(result["resolved_direction"]["visual_motif"], "preserved-deck")
        self.assertTrue(result["existing_deck_inherited"])
        self.assertFalse(result["art_director_required"])

    def test_template_fill_keeps_master_layout_and_brand_authority(self) -> None:
        brief = approved_brief()
        template = {
            "authorized": True,
            "style": {"visual_motif": "institutional", "card_style": "template-card"},
            "protected_brand": {
                "master_id": "MASTER-AUTHORIZED",
                "layout_ids": ["TITLE", "CONTENT"],
                "theme_colors": ["#003B5C", "#FFFFFF"],
                "logo_policy": "preserve-aspect-ratio",
                "footer_policy": "preserve",
                "animation_policy": "preserve",
            },
        }
        result = resolve_visual_direction(
            route="template-fill",
            visual_brief=brief,
            authorized_template=template,
        )
        self.assertEqual(result["primary_authority"], "approved_visual_brief")
        self.assertEqual(result["template_brand_authority"], "authorized_template")
        self.assertEqual(result["resolved_direction"]["master_id"], "MASTER-AUTHORIZED")
        self.assertEqual(result["resolved_direction"]["animation_policy"], "preserve")
        self.assertIn("layout_ids", result["template_protected_fields"])

    def test_visual_brief_cannot_contain_template_brand_override(self) -> None:
        brief = approved_brief()
        brief["direction"]["theme_colors"] = ["#FFFFFF"]
        errors = validate_visual_brief(brief)
        self.assertTrue(any("template-protected" in error for error in errors))

    def test_approved_direction_projects_only_presentation_tokens(self) -> None:
        resolution = resolve_visual_direction(
            route="generate", visual_brief=approved_brief()
        )
        profile = visual_direction_to_style_profile(resolution)
        self.assertIsNotNone(profile)
        self.assertFalse(profile["scientific_content_included"])
        self.assertTrue(profile["theme_colors"])
        self.assertNotIn("claim", json.dumps(profile).casefold())
        tokens = {item["slot"]: item["value"] for item in profile["theme_colors"]}
        self.assertEqual(tokens["dk2"], "14324B")
        self.assertEqual(tokens["lt1"], "DCECEE")

    def test_enhance_and_template_fill_do_not_flatten_native_style(self) -> None:
        for route in ("enhance", "template-fill"):
            resolution = resolve_visual_direction(
                route=route, visual_brief=approved_brief()
            )
            self.assertIsNone(visual_direction_to_style_profile(resolution))


if __name__ == "__main__":
    unittest.main()
