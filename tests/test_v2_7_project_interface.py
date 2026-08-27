from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.project_interface import (  # noqa: E402
    PUBLIC_QUALITIES,
    PUBLIC_ROUTES,
    ProjectInterfaceError,
    ProjectPaths,
    apply_project_brief_defaults,
    configure_project_args,
    prepare_run_output,
    publish_standard_outputs,
    normalise_presentation_brief,
    resolve_project_reference,
    validate_approved_assets,
    validate_public_output_tree,
)
from run_ppt_workflow import build_parser, normalise_public_request  # noqa: E402
from tests.support.path_contract import (  # noqa: E402
    canonical_test_path,
    path_is_within,
    paths_equal,
)


def approved_visual_brief(route: str) -> dict:
    value = json.loads(
        (ROOT / "config/visual_brief.template.yaml").read_text(encoding="utf-8")
    )
    value["visual_brief_id"] = f"approved-{route}-direction"
    value["scope"] = {"route": route, "redesign_requested": True}
    value["approval"] = {
        "status": "approved",
        "approved_by": "authorized-synthetic-reviewer",
        "approved_at": "2026-08-13T10:00:00+08:00",
    }
    return value


def write_mapping(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def scaffold(root: Path, *, route: str = "generate", quality: str = "validated") -> None:
    for relative in (
        "input/documents",
        "input/data",
        "input/figures",
        "input/references",
        "input/style_reference",
        "input/template",
        "input/existing",
        "brief",
        "approved_assets",
        "output",
        "audit",
        "cache",
        "private",
    ):
        (root / relative).mkdir(parents=True, exist_ok=True)
    write_mapping(
        root / "brief/presentation_brief.yaml",
        {"project_name": "Interface fixture", "route": route, "quality": quality},
    )


class PublicRoutingContractTests(unittest.TestCase):
    def test_exactly_four_routes_and_three_qualities_are_public(self) -> None:
        self.assertEqual(
            PUBLIC_ROUTES,
            ("generate", "enhance", "template-fill", "template-create"),
        )
        self.assertEqual(PUBLIC_QUALITIES, ("quick", "validated", "full"))
        parser = build_parser()
        for route in PUBLIC_ROUTES:
            for quality in PUBLIC_QUALITIES:
                with self.subTest(route=route, quality=quality):
                    args = normalise_public_request(
                        parser.parse_args(["--route", route, "--quality", quality])
                    )
                    self.assertEqual(args.public_route, route)
                    self.assertEqual(args.quality, quality)

    def test_public_help_hides_internal_artifacts_and_recovery_controls(self) -> None:
        help_text = build_parser().format_help()
        for term in (
            "--brief",
            "--template",
            "--ppt",
            "--input",
            "operation-plan",
            "changed-candidate-pptx",
            "content-hash",
            "SlideSpec",
            "VisualSpec",
            "retry-slide",
            "checkpoint",
        ):
            self.assertNotIn(term, help_text)
        self.assertIn("--project", help_text)
        self.assertIn("--route", help_text)
        self.assertIn("--quality", help_text)

    def test_project_brief_supplies_route_and_quality_but_cli_wins(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root, route="enhance", quality="quick")
            parser = build_parser()
            inherited = apply_project_brief_defaults(
                parser.parse_args(["--project", str(root)])
            )
            self.assertEqual(inherited.route, "enhance")
            self.assertEqual(inherited.quality, "quick")
            explicit = apply_project_brief_defaults(
                parser.parse_args(
                    [
                        "--project",
                        str(root),
                        "--route",
                        "generate",
                        "--quality",
                        "full",
                    ]
                )
            )
            self.assertEqual(explicit.route, "generate")
            self.assertEqual(explicit.quality, "full")

    def test_project_id_resolves_beneath_workspace_home(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            with patch.dict(os.environ, {"PPT_WORKSPACE_HOME": temporary}):
                resolved = resolve_project_reference("study-01")
            self.assertEqual(
                resolved.root, (Path(temporary) / "projects/study-01").resolve()
            )

    def test_user_brief_is_flattened_without_inventing_science(self) -> None:
        value = normalise_presentation_brief(
            {
                "project_name": "Nested fixture",
                "presentation": {"language": "zh-CN", "target_slide_count": 12},
                "content_contract": {
                    "must_include": ["registered outcome"],
                    "must_not_claim": ["causality"],
                },
            }
        )
        self.assertEqual(value["language"], "zh-CN")
        self.assertEqual(value["target_slide_count"], 12)
        self.assertEqual(value["must_include"], ["registered outcome"])
        self.assertEqual(value["must_exclude"], ["causality"])
        self.assertNotIn("result", value)


class OutputContractTests(unittest.TestCase):
    def test_quick_publishes_only_draft_and_minimal_changed_qa(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root)
            run = prepare_run_output(
                ProjectPaths.from_root(root),
                route="enhance",
                quality="quick",
                run_id="RUN_QUICK",
            )
            engine = root / "private/engine-result"
            engine.mkdir(parents=True)
            (engine / "updated.pptx").write_bytes(b"synthetic pptx")
            (engine / "changed_slide_qa.md").write_text("# QA\n\nPASS\n", encoding="utf-8")
            (engine / "execution_summary.md").write_text("# Summary\n", encoding="utf-8")
            # Internal runtime objects must not be copied into public output.
            (engine / "operation_plan.json").write_text("{}", encoding="utf-8")
            publish_standard_outputs(engine, run)
            result = validate_public_output_tree(run)
            self.assertEqual(result["status"], "PASS")
            self.assertTrue((run.draft / "deck.pptx").is_file())
            self.assertFalse(any(run.final.iterdir()))
            self.assertEqual(result["internal_artifacts_exposed"], [])
            self.assertLessEqual(len(result["present"]), 4)

    def test_validated_and_full_require_delivery_and_minimal_qa(self) -> None:
        for quality in ("validated", "full"):
            with self.subTest(quality=quality), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary) / "project"
                scaffold(root)
                run = prepare_run_output(
                    ProjectPaths.from_root(root),
                    route="generate",
                    quality=quality,
                    run_id=f"RUN_{quality.upper()}",
                )
                engine = root / "private/engine-result"
                (engine / "preview").mkdir(parents=True)
                (engine / "deck.pptx").write_bytes(b"synthetic pptx")
                (engine / "deck.pdf").write_bytes(b"synthetic pdf")
                (engine / "preview/contact_sheet.png").write_bytes(b"synthetic png")
                (engine / "qa_report.md").write_text(
                    "# QA Report\n\n## Scientific QA\n\n- PASS\n\n"
                    "## Visual QA\n\n- PASS\n\n## File QA\n\n- PASS\n",
                    encoding="utf-8",
                )
                publish_standard_outputs(engine, run)
                result = validate_public_output_tree(run)
                self.assertEqual(result["status"], "PASS", result)
                self.assertTrue((run.final / "deck.pptx").is_file())
                self.assertTrue((run.final / "deck.pdf").is_file())
                self.assertTrue((run.preview / "contact_sheet.png").is_file())
                self.assertIn("PASS", (run.audit / "scientific_qa.md").read_text())
                self.assertIn("PASS", (run.audit / "visual_qa.md").read_text())

    def test_validated_enhance_maps_changed_qa_and_preview_to_public_contract(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root, route="enhance", quality="validated")
            run = prepare_run_output(
                ProjectPaths.from_root(root),
                route="enhance",
                quality="validated",
                run_id="RUN_ENHANCE_VALIDATED",
            )
            engine = root / "private/engine-result"
            preview = engine / "changed_preview"
            preview.mkdir(parents=True)
            (engine / "updated.pptx").write_bytes(b"synthetic pptx")
            (engine / "updated.pdf").write_bytes(b"synthetic pdf")
            (engine / "changed_slide_qa.md").write_text(
                "# Changed-slide QA\n\n- Status: `PASS`\n", encoding="utf-8"
            )
            from PIL import Image

            Image.new("RGB", (160, 90), "white").save(preview / "slide_002.png")
            publish_standard_outputs(engine, run)
            result = validate_public_output_tree(run)
            self.assertEqual(result["status"], "PASS", result)
            self.assertTrue((run.preview / "contact_sheet.png").is_file())
            self.assertIn("PASS", (run.audit / "scientific_qa.md").read_text())
            self.assertIn("PASS", (run.audit / "visual_qa.md").read_text())

    def test_configure_project_args_uses_private_engine_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root)
            args = normalise_public_request(
                build_parser().parse_args(
                    ["--project", str(root), "--route", "generate", "--run-id", "RUN1"]
                )
            )
            configured, run = configure_project_args(args)
            self.assertTrue(paths_equal(configured.input_root, root / "input"))
            self.assertTrue(path_is_within(configured.output_root, root / "private"))
            self.assertTrue(path_is_within(configured.staging_root, root / "private"))
            self.assertTrue(paths_equal(run.run_root, root / "output/RUN1"))  # self-containment: generated-temp
            public_contract = run.public_contract()
            self.assertFalse(public_contract["internal_workflow_objects_exposed"])
            self.assertNotIn("hash", json.dumps(public_contract).casefold())

    def test_configure_project_args_uses_external_project_cache_for_all_routes(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            base = Path(temporary)
            for route in ("generate", "enhance", "template-fill"):
                with self.subTest(route=route):
                    root = base / route
                    scaffold(root, route=route, quality="validated")
                    if route == "enhance":
                        existing = root / "input/existing/existing.pptx"
                        existing.write_bytes(b"synthetic")
                        write_mapping(
                            root / "brief/presentation_brief.yaml",
                            {
                                "project_name": "External cache fixture",
                                "route": route,
                                "quality": "validated",
                                "inputs": {"existing_deck_path": str(existing)},
                            },
                        )
                    elif route == "template-fill":
                        template = root / "input/template/template.pptx"
                        template.write_bytes(b"synthetic")
                        write_mapping(
                            root / "brief/presentation_brief.yaml",
                            {
                                "project_name": "External cache fixture",
                                "route": route,
                                "quality": "validated",
                                "inputs": {"template_path": str(template)},
                            },
                        )
                    args = normalise_public_request(
                        build_parser().parse_args(
                            [
                                "--project",
                                str(root),
                                "--route",
                                route,
                                "--quality",
                                "validated",
                                "--run-id",
                                f"RUN_{route.replace('-', '_').upper()}",
                            ]
                        )
                    )
                    with patch(
                        "academic_ppt.project_interface.parse_reference_style",
                        return_value={"visual_motif": "synthetic"},
                    ):
                        configured, _ = configure_project_args(args)
                    self.assertTrue(
                        paths_equal(configured.project_cache_root, root / "cache")
                    )
                    self.assertFalse(
                        path_is_within(configured.project_cache_root, ROOT / ".cache")
                    )


    def test_enhance_project_without_visual_brief_inherits_existing_style(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root, route="enhance", quality="quick")
            existing = root / "input/existing/existing.pptx"
            existing.write_bytes(b"synthetic existing deck")
            write_mapping(
                root / "brief/presentation_brief.yaml",
                {
                    "project_name": "Enhance fixture",
                    "route": "enhance",
                    "quality": "quick",
                    "inputs": {"existing_deck_path": "input/existing/existing.pptx"},
                },
            )
            args = normalise_public_request(
                build_parser().parse_args(
                    ["--project", str(root), "--route", "enhance", "--run-id", "RUN_B"]
                )
            )
            with patch(
                "academic_ppt.project_interface.parse_reference_style",
                return_value={"visual_motif": "preserved-existing-style"},
            ):
                _, run = configure_project_args(args)
            resolution = json.loads(
                (run.audit / "visual_direction_resolution.json").read_text(encoding="utf-8")
            )
            self.assertEqual(resolution["primary_authority"], "existing_deck_style")
            self.assertTrue(resolution["existing_deck_inherited"])
            self.assertFalse(resolution["human_visual_brief_approved"])

    def test_template_fill_visual_brief_cannot_override_native_brand(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root, route="template-fill", quality="validated")
            template = root / "input/template/official.pptx"
            template.write_bytes(b"synthetic authorized template")
            write_mapping(
                root / "brief/presentation_brief.yaml",
                {
                    "project_name": "Template fixture",
                    "route": "template-fill",
                    "quality": "validated",
                    "inputs": {"template_path": "input/template/official.pptx"},
                },
            )
            write_mapping(
                root / "brief/visual_brief.yaml",
                approved_visual_brief("template-fill"),
            )
            args = normalise_public_request(
                build_parser().parse_args(
                    [
                        "--project",
                        str(root),
                        "--route",
                        "template-fill",
                        "--run-id",
                        "RUN_C",
                    ]
                )
            )
            with patch(
                "academic_ppt.project_interface.parse_reference_style",
                return_value={"theme_colors": ["003B5C", "FFFFFF"]},
            ):
                configured, run = configure_project_args(args)
            resolution = json.loads(
                (run.audit / "visual_direction_resolution.json").read_text(encoding="utf-8")
            )
            self.assertTrue(resolution["human_visual_brief_approved"])
            self.assertEqual(resolution["template_brand_authority"], "authorized_template")
            self.assertIn("master_roles", resolution["template_protected_fields"])
            self.assertIn("theme_colors", resolution["template_protected_fields"])
            self.assertIsNone(getattr(configured, "style_profile", None))


class PortablePathContractTests(unittest.TestCase):
    def test_paths_equal_after_canonical_normalization(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            left = root / "private" / "engine"
            right = root / "private" / "nested" / ".." / "engine"
            self.assertEqual(canonical_test_path(left), canonical_test_path(right))
            self.assertTrue(paths_equal(left, right))

    def test_paths_differ_when_logically_distinct(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            self.assertFalse(
                paths_equal(root / "private" / "engine-a", root / "private" / "engine-b")
            )

    def test_no_private_absolute_path_required(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            workspace = Path(temporary) / "workspace"
            with patch.dict(os.environ, {"PPT_WORKSPACE_HOME": str(workspace)}, clear=False):
                resolved = resolve_project_reference("portable-study")
            expected = workspace / "projects" / "portable-study"
            self.assertTrue(paths_equal(resolved.root, expected))
            self.assertFalse(path_is_within(resolved.root, ROOT))


class ApprovedAssetsTests(unittest.TestCase):
    def test_approved_assets_are_non_evidence_and_must_remain_in_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root)
            asset = root / "approved_assets/concept.png"
            asset.write_bytes(b"synthetic")
            write_mapping(
                root / "approved_assets/registry.yaml",
                {
                    "assets": [
                        {
                            "asset_id": "ASSET-001",
                            "path": "concept.png",
                            "role": "conceptual_background",
                            "source": "user_approved",
                            "scientific_evidence": False,
                            "approved_by": "authorized-reviewer",
                            "approved_at": "2026-08-13T10:00:00+08:00",
                        }
                    ]
                },
            )
            rows = validate_approved_assets(ProjectPaths.from_root(root))
            self.assertEqual(len(rows), 1)
            self.assertIs(rows[0]["scientific_evidence"], False)

    def test_asset_cannot_claim_scientific_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            scaffold(root)
            (root / "approved_assets/concept.png").write_bytes(b"synthetic")
            write_mapping(
                root / "approved_assets/registry.yaml",
                {
                    "assets": [
                        {
                            "asset_id": "ASSET-001",
                            "path": "concept.png",
                            "role": "concept",
                            "source": "user",
                            "scientific_evidence": True,
                            "approved_by": "authorized-reviewer",
                            "approved_at": "2026-08-13T10:00:00+08:00",
                        }
                    ]
                },
            )
            with self.assertRaisesRegex(ProjectInterfaceError, "scientific_evidence=false"):
                validate_approved_assets(ProjectPaths.from_root(root))


if __name__ == "__main__":
    unittest.main()
