from __future__ import annotations

import copy
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.fast_enhance import (  # noqa: E402
    FastEnhanceError,
    _assert_clinical_cache_is_ignored,
    _build_failure_record,
    _validate_cached_source_slide_order,
    _validate_changed_spec_source_registry,
    _load_retry_checkpoint,
    _write_checkpoint,
    _write_runtime_outputs,
    _blocking_geometry_messages,
    retry_stage_execution_policy,
    validate_fast_runtime_controls,
    validate_fast_cache_baseline,
)
from academic_ppt.runtime_profile import RuntimeProfiler  # noqa: E402
from academic_ppt.utils import sha256_file  # noqa: E402
from academic_ppt.routing import (  # noqa: E402
    RouteRequest,
    RouteValidationError,
    validate_route_request,
)
from academic_ppt.runner import WorkflowError, execute  # noqa: E402
from run_ppt_workflow import build_parser, normalise_public_request  # noqa: E402


class RuntimeModeRoutingTests(unittest.TestCase):
    def test_enhance_existing_defaults_to_fast(self) -> None:
        args = SimpleNamespace(route="enhance-existing", workflow_mode=None)
        expected = Path("synthetic-fast")
        with (
            patch(
                "academic_ppt.fast_enhance.execute_fast_enhance",
                return_value=expected,
            ) as fast,
            patch("academic_ppt.runner.execute_full_validation") as full,
        ):
            self.assertEqual(execute(args), expected)
        fast.assert_called_once_with(args)
        full.assert_not_called()

    def test_explicit_full_validation_preserves_legacy_enhance_path(self) -> None:
        args = SimpleNamespace(
            route="enhance-existing", workflow_mode="full_validation"
        )
        expected = Path("synthetic-full")
        with (
            patch(
                "academic_ppt.runner.execute_full_validation",
                return_value=expected,
            ) as full,
            patch("academic_ppt.fast_enhance.execute_fast_enhance") as fast,
        ):
            self.assertEqual(execute(args), expected)
        full.assert_called_once_with(args)
        fast.assert_not_called()

    def test_non_enhance_routes_default_to_full_validation(self) -> None:
        for route in ("generate", "create-style-profile", "fill-template"):
            with self.subTest(route=route):
                args = SimpleNamespace(route=route, workflow_mode=None)
                expected = Path(f"synthetic-{route}")
                with (
                    patch(
                        "academic_ppt.runner.execute_full_validation",
                        return_value=expected,
                    ) as full,
                    patch("academic_ppt.fast_enhance.execute_fast_enhance") as fast,
                ):
                    self.assertEqual(execute(args), expected)
                full.assert_called_once_with(args)
                fast.assert_not_called()

    def test_fast_mode_rejects_non_enhance_route(self) -> None:
        with self.assertRaisesRegex(WorkflowError, "restricted"):
            execute(SimpleNamespace(route="generate", workflow_mode="fast_enhance"))

    def test_cli_switches_are_opt_in_and_parse_explicitly(self) -> None:
        parser = build_parser()
        defaults = parser.parse_args([])
        self.assertIsNone(defaults.workflow_mode)
        self.assertEqual(defaults.route, "generate")
        self.assertEqual(defaults.quality, "validated")
        self.assertFalse(defaults.final_delivery)
        self.assertFalse(defaults.cross_renderer_validation)
        self.assertFalse(defaults.audit_full)
        self.assertFalse(defaults.manual_review_approved)
        explicit = parser.parse_args(
            [
                "--route",
                "enhance",
                "--quality",
                "full",
                "--final-delivery",
                "--cross-renderer-validation",
                "--audit-full",
                "--manual-review-approved",
            ]
        )
        self.assertTrue(explicit.final_delivery)
        self.assertTrue(explicit.cross_renderer_validation)
        self.assertTrue(explicit.audit_full)
        self.assertTrue(explicit.manual_review_approved)

    def test_public_routes_are_translated_without_exposing_internal_contracts(self) -> None:
        cases = (
            ("generate", "quick", "generate", "full_validation", False),
            ("template-create", "validated", "create-style-profile", "full_validation", False),
            ("template-fill", "validated", "fill-template", "full_validation", False),
            ("enhance", "quick", "enhance-existing", "fast_enhance", False),
            ("enhance", "validated", "enhance-existing", "fast_enhance", True),
            ("enhance", "full", "enhance-existing", "full_validation", False),
        )
        parser = build_parser()
        for public_route, quality, internal_route, mode, final_delivery in cases:
            with self.subTest(route=public_route, quality=quality):
                args = normalise_public_request(
                    parser.parse_args(["--route", public_route, "--quality", quality])
                )
                self.assertEqual(args.public_route, public_route)
                self.assertEqual(args.route, internal_route)
                self.assertEqual(args.workflow_mode, mode)
                self.assertEqual(args.final_delivery, final_delivery)

    def test_public_help_does_not_expose_low_level_fast_artifacts(self) -> None:
        help_text = build_parser().format_help()
        self.assertIn("template-create", help_text)
        self.assertIn("template-fill", help_text)
        self.assertNotIn("operation_plan.json", help_text)
        self.assertNotIn("Changed-slide-only candidate", help_text)

    def test_legacy_route_spelling_remains_accepted_for_automation_compatibility(self) -> None:
        args = normalise_public_request(
            build_parser().parse_args(
                ["--route", "enhance-existing", "--workflow-mode", "fast_enhance"]
            )
        )
        self.assertEqual(args.public_route, "enhance")
        self.assertEqual(args.route, "enhance-existing")
        self.assertEqual(args.workflow_mode, "fast_enhance")


class FastRuntimeControlTests(unittest.TestCase):
    def test_retry_slide_requires_an_isolated_matching_plan(self) -> None:
        args = SimpleNamespace(
            retry_slide="SLD-002",
            retry_stage=None,
            final_delivery=False,
            manual_review_approved=False,
            cross_renderer_validation=False,
            audit_full=False,
        )
        with self.assertRaisesRegex(FastEnhanceError, "isolated one-slide"):
            validate_fast_runtime_controls(args, ["SLD-001", "SLD-002"])
        controls = validate_fast_runtime_controls(args, ["SLD-002"])
        self.assertEqual(controls["retry_slide"], "SLD-002")

    def test_retry_stage_is_allowlisted_and_mutually_exclusive(self) -> None:
        valid = SimpleNamespace(
            retry_slide=None,
            retry_stage="changed_slide_qa",
            final_delivery=False,
            manual_review_approved=False,
            cross_renderer_validation=False,
            audit_full=False,
        )
        self.assertEqual(
            validate_fast_runtime_controls(valid, ["SLD-001"])["retry_stage"],
            "changed_slide_qa",
        )
        invalid = SimpleNamespace(**{**vars(valid), "retry_stage": "full_deck_qa"})
        with self.assertRaisesRegex(FastEnhanceError, "changed-only"):
            validate_fast_runtime_controls(invalid, ["SLD-001"])
        both = SimpleNamespace(
            **{**vars(valid), "retry_slide": "SLD-001"}
        )
        with self.assertRaisesRegex(FastEnhanceError, "mutually exclusive"):
            validate_fast_runtime_controls(both, ["SLD-001"])

    def test_manual_approval_cannot_bypass_final_delivery(self) -> None:
        args = SimpleNamespace(
            retry_slide=None,
            retry_stage=None,
            final_delivery=False,
            manual_review_approved=True,
            cross_renderer_validation=False,
            audit_full=False,
        )
        with self.assertRaisesRegex(FastEnhanceError, "requires --final-delivery"):
            validate_fast_runtime_controls(args, ["SLD-001"])

    def test_cross_renderer_and_full_audit_default_off(self) -> None:
        controls = validate_fast_runtime_controls(
            SimpleNamespace(), ["SLD-001"]
        )
        self.assertFalse(controls["cross_renderer_validation"])
        self.assertFalse(controls["audit_full"])
        self.assertFalse(controls["final_delivery"])

    def test_retry_execution_policy_proves_expensive_stage_skips(self) -> None:
        changed_qa = retry_stage_execution_policy("changed_slide_qa")
        for key in (
            "run_extract",
            "run_evidence",
            "run_slide_planning",
            "run_candidate_validation",
            "run_powerpoint_apply",
            "run_powerpoint_geometry",
        ):
            self.assertFalse(changed_qa[key], key)
        self.assertTrue(changed_qa["run_changed_slide_qa"])
        self.assertTrue(changed_qa["run_whole_deck_light_qa"])

        whole_deck = retry_stage_execution_policy("whole_deck_light_qa")
        for key in (
            "run_extract",
            "run_evidence",
            "run_slide_planning",
            "run_candidate_validation",
            "run_powerpoint_apply",
            "run_powerpoint_geometry",
            "run_changed_slide_qa",
        ):
            self.assertFalse(whole_deck[key], key)
        self.assertTrue(whole_deck["run_whole_deck_light_qa"])

    def test_content_shape_footer_intrusion_fails_closed(self) -> None:
        report = {
            "status": "PASS",
            "slide_width_pt": 960,
            "slide_height_pt": 540,
            "slides": [
                {
                    "slide_index": 1,
                    "slide_id": "SLD-0123456789ABCDEF0123",
                    "text_shapes": [],
                    "shapes": [
                        {
                            "shape_role": "content",
                            "left_pt": 60,
                            "top_pt": 470,
                            "width_pt": 300,
                            "height_pt": 45,
                        }
                    ],
                }
            ],
        }
        findings = _blocking_geometry_messages(report)
        self.assertTrue(
            any("content shape intrudes" in item for item in findings),
            findings,
        )


class ClinicalCacheGateTests(unittest.TestCase):
    def test_unignored_clinical_cache_fails_before_materialisation(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            repo = Path(raw)
            with self.assertRaisesRegex(FastEnhanceError, "not protected"):
                _assert_clinical_cache_is_ignored(repo, None)
            self.assertFalse((repo / ".cache").exists())

    def test_gitignored_repository_local_cache_is_accepted_without_creation(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            repo = Path(raw)
            (repo / ".gitignore").write_text("/.cache/\n", encoding="utf-8")
            expected = (repo / ".cache" / "project_state").resolve()
            self.assertEqual(
                _assert_clinical_cache_is_ignored(repo, None), expected
            )
            self.assertFalse((repo / ".cache").exists())

    def test_gitignore_negation_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            repo = Path(raw)
            (repo / ".gitignore").write_text(
                ".cache/\n!.cache/project_state/\n",  # self-containment: documentation-only
                encoding="utf-8",
            )
            with self.assertRaisesRegex(FastEnhanceError, "conflicting"):
                _assert_clinical_cache_is_ignored(repo, None)

    def test_clinical_failure_record_redacts_exception_text_and_paths(self) -> None:
        secret = chr(67) + chr(58) + chr(92) + chr(92).join(
            ("private", "patient", "source.docx")
        )
        try:
            raise FastEnhanceError(f"failed while reading {secret}")
        except FastEnhanceError as exc:
            clinical = _build_failure_record(
                exc, clinical_privacy_mode=True
            )
            ordinary = _build_failure_record(
                exc, clinical_privacy_mode=False
            )
        clinical_text = str(clinical)
        self.assertEqual(clinical["error"], "REDACTED_CLINICAL_FAILURE")
        self.assertTrue(clinical["clinical_detail_redacted"])
        self.assertRegex(clinical["traceback_sha256"], r"^[0-9a-f]{64}$")
        self.assertNotIn(secret, clinical_text)
        self.assertNotIn("failed while reading", clinical_text)
        self.assertNotIn("traceback", clinical)
        self.assertFalse(ordinary["clinical_detail_redacted"])
        self.assertIn(secret, ordinary["error"])


class CacheBaselineAndCheckpointTests(unittest.TestCase):
    def test_operation_source_order_must_match_cached_slidespec_order(self) -> None:
        state = {
            "slide_specs": [
                {"slide_id": "SLD-00000000000000000001"},
                {"slide_id": "SLD-00000000000000000002"},
            ]
        }
        _validate_cached_source_slide_order(
            state,
            ["SLD-00000000000000000001", "SLD-00000000000000000002"],
        )
        with self.assertRaisesRegex(FastEnhanceError, "cached SlideSpec order"):
            _validate_cached_source_slide_order(
                state,
                ["SLD-00000000000000000002", "SLD-00000000000000000001"],
            )

    def test_changed_spec_source_bindings_must_resolve_in_current_registry(self) -> None:
        registry = [
            {"source_id": "SRC-CURRENT", "relative_path": "documents/current.txt"}
        ]
        valid = [
            {
                "slide_id": "SLD-00000000000000000001",
                "source_bindings": [
                    {
                        "source_id": "SRC-CURRENT",
                        "source_file": "documents/current.txt",
                    }
                ],
            }
        ]
        _validate_changed_spec_source_registry(valid, registry)

        unknown = copy.deepcopy(valid)
        unknown[0]["source_bindings"][0]["source_id"] = "SRC-STALE"
        with self.assertRaisesRegex(FastEnhanceError, "unknown current source_id"):
            _validate_changed_spec_source_registry(unknown, registry)

        inconsistent = copy.deepcopy(valid)
        inconsistent[0]["source_bindings"][0]["source_file"] = "documents/other.txt"
        with self.assertRaisesRegex(FastEnhanceError, "unknown current source_file"):
            _validate_changed_spec_source_registry(inconsistent, registry)

    def test_fast_cache_must_match_a_passing_exact_source_deck(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            source = root / "source.pptx"
            source.write_bytes(b"synthetic deck package hash fixture")
            state = {
                "cache_contract_fingerprints": {
                    "brief_fingerprint": "0" * 64,
                    "scientific_definition_fingerprint": "1" * 64,
                    "extractor_fingerprint": "2" * 64,
                    "config_fingerprint": "3" * 64,
                    "code_fingerprint": "4" * 64,
                },
                "previous_qa_status": {"status": "PASS"},
                "pptx_state": {
                    "sha256": sha256_file(source),
                    "slide_count": 21,
                },
            }
            validate_fast_cache_baseline(
                state,
                source_pptx=source,
                expected_slide_count=21,
                current_contract_fingerprints=state["cache_contract_fingerprints"],
            )
            for invalid in (
                {**state, "previous_qa_status": {"status": "REVIEW"}},
                {
                    **state,
                    "pptx_state": {"sha256": "0" * 64, "slide_count": 21},
                },
                {
                    **state,
                    "pptx_state": {
                        "sha256": sha256_file(source),
                        "slide_count": 20,
                    },
                },
            ):
                with self.subTest(invalid=invalid):
                    with self.assertRaisesRegex(FastEnhanceError, "full_validation"):
                        validate_fast_cache_baseline(
                            invalid,
                            source_pptx=source,
                            expected_slide_count=21,
                            current_contract_fingerprints=state["cache_contract_fingerprints"],
                        )

            for field in state["cache_contract_fingerprints"]:
                with self.subTest(contract_field=field):
                    changed_contract = dict(state["cache_contract_fingerprints"])
                    changed_contract[field] = "f" * 64
                    with self.assertRaisesRegex(FastEnhanceError, "full_validation"):
                        validate_fast_cache_baseline(
                            state,
                            source_pptx=source,
                            expected_slide_count=21,
                            current_contract_fingerprints=changed_contract,
                        )

    def test_retry_checkpoint_binds_plan_cache_inputs_and_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            staging = root / "staging"
            output = root / "output"
            staging.mkdir()
            output.mkdir()
            source = root / "source.pptx"
            plan = root / "operation_plan.json"
            source.write_bytes(b"source")
            plan.write_text("{}\n", encoding="utf-8")
            (staging / "fast_retry_context.json").write_text(
                '{"parsed_objects":{"SRC":"synthetic"},"changed_specs":[{}],"merged_specs":[{}]}\n',
                encoding="utf-8",
            )
            updated = output / "updated.pptx"
            updated.write_bytes(b"updated")
            geometry = staging / "powerpoint_geometry.json"
            geometry.write_text('{"slides":[]}\n', encoding="utf-8")
            changed_qa = staging / "changed_slide_qa.json"
            changed_qa.write_text('{"status":"PASS"}\n', encoding="utf-8")
            manifest = {"sources": [], "hash_algorithm": "sha256"}
            checkpoint = staging / "fast_enhance_checkpoint.json"
            _write_checkpoint(
                checkpoint,
                completed_stage="changed_slide_qa",
                existing_pptx=source,
                operation_plan_path=plan,
                cache_generation="GEN-1111111111111111",
                current_source_manifest=manifest,
                updated_pptx=updated,
                geometry_path=geometry,
            )
            verified = _load_retry_checkpoint(
                checkpoint,
                retry_stage="whole_deck_light_qa",
                existing_pptx=source,
                operation_plan_path=plan,
                cache_generation="GEN-1111111111111111",
                current_source_manifest=manifest,
                output_dir=output,
                staging_root=staging,
            )
            self.assertEqual(verified["updated_pptx"], updated)
            updated.write_bytes(b"tampered")
            with self.assertRaisesRegex(FastEnhanceError, "hash mismatch"):
                _load_retry_checkpoint(
                    checkpoint,
                    retry_stage="whole_deck_light_qa",
                    existing_pptx=source,
                    operation_plan_path=plan,
                    cache_generation="GEN-1111111111111111",
                    current_source_manifest=manifest,
                    output_dir=output,
                    staging_root=staging,
                )


class MinimalArtifactContractTests(unittest.TestCase):
    def test_minimal_artifact_set_keeps_runtime_profile_in_staging(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            staging = root / "staging"
            output = root / "output"
            staging.mkdir()
            output.mkdir()
            expected = {
                "updated.pptx",
                "change_plan.json",
                "slide_diff.json",
                "changed_slide_qa.md",
                "manual_review_checklist.md",
                "execution_summary.md",
            }
            for name in expected:
                (output / name).write_bytes(b"synthetic")
            profiler = RuntimeProfiler(
                workflow_mode="fast_enhance",
                original_slide_count=21,
                changed_slide_count=7,
                source_count=2,
            )
            profiler.record_model_telemetry(
                model_calls=0,
                input_tokens=0,
                output_tokens=0,
                model_elapsed_seconds=0,
            )
            runtime_paths, _profile, audit_dir = _write_runtime_outputs(
                profiler,
                staging_root=staging,
                output_dir=output,
                audit_full=False,
            )
            self.assertIsNone(audit_dir)
            self.assertTrue(Path(runtime_paths["runtime_profile"]).is_file())
            self.assertFalse((output / "runtime_profile.json").exists())
            self.assertEqual(
                {path.name for path in output.iterdir() if path.is_file()}, expected
            )


class RouteOptionContractTests(unittest.TestCase):
    def _request(self, **options: object) -> RouteRequest:
        return RouteRequest(
            route="enhance-existing",
            brief={"project_name": "synthetic"},
            existing_deck_path="synthetic.pptx",
            options=options,
        )

    def test_retry_controls_are_fast_only_and_mutually_exclusive(self) -> None:
        with self.assertRaisesRegex(RouteValidationError, "mutually exclusive"):
            validate_route_request(
                self._request(
                    retry_slide="SLD-001",
                    retry_stage="changed_slide_qa",
                )
            )
        with self.assertRaisesRegex(RouteValidationError, "fast_enhance"):
            validate_route_request(
                self._request(
                    workflow_mode="full_validation", retry_slide="SLD-001"
                )
            )

    def test_manual_review_requires_final_delivery(self) -> None:
        with self.assertRaisesRegex(RouteValidationError, "requires final_delivery"):
            validate_route_request(
                self._request(manual_review_approved=True)
            )


if __name__ == "__main__":
    unittest.main()
