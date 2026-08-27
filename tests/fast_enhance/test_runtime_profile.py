from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.runtime_profile import (  # noqa: E402
    NOT_MEASURED,
    REQUIRED_STAGES,
    RuntimeProfileError,
    RuntimeProfiler,
    complexity_for_changed_slides,
)


class FakeClock:
    def __init__(self) -> None:
        self.value = 0.0

    def __call__(self) -> float:
        return self.value

    def advance(self, seconds: float) -> None:
        self.value += seconds


class RuntimeProfilerTests(unittest.TestCase):
    def test_runtime_profile_has_all_stage_and_anonymous_workload_fields(self):
        clock = FakeClock()
        profiler = RuntimeProfiler(
            workflow_mode="fast_enhance",
            original_slide_count=21,
            changed_slide_count=7,
            source_count=4,
            clock=clock,
        )
        with profiler.stage("source_parse", input_count=2, output_count=2, cache_hits=2):
            clock.advance(1.25)
        with profiler.stage("changed_slide_generation", input_count=7, output_count=7):
            clock.advance(2.5)
        profile = profiler.to_dict()
        self.assertEqual(profile["anonymous_workload"]["complexity_class"], "M")
        self.assertEqual(
            {stage["name"] for stage in profile["stages"]}, set(REQUIRED_STAGES)
        )
        stage_by_name = {stage["name"]: stage for stage in profile["stages"]}
        self.assertEqual(stage_by_name["source_parse"]["status"], "PASSED")
        self.assertEqual(stage_by_name["optional_cross_renderer"]["status"], "NOT_RUN")
        self.assertEqual(profile["model_telemetry"]["model_calls"], NOT_MEASURED)
        self.assertAlmostEqual(profile["stage_groups_seconds"]["file_parsing"], 1.25)
        self.assertAlmostEqual(profile["stage_groups_seconds"]["generation"], 2.5)

    def test_deterministic_run_records_zero_model_calls_without_guessing_tokens(self):
        profiler = RuntimeProfiler(
            workflow_mode="fast_enhance",
            original_slide_count=21,
            changed_slide_count=3,
            source_count=2,
        )
        profiler.record_model_telemetry(
            model_calls=0,
            input_tokens=0,
            output_tokens=0,
            model_elapsed_seconds=0,
        )
        telemetry = profiler.to_dict(total_seconds=4.0)["model_telemetry"]
        self.assertEqual(telemetry["model_calls"], 0)
        self.assertEqual(telemetry["total_tokens"], 0)
        self.assertEqual(telemetry["model_elapsed_seconds"], 0)

    def test_failed_stage_uses_error_class_code_not_exception_text(self):
        clock = FakeClock()
        profiler = RuntimeProfiler(
            workflow_mode="fast_enhance",
            original_slide_count=21,
            changed_slide_count=7,
            source_count=3,
            clock=clock,
        )
        private_text = "private-source-title-should-never-appear"
        with self.assertRaisesRegex(RuntimeError, "private-source"):
            with profiler.stage("changed_slide_qa", input_count=1):
                clock.advance(0.5)
                raise RuntimeError(private_text)
        encoded = json.dumps(profiler.to_dict(), ensure_ascii=False)
        self.assertNotIn(private_text, encoded)
        self.assertIn("RUNTIMEERROR", encoded)

    def test_m_runtime_over_sixty_minutes_emits_bottleneck_report(self):
        profiler = RuntimeProfiler(
            workflow_mode="fast_enhance",
            original_slide_count=21,
            changed_slide_count=7,
            source_count=4,
        )
        profiler.record_stage("source_parse", 620.0, input_count=2, output_count=2)
        profiler.record_stage("slide_planning", 790.0, input_count=7, output_count=7)
        profiler.record_stage("powerpoint_apply", 930.0, input_count=7, output_count=25)
        profiler.record_stage("changed_slide_qa", 1261.0, input_count=7, output_count=7)
        profiler.record_model_telemetry(
            model_calls=NOT_MEASURED,
            input_tokens=NOT_MEASURED,
            output_tokens=NOT_MEASURED,
            model_elapsed_seconds=NOT_MEASURED,
        )
        with tempfile.TemporaryDirectory() as temporary:
            paths = profiler.write(Path(temporary), total_seconds=3601.0)
            bottleneck = paths["runtime_bottleneck_report"]
            self.assertIsNotNone(bottleneck)
            assert bottleneck is not None
            body = bottleneck.read_text(encoding="utf-8")
            self.assertIn("changed_slide_qa", body)
            self.assertIn("File parsing time", body)
            self.assertIn("PowerPoint time", body)
            self.assertIn("QA time", body)
            profile = json.loads(
                paths["runtime_profile"].read_text(encoding="utf-8")  # type: ignore[union-attr]
            )
            self.assertTrue(profile["budget"]["bottleneck_report_required"])

    def test_m_runtime_at_sixty_minutes_does_not_emit_mandatory_report(self):
        profiler = RuntimeProfiler(
            workflow_mode="fast_enhance",
            original_slide_count=21,
            changed_slide_count=7,
            source_count=4,
        )
        with tempfile.TemporaryDirectory() as temporary:
            paths = profiler.write(Path(temporary), total_seconds=3600.0)
            self.assertIsNone(paths["runtime_bottleneck_report"])

    def test_runtime_budget_classification(self):
        self.assertEqual(complexity_for_changed_slides(0), "S")
        self.assertEqual(complexity_for_changed_slides(3), "S")
        self.assertEqual(complexity_for_changed_slides(4), "M")
        self.assertEqual(complexity_for_changed_slides(10), "M")
        self.assertEqual(complexity_for_changed_slides(11), "L")

    def test_profiler_rejects_free_text_counters_and_unknown_stages(self):
        profiler = RuntimeProfiler(
            workflow_mode="fast_enhance",
            original_slide_count=21,
            changed_slide_count=7,
            source_count=4,
        )
        with self.assertRaises(RuntimeProfileError):
            profiler.record_stage("source_parse", 1.0, source_title="not allowed")
        with self.assertRaises(RuntimeProfileError):
            profiler.record_stage("arbitrary_private_stage", 1.0)


if __name__ == "__main__":
    unittest.main()
