from __future__ import annotations

import json
import math
import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from time import perf_counter
from typing import Any, Callable, Iterator, Mapping


NOT_MEASURED = "NOT_MEASURED"


TIMED_STAGES: tuple[str, ...] = (
    "preflight",
    "source_snapshot",
    "cache_load",
    "delta_classification",
    "source_parse",
    "evidence_rebuild",
    "impact_resolution",
    "slide_planning",
    "changed_slide_generation",
    "powerpoint_apply",
    "changed_slide_render",
    "changed_slide_qa",
    "whole_deck_light_qa",
    "optional_pdf",
    "optional_cross_renderer",
    "package",
    "cache_promote",
)
REQUIRED_STAGES: tuple[str, ...] = (*TIMED_STAGES, "total")


STAGE_GROUPS: dict[str, tuple[str, ...]] = {
    "file_parsing": (
        "source_snapshot",
        "cache_load",
        "delta_classification",
        "source_parse",
    ),
    "generation": (
        "evidence_rebuild",
        "impact_resolution",
        "slide_planning",
        "changed_slide_generation",
    ),
    "powerpoint": (
        "powerpoint_apply",
        "changed_slide_render",
        "optional_pdf",
    ),
    "qa": (
        "changed_slide_qa",
        "whole_deck_light_qa",
        "optional_cross_renderer",
    ),
}


RUNTIME_BUDGETS: dict[str, dict[str, int | None]] = {
    "S": {
        "minimum_changed_slides": 0,
        "maximum_changed_slides": 3,
        "target_min_seconds": 0,
        "target_max_seconds": 15 * 60,
        "bottleneck_report_threshold_seconds": 30 * 60,
    },
    "M": {
        "minimum_changed_slides": 4,
        "maximum_changed_slides": 10,
        "target_min_seconds": 30 * 60,
        "target_max_seconds": 45 * 60,
        "bottleneck_report_threshold_seconds": 60 * 60,
    },
    "L": {
        "minimum_changed_slides": 11,
        "maximum_changed_slides": None,
        "target_min_seconds": 60 * 60,
        "target_max_seconds": 90 * 60,
        "bottleneck_report_threshold_seconds": 90 * 60,
    },
}


_SAFE_NAME = re.compile(r"^[a-z][a-z0-9_]{0,63}$")
_SAFE_ERROR_CODE = re.compile(r"^[A-Z][A-Z0-9_]{0,63}$")


class RuntimeProfileError(ValueError):
    """Raised when timing telemetry violates the de-identified contract."""


def complexity_for_changed_slides(changed_slide_count: int) -> str:
    if isinstance(changed_slide_count, bool) or changed_slide_count < 0:
        raise RuntimeProfileError("changed_slide_count must be a non-negative integer")
    if changed_slide_count <= 3:
        return "S"
    if changed_slide_count <= 10:
        return "M"
    return "L"


def _anonymous_number(value: Any, *, field_name: str) -> int | float | str:
    if value == NOT_MEASURED:
        return NOT_MEASURED
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise RuntimeProfileError(
            f"{field_name} must be a non-negative number or {NOT_MEASURED}"
        )
    if not math.isfinite(float(value)) or value < 0:
        raise RuntimeProfileError(
            f"{field_name} must be a finite, non-negative number"
        )
    return value


def _anonymous_counts(values: Mapping[str, Any]) -> dict[str, int | float | str]:
    result: dict[str, int | float | str] = {}
    for key, value in sorted(values.items()):
        if not _SAFE_NAME.fullmatch(str(key)):
            raise RuntimeProfileError(f"Unsafe anonymous counter name: {key!r}")
        result[str(key)] = _anonymous_number(value, field_name=str(key))
    return result


def _normalise_error_code(value: str | None) -> str:
    if not value:
        return "UNSPECIFIED_ERROR"
    candidate = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_").upper()
    if not candidate:
        candidate = "UNSPECIFIED_ERROR"
    candidate = candidate[:64]
    return candidate if _SAFE_ERROR_CODE.fullmatch(candidate) else "UNSPECIFIED_ERROR"


@dataclass
class _ActiveStage:
    name: str
    attempt: int
    counts: dict[str, int | float | str]
    status: str = "RUNNING"
    error_code: str | None = None

    def add_counts(self, **counts: Any) -> None:
        self.counts.update(_anonymous_counts(counts))

    def fail(self, error_code: str) -> None:
        self.status = "FAILED"
        self.error_code = _normalise_error_code(error_code)


@dataclass(frozen=True)
class StageRecord:
    name: str
    attempt: int
    elapsed_seconds: float
    status: str
    counts: dict[str, int | float | str] = field(default_factory=dict)
    error_code: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "attempt": self.attempt,
            "elapsed_seconds": round(self.elapsed_seconds, 6),
            "status": self.status,
            "counts": dict(self.counts),
            "error_code": self.error_code,
        }


class RuntimeProfiler:
    """Monotonic, de-identified stage profiler for incremental PPT updates.

    The profiler deliberately accepts only numeric anonymous counters.  It has
    no field for source text, prompts, filenames, private paths, patient facts,
    or free-form errors.  A controlled caller may supply measured model counts
    and token totals; otherwise those values remain ``NOT_MEASURED``.
    """

    def __init__(
        self,
        *,
        workflow_mode: str,
        original_slide_count: int,
        changed_slide_count: int,
        source_count: int,
        clock: Callable[[], float] = perf_counter,
    ) -> None:
        if workflow_mode not in {"fast_enhance", "full_validation"}:
            raise RuntimeProfileError(f"Unsupported workflow_mode: {workflow_mode!r}")
        for name, value in {
            "original_slide_count": original_slide_count,
            "changed_slide_count": changed_slide_count,
            "source_count": source_count,
        }.items():
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise RuntimeProfileError(f"{name} must be a non-negative integer")
        self.workflow_mode = workflow_mode
        self.original_slide_count = original_slide_count
        self.changed_slide_count = changed_slide_count
        self.source_count = source_count
        self.complexity_class = complexity_for_changed_slides(changed_slide_count)
        self._clock = clock
        self._started = float(clock())
        self._records: list[StageRecord] = []
        self._attempts: dict[str, int] = {}
        self._model_telemetry: dict[str, int | float | str] = {
            "model_calls": NOT_MEASURED,
            "input_tokens": NOT_MEASURED,
            "output_tokens": NOT_MEASURED,
            "total_tokens": NOT_MEASURED,
            "model_elapsed_seconds": NOT_MEASURED,
        }

    def _validate_stage_name(self, name: str) -> str:
        if name not in TIMED_STAGES:
            raise RuntimeProfileError(f"Unknown runtime stage: {name!r}")
        return name

    @contextmanager
    def stage(self, name: str, **counts: Any) -> Iterator[_ActiveStage]:
        """Time one stage and preserve an attempt record even on failure."""

        name = self._validate_stage_name(name)
        attempt = self._attempts.get(name, 0) + 1
        self._attempts[name] = attempt
        active = _ActiveStage(name, attempt, _anonymous_counts(counts))
        started = float(self._clock())
        try:
            yield active
        except BaseException as exc:
            elapsed = max(0.0, float(self._clock()) - started)
            self._records.append(
                StageRecord(
                    name=name,
                    attempt=attempt,
                    elapsed_seconds=elapsed,
                    status="FAILED",
                    counts=active.counts,
                    error_code=_normalise_error_code(type(exc).__name__),
                )
            )
            raise
        else:
            elapsed = max(0.0, float(self._clock()) - started)
            status = active.status
            if status not in {"PASSED", "FAILED", "SKIPPED"}:
                status = "PASSED"
            self._records.append(
                StageRecord(
                    name=name,
                    attempt=attempt,
                    elapsed_seconds=elapsed,
                    status=status,
                    counts=active.counts,
                    error_code=active.error_code if status == "FAILED" else None,
                )
            )

    def record_stage(
        self,
        name: str,
        elapsed_seconds: float,
        *,
        status: str = "PASSED",
        error_code: str | None = None,
        **counts: Any,
    ) -> StageRecord:
        """Record an externally measured deterministic or tool stage."""

        name = self._validate_stage_name(name)
        elapsed = float(_anonymous_number(elapsed_seconds, field_name="elapsed_seconds"))
        status = str(status).upper()
        if status not in {"PASSED", "FAILED", "SKIPPED"}:
            raise RuntimeProfileError(f"Unsupported stage status: {status!r}")
        attempt = self._attempts.get(name, 0) + 1
        self._attempts[name] = attempt
        record = StageRecord(
            name=name,
            attempt=attempt,
            elapsed_seconds=elapsed,
            status=status,
            counts=_anonymous_counts(counts),
            error_code=_normalise_error_code(error_code) if status == "FAILED" else None,
        )
        self._records.append(record)
        return record

    def mark_skipped(self, name: str, **counts: Any) -> StageRecord:
        return self.record_stage(name, 0.0, status="SKIPPED", **counts)

    def record_model_telemetry(
        self,
        *,
        model_calls: int | str,
        input_tokens: int | str = NOT_MEASURED,
        output_tokens: int | str = NOT_MEASURED,
        model_elapsed_seconds: float | str = NOT_MEASURED,
    ) -> None:
        values = {
            "model_calls": _anonymous_number(model_calls, field_name="model_calls"),
            "input_tokens": _anonymous_number(input_tokens, field_name="input_tokens"),
            "output_tokens": _anonymous_number(output_tokens, field_name="output_tokens"),
            "model_elapsed_seconds": _anonymous_number(
                model_elapsed_seconds, field_name="model_elapsed_seconds"
            ),
        }
        if isinstance(values["model_calls"], float) and not values["model_calls"].is_integer():
            raise RuntimeProfileError("model_calls must be an integer or NOT_MEASURED")
        if values["model_calls"] != NOT_MEASURED:
            values["model_calls"] = int(values["model_calls"])
        for field_name in ("input_tokens", "output_tokens"):
            value = values[field_name]
            if isinstance(value, float) and not value.is_integer():
                raise RuntimeProfileError(f"{field_name} must be an integer or NOT_MEASURED")
            if value != NOT_MEASURED:
                values[field_name] = int(value)
        if values["input_tokens"] == NOT_MEASURED or values["output_tokens"] == NOT_MEASURED:
            total_tokens: int | str = NOT_MEASURED
        else:
            total_tokens = int(values["input_tokens"]) + int(values["output_tokens"])
        self._model_telemetry = {**values, "total_tokens": total_tokens}

    def _stage_totals(self) -> dict[str, float]:
        totals = {name: 0.0 for name in TIMED_STAGES}
        for record in self._records:
            totals[record.name] += record.elapsed_seconds
        return {key: round(value, 6) for key, value in totals.items()}

    def _group_totals(self, stage_totals: Mapping[str, float]) -> dict[str, float | str]:
        groups: dict[str, float | str] = {
            name: round(sum(float(stage_totals[stage]) for stage in stages), 6)
            for name, stages in STAGE_GROUPS.items()
        }
        groups["model"] = self._model_telemetry["model_elapsed_seconds"]
        return groups

    def to_dict(self, *, total_seconds: float | None = None) -> dict[str, Any]:
        if total_seconds is None:
            total = max(0.0, float(self._clock()) - self._started)
        else:
            total = float(_anonymous_number(total_seconds, field_name="total_seconds"))
        totals = self._stage_totals()
        totals["total"] = round(total, 6)
        attempts_by_stage: dict[str, list[dict[str, Any]]] = {
            name: [] for name in TIMED_STAGES
        }
        for record in self._records:
            attempts_by_stage[record.name].append(record.as_dict())
        stages = []
        for name in REQUIRED_STAGES:
            if name == "total":
                stages.append(
                    {
                        "name": "total",
                        "elapsed_seconds": round(total, 6),
                        "status": "PASSED",
                        "attempt_count": 1,
                        "attempts": [
                            {
                                "name": "total",
                                "attempt": 1,
                                "elapsed_seconds": round(total, 6),
                                "status": "PASSED",
                                "counts": {},
                                "error_code": None,
                            }
                        ],
                    }
                )
                continue
            attempts = attempts_by_stage[name]
            if not attempts:
                status = "NOT_RUN"
            elif any(attempt["status"] == "FAILED" for attempt in attempts):
                status = "FAILED"
            elif all(attempt["status"] == "SKIPPED" for attempt in attempts):
                status = "SKIPPED"
            else:
                status = "PASSED"
            stages.append(
                {
                    "name": name,
                    "elapsed_seconds": totals[name],
                    "status": status,
                    "attempt_count": len(attempts),
                    "attempts": attempts,
                }
            )
        budget = dict(RUNTIME_BUDGETS[self.complexity_class])
        threshold = float(budget["bottleneck_report_threshold_seconds"] or math.inf)
        budget.update(
            {
                "complexity_class": self.complexity_class,
                "target_exceeded": total > float(budget["target_max_seconds"] or math.inf),
                "bottleneck_report_required": total > threshold,
            }
        )
        return {
            "schema_version": "2.5.1",
            "workflow_mode": self.workflow_mode,
            "anonymous_workload": {
                "original_slide_count": self.original_slide_count,
                "changed_slide_count": self.changed_slide_count,
                "source_count": self.source_count,
                "complexity_class": self.complexity_class,
            },
            "stages": stages,
            "stage_totals_seconds": totals,
            "stage_groups_seconds": self._group_totals(totals),
            "model_telemetry": dict(self._model_telemetry),
            "total_seconds": round(total, 6),
            "budget": budget,
            "privacy": {
                "deidentified_counts_only": True,
                "contains_source_text": False,
                "contains_prompts_or_responses": False,
                "contains_private_paths": False,
            },
        }

    def bottleneck_report(self, profile: Mapping[str, Any]) -> str:
        stage_totals = {
            name: seconds
            for name, seconds in profile["stage_totals_seconds"].items()
            if name != "total"
        }
        slowest_name, slowest_seconds = max(
            stage_totals.items(), key=lambda item: float(item[1])
        )
        groups = profile["stage_groups_seconds"]
        model_seconds = groups.get("model", NOT_MEASURED)
        return "\n".join(
            [
                "# Fast enhance runtime bottleneck report",
                "",
                f"- Complexity class: `{profile['budget']['complexity_class']}`",
                f"- Total runtime: `{float(profile['total_seconds']):.3f}` seconds",
                f"- Slowest stage: `{slowest_name}` (`{float(slowest_seconds):.3f}` seconds)",
                f"- Model time: `{model_seconds}` seconds",
                f"- File parsing time: `{float(groups['file_parsing']):.3f}` seconds",
                f"- PPT generation/planning time: `{float(groups['generation']):.3f}` seconds",
                f"- PowerPoint time: `{float(groups['powerpoint']):.3f}` seconds",
                f"- QA time: `{float(groups['qa']):.3f}` seconds",
                "",
                "The report contains anonymous counts and timings only; no source text, "
                "prompt, filename, private path, or project identity is recorded.",
                "",
            ]
        )

    @staticmethod
    def _write_json_atomic(path: Path, value: Mapping[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_name(path.name + ".tmp")
        temporary.write_text(
            json.dumps(value, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        temporary.replace(path)

    def write(
        self,
        destination: Path,
        *,
        total_seconds: float | None = None,
    ) -> dict[str, Path | None]:
        """Write ``runtime_profile.json`` and the conditional bottleneck report."""

        destination = Path(destination)
        if destination.suffix.lower() == ".json":
            profile_path = destination
            output_dir = destination.parent
        else:
            output_dir = destination
            profile_path = output_dir / "runtime_profile.json"
        output_dir.mkdir(parents=True, exist_ok=True)
        profile = self.to_dict(total_seconds=total_seconds)
        self._write_json_atomic(profile_path, profile)
        bottleneck_path: Path | None = None
        if profile["budget"]["bottleneck_report_required"]:
            bottleneck_path = output_dir / "runtime_bottleneck_report.md"
            temporary = bottleneck_path.with_name(bottleneck_path.name + ".tmp")
            temporary.write_text(self.bottleneck_report(profile), encoding="utf-8")
            temporary.replace(bottleneck_path)
        return {
            "runtime_profile": profile_path,
            "runtime_bottleneck_report": bottleneck_path,
        }
