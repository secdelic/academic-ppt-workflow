from __future__ import annotations

import hashlib
import json
import math
import re
from collections.abc import Iterable, Mapping
from typing import Any


REQUIRED_FIELDS = (
    "visual_id",
    "slide_id",
    "visual_type",
    "semantic_role",
    "data_contract_id",
    "outcome",
    "estimand",
    "unit",
    "scale",
    "categories",
    "series",
    "estimate",
    "ci_low",
    "ci_high",
    "p_value",
    "interaction_p",
    "numerator",
    "denominator",
    "sample_size",
    "reference_group",
    "reference_value",
    "analysis_label",
    "highlight_rule",
    "annotation_rule",
    "source_ids",
    "wording_boundary",
    "editability_requirement",
    "preferred_backend",
    "fallback_backend",
)

BACKENDS = {
    "native_pptxgenjs",
    "svg_drawingml_adapter",
    "graphviz_adapter",
    "mermaid_adapter",
    "latex_omml_adapter",
}

RESULT_VISUAL_TYPES = {
    "event_rate_plot",
    "forest_plot",
    "sensitivity_plot",
    "subgroup_plot",
    "weighted_risk_curve",
    "distribution",
    "love_plot",
    "volcano",
    "enrichment",
    "funnel_plot",
}

CAUSAL_OR_MECHANISM_UPGRADE = re.compile(
    r"(?i)\b(caused|proved|confirmed mechanism|therapeutic target|"
    r"clinically validated|effect modification)\b"
)


def stable_id(prefix: str, *parts: object, length: int) -> str:
    payload = "\x1f".join(str(part) for part in parts)
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest().upper()
    return f"{prefix}-{digest[:length]}"


def make_visual_ir(
    *,
    slide_key: str,
    visual_type: str,
    semantic_role: str,
    source_ids: Iterable[str],
    data_contract_id: str | None = None,
    outcome: str | None = None,
    estimand: str | None = None,
    unit: str | None = None,
    scale: str | None = "linear",
    categories: Iterable[str] = (),
    series: Iterable[Mapping[str, Any]] = (),
    estimate: Iterable[float | None] = (),
    ci_low: Iterable[float | None] = (),
    ci_high: Iterable[float | None] = (),
    p_value: Iterable[float | None] = (),
    interaction_p: Iterable[float | None] = (),
    numerator: Iterable[float | None] = (),
    denominator: Iterable[float | None] = (),
    sample_size: float | None = None,
    reference_group: str | None = None,
    reference_value: float | None = None,
    analysis_label: Iterable[str] = (),
    highlight_rule: Mapping[str, Any] | None = None,
    annotation_rule: Mapping[str, Any] | None = None,
    wording_boundary: str = "Preserve the source evidence strength.",
    editability_requirement: str = "native_required",
    preferred_backend: str = "native_pptxgenjs",
    fallback_backend: str = "native_pptxgenjs",
) -> dict[str, Any]:
    sources = list(dict.fromkeys(str(item) for item in source_ids if str(item)))
    slide_id = stable_id("SLD", slide_key, length=20)
    visual_id = stable_id(
        "VIS", slide_id, visual_type, data_contract_id or "", length=16
    )
    return {
        "visual_id": visual_id,
        "slide_id": slide_id,
        "visual_type": visual_type,
        "semantic_role": semantic_role,
        "data_contract_id": data_contract_id,
        "outcome": outcome,
        "estimand": estimand,
        "unit": unit,
        "scale": scale,
        "categories": [str(item) for item in categories],
        "series": [dict(item) for item in series],
        "estimate": list(estimate),
        "ci_low": list(ci_low),
        "ci_high": list(ci_high),
        "p_value": list(p_value),
        "interaction_p": list(interaction_p),
        "numerator": list(numerator),
        "denominator": list(denominator),
        "sample_size": sample_size,
        "reference_group": reference_group,
        "reference_value": reference_value,
        "analysis_label": [str(item) for item in analysis_label],
        "highlight_rule": dict(highlight_rule or {}),
        "annotation_rule": dict(annotation_rule or {}),
        "source_ids": sources,
        "wording_boundary": wording_boundary,
        "editability_requirement": editability_requirement,
        "preferred_backend": preferred_backend,
        "fallback_backend": fallback_backend,
    }


def validate_visual_ir_schema(visual: Mapping[str, Any]) -> list[str]:
    errors: list[str] = []
    missing = [field for field in REQUIRED_FIELDS if field not in visual]
    errors.extend(f"missing field: {field}" for field in missing)
    if missing:
        return errors
    if not re.fullmatch(r"VIS-[A-F0-9]{16}", str(visual["visual_id"])):
        errors.append("visual_id does not match VIS-<16 hex>")
    if not re.fullmatch(r"SLD-[A-F0-9]{20}", str(visual["slide_id"])):
        errors.append("slide_id does not match SLD-<20 hex>")
    for field in (
        "categories",
        "series",
        "estimate",
        "ci_low",
        "ci_high",
        "p_value",
        "interaction_p",
        "numerator",
        "denominator",
        "analysis_label",
        "source_ids",
    ):
        if not isinstance(visual[field], list):
            errors.append(f"{field} must be a list")
    if not visual["source_ids"]:
        errors.append("source_ids must not be empty")
    if visual["preferred_backend"] not in BACKENDS:
        errors.append("preferred_backend is not registered")
    if visual["fallback_backend"] != "native_pptxgenjs":
        errors.append("fallback_backend must be native_pptxgenjs")
    if visual["editability_requirement"] not in {
        "native_required",
        "vector_acceptable",
        "raster_acceptable",
    }:
        errors.append("invalid editability_requirement")
    if not str(visual["wording_boundary"]).strip():
        errors.append("wording_boundary must not be empty")
    return errors


def _non_null(values: Iterable[Any]) -> list[float]:
    result = []
    for value in values:
        if value is not None:
            result.append(float(value))
    return result


def validate_visual_semantics(
    visual: Mapping[str, Any],
    *,
    protocol_without_results: bool = False,
) -> list[dict[str, str]]:
    issues: list[dict[str, str]] = []

    def fail(code: str, message: str) -> None:
        issues.append({"severity": "critical", "code": code, "message": message})

    for schema_error in validate_visual_ir_schema(visual):
        fail("VISUAL_IR_SCHEMA", schema_error)
    visual_type = str(visual.get("visual_type", ""))
    estimates = list(visual.get("estimate", []))
    lows = list(visual.get("ci_low", []))
    highs = list(visual.get("ci_high", []))
    series = [
        row for row in visual.get("series", []) if isinstance(row, Mapping)
    ]

    if visual_type in {"forest_plot", "sensitivity_plot", "subgroup_plot"}:
        if not estimates or len(estimates) != len(lows) or len(estimates) != len(highs):
            fail("FOREST_CI_REQUIRED", "Every estimate requires ci_low and ci_high")
        for index, (estimate, low, high) in enumerate(
            zip(estimates, lows, highs, strict=False), start=1
        ):
            if None in {estimate, low, high}:
                fail("FOREST_CI_REQUIRED", f"Row {index} contains a missing CI value")
                continue
            if not float(low) < float(estimate) < float(high):
                fail(
                    "FOREST_CI_ORDER",
                    f"Row {index} must satisfy ci_low < estimate < ci_high",
                )
        series_outcomes = {
            str(row.get("outcome")).strip()
            for row in series
            if row.get("outcome") not in {None, ""}
        }
        if visual.get("outcome"):
            series_outcomes.add(str(visual["outcome"]).strip())
        if len(series_outcomes) > 1:
            fail(
                "OUTCOME_MIXED_WITHOUT_FACET",
                "Different outcomes cannot share one unfaceted forest plot",
            )
        series_estimands = {
            str(row.get("estimand")).strip()
            for row in series
            if row.get("estimand") not in {None, ""}
        }
        if visual.get("estimand"):
            series_estimands.add(str(visual["estimand"]).strip())
        if len(series_estimands) > 1 and not visual.get("annotation_rule", {}).get(
            "show_estimand_labels"
        ):
            fail(
                "ESTIMAND_MIXED_WITHOUT_LABEL",
                "Different estimands require explicit labels or separate plots",
            )
        references = {
            str(row.get("reference_group")).strip()
            for row in series
            if row.get("reference_group") not in {None, ""}
        }
        if visual.get("reference_group"):
            references.add(str(visual["reference_group"]).strip())
        if len(references) > 1:
            fail(
                "REFERENCE_GROUP_INCONSISTENT",
                "Forest rows must use one reference group",
            )
        if visual.get("scale") in {"log", "log10"} and any(
            value <= 0
            for value in _non_null(
                [*estimates, *lows, *highs, visual.get("reference_value")]
            )
        ):
            fail("LOG_SCALE_NONPOSITIVE", "Log scale cannot receive non-positive values")

    if visual_type == "sensitivity_plot":
        labels = list(visual.get("analysis_label", []))
        if len(labels) != len(estimates) or len(set(labels)) != len(labels):
            fail(
                "SENSITIVITY_ANALYSIS_LABEL",
                "Sensitivity rows require unique labels from the analysis field",
            )
        if labels and not any(
            re.search(r"(?i)main|primary|reference|24\s*h", label)
            for label in labels
        ):
            fail(
                "SENSITIVITY_MAIN_NOT_DISTINGUISHED",
                "Main and sensitivity analyses are not distinguishable",
            )

    if visual_type == "subgroup_plot":
        labels = list(visual.get("categories", []))
        if not labels or any(not str(label).strip() for label in labels):
            fail("SUBGROUP_NAME_REQUIRED", "Every subgroup row requires a name")
        interaction = list(visual.get("interaction_p", []))
        if not interaction or all(value is None for value in interaction):
            fail(
                "SUBGROUP_INTERACTION_REQUIRED",
                "Subgroup plot requires interaction P",
            )
        subgroup_boundary = str(visual.get("wording_boundary", "")).lower()
        if (
            "effect modification" in subgroup_boundary
            and not re.search(
                r"(?:do not|does not|cannot|not sufficient to|insufficient to).{0,40}"
                r"effect modification",
                subgroup_boundary,
            )
        ):
            fail(
                "SUBGROUP_OVERSTATEMENT",
                "Point-estimate differences cannot be described as effect modification",
            )

    if visual_type == "event_rate_plot":
        numerator = list(visual.get("numerator", []))
        denominator = list(visual.get("denominator", []))
        if not numerator or len(numerator) != len(denominator) or len(numerator) != len(estimates):
            fail(
                "EVENT_RATE_DENOMINATOR",
                "Event rates require aligned numerator, denominator, and rate",
            )
        precision = int(visual.get("annotation_rule", {}).get("rate_decimals", 1))
        for index, (num, den, rate) in enumerate(
            zip(numerator, denominator, estimates, strict=False), start=1
        ):
            if None in {num, den, rate} or float(den) <= 0:
                fail("EVENT_RATE_DENOMINATOR", f"Invalid row {index}")
                continue
            expected = round(float(num) / float(den) * 100, precision)
            if not math.isclose(expected, round(float(rate), precision), abs_tol=10 ** (-precision)):
                fail(
                    "EVENT_RATE_INCONSISTENT",
                    f"Row {index} rate does not match numerator/denominator",
                )
        if denominator and visual.get("sample_size") is not None:
            if not math.isclose(
                sum(float(value) for value in denominator if value is not None),
                float(visual["sample_size"]),
                abs_tol=1e-9,
            ):
                fail(
                    "EVENT_RATE_SAMPLE_SIZE",
                    "sample_size must equal the ChartDataContract denominator sum",
                )

    if visual_type == "missingness_plot":
        if (
            not visual.get("categories")
            or len(visual.get("categories", [])) != len(visual.get("numerator", []))
            or len(visual.get("categories", [])) != len(estimates)
        ):
            fail(
                "MISSINGNESS_FIELDS",
                "Missingness requires variable, missing n, and missing percentage",
            )
        threshold = visual.get("annotation_rule", {}).get("decision_threshold")
        if threshold is not None and visual.get("highlight_rule", {}).get(
            "threshold_line"
        ) != threshold:
            fail(
                "MISSINGNESS_THRESHOLD",
                "A declared decision threshold requires a matching threshold line",
            )
        if re.search(
            r"(?i)imputation (?:removed|eliminated) bias",
            str(visual.get("wording_boundary", "")),
        ):
            fail(
                "MISSINGNESS_OVERSTATEMENT",
                "Imputation cannot be said to eliminate missing-data bias",
            )

    if visual_type == "matrix_2x2":
        rule = visual.get("annotation_rule", {})
        if (
            not rule.get("row_dimension")
            or not rule.get("column_dimension")
            or len(visual.get("categories", [])) != 4
        ):
            fail(
                "TRUE_2X2_REQUIRED",
                "A 2x2 matrix requires two real binary dimensions and four quadrants",
            )

    if protocol_without_results:
        visible_text = " ".join(
            [
                str(visual.get("outcome") or ""),
                str(visual.get("estimand") or ""),
                " ".join(visual.get("categories", [])),
                str(visual.get("annotation_rule", {})),
            ]
        )
        if (
            visual_type in RESULT_VISUAL_TYPES
            or visual.get("semantic_role") == "result"
            or CAUSAL_OR_MECHANISM_UPGRADE.search(visible_text)
        ):
            fail(
                "PROTOCOL_RESULT_FORBIDDEN",
                "A protocol without results cannot emit result visuals or conclusions",
            )
    return issues


def timeline_similarity(first: Mapping[str, Any], second: Mapping[str, Any]) -> float:
    def tokens(value: Mapping[str, Any]) -> set[str]:
        raw = " ".join(
            [
                str(value.get("semantic_role", "")),
                str(value.get("outcome", "")),
                " ".join(value.get("categories", [])),
                " ".join(value.get("analysis_label", [])),
            ]
        ).casefold()
        return set(re.findall(r"[\w\u3400-\u9fff]+", raw))

    left, right = tokens(first), tokens(second)
    if not left and not right:
        return 1.0
    return len(left & right) / max(1, len(left | right))


def validate_timeline_duplicates(
    visuals: Iterable[Mapping[str, Any]],
    *,
    threshold: float = 0.85,
) -> list[dict[str, str]]:
    timelines = [item for item in visuals if item.get("visual_type") == "timeline"]
    issues = []
    for index, first in enumerate(timelines):
        for second in timelines[index + 1 :]:
            similarity = timeline_similarity(first, second)
            if (
                similarity >= threshold
                and first.get("semantic_role") == second.get("semantic_role")
            ):
                issues.append(
                    {
                        "severity": "critical",
                        "code": "TIMELINE_SEMANTIC_DUPLICATE",
                        "message": (
                            f"{first.get('visual_id')} and {second.get('visual_id')} "
                            f"are {similarity:.0%} similar without distinct roles"
                        ),
                    }
                )
    return issues


def canonical_visual_hash(visuals: Iterable[Mapping[str, Any]]) -> str:
    payload = json.dumps(
        list(visuals), ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


__all__ = [
    "BACKENDS",
    "REQUIRED_FIELDS",
    "RESULT_VISUAL_TYPES",
    "canonical_visual_hash",
    "make_visual_ir",
    "stable_id",
    "timeline_similarity",
    "validate_timeline_duplicates",
    "validate_visual_ir_schema",
    "validate_visual_semantics",
]
