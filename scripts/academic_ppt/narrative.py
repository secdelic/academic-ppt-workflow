from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Mapping


NARRATIVE_MODES = (
    "scientific_problem",
    "conclusion_first",
    "thesis_defense",
    "journal_club",
    "clinical_protocol",
    "methods_teaching",
    "neutral_briefing",
)

REQUIRED_POLICY_FIELDS = frozenset(
    {
        "recommended_slide_order",
        "title_rules",
        "evidence_requirements",
        "conclusion_strength",
        "limitations_page_requirement",
        "citation_density",
        "speaker_pacing",
        "appendix_strategy",
    }
)

PRESENTATION_TYPE_TO_MODE = {
    "research_report": "scientific_problem",
    "bioinformatics_study": "scientific_problem",
    "conference_talk": "conclusion_first",
    "thesis_defense": "thesis_defense",
    "journal_club": "journal_club",
    "clinical_research_protocol": "clinical_protocol",
    "methods_teaching": "methods_teaching",
    "project_update": "neutral_briefing",
}

EVIDENCE_STRENGTHS = frozenset(
    {
        "demonstrated_mechanism",
        "supported_interpretation",
        "reported_fact",
        "observational_association",
        "computational_inference",
        "proposed_mechanism",
        "hypothesis_only",
        "unknown",
    }
)


class EvidenceBoundaryError(ValueError):
    """Raised when wording is stronger than its declared evidence class."""


def load_narrative_modes(path: str | Path) -> dict[str, dict[str, Any]]:
    """Load and validate the JSON-compatible narrative-mode contract."""

    source = Path(path)
    try:
        payload = json.loads(source.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"{source} must remain JSON-compatible YAML (valid JSON)"
        ) from exc
    if not isinstance(payload, dict) or not isinstance(payload.get("modes"), dict):
        raise ValueError("narrative config requires a top-level 'modes' mapping")

    modes = payload["modes"]
    missing_modes = sorted(set(NARRATIVE_MODES) - set(modes))
    extra_modes = sorted(set(modes) - set(NARRATIVE_MODES))
    if missing_modes or extra_modes:
        details = []
        if missing_modes:
            details.append("missing: " + ", ".join(missing_modes))
        if extra_modes:
            details.append("unknown: " + ", ".join(extra_modes))
        raise ValueError("Narrative mode set mismatch (" + "; ".join(details) + ")")

    validated: dict[str, dict[str, Any]] = {}
    for mode_name in NARRATIVE_MODES:
        policy = modes[mode_name]
        if not isinstance(policy, dict):
            raise ValueError(f"Narrative policy {mode_name!r} must be a mapping")
        missing_fields = sorted(REQUIRED_POLICY_FIELDS - set(policy))
        if missing_fields:
            raise ValueError(
                f"Narrative policy {mode_name!r} is missing: "
                + ", ".join(missing_fields)
            )
        if not isinstance(policy["recommended_slide_order"], list) or not all(
            isinstance(item, str) and item.strip()
            for item in policy["recommended_slide_order"]
        ):
            raise ValueError(
                f"{mode_name}.recommended_slide_order must contain non-empty strings"
            )
        if not isinstance(policy["title_rules"], list) or not policy["title_rules"]:
            raise ValueError(f"{mode_name}.title_rules must be a non-empty list")
        if not isinstance(policy["evidence_requirements"], list) or not policy[
            "evidence_requirements"
        ]:
            raise ValueError(
                f"{mode_name}.evidence_requirements must be a non-empty list"
            )
        if policy["limitations_page_requirement"] not in {
            "required",
            "recommended",
            "conditional",
        }:
            raise ValueError(
                f"{mode_name}.limitations_page_requirement has an invalid value"
            )
        validated[mode_name] = dict(policy)
    return validated


def select_narrative_mode(
    brief: Mapping[str, Any], explicit: str | None = None
) -> str:
    """Select one narrative mode without inferring scientific conclusions."""

    requested = explicit or brief.get("narrative_mode")
    if requested:
        if requested not in NARRATIVE_MODES:
            allowed = ", ".join(NARRATIVE_MODES)
            raise ValueError(
                f"Unknown narrative mode {requested!r}; expected one of: {allowed}"
            )
        return str(requested)
    presentation_type = str(brief.get("presentation_type", "")).strip()
    return PRESENTATION_TYPE_TO_MODE.get(presentation_type, "neutral_briefing")


def neutral_title(topic: str, evidence_strength: str = "unknown") -> str:
    """Create a descriptive title from a *topic*, never from an inferred result.

    The caller remains responsible for supplying a source-bound topic rather
    than an unsupported conclusion.  The suffix makes the declared evidence
    class visible when a conclusion title is not justified.
    """

    clean_topic = re.sub(r"\s+", " ", str(topic)).strip(" \t\r\n:：—-")
    if not clean_topic or clean_topic == "INFORMATION_REQUIRED":
        return "INFORMATION_REQUIRED"
    if evidence_strength not in EVIDENCE_STRENGTHS:
        raise ValueError(f"Unknown evidence strength: {evidence_strength!r}")
    suffixes = {
        "demonstrated_mechanism": "reported mechanistic evidence",
        "supported_interpretation": "evidence-supported interpretation",
        "reported_fact": "reported evidence",
        "observational_association": "observed association",
        "computational_inference": "computational inference",
        "proposed_mechanism": "proposed mechanism",
        "hypothesis_only": "hypothesis",
        "unknown": "evidence summary",
    }
    return f"{clean_topic} — {suffixes[evidence_strength]}"


_CAUSAL_PATTERNS = (
    r"\bcaus(?:e|es|ed|ing|al|ality)\b",
    r"\bled to\b",
    r"\bdrives?\b",
    r"\bprevents?\b",
    r"\bimproves?\b",
    r"\breduces?\b",
    r"导致",
    r"驱动",
    r"预防",
    r"改善",
)
_PROOF_PATTERNS = (
    r"\bpro(?:ve|ves|ved|ven|of)\b",
    r"\bconfirm(?:s|ed|ation)?\b",
    r"\bexperimentally validated\b",
    r"\bclinical(?:ly)? effective\b",
    r"证实",
    r"证明",
    r"实验验证",
    r"临床有效",
)
_MECHANISM_PROOF_PATTERNS = (
    r"\bdemonstrat(?:e|es|ed|ion)\b.{0,30}\bmechanism\b",
    r"\bmechanism\b.{0,30}\b(?:proven|confirmed)\b",
    r"机制(?:已)?(?:证实|证明|明确)",
)


def guard_evidence_strength(text: str, evidence_strength: str) -> str:
    """Return source wording unchanged or reject an evidence overstatement.

    This is intentionally a guard, not an automatic paraphraser.  Rewriting a
    scientific conclusion without its source could introduce a new claim.
    """

    if evidence_strength not in EVIDENCE_STRENGTHS:
        raise ValueError(f"Unknown evidence strength: {evidence_strength!r}")
    wording = str(text).strip()
    if not wording or wording == "INFORMATION_REQUIRED":
        return wording

    def has_any(patterns: tuple[str, ...]) -> bool:
        return any(re.search(pattern, wording, flags=re.IGNORECASE) for pattern in patterns)

    if evidence_strength in {
        "observational_association",
        "computational_inference",
        "proposed_mechanism",
        "hypothesis_only",
        "unknown",
    } and has_any(_CAUSAL_PATTERNS):
        raise EvidenceBoundaryError(
            f"{evidence_strength} wording cannot assert causality: {wording!r}"
        )
    if evidence_strength in {
        "computational_inference",
        "proposed_mechanism",
        "hypothesis_only",
        "unknown",
    } and has_any(_PROOF_PATTERNS):
        raise EvidenceBoundaryError(
            f"{evidence_strength} wording cannot assert validation or proof: {wording!r}"
        )
    if evidence_strength in {
        "supported_interpretation",
        "observational_association",
        "computational_inference",
        "proposed_mechanism",
        "hypothesis_only",
        "unknown",
    } and has_any(_MECHANISM_PROOF_PATTERNS):
        raise EvidenceBoundaryError(
            f"{evidence_strength} wording cannot assert a demonstrated mechanism: "
            f"{wording!r}"
        )
    return wording
