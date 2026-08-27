from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Iterable, Mapping


class ChangeImpactError(ValueError):
    """Raised when a change-impact graph or delta document is malformed."""


def _clean_id(value: Any, label: str) -> str:
    result = str(value).strip()
    if not result:
        raise ChangeImpactError(f"{label} cannot be empty")
    return result


def _normalise_ids(values: Iterable[Any]) -> tuple[str, ...]:
    return tuple(sorted({_clean_id(value, "entity id") for value in values}))


@dataclass(frozen=True)
class ChangeImpactResult:
    strategy: str
    affected_slides: tuple[str, ...]
    escalation_reasons: tuple[str, ...]
    changed_entities: Mapping[str, tuple[str, ...]]
    traversed_edges: tuple[tuple[str, str, str], ...]

    @property
    def requires_full_rebuild(self) -> bool:
        return self.strategy == "FULL"

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": "academic-ppt-change-impact-result/1",
            "strategy": self.strategy,
            "requires_full_rebuild": self.requires_full_rebuild,
            "affected_slides": list(self.affected_slides),
            "escalation_reasons": list(self.escalation_reasons),
            "changed_entities": {
                key: list(values) for key, values in self.changed_entities.items()
            },
            "traversed_edges": [
                {"edge_type": edge_type, "source": source, "target": target}
                for edge_type, source, target in self.traversed_edges
            ],
        }


class ChangeImpactGraph:
    """Explicit source -> claim -> visual -> slide dependency graph.

    Fast enhance is permitted only when every changed upstream entity has a
    complete path to a registered slide.  Missing lineage, any removal, or a
    deck-wide presentation/scientific change fails closed to ``FULL``.
    """

    def __init__(self) -> None:
        self._source_to_claim: dict[str, set[str]] = defaultdict(set)
        self._claim_to_visual: dict[str, set[str]] = defaultdict(set)
        self._visual_to_slide: dict[str, set[str]] = defaultdict(set)
        self._sources: set[str] = set()
        self._claims: set[str] = set()
        self._visuals: set[str] = set()
        self._slides: set[str] = set()

    def register_source(self, source_id: str) -> None:
        self._sources.add(_clean_id(source_id, "source_id"))

    def register_claim(self, claim_id: str) -> None:
        self._claims.add(_clean_id(claim_id, "claim_id"))

    def register_visual(self, visual_id: str) -> None:
        self._visuals.add(_clean_id(visual_id, "visual_id"))

    def register_slide(self, slide_id: str) -> None:
        self._slides.add(_clean_id(slide_id, "slide_id"))

    def add_source_claim(self, source_id: str, claim_id: str) -> None:
        source = _clean_id(source_id, "source_id")
        claim = _clean_id(claim_id, "claim_id")
        self._sources.add(source)
        self._claims.add(claim)
        self._source_to_claim[source].add(claim)

    def add_claim_visual(self, claim_id: str, visual_id: str) -> None:
        claim = _clean_id(claim_id, "claim_id")
        visual = _clean_id(visual_id, "visual_id")
        self._claims.add(claim)
        self._visuals.add(visual)
        self._claim_to_visual[claim].add(visual)

    def add_visual_slide(self, visual_id: str, slide_id: str) -> None:
        visual = _clean_id(visual_id, "visual_id")
        slide = _clean_id(slide_id, "slide_id")
        self._visuals.add(visual)
        self._slides.add(slide)
        self._visual_to_slide[visual].add(slide)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": "academic-ppt-change-impact-graph/1",
            "nodes": {
                "sources": sorted(self._sources),
                "claims": sorted(self._claims),
                "visuals": sorted(self._visuals),
                "slides": sorted(self._slides),
            },
            "edges": {
                "source_to_claim": [
                    {"source_id": source, "claim_id": claim}
                    for source in sorted(self._source_to_claim)
                    for claim in sorted(self._source_to_claim[source])
                ],
                "claim_to_visual": [
                    {"claim_id": claim, "visual_id": visual}
                    for claim in sorted(self._claim_to_visual)
                    for visual in sorted(self._claim_to_visual[claim])
                ],
                "visual_to_slide": [
                    {"visual_id": visual, "slide_id": slide}
                    for visual in sorted(self._visual_to_slide)
                    for slide in sorted(self._visual_to_slide[visual])
                ],
            },
        }

    @classmethod
    def from_dict(cls, document: Mapping[str, Any]) -> "ChangeImpactGraph":
        if not isinstance(document, Mapping):
            raise ChangeImpactError("ChangeImpactGraph document must be an object")
        graph = cls()
        nodes = document.get("nodes", {})
        if not isinstance(nodes, Mapping):
            raise ChangeImpactError("ChangeImpactGraph nodes must be an object")
        registrations = (
            ("sources", graph.register_source),
            ("claims", graph.register_claim),
            ("visuals", graph.register_visual),
            ("slides", graph.register_slide),
        )
        for name, register in registrations:
            values = nodes.get(name, [])
            if not isinstance(values, list):
                raise ChangeImpactError(f"nodes.{name} must be a list")
            for value in values:
                register(value)

        edges = document.get("edges", {})
        if not isinstance(edges, Mapping):
            raise ChangeImpactError("ChangeImpactGraph edges must be an object")
        edge_specs = (
            ("source_to_claim", "source_id", "claim_id", graph.add_source_claim),
            ("claim_to_visual", "claim_id", "visual_id", graph.add_claim_visual),
            ("visual_to_slide", "visual_id", "slide_id", graph.add_visual_slide),
        )
        for edge_name, left_name, right_name, add_edge in edge_specs:
            rows = edges.get(edge_name, [])
            if not isinstance(rows, list):
                raise ChangeImpactError(f"edges.{edge_name} must be a list")
            for row in rows:
                if not isinstance(row, Mapping):
                    raise ChangeImpactError(f"edges.{edge_name} entries must be objects")
                add_edge(row.get(left_name, ""), row.get(right_name, ""))
        return graph

    def _trace_claim(
        self,
        claim: str,
        affected: set[str],
        reasons: set[str],
        traversed: set[tuple[str, str, str]],
    ) -> None:
        if claim not in self._claims:
            reasons.add(f"MISSING_CLAIM_LINEAGE:{claim}")
            return
        visuals = self._claim_to_visual.get(claim, set())
        if not visuals:
            reasons.add(f"MISSING_CLAIM_TO_VISUAL:{claim}")
            return
        for visual in visuals:
            traversed.add(("claim_to_visual", claim, visual))
            self._trace_visual(visual, affected, reasons, traversed)

    def _trace_visual(
        self,
        visual: str,
        affected: set[str],
        reasons: set[str],
        traversed: set[tuple[str, str, str]],
    ) -> None:
        if visual not in self._visuals:
            reasons.add(f"MISSING_VISUAL_LINEAGE:{visual}")
            return
        slides = self._visual_to_slide.get(visual, set())
        if not slides:
            reasons.add(f"MISSING_VISUAL_TO_SLIDE:{visual}")
            return
        for slide in slides:
            traversed.add(("visual_to_slide", visual, slide))
            if slide not in self._slides:
                reasons.add(f"MISSING_SLIDE_LINEAGE:{slide}")
            else:
                affected.add(slide)

    def analyze(
        self,
        *,
        changed_sources: Iterable[str] = (),
        changed_claims: Iterable[str] = (),
        changed_visuals: Iterable[str] = (),
        changed_slides: Iterable[str] = (),
        removed_entities: Iterable[str] = (),
        deck_wide_change: bool = False,
        deck_wide_reason: str = "DECK_WIDE_CHANGE",
    ) -> ChangeImpactResult:
        changed = {
            "sources": _normalise_ids(changed_sources),
            "claims": _normalise_ids(changed_claims),
            "visuals": _normalise_ids(changed_visuals),
            "slides": _normalise_ids(changed_slides),
            "removed": _normalise_ids(removed_entities),
        }
        affected: set[str] = set()
        reasons: set[str] = set()
        traversed: set[tuple[str, str, str]] = set()

        if deck_wide_change:
            reasons.add(_clean_id(deck_wide_reason, "deck_wide_reason"))
        for entity in changed["removed"]:
            reasons.add(f"REMOVED_ENTITY:{entity}")

        for slide in changed["slides"]:
            if slide not in self._slides:
                reasons.add(f"MISSING_SLIDE_LINEAGE:{slide}")
            else:
                affected.add(slide)
        for visual in changed["visuals"]:
            self._trace_visual(visual, affected, reasons, traversed)
        for claim in changed["claims"]:
            self._trace_claim(claim, affected, reasons, traversed)
        for source in changed["sources"]:
            if source not in self._sources:
                reasons.add(f"MISSING_SOURCE_LINEAGE:{source}")
                continue
            claims = self._source_to_claim.get(source, set())
            if not claims:
                reasons.add(f"MISSING_SOURCE_TO_CLAIM:{source}")
                continue
            for claim in claims:
                traversed.add(("source_to_claim", source, claim))
                self._trace_claim(claim, affected, reasons, traversed)

        strategy = "FULL" if reasons else "TARGETED"
        if strategy == "FULL":
            affected = set(self._slides)
        return ChangeImpactResult(
            strategy=strategy,
            affected_slides=tuple(sorted(affected)),
            escalation_reasons=tuple(sorted(reasons)),
            changed_entities=changed,
            traversed_edges=tuple(sorted(traversed)),
        )

    def analyze_delta(
        self,
        delta_manifest: Mapping[str, Any],
        *,
        changed_claims: Iterable[str] = (),
        changed_visuals: Iterable[str] = (),
        changed_slides: Iterable[str] = (),
        deck_wide_change: bool = False,
        deck_wide_reason: str = "DECK_WIDE_CHANGE",
    ) -> ChangeImpactResult:
        rows = delta_manifest.get("entries", [])
        if not isinstance(rows, list):
            raise ChangeImpactError("Delta manifest entries must be a list")
        changed_sources: list[str] = []
        removed: list[str] = []
        allowed = {"UNCHANGED", "MODIFIED", "NEW", "REMOVED"}
        for row in rows:
            if not isinstance(row, Mapping):
                raise ChangeImpactError("Delta manifest entries must be objects")
            source = _clean_id(row.get("source_key", ""), "source_key")
            status = str(row.get("status", "")).strip().upper()
            if status not in allowed:
                raise ChangeImpactError(f"Unsupported delta status: {status or '<empty>'}")
            if status in {"MODIFIED", "NEW"}:
                changed_sources.append(source)
            elif status == "REMOVED":
                removed.append(f"source:{source}")
        return self.analyze(
            changed_sources=changed_sources,
            changed_claims=changed_claims,
            changed_visuals=changed_visuals,
            changed_slides=changed_slides,
            removed_entities=removed,
            deck_wide_change=deck_wide_change,
            deck_wide_reason=deck_wide_reason,
        )

