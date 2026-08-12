from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.change_impact import (  # noqa: E402
    ChangeImpactGraph,
)
from academic_ppt.project_cache import (  # noqa: E402
    CachePolicyError,
    CacheStateError,
    ProjectCache,
    build_cache_contract_fingerprints,
    build_project_state,
    build_source_manifest,
    compute_delta_manifest,
    derive_opaque_project_key,
    validate_cache_root,
    validate_project_state,
)


def _digest(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _manifest(rows: list[dict[str, object]]) -> dict[str, object]:
    return {
        "schema_version": "academic-ppt-source-manifest/1",
        "hash_algorithm": "sha256",
        "sources": rows,
    }


def _fingerprints(seed: str = "baseline") -> dict[str, str]:
    return {
        "brief_fingerprint": _digest(f"brief:{seed}".encode()),
        "scientific_definition_fingerprint": _digest(f"science:{seed}".encode()),
        "extractor_fingerprint": _digest(f"extractor:{seed}".encode()),
        "config_fingerprint": _digest(f"config:{seed}".encode()),
        "code_fingerprint": _digest(f"code:{seed}".encode()),
    }


class SourceDeltaTests(unittest.TestCase):
    def test_sha256_is_the_only_change_authority(self) -> None:
        previous = _manifest(
            [
                {"source_key": "SRC-A", "sha256": _digest(b"old"), "size_bytes": 3},
                {"source_key": "SRC-B", "sha256": _digest(b"same"), "size_bytes": 4},
                {"source_key": "SRC-C", "sha256": _digest(b"removed")},
            ]
        )
        current = _manifest(
            [
                {"source_key": "SRC-A", "sha256": _digest(b"new"), "size_bytes": 3},
                # Deliberately different metadata with the same authoritative hash.
                {"source_key": "SRC-B", "sha256": _digest(b"same"), "size_bytes": 999},
                {"source_key": "SRC-D", "sha256": _digest(b"added")},
            ]
        )
        delta = compute_delta_manifest(previous, current)
        status = {row["source_key"]: row["status"] for row in delta["entries"]}
        self.assertEqual(
            status,
            {
                "SRC-A": "MODIFIED",
                "SRC-B": "UNCHANGED",
                "SRC-C": "REMOVED",
                "SRC-D": "NEW",
            },
        )
        self.assertEqual(delta["parse_source_keys"], ["SRC-A", "SRC-D"])
        self.assertEqual(delta["unchanged_source_keys"], ["SRC-B"])
        self.assertEqual(delta["removed_source_keys"], ["SRC-C"])

    def test_source_manifest_uses_registered_keys_and_relative_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "synthetic" / "supplement.txt"
            source.parent.mkdir()
            source.write_bytes(b"fully synthetic fixture")
            manifest = build_source_manifest({"SRC-SYN-001": source}, relative_to=root)
            row = manifest["sources"][0]
            self.assertEqual(row["source_key"], "SRC-SYN-001")
            self.assertEqual(row["relative_path"], "synthetic/supplement.txt")
            self.assertEqual(row["sha256"], _digest(b"fully synthetic fixture"))
            self.assertNotIn(str(root.resolve()), json.dumps(manifest))


class ProjectCacheTests(unittest.TestCase):
    @staticmethod
    def _state(qa_status: str) -> dict[str, object]:
        return build_project_state(
            cache_contract_fingerprints=_fingerprints(),
            source_manifest=_manifest([]),
            parsed_document_objects={},
            evidence_registry=[],
            slide_specs=[],
            visual_specs=[],
            master_layout_map={},
            font_map={},
            image_registry=[],
            source_bindings=[],
            previous_qa_status={"status": qa_status},
        )

    def test_opaque_identity_and_atomic_generation_pointers(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            repo = Path(temporary)
            identity = "SYNTHETIC CASE DISPLAY NAME"
            cache = ProjectCache.from_identity(repo, identity)
            self.assertRegex(cache.project_key, r"^PRJ-[A-F0-9]{32}$")
            self.assertNotIn("SYNTHETIC", cache.project_key)

            first = cache.commit_generation(
                self._state("PASS"),
                generation_id="GEN-1111111111111111",
            )
            second = cache.commit_generation(
                self._state("REVIEW"),
                generation_id="GEN-2222222222222222",
            )
            self.assertEqual(first["current"], "GEN-1111111111111111")
            self.assertEqual(second["current"], "GEN-2222222222222222")
            self.assertEqual(second["previous"], "GEN-1111111111111111")
            self.assertEqual(
                cache.load_current()["previous_qa_status"]["status"], "REVIEW"
            )
            self.assertEqual(
                cache.load_previous()["previous_qa_status"]["status"], "PASS"
            )
            self.assertTrue(
                (cache.generations_dir / "GEN-1111111111111111" / "state.json").is_file()
            )

            all_cache_text = "\n".join(
                path.read_text(encoding="utf-8")
                for path in cache.cache_root.rglob("*.json")
            )
            self.assertNotIn(identity, all_cache_text)
            self.assertFalse(list(cache.cache_root.rglob("*.pkl")))
            self.assertFalse(list(cache.cache_root.rglob("*.pickle")))
            current_manifest = _manifest(
                [{"source_key": "SRC-NEW", "sha256": _digest(b"new")}]
            )
            delta = cache.compute_delta(current_manifest)
            self.assertEqual(delta["parse_source_keys"], ["SRC-NEW"])

    def test_non_json_state_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            cache = ProjectCache.from_identity(Path(temporary), "synthetic")
            with self.assertRaisesRegex(CacheStateError, "JSON values only"):
                state = self._state("PASS")
                state["not_json"] = b"bytes"
                cache.stage_generation(state)

    def test_incomplete_or_inconsistent_state_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            cache = ProjectCache.from_identity(Path(temporary), "synthetic")
            with self.assertRaisesRegex(CacheStateError, "missing required"):
                cache.stage_generation({"source_manifest": _manifest([])})
            state = self._state("PASS")
            state["source_hashes"] = {"SRC-NOT-IN-MANIFEST": "0" * 64}
            with self.assertRaisesRegex(CacheStateError, "exactly match"):
                cache.stage_generation(state)

    def test_all_cache_contract_fingerprints_are_required_sha256(self) -> None:
        state = self._state("PASS")
        del state["cache_contract_fingerprints"]["code_fingerprint"]
        with self.assertRaisesRegex(CacheStateError, "fields are invalid"):
            validate_project_state(state)
        state = self._state("PASS")
        state["cache_contract_fingerprints"]["brief_fingerprint"] = "not-a-digest"
        with self.assertRaisesRegex(CacheStateError, "not SHA-256"):
            validate_project_state(state)

    def test_pre_fingerprint_schema_fails_closed(self) -> None:
        state = self._state("PASS")
        state["schema_version"] = "academic-ppt-project-state-payload/1"
        with self.assertRaisesRegex(CacheStateError, "full_validation"):
            validate_project_state(state)

    def test_fingerprints_change_for_brief_science_extractor_config_and_code(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            repo = Path(temporary)
            module_dir = repo / "scripts" / "academic_ppt"
            config_dir = repo / "config"
            module_dir.mkdir(parents=True)
            config_dir.mkdir()
            (module_dir / "extractors.py").write_text("EXTRACTOR = 1\n", encoding="utf-8")
            (module_dir / "runtime.py").write_text("RUNTIME = 1\n", encoding="utf-8")
            (config_dir / "workflow.yaml").write_text('{"mode":"safe"}\n', encoding="utf-8")
            brief = {"project_name": "synthetic", "outcome_definition": "AKI"}
            baseline = build_cache_contract_fingerprints(repo_root=repo, brief=brief)

            changed_brief = build_cache_contract_fingerprints(
                repo_root=repo,
                brief={**brief, "language": "zh-CN"},
            )
            changed_science = build_cache_contract_fingerprints(
                repo_root=repo,
                brief={**brief, "outcome_definition": "mortality"},
            )
            self.assertNotEqual(
                baseline["brief_fingerprint"], changed_brief["brief_fingerprint"]
            )
            self.assertNotEqual(
                baseline["scientific_definition_fingerprint"],
                changed_science["scientific_definition_fingerprint"],
            )

            (module_dir / "extractors.py").write_text("EXTRACTOR = 2\n", encoding="utf-8")
            changed_extractor = build_cache_contract_fingerprints(
                repo_root=repo, brief=brief
            )
            self.assertNotEqual(
                baseline["extractor_fingerprint"],
                changed_extractor["extractor_fingerprint"],
            )
            (config_dir / "workflow.yaml").write_text('{"mode":"strict"}\n', encoding="utf-8")
            changed_config = build_cache_contract_fingerprints(
                repo_root=repo, brief=brief
            )
            self.assertNotEqual(
                changed_extractor["config_fingerprint"],
                changed_config["config_fingerprint"],
            )
            (module_dir / "runtime.py").write_text("RUNTIME = 2\n", encoding="utf-8")
            changed_code = build_cache_contract_fingerprints(
                repo_root=repo, brief=brief
            )
            self.assertNotEqual(
                changed_config["code_fingerprint"], changed_code["code_fingerprint"]
            )

    def test_cache_root_is_repository_local_and_not_in_publish_trees(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            repo = Path(temporary)
            self.assertEqual(
                validate_cache_root(repo),
                (repo / ".cache" / "project_state").resolve(),
            )
            for blocked in ("tests", "regression", "benchmark", "output", "archive"):
                with self.subTest(blocked=blocked):
                    with self.assertRaises(CachePolicyError):
                        validate_cache_root(repo, repo / blocked / "project_state")

    def test_project_key_requires_keyed_opaque_digest(self) -> None:
        first = derive_opaque_project_key("synthetic-a", secret=b"x" * 32)
        second = derive_opaque_project_key("synthetic-a", secret=b"x" * 32)
        other = derive_opaque_project_key("synthetic-b", secret=b"x" * 32)
        self.assertEqual(first, second)
        self.assertNotEqual(first, other)
        self.assertNotIn("synthetic", first.casefold())


def _complete_graph() -> ChangeImpactGraph:
    graph = ChangeImpactGraph()
    graph.add_source_claim("SRC-1", "CLM-1")
    graph.add_claim_visual("CLM-1", "VIS-1")
    graph.add_visual_slide("VIS-1", "SLD-1")
    graph.register_slide("SLD-2")
    return graph


class ChangeImpactGraphTests(unittest.TestCase):
    def test_complete_lineage_targets_only_affected_slide(self) -> None:
        graph = _complete_graph()
        result = graph.analyze(changed_sources=["SRC-1"])
        self.assertEqual(result.strategy, "TARGETED")
        self.assertEqual(result.affected_slides, ("SLD-1",))
        self.assertFalse(result.requires_full_rebuild)
        self.assertEqual(len(result.traversed_edges), 3)

    def test_missing_lineage_escalates_to_full(self) -> None:
        graph = _complete_graph()
        graph.register_source("SRC-NO-CLAIM")
        result = graph.analyze(changed_sources=["SRC-NO-CLAIM"])
        self.assertTrue(result.requires_full_rebuild)
        self.assertEqual(result.affected_slides, ("SLD-1", "SLD-2"))
        self.assertIn("MISSING_SOURCE_TO_CLAIM:SRC-NO-CLAIM", result.escalation_reasons)

    def test_removed_or_deck_wide_change_escalates_to_full(self) -> None:
        graph = _complete_graph()
        removed = graph.analyze(removed_entities=["source:SRC-OLD"])
        self.assertTrue(removed.requires_full_rebuild)
        self.assertIn("REMOVED_ENTITY:source:SRC-OLD", removed.escalation_reasons)
        deck_wide = graph.analyze(
            changed_sources=["SRC-1"],
            deck_wide_change=True,
            deck_wide_reason="MASTER_LAYOUT_CHANGED",
        )
        self.assertTrue(deck_wide.requires_full_rebuild)
        self.assertIn("MASTER_LAYOUT_CHANGED", deck_wide.escalation_reasons)

    def test_delta_manifest_ignores_unchanged_and_escalates_removed(self) -> None:
        graph = _complete_graph()
        unchanged = graph.analyze_delta(
            {"entries": [{"source_key": "SRC-1", "status": "UNCHANGED"}]}
        )
        self.assertEqual(unchanged.strategy, "TARGETED")
        self.assertEqual(unchanged.affected_slides, ())
        removed = graph.analyze_delta(
            {"entries": [{"source_key": "SRC-1", "status": "REMOVED"}]}
        )
        self.assertTrue(removed.requires_full_rebuild)

    def test_graph_json_round_trip_preserves_explicit_edges(self) -> None:
        graph = _complete_graph()
        document = graph.to_dict()
        rebuilt = ChangeImpactGraph.from_dict(json.loads(json.dumps(document)))
        result = rebuilt.analyze(changed_claims=["CLM-1"])
        self.assertEqual(result.strategy, "TARGETED")
        self.assertEqual(result.affected_slides, ("SLD-1",))


if __name__ == "__main__":
    unittest.main()
