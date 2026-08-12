from __future__ import annotations

import csv
import argparse
import hashlib
import json
import os
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR = Path(__file__).with_name("generate_synthetic_decks.cjs")


def _stable_id(logical_key: str) -> str:
    digest = hashlib.sha256(f"fast-enhance-fixture:{logical_key}".encode("utf-8")).hexdigest()
    return f"SLD-{digest[:20].upper()}"


def _content_hash(payload: dict[str, Any]) -> str:
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _slide(
    logical_key: str,
    title: str,
    message: str,
    *,
    revision: int = 1,
    source_ids: tuple[str, ...] = ("SRC-SYN-A",),
    layout: str = "content",
) -> dict[str, Any]:
    core = {
        "logical_key": logical_key,
        "title": title,
        "key_message": message,
        "source_ids": list(source_ids),
        "layout": layout,
    }
    return {
        "slide_id": _stable_id(logical_key),
        "slide_revision": revision,
        **core,
        "content_hash": _content_hash({**core, "revision": revision}),
    }


def build_fixture_definition() -> dict[str, Any]:
    """Return an anonymous 21-slide baseline plus seven-slide delta contract."""

    rows = [
        ("opening/purpose", "Synthetic update contract", "The fixture tests structure only.", "divider"),
        ("context/scope", "Scope remains bounded", "Only registered synthetic sources are used.", "content"),
        ("context/constraints", "Constraints are explicit", "No external service participates.", "content"),
        ("system/baseline", "Baseline state is registered", "The baseline contains stable semantic identities.", "content"),
        ("system/components", "Components remain traceable", "Each component binds to an opaque source ID.", "content"),
        ("system/interfaces", "Interfaces are deterministic", "Mechanical operations do not require inference.", "content"),
        ("method/snapshot", "Snapshots use content hashes", "Modification time alone does not establish identity.", "content"),
        ("method/delta", "Delta classification is explicit", "Changed objects are separated from reusable objects.", "content"),
        ("method/impact", "Impact follows registered edges", "Stable slide IDs identify affected consumers.", "content"),
        ("method/plan", "The operation plan is auditable", "Keep, replace, and insert actions are explicit.", "content"),
        ("evidence/registry", "Synthetic evidence stays local", "The registry contains no personal information.", "content"),
        ("evidence/bindings", "Bindings share one authority", "Notes and manifests derive from the same binding.", "content"),
        ("evidence/verification", "Verification precedes mutation", "Inputs are checked before a working copy changes.", "content"),
        ("results/summary", "Baseline execution is measurable", "Stage timings are recorded without source text.", "content"),
        ("results/cache", "Cache reuse is conditional", "Only validated unchanged objects are reused.", "content"),
        ("results/render", "Rendering remains bounded", "Only affected slides enter the generation queue.", "content"),
        ("quality/changed", "Changed slides receive deep QA", "Scientific and geometric checks remain mandatory.", "content"),
        ("quality/deck", "The complete deck receives light QA", "Openability and package integrity are retained.", "content"),
        ("discussion/questions", "Review questions remain explicit", "Unresolved decisions stay under human control.", "content"),
        ("conclusion/next_steps", "The next action is minimal", "Failed slides can be retried independently.", "content"),
        ("closing/boundary", "Synthetic validation is not real-world approval", "Human approval remains outside this fixture.", "divider"),
    ]
    baseline = [
        _slide(key, title, message, layout=layout) for key, title, message, layout in rows
    ]
    by_key = {slide["logical_key"]: slide for slide in baseline}

    replacements = [
        _slide(
            "method/delta",
            "Only modified and new sources are parsed",
            "Unchanged sources produce zero parser calls.",
            revision=2,
            source_ids=("SRC-SYN-B",),
        ),
        _slide(
            "results/summary",
            "Seven affected slides define the render queue",
            "Three replacements and four insertions are isolated.",
            revision=2,
            source_ids=("SRC-SYN-B", "SRC-SYN-C"),
        ),
        _slide(
            "discussion/questions",
            "Manual review focuses on changed content",
            "Unchanged pages are not repeatedly reinterpreted.",
            revision=2,
            source_ids=("SRC-SYN-C",),
        ),
    ]
    insertions = [
        _slide(
            "supplement/change_summary",
            "The supplement contributes a bounded change",
            "A new source affects only registered downstream objects.",
            source_ids=("SRC-SYN-B",),
        ),
        _slide(
            "supplement/impact_summary",
            "Impact resolution names exact consumers",
            "No title or page-number heuristic is used.",
            source_ids=("SRC-SYN-B",),
        ),
        _slide(
            "supplement/qa_scope",
            "QA scope follows the affected set",
            "Deep checks remain local while deck integrity stays global.",
            source_ids=("SRC-SYN-C",),
        ),
        _slide(
            "supplement/retry_scope",
            "Retry reuses prior passing artifacts",
            "Only one failed semantic slide returns to its failed stage.",
            source_ids=("SRC-SYN-C",),
        ),
    ]
    candidate = [*replacements, *insertions]

    insertion_after = {
        insertions[0]["slide_id"]: by_key["evidence/verification"]["slide_id"],
        insertions[1]["slide_id"]: insertions[0]["slide_id"],
        insertions[2]["slide_id"]: by_key["results/render"]["slide_id"],
        insertions[3]["slide_id"]: insertions[2]["slide_id"],
    }
    replacements_by_id = {slide["slide_id"]: slide for slide in replacements}
    insertions_by_anchor: dict[str, list[dict[str, Any]]] = {}
    for slide in insertions:
        insertions_by_anchor.setdefault(insertion_after[slide["slide_id"]], []).append(slide)

    final: list[dict[str, Any]] = []
    for slide in baseline:
        final.append(replacements_by_id.get(slide["slide_id"], slide))
        anchor = final[-1]["slide_id"]
        while anchor in insertions_by_anchor:
            inserted = insertions_by_anchor[anchor]
            if len(inserted) != 1:
                raise AssertionError("Synthetic fixture anchor must be unambiguous")
            final.append(inserted[0])
            anchor = inserted[0]["slide_id"]

    replacement_ids = set(replacements_by_id)
    operations: list[dict[str, Any]] = []
    candidate_indexes = {
        slide["slide_id"]: index for index, slide in enumerate(candidate, start=1)
    }
    for slide in baseline:
        if slide["slide_id"] in replacement_ids:
            operations.append(
                {
                    "action": "REPLACE",
                    "slide_id": slide["slide_id"],
                    "candidate_index": candidate_indexes[slide["slide_id"]],
                    "expected_previous_content_hash": slide["content_hash"],
                }
            )
        else:
            operations.append({"action": "KEEP", "slide_id": slide["slide_id"]})
    for slide in insertions:
        operations.append(
            {
                "action": "INSERT_AFTER",
                "slide_id": slide["slide_id"],
                "after_slide_id": insertion_after[slide["slide_id"]],
                "candidate_index": candidate_indexes[slide["slide_id"]],
                "layout_donor_slide_id": by_key["results/render"]["slide_id"],
            }
        )

    action_counts = {
        action: sum(1 for item in operations if item["action"] == action)
        for action in ("KEEP", "REPLACE", "INSERT_AFTER")
    }
    return {
        "schema_version": "2.5.1",
        "synthetic_only": True,
        "baseline_slides": baseline,
        "candidate_slides": candidate,
        "final_slides": final,
        "operations": operations,
        "expected": {
            "original_slide_count": 21,
            "changed_slide_count": 7,
            "final_slide_count": 25,
            "operation_counts": action_counts,
            "parser_calls_on_unchanged_rerun": 0,
            "libreoffice_calls_in_fast_mode": 0,
        },
    }


@dataclass(frozen=True)
class SyntheticFixture:
    root: Path
    prior_inputs: Path
    current_inputs: Path
    fixture_contract: Path
    baseline_slide_specs: Path
    candidate_slide_specs: Path
    final_slide_specs: Path
    baseline_pptx: Path | None
    candidate_pptx: Path | None
    generator_manifest: Path | None


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _write_sources(prior: Path, current: Path) -> None:
    prior.mkdir(parents=True, exist_ok=False)
    current.mkdir(parents=True, exist_ok=False)
    brief = {
        "presentation_type": "software_verification",
        "audience": "internal engineering reviewers",
        "language": "en-US",
        "privacy_mode": "synthetic_only",
    }
    _write_json(prior / "brief.json", brief)
    _write_json(current / "brief.json", brief)
    rows = [
        {"module": "parser", "state": "ready", "revision": "1"},
        {"module": "renderer", "state": "ready", "revision": "1"},
        {"module": "validator", "state": "ready", "revision": "1"},
    ]
    for root in (prior, current):
        with (root / "component_matrix.csv").open(
            "w", encoding="utf-8", newline=""
        ) as handle:
            writer = csv.DictWriter(handle, fieldnames=["module", "state", "revision"])
            writer.writeheader()
            writer.writerows(rows)
    _write_json(
        prior / "supplement.json",
        {"revision": 1, "labels": ["cache", "impact", "retry"]},
    )
    _write_json(
        current / "supplement.json",
        {"revision": 2, "labels": ["cache", "impact", "retry", "timing"]},
    )
    _write_json(
        current / "addendum.json",
        {"revision": 1, "checks": ["changed_slide_qa", "lightweight_deck_qa"]},
    )


def discover_node_runtime(repo_root: Path = REPO_ROOT) -> tuple[Path, Path] | None:
    """Discover a local Node/PptxGenJS pair without installing anything."""

    explicit_node = os.environ.get("ACADEMIC_PPT_NODE")
    explicit_modules = os.environ.get("ACADEMIC_PPT_NODE_MODULES")
    candidates: list[tuple[Path, Path]] = []
    if explicit_node and explicit_modules:
        candidates.append((Path(explicit_node), Path(explicit_modules)))
    system_node = shutil.which("node")
    if system_node:
        candidates.append((Path(system_node), repo_root / "node_modules"))
    bundled = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "node"
    )
    candidates.append((bundled / "bin" / "node.exe", bundled / "node_modules"))
    candidates.append((bundled / "bin" / "node", bundled / "node_modules"))
    for node, modules in candidates:
        package_dir = modules / "pptxgenjs"
        if node.is_file() and package_dir.is_dir():
            return node.resolve(), modules.resolve()
    return None


def generate_synthetic_pptx(
    fixture: SyntheticFixture,
    *,
    timeout_seconds: int = 60,
    repo_root: Path = REPO_ROOT,
) -> SyntheticFixture:
    runtime = discover_node_runtime(repo_root)
    if runtime is None:
        raise RuntimeError("Local Node/PptxGenJS runtime is unavailable; installation is prohibited")
    node, modules = runtime
    baseline_pptx = fixture.root / "generated" / "baseline_21_slides.pptx"
    candidate_pptx = fixture.root / "generated" / "candidate_7_changed_slides.pptx"
    manifest = fixture.root / "generated" / "generation_manifest.json"
    baseline_pptx.parent.mkdir(parents=True, exist_ok=True)
    environment = dict(os.environ)
    prior_node_path = environment.get("NODE_PATH", "")
    environment["NODE_PATH"] = str(modules) + (os.pathsep + prior_node_path if prior_node_path else "")
    command = [
        str(node),
        str(GENERATOR),
        str(fixture.baseline_slide_specs),
        str(fixture.candidate_slide_specs),
        str(baseline_pptx),
        str(candidate_pptx),
        str(manifest),
    ]
    completed = subprocess.run(
        command,
        cwd=repo_root,
        env=environment,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout_seconds,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"Synthetic PptxGenJS fixture generation failed ({completed.returncode}): "
            f"{(completed.stdout + completed.stderr).strip()}"
        )
    for path in (baseline_pptx, candidate_pptx, manifest):
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(f"Synthetic generator did not create {path.name}")
    return SyntheticFixture(
        root=fixture.root,
        prior_inputs=fixture.prior_inputs,
        current_inputs=fixture.current_inputs,
        fixture_contract=fixture.fixture_contract,
        baseline_slide_specs=fixture.baseline_slide_specs,
        candidate_slide_specs=fixture.candidate_slide_specs,
        final_slide_specs=fixture.final_slide_specs,
        baseline_pptx=baseline_pptx,
        candidate_pptx=candidate_pptx,
        generator_manifest=manifest,
    )


def materialize_synthetic_fixture(
    root: Path,
    *,
    render_pptx: bool = False,
    repo_root: Path = REPO_ROOT,
) -> SyntheticFixture:
    """Create the complete fixture under a caller-owned temporary directory."""

    root = Path(root).resolve()
    if root.exists() and any(root.iterdir()):
        raise FileExistsError(f"Fixture root must be empty: {root}")
    root.mkdir(parents=True, exist_ok=True)
    definition = build_fixture_definition()
    prior = root / "inputs_prior"
    current = root / "inputs_current"
    _write_sources(prior, current)
    contract = root / "fixture_contract.json"
    baseline = root / "baseline_slide_specs.json"
    candidate = root / "candidate_slide_specs.json"
    final = root / "final_slide_specs.json"
    _write_json(
        contract,
        {
            "schema_version": definition["schema_version"],
            "synthetic_only": True,
            "expected": definition["expected"],
            "operations": definition["operations"],
        },
    )
    _write_json(baseline, {"slides": definition["baseline_slides"]})
    _write_json(candidate, {"slides": definition["candidate_slides"]})
    _write_json(final, {"slides": definition["final_slides"]})
    fixture = SyntheticFixture(
        root=root,
        prior_inputs=prior,
        current_inputs=current,
        fixture_contract=contract,
        baseline_slide_specs=baseline,
        candidate_slide_specs=candidate,
        final_slide_specs=final,
        baseline_pptx=None,
        candidate_pptx=None,
        generator_manifest=None,
    )
    return generate_synthetic_pptx(fixture, repo_root=repo_root) if render_pptx else fixture


__all__ = [
    "SyntheticFixture",
    "build_fixture_definition",
    "discover_node_runtime",
    "generate_synthetic_pptx",
    "materialize_synthetic_fixture",
]


def _main() -> int:
    parser = argparse.ArgumentParser(
        description="Materialize an anonymous 21-slide baseline and seven-slide delta fixture"
    )
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--render-pptx", action="store_true")
    args = parser.parse_args()
    fixture = materialize_synthetic_fixture(args.out, render_pptx=args.render_pptx)
    definition = build_fixture_definition()
    print(
        json.dumps(
            {
                "synthetic_only": True,
                **definition["expected"],
                "pptx_generated": fixture.baseline_pptx is not None,
                "external_service_usage": 0,
            },
            ensure_ascii=False,
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
