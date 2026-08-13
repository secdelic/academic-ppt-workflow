"""Unified, repository-independent runtime workspace for tests.

The helper intentionally creates every mutable workflow area below one
operating-system temporary root.  Tests must never use ignored directories in
the developer checkout as implicit fixtures or scratch space.
"""

from __future__ import annotations

import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator


@dataclass(frozen=True)
class TestRuntimeWorkspace:
    root: Path
    staging: Path
    audit: Path
    benchmark: Path
    output: Path
    cache: Path


@contextmanager
def temporary_runtime_workspace() -> Iterator[TestRuntimeWorkspace]:
    """Yield a complete test runtime tree and remove it on exit."""

    with tempfile.TemporaryDirectory(prefix="academic-ppt-test-") as temporary:
        root = Path(temporary).resolve()
        workspace = TestRuntimeWorkspace(
            root=root,
            staging=root / "staging",
            audit=root / "audit",
            benchmark=root / "benchmark",
            output=root / "output",
            cache=root / "cache",
        )
        for path in (
            workspace.staging,
            workspace.audit,
            workspace.benchmark,
            workspace.output,
            workspace.cache,
        ):
            path.mkdir(parents=True, exist_ok=False)
        yield workspace


__all__ = ["TestRuntimeWorkspace", "temporary_runtime_workspace"]
