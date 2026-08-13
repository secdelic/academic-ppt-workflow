"""Portable path comparisons used only by the test suite.

Both operands must pass through the same normalization pipeline before a test
compares them.  This avoids Windows CI differences such as short/long temp
paths, path case, and slash direction without changing production path rules.
"""

from __future__ import annotations

import os
from os import PathLike
from pathlib import Path


def canonical_test_path(value: str | PathLike[str]) -> str:
    """Return a canonical, comparison-only representation of *value*."""

    resolved = Path(value).resolve(strict=False)
    normalized = os.path.normcase(os.path.normpath(str(resolved)))
    return normalized.replace("\\", "/")


def paths_equal(left: str | PathLike[str], right: str | PathLike[str]) -> bool:
    """Compare two paths after applying the identical canonical pipeline."""

    return canonical_test_path(left) == canonical_test_path(right)


def path_is_within(
    candidate: str | PathLike[str], parent: str | PathLike[str]
) -> bool:
    """Return whether *candidate* is beneath *parent* after normalization."""

    candidate_value = canonical_test_path(candidate)
    parent_value = canonical_test_path(parent)
    try:
        common_value = os.path.commonpath((candidate_value, parent_value))
        return canonical_test_path(common_value) == parent_value
    except ValueError:
        # Different Windows drives cannot have a common path.
        return False
