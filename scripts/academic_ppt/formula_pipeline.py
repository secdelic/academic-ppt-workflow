"""Deterministic formula manifests and capability probes.

The module plans a formula rendering route only. It never installs software,
executes a renderer, opens a shell, or injects content into a PPTX package.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
from pathlib import Path
from typing import Any, Callable


REQUESTED_METHODS = ("native_omml", "latex_fallback", "plain_text")
SOURCE_NOTATIONS = ("latex", "plain_text", "omml")
DEFAULT_FORMULA_TIMEOUT_SECONDS = 30

FORMULA_EXECUTABLE_ALLOWLIST = (
    "pandoc",
    "latex",
    "pdflatex",
    "xelatex",
    "dvisvgm",
    "pdftocairo",
)

UNSAFE_TEX_REJECT_LIST = (
    r"\input",
    r"\include",
    r"\includegraphics",
    r"\openin",
    r"\openout",
    r"\read",
    r"\write",
    r"\write18",
    r"\immediate",
    r"\special",
    r"\usepackage",
    r"\documentclass",
    r"\begin{filecontents",
    r"\href",
    r"\url",
)

_UNSAFE_PATTERNS = (
    re.compile(r"\\input(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\include(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\includegraphics(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\open(?:in|out)(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\(?:read|write)(?:\s*18)?(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\immediate(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\special(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\usepackage(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\documentclass(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\begin\s*\{\s*filecontents\*?\s*\}", re.IGNORECASE),
    re.compile(r"\\href(?![A-Za-z@])", re.IGNORECASE),
    re.compile(r"\\url(?![A-Za-z@])", re.IGNORECASE),
)


def _canonical_hash(value: Any) -> str:
    payload = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def validate_formula_source(source_text: str, source_notation: str = "latex") -> None:
    """Reject empty, control-character, or unsafe TeX formula input."""

    if source_notation not in SOURCE_NOTATIONS:
        raise ValueError(
            f"source_notation must be one of: {', '.join(SOURCE_NOTATIONS)}"
        )
    if not isinstance(source_text, str) or not source_text.strip():
        raise ValueError("Formula source_text must be non-empty text")
    if len(source_text) > 20000:
        raise ValueError("Formula source_text exceeds the 20000-character limit")
    if any(ord(char) < 32 and char not in "\t\r\n" for char in source_text):
        raise ValueError("Formula source_text contains unsafe control characters")
    if source_notation == "latex":
        for pattern in _UNSAFE_PATTERNS:
            match = pattern.search(source_text)
            if match:
                raise ValueError(f"Unsafe TeX command rejected: {match.group(0)}")


def probe_formula_capabilities(
    executable_finder: Callable[[str], str | None] | None = None,
) -> dict[str, dict[str, Any]]:
    """Detect allowlisted local tools without executing or installing anything."""

    finder = executable_finder or shutil.which
    resolved = {name: finder(name) for name in FORMULA_EXECUTABLE_ALLOWLIST}
    latex_engine = next(
        (name for name in ("latex", "pdflatex", "xelatex") if resolved[name]), None
    )
    converter = next(
        (name for name in ("dvisvgm", "pdftocairo") if resolved[name]), None
    )
    return {
        "native_omml": {
            "available": bool(resolved["pandoc"]),
            "execution_attempted": False,
            "executable": "pandoc",
            "resolved_path": (
                str(Path(resolved["pandoc"])) if resolved["pandoc"] else None
            ),
            "timeout_seconds": DEFAULT_FORMULA_TIMEOUT_SECONDS,
            "installation_attempted": False,
        },
        "latex_fallback": {
            "available": bool(latex_engine and converter),
            "execution_attempted": False,
            "engine": latex_engine,
            "engine_path": (
                str(Path(resolved[latex_engine])) if latex_engine else None
            ),
            "converter": converter,
            "converter_path": (
                str(Path(resolved[converter])) if converter else None
            ),
            "timeout_seconds": DEFAULT_FORMULA_TIMEOUT_SECONDS,
            "installation_attempted": False,
        },
        "plain_text": {
            "available": True,
            "execution_attempted": False,
            "executable": None,
            "timeout_seconds": 0,
            "installation_attempted": False,
        },
    }


detect_formula_capabilities = probe_formula_capabilities


def _fallback_methods(requested_method: str) -> tuple[str, ...]:
    if requested_method == "native_omml":
        return ("native_omml", "latex_fallback", "plain_text")
    if requested_method == "latex_fallback":
        return ("latex_fallback", "plain_text")
    return ("plain_text",)


def build_formula_manifest(
    source_text: str,
    *,
    source_notation: str = "latex",
    requested_method: str = "native_omml",
    formula_id: str | None = None,
    executable_finder: Callable[[str], str | None] | None = None,
) -> dict[str, Any]:
    """Build a deterministic, non-executing formula rendering manifest."""

    if requested_method not in REQUESTED_METHODS:
        raise ValueError(
            f"requested_method must be one of: {', '.join(REQUESTED_METHODS)}"
        )
    validate_formula_source(source_text, source_notation)
    source_hash = _canonical_hash(
        {"source_notation": source_notation, "source_text": source_text}
    )
    if formula_id is None:
        formula_id = f"FORMULA-{source_hash[:16]}"
    elif not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,79}", formula_id):
        raise ValueError("formula_id must be 1-80 safe identifier characters")

    capabilities = probe_formula_capabilities(executable_finder)
    fallback_plan = [
        {
            "order": order,
            "method": method,
            "available": bool(capabilities[method]["available"]),
            "requires_external_execution": method != "plain_text",
            "terminal_fallback": method == "plain_text",
        }
        for order, method in enumerate(_fallback_methods(requested_method), start=1)
    ]
    manifest: dict[str, Any] = {
        "formula_id": formula_id,
        "source_notation": source_notation,
        "source_text": source_text,
        "source_hash": source_hash,
        "requested_method": requested_method,
        "capability_detection": capabilities,
        "unsafe_tex_reject_list": list(UNSAFE_TEX_REJECT_LIST),
        "fallback_plan": fallback_plan,
        "execution_policy": {
            "execute_by_default": False,
            "shell_allowed": False,
            "installation_allowed": False,
            "pptx_injection_performed": False,
            "allowlisted_executables": list(FORMULA_EXECUTABLE_ALLOWLIST),
        },
    }
    manifest["manifest_hash"] = _canonical_hash(manifest)
    return manifest
