from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any

from .deck_ir import validate_deck_ir
from .utils import sha256_file


class TemplateFillError(RuntimeError):
    pass


def _load_and_validate_deck_ir(path: Path) -> dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            value = json.load(handle)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise TemplateFillError(f"Deck IR is not valid JSON: {path}") from exc
    if not isinstance(value, dict):
        raise TemplateFillError("Deck IR must be a JSON object")
    try:
        validate_deck_ir(value)
    except (TypeError, ValueError) as exc:
        raise TemplateFillError(f"Deck IR contract validation failed: {exc}") from exc
    if not value["slides"]:
        raise TemplateFillError("Deck IR contains no slides")
    return value


def _validate_template_paths(
    template_path: Path,
    deck_ir_path: Path,
    output_pptx: Path,
    script_path: Path,
) -> tuple[Path, Path, Path, Path]:
    template = template_path.resolve()
    deck_ir = deck_ir_path.resolve()
    output = output_pptx.resolve()
    script = script_path.resolve()
    if template.suffix.lower() not in {".pptx", ".potx"}:
        raise TemplateFillError("Native template fill accepts only .pptx or .potx")
    for path, label, suffix in [
        (template, "template", None),
        (deck_ir, "Deck IR", ".json"),
        (script, "PowerShell template-fill script", ".ps1"),
    ]:
        if not path.is_file():
            raise TemplateFillError(f"Missing {label}: {path}")
        if suffix and path.suffix.lower() != suffix:
            raise TemplateFillError(f"{label} must use the {suffix} extension")
    if output.suffix.lower() != ".pptx":
        raise TemplateFillError("Template-fill output must be a .pptx file")
    if output.exists():
        raise TemplateFillError(f"Refusing to overwrite template-fill output: {output}")
    if output in {template, deck_ir, script}:
        raise TemplateFillError("Output path must be distinct from every input")
    return template, deck_ir, output, script


def run_native_template_fill(
    *,
    template_path: Path,
    deck_ir_path: Path,
    output_pptx: Path,
    script_path: Path,
    timeout_seconds: int = 120,
) -> str:
    if timeout_seconds <= 0:
        raise TemplateFillError("timeout_seconds must be positive")
    template_path, deck_ir_path, output_pptx, script_path = (
        _validate_template_paths(
            template_path, deck_ir_path, output_pptx, script_path
        )
    )
    _load_and_validate_deck_ir(deck_ir_path)
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    input_hashes = {
        template_path: sha256_file(template_path),
        deck_ir_path: sha256_file(deck_ir_path),
        script_path: sha256_file(script_path),
    }
    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-Template",
        str(template_path),
        "-DeckIr",
        str(deck_ir_path),
        "-OutputPptx",
        str(output_pptx),
    ]
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        changed_inputs = [
            str(path)
            for path, digest in input_hashes.items()
            if not path.is_file() or sha256_file(path) != digest
        ]
        if changed_inputs:
            raise TemplateFillError(
                "Input changed during timed-out native fill: "
                + ", ".join(changed_inputs)
            ) from exc
        raise TemplateFillError(
            f"PowerPoint native template fill timed out after {timeout_seconds}s"
        ) from exc
    changed_inputs = [
        str(path)
        for path, digest in input_hashes.items()
        if not path.is_file() or sha256_file(path) != digest
    ]
    if changed_inputs:
        raise TemplateFillError(
            "Input changed during PowerPoint native fill: "
            + ", ".join(changed_inputs)
        )
    log = (completed.stdout + "\n" + completed.stderr).strip()
    if completed.returncode != 0:
        raise TemplateFillError(
            f"PowerPoint native template fill failed ({completed.returncode}): {log}"
        )
    if not output_pptx.is_file() or output_pptx.stat().st_size == 0:
        raise TemplateFillError("Native template fill produced no PPTX")
    return log
