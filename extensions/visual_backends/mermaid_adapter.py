from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from typing import Any, Mapping

from .base import DEFAULT_BOUNDS, VisualArtifact, VisualBackend, source_map
from .native_pptxgenjs import NativePptxGenJSBackend


class MermaidAdapter(VisualBackend):
    backend_id = "mermaid_adapter"

    def __init__(self, executable: str | None = None) -> None:
        self.executable = executable or shutil.which("mmdc")

    def render(
        self, visual_ir: Mapping[str, Any], output_dir: Path
    ) -> VisualArtifact:
        output_dir.mkdir(parents=True, exist_ok=True)
        categories = [str(item) for item in visual_ir.get("categories", [])][:8]
        source = ["flowchart LR"]
        manifest = []
        for index, label in enumerate(categories):
            safe = label.replace('"', "'").replace("[", "(").replace("]", ")")
            source.append(f'  N{index}["{safe}"]')
            manifest.append(
                {
                    "object_id": f"{visual_ir['visual_id']}:node:{index + 1}",
                    "kind": "node",
                    "text": label,
                    "text_editable_in_mermaid_source": True,
                }
            )
            if index:
                source.append(f"  N{index - 1} --> N{index}")
        source_path = output_dir / f"{visual_ir['visual_id']}.mmd"
        source_path.write_text("\n".join(source), encoding="utf-8")
        if not self.executable:
            native = NativePptxGenJSBackend().render(visual_ir, output_dir)
            return VisualArtifact(
                artifact_path=native.artifact_path,
                artifact_type=native.artifact_type,
                editable_level=native.editable_level,
                object_manifest=native.object_manifest,
                render_backend=self.backend_id,
                source_map=native.source_map,
                bounding_box=native.bounding_box,
                fallback_used=True,
                warnings=[
                    "Mermaid CLI is not installed; source was retained and "
                    "native_pptxgenjs fallback was used."
                ],
            )
        svg_path = output_dir / f"{visual_ir['visual_id']}.mermaid.svg"
        completed = subprocess.run(
            [
                self.executable,
                "-i",
                str(source_path),
                "-o",
                str(svg_path),
                "--backgroundColor",
                "transparent",
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
        )
        if completed.returncode != 0 or not svg_path.is_file():
            native = NativePptxGenJSBackend().render(visual_ir, output_dir)
            return VisualArtifact(
                artifact_path=native.artifact_path,
                artifact_type=native.artifact_type,
                editable_level=native.editable_level,
                object_manifest=native.object_manifest,
                render_backend=self.backend_id,
                source_map=native.source_map,
                bounding_box=native.bounding_box,
                fallback_used=True,
                warnings=[f"Mermaid render failed: {completed.stderr.strip()}"],
            )
        return VisualArtifact(
            artifact_path=str(svg_path),
            artifact_type="mermaid_svg",
            editable_level="source_editable",
            object_manifest=manifest,
            render_backend=self.backend_id,
            source_map=source_map(visual_ir),
            bounding_box=dict(DEFAULT_BOUNDS),
            fallback_used=False,
            warnings=[],
        )


__all__ = ["MermaidAdapter"]
