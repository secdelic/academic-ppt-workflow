from __future__ import annotations

import json
import shlex
import shutil
import subprocess
from pathlib import Path
from typing import Any, Mapping

from .base import DEFAULT_BOUNDS, VisualArtifact, VisualBackend, source_map
from .native_pptxgenjs import NativePptxGenJSBackend


class GraphvizAdapter(VisualBackend):
    backend_id = "graphviz_adapter"

    def __init__(self, dot_executable: str | None = None) -> None:
        self.dot_executable = dot_executable or shutil.which("dot")

    def render(
        self, visual_ir: Mapping[str, Any], output_dir: Path
    ) -> VisualArtifact:
        output_dir.mkdir(parents=True, exist_ok=True)
        categories = [str(item) for item in visual_ir.get("categories", [])][:10]
        if len(categories) < 2 or not self.dot_executable:
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
                    "Graphviz unavailable or graph has fewer than two nodes; "
                    "native_pptxgenjs fallback used."
                ],
            )
        nodes = [
            {
                "object_id": f"{visual_ir['visual_id']}:node:{index + 1}",
                "kind": "node",
                "text": label,
                "text_editable_in_dot_source": True,
                "native_shape_translation_available": True,
            }
            for index, label in enumerate(categories)
        ]
        edges = [
            {
                "object_id": f"{visual_ir['visual_id']}:edge:{index + 1}",
                "kind": "directed_edge",
                "source": f"n{index}",
                "target": f"n{index + 1}",
                "native_connector_translation_available": True,
            }
            for index in range(len(categories) - 1)
        ]
        dot_lines = [
            "digraph G {",
            'graph [rankdir=LR, bgcolor="transparent", margin=0.05];',
            'node [shape=box, style="rounded,filled", fillcolor="#FFFFFF", '
            'color="#0F766E", fontname="Arial", fontsize=16];',
            'edge [color="#627D98", penwidth=1.5, arrowsize=0.8];',
        ]
        for index, label in enumerate(categories):
            escaped = label.replace("\\", "\\\\").replace('"', '\\"')
            dot_lines.append(f'n{index} [label="{escaped}"];')
        for index in range(len(categories) - 1):
            dot_lines.append(f"n{index} -> n{index + 1};")
        dot_lines.append("}")
        dot_path = output_dir / f"{visual_ir['visual_id']}.dot"
        svg_path = output_dir / f"{visual_ir['visual_id']}.graphviz.svg"
        dot_path.write_text("\n".join(dot_lines), encoding="utf-8")
        completed = subprocess.run(
            [self.dot_executable, "-Tsvg", str(dot_path), "-o", str(svg_path)],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
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
                warnings=[f"Graphviz failed: {completed.stderr.strip()}"],
            )
        plain_path = output_dir / f"{visual_ir['visual_id']}.graphviz.plain"
        plain = subprocess.run(
            [self.dot_executable, "-Tplain", str(dot_path), "-o", str(plain_path)],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )
        positioned_nodes: dict[str, dict[str, float]] = {}
        graph_width = graph_height = 1.0
        if plain.returncode == 0 and plain_path.is_file():
            for raw in plain_path.read_text(encoding="utf-8").splitlines():
                parts = shlex.split(raw)
                if not parts:
                    continue
                if parts[0] == "graph" and len(parts) >= 4:
                    graph_width = max(0.001, float(parts[2]))
                    graph_height = max(0.001, float(parts[3]))
                elif parts[0] == "node" and len(parts) >= 6:
                    positioned_nodes[parts[1]] = {
                        "x": float(parts[2]) / graph_width,
                        "y": 1.0 - float(parts[3]) / graph_height,
                        "w": float(parts[4]) / graph_width,
                        "h": float(parts[5]) / graph_height,
                    }
        for index, node in enumerate(nodes):
            node.update(
                positioned_nodes.get(
                    f"n{index}",
                    {
                        "x": (index + 0.5) / len(nodes),
                        "y": 0.5,
                        "w": 0.16,
                        "h": 0.18,
                    },
                )
            )
        manifest = [*edges, *nodes]
        (output_dir / f"{visual_ir['visual_id']}.graphviz.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return VisualArtifact(
            artifact_path=str(svg_path),
            artifact_type="graphviz_svg_with_native_manifest",
            editable_level="source_and_manifest_editable",
            object_manifest=manifest,
            render_backend=self.backend_id,
            source_map=source_map(visual_ir),
            bounding_box=dict(DEFAULT_BOUNDS),
            fallback_used=False,
            warnings=[],
        )


__all__ = ["GraphvizAdapter"]
