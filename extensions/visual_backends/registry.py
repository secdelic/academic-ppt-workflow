from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from .base import VisualArtifact
from .graphviz_adapter import GraphvizAdapter
from .latex_omml_adapter import LatexOmmlAdapter
from .mermaid_adapter import MermaidAdapter
from .native_pptxgenjs import NativePptxGenJSBackend
from .svg_drawingml_adapter import SvgDrawingMLAdapter


def render_visual(
    visual_ir: Mapping[str, Any],
    output_dir: Path,
    *,
    backend: str | None = None,
) -> VisualArtifact:
    selected = backend or str(
        visual_ir.get("preferred_backend", "native_pptxgenjs")
    )
    adapters = {
        "native_pptxgenjs": NativePptxGenJSBackend(),
        "svg_drawingml_adapter": SvgDrawingMLAdapter(),
        "graphviz_adapter": GraphvizAdapter(),
        "mermaid_adapter": MermaidAdapter(),
        "latex_omml_adapter": LatexOmmlAdapter(),
    }
    adapter = adapters.get(selected)
    if adapter is None:
        adapter = adapters["native_pptxgenjs"]
    return adapter.render(visual_ir, output_dir)


__all__ = ["render_visual"]
