from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any, Mapping

from .base import DEFAULT_BOUNDS, VisualArtifact, VisualBackend, source_map


class SvgDrawingMLAdapter(VisualBackend):
    """Clean-room vector adapter.

    It emits a deterministic SVG plus an explicit object manifest. The v2.2
    benchmark may translate the manifest to native PptxGenJS shapes. No
    upstream converter code, server, provider, or install script is used.
    """

    backend_id = "svg_drawingml_adapter"

    def render(
        self, visual_ir: Mapping[str, Any], output_dir: Path
    ) -> VisualArtifact:
        output_dir.mkdir(parents=True, exist_ok=True)
        categories = list(visual_ir.get("categories", []))[:8]
        objects: list[dict[str, Any]] = []
        y_step = 600 / max(1, len(categories))
        pieces = [
            '<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="675" '
            'viewBox="0 0 1200 675">',
            '<rect width="1200" height="675" fill="#F7FAFC"/>',
        ]
        for index, label in enumerate(categories):
            y = 35 + index * y_step
            width = 760 - index * 18
            pieces.extend(
                [
                    f'<rect x="210" y="{y:.1f}" width="{width:.1f}" '
                    'height="52" rx="10" fill="#FFFFFF" '
                    'stroke="#0F766E" stroke-width="2"/>',
                    f'<text x="590" y="{y + 33:.1f}" text-anchor="middle" '
                    'font-family="Arial, sans-serif" font-size="20" '
                    f'fill="#102A43">{html.escape(str(label))}</text>',
                ]
            )
            objects.append(
                {
                    "object_id": f"{visual_ir['visual_id']}:node:{index + 1}",
                    "kind": "round_rect_with_text",
                    "text": str(label),
                    "drawingml_translatable": True,
                    "text_editable_after_manifest_translation": True,
                }
            )
        pieces.append("</svg>")
        svg_path = output_dir / f"{visual_ir['visual_id']}.svg"
        svg_path.write_text("\n".join(pieces), encoding="utf-8")
        manifest_path = output_dir / f"{visual_ir['visual_id']}.drawingml.json"
        manifest_path.write_text(
            json.dumps(objects, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return VisualArtifact(
            artifact_path=str(svg_path),
            artifact_type="svg_with_drawingml_manifest",
            editable_level="manifest_translatable",
            object_manifest=objects,
            render_backend=self.backend_id,
            source_map=source_map(visual_ir),
            bounding_box=dict(DEFAULT_BOUNDS),
            fallback_used=False,
            warnings=[
                "SVG insertion alone is vector but not natively text-editable; "
                "production use requires the manifest-to-native-shape path."
            ],
        )


__all__ = ["SvgDrawingMLAdapter"]
