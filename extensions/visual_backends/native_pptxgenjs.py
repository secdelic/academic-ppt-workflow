from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Mapping

from .base import DEFAULT_BOUNDS, VisualArtifact, VisualBackend, source_map


class NativePptxGenJSBackend(VisualBackend):
    backend_id = "native_pptxgenjs"

    def render(
        self, visual_ir: Mapping[str, Any], output_dir: Path
    ) -> VisualArtifact:
        output_dir.mkdir(parents=True, exist_ok=True)
        path = output_dir / f"{visual_ir['visual_id']}.native.json"
        objects = [
            {
                "object_id": f"{visual_ir['visual_id']}:native-root",
                "kind": str(visual_ir["visual_type"]),
                "editable": True,
                "text_editable": True,
            }
        ]
        path.write_text(
            json.dumps(
                {
                    "backend": self.backend_id,
                    "visual_ir": dict(visual_ir),
                    "object_manifest": objects,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return VisualArtifact(
            artifact_path=str(path),
            artifact_type="native_shape_manifest",
            editable_level="fully_editable",
            object_manifest=objects,
            render_backend=self.backend_id,
            source_map=source_map(visual_ir),
            bounding_box=dict(DEFAULT_BOUNDS),
            fallback_used=False,
            warnings=[],
        )


__all__ = ["NativePptxGenJSBackend"]
