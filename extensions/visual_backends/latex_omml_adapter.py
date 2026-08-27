from __future__ import annotations

import html
import json
import re
from pathlib import Path
from typing import Any, Mapping

from .base import DEFAULT_BOUNDS, VisualArtifact, VisualBackend, source_map


SAFE_LATEX = re.compile(r"^[A-Za-z0-9_{}^+\-*/=().,\s\\]+$")


class LatexOmmlAdapter(VisualBackend):
    backend_id = "latex_omml_adapter"

    def render(
        self, visual_ir: Mapping[str, Any], output_dir: Path
    ) -> VisualArtifact:
        output_dir.mkdir(parents=True, exist_ok=True)
        expression = str(
            visual_ir.get("annotation_rule", {}).get(
                "latex", visual_ir.get("estimand") or "INFORMATION_REQUIRED"
            )
        )
        warnings = []
        fallback = False
        if not SAFE_LATEX.fullmatch(expression):
            expression = re.sub(r"[^A-Za-z0-9_{}^+\-*/=().,\s\\]", "", expression)
            warnings.append("Unsafe LaTeX characters were removed before OMML fallback.")
            fallback = True
        tex_path = output_dir / f"{visual_ir['visual_id']}.tex"
        tex_path.write_text(expression, encoding="utf-8")
        omml_path = output_dir / f"{visual_ir['visual_id']}.omml.xml"
        omml_path.write_text(
            '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/'
            'officeDocument/2006/math"><m:oMath><m:r><m:t>'
            + html.escape(expression)
            + "</m:t></m:r></m:oMath></m:oMathPara>",
            encoding="utf-8",
        )
        objects = [
            {
                "object_id": f"{visual_ir['visual_id']}:equation",
                "kind": "omml_equation",
                "text_editable_after_ooxml_injection": True,
                "plain_text_fallback": expression,
            }
        ]
        (output_dir / f"{visual_ir['visual_id']}.formula.json").write_text(
            json.dumps(objects, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return VisualArtifact(
            artifact_path=str(omml_path),
            artifact_type="omml_with_editable_text_fallback",
            editable_level="native_after_injection",
            object_manifest=objects,
            render_backend=self.backend_id,
            source_map=source_map(visual_ir),
            bounding_box=dict(DEFAULT_BOUNDS),
            fallback_used=fallback,
            warnings=warnings,
        )


__all__ = ["LatexOmmlAdapter", "SAFE_LATEX"]
