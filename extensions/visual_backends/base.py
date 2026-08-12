from __future__ import annotations

import hashlib
import json
from abc import ABC, abstractmethod
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Mapping


@dataclass(frozen=True)
class VisualArtifact:
    artifact_path: str
    artifact_type: str
    editable_level: str
    object_manifest: list[dict[str, Any]]
    render_backend: str
    source_map: list[dict[str, str]]
    bounding_box: dict[str, float]
    fallback_used: bool
    warnings: list[str]

    def to_dict(self) -> dict[str, Any]:
        value = asdict(self)
        payload = json.dumps(value, sort_keys=True, ensure_ascii=False)
        value["artifact_hash"] = hashlib.sha256(payload.encode("utf-8")).hexdigest()
        return value


class VisualBackend(ABC):
    backend_id: str

    @abstractmethod
    def render(
        self,
        visual_ir: Mapping[str, Any],
        output_dir: Path,
    ) -> VisualArtifact:
        """Render one validated Visual IR object without altering source files."""


def source_map(visual_ir: Mapping[str, Any]) -> list[dict[str, str]]:
    return [
        {
            "visual_id": str(visual_ir["visual_id"]),
            "source_id": str(source_id),
        }
        for source_id in visual_ir.get("source_ids", [])
    ]


DEFAULT_BOUNDS = {"x": 0.0, "y": 0.0, "w": 1.0, "h": 1.0}


__all__ = [
    "DEFAULT_BOUNDS",
    "VisualArtifact",
    "VisualBackend",
    "source_map",
]
