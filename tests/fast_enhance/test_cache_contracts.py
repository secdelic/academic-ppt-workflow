from __future__ import annotations

import hashlib
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.fast_enhance import (  # noqa: E402
    FastEnhanceError,
    _promote_fast_cache,
    _refreshed_pptx_asset_state,
    promote_full_validation_cache,
)


def _fingerprints() -> dict[str, str]:
    return {
        "brief_fingerprint": "0" * 64,
        "scientific_definition_fingerprint": "1" * 64,
        "extractor_fingerprint": "2" * 64,
        "config_fingerprint": "3" * 64,
        "code_fingerprint": "4" * 64,
    }


class FullValidationClinicalCacheGateTests(unittest.TestCase):
    def test_clinical_full_promotion_fails_before_cache_materialisation(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            repo = Path(raw)
            with self.assertRaisesRegex(FastEnhanceError, "not protected"):
                promote_full_validation_cache(
                    repo_root=repo,
                    project_identity="synthetic-private-project",
                    input_root=repo / "input",
                    manifest=[],
                    extracted={},
                    claims=[],
                    deck_ir={"slides": []},
                    storyboard=[],
                    style_profile=None,
                    figures=[],
                    qa_status={"status": "PASS"},
                    output_pptx=repo / "not-created.pptx",
                    brief={"project_name": "synthetic-private-project"},
                    clinical_privacy_mode=True,
                )
            self.assertFalse((repo / ".cache").exists())


class FastPromotionAssetRefreshTests(unittest.TestCase):
    def test_output_package_assets_are_rescanned_without_media_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            output = root / "updated.pptx"
            payload = b"synthetic-image-bytes"
            with zipfile.ZipFile(output, "w") as package:
                package.writestr("ppt/media/image1.png", payload)
            profile = {
                "slide_dimensions": {"width_emu": 1, "height_emu": 1},
                "slide_masters": [{"part": "ppt/slideMasters/slideMaster1.xml"}],
                "slide_layouts": [{"part": "ppt/slideLayouts/slideLayout1.xml"}],
                "common_page_roles": [{"role": "content", "layout_count": 1}],
                "footer_and_page_number_rules": {},
                "theme_fonts": {"major": {"latin": "Aptos"}},
            }
            with patch(
                "academic_ppt.fast_enhance.parse_reference_style",
                return_value=profile,
            ):
                masters, fonts, images = _refreshed_pptx_asset_state(output)
            self.assertEqual(masters["cache_status"], "REFRESHED_FROM_OUTPUT_PPTX")
            self.assertEqual(fonts["theme_fonts"], profile["theme_fonts"])
            self.assertEqual(images[0]["part"], "ppt/media/image1.png")
            self.assertEqual(images[0]["sha256"], hashlib.sha256(payload).hexdigest())
            self.assertNotIn(payload.decode(), str(images))

    def test_asset_scan_failure_invalidates_instead_of_reusing_old_state(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            output = Path(raw) / "updated.pptx"
            output.write_bytes(b"not-a-package")
            masters, fonts, images = _refreshed_pptx_asset_state(output)
            self.assertEqual(
                masters["cache_status"], "INVALIDATED_AFTER_FAST_UPDATE"
            )
            self.assertEqual(fonts["cache_status"], "INVALIDATED_AFTER_FAST_UPDATE")
            self.assertEqual(
                images["cache_status"], "INVALIDATED_AFTER_FAST_UPDATE"
            )

    def test_fast_promotion_uses_refreshed_assets_not_cached_assets(self) -> None:
        class CaptureCache:
            def __init__(self) -> None:
                self.state = None

            def commit_generation(self, state):
                self.state = state
                return {"current": "GEN-synthetic"}

        with tempfile.TemporaryDirectory() as raw:
            output = Path(raw) / "updated.pptx"
            output.write_bytes(b"synthetic-output")
            cache = CaptureCache()
            cached = {
                "master_layout_map": {"old": True},
                "font_map": {"old": True},
                "image_registry": [{"old": True}],
                "visual_specs": [],
            }
            refreshed = (
                {"cache_status": "REFRESHED_FROM_OUTPUT_PPTX"},
                {"cache_status": "REFRESHED_FROM_OUTPUT_PPTX"},
                [{"part": "ppt/media/image1.png", "sha256": "f" * 64}],
            )
            with patch(
                "academic_ppt.fast_enhance._refreshed_pptx_asset_state",
                return_value=refreshed,
            ):
                _promote_fast_cache(
                    cache=cache,
                    cached_state=cached,
                    current_source_manifest={
                        "schema_version": "academic-ppt-source-manifest/1",
                        "hash_algorithm": "sha256",
                        "sources": [],
                    },
                    current_registry=[],
                    parsed_objects={},
                    evidence_registry=[],
                    slide_specs=[],
                    qa_status={"status": "PASS"},
                    output_pptx=output,
                    contract_fingerprints=_fingerprints(),
                )
            self.assertEqual(cache.state["master_layout_map"], refreshed[0])
            self.assertEqual(cache.state["font_map"], refreshed[1])
            self.assertEqual(cache.state["image_registry"], refreshed[2])
            self.assertNotIn("old", str(cache.state["master_layout_map"]))


if __name__ == "__main__":
    unittest.main()
