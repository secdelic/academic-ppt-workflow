from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.formula_pipeline import (
    FORMULA_EXECUTABLE_ALLOWLIST,
    build_formula_manifest,
    probe_formula_capabilities,
)
from academic_ppt.scientific_visuals import (
    MECHANISM_EVIDENCE_STATUSES,
    TEMPLATE_IDS,
    build_visual_spec,
    probe_visual_capabilities,
    render_safe_svg,
)


class ScientificVisualTests(unittest.TestCase):
    def test_all_eleven_templates_are_declarative_and_hashable(self):
        self.assertEqual(len(TEMPLATE_IDS), 11)
        for template_id in TEMPLATE_IDS:
            spec = build_visual_spec(template_id)
            self.assertTrue(spec["nodes"], template_id)
            self.assertIn("edges", spec)
            self.assertEqual(
                spec["layout"]["coordinate_system"], "normalized_0_to_1"
            )
            self.assertTrue(spec["metadata"]["native_shape_compatible"])
            self.assertEqual(len(spec["content_hash"]), 64)

    def test_mechanism_evidence_status_defaults_and_exact_values(self):
        default = build_visual_spec("mechanism_hypothesis")
        self.assertEqual(
            default["metadata"]["evidence_status"], "hypothesis_only"
        )
        for status in MECHANISM_EVIDENCE_STATUSES:
            spec = build_visual_spec(
                "mechanism_hypothesis", evidence_status=status
            )
            self.assertEqual(spec["metadata"]["evidence_status"], status)
        with self.assertRaises(ValueError):
            build_visual_spec(
                "mechanism_hypothesis", evidence_status="mechanism_proven"
            )

    def test_safe_svg_escapes_text_and_has_no_external_references(self):
        spec = build_visual_spec(
            "cohort_flow",
            title='A < B & "quoted"',
            labels={"eligible": "<script>alert(1)</script> & records"},
        )
        svg = render_safe_svg(spec)
        self.assertIn("&lt;script&gt;", svg)
        self.assertNotIn("<script", svg.lower())
        self.assertNotIn("<image", svg.lower())
        self.assertNotIn("<foreignobject", svg.lower())
        self.assertNotIn("href=", svg.lower())
        self.assertNotIn("http://", svg.lower())
        self.assertNotIn("https://", svg.lower())
        self.assertNotIn("file://", svg.lower())
        self.assertNotIn("data:", svg.lower())

    def test_visual_hash_is_deterministic_and_content_sensitive(self):
        first = build_visual_spec(
            "timeline", labels={"baseline": "Enrollment"}
        )
        second = build_visual_spec(
            "timeline", labels={"baseline": "Enrollment"}
        )
        changed = build_visual_spec(
            "timeline", labels={"baseline": "Index date"}
        )
        self.assertEqual(first["content_hash"], second["content_hash"])
        self.assertNotEqual(first["content_hash"], changed["content_hash"])

    def test_visual_probe_uses_only_allowlisted_names_and_never_executes(self):
        requested: list[str] = []

        def fake_finder(name: str) -> str | None:
            requested.append(name)
            drive = chr(67) + chr(58)
            return f"{drive}/tools/{name}.exe"

        capabilities = probe_visual_capabilities(fake_finder)
        self.assertEqual(set(requested), {"dot", "mmdc"})
        self.assertTrue(capabilities["graphviz"]["available"])
        self.assertFalse(capabilities["mermaid"]["execution_attempted"])


class FormulaPipelineTests(unittest.TestCase):
    @staticmethod
    def _finder(name: str) -> str | None:
        drive = chr(67) + chr(58)
        available = {
            "pandoc": f"{drive}/tools/pandoc.exe",
            "pdflatex": f"{drive}/tools/pdflatex.exe",
            "dvisvgm": f"{drive}/tools/dvisvgm.exe",
        }
        return available.get(name)

    def test_formula_manifest_and_hashes_are_deterministic(self):
        first = build_formula_manifest(
            r"E = mc^2", executable_finder=self._finder
        )
        second = build_formula_manifest(
            r"E = mc^2", executable_finder=self._finder
        )
        self.assertEqual(first["formula_id"], second["formula_id"])
        self.assertEqual(first["source_hash"], second["source_hash"])
        self.assertEqual(first["manifest_hash"], second["manifest_hash"])
        self.assertEqual(
            [step["method"] for step in first["fallback_plan"]],
            ["native_omml", "latex_fallback", "plain_text"],
        )
        self.assertFalse(first["execution_policy"]["execute_by_default"])
        self.assertFalse(first["execution_policy"]["pptx_injection_performed"])

    def test_unsafe_tex_commands_are_rejected(self):
        unsafe = (
            r"\input{secrets.txt}",
            r"\includegraphics{patient.png}",
            r"\write18{cmd.exe}",
            r"\href{https://example.test}{x}",
            r"\begin{filecontents}{x}unsafe\end{filecontents}",
        )
        for formula in unsafe:
            with self.subTest(formula=formula):
                with self.assertRaises(ValueError):
                    build_formula_manifest(formula)

    def test_formula_probe_is_detection_only_and_allowlisted(self):
        requested: list[str] = []

        def fake_finder(name: str) -> str | None:
            requested.append(name)
            return None

        capabilities = probe_formula_capabilities(fake_finder)
        self.assertEqual(tuple(requested), FORMULA_EXECUTABLE_ALLOWLIST)
        self.assertTrue(capabilities["plain_text"]["available"])
        for capability in capabilities.values():
            self.assertFalse(capability["execution_attempted"])
            self.assertFalse(capability["installation_attempted"])


if __name__ == "__main__":
    unittest.main()
