from pathlib import Path
import re
import unittest


REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "scripts" / "v2_5" / "render_v2_5_deck.mjs"


def function_body(source: str, start: str, end: str) -> str:
    return source.split(f"function {start}", 1)[1].split(f"function {end}", 1)[0]


class V25SemanticGraphRenderingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = RENDERER.read_text(encoding="utf-8")

    def test_process_branch_uses_explicit_edges_and_topological_depth(self):
        body = function_body(self.source, "renderBranch", "renderTimeWindow")
        self.assertIn("visual.payload.edges || visual.payload.branches", body)
        self.assertIn("implicit sequencing is prohibited", body)
        self.assertIn("depth[target] = Math.max(depth[target], depth[source] + 1)", body)
        self.assertIn("must be acyclic", body)
        self.assertNotRegex(body, r"if \(count <= 6\).*index \*", "node order must not masquerade as a linear process")

    def test_dag_routes_cross_layer_edges_without_dropping_them(self):
        body = function_body(self.source, "renderDAG", "renderRiskMatrix")
        self.assertIn("edges.forEach(routeDAGEdge)", body)
        self.assertIn("obstructed = nodes.some", body)
        self.assertIn("confounder-to-", body)
        self.assertRegex(body, re.compile(r"edgeName}:a[\s\S]+edgeName}:e"))
        self.assertIn("DAG edge references missing node", body)

    def test_assessment_timeline_has_two_dimensional_fail_closed_packing(self):
        body = function_body(self.source, "renderAssessmentTimeline", "renderVariables")
        self.assertIn('side: "top"', body)
        self.assertIn('side: "bottom"', body)
        self.assertIn("findTimelineSlot", body)
        self.assertIn("intervalAvailable", body)
        self.assertIn("estimatedLabelHeight", body)
        self.assertIn("2D packing budget exceeded", body)
        self.assertNotIn("lastRight: -Infinity", body)

    def test_cover_title_wrap_is_budgeted_balanced_and_token_safe(self):
        protected = function_body(self.source, "titleProtectedRanges", "wrapTitleText")
        wrapped = function_body(self.source, "wrapTitleText", "truncateDisplayText")
        cover = function_body(self.source, "renderCover", "renderHeroQuestion")
        self.assertIn("titleProtectedRanges(value)", wrapped)
        self.assertIn("candidate.overflow === 0", wrapped)
        self.assertIn("a.balance - b.balance", wrapped)
        self.assertLess(wrapped.index("b.punctuationBreak"), wrapped.index("a.balance - b.balance"))
        self.assertIn("test(right)", wrapped)
        self.assertIn("小时|分钟|天|周|月|年", protected)
        self.assertIn("TITLE_GLOSSARY_TERMS", protected)
        self.assertIn("Object.values(ZH_LABELS)", self.source)
        self.assertIn("(?:\\.\\d+)+", protected)
        self.assertIn("个百分点", protected)
        self.assertIn("OR|HR|RR|CI|FDR|NES|SMD", protected)
        self.assertIn("Script=Han", self.source)
        self.assertIn("Math.min(6, characters.length)", self.source)
        self.assertIn('coverTitle.split("\\n")', cover)
        self.assertIn("coverLines.length > 2", cover)
        self.assertIn("wrap: false", cover)
        self.assertRegex(cover, r"coverFontSize[^;]+38")
        self.assertIn("margin: [0.05, 0, 0.05, 0]", cover)

    def test_regular_titles_use_top_aligned_cross_renderer_safe_box(self):
        title_body = function_body(self.source, "title", "footer")
        self.assertIn('valign: "top"', title_body)
        self.assertRegex(self.source, r'title:\s*Object\.freeze\(\{ x: 0\.74, y: 0\.46, w: 11\.85, h: 0\.97, topGutter: 0 \}\)')
        self.assertIn("zone.y + zone.h > SAFE_ZONES.content.y", title_body)
        self.assertIn("margin: [zone.topGutter, 0, 0, 0]", title_body)
        self.assertIn("wrap: false", title_body)
        self.assertRegex(title_body, r"twoLines \? 30 : TYPO\.title")

    def test_footer_and_page_number_are_explicit_slide_level_objects(self):
        footer_body = function_body(self.source, "footer", "speakerNotes")
        master_block = self.source.split("for (const family of FAMILIES)", 1)[1].split("function addText", 1)[0]
        self.assertIn("slideNumber: undefined", master_block)
        self.assertNotIn("slideNumber: family", master_block)
        for object_name in ("footer-band", "footer-label", "page-number"):
            self.assertIn(object_name, footer_body)
        self.assertNotIn("footer-prefix", footer_body)
        self.assertIn('addText(slide, `来源： ${text}`', footer_body)
        self.assertGreaterEqual(footer_body.count("11,"), 2)
        self.assertIn("wrap: false", footer_body)
        self.assertIn("footer(slide, spec, slideIndex + 1)", self.source)
        self.assertLess(self.source.index("annotationEvidence[spec.slide_id] = renderSlideAnnotations"), self.source.index("footer(slide, spec, slideIndex + 1)"))

    def test_independent_network_edges_end_outside_receiver_nodes(self):
        helper = function_body(self.source, "clippedEllipseDirectedEdge", "renderNetwork")
        network = function_body(self.source, "renderNetwork", "renderLadder")
        self.assertIn("boundaryDenominator", helper)
        self.assertIn("target.x - dx * boundaryFraction - ux * gap", helper)
        self.assertIn("source.x + dx * boundaryFraction + ux * gap", helper)
        self.assertIn("clippedEllipseDirectedEdge(a, b)", network)
        self.assertIn('endArrowType: "triangle"', network)
        self.assertIn('dash: predicted ? "dash" : undefined', network)
        self.assertIn("visual.sequential === true", network)
        self.assertNotIn("line(slide, a.x, a.y, b.x - a.x, b.y - a.y", network)


if __name__ == "__main__":
    unittest.main()
