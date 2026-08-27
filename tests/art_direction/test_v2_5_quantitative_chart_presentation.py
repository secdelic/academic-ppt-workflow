from pathlib import Path
import unittest


REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "scripts" / "v2_5" / "render_v2_5_deck.mjs"


def function_body(source: str, start: str, end: str) -> str:
    return source.split(f"function {start}", 1)[1].split(f"function {end}", 1)[0]


class V25QuantitativeChartPresentationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = RENDERER.read_text(encoding="utf-8")

    def test_faceted_volcano_has_data_derived_quantitative_axes(self):
        body = function_body(self.source, "renderFacetedVolcano", "renderEnrichment")
        self.assertIn("0 < FDR <= 1", body)
        self.assertIn("maxAbsEffect", body)
        self.assertIn("maxScore", body)
        self.assertIn("xTicks", body)
        self.assertIn("yTicks", body)
        self.assertIn("-log₁₀(FDR)", body)
        self.assertIn("x + 0.62", body)
        self.assertIn("tick === 0 ? ty - 0.30 : ty - 0.10", body)
        self.assertIn("log₂FC", body)
        self.assertNotIn("Math.min(18, score)", body)

    def test_faceted_volcano_retains_all_points_and_top_k_labels_only(self):
        body = function_body(self.source, "renderFacetedVolcano", "renderEnrichment")
        self.assertIn("subset.forEach", body)
        self.assertIn("slice(0, 2)", body)
        self.assertIn("rowName(visual, rows.indexOf(item))", body)

    def test_grouped_composition_localizes_display_without_reparsing_data(self):
        body = function_body(self.source, "renderNativeChart", "resolveSourceFigureTreatment")
        self.assertIn("isGroupedComposition", body)
        self.assertIn("displaySeries", body)
        self.assertIn("labels.map(zhLabel)", body)
        self.assertIn("chartWorkbookFields: data.workbookFields", body)
        self.assertIn("slide.addChart(data.pptxType, displaySeries", body)
        self.assertIn("nativeDataLabelCount", body)
        self.assertIn("nativeDataLabelCount <= 8", body)
        self.assertIn("showValue: showDirectValues", body)
        self.assertIn('dataLabelPolicy: isGroupedComposition && !showDirectValues ? "axis_legend_plus_annotation"', body)
        self.assertNotIn('showValue: isBar && data.categoryCount <= 10', body)

    def test_annotation_entity_is_localized_at_render_time(self):
        normalize = function_body(self.source, "normalizeAnnotation", "annotationNeedsExternalPanel")
        annotation = function_body(self.source, "annotationAudienceText", "localizeAudienceText")
        render = function_body(self.source, "renderSlideAnnotations", "renderVisual")
        self.assertIn("entity: item.entity", normalize)
        self.assertIn("zhLabel(entity)", annotation)
        self.assertIn("个百分点", annotation)
        self.assertIn("annotationAudienceText(annotation)", render)

    def test_time_window_positions_use_hour_not_event_order(self):
        body = function_body(self.source, "renderTimeWindow", "renderLove")
        self.assertIn("Number(event.hour) / 24", body)
        self.assertNotIn("index / events.length", body)


if __name__ == "__main__":
    unittest.main()
