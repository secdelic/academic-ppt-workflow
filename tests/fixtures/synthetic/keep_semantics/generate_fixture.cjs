"use strict";

const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const output = process.argv[2];
const changedChart = process.argv[3] === "changed-chart";
if (!output) throw new Error("Usage: generate_fixture.cjs <output.pptx>");

const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFElEQVR4nGP8z8DAwMDAxMDAwMAAAAoAAf8CB0kAAAAASUVORK5CYII=";
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow";
pptx.subject = "SYNTHETIC KEEP SEMANTIC REGRESSION";
pptx.title = "SYNTHETIC KEEP SEMANTIC REGRESSION";
pptx.company = "SYNTHETIC";
pptx.lang = "en-US";
pptx.theme = { headFontFace: "Arial", bodyFontFace: "Arial", lang: "en-US" };

function title(slide, value) {
  slide.addText(value, { x: 0.6, y: 0.3, w: 12.0, h: 0.45, fontFace: "Arial", fontSize: 24, bold: true, color: "17324D", margin: 0 });
}
function note(slide, id) {
  if (typeof slide.addNotes === "function") slide.addNotes(`[Sources]\nSYNTHETIC\n[Slide-ID]\n${id}`);
}

let slide = pptx.addSlide();
title(slide, "Text-only KEEP");
slide.addText("Visible synthetic text", { x: 0.8, y: 1.5, w: 4.8, h: 0.5, fontFace: "Arial", fontSize: 20, color: "334155", margin: 0 });
note(slide, "SYN-KEEP-TEXT");

slide = pptx.addSlide();
title(slide, "Table KEEP");
slide.addTable([["Item", "State", "Value"], ["Alpha", "Stable", "10"]], {
  x: 0.8, y: 1.4, w: 8.8, h: 1.5, border: { type: "solid", color: "D9E2EC", pt: 1 },
  fill: "F7FAFC", color: "17324D", fontFace: "Arial", fontSize: 16,
  margin: 0.05, rowH: 0.55, colW: [3.0, 3.0, 2.8],
});
note(slide, "SYN-KEEP-TABLE");

slide = pptx.addSlide();
title(slide, "Image KEEP");
slide.addImage({ data: png, x: 1.0, y: 1.4, w: 2.4, h: 2.4 });
slide.addText("Synthetic raster", { x: 3.8, y: 2.1, w: 3.0, h: 0.4, fontFace: "Arial", fontSize: 18, margin: 0 });
note(slide, "SYN-KEEP-IMAGE");

slide = pptx.addSlide();
title(slide, "Diagram KEEP");
slide.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.6, w: 2.2, h: 1.0, fill: { color: "DCEAF7" }, line: { color: "5A7D9A", pt: 1 } });
slide.addText("Node A", { x: 1.2, y: 1.9, w: 1.8, h: 0.3, fontFace: "Arial", fontSize: 18, align: "center", margin: 0 });
slide.addShape(pptx.ShapeType.chevron, { x: 3.6, y: 1.85, w: 1.2, h: 0.5, fill: { color: "5A7D9A" }, line: { color: "5A7D9A" } });
slide.addShape(pptx.ShapeType.rect, { x: 5.2, y: 1.6, w: 2.2, h: 1.0, fill: { color: "E6F2EC" }, line: { color: "4A8062", pt: 1 } });
slide.addText("Node B", { x: 5.4, y: 1.9, w: 1.8, h: 0.3, fontFace: "Arial", fontSize: 18, align: "center", margin: 0 });
note(slide, "SYN-KEEP-DIAGRAM");

slide = pptx.addSlide();
title(slide, "Mixed KEEP");
slide.addText("Mixed content", { x: 0.9, y: 1.25, w: 3.0, h: 0.4, fontFace: "Arial", fontSize: 18, margin: 0 });
slide.addImage({ data: png, x: 0.9, y: 2.0, w: 1.5, h: 1.5 });
slide.addShape(pptx.ShapeType.roundRect, { x: 3.0, y: 2.0, w: 2.5, h: 1.0, fill: { color: "F7E8D0" }, line: { color: "B7791F", pt: 1 } });
slide.addText("Editable card", { x: 3.2, y: 2.3, w: 2.1, h: 0.3, fontFace: "Arial", fontSize: 17, align: "center", margin: 0 });
note(slide, "SYN-KEEP-MIXED");

slide = pptx.addSlide();
title(slide, "Chart change detector");
slide.addChart(pptx.ChartType.bar, [{ name: "Synthetic series", labels: ["A", "B", "C"], values: changedChart ? [1, 2, 9] : [1, 2, 3] }], {
  x: 0.9, y: 1.2, w: 7.2, h: 4.4, catAxisLabelFontFace: "Arial", valAxisLabelFontFace: "Arial",
  showLegend: false, showTitle: false, showValue: true, chartColors: ["2A7F9E"],
});
note(slide, "SYN-CHANGED-CHART");

fs.mkdirSync(path.dirname(output), { recursive: true });
pptx.writeFile({ fileName: output }).catch((error) => {
  process.stderr.write(String(error && error.stack ? error.stack : error) + "\n");
  process.exitCode = 1;
});
