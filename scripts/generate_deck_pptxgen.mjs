import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");

const [specPath, outputPptx, previewRoot] = process.argv.slice(2);
if (!specPath || !outputPptx || !previewRoot) {
  throw new Error("Usage: generate_deck_pptxgen.mjs <spec.json> <output.pptx> <preview-root>");
}
const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow";
pptx.subject = "Source-bound academic presentation";
pptx.title = spec.brief.project_name;
pptx.company = "Local assisted workflow";
pptx.lang = spec.brief.language || "zh-CN";
pptx.theme = {
  headFontFace: spec.brief.language?.startsWith("zh") ? "Microsoft YaHei" : "Arial",
  bodyFontFace: spec.brief.language?.startsWith("zh") ? "Microsoft YaHei" : "Arial",
  lang: spec.brief.language || "zh-CN",
};

const C = { bg: "F7FAFC", ink: "102A43", muted: "627D98", primary: "0F766E", secondary: "2563EB", accent: "D97706", line: "D9E2EC", white: "FFFFFF" };
const font = spec.brief.language?.startsWith("zh") ? "Microsoft YaHei" : "Arial";

function notes(item) {
  const sourceLines = item.source_ids?.length ? item.source_ids.map((s) => `- ${s}`) : ["- INFORMATION_REQUIRED"];
  return `${item.speaker_note_summary || ""}\n\n[Sources]\n${sourceLines.join("\n")}`;
}
function addFooter(slide, item, index) {
  slide.addShape(pptx.ShapeType.line, { x: 0.75, y: 6.93, w: 11.83, h: 0, line: { color: C.line, width: 1 } });
  const sourceText = item.source_ids?.length ? `Sources: ${item.source_ids.join(", ")}` : "Source: INFORMATION_REQUIRED";
  slide.addText(sourceText, { x: 0.75, y: 6.98, w: 10.3, h: 0.28, fontFace: font, fontSize: 11, color: C.muted, margin: 0, breakLine: false });
  slide.addText(String(index + 1), { x: 11.8, y: 6.98, w: 0.75, h: 0.28, fontFace: font, fontSize: 11, color: C.muted, align: "right", margin: 0 });
}
function addBase(slide, item, index) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0.75, y: 0.56, w: 0.54, h: 0.07, fill: { color: C.primary }, line: { color: C.primary, transparency: 100 } });
  slide.addText(item.slide_title, { x: 0.75, y: 0.76, w: 11.83, h: 1.12, fontFace: font, fontSize: 35, bold: true, color: C.ink, margin: 0, breakLine: false });
}

for (const [index, item] of spec.slides.entries()) {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  if (index === 0) {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.25, h: 7.5, fill: { color: C.primary }, line: { color: C.primary } });
    slide.addText("ACADEMIC PRESENTATION", { x: 0.83, y: 0.9, w: 5.2, h: 0.35, fontFace: font, fontSize: 12, bold: true, color: C.primary, margin: 0 });
    slide.addText(item.slide_title.replaceAll("_", " "), { x: 0.83, y: 1.45, w: 10.8, h: 2.05, fontFace: font, fontSize: 50, bold: true, color: C.ink, margin: 0, breakLine: false });
    slide.addText(item.single_key_message, { x: 0.83, y: 3.9, w: 9.4, h: 1.15, fontFace: font, fontSize: 21, color: C.muted, margin: 0, fit: "shrink" });
    slide.addText(`${spec.brief.presentation_type} · ${spec.brief.duration_minutes} min`, { x: 0.83, y: 6.15, w: 7, h: 0.35, fontFace: font, fontSize: 13, color: C.muted, margin: 0 });
  } else {
    addBase(slide, item, index);
    if (item.chart_data) {
      const unit = item.chart_data.unit || "unit not provided";
      slide.addText(item.single_key_message, { x: 0.75, y: 2.05, w: 4.2, h: 1.8, fontFace: font, fontSize: 22, bold: true, color: C.ink, margin: 0.08, fit: "shrink" });
      slide.addText(`Editable chart · unit: ${unit} · sample size: not provided`, { x: 0.75, y: 4.45, w: 4.2, h: 0.7, fontFace: font, fontSize: 16, color: C.muted, margin: 0, fit: "shrink" });
      slide.addChart(pptx.ChartType.bar, [{ name: "Value", labels: item.chart_data.categories, values: item.chart_data.values }], {
        x: 5.45, y: 2.0, w: 7.05, h: 4.25,
        catAxisLabelFontFace: font, catAxisLabelFontSize: 16,
        valAxisLabelFontFace: font, valAxisLabelFontSize: 16,
        showLegend: false, showValue: true, chartColors: [C.secondary],
        showTitle: false, showCatName: false, showSerName: false,
        showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: unit === "percent" ? '0"%"' : "0.0",
        showValAxisTitle: true, valAxisTitle: unit || "Value", valAxisLabelFormatCode: unit === "percent" ? '0"%"' : "0.0", showCatAxisTitle: false,
        showValue: true, valGridLine: { color: C.line, width: 1 },
        border: { color: C.white, transparency: 100 },
      });
    } else if (item.visual_type === "source_map") {
      slide.addText(item.single_key_message, { x: 0.75, y: 2.05, w: 11.83, h: 0.5, fontFace: font, fontSize: 24, bold: true, color: C.ink, margin: 0 });
      for (const [sourceIndex, source] of (item.source_records || []).entries()) {
        const y = 3.0 + sourceIndex * 0.78;
        slide.addText(String(sourceIndex + 1).padStart(2, "0"), { x: 1.0, y, w: 0.55, h: 0.3, fontFace: font, fontSize: 16, bold: true, color: C.primary, margin: 0 });
        slide.addText(source.file_name, { x: 1.9, y, w: 5.2, h: 0.35, fontFace: font, fontSize: 18, bold: true, color: C.ink, margin: 0 });
        slide.addText(`${source.file_type.toUpperCase()} · ${source.source_id}`, { x: 7.4, y, w: 4.6, h: 0.35, fontFace: font, fontSize: 14, color: C.muted, align: "right", margin: 0 });
      }
    } else if (item.visual_type === "comparison") {
      slide.addText("Allowed", { x: 0.75, y: 2.15, w: 4.6, h: 0.45, fontFace: font, fontSize: 24, bold: true, color: C.primary, margin: 0 });
      slide.addText("Preserve the study design, uncertainty, and source wording.", { x: 0.75, y: 2.85, w: 4.8, h: 1.6, fontFace: font, fontSize: 20, color: C.ink, margin: 0, fit: "shrink" });
      slide.addShape(pptx.ShapeType.line, { x: 6.45, y: 2.05, w: 0, h: 3.7, line: { color: C.line, width: 2 } });
      slide.addText("Not allowed", { x: 7.15, y: 2.15, w: 4.3, h: 0.45, fontFace: font, fontSize: 24, bold: true, color: C.accent, margin: 0 });
      slide.addText("Do not upgrade association or computational inference into causal, mechanistic, or clinical proof.", { x: 7.15, y: 2.85, w: 4.8, h: 1.8, fontFace: font, fontSize: 20, color: C.ink, margin: 0, fit: "shrink" });
    } else {
      slide.addText(item.single_key_message, { x: 1.15, y: 2.55, w: 11.0, h: 2.1, fontFace: font, fontSize: 25, bold: true, color: C.ink, align: "center", valign: "mid", margin: 0.1, fit: "shrink" });
    }
    addFooter(slide, item, index);
  }
  if (typeof slide.addNotes === "function") slide.addNotes(notes(item));
}
fs.mkdirSync(path.dirname(outputPptx), { recursive: true });
fs.mkdirSync(path.join(previewRoot, "layouts"), { recursive: true });
fs.writeFileSync(path.join(previewRoot, "layouts", "README.txt"), "PptxGenJS backend relies on post-generation PowerPoint render QA.\n");
await pptx.writeFile({ fileName: outputPptx });
console.log(`PPTXGENJS_SLIDES=${spec.slides.length}`);
console.log(`PPTX=${outputPptx}`);
