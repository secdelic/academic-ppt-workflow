import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);
const artifactPath = require.resolve("@oai/artifact-tool");
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactPath).href);

const [specPath, outputPptx, previewRoot] = process.argv.slice(2);
if (!specPath || !outputPptx || !previewRoot) {
  throw new Error("Usage: generate_deck_artifact.mjs <spec.json> <output.pptx> <preview-root>");
}
const spec = JSON.parse(await fs.readFile(specPath, "utf8"));
const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });

const C = {
  bg: "#F7FAFC",
  surface: "#FFFFFF",
  ink: "#102A43",
  muted: "#627D98",
  primary: "#0F766E",
  secondary: "#2563EB",
  accent: "#D97706",
  line: "#D9E2EC",
};
const font = spec.brief.language?.startsWith("zh") ? "Microsoft YaHei" : "Arial";

function addText(slide, text, position, style = {}, name = "text") {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = String(text ?? "");
  shape.text.style = { fontFamily: font, color: C.ink, ...style };
  return shape;
}

function addFooter(slide, item, index) {
  slide.shapes.add({
    geometry: "line",
    position: { left: 72, top: 666, width: 1136, height: 0 },
    fill: "none",
    line: { style: "solid", fill: C.line, width: 1 },
  });
  const sourceText = item.source_ids?.length
    ? `Sources: ${item.source_ids.join(", ")}`
    : "Source: INFORMATION_REQUIRED";
  addText(slide, sourceText, { left: 72, top: 672, width: 980, height: 24 }, { fontSize: 16, color: C.muted }, `sources-${index}`);
  addText(slide, String(index + 1), { left: 1130, top: 672, width: 78, height: 24 }, { fontSize: 16, color: C.muted, alignment: "right" }, `page-${index}`);
}

function addNotes(slide, item) {
  const sources = item.source_ids?.length
    ? item.source_ids.map((source) => `- ${source}`).join("\n")
    : "- INFORMATION_REQUIRED";
  slide.speakerNotes.textFrame.setText(
    `${item.speaker_note_summary || ""}\n\n[Sources]\n${sources}`,
  );
  slide.speakerNotes.setVisible(true);
}

function addCover(slide, item) {
  slide.background.fill = C.bg;
  slide.shapes.add({
    geometry: "rect",
    position: { left: 0, top: 0, width: 24, height: 720 },
    fill: C.primary,
    line: { style: "solid", fill: C.primary, width: 0 },
  });
  addText(slide, "ACADEMIC PRESENTATION", { left: 80, top: 86, width: 500, height: 32 }, { fontSize: 16, bold: true, color: C.primary }, "cover-eyebrow");
  addText(slide, item.slide_title.replaceAll("_", " "), { left: 80, top: 145, width: 1040, height: 190 }, { fontSize: 68, bold: true, color: C.ink }, "cover-title");
  addText(slide, item.single_key_message, { left: 80, top: 372, width: 900, height: 110 }, { fontSize: 26, color: C.muted }, "cover-objective");
  addText(slide, `${spec.brief.presentation_type} · ${spec.brief.duration_minutes} min`, { left: 80, top: 588, width: 700, height: 34 }, { fontSize: 18, color: C.muted }, "cover-meta");
}

function addBase(slide, item, index) {
  slide.background.fill = C.bg;
  slide.shapes.add({
    geometry: "rect",
    position: { left: 72, top: 54, width: 52, height: 7 },
    fill: C.primary,
    line: { style: "solid", fill: C.primary, width: 0 },
  });
  addText(slide, item.slide_title, { left: 72, top: 78, width: 1136, height: 104 }, { fontSize: 47, bold: true, color: C.ink }, `title-${index}`);
}

function addContent(slide, item, index) {
  addBase(slide, item, index);
  if (item.chart_data) {
    addText(slide, item.single_key_message, { left: 72, top: 208, width: 420, height: 190 }, { fontSize: 27, bold: true, color: C.ink }, `key-message-${index}`);
    const unit = item.chart_data.unit || "unit not provided";
    addText(slide, `Editable chart · unit: ${unit} · sample size: not provided`, { left: 72, top: 430, width: 430, height: 70 }, { fontSize: 21, color: C.muted }, `chart-note-${index}`);
    slide.charts.add("bar", {
      position: { left: 540, top: 206, width: 668, height: 408 },
      categories: item.chart_data.categories,
      series: [{ name: "Value", values: item.chart_data.values, fill: C.secondary, valuesFormatCode: unit === "percent" ? '0"%"' : "0.0" }],
      barOptions: { direction: "column", grouping: "clustered", gapWidth: 70 },
      hasLegend: false,
      chartFill: C.surface,
      plotAreaFill: C.surface,
      xAxis: { textStyle: { fontSize: 21, fill: C.muted }, line: { style: "solid", fill: C.line, width: 1 } },
      yAxis: { title: unit || "Value", numberFormatCode: unit === "percent" ? '0"%"' : "0.0", textStyle: { fontSize: 19, fill: C.muted }, majorGridlines: { style: "solid", fill: C.line, width: 1 } },
      dataLabels: { showValue: true, position: "outEnd", textStyle: { fontSize: 21, fill: C.ink, bold: true } },
    });
  } else if (item.visual_type === "source_map") {
    addText(slide, item.single_key_message, { left: 72, top: 202, width: 1136, height: 50 }, { fontSize: 29, bold: true, color: C.ink }, `key-message-${index}`);
    for (const [sourceIndex, source] of (item.source_records || []).entries()) {
      const top = 300 + sourceIndex * 76;
      addText(slide, String(sourceIndex + 1).padStart(2, "0"), { left: 96, top, width: 62, height: 34 }, { fontSize: 21, bold: true, color: C.primary }, `source-index-${index}-${sourceIndex}`);
      addText(slide, source.file_name, { left: 190, top, width: 520, height: 34 }, { fontSize: 24, bold: true, color: C.ink }, `source-name-${index}-${sourceIndex}`);
      addText(slide, `${source.file_type.toUpperCase()} · ${source.source_id}`, { left: 740, top: top + 2, width: 420, height: 34 }, { fontSize: 18, color: C.muted, alignment: "right" }, `source-id-${index}-${sourceIndex}`);
    }
  } else if (item.visual_type === "comparison") {
    addText(slide, "Allowed", { left: 72, top: 220, width: 420, height: 42 }, { fontSize: 32, bold: true, color: C.primary }, `allowed-${index}`);
    addText(slide, "Preserve the study design, uncertainty, and source wording.", { left: 72, top: 286, width: 460, height: 150 }, { fontSize: 27, color: C.ink }, `allowed-body-${index}`);
    slide.shapes.add({ geometry: "line", position: { left: 625, top: 208, width: 0, height: 360 }, fill: "none", line: { style: "solid", fill: C.line, width: 2 } });
    addText(slide, "Not allowed", { left: 690, top: 220, width: 420, height: 42 }, { fontSize: 32, bold: true, color: C.accent }, `not-allowed-${index}`);
    addText(slide, "Do not upgrade association or computational inference into causal, mechanistic, or clinical proof.", { left: 690, top: 286, width: 470, height: 190 }, { fontSize: 27, color: C.ink }, `not-allowed-body-${index}`);
  } else {
    addText(slide, item.single_key_message, { left: 104, top: 238, width: 1040, height: 230 }, { fontSize: 34, bold: true, color: C.ink, alignment: "center" }, `key-message-${index}`);
  }
  addFooter(slide, item, index);
}

for (const [index, item] of spec.slides.entries()) {
  const slide = presentation.slides.add();
  if (index === 0) addCover(slide, item);
  else addContent(slide, item, index);
  addNotes(slide, item);
}

await fs.mkdir(path.dirname(outputPptx), { recursive: true });
await fs.mkdir(path.join(previewRoot, "slides"), { recursive: true });
await fs.mkdir(path.join(previewRoot, "layouts"), { recursive: true });
for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(3, "0")}`;
  const png = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(previewRoot, "slides", `${stem}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(previewRoot, "layouts", `${stem}.json`), await layout.text());
}
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(outputPptx);
console.log(`ARTIFACT_TOOL_SLIDES=${spec.slides.length}`);
console.log(`PPTX=${outputPptx}`);
