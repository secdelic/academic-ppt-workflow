import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");
const { imageSize } = require("image-size");

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

const styleProfile = spec.style_profile || null;
function themeColor(slot, fallback) {
  const tokens = Array.isArray(styleProfile?.theme_colors) ? styleProfile.theme_colors : [];
  const match = tokens.find((item) => item?.slot === slot && /^[0-9A-Fa-f]{6}$/.test(item?.value || ""));
  return match ? match.value.toUpperCase() : fallback;
}
function themeFont(kind, fallback) {
  const value = styleProfile?.theme_fonts?.[kind]?.latin;
  return value && value !== "UNKNOWN" ? value : fallback;
}
const languageFallback = /^(zh|中文|简体)/i.test(spec.brief.language || "") ? "Microsoft YaHei" : "Arial";
const headFont = themeFont("major", languageFallback);
const bodyFont = themeFont("minor", languageFallback);
pptx.theme = {
  headFontFace: headFont,
  bodyFontFace: bodyFont,
  lang: spec.brief.language || "zh-CN",
};

const C = {
  bg: themeColor("lt1", "F7FAFC"),
  ink: themeColor("dk2", "102A43"),
  muted: themeColor("accent2", "627D98"),
  primary: themeColor("accent1", "0F766E"),
  secondary: themeColor("accent5", "2563EB"),
  accent: themeColor("accent4", "D97706"),
  positive: themeColor("accent6", "70AD47"),
  line: themeColor("lt2", "D9E2EC"),
  white: themeColor("lt1", "FFFFFF"),
};
const font = bodyFont;
const slideWidth = Number(spec.layout_contract?.slide?.width_in || 13.333);
const slideHeight = Number(spec.layout_contract?.slide?.height_in || 7.5);

function frame(item, suffix, required = true) {
  const target = (item.planned_geometry || []).find((entry) => String(entry.object_id || "").endsWith(`:${suffix}`));
  if (!target) {
    if (required) throw new Error(`Layout frame ${suffix} missing for ${item.slide_id}`);
    return null;
  }
  return { ...target.bounds };
}

function intersects(a, b) {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

function validatePlannedGeometry(item) {
  const footer = spec.layout_contract?.zones?.footer_exclusion;
  if (!Array.isArray(item.planned_geometry) || item.planned_geometry.length === 0) {
    throw new Error(`Every slide requires planned_geometry: ${item.slide_id}`);
  }
  for (const object of item.planned_geometry) {
    const b = object.bounds || {};
    if (![b.x, b.y, b.w, b.h].every(Number.isFinite)) {
      throw new Error(`Invalid bounds for ${object.object_id}`);
    }
    if (b.x < 0 || b.y < 0 || b.x + b.w > slideWidth + 0.001 || b.y + b.h > slideHeight + 0.001) {
      throw new Error(`Off-slide planned object: ${object.object_id}`);
    }
    if (footer && !["footer", "page_number"].includes(object.role) && intersects(b, footer)) {
      throw new Error(`Footer intrusion blocked before generation: ${object.object_id}`);
    }
  }
}

function notes(item) {
  const sourceLines = item.source_ids?.length ? item.source_ids.map((s) => `- ${s}`) : ["- INFORMATION_REQUIRED"];
  const claimLines = item.claim_ids?.length ? item.claim_ids.map((s) => `- ${s}`) : ["- NONE"];
  const boundary = item.prohibited_overstatement || item.evidence_strength || "Preserve source wording and uncertainty.";
  const chart = item.chart_data?.chart_id ? `\n[ChartDataContract]\n- ${item.chart_data.chart_id}\n- ${item.chart_data.source_path || ""}` : "";
  return `${item.speaker_note_summary || item.speaker_notes || ""}\n\n[Slide-ID]\n${item.slide_id || "INFORMATION_REQUIRED"}\n[Slide-Revision]\n${item.slide_revision || 1}\n[Content-Hash]\n${item.content_hash || "INFORMATION_REQUIRED"}\n[Claims]\n${claimLines.join("\n")}\n[Wording-Boundary]\n${boundary}\n[Sources]\n${sourceLines.join("\n")}${chart}`;
}

function addFooter(slide, item, index) {
  const source = frame(item, "source");
  const page = frame(item, "page");
  const lineY = Math.min(source.y - 0.08, slideHeight * 0.90);
  slide.addShape(pptx.ShapeType.line, {
    x: 0.75, y: lineY, w: 11.83, h: 0,
    line: { color: C.line, width: 1 },
    objectName: `awf:${item.slide_id}:footer-line`,
  });
  slide.addText(item.short_source_label || "来源：输入材料", {
    ...source, fontFace: font, fontSize: 10, color: C.muted,
    margin: 0, breakLine: false,
    objectName: `awf:${item.slide_id}:sources`,
  });
  slide.addText(String(index + 1), {
    ...page, fontFace: font, fontSize: 11, color: C.muted,
    align: "right", margin: 0,
    objectName: `awf:${item.slide_id}:page-number`,
  });
}

function addTitle(slide, item) {
  const title = frame(item, "title");
  slide.addShape(pptx.ShapeType.rect, {
    x: title.x, y: Math.max(0.42, title.y - 0.13), w: 0.54, h: 0.07,
    fill: { color: C.primary }, line: { color: C.primary, transparency: 100 },
    objectName: `awf:${item.slide_id}:title-accent`,
  });
  slide.addText(item.slide_title, {
    ...title, fontFace: headFont, fontSize: 32, bold: true,
    color: C.ink, margin: 0, breakLine: false,
    objectName: `awf:${item.slide_id}:title`,
  });
}

function addMessage(slide, item, suffix = "message") {
  const box = frame(item, suffix);
  slide.addText(item.single_key_message, {
    ...box, fontFace: font, fontSize: item.visual_type === "forest_plot" || item.visual_type === "subgroup_forest_plot" ? 20 : 22,
    bold: true, color: C.ink, margin: 0.05, valign: "mid",
    objectName: `awf:${item.slide_id}:key-message`,
  });
}

function addChartCaption(slide, item) {
  const box = frame(item, "metadata", false);
  if (!box || !item.chart_caption) return;
  slide.addText(item.chart_caption, {
    ...box, fontFace: font, fontSize: 15, color: C.muted,
    margin: 0, valign: "top",
    objectName: `awf:${item.slide_id}:chart-metadata`,
  });
}

function addBarChart(slide, item, missingness = false) {
  addMessage(slide, item);
  addChartCaption(slide, item);
  const box = frame(item, "visual");
  const data = item.chart_data;
  slide.addChart(
    pptx.ChartType.bar,
    [{ name: missingness ? "Missing" : "Rate", labels: data.categories, values: data.values }],
    {
      ...box,
      catAxisLabelFontFace: font,
      catAxisLabelFontSize: 15,
      valAxisLabelFontFace: font,
      valAxisLabelFontSize: 15,
      showLegend: false,
      showValue: true,
      chartColors: [missingness ? C.accent : C.secondary],
      dataLabelPosition: "outEnd",
      dataLabelFormatCode: data.units === "percent" ? '0.0"%"' : "0.0",
      showValAxisTitle: true,
      valAxisTitle: data.units || "Value",
      valAxisLabelFormatCode: data.units === "percent" ? '0"%"' : "0.0",
      valGridLine: { color: C.line, width: 1 },
      border: { color: C.white, transparency: 100 },
      objectName: `awf:${item.slide_id}:chart`,
    },
  );
}

function addForestPlot(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "visual");
  const data = item.chart_data;
  const intervals = data.confidence_interval || [];
  const lows = intervals.map((value) => Number(value.low));
  const highs = intervals.map((value) => Number(value.high));
  const estimates = data.values.map(Number);
  const reference = Number(data.reference_value ?? 1);
  const rawMin = Math.min(reference, ...lows);
  const rawMax = Math.max(reference, ...highs);
  const padding = Math.max(0.2, (rawMax - rawMin) * 0.08);
  const min = Math.max(0, rawMin - padding);
  const max = rawMax + padding;
  const labelWidth = box.w * 0.31;
  const plotX = box.x + labelWidth;
  const plotW = box.w - labelWidth - 0.12;
  const scaleX = (value) => plotX + ((value - min) / Math.max(0.001, max - min)) * plotW;
  const refX = scaleX(reference);
  slide.addShape(pptx.ShapeType.line, {
    x: refX, y: box.y + 0.12, w: 0, h: box.h - 0.30,
    line: { color: C.muted, width: 1.3, dash: "dash" },
    objectName: `awf:${item.slide_id}:reference-line`,
  });
  const count = Math.max(1, data.categories.length);
  const rowH = (box.h - 0.35) / count;
  data.categories.forEach((label, index) => {
    const cy = box.y + 0.18 + rowH * (index + 0.5);
    slide.addText(String(label), {
      x: box.x, y: cy - rowH * 0.34, w: labelWidth - 0.10, h: rowH * 0.68,
      fontFace: font, fontSize: 15, color: C.ink, margin: 0,
      valign: "mid", objectName: `awf:${item.slide_id}:forest-label:${index + 1}`,
    });
    const low = Number(intervals[index]?.low);
    const high = Number(intervals[index]?.high);
    const estimate = Number(estimates[index]);
    const x1 = scaleX(low);
    const x2 = scaleX(high);
    const px = scaleX(estimate);
    slide.addShape(pptx.ShapeType.line, {
      x: x1, y: cy, w: x2 - x1, h: 0,
      line: { color: C.secondary, width: 2 },
      objectName: `awf:${item.slide_id}:ci:${index + 1}`,
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: px - 0.065, y: cy - 0.065, w: 0.13, h: 0.13,
      fill: { color: C.primary }, line: { color: C.primary },
      objectName: `awf:${item.slide_id}:estimate:${index + 1}`,
    });
    slide.addText(`${estimate.toFixed(2)} [${low.toFixed(2)}, ${high.toFixed(2)}]`, {
      x: Math.min(box.x + box.w - 1.75, x2 + 0.08), y: cy - 0.14, w: 1.65, h: 0.28,
      fontFace: font, fontSize: 15, color: C.muted, margin: 0,
      objectName: `awf:${item.slide_id}:forest-value:${index + 1}`,
    });
  });
  slide.addText(item.chart_caption || "", {
    x: plotX, y: box.y + box.h - 0.22, w: plotW, h: 0.22,
    fontFace: font, fontSize: 10, color: C.muted, margin: 0,
    align: "center", objectName: `awf:${item.slide_id}:forest-caption`,
  });
}

function addSourceFigure(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "visual");
  if (!item.visual_asset_path || !fs.existsSync(item.visual_asset_path)) {
    throw new Error(`Source figure is missing: ${item.visual_asset_path}`);
  }
  const dimensions = imageSize(fs.readFileSync(item.visual_asset_path));
  const scale = Math.min(box.w / dimensions.width, box.h / dimensions.height);
  const imageW = dimensions.width * scale;
  const imageH = dimensions.height * scale;
  slide.addImage({
    path: item.visual_asset_path,
    x: box.x + (box.w - imageW) / 2,
    y: box.y + (box.h - imageH) / 2,
    w: imageW,
    h: imageH,
    objectName: `awf:${item.slide_id}:source-figure`,
  });
}

function addTimeline(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "timeline");
  const milestones = item.timeline_spec?.milestones || [];
  const count = Math.max(2, milestones.length);
  const lineY = box.y + box.h * 0.47;
  slide.addShape(pptx.ShapeType.line, {
    x: box.x + 0.35, y: lineY, w: box.w - 0.70, h: 0,
    line: { color: C.secondary, width: 2.2 },
    objectName: `awf:${item.slide_id}:timeline-line`,
  });
  milestones.forEach((milestone, index) => {
    const x = box.x + 0.35 + ((box.w - 0.70) * index) / Math.max(1, count - 1);
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x - 0.10, y: lineY - 0.10, w: 0.20, h: 0.20,
      fill: { color: C.primary }, line: { color: C.primary },
      objectName: `awf:${item.slide_id}:timeline-node:${index + 1}`,
    });
    slide.addText(String(milestone.value || ""), {
      x: x - 0.55, y: lineY - 0.75, w: 1.10, h: 0.35,
      fontFace: font, fontSize: 18, bold: true, color: C.ink,
      align: "center", margin: 0,
      objectName: `awf:${item.slide_id}:timeline-value:${index + 1}`,
    });
    slide.addText(String(milestone.label || ""), {
      x: x - 0.85, y: lineY + 0.25, w: 1.70, h: 0.50,
      fontFace: font, fontSize: 16, color: C.muted,
      align: "center", margin: 0,
      objectName: `awf:${item.slide_id}:timeline-label:${index + 1}`,
    });
  });
}

function addLineChart(slide, item) {
  addMessage(slide, item);
  addChartCaption(slide, item);
  const box = frame(item, "visual");
  const data = item.chart_data || {};
  const chartSeries = Array.isArray(data.series) && data.series.length
    ? data.series.map((series) => ({
        name: String(series.name || "Series"),
        labels: series.labels || data.categories || [],
        values: series.values || [],
      }))
    : [{ name: "Value", labels: data.categories || [], values: data.values || [] }];
  slide.addChart(pptx.ChartType.line, chartSeries, {
    ...box,
    catAxisLabelFontFace: font,
    catAxisLabelFontSize: 15,
    valAxisLabelFontFace: font,
    valAxisLabelFontSize: 15,
    showLegend: chartSeries.length > 1,
    legendFontFace: font,
    legendFontSize: 14,
    showValue: false,
    showTitle: false,
    showValAxisTitle: true,
    valAxisTitle: data.units || "Value",
    valGridLine: { color: C.line, width: 1 },
    chartColors: [C.primary, C.secondary, C.accent, C.positive],
    lineSize: 2.2,
    showMarker: true,
    border: { color: C.white, transparency: 100 },
    objectName: `awf:${item.slide_id}:line-chart`,
  });
}

function addMatrix(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "structure");
  const specMatrix = item.matrix_spec || {};
  const cells = Array.isArray(specMatrix.cells) ? specMatrix.cells.slice(0, 4) : [];
  const labels = cells.length === 4
    ? cells.map((cell) => String(cell.label || cell.value || ""))
    : ["0 / 0", "0 / 1", "1 / 0", "1 / 1"];
  const gap = 0.18;
  const cellW = (box.w - gap) / 2;
  const cellH = (box.h - gap) / 2;
  labels.forEach((label, index) => {
    const column = index % 2;
    const row = Math.floor(index / 2);
    const x = box.x + column * (cellW + gap);
    const y = box.y + row * (cellH + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: cellW, h: cellH,
      fill: { color: index === 3 ? "E9F5F3" : C.white },
      line: { color: index === 3 ? C.primary : C.line, width: 1.3 },
      radius: 0.04,
      objectName: `awf:${item.slide_id}:matrix-cell:${index + 1}`,
    });
    slide.addText(label, {
      x: x + 0.18, y: y + 0.15, w: cellW - 0.36, h: cellH - 0.30,
      fontFace: font, fontSize: 18, bold: true, color: C.ink,
      align: "center", valign: "mid", margin: 0,
      objectName: `awf:${item.slide_id}:matrix-label:${index + 1}`,
    });
  });
}

function addProblemGapObjective(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "structure");
  const payload = item.problem_gap_objective || {};
  const panels = [
    { heading: "Problem", text: payload.problem || "INFORMATION_REQUIRED" },
    { heading: "Gap", text: payload.gap || "INFORMATION_REQUIRED" },
    { heading: "Objective", text: payload.objective || item.single_key_message },
  ];
  const gap = 0.22;
  const panelW = (box.w - gap * 2) / 3;
  panels.forEach((panel, index) => {
    const x = box.x + index * (panelW + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: box.y, w: panelW, h: box.h,
      fill: { color: index === 1 ? "FFF7ED" : C.white },
      line: { color: index === 1 ? C.accent : C.line, width: 1.3 },
      radius: 0.04,
      objectName: `awf:${item.slide_id}:pgo-panel:${index + 1}`,
    });
    slide.addText(panel.heading, {
      x: x + 0.22, y: box.y + 0.24, w: panelW - 0.44, h: 0.34,
      fontFace: font, fontSize: 21, bold: true,
      color: index === 1 ? C.accent : C.primary, margin: 0,
      objectName: `awf:${item.slide_id}:pgo-heading:${index + 1}`,
    });
    slide.addText(String(panel.text), {
      x: x + 0.22, y: box.y + 0.82, w: panelW - 0.44, h: box.h - 1.06,
      fontFace: font, fontSize: 18, color: C.ink, margin: 0,
      valign: "mid",
      objectName: `awf:${item.slide_id}:pgo-body:${index + 1}`,
    });
  });
}

function addNativeFlow(slide, item) {
  const box = frame(item, "structure");
  const nodes = item.flow_spec?.nodes || [];
  const usable = nodes.length ? nodes.slice(0, 6) : [];
  const count = Math.max(1, usable.length);
  const gap = 0.20;
  const nodeH = Math.min(0.62, (box.h - gap * (count - 1)) / count);
  usable.forEach((node, index) => {
    const y = box.y + index * (nodeH + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: box.x + 1.25, y, w: box.w - 2.50, h: nodeH,
      fill: { color: C.white },
      line: { color: node.type === "exclusion" ? C.accent : C.primary, width: 1.4 },
      radius: 0.04,
      objectName: `awf:${item.slide_id}:flow-node:${index + 1}`,
    });
    slide.addText(String(node.label || node.value || ""), {
      x: box.x + 1.50, y: y + 0.05, w: box.w - 3.00, h: nodeH - 0.10,
      fontFace: font, fontSize: 17, color: C.ink,
      align: "center", valign: "mid", margin: 0,
      objectName: `awf:${item.slide_id}:flow-label:${index + 1}`,
    });
    if (index < count - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: box.x + box.w / 2, y: y + nodeH, w: 0, h: gap,
        line: { color: C.muted, width: 1.2, endArrowType: "triangle" },
        objectName: `awf:${item.slide_id}:flow-edge:${index + 1}`,
      });
    }
  });
}

function addSourceRegistry(slide, item) {
  const box = frame(item, "structure");
  const rows = (item.source_records || []).slice(0, 8).map((source, index) => [
    String(index + 1).padStart(2, "0"),
    source.file_name,
    String(source.file_type || "").toUpperCase(),
  ]);
  if ((item.source_records || []).length > rows.length) {
    rows.push(["…", `其余 ${(item.source_records || []).length - rows.length} 项见 source_manifest.csv`, "AUDIT"]);
  }
  slide.addTable(
    [
      [
        { text: "序号", options: { bold: true, color: C.white } },
        { text: "输入文件", options: { bold: true, color: C.white } },
        { text: "类型", options: { bold: true, color: C.white } },
      ],
      ...rows,
    ],
    {
      ...box, colW: [0.8, box.w - 2.0, 1.2], rowH: 0.38,
      fontFace: font, fontSize: 16, color: C.ink,
      border: { color: C.line, width: 1 },
      fill: C.white, margin: 0.05,
      autoFit: false,
      objectName: `awf:${item.slide_id}:source-registry`,
    },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: box.x, y: box.y, w: box.w, h: 0.38,
    fill: { color: C.secondary, transparency: 100 },
    line: { color: C.secondary, transparency: 100 },
    objectName: `awf:${item.slide_id}:source-registry-header-anchor`,
  });
}

function addComparison(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "structure");
  const gap = 0.35;
  const panelW = (box.w - gap) / 2;
  const panels = [
    { title: "Allowed", color: C.positive, text: "保留研究设计、统计不确定性和来源措辞。" },
    { title: "Not allowed", color: C.accent, text: "不得将关联或计算推断升级为因果、机制或临床证明。" },
  ];
  panels.forEach((panel, index) => {
    const x = box.x + index * (panelW + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: box.y, w: panelW, h: box.h,
      fill: { color: C.white }, line: { color: panel.color, width: 1.4 },
      radius: 0.05, objectName: `awf:${item.slide_id}:comparison-panel:${index + 1}`,
    });
    slide.addText(panel.title, {
      x: x + 0.25, y: box.y + 0.25, w: panelW - 0.5, h: 0.38,
      fontFace: font, fontSize: 24, bold: true, color: panel.color, margin: 0,
      objectName: `awf:${item.slide_id}:comparison-title:${index + 1}`,
    });
    slide.addText(panel.text, {
      x: x + 0.25, y: box.y + 0.92, w: panelW - 0.5, h: box.h - 1.20,
      fontFace: font, fontSize: 18, color: C.ink, margin: 0,
      objectName: `awf:${item.slide_id}:comparison-body:${index + 1}`,
    });
  });
}

function addTakeaway(slide, item) {
  addMessage(slide, item);
  const box = frame(item, "structure");
  const items = (item.takeaways || [item.single_key_message, item.prohibited_overstatement]).slice(0, 3);
  const count = Math.max(1, items.length);
  const gap = 0.22;
  const itemW = (box.w - gap * (count - 1)) / count;
  items.forEach((text, index) => {
    const x = box.x + index * (itemW + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: box.y, w: itemW, h: box.h,
      fill: { color: index === 0 ? "F2F6FC" : C.white },
      line: { color: index === 0 ? C.secondary : C.line, width: 1.2 },
      radius: 0.05, objectName: `awf:${item.slide_id}:takeaway-card:${index + 1}`,
    });
    slide.addText(String(text || ""), {
      x: x + 0.25, y: box.y + 0.35, w: itemW - 0.5, h: box.h - 0.70,
      fontFace: font, fontSize: 18, bold: index === 0,
      color: C.ink, margin: 0, valign: "mid",
      objectName: `awf:${item.slide_id}:takeaway-text:${index + 1}`,
    });
  });
}

function addAuditTable(slide, item) {
  const box = frame(item, "structure");
  const rows = (item.audit_rows || []).map((row) => [row.item || "", row.decision || "", row.status || ""]);
  slide.addTable(
    [
      [
        { text: "审计项", options: { bold: true, color: C.white, fill: C.secondary } },
        { text: "处理", options: { bold: true, color: C.white, fill: C.secondary } },
        { text: "状态", options: { bold: true, color: C.white, fill: C.secondary } },
      ],
      ...rows,
    ],
    {
      ...box, colW: [1.35, box.w - 3.15, 1.80], rowH: 0.70,
      fontFace: font, fontSize: 16, color: C.ink,
      border: { color: C.line, width: 1 }, fill: C.white,
      margin: 0.08, autoFit: false,
      objectName: `awf:${item.slide_id}:audit-table`,
    },
  );
}

function addDiagram(slide, item) {
  const diagram = item.diagram_spec;
  const nodes = diagram?.nodes || [];
  const nodeMap = new Map(nodes.map((node) => [node.node_id, node]));
  const area = frame(item, "structure", false) || frame(item, "visual", false) || { x: 0.85, y: 2.0, w: 11.65, h: 4.25 };
  for (const [edgeIndex, edge] of (diagram?.edges || []).entries()) {
    const source = nodeMap.get(edge.source);
    const target = nodeMap.get(edge.target);
    if (!source || !target) continue;
    const x1 = area.x + (source.x + source.w / 2) * area.w;
    const y1 = area.y + (source.y + source.h / 2) * area.h;
    const x2 = area.x + (target.x + target.w / 2) * area.w;
    const y2 = area.y + (target.y + target.h / 2) * area.h;
    slide.addShape(pptx.ShapeType.line, {
      x: x1, y: y1, w: x2 - x1, h: y2 - y1,
      line: { color: C.muted, width: 1.5, endArrowType: "triangle" },
      objectName: `awf:${item.slide_id}:edge:${edgeIndex + 1}`,
    });
  }
  for (const node of nodes) {
    const roleColor = node.role === "exclusion" ? C.accent : node.role === "outcome" ? C.secondary : C.primary;
    const shapeType = node.shape === "ellipse" ? pptx.ShapeType.ellipse : node.shape === "diamond" ? pptx.ShapeType.diamond : node.shape === "hexagon" ? pptx.ShapeType.hexagon : pptx.ShapeType.roundRect;
    const x = area.x + node.x * area.w;
    const y = area.y + node.y * area.h;
    const w = node.w * area.w;
    const h = node.h * area.h;
    slide.addShape(shapeType, {
      x, y, w, h, fill: { color: C.white },
      line: { color: roleColor, width: 1.6 }, radius: 0.08,
      objectName: `awf:${item.slide_id}:node:${node.node_id}`,
    });
    slide.addText(node.label, {
      x: x + 0.05, y: y + 0.02, w: Math.max(0.1, w - 0.1), h: Math.max(0.1, h - 0.04),
      fontFace: font, fontSize: 16, bold: true, color: C.ink,
      align: "center", valign: "mid", margin: 0.02,
      objectName: `awf:${item.slide_id}:node-label:${node.node_id}`,
    });
  }
}

for (const [index, item] of spec.slides.entries()) {
  validatePlannedGeometry(item);
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  if (item.slide_role === "cover") {
    const title = frame(item, "cover-title");
    const message = frame(item, "cover-message");
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 0.25, h: slideHeight * 0.885,
      fill: { color: C.primary }, line: { color: C.primary },
      objectName: `awf:${item.slide_id}:cover-bar`,
    });
    slide.addText("ACADEMIC PRESENTATION", {
      x: title.x, y: 0.9, w: 5.2, h: 0.35,
      fontFace: font, fontSize: 12, bold: true, color: C.primary, margin: 0,
      objectName: `awf:${item.slide_id}:eyebrow`,
    });
    slide.addText(item.slide_title.replaceAll("_", " "), {
      ...title, fontFace: headFont, fontSize: 50, bold: true,
      color: C.ink, margin: 0,
      objectName: `awf:${item.slide_id}:title`,
    });
    slide.addText(item.single_key_message, {
      ...message, fontFace: font, fontSize: 21, color: C.muted, margin: 0,
      objectName: `awf:${item.slide_id}:key-message`,
    });
  } else {
    addTitle(slide, item);
    if (item.visual_type === "event_rate_chart" || item.visual_type === "editable_bar_chart") {
      addBarChart(slide, item, false);
    } else if (item.visual_type === "missingness_chart") {
      addBarChart(slide, item, true);
    } else if (item.visual_type === "forest_plot" || item.visual_type === "subgroup_forest_plot") {
      addForestPlot(slide, item);
    } else if (item.visual_type === "source_figure") {
      addSourceFigure(slide, item);
    } else if (item.visual_type === "timeline") {
      addTimeline(slide, item);
    } else if (item.visual_type === "line_chart") {
      addLineChart(slide, item);
    } else if (item.visual_type === "matrix_2x2") {
      addMatrix(slide, item);
    } else if (item.visual_type === "problem_gap_objective") {
      addProblemGapObjective(slide, item);
    } else if (item.visual_type === "native_flow") {
      addNativeFlow(slide, item);
    } else if (item.visual_type === "source_registry") {
      addSourceRegistry(slide, item);
    } else if (item.visual_type === "comparison") {
      addComparison(slide, item);
    } else if (item.visual_type === "takeaway") {
      addTakeaway(slide, item);
    } else if (item.visual_type === "audit_table") {
      addAuditTable(slide, item);
    } else if (item.diagram_spec) {
      addDiagram(slide, item);
    } else {
      addMessage(slide, item);
    }
    addFooter(slide, item, index);
  }
  if (typeof slide.addNotes === "function") slide.addNotes(notes(item));
}

fs.mkdirSync(path.dirname(outputPptx), { recursive: true });
fs.mkdirSync(path.join(previewRoot, "layouts"), { recursive: true });
fs.writeFileSync(
  path.join(previewRoot, "layouts", "planned-layout-contract.json"),
  JSON.stringify({ schema_version: "2.1", layout_contract: spec.layout_contract, slides: spec.slides.map((slide) => ({ slide_id: slide.slide_id, planned_geometry: slide.planned_geometry })) }, null, 2),
);
await pptx.writeFile({ fileName: outputPptx });
console.log(`PPTXGENJS_SLIDES=${spec.slides.length}`);
console.log(`PPTX=${outputPptx}`);
