import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");
const imageSizeModule = require("image-size");
const imageSize = imageSizeModule.imageSize || imageSizeModule.default || imageSizeModule;

const [specPath, outputPath] = process.argv.slice(2);
if (!specPath || !outputPath) {
  throw new Error("Usage: node render_benchmark_deck.mjs <spec.json> <output.pptx>");
}
const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow v2.2";
pptx.subject = "Controlled synthetic visual backend benchmark";
pptx.title = spec.project_title;
pptx.company = "Academic PPT Workflow canonical repository";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "zh-CN",
};
pptx.defineSlideMaster({
  title: "MASTER",
  background: { color: "F7FAFC" },
  objects: [
    { rect: { x: 0, y: 7.14, w: 13.333, h: 0.36, fill: { color: "17324D" }, line: { color: "17324D" } } },
    { line: { x: 0.72, y: 0.55, w: 0.42, h: 0, line: { color: "3AAFA9", width: 4 } } },
  ],
  slideNumber: { x: 12.45, y: 7.17, w: 0.42, h: 0.18, color: "FFFFFF", fontFace: "Arial", fontSize: 9, align: "right" },
});

const C = {
  navy: "17324D", teal: "167C80", mint: "DDF3F0", blue: "4E79A7",
  amber: "F2B134", coral: "E76F51", ink: "243B53", muted: "627D98",
  pale: "EEF4F7", white: "FFFFFF", red: "B23A48", green: "2A9D8F",
};
const Z = { left: 0.8, right: 12.53, top: 1.55, bottom: 6.78, footer: 7.14 };

function addTitle(slide, title, role) {
  slide.addText(title, {
    x: 0.82, y: 0.43, w: 11.75, h: 0.88,
    fontFace: "Aptos Display", fontSize: title.length > 54 ? 23 : (title.length > 38 ? 26 : 30),
    bold: true, color: C.navy, margin: 0, breakLine: false, fit: "shrink",
  });
  slide.addText(role.toUpperCase(), {
    x: 0.83, y: 1.36, w: 2.2, h: 0.22, fontFace: "Arial",
    fontSize: 9, bold: true, color: C.teal, charSpacing: 1.5, margin: 0,
  });
}

function addFooter(slide, sources) {
  const short = (sources || []).slice(0, 2).join(" · ");
  if (short) slide.addText(`来源：${short}`, {
    x: 0.82, y: 7.17, w: 10.8, h: 0.16, fontFace: "Microsoft YaHei",
    fontSize: 9.5, color: "FFFFFF", margin: 0, valign: "mid",
  });
}

function addNotes(slide, item) {
  const notes = [
    "[Sources]",
    ...(item.notes_sources || []),
    "",
    "[Claims]",
    ...(item.claim_ids || []),
    "",
    "[Wording boundary]",
    item.wording_boundary || "Source-bound synthetic evidence only.",
  ];
  slide.addNotes(notes.join("\n"));
}

function addEmptyState(slide, label) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.35, y: 2.25, w: 10.6, h: 2.6, rectRadius: 0.08,
    fill: { color: C.pale }, line: { color: "C8D6E0", width: 1.2 },
  });
  slide.addText(label, {
    x: 1.75, y: 3.05, w: 9.8, h: 0.85, fontSize: 24, bold: true,
    align: "center", color: C.ink, margin: 0.05,
  });
}

function extent(values, reference) {
  const clean = values.filter(v => Number.isFinite(Number(v))).map(Number);
  if (Number.isFinite(reference)) clean.push(Number(reference));
  if (!clean.length) return [0, 1];
  let lo = Math.min(...clean), hi = Math.max(...clean);
  if (lo === hi) { lo -= 1; hi += 1; }
  const pad = (hi - lo) * 0.12;
  return [lo - pad, hi + pad];
}

function forest(slide, v) {
  const labels = v.categories || [];
  const estimates = v.estimate || [], lows = v.ci_low || [], highs = v.ci_high || [];
  if (!labels.length || estimates.length !== lows.length || estimates.length !== highs.length) {
    addEmptyState(slide, "Forest plot requirements were not met; structured fallback retained.");
    return;
  }
  const n = Math.min(labels.length, 8);
  const [minV, maxV] = extent([...lows.slice(0, n), ...highs.slice(0, n)], v.reference_value);
  const x0 = 5.25, plotW = 6.65, y0 = 1.9, rowH = 0.56;
  const mapX = value => x0 + (Number(value) - minV) / (maxV - minV) * plotW;
  if (Number.isFinite(v.reference_value)) {
    const rx = mapX(v.reference_value);
    slide.addShape(pptx.ShapeType.line, { x: rx, y: y0 - 0.2, w: 0, h: rowH * n + 0.25, line: { color: "91A8B8", width: 1.2, dash: "dash" } });
  }
  for (let i = 0; i < n; i++) {
    const y = y0 + i * rowH;
    slide.addText(String(labels[i]), { x: 0.92, y: y - 0.02, w: 3.95, h: 0.34, fontSize: 15.5, color: C.ink, margin: 0, fit: "shrink" });
    const lx = mapX(lows[i]), hx = mapX(highs[i]), ex = mapX(estimates[i]);
    slide.addShape(pptx.ShapeType.line, { x: lx, y: y + 0.13, w: Math.max(0.02, hx - lx), h: 0, line: { color: C.blue, width: 2.2 } });
    slide.addShape(pptx.ShapeType.ellipse, { x: ex - 0.07, y: y + 0.06, w: 0.14, h: 0.14, fill: { color: i === 0 ? C.coral : C.teal }, line: { color: "FFFFFF", transparency: 100 } });
    slide.addText(`${Number(estimates[i]).toFixed(2)} [${Number(lows[i]).toFixed(2)}, ${Number(highs[i]).toFixed(2)}]`, { x: 11.98, y: y - 0.04, w: 1.05, h: 0.32, fontSize: 10.5, align: "right", color: C.muted, margin: 0, fit: "shrink" });
  }
  slide.addText(`Reference ${v.reference_value ?? ""}`, { x: 5.25, y: 6.45, w: 3.5, h: 0.25, fontSize: 10.5, color: C.muted, margin: 0 });
}

function bars(slide, v) {
  const labels = v.categories || [], values = v.estimate || [];
  const n = Math.min(labels.length, 8);
  const max = Math.max(1, ...values.slice(0, n).map(Number));
  for (let i = 0; i < n; i++) {
    const y = 1.78 + i * 0.61;
    const width = Math.max(0.03, Number(values[i] || 0) / max * 6.25);
    slide.addText(String(labels[i]), { x: 0.95, y, w: 3.7, h: 0.32, fontSize: 14.5, color: C.ink, margin: 0, fit: "shrink" });
    slide.addShape(pptx.ShapeType.roundRect, { x: 4.75, y: y + 0.02, w: width, h: 0.26, rectRadius: 0.03, fill: { color: i % 2 ? C.blue : C.teal }, line: { color: "FFFFFF", transparency: 100 } });
    const count = v.numerator?.[i];
    const den = v.denominator?.[i];
    const suffix = Number.isFinite(Number(count)) && Number.isFinite(Number(den)) ? `${count}/${den}` : "";
    slide.addText(`${Number(values[i] || 0).toFixed(1)}${v.unit === "%" || v.visual_type === "event_rate_plot" ? "%" : ""} ${suffix}`, { x: 11.25, y, w: 1.5, h: 0.3, fontSize: 12, color: C.muted, align: "right", margin: 0 });
  }
}

function love(slide, v) {
  const labels = v.categories || [];
  const a = v.series?.[0]?.values || [], b = v.series?.[1]?.values || [];
  const n = Math.min(labels.length, 8), x0 = 4.5, w = 7.2;
  const all = [...a, ...b].map(Number).filter(Number.isFinite);
  const max = Math.max(0.2, ...all);
  const mapX = val => x0 + Number(val) / max * w;
  const tx = mapX(v.reference_value ?? 0.1);
  slide.addShape(pptx.ShapeType.line, { x: tx, y: 1.72, w: 0, h: 4.75, line: { color: C.coral, width: 1.3, dash: "dash" } });
  for (let i = 0; i < n; i++) {
    const y = 1.9 + i * 0.57;
    slide.addText(String(labels[i]), { x: 0.95, y, w: 3.2, h: 0.28, fontSize: 14, color: C.ink, margin: 0, fit: "shrink" });
    slide.addShape(pptx.ShapeType.line, { x: mapX(a[i]), y: y + 0.12, w: mapX(b[i]) - mapX(a[i]), h: 0, line: { color: "AFC4D2", width: 1 } });
    slide.addShape(pptx.ShapeType.ellipse, { x: mapX(a[i]) - 0.07, y: y + 0.05, w: 0.14, h: 0.14, fill: { color: C.coral }, line: { color: C.coral } });
    slide.addShape(pptx.ShapeType.ellipse, { x: mapX(b[i]) - 0.07, y: y + 0.05, w: 0.14, h: 0.14, fill: { color: C.teal }, line: { color: C.teal } });
  }
}

function graphFromManifest(slide, artifact) {
  const manifest = artifact?.object_manifest || [];
  const nodes = manifest.filter(o => o.kind === "node");
  const edges = manifest.filter(o => o.kind === "directed_edge");
  if (!nodes.length) { addEmptyState(slide, "Graph adapter used the native fallback."); return; }
  const positions = {};
  const cols = nodes.length <= 4 ? nodes.length : 4;
  nodes.forEach((node, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = cols === 1 ? 5.45 : 1.2 + col * (9.85 / Math.max(1, cols - 1));
    const y = nodes.length <= 4 ? 3.25 : 2.25 + row * 2.05;
    positions[`n${i}`] = { x, y };
  });
  edges.forEach(edge => {
    const a = positions[edge.source], b = positions[edge.target];
    if (a && b) {
      const ax = a.x + 1.1, ay = a.y + 0.36;
      const bx = b.x - 1.1, by = b.y + 0.36;
      const left = Math.min(ax, bx), top = Math.min(ay, by);
      const width = Math.max(0.02, Math.abs(bx - ax));
      const height = Math.max(0, Math.abs(by - ay));
      slide.addShape(pptx.ShapeType.line, {
        x: left, y: top, w: width, h: height,
        line: { color: C.muted, width: 1.8, beginArrowType: "none", endArrowType: "triangle" },
      });
    }
  });
  nodes.forEach((node, i) => {
    const p = positions[`n${i}`];
    slide.addShape(pptx.ShapeType.roundRect, { x: p.x - 1.1, y: p.y, w: 2.2, h: 0.78, rectRadius: 0.06, fill: { color: i % 2 ? C.mint : C.white }, line: { color: C.teal, width: 1.4 } });
    slide.addText(String(node.text || ""), { x: p.x - 1.0, y: p.y + 0.1, w: 2.0, h: 0.54, fontSize: 12.2, bold: true, color: C.ink, align: "center", valign: "mid", margin: 0.02, fit: "shrink" });
  });
}

function timeline(slide, v) {
  const labels = (v.categories || []).slice(0, 7);
  if (!labels.length) { addEmptyState(slide, "No timeline stages were available."); return; }
  slide.addShape(pptx.ShapeType.line, { x: 1.2, y: 3.65, w: 10.9, h: 0, line: { color: C.teal, width: 3 } });
  labels.forEach((label, i) => {
    const x = 1.35 + i * (10.55 / Math.max(1, labels.length - 1));
    slide.addShape(pptx.ShapeType.ellipse, { x: x - 0.11, y: 3.54, w: 0.22, h: 0.22, fill: { color: i % 2 ? C.amber : C.teal }, line: { color: "FFFFFF", width: 1 } });
    slide.addText(String(label), { x: x - 0.7, y: i % 2 ? 3.94 : 2.65, w: 1.4, h: 0.62, fontSize: 12.5, bold: true, color: C.ink, align: "center", margin: 0.02, fit: "shrink" });
  });
}

function cards(slide, labels, series) {
  const items = (labels || []).slice(0, 6);
  const cols = items.length <= 3 ? items.length : 3;
  items.forEach((label, i) => {
    const row = Math.floor(i / 3), col = i % 3;
    const x = 0.95 + col * 4.05, y = 1.78 + row * 2.15;
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 3.65, h: 1.65, rectRadius: 0.07, fill: { color: row % 2 ? "FFFFFF" : C.pale }, line: { color: "C9D8E2", width: 1.1 } });
    slide.addText(String(label), { x: x + 0.22, y: y + 0.24, w: 3.2, h: 0.52, fontSize: 16.5, bold: true, color: C.navy, margin: 0, fit: "shrink" });
    if (series?.[i]) slide.addText(Object.values(series[i]).filter(v => v != null).slice(0, 3).join(" · "), { x: x + 0.22, y: y + 0.9, w: 3.2, h: 0.42, fontSize: 11, color: C.muted, margin: 0, fit: "shrink" });
  });
}

function scatter(slide, v) {
  const xs = v.estimate || [];
  const aux = v.series?.[0]?.fdr || [];
  const n = Math.min(xs.length, aux.length || xs.length, 60);
  const xVals = xs.slice(0, n).map(Number);
  const yVals = aux.length ? aux.slice(0, n).map(q => -Math.log10(Math.max(1e-8, Number(q)))) : xVals.map((_, i) => (i % 9) / 2);
  const [xmin, xmax] = extent(xVals, 0), [ymin, ymax] = extent(yVals, 0);
  const mapX = val => 1.4 + (val - xmin) / (xmax - xmin) * 10.1;
  const mapY = val => 6.25 - (val - ymin) / (ymax - ymin) * 4.45;
  slide.addShape(pptx.ShapeType.line, { x: 1.4, y: 6.25, w: 10.1, h: 0, line: { color: C.muted, width: 1 } });
  slide.addShape(pptx.ShapeType.line, { x: mapX(0), y: 1.8, w: 0, h: 4.45, line: { color: "AFC4D2", width: 1, dash: "dash" } });
  for (let i = 0; i < n; i++) {
    slide.addShape(pptx.ShapeType.ellipse, { x: mapX(xVals[i]) - 0.045, y: mapY(yVals[i]) - 0.045, w: 0.09, h: 0.09, fill: { color: xVals[i] >= 0 ? C.teal : C.coral, transparency: 15 }, line: { color: "FFFFFF", transparency: 100 } });
  }
}

function sourceImage(slide, v) {
  const asset = v.annotation_rule?.source_asset;
  if (!asset || !fs.existsSync(asset)) { addEmptyState(slide, "Registered source asset was unavailable."); return; }
  const dimensions = imageSize(fs.readFileSync(asset));
  const box = { x: 1.25, y: 1.65, w: 10.85, h: 4.95 };
  const scale = Math.min(box.w / dimensions.width, box.h / dimensions.height);
  const w = dimensions.width * scale, h = dimensions.height * scale;
  slide.addImage({ path: asset, x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h });
}

function formula(slide, v, artifact) {
  const labels = (v.categories || []).slice(0, 5);
  slide.addShape(pptx.ShapeType.roundRect, { x: 1.0, y: 1.82, w: 5.15, h: 4.45, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.08 });
  slide.addText("n = f(α, power, effect, variance, attrition)", { x: 1.38, y: 2.55, w: 4.35, h: 1.05, fontFace: "Cambria Math", fontSize: 23, color: C.white, align: "center", margin: 0.04, fit: "shrink" });
  slide.addText("Editable text fallback retained", { x: 1.55, y: 4.35, w: 4.0, h: 0.45, fontSize: 12, color: "BFE6E3", align: "center", margin: 0 });
  labels.slice(0, 4).forEach((label, i) => {
    const y = 1.85 + i * 1.08;
    slide.addShape(pptx.ShapeType.roundRect, { x: 6.55, y, w: 5.35, h: 0.82, rectRadius: 0.04, fill: { color: i % 2 ? "FFFFFF" : C.pale }, line: { color: "C9D8E2" } });
    slide.addText(String(label), { x: 6.82, y: y + 0.18, w: 4.8, h: 0.36, fontSize: 15, bold: true, color: C.ink, margin: 0, fit: "shrink" });
  });
}

function renderVisual(slide, item) {
  const v = item.visual || {};
  switch (v.visual_type) {
    case "forest_plot":
    case "sensitivity_plot":
    case "subgroup_plot": forest(slide, v); break;
    case "event_rate_plot":
    case "composition_bar":
    case "enrichment":
    case "distribution": bars(slide, v); break;
    case "love_plot": love(slide, v); break;
    case "flow_diagram":
    case "dag":
    case "network": graphFromManifest(slide, item.artifact); break;
    case "timeline": timeline(slide, v); break;
    case "volcano":
    case "funnel_plot": scatter(slide, v); break;
    case "formula": formula(slide, v, item.artifact); break;
    case "risk_matrix":
    case "grade_summary":
    case "matrix_2x2": cards(slide, v.categories, v.series); break;
    case "weighted_risk_curve":
    case "umap":
    case "source_figure": sourceImage(slide, v); break;
    default: cards(slide, v.categories, v.series);
  }
}

for (const item of spec.slides) {
  const slide = pptx.addSlide("MASTER");
  slide.background = { color: item.role === "cover" ? C.navy : "F7FAFC" };
  if (item.role === "cover") {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.navy }, line: { color: C.navy } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.0, w: 0.12, h: 4.6, fill: { color: C.teal }, line: { color: C.teal } });
    slide.addText(item.title, { x: 1.25, y: 1.45, w: 10.5, h: 1.45, fontFace: "Aptos Display", fontSize: 38, bold: true, color: C.white, margin: 0.02, fit: "shrink" });
    slide.addText(item.subtitle || "", { x: 1.28, y: 3.25, w: 9.8, h: 0.6, fontSize: 20, color: "BFE6E3", margin: 0 });
    slide.addText(item.synthetic_warning || "", { x: 1.28, y: 5.65, w: 10.4, h: 0.48, fontSize: 13, bold: true, color: C.amber, charSpacing: 1.2, margin: 0 });
    slide.addText("Academic PPT Workflow v2.2", { x: 1.28, y: 6.75, w: 4.8, h: 0.3, fontSize: 11, color: "AFC4D2", margin: 0 });
  } else if (item.cards) {
    addTitle(slide, item.title, item.role);
    cards(slide, item.cards, []);
    addFooter(slide, item.sources);
  } else {
    addTitle(slide, item.title, item.role);
    renderVisual(slide, item);
    addFooter(slide, item.sources);
    addNotes(slide, item);
  }
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
await pptx.writeFile({ fileName: outputPath });
console.log(`WROTE=${outputPath}`);
console.log(`SLIDES=${spec.slides.length}`);
