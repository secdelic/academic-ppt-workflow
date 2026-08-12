/**
 * Isolated Arm B behavioral benchmark adapter.
 *
 * This file is an independent local implementation.  It does not copy or
 * execute any audited upstream code.  It benchmarks only high-level behaviors
 * selected after static review: action-title hierarchy, one dominant exhibit,
 * slide-role variation, restrained density, and editable native objects.
 *
 * It is deliberately outside the canonical runner and is not a production
 * backend.  Scientific content is supplied by the canonical Deck IR adapter;
 * this module does not research, infer, or complete facts.
 */

import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");

const [specPath, outputPptx, auditRoot] = process.argv.slice(2);
if (!specPath || !outputPptx || !auditRoot) {
  throw new Error(
    "Usage: generate_upstream_behavior_arm.mjs <build_spec.json> <output.pptx> <audit-root>",
  );
}

const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
if (!Array.isArray(spec.slides) || spec.slides.length === 0) {
  throw new Error("Benchmark adapter requires a non-empty slides array");
}

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow v2 benchmark";
pptx.company = "Local controlled benchmark";
pptx.subject = "Isolated upstream-behavior comparison arm";
pptx.title = `${spec.brief.project_name} — Arm B`;
pptx.lang = spec.brief.language || "zh-CN";

const font = spec.brief.language?.startsWith("zh") ? "Microsoft YaHei" : "Aptos";
const C = {
  navy: "132A3A",
  navy2: "1E3A4C",
  teal: "0D9488",
  cyan: "56B4E9",
  amber: "E69F00",
  coral: "D55E00",
  ink: "172B3A",
  muted: "5F7180",
  paper: "F8FAFC",
  white: "FFFFFF",
  line: "D8E1E8",
  paleTeal: "E8F5F3",
  paleBlue: "EAF3FA",
  paleAmber: "FBF2DD",
};

pptx.theme = {
  headFontFace: font,
  bodyFontFace: font,
  lang: spec.brief.language || "zh-CN",
};

function sourceLines(item) {
  return item.source_ids?.length
    ? item.source_ids.map((source) => `- ${source}`)
    : ["- INFORMATION_REQUIRED"];
}

function noteText(item) {
  return [
    item.speaker_note_summary || item.speaker_notes || "",
    "",
    "[Slide-ID]",
    item.slide_id || "INFORMATION_REQUIRED",
    "[Slide-Revision]",
    String(item.slide_revision || 1),
    "[Content-Hash]",
    item.content_hash || "INFORMATION_REQUIRED",
    "[Sources]",
    ...sourceLines(item),
    "",
    "[Benchmark-Arm]",
    "B — isolated high-level upstream-behavior benchmark; no upstream code executed",
  ].join("\n");
}

function addText(slide, text, options, objectName) {
  slide.addText(String(text ?? ""), { ...options, objectName });
}

function addHeader(slide, item, index) {
  slide.background = { color: C.paper };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 0.2,
    h: 7.5,
    fill: { color: C.teal },
    line: { color: C.teal, transparency: 100 },
    objectName: `benchb:${item.slide_id}:rail`,
  });
  addText(
    slide,
    String(index + 1).padStart(2, "0"),
    {
      x: 0.55,
      y: 0.52,
      w: 0.55,
      h: 0.3,
      fontFace: font,
      fontSize: 13,
      bold: true,
      color: C.teal,
      margin: 0,
    },
    `benchb:${item.slide_id}:index`,
  );
  addText(
    slide,
    item.slide_title,
    {
      x: 1.25,
      y: 0.43,
      w: 10.9,
      h: 0.95,
      fontFace: font,
      fontSize: 31,
      bold: true,
      color: C.ink,
      margin: 0,
      fit: "shrink",
      breakLine: false,
    },
    `benchb:${item.slide_id}:title`,
  );
  slide.addShape(pptx.ShapeType.line, {
    x: 1.25,
    y: 1.55,
    w: 11.25,
    h: 0,
    line: { color: C.line, width: 1 },
    objectName: `benchb:${item.slide_id}:header-rule`,
  });
}

function addFooter(slide, item) {
  const source = item.source_ids?.length
    ? `Sources: ${item.source_ids.join(", ")}`
    : "Source: INFORMATION_REQUIRED";
  addText(
    slide,
    source,
    {
      x: 1.25,
      y: 7.03,
      w: 10.9,
      h: 0.24,
      fontFace: font,
      fontSize: 10.5,
      color: C.muted,
      margin: 0,
      fit: "shrink",
    },
    `benchb:${item.slide_id}:sources`,
  );
}

function addCover(slide, item) {
  slide.background = { color: C.navy };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.18,
    fill: { color: C.teal },
    line: { color: C.teal, transparency: 100 },
    objectName: `benchb:${item.slide_id}:cover-accent`,
  });
  addText(
    slide,
    "CONTROLLED ACADEMIC BENCHMARK",
    {
      x: 0.85,
      y: 0.78,
      w: 5.5,
      h: 0.3,
      fontFace: font,
      fontSize: 11,
      bold: true,
      color: C.cyan,
      charSpacing: 1.2,
      margin: 0,
    },
    `benchb:${item.slide_id}:eyebrow`,
  );
  addText(
    slide,
    item.slide_title,
    {
      x: 0.85,
      y: 1.42,
      w: 10.9,
      h: 1.9,
      fontFace: font,
      fontSize: 43,
      bold: true,
      color: C.white,
      margin: 0,
      fit: "shrink",
      breakLine: false,
    },
    `benchb:${item.slide_id}:title`,
  );
  addText(
    slide,
    item.single_key_message,
    {
      x: 0.85,
      y: 3.62,
      w: 9.8,
      h: 1.12,
      fontFace: font,
      fontSize: 20,
      color: "D9E7EF",
      margin: 0,
      fit: "shrink",
    },
    `benchb:${item.slide_id}:message`,
  );
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 9.95,
    y: 5.32,
    w: 2.25,
    h: 0.7,
    rectRadius: 0.08,
    fill: { color: C.teal },
    line: { color: C.teal, transparency: 100 },
    objectName: `benchb:${item.slide_id}:mode-card`,
  });
  addText(
    slide,
    spec.deck_ir?.narrative_mode || spec.brief.presentation_type,
    {
      x: 10.15,
      y: 5.52,
      w: 1.85,
      h: 0.24,
      fontFace: font,
      fontSize: 11,
      bold: true,
      color: C.white,
      align: "center",
      margin: 0,
      fit: "shrink",
    },
    `benchb:${item.slide_id}:mode`,
  );
}

function addSourceMap(slide, item) {
  addText(
    slide,
    item.single_key_message,
    {
      x: 1.25,
      y: 1.92,
      w: 10.9,
      h: 0.55,
      fontFace: font,
      fontSize: 20,
      bold: true,
      color: C.navy2,
      margin: 0,
      fit: "shrink",
    },
    `benchb:${item.slide_id}:message`,
  );
  const records = (item.source_records || []).slice(0, 6);
  records.forEach((record, index) => {
    const column = index % 2;
    const row = Math.floor(index / 2);
    const x = 1.25 + column * 5.65;
    const y = 2.78 + row * 1.08;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 5.25,
      h: 0.82,
      rectRadius: 0.05,
      fill: { color: index % 3 === 0 ? C.paleTeal : C.white },
      line: { color: C.line, width: 1 },
      objectName: `benchb:${item.slide_id}:source-card:${index + 1}`,
    });
    addText(
      slide,
      record.file_name,
      {
        x: x + 0.23,
        y: y + 0.16,
        w: 3.35,
        h: 0.25,
        fontFace: font,
        fontSize: 15,
        bold: true,
        color: C.ink,
        margin: 0,
        fit: "shrink",
      },
      `benchb:${item.slide_id}:source-name:${index + 1}`,
    );
    addText(
      slide,
      `${String(record.file_type || "").toUpperCase()} · ${record.source_id}`,
      {
        x: x + 0.23,
        y: y + 0.48,
        w: 4.75,
        h: 0.19,
        fontFace: font,
        fontSize: 9.5,
        color: C.muted,
        margin: 0,
        fit: "shrink",
      },
      `benchb:${item.slide_id}:source-id:${index + 1}`,
    );
  });
}

function addChart(slide, item) {
  const values = item.chart_data?.values || [];
  const labels = item.chart_data?.categories || [];
  const unit = item.chart_data?.unit || "unit not provided";
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.25,
    y: 2.0,
    w: 3.45,
    h: 3.95,
    rectRadius: 0.06,
    fill: { color: C.navy },
    line: { color: C.navy, transparency: 100 },
    objectName: `benchb:${item.slide_id}:message-card`,
  });
  addText(
    slide,
    item.single_key_message,
    {
      x: 1.58,
      y: 2.42,
      w: 2.78,
      h: 1.86,
      fontFace: font,
      fontSize: 21,
      bold: true,
      color: C.white,
      margin: 0,
      fit: "shrink",
      valign: "mid",
    },
    `benchb:${item.slide_id}:message`,
  );
  addText(
    slide,
    `Editable native chart · ${unit}`,
    {
      x: 1.58,
      y: 5.2,
      w: 2.78,
      h: 0.28,
      fontFace: font,
      fontSize: 10.5,
      color: "C9D9E3",
      margin: 0,
    },
    `benchb:${item.slide_id}:chart-caption`,
  );
  slide.addChart(
    pptx.ChartType.bar,
    [{ name: "Value", labels, values }],
    {
      x: 5.15,
      y: 2.0,
      w: 7.15,
      h: 3.95,
      chartColors: [C.teal],
      showLegend: false,
      showTitle: false,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelColor: C.ink,
      dataLabelFormatCode: unit === "percent" ? '0"%"' : "0.0",
      catAxisLabelFontFace: font,
      catAxisLabelFontSize: 13,
      valAxisLabelFontFace: font,
      valAxisLabelFontSize: 12,
      valGridLine: { color: C.line, width: 1 },
      showValAxisTitle: true,
      valAxisTitle: unit,
      showCatAxisTitle: false,
      border: { color: C.white, transparency: 100 },
      objectName: `benchb:${item.slide_id}:chart`,
    },
  );
}

function addDiagram(slide, item) {
  const diagram = item.diagram_spec || {};
  const nodes = diagram.nodes || [];
  const byId = new Map(nodes.map((node) => [node.node_id, node]));
  const area = { x: 1.35, y: 2.0, w: 10.8, h: 4.35 };
  (diagram.edges || []).forEach((edge, index) => {
    const a = byId.get(edge.source);
    const b = byId.get(edge.target);
    if (!a || !b) return;
    const x1 = area.x + (a.x + a.w / 2) * area.w;
    const y1 = area.y + (a.y + a.h / 2) * area.h;
    const x2 = area.x + (b.x + b.w / 2) * area.w;
    const y2 = area.y + (b.y + b.h / 2) * area.h;
    slide.addShape(pptx.ShapeType.line, {
      x: x1,
      y: y1,
      w: x2 - x1,
      h: y2 - y1,
      line: { color: C.muted, width: 1.5, endArrowType: "triangle" },
      objectName: `benchb:${item.slide_id}:edge:${index + 1}`,
    });
  });
  nodes.forEach((node, index) => {
    const x = area.x + node.x * area.w;
    const y = area.y + node.y * area.h;
    const w = node.w * area.w;
    const h = node.h * area.h;
    const fill = [C.paleTeal, C.paleBlue, C.paleAmber][index % 3];
    const shape =
      node.shape === "ellipse"
        ? pptx.ShapeType.ellipse
        : node.shape === "diamond"
          ? pptx.ShapeType.diamond
          : pptx.ShapeType.roundRect;
    slide.addShape(shape, {
      x,
      y,
      w,
      h,
      rectRadius: 0.05,
      fill: { color: fill },
      line: { color: C.teal, width: 1.3 },
      objectName: `benchb:${item.slide_id}:node:${node.node_id}`,
    });
    addText(
      slide,
      node.label,
      {
        x: x + 0.06,
        y: y + 0.03,
        w: Math.max(0.1, w - 0.12),
        h: Math.max(0.1, h - 0.06),
        fontFace: font,
        fontSize: 13.5,
        bold: true,
        color: C.ink,
        align: "center",
        valign: "mid",
        margin: 0.02,
        fit: "shrink",
      },
      `benchb:${item.slide_id}:node-label:${node.node_id}`,
    );
  });
  if (diagram.metadata?.evidence_status) {
    addText(
      slide,
      `Evidence status: ${diagram.metadata.evidence_status.replaceAll("_", " ")}`,
      {
        x: 1.35,
        y: 6.46,
        w: 6.5,
        h: 0.28,
        fontFace: font,
        fontSize: 11,
        bold: true,
        color: C.coral,
        margin: 0,
      },
      `benchb:${item.slide_id}:evidence-status`,
    );
  }
}

function addComparison(slide, item) {
  const cards = [
    {
      x: 1.25,
      color: C.paleTeal,
      border: C.teal,
      label: "Evidence-preserving",
      body: "Keep the study design, uncertainty, source wording, and explicit limitations.",
    },
    {
      x: 6.95,
      color: "FCEBE7",
      border: C.coral,
      label: "Not permitted",
      body: "Do not upgrade association or computational inference into causal, mechanistic, or clinical proof.",
    },
  ];
  cards.forEach((card, index) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: card.x,
      y: 2.08,
      w: 5.25,
      h: 3.78,
      rectRadius: 0.07,
      fill: { color: card.color },
      line: { color: card.border, width: 1.5 },
      objectName: `benchb:${item.slide_id}:guardrail:${index + 1}`,
    });
    addText(
      slide,
      card.label,
      {
        x: card.x + 0.35,
        y: 2.55,
        w: 4.5,
        h: 0.45,
        fontFace: font,
        fontSize: 22,
        bold: true,
        color: card.border,
        margin: 0,
      },
      `benchb:${item.slide_id}:guardrail-label:${index + 1}`,
    );
    addText(
      slide,
      card.body,
      {
        x: card.x + 0.35,
        y: 3.35,
        w: 4.5,
        h: 1.5,
        fontFace: font,
        fontSize: 17,
        color: C.ink,
        margin: 0,
        fit: "shrink",
        valign: "mid",
      },
      `benchb:${item.slide_id}:guardrail-body:${index + 1}`,
    );
  });
}

function addStatement(slide, item) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1.35,
    y: 2.08,
    w: 10.65,
    h: 3.85,
    rectRadius: 0.08,
    fill: { color: C.white },
    line: { color: C.line, width: 1.2 },
    shadow: { type: "outer", color: "AAB8C2", opacity: 0.14, blur: 1, angle: 45, distance: 1 },
    objectName: `benchb:${item.slide_id}:message-card`,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.35,
    y: 2.08,
    w: 0.18,
    h: 3.85,
    fill: { color: item.slide_role === "conclusion" ? C.amber : C.teal },
    line: { color: C.white, transparency: 100 },
    objectName: `benchb:${item.slide_id}:message-accent`,
  });
  addText(
    slide,
    item.single_key_message,
    {
      x: 2.05,
      y: 2.62,
      w: 9.1,
      h: 2.35,
      fontFace: font,
      fontSize: 24,
      bold: true,
      color: C.ink,
      align: "center",
      valign: "mid",
      margin: 0.08,
      fit: "shrink",
    },
    `benchb:${item.slide_id}:message`,
  );
  addText(
    slide,
    String(item.slide_role || "content").replaceAll("_", " ").toUpperCase(),
    {
      x: 9.05,
      y: 5.35,
      w: 2.45,
      h: 0.24,
      fontFace: font,
      fontSize: 9.5,
      bold: true,
      color: C.muted,
      align: "right",
      margin: 0,
      charSpacing: 0.8,
    },
    `benchb:${item.slide_id}:role`,
  );
}

for (const [index, item] of spec.slides.entries()) {
  const slide = pptx.addSlide();
  if (index === 0 || item.slide_role === "cover") {
    addCover(slide, item);
  } else {
    addHeader(slide, item, index);
    if (item.diagram_spec) {
      addDiagram(slide, item);
    } else if (item.chart_data) {
      addChart(slide, item);
    } else if (item.visual_type === "source_map") {
      addSourceMap(slide, item);
    } else if (item.visual_type === "comparison") {
      addComparison(slide, item);
    } else {
      addStatement(slide, item);
    }
    addFooter(slide, item);
  }
  if (typeof slide.addNotes === "function") {
    slide.addNotes(noteText(item));
  }
}

fs.mkdirSync(path.dirname(outputPptx), { recursive: true });
fs.mkdirSync(auditRoot, { recursive: true });
fs.writeFileSync(
  path.join(auditRoot, "arm_b_behavior_contract.json"),
  JSON.stringify(
    {
      schema_version: "1.0",
      arm: "B",
      status: "BENCHMARK_ONLY",
      external_code_executed: false,
      external_services_used: false,
      network_access: false,
      source_ir_hash: spec.deck_ir?.canonical_hash || "NOT_PROVIDED",
      behaviors: [
        "action-title hierarchy",
        "one dominant exhibit",
        "slide-role layout variation",
        "restrained density",
        "editable native PowerPoint objects",
      ],
      limitation:
        "This arm is not an independent scientific pipeline; it consumes the canonical source-bound Deck IR solely to isolate presentation behavior.",
    },
    null,
    2,
  ) + "\n",
  "utf8",
);

await pptx.writeFile({ fileName: outputPptx });
console.log(`ARM_B_SLIDES=${spec.slides.length}`);
console.log(`ARM_B_PPTX=${outputPptx}`);
