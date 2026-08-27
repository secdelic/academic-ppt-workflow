"use strict";

const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const [baselineSpecPath, candidateSpecPath, baselineOutput, candidateOutput, manifestOutput] = process.argv.slice(2);
if (!baselineSpecPath || !candidateSpecPath || !baselineOutput || !candidateOutput || !manifestOutput) {
  throw new Error(
    "Usage: generate_synthetic_decks.cjs <baseline-spec.json> <candidate-spec.json> " +
    "<baseline.pptx> <candidate.pptx> <manifest.json>"
  );
}

function loadSlides(specPath) {
  const payload = JSON.parse(fs.readFileSync(specPath, "utf8"));
  if (!Array.isArray(payload.slides)) throw new Error(`slides array missing: ${path.basename(specPath)}`);
  return payload.slides;
}

function defineMasters(pptx) {
  const common = [
    {
      text: {
        text: "SYNTHETIC PERFORMANCE FIXTURE",
        options: { x: 0.65, y: 7.08, w: 5.5, h: 0.18, fontFace: "Arial", fontSize: 9, color: "64748B", margin: 0 },
      },
    },
  ];
  pptx.defineSlideMaster({
    title: "SYN_CONTENT",
    background: { color: "F8FAFC" },
    objects: common,
    slideNumber: { x: 12.15, y: 7.03, w: 0.55, h: 0.2, fontFace: "Arial", fontSize: 9, color: "64748B", align: "right", margin: 0 },
  });
  pptx.defineSlideMaster({
    title: "SYN_DIVIDER",
    background: { color: "172554" },
    objects: [
      {
        text: {
          text: "SYNTHETIC PERFORMANCE FIXTURE",
          options: { x: 0.65, y: 7.08, w: 5.5, h: 0.18, fontFace: "Arial", fontSize: 9, color: "BFDBFE", margin: 0 },
        },
      },
    ],
    slideNumber: { x: 12.15, y: 7.03, w: 0.55, h: 0.2, fontFace: "Arial", fontSize: 9, color: "BFDBFE", align: "right", margin: 0 },
  });
}

async function buildDeck(slides, outputPath, title) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Academic PPT Workflow";
  pptx.company = "Local synthetic regression";
  pptx.subject = "Anonymous fast-enhance performance fixture";
  pptx.title = title;
  pptx.lang = "en-US";
  pptx.theme = { headFontFace: "Arial", bodyFontFace: "Arial", lang: "en-US" };
  defineMasters(pptx);
  for (const spec of slides) {
    const divider = spec.layout === "divider";
    const slide = pptx.addSlide(divider ? "SYN_DIVIDER" : "SYN_CONTENT");
    slide.background = { color: divider ? "172554" : "F8FAFC" };
    slide.addText(String(spec.title), {
      x: 0.75, y: divider ? 2.25 : 0.62, w: 11.8, h: divider ? 1.15 : 0.72,
      fontFace: "Arial", fontSize: divider ? 32 : 28, bold: true,
      color: divider ? "F8FAFC" : "172554", margin: 0, breakLine: false,
    });
    slide.addText(String(spec.key_message), {
      x: 0.78, y: divider ? 3.55 : 2.05, w: 10.9, h: 1.2,
      fontFace: "Arial", fontSize: 18, color: divider ? "BFDBFE" : "334155",
      margin: 0, valign: "mid",
    });
    slide.addText(`${spec.slide_id}  r${spec.slide_revision}`, {
      x: 0.78, y: 6.62, w: 4.8, h: 0.2, fontFace: "Arial", fontSize: 9,
      color: divider ? "93C5FD" : "94A3B8", margin: 0,
    });
    if (typeof slide.addNotes === "function") {
      slide.addNotes(
        `[Sources]\nSynthetic local fixture only.\n[Slide-ID]\n${spec.slide_id}` +
        `\n[Slide-Revision]\n${spec.slide_revision}\n[Content-Hash]\n${spec.content_hash}`
      );
    }
  }
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await pptx.writeFile({ fileName: outputPath });
}

(async () => {
  const baselineSlides = loadSlides(baselineSpecPath);
  const candidateSlides = loadSlides(candidateSpecPath);
  await buildDeck(baselineSlides, baselineOutput, "Synthetic 21-slide baseline");
  await buildDeck(candidateSlides, candidateOutput, "Synthetic seven-slide delta");
  const manifest = {
    schema_version: "2.5.1",
    synthetic_only: true,
    baseline_slide_count: baselineSlides.length,
    candidate_slide_count: candidateSlides.length,
    baseline_slide_ids: baselineSlides.map((slide) => slide.slide_id),
    candidate_slide_ids: candidateSlides.map((slide) => slide.slide_id),
    external_service_usage: 0,
  };
  fs.mkdirSync(path.dirname(manifestOutput), { recursive: true });
  fs.writeFileSync(manifestOutput, JSON.stringify(manifest, null, 2) + "\n", "utf8");
})().catch((error) => {
  process.stderr.write(String(error && error.stack ? error.stack : error) + "\n");
  process.exitCode = 1;
});
