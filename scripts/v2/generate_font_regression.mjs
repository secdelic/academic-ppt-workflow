import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");
const output = process.argv[2];
if (!output) throw new Error("Output PPTX path is required");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow";
pptx.subject = "Font fallback regression fixture";
pptx.title = "Generic font fallback regression";
pptx.lang = "zh-CN";

const fonts = [
  ["APTOS", "Aptos"],
  ["ARIAL", "Arial"],
  ["MICROSOFT_YAHEI", "Microsoft YaHei"],
  ["NOTO_SANS_CJK", "Noto Sans CJK SC"],
  ["MISSING_PREFERRED", "AWF Deliberately Missing Font"],
];
fonts.forEach(([key, font], index) => {
  const slide = pptx.addSlide();
  slide.background = { color: "F7FAFC" };
  slide.addText(`长中文标题与 Mixed English title 保持在标题安全区 — ${key}`, {
    x: 0.74, y: 0.42, w: 11.86, h: 1.12,
    fontFace: font, fontSize: 32, bold: true, color: "102A43",
    margin: 0, objectName: `awf:font-${key}:title`,
  });
  slide.addText(
    "这是用于验证 PowerPoint 实际字体替换后文本边界的合成段落。"
      + "正文保持 18 pt 以上，不使用自动无限缩小，也不得进入页脚安全区。",
    {
      x: 0.80, y: 1.78, w: 11.73, h: 2.30,
      fontFace: font, fontSize: 20, color: "102A43",
      margin: 0.05,
      objectName: `awf:font-${key}:body`,
    },
  );
  slide.addShape(pptx.ShapeType.line, {
    x: 0.74, y: 6.76, w: 11.84, h: 0,
    line: { color: "D9E2EC", width: 1 },
    objectName: `awf:font-${key}:footer-line`,
  });
  slide.addText("来源：合成字体回归夹具", {
    x: 0.74, y: 6.84, w: 10.93, h: 0.37,
    fontFace: font, fontSize: 10, color: "627D98",
    margin: 0, objectName: `awf:font-${key}:sources`,
  });
  slide.addText(String(index + 1), {
    x: 12.00, y: 6.86, w: 0.73, h: 0.41,
    fontFace: font, fontSize: 11, color: "627D98",
    align: "right", margin: 0,
    objectName: `awf:font-${key}:page-number`,
  });
});
fs.mkdirSync(path.dirname(output), { recursive: true });
await pptx.writeFile({ fileName: output });
console.log(`FONT_REGRESSION_PPTX=${output}`);
