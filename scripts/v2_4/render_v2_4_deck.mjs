import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");
const imageSizeModule = require("image-size");
const imageSize = imageSizeModule.imageSize || imageSizeModule.default || imageSizeModule;

const [graphPath, outputPath] = process.argv.slice(2);
if (!graphPath || !outputPath) throw new Error("Usage: render_v2_4_deck.mjs <canonical_object_graph.json> <output.pptx>");
const graph = JSON.parse(fs.readFileSync(graphPath, "utf8"));

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow v2.4";
pptx.subject = "Synthetic presentation design and semantic chart grammar validation";
pptx.company = "Academic PPT Workflow";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};

const S = pptx.ShapeType;
const W = 13.333;
const H = 7.5;
const TYPO = { deck: 50, title: 35, sub: 24, body: 18, detail: 16, chart: 15, footer: 10 };
const SHARED = { canvas: "F6F9FB", ink: "17324D", muted: "627D98", white: "FFFFFF", pale: "EAF1F5", warning: "D97706", danger: "B23A48", line: "B9CAD6" };
const PROFILES = {
  clinical_methods: { primary: "17324D", accent: "178A8A", secondary: "E76F51", soft: "DDF3F0", soft2: "FCE8E2" },
  bioinformatics: { primary: "213A5C", accent: "2A9D8F", secondary: "D95F59", negative: "3B82C4", soft: "DDF3F0", soft2: "FBE4E2" },
  evidence_synthesis: { primary: "263B57", accent: "6B7F9E", secondary: "C74B50", soft: "E7EDF4", soft2: "F5E5E5" },
  clinical_protocol: { primary: "1D3B46", accent: "3A8D8A", secondary: "6C8EBF", soft: "DDEFEF", soft2: "E5ECF7" },
};
const ROLE_ZH = { cover: "", context: "问题", method: "方法", result: "结果", limitation: "局限性", conclusion: "结论", planning: "计划", appendix: "附录" };
const FAMILIES = ["cover", "section_divider", "hero_insight", "chart_led", "figure_with_callout", "split_screen", "comparison_matrix", "process_branch", "evidence_ladder", "risk_heatmap", "conclusion_synthesis", "appendix_audit"];

// Canonical source values stay unchanged in VisualSpec and audit artifacts.
// This dictionary is only used at the audience-facing render boundary.
const ZH_LABELS = {
  "Control": "对照", "SIMD": "SIMD",
  "T cell": "T细胞", "B cell": "B细胞", "NK cell": "NK细胞", "Monocyte": "单核细胞",
  "Macrophage": "巨噬细胞", "Dendritic": "树突状细胞", "Endothelial": "内皮细胞",
  "Fibroblast": "成纤维细胞", "Pericyte": "周细胞", "Cardiomyocyte": "心肌细胞",
  "Biventricular dysfunction phenotype": "双心室功能障碍表型",
  "AKI progression by 72 h": "72小时内AKI进展", "SOFA change at 48 h": "48小时SOFA变化",
  "28-day mortality": "28天死亡", "Age": "年龄", "SOFA score": "SOFA评分", "Lactate": "乳酸",
  "Norepinephrine-equivalent dose": "去甲肾上腺素等效剂量", "Mechanical ventilation": "机械通气",
  "TTE image quality": "TTE图像质量",
  "Age >=18 years": "年龄≥18岁", "Suspected or confirmed infection": "疑似或确诊感染",
  "SOFA increase >=2": "SOFA增加≥2分", "TTE feasible within 6 h of enrollment": "入组后6小时内可完成TTE",
  "Pregnancy": "妊娠", "Known severe chronic pulmonary hypertension": "已知重度慢性肺动脉高压",
  "Mechanical circulatory support before enrollment": "入组前已接受机械循环支持", "Expected survival <24 h": "预期生存<24小时",
  "Screening": "筛选", "Eligibility, consent, baseline history": "资格、知情同意与基线病史",
  "Hemodynamics-1": "血流动力学-1", "MAP, HR, CVP, lactate, vasopressor dose": "MAP、HR、CVP、乳酸与升压药剂量",
  "LV/RV function, RV-PA coupling, IVC": "LV/RV功能、RV-PA耦联与IVC", "Repeat focused TTE": "重复重点TTE",
  "Outcome-1": "结局-1", "SOFA trajectory and organ support": "SOFA轨迹与器官支持",
  "Outcome-2": "结局-2", "AKI progression and RRT": "AKI进展与RRT", "Follow-up": "随访",
  "ICU and 28-day mortality": "ICU/28天死亡", "Primary endpoint": "主要终点", "Secondary": "次要",
  "Target enrollment": "目标入组", "Required analyzable sample": "所需可分析样本",
  "Loss/non-evaluable rate": "失访/不可评估比例", "Primary anticipated event rate": "主要结局预期事件率",
  "Minimum detectable adjusted odds ratio": "最小可检出的调整OR", "Two-sided alpha": "双侧α", "Power": "把握度",
  "Participants": "参与者", "Proportion": "比例", "Probability": "概率", "Categorical": "分类变量",
  "Binary": "二分类", "Continuous": "连续变量", "Ordinal": "有序分类",
  "TTE at 0-6 h": "0–6小时TTE", "KDIGO stage": "KDIGO分期", "Baseline": "基线",
  "Before TTE": "TTE前", "At acquisition": "采集时", "0-6 h": "0–6小时", "0 h": "0小时",
  "24 h +/- 4 h": "24小时±4小时", "48 h": "48小时", "72 h": "72小时", "Day 28": "第28天",
  "Delayed TTE": "TTE延迟", "Dedicated on-call sonographer": "设置专职值班超声医师",
  "Inter-observer variability": "观察者间差异", "Core-lab training and duplicate reads": "核心实验室培训与重复判读",
  "Missing 72-h creatinine": "缺失72小时肌酐", "Automated follow-up query": "自动随访查询",
  "Confounding by disease severity": "疾病严重程度混杂", "Prespecified adjustment and sensitivity analysis": "预设调整与敏感性分析",
  "Center-level heterogeneity": "中心间异质性", "Multilevel sensitivity analysis": "多层敏感性分析",
  "Enrollment imbalance": "入组不均衡", "Monthly site monitoring": "每月中心监测",
  "Protocol finalization": "方案定稿", "Ethics and registration": "伦理与注册", "Site training": "中心培训",
  "Pilot enrollment": "试点入组", "Full enrollment": "全面入组", "Database lock": "数据库锁定",
  "Analysis and reporting": "分析与报告", "Planned": "计划中", "Required": "必需", "Exclude": "排除",
  "transfuse within 24 h vs no transfusion within 24 h": "24小时内输血 vs 24小时内未输血",
  "transfuse_within_24h": "24小时内输血", "no_transfusion_within_24h": "24小时内未输血",
  "transfusion_24h": "24小时内输血", "no_transfusion_24h": "24小时内未输血",
  "transfusion_vs_no_transfusion": "输血 vs 未输血", "canonical": "规范估计", "supportive": "支持性估计",
  "Weighted analysis": "加权分析", "Per-protocol analysis": "符合方案分析", "Complete-case analysis": "完整病例分析",
  "Grace period 12 h": "宽限期12小时", "Grace period 24 h": "宽限期24小时",
  "Weight truncation 1st/99th": "权重截尾1%/99%", "Weight truncation 5th/95th": "权重截尾5%/95%",
  "Index hemoglobin": "基线血红蛋白", "Vasopressor use": "升压药使用", "Site": "研究中心",
  "Input cells": "输入细胞", "After gene-count filter": "基因数过滤后",
  "After mitochondrial filter": "线粒体比例过滤后", "After doublet removal": "双细胞去除后",
  "Prospective studies": "前瞻性研究", "Retrospective studies": "回顾性研究",
  "RV dysfunction by TAPSE": "按TAPSE定义右心功能障碍",
  "RV dysfunction by qualitative assessment": "按定性评估定义右心功能障碍",
  "Septic shock only": "仅脓毒性休克", "Mixed sepsis severity": "混合脓毒症严重度",
  "Selection": "选择偏倚", "Measurement": "测量偏倚", "Confounding": "混杂偏倚",
  "Missing data": "缺失数据", "Reporting": "报告偏倚",
  "Short-term mortality": "短期死亡", "ICU mortality": "ICU死亡",
  "Mechanical ventilation duration": "机械通气时长",
  "Serious": "严重", "Not serious": "不严重", "Low": "低", "Very low": "极低",
  "Moderate": "中等", "None": "无", "Screening / time zero": "筛选/时间零点",
};

function zhLabel(value) {
  return String(value ?? "").split("\n").map(part => {
    const normalized = part.trim();
    return ZH_LABELS[normalized] || normalized;
  }).join("\n");
}

function profile(slideSpec) { return PROFILES[slideSpec.design_profile] || PROFILES.clinical_methods; }
function masterName(family) { return `V24_${family}`; }
for (const family of FAMILIES) {
  const objects = family === "cover" ? [] : [
    { rect: { x: 0, y: 7.16, w: W, h: 0.34, fill: { color: "17324D" }, line: { color: "17324D" } } },
    { line: { x: 0.72, y: 0.30, w: 0.44, h: 0, line: { color: "2A9D8F", width: 4 } } },
  ];
  pptx.defineSlideMaster({
    title: masterName(family),
    background: { color: family === "cover" ? "17324D" : SHARED.canvas },
    objects,
    // PptxGenJS passes unsupported slide-number-only valign/margin options
    // through to the master field in a form PowerPoint rejects as corrupted.
    // Keep this contract to the supported properties used by the stable v2.3
    // backend; position and height still provide the 10 pt footer-safe page id.
    slideNumber: family === "cover" ? undefined : { x: 12.40, y: 7.17, w: 0.43, h: 0.24, color: SHARED.white, fontFace: "Arial", fontSize: 10, align: "right" },
  });
}

function addText(slide, text, x, y, w, h, size = TYPO.body, options = {}) {
  slide.addText(String(text), {
    x, y, w, h,
    fontFace: options.fontFace || "Microsoft YaHei",
    fontSize: size,
    bold: Boolean(options.bold),
    color: options.color || SHARED.ink,
    margin: options.margin ?? 0,
    align: options.align || "left",
    valign: options.valign || "mid",
    rotate: options.rotate,
    breakLine: false,
    name: options.name,
  });
}

function box(slide, x, y, w, h, fill = SHARED.white, line = SHARED.line, radius = 0.04, name = undefined) {
  slide.addShape(S.roundRect, { x, y, w, h, rectRadius: radius, fill: { color: fill }, line: { color: line, width: 1 }, name });
}

function line(slide, x, y, w, h, color = SHARED.line, width = 1, options = {}) {
  // OOXML extents must be non-negative for PowerPoint compatibility.  Express
  // reverse-direction edges through shape flips instead of negative cx/cy;
  // LibreOffice accepts negative extents, but PowerPoint may reject the deck.
  const endX = x + w, endY = y + h;
  const flipH = endX < x, flipV = endY < y;
  slide.addShape(S.line, {
    x: Math.min(x, endX), y: Math.min(y, endY),
    w: Math.abs(w), h: Math.abs(h), flipH, flipV,
    line: { color, width, dash: options.dash, beginArrowType: options.beginArrowType, endArrowType: options.endArrowType },
    name: options.name,
  });
}

function ellipse(slide, x, y, w, h, fill, lineColor = fill, name = undefined) {
  slide.addShape(S.ellipse, { x, y, w, h, fill: { color: fill }, line: { color: lineColor, width: 1 }, name });
}

function rowName(visual, index) {
  const key = (visual.expected_row_keys || [])[index];
  if (!key) throw new Error(`Missing expected row key for ${visual.visual_id} row ${index}`);
  return `${visual.visual_id}:row:${key}`;
}

function titleDisplayWidth(text) {
  return [...String(text)].reduce((total, character) => total + (/[^\u0000-\u00ff]/.test(character) ? 2 : 1), 0);
}

function wrapTitleText(text, maxWidth = 38) {
  const value = String(text);
  if (titleDisplayWidth(value) <= maxWidth) return value;
  const characters = [...value];
  const target = titleDisplayWidth(value) / 2;
  let width = 0, best = -1, bestDistance = Number.POSITIVE_INFINITY;
  characters.forEach((character, index) => {
    width += /[^\u0000-\u00ff]/.test(character) ? 2 : 1;
    if (/[：；，、—:;,-]/.test(character) || character === " ") {
      const distance = Math.abs(width - target);
      if (distance < bestDistance) { best = index + 1; bestDistance = distance; }
    }
  });
  if (best < 0) {
    width = 0;
    best = Math.max(1, Math.floor(characters.length / 2));
    for (let index = 0; index < characters.length; index++) {
      width += /[^\u0000-\u00ff]/.test(characters[index]) ? 2 : 1;
      if (width >= target) { best = index + 1; break; }
    }
  }
  // Never split an ASCII scientific token or decimal (for example 4.3,
  // 95%CI, TTE, or pseudobulk) merely to reach the visual midpoint.
  const tokenChar = character => /[A-Za-z0-9._%+\-]/.test(character || "");
  while (best < characters.length && tokenChar(characters[best - 1]) && tokenChar(characters[best])) best += 1;
  return `${characters.slice(0, best).join("").trim()}\n${characters.slice(best).join("").trim()}`;
}

function title(slide, spec) {
  const wrapped = wrapTitleText(spec.title);
  const twoLines = wrapped.includes("\n");
  addText(slide, wrapped, 0.74, 0.27, 11.85, 0.90, twoLines ? 30 : TYPO.title, { bold: true, color: profile(spec).primary, name: `${spec.slide_id}:title`, valign: "mid" });
  const label = ROLE_ZH[spec.slide_role] || spec.slide_role;
  const reservesUpperBand = (spec.visual_specs || []).some(visual => visual.visual_type === "assessment_timeline");
  // A two-line title already consumes the navigation-label row.  Omitting the
  // redundant role label preserves a measured gutter without shrinking the
  // title below the 30 pt contract or pushing content into its safe zone.
  if (label && !twoLines && !reservesUpperBand) addText(slide, label, 0.76, 1.17, 1.55, 0.24, TYPO.chart, { bold: true, color: profile(spec).accent, name: `${spec.slide_id}:section` });
}

function footer(slide, spec) {
  const labels = [...new Set((spec.source_bindings || []).map(binding => binding.short_label_zh || "输入材料"))];
  const visible = labels.slice(0, 3);
  const suffix = labels.length > 3 ? ` 等${labels.length}项` : "";
  if (visible.length) addText(slide, `来源：${visible.join("、")}${suffix}`, 0.78, 7.17, 10.8, 0.22, TYPO.footer, { color: SHARED.white, name: `${spec.slide_id}:footer` });
}

function speakerNotes(spec) {
  const lines = ["[Sources]"];
  for (const binding of spec.source_bindings || []) {
    lines.push(`${binding.source_id} | ${binding.source_file} | ${binding.source_location} | fields=${(binding.fields_used || []).join(",")} | filter=${binding.row_filter} | aggregation=${binding.aggregation}`);
  }
  lines.push("", "[Claims]");
  for (const claim of spec.claim_bindings || []) lines.push(`${claim.claim_id} | ${claim.claim_text} | ${claim.wording_boundary}`);
  lines.push("", "[Conflicts]");
  for (const item of spec.conflict_bindings || []) lines.push(`${item.conflict_id} | ${item.conflicting_value} -> ${item.canonical_value}`);
  lines.push("", "[Unresolved]");
  for (const item of spec.unresolved_bindings || []) lines.push(`${item.marker} | INFORMATION_REQUIRED`);
  return lines.join("\n");
}

function imageContain(slide, file, x, y, w, h, name) {
  const dimensions = imageSize(fs.readFileSync(file));
  const scale = Math.min(w / dimensions.width, h / dimensions.height);
  const iw = dimensions.width * scale;
  const ih = dimensions.height * scale;
  slide.addImage({ path: file, x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih, name });
}

function addAxis(slide, x, y, w, ticks, scale, label, p, format = value => String(value)) {
  line(slide, x, y, w, 0, SHARED.muted, 1.1);
  for (const tick of ticks) {
    const tx = scale(tick);
    line(slide, tx, y - 0.05, 0, 0.10, SHARED.muted, 0.8);
    addText(slide, format(tick), tx - 0.32, y + 0.07, 0.64, 0.22, TYPO.chart, { align: "center", color: SHARED.muted });
  }
  if (label) addText(slide, label, x, y + 0.36, w, 0.24, TYPO.chart, { align: "center", color: SHARED.muted });
}

function linearScale(domainMin, domainMax, rangeMin, rangeMax) {
  const span = domainMax - domainMin || 1;
  return value => rangeMin + (Number(value) - domainMin) / span * (rangeMax - rangeMin);
}

function logScale(domainMin, domainMax, rangeMin, rangeMax) {
  const lo = Math.log(domainMin), hi = Math.log(domainMax);
  return value => rangeMin + (Math.log(Number(value)) - lo) / (hi - lo) * (rangeMax - rangeMin);
}

function renderCover(slide, spec, projectKind) {
  const p = profile(spec);
  slide.background = { color: p.primary };
  slide.addShape(S.rect, { x: 0, y: 0, w: W, h: H, fill: { color: p.primary }, line: { color: p.primary } });
  slide.addShape(S.rect, { x: 0.82, y: 1.05, w: 0.12, h: 4.85, fill: { color: p.accent }, line: { color: p.accent } });
  addText(slide, spec.title, 1.28, 1.46, 10.9, 1.22, TYPO.deck, { bold: true, color: SHARED.white, name: `${spec.slide_id}:deck-title` });
  const subtitle = { target_trial: "临床方法 / 目标试验模拟", single_cell: "生物信息学 / 单细胞分析", meta_analysis: "证据综合 / Meta分析", protocol: "临床研究方案 / 前瞻性TTE" }[projectKind];
  addText(slide, subtitle, 1.30, 3.10, 8.7, 0.42, TYPO.sub, { color: "CBE4E6" });
  addText(slide, "完全合成回归测试材料｜不构成真实临床证据", 1.30, 5.45, 10.6, 0.34, TYPO.detail, { bold: true, color: "F2BE6A" });
}

function renderHeroQuestion(slide, visual, spec) {
  const p = profile(spec), data = visual.payload;
  addText(slide, data.question, 0.86, 1.52, 11.3, 1.05, 30, { bold: true, color: p.primary, name: rowName(visual, 0) });
  line(slide, 0.90, 2.86, 11.4, 0, p.accent, 2.2);
  const entries = Object.entries(data).filter(([key]) => key !== "question");
  entries.forEach(([key, value], index) => {
    const x = 0.92 + index * (11.4 / entries.length);
    addText(slide, ({ gap: "知识缺口", objective: "目标", population: "人群", exposure: "暴露", outcome: "结局", design: "设计", sites: "站点", boundary: "边界" }[key] || key), x, 3.16, 11.0 / entries.length, 0.30, TYPO.detail, { bold: true, color: p.accent });
    addText(slide, value, x, 3.62, 10.7 / entries.length, 1.34, TYPO.body, { color: SHARED.ink, valign: "top" });
  });
}

function renderMatrix(slide, visual, spec) {
  const items = visual.payload.items || [];
  const cols = items.length > 6 ? 4 : 2;
  const rows = Math.ceil(items.length / cols);
  const gapX = 0.18, gapY = 0.20, x0 = 0.82, y0 = 1.45, totalW = 11.72, totalH = 5.35;
  const w = (totalW - gapX * (cols - 1)) / cols, h = (totalH - gapY * (rows - 1)) / rows;
  items.forEach((item, index) => {
    const col = index % cols, row = Math.floor(index / cols), x = x0 + col * (w + gapX), y = y0 + row * (h + gapY);
    box(slide, x, y, w, h, row % 2 ? SHARED.white : profile(spec).soft, SHARED.line, 0.03, rowName(visual, index));
    addText(slide, item[0], x + 0.16, y + 0.12, w - 0.32, 0.34, TYPO.detail, { bold: true, color: profile(spec).primary });
    addText(slide, item[1], x + 0.16, y + 0.58, w - 0.32, h - 0.70, TYPO.detail, { color: SHARED.muted, valign: "top" });
    if (item[2]) {
      const severity = String(item[2]).toLowerCase();
      const badgeColor = severity.includes("high") || severity.includes("高") ? SHARED.warning : profile(spec).accent;
      addText(slide, item[2], x + w - 0.98, y + 0.12, 0.80, 0.24, TYPO.chart, { bold: true, color: badgeColor, align: "right" });
    }
  });
}

function renderBranch(slide, visual, spec) {
  const nodes = visual.payload.nodes || [];
  const branches = visual.payload.branches || nodes.slice(1).map((_, index) => [index, index + 1]);
  const count = nodes.length;
  const nodeWidth = count <= 5 ? Math.min(1.75, 9.80 / Math.max(1, count)) : count === 6 ? 1.55 : 1.48;
  const positions = nodes.map((_, index) => {
    if (count <= 6) return { x: 0.82 + index * ((11.70 - nodeWidth) / Math.max(1, count - 1)), y: 3.1 };
    if (index < 4) return { x: 0.95 + index * 3.25, y: 2.35 };
    return { x: 2.55 + (index - 4) * 3.25, y: 4.65 };
  });
  for (const [from, to] of branches) {
    const a = positions[from], b = positions[to];
    line(slide, a.x + nodeWidth / 2, a.y + 0.36, b.x - a.x - nodeWidth / 2, b.y - a.y, profile(spec).accent, 1.8, { endArrowType: "triangle", name: `${visual.visual_id}:edge:${from}-${to}` });
  }
  nodes.forEach((label, index) => {
    const pos = positions[index];
    box(slide, pos.x, pos.y, nodeWidth, 0.78, index % 2 ? SHARED.white : profile(spec).soft, profile(spec).accent, 0.03, rowName(visual, index));
    addText(slide, zhLabel(label), pos.x + 0.10, pos.y + 0.10, nodeWidth - 0.20, 0.56, TYPO.detail, { bold: true, align: "center" });
  });
  if (visual.payload.warning) {
    box(slide, 3.5, 5.82, 6.35, 0.68, "FFF4E5", SHARED.warning, 0.03);
    addText(slide, visual.payload.warning, 3.70, 5.98, 5.95, 0.32, TYPO.body, { bold: true, color: SHARED.warning, align: "center" });
  }
}

function renderTimeWindow(slide, visual, spec) {
  const events = visual.payload.events || [];
  const p = profile(spec);
  line(slide, 1.05, 3.50, 8.75, 0, p.accent, 3);
  line(slide, 10.0, 3.50, 1.90, 0, p.accent, 3);
  addText(slide, "//", 9.72, 3.34, 0.35, 0.34, TYPO.sub, { bold: true, color: SHARED.muted, align: "center" });
  events.forEach((event, index) => {
    const x = Number(event.hour) <= 24 ? 1.05 + Number(event.hour) / 24 * 8.75 : 11.55;
    ellipse(slide, x - 0.10, 3.40, 0.20, 0.20, index === 2 ? p.secondary : p.accent, undefined, rowName(visual, index));
    addText(slide, zhLabel(event.label), x - 0.85, index % 2 ? 3.82 : 2.43, 1.70, 0.62, TYPO.detail, { bold: true, align: "center" });
  });
  addText(slide, "0–24小时：策略宽限期", 1.10, 5.15, 4.3, 0.35, TYPO.body, { bold: true, color: p.primary });
  addText(slide, "轴断点", 9.30, 5.15, 1.2, 0.35, TYPO.detail, { color: SHARED.muted, align: "center" });
  addText(slide, "28天结局", 10.65, 5.15, 1.9, 0.35, TYPO.body, { bold: true, color: p.primary, align: "center" });
}

function renderLove(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const x0 = 4.35, x1 = 11.45, y0 = 1.72, rowH = 0.58;
  const maxValue = Math.max(0.30, ...rows.flatMap(row => [Number(row.absolute_smd_unweighted), Number(row.absolute_smd_weighted)]));
  const scale = linearScale(0, Math.ceil(maxValue * 10) / 10, x0, x1);
  addAxis(slide, x0, 6.35, x1 - x0, [0, 0.1, 0.2, 0.3], scale, "绝对标准化均差（|SMD|）", p, value => value.toFixed(1));
  line(slide, scale(0.10), 1.48, 0, 4.65, p.secondary, 1.3, { dash: "dash" });
  rows.forEach((row, index) => {
    const y = y0 + index * rowH;
    addText(slide, zhLabel(row.covariate), 0.86, y - 0.05, 3.05, 0.30, TYPO.chart, { color: SHARED.ink });
    const before = Number(row.absolute_smd_unweighted), after = Number(row.absolute_smd_weighted);
    line(slide, scale(Math.min(before, after)), y + 0.12, Math.abs(scale(before) - scale(after)), 0, "A9BBC8", 1.2);
    ellipse(slide, scale(before) - 0.065, y + 0.055, 0.13, 0.13, p.secondary, p.secondary, rowName(visual, index));
    ellipse(slide, scale(after) - 0.065, y + 0.055, 0.13, 0.13, p.accent, p.accent);
    addText(slide, `${before.toFixed(2)} → ${after.toFixed(2)}`, 11.62, y - 0.05, 1.45, 0.30, TYPO.chart, { align: "right", color: SHARED.muted, fontFace: "Arial" });
  });
  ellipse(slide, 8.65, 1.27, 0.13, 0.13, p.secondary); addText(slide, "加权前", 8.85, 1.19, 0.95, 0.28, TYPO.chart, { color: SHARED.muted });
  ellipse(slide, 10.05, 1.27, 0.13, 0.13, p.accent); addText(slide, "加权后", 10.25, 1.19, 0.95, 0.28, TYPO.chart, { color: SHARED.muted });
  addText(slide, "阈值0.10", scale(0.10) - 0.42, 1.17, 0.84, 0.25, TYPO.chart, { color: p.secondary, align: "center" });
}

function renderHistogram(slide, visual, spec) {
  const bins = visual.payload.bins, summary = visual.payload.summary, p = profile(spec);
  const x0 = 1.02, yBase = 5.98, plotW = 8.35, plotH = 4.45;
  const maxCount = Math.max(...bins.map(bin => Number(bin.count)));
  const barW = plotW / bins.length;
  bins.forEach((bin, index) => {
    const h = Number(bin.count) / maxCount * plotH;
    slide.addShape(S.rect, { x: x0 + index * barW + 0.01, y: yBase - h, w: Math.max(0.03, barW - 0.02), h, fill: { color: p.accent }, line: { color: p.accent }, name: rowName(visual, index) });
  });
  line(slide, x0, yBase, plotW, 0, SHARED.muted, 1);
  addText(slide, "稳定权重", x0, 6.20, plotW, 0.25, TYPO.chart, { align: "center", color: SHARED.muted });
  box(slide, 9.78, 1.70, 2.65, 3.90, p.soft, p.accent, 0.03);
  const values = [
    ["记录数", summary.n],
    ["中位数", Number(summary.median).toFixed(2)],
    ["第99百分位", Number(summary.p99).toFixed(2)],
    ["最大值", Number(summary.max).toFixed(2)],
  ];
  values.forEach((entry, index) => {
    addText(slide, entry[0], 10.02, 1.96 + index * 0.84, 1.45, 0.28, TYPO.detail, { color: SHARED.muted });
    addText(slide, entry[1], 10.02, 2.28 + index * 0.84, 1.95, 0.38, TYPO.sub, { bold: true, color: p.primary });
  });
}

function renderAbsoluteRisk(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const differenceRow = rows.find(row => String(row.estimand || "").toLowerCase().includes("difference"));
  const strategyRows = rows.filter(row => row !== differenceRow);
  if (!differenceRow || strategyRows.length < 2) throw new Error(`Absolute-risk contract requires strategy risks and a risk difference: ${visual.visual_id}`);
  const entries = strategyRows.map((row, index) => ({ rowIndex: rows.indexOf(row), label: zhLabel(row.contrast || row.estimand), value: Number(row.estimate) * 100, color: index % 2 ? p.secondary : p.accent }));
  const rd = Number(differenceRow.estimate) * 100;
  const max = Math.max(10, Math.ceil(Math.max(...entries.map(entry => entry.value)) / 5) * 5 + 5);
  const scale = linearScale(0, max, 3.42, 8.60);
  entries.forEach((entry, index) => {
    const y = 2.15 + index * 1.52;
    addText(slide, entry.label, 0.92, y - 0.05, 2.45, 0.34, TYPO.body, { bold: true });
    slide.addShape(S.rect, { x: 3.42, y, w: Math.max(0.04, scale(entry.value) - 3.42), h: 0.48, fill: { color: entry.color }, line: { color: entry.color }, name: rowName(visual, entry.rowIndex) });
    addText(slide, `${entry.value.toFixed(1)}%`, scale(entry.value) + 0.12, y - 0.02, 1.15, 0.42, TYPO.sub, { bold: true, color: entry.color });
  });
  box(slide, 9.78, 1.78, 2.65, 3.55, p.soft, p.accent, 0.03, rowName(visual, rows.indexOf(differenceRow)));
  addText(slide, "风险差", 10.02, 2.03, 2.10, 0.36, TYPO.detail, { bold: true, color: p.accent });
  addText(slide, `${rd.toFixed(1)} pp`, 10.02, 2.62, 2.10, 0.62, 34, { bold: true, color: p.primary });
  addText(slide, `95%CI ${(Number(differenceRow.ci_low) * 100).toFixed(1)} 至 ${(Number(differenceRow.ci_high) * 100).toFixed(1)} pp`, 10.02, 3.55, 2.05, 0.62, TYPO.detail, { color: SHARED.muted });
  addText(slide, "合成估计｜不构成输血建议", 10.02, 4.52, 2.05, 0.38, TYPO.detail, { bold: true, color: SHARED.warning });
}

function renderEstimandComparison(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  rows.forEach((row, index) => {
    const x = index === 0 ? 0.90 : 6.78;
    box(slide, x, 1.65, 5.55, 4.70, index === 0 ? p.soft : SHARED.white, p.accent, 0.03, rowName(visual, index));
    const estimand = String(row.estimand || "");
    const heading = /hazard/i.test(estimand) ? "加权风险比（HR）" : /risk ratio/i.test(estimand) ? "风险比（RR）" : estimand;
    addText(slide, heading, x + 0.28, 1.98, 4.90, 0.42, TYPO.sub, { bold: true, color: p.primary });
    addText(slide, Number(row.estimate).toFixed(2), x + 0.28, 2.70, 2.20, 0.74, 38, { bold: true, color: index === 0 ? p.accent : p.secondary });
    addText(slide, `95%CI ${row.ci_low}–${row.ci_high}`, x + 0.28, 3.62, 3.70, 0.36, TYPO.body, { color: SHARED.muted });
    addText(slide, zhLabel(row.contrast || estimand), x + 0.28, 4.45, 4.70, 0.42, TYPO.body, { color: SHARED.ink });
    addText(slide, zhLabel(row.status || "已登记估计量"), x + 0.28, 5.35, 4.70, 0.36, TYPO.detail, { bold: true, color: index === 0 ? p.accent : SHARED.warning });
  });
}

function renderForest(slide, visual, spec, isSubgroup = false) {
  const rows = visual.payload.rows, p = profile(spec), log = visual.payload.scale === "log";
  const labels = rows.map(row => zhLabel(row[visual.label_field]));
  const estimates = rows.map(row => Number(row[visual.x_field]));
  const lows = rows.map(row => Number(row[visual.ci_low_field]));
  const highs = rows.map(row => Number(row[visual.ci_high_field]));
  const reference = Number(visual.reference_value);
  const x0 = 5.10, x1 = 10.65, plotTop = 1.57, plotBottom = visual.payload.pooled ? 5.78 : 5.95;
  let domainMin, domainMax, ticks, scale;
  if (log) {
    domainMin = Math.min(0.5, ...lows); domainMax = Math.max(6, ...highs);
    ticks = [0.5, 1, 2, 4, 8].filter(tick => tick >= domainMin && tick <= domainMax);
    scale = logScale(domainMin, domainMax, x0, x1);
  } else {
    domainMin = Math.min(-0.10, ...lows, reference); domainMax = Math.max(0.02, ...highs, reference);
    ticks = [-0.08, -0.04, 0, 0.04].filter(tick => tick >= domainMin && tick <= domainMax);
    scale = linearScale(domainMin, domainMax, x0, x1);
  }
  const rowH = (plotBottom - plotTop) / rows.length;
  line(slide, scale(reference), plotTop - 0.12, 0, plotBottom - plotTop + 0.25, SHARED.muted, 1, { dash: "dash" });
  rows.forEach((row, index) => {
    const y = plotTop + index * rowH;
    addText(slide, labels[index], 0.82, y - 0.02, 3.95, Math.max(0.22, rowH - 0.02), TYPO.chart, { color: SHARED.ink });
    line(slide, scale(lows[index]), y + rowH * 0.42, Math.max(0.02, scale(highs[index]) - scale(lows[index])), 0, p.accent, 1.6);
    ellipse(slide, scale(estimates[index]) - 0.055, y + rowH * 0.42 - 0.055, 0.11, 0.11, p.accent, undefined, rowName(visual, index));
    const interaction = isSubgroup ? `；交互P=${row.interaction_p}` : "";
    addText(slide, `${estimates[index].toFixed(2)}（${lows[index].toFixed(2)}–${highs[index].toFixed(2)}）${interaction}`, 10.82, y - 0.02, 2.05, Math.max(0.22, rowH - 0.02), TYPO.chart, { align: "right", color: SHARED.muted });
  });
  if (visual.payload.pooled) {
    const pooled = visual.payload.pooled, e = Number(pooled.estimate), low = Number(pooled.ci_low), high = Number(pooled.ci_high), y = 6.01;
    addText(slide, "随机效应合并", 0.82, y - 0.10, 3.95, 0.30, TYPO.chart, { bold: true, color: p.primary });
    const cx = scale(e), left = scale(low), right = scale(high);
    const diamondWidth = Math.max(0.20, right - left);
    slide.addShape(S.diamond, { x: cx - diamondWidth / 2, y: y - 0.10, w: diamondWidth, h: 0.28, fill: { color: p.secondary }, line: { color: p.secondary }, name: `${visual.visual_id}:pooled-diamond` });
    addText(slide, `${e.toFixed(2)}（${low.toFixed(2)}–${high.toFixed(2)}）`, 10.82, y - 0.12, 2.05, 0.32, TYPO.chart, { bold: true, align: "right", color: p.primary });
  }
  addAxis(slide, x0, 6.52, x1 - x0, ticks, scale, visual.payload.axis_label || visual.unit, p, value => log ? value.toFixed(value < 1 ? 1 : 0) : value.toFixed(2));
}

function renderBoundary(slide, visual, spec) {
  const p = profile(spec);
  const allowed = visual.payload.allowed || visual.payload.available || visual.payload.planned || [];
  const prohibited = visual.payload.prohibited || visual.payload.missing || visual.payload.not_available || [];
  const leftTitle = visual.payload.available ? "当前已有" : visual.payload.planned ? "已预设" : "允许表述";
  const rightTitle = visual.payload.missing ? "仍需验证" : visual.payload.not_available ? "当前没有" : "禁止升级";
  box(slide, 0.88, 1.55, 5.55, 4.95, p.soft, p.accent, 0.03);
  box(slide, 6.90, 1.55, 5.55, 4.95, "FFF3E7", SHARED.warning, 0.03);
  addText(slide, leftTitle, 1.16, 1.86, 4.90, 0.40, TYPO.sub, { bold: true, color: p.primary });
  addText(slide, rightTitle, 7.18, 1.86, 4.90, 0.40, TYPO.sub, { bold: true, color: SHARED.warning });
  allowed.forEach((item, index) => addText(slide, `● ${item}`, 1.18, 2.62 + index * 0.80, 4.85, 0.46, TYPO.body, { color: SHARED.ink, name: rowName(visual, index) }));
  prohibited.forEach((item, index) => addText(slide, `— ${item}`, 7.20, 2.62 + index * 0.80, 4.85, 0.46, TYPO.body, { color: SHARED.ink, name: rowName(visual, allowed.length + index) }));
}

function renderTakeaways(slide, visual, spec) {
  const p = profile(spec), items = visual.payload.takeaways || [];
  const w = 11.45 / items.length;
  items.forEach((item, index) => {
    const x = 0.84 + index * w;
    addText(slide, String(index + 1).padStart(2, "0"), x, 1.80, w - 0.24, 0.55, 30, { bold: true, color: index === 1 ? p.secondary : p.accent });
    line(slide, x, 2.52, w - 0.34, 0, index === 1 ? p.secondary : p.accent, 2.2);
    addText(slide, item, x, 2.88, w - 0.42, 2.45, TYPO.sub, { bold: true, color: p.primary, valign: "top", name: rowName(visual, index) });
  });
}

function renderQCFunnel(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const max = Math.max(...rows.map(row => Number(row.remaining_cells)));
  rows.forEach((row, index) => {
    const width = 9.8 * Number(row.remaining_cells) / max;
    const x = (W - width) / 2;
    const y = 1.50 + index * 1.22;
    slide.addShape(S.chevron, { x, y, w: width, h: 0.78, fill: { color: index % 2 ? p.soft : p.accent, transparency: index % 2 ? 0 : 10 }, line: { color: p.accent, width: 1 }, name: rowName(visual, index) });
    addText(slide, `${zhLabel(row.stage)}：${Number(row.remaining_cells).toLocaleString()}（本步剔除${row.removed_at_stage}）`, x + 0.22, y + 0.16, width - 0.44, 0.38, TYPO.body, { bold: true, color: index % 2 ? p.primary : SHARED.white, align: "center" });
  });
}

function renderFigureCallout(slide, visual, spec) {
  const p = profile(spec);
  imageContain(slide, visual.payload.image_path, 0.82, 1.40, 8.75, 5.45, `${visual.visual_id}:image`);
  box(slide, 9.82, 1.72, 2.62, 3.12, p.soft, p.accent, 0.03, rowName(visual, 0));
  addText(slide, "解释边界", 10.08, 2.00, 2.05, 0.40, TYPO.sub, { bold: true, color: p.primary });
  addText(slide, visual.payload.callout, 10.08, 2.74, 2.06, 1.58, TYPO.body, { color: SHARED.ink, valign: "top" });
  addText(slide, "来源图仅用于展示已登记信息", 10.08, 5.42, 2.05, 0.62, TYPO.detail, { color: SHARED.muted });
}

const CELL_COLORS = {
  "T cell": "3B82C4", "B cell": "6F5BD3", "NK cell": "7A9E33", "Monocyte": "E69F00",
  "Macrophage": "D95F59", "Dendritic": "B07AA1", "Endothelial": "2A9D8F", "Fibroblast": "9C755F",
  "Pericyte": "76B7B2", "Cardiomyocyte": "4E79A7",
};

function renderComposition(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const groups = [...new Set(rows.map(row => row.group))];
  if (groups.length !== 2) throw new Error(`grouped_composition requires exactly two groups; received ${groups.length}`);
  const [groupA, groupB] = groups;
  const cellTypes = [...new Set(rows.map(row => row.cell_type))];
  const comparisons = cellTypes.map(cellType => {
    const first = rows.find(row => row.cell_type === cellType && row.group === groupA);
    const second = rows.find(row => row.cell_type === cellType && row.group === groupB);
    if (!first || !second) throw new Error(`Missing composition row for ${cellType}`);
    return { cellType, first: Number(first.percentage), second: Number(second.percentage), delta: Number(second.percentage) - Number(first.percentage), firstIndex: rows.indexOf(first), secondIndex: rows.indexOf(second) };
  }).sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
  const maxValue = Math.max(5, Math.ceil(Math.max(...comparisons.flatMap(row => [row.first, row.second])) / 5) * 5);
  const x0 = 4.08, x1 = 10.80, y0 = 1.52, rowH = 0.48, scale = linearScale(0, maxValue, x0, x1);
  comparisons.forEach((row, index) => {
    const y = y0 + index * rowH;
    addText(slide, zhLabel(row.cellType), 0.82, y - 0.02, 2.78, 0.28, TYPO.chart, { color: SHARED.ink });
    line(slide, scale(Math.min(row.first, row.second)), y + 0.11, Math.abs(scale(row.second) - scale(row.first)), 0, "B7C7D2", 1.3);
    ellipse(slide, scale(row.first) - 0.06, y + 0.05, 0.12, 0.12, p.accent, undefined, rowName(visual, row.firstIndex));
    ellipse(slide, scale(row.second) - 0.06, y + 0.05, 0.12, 0.12, p.secondary, undefined, rowName(visual, row.secondIndex));
    addText(slide, `${row.first.toFixed(1)}% / ${row.second.toFixed(1)}%`, 10.95, y - 0.04, 1.35, 0.28, TYPO.chart, { color: SHARED.muted, align: "right" });
    addText(slide, `${row.delta >= 0 ? "+" : ""}${row.delta.toFixed(1)} pp`, 12.35, y - 0.04, 0.78, 0.28, TYPO.chart, { bold: Math.abs(row.delta) > 5, color: Math.abs(row.delta) > 5 ? p.secondary : SHARED.muted, align: "right" });
  });
  const ticks = Array.from({ length: 6 }, (_, index) => index * maxValue / 5);
  addAxis(slide, x0, 6.38, x1 - x0, ticks, scale, "细胞比例（%）", p, value => String(Math.round(value)));
  ellipse(slide, 8.45, 1.17, 0.12, 0.12, p.accent); addText(slide, zhLabel(groupA), 8.63, 1.09, 1.35, 0.28, TYPO.chart, { color: SHARED.muted });
  ellipse(slide, 10.25, 1.17, 0.12, 0.12, p.secondary); addText(slide, zhLabel(groupB), 10.43, 1.09, 1.35, 0.28, TYPO.chart, { color: SHARED.muted });
}

function renderFacetedVolcano(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const facets = [...new Set(rows.map(row => row.cell_type))];
  const effectField = visual.x_field;
  facets.forEach((facet, index) => {
    const col = index % 3, rowIndex = Math.floor(index / 3);
    const x = 0.78 + col * 4.18, y = 1.45 + rowIndex * 2.63, w = 3.82, h = 2.20;
    box(slide, x, y, w, h, SHARED.white, "D6E1E8", 0.02);
    addText(slide, zhLabel(facet), x + 0.14, y + 0.10, w - 0.28, 0.28, TYPO.detail, { bold: true, color: p.primary });
    const subset = rows.filter(item => item.cell_type === facet);
    const scaleX = linearScale(-2.2, 2.2, x + 0.38, x + w - 0.38);
    const scaleY = linearScale(0, 18, y + h - 0.34, y + 0.52);
    line(slide, scaleX(0), y + 0.48, 0, h - 0.76, SHARED.line, 0.8, { dash: "dash" });
    const labelRows = new Set([...subset].sort((a, b) => Number(a.FDR) - Number(b.FDR)).slice(0, 2));
    const placedLabels = [];
    subset.forEach(item => {
      const effect = Number(item[effectField]), score = -Math.log10(Math.max(Number(item.FDR), 1e-20));
      const cx = scaleX(effect), cy = scaleY(Math.min(18, score));
      ellipse(slide, cx - 0.055, cy - 0.055, 0.11, 0.11, effect >= 0 ? p.secondary : p.negative || "3B82C4", undefined, rowName(visual, rows.indexOf(item)));
      if (labelRows.has(item)) {
        const labelW = 1.50, labelH = 0.30;
        let labelY = cy - 0.38;
        const labelX = Math.max(x + 0.04, Math.min(x + w - labelW - 0.04, cx - labelW / 2));
        for (const placed of placedLabels) {
          if (Math.abs(labelX - placed.x) < labelW && Math.abs(labelY - placed.y) < labelH + 0.08) {
            labelY = Math.min(y + h - labelH - 0.10, placed.y + labelH + 0.10);
          }
        }
        placedLabels.push({ x: labelX, y: labelY });
        addText(slide, item.gene, labelX, labelY, labelW, labelH, TYPO.chart, { bold: score > 5, align: "center", color: SHARED.ink, fontFace: "Arial" });
      }
    });
    addText(slide, "log2FC", x + w - 0.90, y + h - 0.28, 0.72, 0.20, TYPO.chart, { color: SHARED.muted, align: "right" });
  });
}

function renderEnrichment(slide, visual, spec) {
  const rows = [...visual.payload.rows].sort((a, b) => a.cell_type.localeCompare(b.cell_type) || Number(b.normalized_enrichment_score) - Number(a.normalized_enrichment_score));
  const p = profile(spec), x0 = 5.45, x1 = 11.25, zero = (x0 + x1) / 2, maxAbs = 2.5;
  const scale = linearScale(-maxAbs, maxAbs, x0, x1);
  line(slide, zero, 1.48, 0, 4.92, SHARED.muted, 1.1);
  let lastFacet = "";
  rows.forEach((row, index) => {
    const y = 1.58 + index * 0.76, value = Number(row.normalized_enrichment_score), x = scale(value);
    if (row.cell_type !== lastFacet) {
      addText(slide, zhLabel(row.cell_type), 0.80, y - 0.10, 1.45, 0.28, TYPO.detail, { bold: true, color: p.primary });
      lastFacet = row.cell_type;
    }
    addText(slide, zhLabel(row.pathway), 2.25, y - 0.08, 2.85, 0.30, TYPO.chart, { color: SHARED.ink });
    const originalIndex = visual.payload.rows.indexOf(row);
    slide.addShape(S.rect, { x: Math.min(zero, x), y: y + 0.02, w: Math.max(0.03, Math.abs(x - zero)), h: 0.24, fill: { color: value >= 0 ? p.secondary : p.negative || "3B82C4" }, line: { color: value >= 0 ? p.secondary : p.negative || "3B82C4" }, name: rowName(visual, originalIndex) });
    addText(slide, `NES ${value.toFixed(2)}；FDR ${row.FDR}`, 11.38, y - 0.08, 1.72, 0.30, TYPO.chart, { align: "right", color: SHARED.muted });
  });
  addAxis(slide, x0, 6.52, x1 - x0, [-2, -1, 0, 1, 2], scale, "标准化富集分数（NES）", p, value => String(value));
}

function renderNetwork(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const names = [...new Set(rows.flatMap(row => [row.sender, row.receiver]))];
  const center = { x: 6.60, y: 3.78 }, radiusX = 4.55, radiusY = 2.20;
  const positions = Object.fromEntries(names.map((name, index) => {
    const angle = -Math.PI / 2 + index / names.length * Math.PI * 2;
    return [name, { x: center.x + Math.cos(angle) * radiusX, y: center.y + Math.sin(angle) * radiusY }];
  }));
  rows.forEach((edge, index) => {
    const a = positions[edge.sender], b = positions[edge.receiver], probability = Number(edge.communication_probability);
    line(slide, a.x, a.y, b.x - a.x, b.y - a.y, p.accent, 0.9 + probability * 2.2, { dash: "dash", endArrowType: "triangle", name: rowName(visual, index) });
    const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2;
    box(slide, mx - 0.88, my - 0.19, 1.76, 0.38, SHARED.white, p.accent, 0.02);
    addText(slide, `${edge.ligand_receptor} · ${probability.toFixed(2)}`, mx - 0.82, my - 0.15, 1.64, 0.28, TYPO.chart, { align: "center", color: SHARED.ink });
  });
  names.forEach((name, index) => {
    const pos = positions[name], color = CELL_COLORS[name] || (index % 2 ? p.secondary : p.accent);
    ellipse(slide, pos.x - 0.55, pos.y - 0.34, 1.10, 0.68, color, SHARED.white, `${visual.visual_id}:node:${index}`);
    addText(slide, zhLabel(name), pos.x - 0.48, pos.y - 0.21, 0.96, 0.40, TYPO.chart, { bold: true, color: SHARED.white, align: "center" });
  });
  box(slide, 9.72, 6.02, 2.65, 0.52, "FFF4E5", SHARED.warning, 0.02);
  addText(slide, "虚线＝预测/优先验证互作", 9.88, 6.14, 2.33, 0.26, TYPO.chart, { bold: true, color: SHARED.warning, align: "center" });
}

function renderLadder(slide, visual, spec) {
  const steps = visual.payload.steps || [], status = visual.payload.status || [], p = profile(spec);
  const x0 = 0.92, y0 = 5.82, totalW = 11.45, stepW = totalW / steps.length;
  steps.forEach((step, index) => {
    const x = x0 + index * stepW, y = y0 - index * 0.78;
    slide.addShape(S.chevron, { x, y, w: stepW - 0.12, h: 0.68, fill: { color: index === steps.length - 1 ? "FFF4E5" : index % 2 ? p.soft : p.accent, transparency: index % 2 ? 0 : 8 }, line: { color: index === steps.length - 1 ? SHARED.warning : p.accent }, name: rowName(visual, index) });
    const stepColor = index === steps.length - 1 ? SHARED.warning : index % 2 ? p.primary : SHARED.white;
    addText(slide, zhLabel(step), x + 0.10, y + 0.10, stepW - 0.32, 0.28, TYPO.detail, { bold: true, color: stepColor, align: "center" });
    if (status[index]) addText(slide, zhLabel(status[index]), x + 0.08, y + 0.78, stepW - 0.28, 0.28, TYPO.chart, { align: "center", color: index === steps.length - 1 ? SHARED.warning : SHARED.muted });
  });
}

function renderStudyCharacteristics(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  addText(slide, `${rows.length}项研究`, 0.82, 1.50, 2.25, 0.32, TYPO.body, { bold: true, color: p.primary });
  addText(slide, `总样本量 ${Number(visual.payload.total_n).toLocaleString()}｜死亡事件 ${Number(visual.payload.total_events).toLocaleString()}`, 3.05, 1.50, 5.60, 0.32, TYPO.body, { bold: true, color: p.accent });
  addText(slide, "点的位置表示每项研究样本量；颜色表示来源中的偏倚风险概括", 8.28, 1.50, 4.15, 0.32, TYPO.chart, { color: SHARED.muted, align: "right" });
  const sizes = rows.map(row => Number(row.sample_size));
  const maxSize = Math.max(...sizes), scaleX = linearScale(0, maxSize, 4.55, 11.20);
  const plotTop = 1.92, plotBottom = 6.12, rowH = (plotBottom - plotTop) / rows.length;
  rows.forEach((row, index) => {
    const y = plotTop + index * rowH, x = scaleX(Number(row.sample_size));
    const studyLabel = String(row.study_label || "").trim();
    const year = String(row.year || "").trim();
    const displayLabel = year && !studyLabel.endsWith(year) ? `${studyLabel} · ${year}` : studyLabel;
    addText(slide, displayLabel, 0.82, y - 0.01, 3.48, Math.max(0.23, rowH - 0.01), TYPO.chart, { color: SHARED.ink });
    line(slide, 4.55, y + rowH * 0.42, Math.max(0.03, x - 4.55), 0, "CAD7E1", 1.0);
    const color = row.risk_of_bias === "High" ? "C74B50" : row.risk_of_bias === "Low" ? "4C956C" : "E9A23B";
    ellipse(slide, x - 0.065, y + rowH * 0.42 - 0.065, 0.13, 0.13, color, undefined, rowName(visual, index));
    addText(slide, `n=${Number(row.sample_size).toLocaleString()}`, 11.35, y - 0.01, 1.05, Math.max(0.23, rowH - 0.01), TYPO.chart, { align: "right", color: SHARED.muted });
  });
  addAxis(slide, 4.55, 6.35, 6.65, [0, maxSize / 2, maxSize], scaleX, "单项研究样本量", p, value => String(Math.round(value)));
}

function renderPooledPanel(slide, visual, spec) {
  const p = profile(spec), rows = visual.payload.rows || [];
  if (rows.length !== 3) throw new Error(`Pooled evidence panel requires fixed, random, and prediction rows: ${visual.visual_id}`);
  const entries = rows.map(row => {
    const estimand = String(row.estimand || "");
    if (/prediction interval/i.test(estimand)) return { heading: "预测区间", value: `${row.ci_low}–${row.ci_high}`, detail: "未来研究真实效应的可能范围", prediction: true };
    if (/random/i.test(estimand)) return { heading: "随机效应OR", value: row.estimate, detail: `95%CI ${row.ci_low}–${row.ci_high}；I² ${row.I2_percent}%` };
    if (/fixed/i.test(estimand)) return { heading: "固定效应OR", value: row.estimate, detail: `95%CI ${row.ci_low}–${row.ci_high}` };
    return { heading: estimand, value: row.estimate || `${row.ci_low}–${row.ci_high}`, detail: row.status || "已登记合并估计" };
  });
  entries.forEach((entry, index) => {
    const x = 0.84 + index * 4.02;
    line(slide, x, 1.75, 3.55, 0, entry.prediction ? p.secondary : p.accent, 3, { name: rowName(visual, index) });
    addText(slide, entry.heading, x, 2.05, 3.45, 0.36, TYPO.detail, { bold: true, color: SHARED.muted });
    addText(slide, entry.value, x, 2.66, 3.45, 0.72, 32, { bold: true, color: p.primary });
    addText(slide, entry.detail, x, 3.66, 3.45, 0.72, TYPO.detail, { color: SHARED.ink });
  });
  box(slide, 1.55, 5.18, 10.25, 0.92, "FFF4E5", SHARED.warning, 0.03);
  addText(slide, "预测区间不是置信区间：它描述未来研究真实效应可能出现的范围。", 1.82, 5.42, 9.72, 0.38, TYPO.body, { bold: true, color: SHARED.warning, align: "center" });
}

function renderROB(slide, visual, spec) {
  const rows = visual.payload.rows;
  const studies = [...new Set(rows.map(row => row.study_label))];
  const domains = [...new Set(rows.map(row => row.domain))];
  const colors = { Low: "4C956C", "Some concerns": "E9A23B", High: "C74B50" };
  const x0 = 4.38, y0 = 1.58, colW = 1.38, rowH = 0.335;
  domains.forEach((domain, index) => addText(slide, zhLabel(domain), x0 + index * colW - 0.08, 1.14, colW, 0.40, TYPO.chart, { bold: true, align: "center", color: SHARED.muted }));
  studies.forEach((study, rowIndex) => {
    const y = y0 + rowIndex * rowH;
    addText(slide, study, 0.82, y - 0.01, 3.32, rowH - 0.01, TYPO.chart, { color: SHARED.ink });
    domains.forEach((domain, colIndex) => {
      const item = rows.find(row => row.study_label === study && row.domain === domain);
      const fill = item ? colors[item.judgment] : "D9E2E8";
      const sourceIndex = item ? rows.indexOf(item) : -1;
      slide.addShape(S.rect, { x: x0 + colIndex * colW, y, w: colW - 0.08, h: rowH - 0.055, fill: { color: fill }, line: { color: SHARED.white, width: 0.6 }, name: sourceIndex >= 0 ? rowName(visual, sourceIndex) : `${visual.visual_id}:missing:${rowIndex}-${colIndex}` });
    });
  });
  const legend = [["低风险", colors.Low], ["存在担忧", colors["Some concerns"]], ["高风险", colors.High]];
  legend.forEach((item, index) => {
    slide.addShape(S.rect, { x: 8.05 + index * 1.62, y: 6.48, w: 0.24, h: 0.18, fill: { color: item[1] }, line: { color: item[1] } });
    addText(slide, item[0], 8.35 + index * 1.62, 6.40, 1.15, 0.30, TYPO.chart, { color: SHARED.muted });
  });
}

function renderGrade(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const cols = [0.82, 4.95, 6.55, 8.10, 9.65, 11.15];
  const widths = [3.95, 1.42, 1.38, 1.38, 1.38, 1.30];
  const headers = ["结局", "研究数", "参与者", "偏倚风险", "间接性", "确定性"];
  headers.forEach((header, index) => addText(slide, header, cols[index], 1.47, widths[index], 0.34, TYPO.chart, { bold: true, color: SHARED.muted, align: index ? "center" : "left" }));
  rows.forEach((row, rowIndex) => {
    const y = 2.02 + rowIndex * 1.35;
    if (rowIndex % 2 === 0) slide.addShape(S.rect, { x: 0.76, y: y - 0.12, w: 11.82, h: 1.00, fill: { color: p.soft }, line: { color: p.soft } });
    const certaintyColor = ({ High: "2F855A", Moderate: "4C78A8", Low: "E9A23B", "Very low": "C74B50" })[row.certainty] || SHARED.muted;
    const values = [row.outcome, row.studies, row.participants, row.risk_of_bias, row.indirectness, row.certainty];
    values.forEach((value, index) => addText(slide, zhLabel(value), cols[index], y, widths[index], 0.48, index === 5 ? TYPO.body : TYPO.detail, { bold: index === 0 || index === 5, color: index === 5 ? certaintyColor : SHARED.ink, align: index ? "center" : "left", name: index === 0 ? rowName(visual, rowIndex) : undefined }));
    addText(slide, `不一致性：${zhLabel(row.inconsistency)}；不精确性：${zhLabel(row.imprecision)}`, cols[0], y + 0.50, 8.25, 0.30, TYPO.chart, { color: SHARED.muted });
  });
}

function renderEligibility(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const groups = ["Inclusion", "Exclusion"];
  groups.forEach((group, groupIndex) => {
    const subset = rows.filter(row => row.category === group);
    const x = groupIndex === 0 ? 0.82 : 6.85;
    addText(slide, groupIndex === 0 ? "纳入标准" : "排除标准", x, 1.45, 5.35, 0.42, TYPO.sub, { bold: true, color: groupIndex === 0 ? p.accent : p.secondary });
    line(slide, x, 1.98, 5.30, 0, groupIndex === 0 ? p.accent : p.secondary, 2);
    subset.forEach((row, index) => {
      const y = 2.30 + index * 0.94;
      ellipse(slide, x + 0.02, y + 0.05, 0.22, 0.22, groupIndex === 0 ? p.accent : p.secondary, undefined, rowName(visual, rows.indexOf(row)));
      addText(slide, zhLabel(row.criterion), x + 0.42, y - 0.02, 4.78, 0.46, TYPO.body, { color: SHARED.ink });
      addText(slide, row.operational_status === "Required" ? "必须满足" : "排除", x + 0.42, y + 0.47, 2.00, 0.26, TYPO.chart, { bold: true, color: groupIndex === 0 ? p.accent : p.secondary });
    });
  });
}

function renderAssessmentTimeline(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  line(slide, 1.00, 3.55, 8.60, 0, p.accent, 3);
  line(slide, 9.86, 3.55, 2.15, 0, p.accent, 3);
  addText(slide, "//", 9.56, 3.37, 0.35, 0.34, TYPO.sub, { bold: true, color: SHARED.muted, align: "center" });
  const lanes = [
    // The four lanes have disjoint vertical bands.  Horizontal packing alone
    // is insufficient when several assessments share the same 0–6 h window.
    { visitY: 2.38, assessmentY: 2.66, lastRight: -Infinity },
    { visitY: 1.28, assessmentY: 1.56, lastRight: -Infinity },
    { visitY: 3.82, assessmentY: 4.10, lastRight: -Infinity },
    { visitY: 4.98, assessmentY: 5.26, lastRight: -Infinity },
  ];
  const labelW = 1.64;
  [...rows].map((row, originalIndex) => ({ row, originalIndex })).sort((a, b) => Number(a.row.position_hours) - Number(b.row.position_hours)).forEach(({ row, originalIndex }) => {
    const hour = Number(row.position_hours);
    const x = hour <= 72 ? 1.0 + hour / 72 * 8.60 : 11.65;
    const labelX = Math.max(0.76, Math.min(11.76, x - labelW / 2));
    const lane = lanes.find(candidate => labelX >= candidate.lastRight + 0.08);
    if (!lane) throw new Error(`Assessment timeline label budget exceeded near hour ${hour}`);
    lane.lastRight = labelX + labelW;
    ellipse(slide, x - 0.10, 3.45, 0.20, 0.20, row.status === "Primary endpoint" ? p.secondary : p.accent, undefined, rowName(visual, originalIndex));
    line(slide, x, 3.45, 0, lane.visitY < 3.45 ? lane.assessmentY + 0.56 - 3.45 : lane.visitY - 3.45, SHARED.line, 0.7);
    addText(slide, zhLabel(row.visit), labelX, lane.visitY, labelW, 0.26, TYPO.chart, { bold: true, align: "center", color: p.primary });
    addText(slide, `${zhLabel(row.assessment)}\n${zhLabel(row.window)}`, labelX, lane.assessmentY, labelW, 0.56, TYPO.chart, { align: "center", color: SHARED.ink, valign: "top" });
  });
  addText(slide, "0–72小时主要观察窗", 2.65, 6.34, 4.65, 0.30, TYPO.body, { bold: true, color: p.primary, align: "center" });
  addText(slide, "第28天", 10.52, 6.34, 2.10, 0.30, TYPO.body, { bold: true, color: p.primary, align: "center" });
}

function renderVariables(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const primary = rows.filter(row => row.role.startsWith("Primary"));
  const others = rows.filter(row => !row.role.startsWith("Primary"));
  box(slide, 0.82, 1.48, 5.25, 4.98, p.soft, p.accent, 0.03);
  addText(slide, "主要暴露与结局", 1.08, 1.78, 4.70, 0.38, TYPO.sub, { bold: true, color: p.primary });
  primary.forEach((row, index) => {
    const y = 2.44 + index * 1.52;
    addText(slide, row.role === "Primary exposure" ? "主要暴露" : "主要结局", 1.08, y, 1.35, 0.28, TYPO.chart, { bold: true, color: p.accent });
    addText(slide, zhLabel(row.variable), 1.08, y + 0.36, 4.55, 0.38, TYPO.body, { bold: true, name: rowName(visual, rows.indexOf(row)) });
    addText(slide, `${zhLabel(row.timepoint)}｜${zhLabel(row.type)}`, 1.08, y + 0.86, 4.55, 0.30, TYPO.detail, { color: SHARED.muted });
  });
  addText(slide, "次要结局、协变量与质量变量", 6.52, 1.62, 5.75, 0.36, TYPO.sub, { bold: true, color: p.primary });
  others.forEach((row, index) => {
    const y = 2.15 + index * 0.57;
    addText(slide, zhLabel(row.variable), 6.54, y, 3.25, 0.28, TYPO.detail, { bold: row.role === "Quality variable", name: rowName(visual, rows.indexOf(row)) });
    addText(slide, zhLabel(row.timepoint), 9.80, y, 1.35, 0.28, TYPO.chart, { color: SHARED.muted });
    addText(slide, row.role.replace("Secondary outcome", "次要结局").replace("Covariate", "协变量").replace("Quality variable", "质量变量"), 11.15, y, 1.18, 0.28, TYPO.chart, { color: row.role === "Quality variable" ? p.secondary : SHARED.muted, align: "right" });
  });
}

function renderSampleSize(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const get = name => rows.find(row => row.parameter === name);
  const targetRow = get("Target enrollment"), analyzableRow = get("Required analyzable sample"), lossRow = get("Loss/non-evaluable rate");
  if (!targetRow || !analyzableRow || !lossRow) throw new Error(`Sample-size contract is missing required planning rows: ${visual.visual_id}`);
  const target = Number(targetRow.value), analyzable = Number(analyzableRow.value), loss = Number(lossRow.value) * 100;
  addText(slide, target, 0.88, 1.55, 3.25, 0.78, 42, { bold: true, color: p.accent, name: rowName(visual, rows.indexOf(targetRow)) });
  addText(slide, "计划目标入组", 0.90, 2.38, 3.0, 0.34, TYPO.body, { color: SHARED.muted });
  line(slide, 4.15, 2.10, 2.10, 0, p.accent, 2.4, { endArrowType: "triangle" });
  addText(slide, `预计${loss.toFixed(0)}%不可评估`, 4.05, 2.48, 2.35, 0.34, TYPO.detail, { bold: true, color: SHARED.warning, align: "center", name: rowName(visual, rows.indexOf(lossRow)) });
  addText(slide, analyzable, 6.55, 1.55, 3.25, 0.78, 42, { bold: true, color: p.secondary, name: rowName(visual, rows.indexOf(analyzableRow)) });
  addText(slide, "所需可分析样本", 6.57, 2.38, 3.0, 0.34, TYPO.body, { color: SHARED.muted });
  const assumptions = rows.filter(row => !["Target enrollment", "Required analyzable sample", "Loss/non-evaluable rate"].includes(row.parameter));
  line(slide, 0.90, 3.24, 11.45, 0, SHARED.line, 1);
  assumptions.forEach((row, index) => {
    const x = 0.90 + index * 2.34;
    addText(slide, zhLabel(row.parameter), x, 3.62, 2.10, 0.62, TYPO.detail, { bold: true, color: p.primary, name: rowName(visual, rows.indexOf(row)) });
    let display = row.value;
    if (row.unit === "Proportion" || row.unit === "Probability") display = `${(Number(row.value) * 100).toFixed(0)}%`;
    if (row.unit === "OR") display = `OR ${row.value}`;
    addText(slide, display, x, 4.55, 2.10, 0.52, TYPO.sub, { bold: true, color: index % 2 ? p.secondary : p.accent });
    addText(slide, zhLabel(row.unit), x, 5.22, 2.10, 0.28, TYPO.chart, { color: SHARED.muted });
  });
}

function renderDAG(slide, visual, spec) {
  const p = profile(spec), nodes = visual.payload.nodes, edges = visual.payload.edges;
  const nodeWidth = 2.45, nodeHeight = 0.92;
  const byRole = role => nodes.filter(node => node.role === role);
  const positions = {};
  const distribute = (items, x, top, bottom) => items.forEach((node, index) => {
    const y = items.length === 1 ? (top + bottom) / 2 : top + index * ((bottom - top) / Math.max(1, items.length - 1));
    positions[node.id] = { x, y };
  });
  distribute(byRole("confounder"), 0.92, 1.55, 4.72);
  distribute(byRole("exposure"), 4.55, 2.15, 4.05);
  distribute(byRole("mediator").concat(byRole("management")), 7.05, 1.60, 4.65);
  distribute(byRole("outcome"), 10.02, 2.15, 4.05);
  const placed = new Set(Object.keys(positions));
  distribute(nodes.filter(node => !placed.has(node.id)), 7.05, 1.60, 4.65);
  edges.forEach((edge, index) => {
    const a = positions[edge.source], b = positions[edge.target];
    if (!a || !b) throw new Error(`DAG edge references missing node: ${edge.source}->${edge.target}`);
    line(slide, a.x + nodeWidth / 2, a.y + nodeHeight / 2, b.x - a.x - nodeWidth / 2, b.y - a.y, p.accent, nodes.find(node => node.id === edge.source)?.role === "exposure" ? 2.4 : 1.5, { endArrowType: "triangle", name: `${visual.visual_id}:edge:${index}` });
  });
  nodes.forEach((node, index) => {
    const pos = positions[node.id];
    const fill = node.role === "exposure" ? p.soft : node.role === "outcome" ? p.soft2 : SHARED.white;
    const stroke = node.role === "outcome" ? p.secondary : p.accent;
    box(slide, pos.x, pos.y, nodeWidth, nodeHeight, fill, stroke, 0.03, rowName(visual, index));
    addText(slide, node.label, pos.x + 0.12, pos.y + 0.10, nodeWidth - 0.24, nodeHeight - 0.20, TYPO.chart, { bold: true, align: "center", color: p.primary });
  });
  box(slide, 2.10, 5.55, 9.20, 0.98, "FFF4E5", SHARED.warning, 0.03);
  addText(slide, `预设调整集：${visual.payload.adjustment_set.join("、")}`, 2.32, 5.66, 8.76, 0.28, TYPO.chart, { bold: true, color: SHARED.warning, align: "center" });
  if (visual.payload.review_note) addText(slide, visual.payload.review_note, 2.32, 6.02, 8.76, 0.28, TYPO.chart, { color: SHARED.warning, align: "center" });
}

function renderRiskMatrix(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec), x0 = 4.45, y0 = 1.52, cellW = 2.15, cellH = 1.55;
  const fills = [
    ["EAF5ED", "EEF3DF", "FFF0D9"],
    ["EEF3DF", "FFF0D9", "FADFD9"],
    ["FFF0D9", "FADFD9", "F3C7C7"],
  ];
  for (let impact = 1; impact <= 3; impact++) {
    for (let probability = 1; probability <= 3; probability++) {
      const x = x0 + (probability - 1) * cellW, y = y0 + (3 - impact) * cellH;
      slide.addShape(S.rect, { x, y, w: cellW - 0.05, h: cellH - 0.05, fill: { color: fills[impact - 1][probability - 1] }, line: { color: SHARED.white, width: 1 } });
    }
  }
  ["低", "中", "高"].forEach((label, index) => addText(slide, label, x0 + index * cellW, 6.28, cellW - 0.05, 0.30, TYPO.chart, { align: "center", color: SHARED.muted }));
  ["低", "中", "高"].forEach((label, index) => addText(slide, label, 4.02, y0 + (2 - index) * cellH + 0.48, 0.32, 0.30, TYPO.chart, { align: "center", color: SHARED.muted }));
  addText(slide, "发生概率 →", x0, 6.62, cellW * 3, 0.26, TYPO.chart, { align: "center", color: SHARED.muted });
  addText(slide, "影响 ↑", x0 - 0.62, 1.20, 0.58, 0.26, TYPO.chart, { bold: true, color: SHARED.muted, align: "right" });
  const offsets = {};
  rows.forEach((row, index) => {
    const key = `${row.probability_score}-${row.impact_score}`;
    const offset = offsets[key] || 0; offsets[key] = offset + 1;
    const cx = x0 + (Number(row.probability_score) - 1) * cellW + 0.45 + offset * 0.72;
    const cy = y0 + (3 - Number(row.impact_score)) * cellH + 0.32 + (offset % 2) * 0.48;
    ellipse(slide, cx, cy, 0.42, 0.42, p.accent, SHARED.white, rowName(visual, index));
    addText(slide, String(index + 1), cx, cy + 0.03, 0.42, 0.30, TYPO.chart, { bold: true, color: SHARED.white, align: "center" });
    const listY = 1.43 + index * 0.88;
    addText(slide, `${index + 1}. ${zhLabel(row.risk)}`, 0.72, listY, 3.18, 0.36, TYPO.chart, { bold: true, color: SHARED.ink, valign: "top" });
    addText(slide, `缓解：${zhLabel(row.mitigation)}`, 0.90, listY + 0.38, 3.00, 0.44, TYPO.chart, { color: SHARED.muted, valign: "top" });
  });
}

function renderGantt(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  const dates = rows.flatMap(row => [new Date(row.start_date), new Date(row.end_date)]);
  const min = Math.min(...dates.map(date => date.getTime())), max = Math.max(...dates.map(date => date.getTime()));
  const plotLeft = 4.10, plotRight = 9.92;
  const scale = linearScale(min, max, plotLeft, plotRight), y0 = 1.55, rowH = 0.62;
  rows.forEach((row, index) => {
    const y = y0 + index * rowH, start = new Date(row.start_date).getTime(), end = new Date(row.end_date).getTime();
    addText(slide, zhLabel(row.milestone), 0.78, y, 3.05, 0.32, TYPO.chart, { color: SHARED.ink });
    slide.addShape(S.roundRect, { x: scale(start), y: y + 0.02, w: Math.max(0.18, scale(end) - scale(start)), h: 0.27, rectRadius: 0.03, fill: { color: index % 2 ? p.secondary : p.accent }, line: { color: index % 2 ? p.secondary : p.accent }, name: rowName(visual, index) });
    addText(slide, `${row.start_date}–${row.end_date}`, 10.10, y - 0.01, 2.55, 0.32, TYPO.chart, { align: "right", color: SHARED.muted });
  });
  const years = [...new Set(dates.map(date => date.getFullYear()))];
  years.forEach(year => {
    const timestamp = Math.max(min, Math.min(max, new Date(`${year}-01-01`).getTime()));
    const x = scale(timestamp);
    line(slide, x, 1.36, 0, 4.82, SHARED.line, 0.7, { dash: "dash" });
    addText(slide, year, x - 0.35, 6.30, 0.70, 0.28, TYPO.chart, { align: "center", color: SHARED.muted });
  });
}

function renderReview(slide, spec) {
  const conflicts = spec.conflict_bindings || [], unresolved = spec.unresolved_bindings || [];
  addText(slide, "来源冲突", 0.82, 1.45, 5.70, 0.38, TYPO.sub, { bold: true, color: profile(spec).primary });
  addText(slide, "待用户确认", 6.90, 1.45, 5.45, 0.38, TYPO.sub, { bold: true, color: SHARED.warning });
  conflicts.forEach((item, index) => {
    const y = 2.03 + index * 0.86;
    box(slide, 0.82, y, 5.72, 0.65, SHARED.white, SHARED.line, 0.02);
    addText(slide, `${item.conflicting_value} → ${item.canonical_value}`, 1.02, y + 0.13, 5.32, 0.32, TYPO.detail, { bold: true });
  });
  unresolved.forEach((item, index) => {
    const y = 2.03 + index * 0.86;
    box(slide, 6.90, y, 5.45, 0.65, "FFF4E5", SHARED.warning, 0.02);
    addText(slide, `${item.marker}｜INFORMATION_REQUIRED`, 7.10, y + 0.13, 5.05, 0.32, TYPO.detail, { bold: true, color: SHARED.warning });
  });
  addText(slide, "所有冲突和未解决项均保留在审计产物及人工审核清单中。", 0.84, 6.25, 11.45, 0.34, TYPO.body, { color: SHARED.muted, align: "center" });
}

function renderVisual(slide, visual, spec) {
  const before = (slide._slideObjects || []).length;
  const payload = visual.payload || {};
  switch (visual.visual_type) {
    case "hero_question": case "protocol_question": renderHeroQuestion(slide, visual, spec); break;
    case "specification_matrix": case "assumption_warning_matrix": renderMatrix(slide, visual, spec); break;
    case "clone_censor_weight_branch": case "single_cell_workflow": case "protocol_design": case "analysis_plan": renderBranch(slide, visual, spec); break;
    case "time_window_timeline": renderTimeWindow(slide, visual, spec); break;
    case "love_plot": renderLove(slide, visual, spec); break;
    case "weight_histogram": renderHistogram(slide, visual, spec); break;
    case "absolute_risk_comparison": renderAbsoluteRisk(slide, visual, spec); break;
    case "estimand_comparison": renderEstimandComparison(slide, visual, spec); break;
    case "forest_plot": renderForest(slide, visual, spec, false); break;
    case "subgroup_forest": renderForest(slide, visual, spec, true); break;
    case "evidence_boundary": case "validation_gap": case "no_result_boundary": renderBoundary(slide, visual, spec); break;
    case "takeaway_synthesis": renderTakeaways(slide, visual, spec); break;
    case "qc_funnel": renderQCFunnel(slide, visual, spec); break;
    case "umap_figure": case "source_figure": renderFigureCallout(slide, visual, spec); break;
    case "grouped_composition": renderComposition(slide, visual, spec); break;
    case "faceted_volcano": renderFacetedVolcano(slide, visual, spec); break;
    case "diverging_enrichment": renderEnrichment(slide, visual, spec); break;
    case "communication_network": renderNetwork(slide, visual, spec); break;
    case "evidence_ladder": case "quality_control_ladder": renderLadder(slide, visual, spec); break;
    case "study_characteristics": renderStudyCharacteristics(slide, visual, spec); break;
    case "pooled_evidence_panel": renderPooledPanel(slide, visual, spec); break;
    case "risk_of_bias_matrix": renderROB(slide, visual, spec); break;
    case "grade_summary": renderGrade(slide, visual, spec); break;
    case "eligibility_split": renderEligibility(slide, visual, spec); break;
    case "assessment_timeline": renderAssessmentTimeline(slide, visual, spec); break;
    case "variable_definition_matrix": renderVariables(slide, visual, spec); break;
    case "sample_size_waterfall": renderSampleSize(slide, visual, spec); break;
    case "dag": renderDAG(slide, visual, spec); break;
    case "risk_matrix": renderRiskMatrix(slide, visual, spec); break;
    case "timeline_gantt": renderGantt(slide, visual, spec); break;
    default: throw new Error(`No renderer registered for VisualSpec type ${visual.visual_type}`);
  }
  const created = (slide._slideObjects || []).slice(before);
  const createdObjects = created.length;
  const prefix = `${visual.visual_id}:row:`;
  const renderedRowKeys = [...new Set(created.map(object => object?.options?.name || "").filter(name => name.startsWith(prefix)).map(name => name.slice(prefix.length)))];
  const expectedRowKeys = visual.expected_row_keys || [];
  const missing = expectedRowKeys.filter(key => !renderedRowKeys.includes(key));
  const unexpected = renderedRowKeys.filter(key => !expectedRowKeys.includes(key));
  if (createdObjects <= 0) throw new Error(`Renderer created no PowerPoint objects for ${visual.visual_id}`);
  if (missing.length || unexpected.length || renderedRowKeys.length !== expectedRowKeys.length) {
    throw new Error(`Renderer row evidence mismatch for ${visual.visual_id}: missing=${missing.join(",")} unexpected=${unexpected.join(",")}`);
  }
  return { renderedRows: renderedRowKeys.length, renderedRowKeys, createdObjects };
}

const objectManifest = [];
const renderEvidence = {};
for (const spec of graph.slide_specs) {
  const slide = pptx.addSlide(masterName(spec.layout_family));
  if (spec.slide_role === "cover") {
    renderCover(slide, spec, graph.kind);
  } else {
    title(slide, spec);
    if (spec.slide_role === "appendix") renderReview(slide, spec);
    else for (const visual of spec.visual_specs) renderEvidence[visual.visual_id] = renderVisual(slide, visual, spec);
    footer(slide, spec);
  }
  slide.addNotes(speakerNotes(spec));
  for (const visual of spec.visual_specs) {
    objectManifest.push({
      slide_id: spec.slide_id,
      visual_id: visual.visual_id,
      visual_type: visual.visual_type,
      object_type: ["umap_figure", "source_figure"].includes(visual.visual_type) ? "registered_image_with_native_callout" : "native_powerpoint_shapes",
      editability: ["umap_figure", "source_figure"].includes(visual.visual_type) ? "mixed" : "full",
      data_contract_id: visual.data_contract_id,
      source_bindings: visual.source_bindings,
      expected_row_count: visual.expected_row_count,
      rendered_row_count: renderEvidence[visual.visual_id]?.renderedRows ?? 0,
      rendered_row_keys: renderEvidence[visual.visual_id]?.renderedRowKeys ?? [],
      expected_row_keys: visual.expected_row_keys,
      created_object_count: renderEvidence[visual.visual_id]?.createdObjects ?? 0,
      fallback_used: false,
      warnings: [],
    });
  }
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
await pptx.writeFile({ fileName: outputPath });
const manifestPath = outputPath.replace(/\.pptx$/i, ".object_manifest.json");
fs.writeFileSync(manifestPath, JSON.stringify(objectManifest, null, 2), "utf8");
console.log(`WROTE=${outputPath}`);
console.log(`SLIDES=${graph.slide_specs.length}`);
console.log(`OBJECT_MANIFEST=${manifestPath}`);
