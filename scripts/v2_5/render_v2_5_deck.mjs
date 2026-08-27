import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");
const imageSizeModule = require("image-size");
const imageSize = imageSizeModule.imageSize || imageSizeModule.default || imageSizeModule;

const [graphPath, outputPath] = process.argv.slice(2);
if (!graphPath || !outputPath) throw new Error("Usage: render_v2_5_deck.mjs <canonical_object_graph.json> <output.pptx>");
const graph = JSON.parse(fs.readFileSync(graphPath, "utf8"));

// v2.5 is a render-only compatibility layer.  It consumes presentation fields
// when the upstream canonical graph supplies them, but never reparses source
// files or mutates evidence/model state.
const ART_DIRECTION_SPEC = graph.art_direction_spec || graph.art_direction || graph.presentation?.art_direction || {};
const RAW_RHYTHM_PLAN = graph.deck_rhythm_plan || graph.rhythm_plan || graph.presentation?.deck_rhythm_plan || [];
const RHYTHM_ENTRIES = Array.isArray(RAW_RHYTHM_PLAN) ? RAW_RHYTHM_PLAN : (RAW_RHYTHM_PLAN.slides || RAW_RHYTHM_PLAN.entries || []);
const RHYTHM_BY_SLIDE = new Map(RHYTHM_ENTRIES.filter(entry => entry && entry.slide_id).map(entry => [entry.slide_id, entry]));
const RENDER_MODES = new Set(["native_chart", "editable_shapes", "source_figure", "svg"]);
const NATIVE_CHART_TYPES = new Set([
  "bar", "bar_chart", "grouped_bar", "grouped_bar_chart", "line", "line_chart",
  "area_chart", "scatter", "scatter_plot", "histogram", "weight_histogram", "simple_time_series", "time_series",
  "grouped_composition", "umap_scatter", "funnel_plot",
]);
const NATIVE_CHART_ROW_LIMIT = 1000;
const NATIVE_ROW_FIELD_ALLOWLIST = Object.freeze({
  weight_histogram: new Set(["bin_low", "bin_high", "count"]),
  grouped_composition: new Set(["group", "cell_type", "cell_count", "percentage"]),
  funnel_plot: new Set(["study_label", "log_effect", "standard_error"]),
});
const PROHIBITED_NATIVE_ROW_FIELDS = new Set([
  ["patient", "id"].join("_"),
  ["synthetic", "id"].join("_"),
  ["cell", "id"].join("_"),
  ["sample", "id"].join("_"),
  ["raw", "records"].join("_"),
]);

const ART_TYPOGRAPHY = ART_DIRECTION_SPEC.typography || ART_DIRECTION_SPEC.fonts || {};
const DECK_LANGUAGE = graph.brief?.language || graph.story_graph?.language || graph.language || "zh-CN";
const HEAD_FONT = ART_TYPOGRAPHY.head_font || ART_TYPOGRAPHY.heading_font || ART_TYPOGRAPHY.headFontFace || "Microsoft YaHei";
const BODY_FONT = ART_TYPOGRAPHY.body_font || ART_TYPOGRAPHY.bodyFontFace || "Microsoft YaHei";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow v2.5";
pptx.subject = "Presentation art direction, rhythm, annotation, and native-chart rendering";
pptx.company = "Academic PPT Workflow";
pptx.lang = DECK_LANGUAGE;
pptx.theme = {
  headFontFace: HEAD_FONT,
  bodyFontFace: BODY_FONT,
  lang: DECK_LANGUAGE,
};

const S = pptx.ShapeType;
const W = 13.333;
const H = 7.5;
const TYPO = { deck: 50, title: 35, sub: 24, body: 18, detail: 16, chart: 15, footer: 10 };
const SAFE_ZONES = Object.freeze({
  // PowerPoint and LibreOffice position CJK glyph ascenders differently when
  // the text frame starts on the decorative rule.  Keep the frame below that
  // rule with a geometric gutter and stop before the content safe zone.
  title: Object.freeze({ x: 0.74, y: 0.46, w: 11.85, h: 0.97, topGutter: 0 }),
  content: Object.freeze({ x: 0.78, y: 1.45, w: 11.77, h: 5.34 }),
  footerExclusion: Object.freeze({ x: 0, y: 7.16, w: W, h: 0.34 }),
});
const LAYOUT_VARIANTS = new Set([
  "minimal_dark", "motif_dark", "dark_band", "tinted_field",
  "hero_left", "hero_right", "hero_number", "full_width",
  "standard", "chart_plus_insight", "chart_plus_secondary_metric",
  "large_left", "large_right", "full_bleed", "top_figure_bottom_callout",
  "balanced", "asymmetric_left", "asymmetric_right", "boundary_contrast",
  "2_column", "3_column", "asymmetric", "delta_focus",
  "horizontal", "vertical", "central_branch", "swimlane",
  "ascending", "central_spine", "certainty_focus",
  "matrix_led", "matrix_plus_mitigation",
  "numbered_takeaways", "hero_number_plus_takeaways", "evidence_boundary_summary",
  "audit_table", "review_register",
]);
const SOURCE_FIGURE_TREATMENTS = new Set(["contain", "cover", "figure_with_callout", "full_bleed"]);
const SHARED = { canvas: "F6F9FB", ink: "17324D", muted: "627D98", white: "FFFFFF", pale: "EAF1F5", warning: "D97706", danger: "B23A48", line: "B9CAD6" };
// Risk-of-bias colors are a presentation semantic contract shared by the
// study-characteristics overview and the full ROB matrix.  The legend is
// populated only from categories that are actually present in the bound rows;
// an unsupported category fails closed instead of being silently colored as
// "Some concerns".
const RISK_OF_BIAS_SEMANTICS = Object.freeze({
  Low: Object.freeze({ color: "4C956C", zh: "低风险" }),
  "Some concerns": Object.freeze({ color: "E9A23B", zh: "存在疑虑" }),
  High: Object.freeze({ color: "C74B50", zh: "高风险" }),
});
const RISK_OF_BIAS_ORDER = Object.freeze(["Low", "Some concerns", "High"]);
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

// Title wrapping is called on every slide. Build the glossary n-gram set once
// per renderer process instead of regenerating identical terms for each title.
const TITLE_GLOSSARY_VALUES = [...new Set(Object.values(ZH_LABELS).map(String))];
const TITLE_GLOSSARY_TERMS = (() => {
  const terms = new Set(TITLE_GLOSSARY_VALUES.filter(term => term.length >= 2));
  for (const glossaryValue of TITLE_GLOSSARY_VALUES) {
    for (const match of glossaryValue.matchAll(/[\p{Script=Han}]{2,}/gu)) {
      const characters = [...match[0]];
      for (let width = 2; width <= Math.min(6, characters.length); width++) {
        for (let start = 0; start + width <= characters.length; start++) {
          terms.add(characters.slice(start, start + width).join(""));
        }
      }
    }
  }
  return [...terms];
})();
const ZH_AUDIENCE_TEXT_REPLACEMENTS = Object.freeze([
  [/TTE at 0\s*[–-]\s*6\s*h/gi, "0–6小时TTE"],
  [/KDIGO stage/gi, "KDIGO分期"],
  [/ICU and 28-day mortality/gi, "ICU与28天死亡"],
  [/Day 28/gi, "第28天"],
  [/72\s*h\b/gi, "72小时"],
  [/\btarget enrollment\b/gi, "目标入组"],
  [/\bprimary endpoint\b/gi, "主要结局"],
  [/\bgrace period\b/gi, "宽限期"],
  [/\bpercentage points?\b/gi, "个百分点"],
  [/\bretained cells?\b/gi, "个保留细胞"],
  [/\bpooled\b/gi, "合并"],
  [/\bstudies\b/gi, "项研究"],
  [/\boutdated\b/gi, "旧版"],
  [/\bthe\b\s*/gi, ""],
  [/(\d+(?:\.\d+)?)\s*h\b/gi, "$1小时"],
]);

function zhLabel(value) {
  return String(value ?? "").split("\n").map(part => {
    const normalized = part.trim();
    return ZH_LABELS[normalized] || normalized;
  }).join("\n");
}

function isChineseDeck() {
  return String(DECK_LANGUAGE).toLowerCase().startsWith("zh");
}

function riskOfBiasSemantic(value, visualId = "visual") {
  const category = String(value ?? "").trim();
  const semantic = RISK_OF_BIAS_SEMANTICS[category];
  if (!semantic) throw new Error(`Unsupported risk_of_bias category "${category}" in ${visualId}`);
  return {
    category,
    color: semantic.color,
    label: isChineseDeck() ? semantic.zh : category,
  };
}

function observedRiskOfBiasLegend(rows, visualId = "visual") {
  const observed = new Set((rows || []).map(row => String(row.risk_of_bias ?? "").trim()));
  for (const category of observed) riskOfBiasSemantic(category, visualId);
  return RISK_OF_BIAS_ORDER
    .filter(category => observed.has(category))
    .map(category => riskOfBiasSemantic(category, visualId));
}

function audienceAxisLabel(rawValue, axisRole, visualType) {
  const raw = String(rawValue ?? "").trim();
  if (!isChineseDeck()) return raw;
  const normalized = raw.toLowerCase();
  if (visualType === "grouped_composition") {
    if (axisRole === "category" || normalized === "cell_type") return "细胞类型";
    if (axisRole === "value" || ["percent", "percentage", "%"].includes(normalized)) return "细胞比例（%）";
  }
  return zhLabel(raw);
}

function annotationAudienceText(annotation) {
  let text = String(annotation.text ?? "");
  if (isChineseDeck() && annotation.entity) {
    const entity = String(annotation.entity);
    text = text.replace(entity, zhLabel(entity));
  }
  text = localizeAudienceText(text);
  if (isChineseDeck()) text = text.replace(/\bpp\b/gi, "个百分点");
  return text;
}

function localizeAudienceText(value) {
  let text = String(value ?? "");
  if (!String(DECK_LANGUAGE).toLowerCase().startsWith("zh")) return text;
  for (const [pattern, replacement] of ZH_AUDIENCE_TEXT_REPLACEMENTS) text = text.replace(pattern, replacement);
  return text;
}

function normalizeScientificTypography(value) {
  return String(value ?? "")
    .replace(/95\s*%\s*CI/gi, "95% CI")
    .replace(/(\d)\s+(小时|分钟|天|周|月|年)/g, "$1$2")
    .replace(/([+\-−]?\d+(?:\.\d+)?)\s*pp\b/gi, "$1 pp")
    .replace(/\s+([，。；：！？,.!?;:])/g, "$1");
}

function bodyProtectedRanges(value) {
  const ranges = [];
  const addMatches = pattern => {
    for (const match of value.matchAll(pattern)) ranges.push([match.index, match.index + match[0].length]);
  };
  addMatches(/[A-Za-z][A-Za-z0-9._%+\-–]*/g);
  addMatches(/[+\-−]?\d+(?:\.\d+)?(?:小时|分钟|天|周|月|年|个百分点|%|％|pp)/giu);
  addMatches(/(?:95\s*%\s*)?(?:OR|HR|RR|CI|FDR|NES|SMD|I²|I2)(?:\s*[=:<>≤≥]?\s*[+\-−]?\d+(?:\.\d+)?(?:%|％|pp)?)?/giu);
  return ranges;
}

function wrapAudienceBodyText(value, lineWidthBudget) {
  let text = normalizeScientificTypography(localizeAudienceText(value));
  if (titleDisplayWidth(text) > lineWidthBudget) {
    // A comparison connector should introduce the second strategy instead of
    // being stranded at the end of the first line. This is content-structural,
    // not project- or page-specific.
    text = text.replace(/\s+(vs\.?|versus)\s+/i, "\n$1 ");
  }
  const lines = [];
  for (const paragraph of text.split("\n")) {
    let remaining = paragraph.trim();
    while (remaining && titleDisplayWidth(remaining) > lineWidthBudget) {
      const characters = [...remaining];
      const protectedRanges = bodyProtectedRanges(remaining);
      let width = 0;
      let end = 0;
      while (end < characters.length) {
        const next = /[^\u0000-\u00ff]/.test(characters[end]) ? 2 : 1;
        if (width + next > lineWidthBudget) break;
        width += next;
        end += 1;
      }
      if (end <= 0) end = 1;
      const crossing = protectedRanges.find(([start, finish]) => start < end && end < finish);
      if (crossing) end = crossing[0] > 0 ? crossing[0] : crossing[1];
      const minimumSemanticBreak = Math.max(1, Math.floor(end * 0.55));
      for (let index = end; index >= minimumSemanticBreak; index--) {
        if (/\s|[，；：、,;:]/.test(characters[index - 1] || "")) {
          if (!protectedRanges.some(([start, finish]) => start < index && index < finish)) {
            end = index;
            break;
          }
        }
      }
      const lineText = characters.slice(0, end).join("").trim();
      if (!lineText) throw new Error("Audience body text has no token-safe line break");
      lines.push(lineText);
      remaining = characters.slice(end).join("").trim();
    }
    if (remaining) lines.push(remaining);
  }
  // Punctuation belongs with the preceding sentence; never emit a punctuation-
  // only line even if a renderer applies different CJK metrics.
  for (let index = 1; index < lines.length; index++) {
    if (/^[，。；：！？,.!?;:]+$/.test(lines[index])) {
      lines[index - 1] += lines[index];
      lines.splice(index, 1);
      index -= 1;
    }
  }
  return lines.join("\n");
}

function needsTokenSafeBodyWrap(value, lineWidthBudget) {
  const text = String(value ?? "");
  if (titleDisplayWidth(text) <= lineWidthBudget) return false;
  return /(?:[+\-−]?\d+(?:\.\d+)?(?:小时|分钟|天|周|月|年|个百分点|%|％|pp)|\b(?:vs\.?|versus)\b|[，。；：！？,.!?;:]$)/iu.test(text);
}

function firstDefined(...values) { return values.find(value => value !== undefined && value !== null && value !== ""); }
function clamp(value, minimum, maximum) { return Math.min(maximum, Math.max(minimum, Number(value)));
}
function colorToken(value, fallback) {
  const token = String(value || "").replace(/^#/, "");
  return /^[0-9A-F]{6}$/i.test(token) ? token.toUpperCase() : fallback;
}

function tintSuitability(hex) {
  const token = colorToken(hex, null);
  if (!token) return null;
  const channels = [0, 2, 4].map(offset => parseInt(token.slice(offset, offset + 2), 16) / 255);
  const maximum = Math.max(...channels), minimum = Math.min(...channels);
  const lightness = (maximum + minimum) / 2;
  const saturation = maximum === minimum ? 0 : (maximum - minimum) / (1 - Math.abs(2 * lightness - 1));
  if (lightness < 0.78 || saturation > 0.50) return null;
  // Prefer an airy but visibly tinted field instead of either a saturated
  // domain accent or a near-white duplicate of the main canvas.
  return { token, score: Math.abs(lightness - 0.91) + saturation * 0.48 };
}

function selectTintColor(palette, fallback, rank = 0) {
  const ranked = [...new Set([...(palette || []), fallback])]
    .map(tintSuitability)
    .filter(Boolean)
    .sort((a, b) => a.score - b.score || a.token.localeCompare(b.token));
  return ranked[Math.min(rank, Math.max(0, ranked.length - 1))]?.token || colorToken(fallback, SHARED.pale);
}

const PROFILE_CACHE = new Map();
function profile(slideSpec) {
  const cacheKey = String(slideSpec.design_profile || "clinical_methods");
  if (PROFILE_CACHE.has(cacheKey)) return PROFILE_CACHE.get(cacheKey);
  const base = PROFILES[slideSpec.design_profile] || PROFILES.clinical_methods;
  const primaryPalette = ART_DIRECTION_SPEC.primary_palette || [];
  const secondaryPalette = ART_DIRECTION_SPEC.secondary_palette || [];
  const semanticColors = ART_DIRECTION_SPEC.semantic_colors || {};
  const artPalette = ART_DIRECTION_SPEC.palette || ART_DIRECTION_SPEC.colors || {};
  const scopedPalette = (ART_DIRECTION_SPEC.domain_profiles || {})[slideSpec.design_profile] || {};
  const merged = {
    primary: colorToken(firstDefined(scopedPalette.primary, artPalette.primary, primaryPalette[0]), base.primary),
    accent: colorToken(firstDefined(scopedPalette.accent, artPalette.accent, primaryPalette[1], semanticColors.planned), base.accent),
    secondary: colorToken(firstDefined(scopedPalette.secondary, artPalette.secondary, primaryPalette[2], semanticColors.review), base.secondary),
    negative: colorToken(firstDefined(scopedPalette.negative, artPalette.negative, semanticColors.negative), base.negative || SHARED.danger),
    prediction: colorToken(firstDefined(scopedPalette.prediction, artPalette.prediction, semanticColors.prediction), base.secondary),
    soft: colorToken(firstDefined(scopedPalette.soft, artPalette.soft, selectTintColor(secondaryPalette, base.soft, 0)), base.soft),
    soft2: colorToken(firstDefined(scopedPalette.soft2, artPalette.soft2, selectTintColor(secondaryPalette, base.soft2, 1)), base.soft2),
  };
  PROFILE_CACHE.set(cacheKey, merged);
  return merged;
}

function normalizeAnnotation(item, index) {
  if (typeof item === "string") return { annotation_id: `annotation-${index + 1}`, role: index === 0 ? "primary" : "secondary", text: item };
  if (!item || typeof item !== "object") return null;
  const roleValue = String(firstDefined(item.priority, item.role, item.annotation_role, item.type, index === 0 ? "primary" : "secondary")).toLowerCase();
  const primaryRoles = new Set(["primary", "takeaway", "key_message", "hero", "insight"]);
  const secondaryRoles = new Set(["secondary", "caveat", "boundary", "qualification", "context", "warning"]);
  const role = primaryRoles.has(roleValue) ? "primary" : secondaryRoles.has(roleValue) ? "secondary" : (index === 0 ? "primary" : "secondary");
  const text = firstDefined(item.text, item.message, item.body, item.label);
  if (!text) return null;
  return {
    annotation_id: item.annotation_id || item.id || `annotation-${index + 1}`,
    annotation_type: item.annotation_type || item.type || "direct_annotation",
    role,
    title: item.title || "",
    text: String(text),
    entity: item.entity || "",
    value: item.value,
    source_bindings: item.source_bindings || [],
    visual_id: item.visual_id || item._visual_id || null,
    visual_type: item.visual_type || item._visual_type || null,
  };
}

function annotationNeedsExternalPanel(annotation) {
  const inlineTypes = new Set(["direct_label", "reference_line_label", "threshold_annotation", "highlight_band"]);
  if (inlineTypes.has(annotation.annotation_type)) return false;
  const integrated = new Set([
    "absolute_risk_comparison:delta_annotation",
    "pooled_evidence_panel:key_value_callout",
    "pooled_evidence_panel:confidence_annotation",
    "sample_size_waterfall:bracket_comparison",
  ]);
  return !integrated.has(`${annotation.visual_type}:${annotation.annotation_type}`);
}

const PRESENTATION_FIELD_CACHE = new Map();
const PRESENTATION_GEOMETRY_CACHE = new Map();
function presentationFields(spec, visual = null) {
  const visualBelongsToSlide = Boolean(visual && (spec.visual_specs || []).some(item => item.visual_id === visual.visual_id));
  const cacheable = !visual || visualBelongsToSlide;
  if (cacheable && PRESENTATION_FIELD_CACHE.has(spec.slide_id)) return PRESENTATION_FIELD_CACHE.get(spec.slide_id);
  const rhythm = RHYTHM_BY_SLIDE.get(spec.slide_id) || {};
  const presentation = spec.presentation_fields || spec.presentation || {};
  const rawVariant = String(firstDefined(spec.layout_variant, presentation.layout_variant, rhythm.layout_variant, ART_DIRECTION_SPEC.default_layout_variant, "standard"));
  if (!LAYOUT_VARIANTS.has(rawVariant)) throw new Error(`Unsupported layout_variant ${rawVariant} on ${spec.slide_id}`);
  const slideVisualAnnotations = (spec.visual_specs || []).flatMap(item =>
    (item.annotation_specs || item.annotations || []).map(annotation => ({
      ...annotation,
      _visual_id: item.visual_id,
      _visual_type: item.render_visual_type || item.visual_type,
    })),
  );
  const adHocVisualAnnotations = visual && !visualBelongsToSlide ? (visual.annotation_specs || visual.annotations || []) : [];
  const rawAnnotations = firstDefined(
    slideVisualAnnotations.length ? slideVisualAnnotations : undefined,
    adHocVisualAnnotations.length ? adHocVisualAnnotations : undefined,
    spec.annotation_specs,
    spec.annotations,
    presentation.annotation_specs,
    presentation.annotations,
    rhythm.annotation_specs,
    [],
  );
  const annotationList = (Array.isArray(rawAnnotations) ? rawAnnotations : [rawAnnotations]).map(normalizeAnnotation).filter(Boolean);
  const panelAnnotations = annotationList.filter(annotationNeedsExternalPanel);
  const primary = panelAnnotations.find(item => item.role === "primary") || null;
  const secondary = panelAnnotations.find(item => item.role === "secondary") || null;
  const annotationWarnings = [];
  if (annotationList.filter(item => item.role === "primary").length > 1) annotationWarnings.push("annotation_primary_truncated_to_one");
  if (annotationList.filter(item => item.role === "secondary").length > 1) annotationWarnings.push("annotation_secondary_truncated_to_one");
  const heroContract = spec.hero_visual || presentation.hero_visual || {};
  const heroEnabled = Boolean(firstDefined(heroContract.enabled, rhythm.is_hero, rhythm.hero_priority === "required", false));
  const requestedHeroShare = heroEnabled
    ? clamp(firstDefined(heroContract.target_area_ratio, spec.hero_share, spec.hero_ratio, presentation.hero_share, presentation.hero_ratio, rhythm.hero_area_ratio, ART_DIRECTION_SPEC.hero_share, 0.52), 0.35, 0.65)
    : 0.62;
  const backgroundVariant = String(firstDefined(spec.background_variant, presentation.background_variant, rhythm.background_variant, "light"));
  if (!["light", "tinted", "dark_emphasis"].includes(backgroundVariant)) throw new Error(`Unsupported background_variant ${backgroundVariant} on ${spec.slide_id}`);
  const backgroundTreatment = firstDefined(spec.background_treatment, presentation.background_treatment, rhythm.background_treatment, ART_DIRECTION_SPEC.background_treatment, "canvas");
  const sourceFigureTreatment = String(firstDefined(spec.figure_treatment, presentation.figure_treatment, rhythm.figure_treatment, ART_DIRECTION_SPEC.figure_treatment, "figure_with_callout"));
  if (!SOURCE_FIGURE_TREATMENTS.has(sourceFigureTreatment)) throw new Error(`Unsupported source_figure_treatment ${sourceFigureTreatment} on ${spec.slide_id}`);
  const fields = {
    layoutVariant: rawVariant,
    heroEnabled,
    requestedHeroShare,
    backgroundVariant,
    backgroundTreatment,
    sourceFigureTreatment,
    annotations: { primary, secondary, all: annotationList },
    annotationWarnings,
    rhythmBeat: firstDefined(rhythm.rhythm_beat, rhythm.beat, rhythm.slide_role, null),
    rhythmIndex: firstDefined(rhythm.rhythm_index, rhythm.index, null),
    renderMode: firstDefined(spec.render_mode, presentation.render_mode, rhythm.render_mode, null),
    visualIntensity: firstDefined(spec.visual_intensity, rhythm.visual_intensity, null),
    densityTarget: firstDefined(spec.density_target, rhythm.density_target, null),
  };
  if (cacheable) PRESENTATION_FIELD_CACHE.set(spec.slide_id, fields);
  return fields;
}

function presentationGeometry(spec, visual = null) {
  const visualBelongsToSlide = Boolean(visual && (spec.visual_specs || []).some(item => item.visual_id === visual.visual_id));
  const cacheable = !visual || visualBelongsToSlide;
  if (cacheable && PRESENTATION_GEOMETRY_CACHE.has(spec.slide_id)) return PRESENTATION_GEOMETRY_CACHE.get(spec.slide_id);
  const fields = presentationFields(spec, visual);
  const hasAnnotations = Boolean(fields.annotations.primary || fields.annotations.secondary);
  const content = { ...SAFE_ZONES.content };
  let targetShare = fields.requestedHeroShare;
  const warnings = [...fields.annotationWarnings];
  const shouldSplit = hasAnnotations || ["hero_left", "hero_right", "large_left", "large_right", "asymmetric_left", "asymmetric_right", "chart_plus_insight", "chart_plus_secondary_metric"].includes(fields.layoutVariant);
  let hero = { ...content };
  let annotation = null;
  if (shouldSplit) {
    const gutter = 0.30;
    const widthRatio = clamp(targetShare, 0.35, 0.65);
    const heroWidth = content.w * widthRatio;
    const annotationWidth = content.w - heroWidth - gutter;
    if (annotationWidth < 1.85) throw new Error(`Annotation column is below safe width on ${spec.slide_id}`);
    const heroOnRight = ["hero_right", "large_right", "asymmetric_right"].includes(fields.layoutVariant);
    hero = { x: heroOnRight ? content.x + content.w - heroWidth : content.x, y: content.y, w: heroWidth, h: content.h };
    annotation = { x: heroOnRight ? content.x : content.x + heroWidth + gutter, y: content.y, w: annotationWidth, h: content.h };
  } else if (fields.heroEnabled) {
    if (["full_width", "full_bleed", "matrix_led"].includes(fields.layoutVariant)) {
      const heroHeight = content.h * targetShare;
      hero = { x: content.x, y: content.y + (content.h - heroHeight) / 2, w: content.w, h: heroHeight };
    } else {
      const heroWidth = content.w * targetShare;
      hero = { x: content.x + (content.w - heroWidth) / 2, y: content.y, w: heroWidth, h: content.h };
    }
  }
  const heroShare = hero.w * hero.h / (content.w * content.h);
  if (fields.heroEnabled && (heroShare < 0.35 - 0.001 || heroShare > 0.65 + 0.001)) throw new Error(`Hero content-area share ${heroShare.toFixed(3)} violates 0.35-0.65 on ${spec.slide_id}`);
  const geometry = { fields, hero, annotation, heroShare, warnings };
  if (cacheable) PRESENTATION_GEOMETRY_CACHE.set(spec.slide_id, geometry);
  return geometry;
}

function boundsWithin(bounds, zone, tolerance = 0.002) {
  return bounds.x >= zone.x - tolerance && bounds.y >= zone.y - tolerance && bounds.x + bounds.w <= zone.x + zone.w + tolerance && bounds.y + bounds.h <= zone.y + zone.h + tolerance;
}

function assertSafeContent(bounds, label) {
  if (!boundsWithin(bounds, SAFE_ZONES.content)) throw new Error(`${label} leaves the content safe zone`);
  if (bounds.y + bounds.h > SAFE_ZONES.footerExclusion.y) throw new Error(`${label} enters the footer exclusion zone`);
}

function applySlideBackground(slide, spec) {
  const fields = presentationFields(spec);
  const treatment = fields.backgroundTreatment;
  const p = profile(spec);
  const type = typeof treatment === "string" ? treatment : (treatment.type || "canvas");
  const variantColor = fields.backgroundVariant === "dark_emphasis" ? p.primary : fields.backgroundVariant === "tinted" ? p.soft : SHARED.canvas;
  const color = colorToken(typeof treatment === "object" ? treatment.color : null, variantColor);
  slide.background = { color };
  if (fields.backgroundVariant === "dark_emphasis" && spec.slide_role !== "cover") {
    slide.addShape(S.roundRect, { x: SAFE_ZONES.content.x - 0.08, y: SAFE_ZONES.content.y - 0.07, w: SAFE_ZONES.content.w + 0.16, h: SAFE_ZONES.content.h + 0.10, rectRadius: 0.03, fill: { color: "F9FBFC" }, line: { color: p.accent, transparency: 40, width: 1 }, name: `${spec.slide_id}:dark-content-field` });
  }
  if (type === "accent_band") {
    slide.addShape(S.rect, { x: 0, y: 0, w: 0.18, h: SAFE_ZONES.footerExclusion.y, fill: { color: p.accent }, line: { color: p.accent }, name: `${spec.slide_id}:background-band` });
  } else if (!["canvas", "soft", "solid", "light", "tinted", "dark_emphasis"].includes(type)) {
    throw new Error(`Unsupported non-decorative background treatment ${type} on ${spec.slide_id}`);
  }
}

function masterName(family) { return `V25_${family}`; }
for (const family of FAMILIES) {
  const objects = family === "cover" ? [] : [
    { line: { x: 0.72, y: 0.30, w: 0.44, h: 0, line: { color: "2A9D8F", width: 4 } } },
  ];
  pptx.defineSlideMaster({
    title: masterName(family),
    background: { color: family === "cover" ? "17324D" : SHARED.canvas },
    objects,
    // Footer band, source label, and page number are deliberately slide-level
    // objects so PowerPoint and LibreOffice receive identical geometry.
    slideNumber: undefined,
  });
}

function addText(slide, text, x, y, w, h, size = TYPO.body, options = {}) {
  slide.addText(String(text), {
    x, y, w, h,
    fontFace: options.fontFace || BODY_FONT,
    fontSize: size,
    bold: Boolean(options.bold),
    color: options.color || SHARED.ink,
    margin: options.margin ?? 0,
    align: options.align || "left",
    valign: options.valign || "mid",
    rotate: options.rotate,
    wrap: options.wrap,
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
    line: { color, width, dashType: options.dash, beginArrowType: options.beginArrowType, endArrowType: options.endArrowType },
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

function titleProtectedRanges(value) {
  const ranges = [];
  const addMatches = pattern => {
    for (const match of value.matchAll(pattern)) ranges.push([match.index, match.index + match[0].length]);
  };
  addMatches(/[A-Za-z][A-Za-z0-9._%+\-–]*/g);
  // Preserve quantitative tokens even when no unit follows.  A title must
  // never turn 4.3 into `4.` / `3`, or detach a signed percentage/pp value.
  addMatches(/[+\-−]?\d+(?:\.\d+)+(?:%|％|pp)?|[+\-−]?\d+(?:%|％|pp|个百分点)/g);
  addMatches(/\d+(?:\.\d+)?(?:小时|分钟|天|周|月|年|%|％)/g);
  // Statistical abbreviations and an attached value form one audience-facing
  // token.  The input value remains untouched; this only constrains wrapping.
  addMatches(/(?:95\s*%\s*)?(?:OR|HR|RR|CI|FDR|NES|SMD|I²|I2)(?:\s*[=:<>≤≥]?\s*[+\-−]?\d+(?:\.\d+)?(?:%|％|pp)?)?/giu);

  // Reuse the precomputed localization glossary as a general token source
  // instead of hard-coding a project title. CJK n-grams protect short medical
  // terms such as a two-character intervention name in a longer title.
  for (const term of TITLE_GLOSSARY_TERMS) {
    if (!value.includes(term)) continue;
    let start = value.indexOf(term);
    while (start >= 0) {
      ranges.push([start, start + term.length]);
      start = value.indexOf(term, start + term.length);
    }
  }
  return ranges;
}

function wrapTitleText(text, lineWidthBudget = 38) {
  const value = String(text);
  if (titleDisplayWidth(value) <= lineWidthBudget) return value;
  const characters = [...value];
  const protectedRanges = titleProtectedRanges(value);
  const candidates = [];
  for (let index = 1; index < characters.length; index++) {
    if (protectedRanges.some(([start, end]) => start < index && index < end)) continue;
    const left = characters.slice(0, index).join("").trim();
    const right = characters.slice(index).join("").trim();
    if (!left || !right) continue;
    if (/^[，。；：、！？）】》〉,.!?;:]/.test(right) || /[（【《〈]$/.test(left)) continue;
    const leftWidth = titleDisplayWidth(left), rightWidth = titleDisplayWidth(right);
    const overflow = Math.max(0, leftWidth - lineWidthBudget) + Math.max(0, rightWidth - lineWidthBudget);
    const punctuationBreak = /[：；，、—:;,-\s]/.test(characters[index - 1]);
    candidates.push({ index, overflow, balance: Math.abs(leftWidth - rightWidth), punctuationBreak });
  }
  if (!candidates.length) throw new Error("Title has no token-safe two-line break");
  const withinBudget = candidates.filter(candidate => candidate.overflow === 0);
  const pool = withinBudget.length ? withinBudget : candidates;
  // Within the width budget, a semantic punctuation boundary is safer than a
  // mathematically perfect split through a phrase.  Balance remains the
  // deterministic tie-breaker.
  pool.sort((a, b) => a.overflow - b.overflow || Number(b.punctuationBreak) - Number(a.punctuationBreak) || a.balance - b.balance);
  const best = pool[0].index;
  return `${characters.slice(0, best).join("").trim()}\n${characters.slice(best).join("").trim()}`;
}

function truncateDisplayText(text, maximumWidth) {
  const value = String(text ?? "").trim();
  if (titleDisplayWidth(value) <= maximumWidth) return value;
  const suffix = "…";
  const budget = Math.max(1, maximumWidth - titleDisplayWidth(suffix));
  let width = 0;
  const kept = [];
  for (const character of [...value]) {
    const next = /[^\u0000-\u00ff]/.test(character) ? 2 : 1;
    if (width + next > budget) break;
    kept.push(character);
    width += next;
  }
  return `${kept.join("").trim()}${suffix}`;
}

function compactFooterSource(labels, maximumWidth = 56) {
  const unique = [...new Set(labels.map(label => String(label || "输入材料").trim()).filter(Boolean))];
  if (!unique.length) return "";
  let output = "";
  let included = 0;
  for (const label of unique) {
    const separator = included ? "、" : "";
    const remainingAfter = unique.length - included - 1;
    const suffix = remainingAfter > 0 ? ` 等${unique.length}项` : "";
    const available = maximumWidth - titleDisplayWidth(output + separator + suffix);
    if (available <= 2) break;
    const candidate = truncateDisplayText(label, available);
    output += `${separator}${candidate}`;
    included += 1;
    if (candidate !== label) break;
    if (titleDisplayWidth(output + (included < unique.length ? ` 等${unique.length}项` : "")) >= maximumWidth) break;
  }
  if (included < unique.length && !output.endsWith("…")) {
    const suffix = ` 等${unique.length}项`;
    output = `${truncateDisplayText(output, maximumWidth - titleDisplayWidth(suffix))}${suffix}`;
  }
  return output;
}

function title(slide, spec) {
  const wrapped = wrapTitleText(localizeAudienceText(spec.title));
  const twoLines = wrapped.includes("\n");
  const zone = SAFE_ZONES.title;
  const dark = presentationFields(spec).backgroundVariant === "dark_emphasis";
  if (zone.y + zone.h > SAFE_ZONES.content.y) throw new Error(`Title safe zone enters content on ${spec.slide_id}`);
  addText(slide, wrapped, zone.x, zone.y, zone.w, zone.h, twoLines ? 30 : TYPO.title, {
    bold: true,
    color: dark ? SHARED.white : profile(spec).primary,
    fontFace: HEAD_FONT,
    name: `${spec.slide_id}:title`,
    valign: "top",
    margin: [zone.topGutter, 0, 0, 0],
    wrap: false,
  });
  const label = ROLE_ZH[spec.slide_role] || spec.slide_role;
  const reservesUpperBand = (spec.visual_specs || []).some(visual => visual.visual_type === "assessment_timeline");
  // A two-line title already consumes the navigation-label row.  Omitting the
  // redundant role label preserves a measured gutter without shrinking the
  // title below the 30 pt contract or pushing content into its safe zone.
  if (label && !twoLines && !reservesUpperBand) addText(slide, label, 0.76, 1.17, 1.55, 0.24, TYPO.chart, { bold: true, color: profile(spec).accent, name: `${spec.slide_id}:section` });
}

function footer(slide, spec, pageNumber) {
  const labels = [...new Set((spec.source_bindings || []).map(binding => binding.short_label_zh || "输入材料"))];
  const text = compactFooterSource(labels, 52) || "未登记";
  slide.addShape(S.rect, { x: 0, y: 7.12, w: W, h: 0.38, fill: { color: "17324D" }, line: { color: "17324D" }, name: `${spec.slide_id}:footer-band` });
  // Keep the prefix and short source label in one text object.  Separate
  // adjacent boxes render inconsistently in PowerPoint on some dark canvases
  // even when their OOXML geometry is valid.
  addText(slide, `来源： ${text}`, 0.72, 7.155, 11.28, 0.28, 11, { color: SHARED.white, name: `${spec.slide_id}:footer-label`, wrap: false, valign: "mid" });
  addText(slide, String(pageNumber), 12.40, 7.155, 0.46, 0.28, 11, { color: SHARED.white, fontFace: "Arial", name: `${spec.slide_id}:page-number`, wrap: false, valign: "mid", align: "right" });
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
  lines.push("", "[Presentation Annotations]");
  for (const visual of spec.visual_specs || []) {
    for (const item of visual.annotation_specs || []) lines.push(`${visual.visual_id} | ${item.priority || item.role || "secondary"} | ${item.annotation_type || "direct_annotation"} | ${item.text}`);
  }
  return lines.join("\n");
}

function imageContain(slide, file, x, y, w, h, name) {
  const dimensions = imageSize(fs.readFileSync(file));
  const scale = Math.min(w / dimensions.width, h / dimensions.height);
  const iw = dimensions.width * scale;
  const ih = dimensions.height * scale;
  slide.addImage({ path: file, x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih, name });
}

function imageCover(slide, file, x, y, w, h, name) {
  slide.addImage({ path: file, x, y, w, h, sizing: { type: "cover", w, h }, name });
}

function resolveAssetPath(rawPath) {
  if (!rawPath) throw new Error("Source-figure VisualSpec does not declare an asset path");
  const value = String(rawPath);
  if (/^[a-z]+:\/\//i.test(value)) throw new Error(`External asset URLs are prohibited: ${value}`);
  const candidates = path.isAbsolute(value)
    ? [value]
    : [path.resolve(path.dirname(graphPath), value), path.resolve(process.cwd(), value)];
  const resolved = candidates.find(candidate => fs.existsSync(candidate) && fs.statSync(candidate).isFile());
  if (!resolved) throw new Error(`Registered visual asset was not found: ${value}`);
  return resolved;
}

function resolveRenderMode(visual, spec) {
  const presentation = presentationFields(spec);
  const rawAsset = firstDefined(visual.presentation_payload?.svg_path, visual.presentation_payload?.image_path, visual.presentation_payload?.asset_path, visual.payload?.svg_path, visual.payload?.image_path, visual.payload?.asset_path, visual.asset_path, "");
  const renderVisualType = String(visual.render_visual_type || visual.visual_type);
  const inferred = ["umap_figure", "source_figure"].includes(renderVisualType)
    ? (String(rawAsset).toLowerCase().endsWith(".svg") ? "svg" : "source_figure")
    : NATIVE_CHART_TYPES.has(renderVisualType) ? "native_chart" : "editable_shapes";
  const mode = String(firstDefined(visual.render_mode, visual.presentation?.render_mode, presentation.renderMode, inferred));
  if (!RENDER_MODES.has(mode)) throw new Error(`Unsupported render_mode ${mode} for ${visual.visual_id}`);
  if (mode === "native_chart" && !NATIVE_CHART_TYPES.has(renderVisualType)) {
    throw new Error(`VisualSpec ${visual.visual_id} requests native_chart for unsupported type ${renderVisualType}`);
  }
  if (mode === "svg" && !String(rawAsset).toLowerCase().endsWith(".svg")) {
    throw new Error(`VisualSpec ${visual.visual_id} requests svg mode without an SVG asset`);
  }
  return mode;
}

function aggregateRows(visual) {
  const renderVisualType = String(visual.render_visual_type || visual.visual_type);
  const presentationRows = firstDefined(visual.presentation_payload?.aggregate_rows, visual.presentation_payload?.rows);
  const rows = firstDefined(
    presentationRows,
    visual.payload?.aggregate_rows,
    visual.payload?.rows,
    renderVisualType.includes("histogram") ? visual.presentation_payload?.bins : undefined,
    renderVisualType.includes("histogram") ? visual.payload?.bins : undefined,
  );
  if (!Array.isArray(rows)) throw new Error(`Native chart ${visual.visual_id} requires aggregate rows in VisualSpec payload`);
  const isPresentationRebuild = Array.isArray(presentationRows);
  const expectedAggregateCount = Number(isPresentationRebuild ? (visual.native_chart_privacy?.embedded_row_count ?? rows.length) : (visual.expected_row_count ?? rows.length));
  if (rows.length !== expectedAggregateCount) throw new Error(`Native chart aggregate-row mismatch for ${visual.visual_id}: ${rows.length} != ${expectedAggregateCount}`);
  const expectedEvidenceCount = Number(visual.expected_row_count ?? (visual.expected_row_keys || []).length);
  if ((visual.expected_row_keys || []).length !== expectedEvidenceCount) throw new Error(`Native chart ${visual.visual_id} lacks one expected_row_key per renderer-evidence row`);
  return rows.map(row => ({ ...row }));
}

function visualField(visual, key, fallback = "") {
  return firstDefined(visual.presentation_payload?.[key], visual[key], fallback);
}

function validateNativeChartPrivacy(visual, rows) {
  const privacy = visual.native_chart_privacy || {};
  for (const flag of ["patient_level_data_embedded", "cell_level_data_embedded", "absolute_source_path_embedded", "unrelated_fields_embedded"]) {
    if (privacy[flag] === true) throw new Error(`Native chart privacy violation ${flag} in ${visual.visual_id}`);
  }
  if (privacy.embedded_row_count !== undefined && Number(privacy.embedded_row_count) !== rows.length) {
    throw new Error(`Native chart privacy row count mismatch for ${visual.visual_id}`);
  }
  if (rows.length > NATIVE_CHART_ROW_LIMIT) throw new Error(`Native chart ${visual.visual_id} exceeds the aggregate row limit (${rows.length} > ${NATIVE_CHART_ROW_LIMIT})`);
  const renderVisualType = String(visual.render_visual_type || visual.visual_type);
  const actualFields = [...new Set(rows.flatMap(row => Object.keys(row || {})))].sort();
  const prohibitedHits = actualFields.filter(field => PROHIBITED_NATIVE_ROW_FIELDS.has(field.toLowerCase()));
  if (prohibitedHits.length) throw new Error(`Native chart ${visual.visual_id} includes prohibited identifier fields: ${prohibitedHits.join(",")}`);
  const absolutePathHits = [];
  rows.forEach((row, rowIndex) => Object.entries(row || {}).forEach(([field, value]) => {
    if (typeof value === "string" && (/^[A-Za-z]:[\\/]/.test(value) || /^\\\\/.test(value) || /^\//.test(value))) absolutePathHits.push(`${rowIndex}:${field}`);
  }));
  if (absolutePathHits.length) throw new Error(`Native chart ${visual.visual_id} includes absolute path values: ${absolutePathHits.join(",")}`);
  let allowlist = NATIVE_ROW_FIELD_ALLOWLIST[renderVisualType];
  let aggregateContract = "visual_type_allowlist";
  if (!allowlist) {
    const explicitAggregateRows = visual.presentation_payload?.aggregate_rows || visual.payload?.aggregate_rows;
    if (!Array.isArray(explicitAggregateRows)) throw new Error(`Native chart ${visual.visual_id} requires an explicit aggregate_rows contract`);
    const declaredFields = visual.native_chart_privacy?.embedded_fields || visual.presentation_payload?.embedded_fields || visual.payload?.aggregate_fields || [];
    if (!Array.isArray(declaredFields) || !declaredFields.length) throw new Error(`Native chart ${visual.visual_id} requires an explicit aggregate field allowlist`);
    allowlist = new Set(declaredFields.map(String));
    aggregateContract = "explicit_aggregate_rows";
  }
  const unexpectedFields = actualFields.filter(field => !allowlist.has(field));
  if (unexpectedFields.length) throw new Error(`Native chart ${visual.visual_id} includes fields outside its allowlist: ${unexpectedFields.join(",")}`);
  return {
    status: "PASS",
    data_scope: privacy.data_scope || aggregateContract,
    actual_row_count: rows.length,
    actual_fields: actualFields,
    allowed_fields: [...allowlist].sort(),
    aggregate_contract: aggregateContract,
    prohibited_field_hits: [],
    absolute_path_hits: [],
    row_limit: NATIVE_CHART_ROW_LIMIT,
  };
}

function numericValue(value, field, visual) {
  const number = Number(value);
  if (!Number.isFinite(number)) throw new Error(`Native chart ${visual.visual_id} has non-numeric ${field}: ${value}`);
  return number;
}

function nativeChartData(visual) {
  const rows = aggregateRows(visual);
  const privacyValidation = validateNativeChartPrivacy(visual, rows);
  const type = String(visual.render_visual_type || visual.visual_type);
  const analysisName = String(firstDefined(visual.analysis_label, visual.outcome, visual.semantic_role, visual.visual_id));
  if (["histogram", "weight_histogram"].includes(type)) {
    const valueField = visualField(visual, "y_field", "count");
    const labels = rows.map(row => {
      if (row.bin !== undefined) return String(row.bin);
      if (row.bin_low !== undefined && row.bin_high !== undefined) return `${Number(row.bin_low).toFixed(2)}–${Number(row.bin_high).toFixed(2)}`;
      throw new Error(`Histogram ${visual.visual_id} requires bin or bin_low/bin_high`);
    });
    return {
      chartType: "histogram",
      // PptxGenJS uses ChartType.bar for both horizontal bars and columns;
      // barDir="col" selects a column chart.  ChartType.column is undefined
      // and produces axis-only OOXML that PowerPoint rejects as corrupt.
      pptxType: pptx.ChartType.bar,
      rows,
      series: [{ name: analysisName, labels, values: rows.map(row => numericValue(row[valueField], valueField, visual)) }],
      categoryCount: labels.length,
      workbookFields: rows.some(row => row.bin !== undefined) ? ["bin", valueField] : ["bin_low", "bin_high", valueField],
      privacyValidation,
    };
  }

  const isScatter = ["scatter", "scatter_plot", "umap_scatter", "funnel_plot"].includes(type);
  const isTimeSeries = ["simple_time_series", "time_series"].includes(type);
  const xField = isTimeSeries ? visualField(visual, "time_field") : visualField(visual, "x_field");
  const labelField = visualField(visual, "label_field");
  const yField = visualField(visual, "y_field") || ((!isScatter && labelField && xField !== labelField) ? xField : "");
  if (!xField || !yField) throw new Error(`Native chart ${visual.visual_id} requires explicit x/time and y field roles`);
  const groupField = visualField(visual, "group_field");

  if (type === "funnel_plot") {
    if (!labelField) throw new Error(`Funnel plot ${visual.visual_id} requires an aggregate study label field`);
    const xValues = rows.map(row => Number(numericValue(row[xField], xField, visual).toFixed(3)));
    const series = [{ name: xField, values: xValues }];
    rows.forEach((row, index) => {
      const values = Array(rows.length).fill(null);
      values[index] = Number(numericValue(row[yField], yField, visual).toFixed(3));
      series.push({ name: String(row[labelField]), values });
    });
    return { chartType: "funnel_plot", pptxType: pptx.ChartType.scatter, rows, series, categoryCount: rows.length, workbookFields: [labelField, xField, yField], privacyValidation };
  }

  if (isScatter) {
    const groups = groupField ? [...new Set(rows.map(row => String(row[groupField])))] : [analysisName];
    const xValues = [...new Set(rows.map(row => numericValue(row[xField], xField, visual)))].sort((a, b) => a - b);
    const series = [{ name: xField, values: xValues }];
    for (const group of groups) {
      const subset = groupField ? rows.filter(row => String(row[groupField]) === group) : rows;
      const byX = new Map(subset.map(row => [numericValue(row[xField], xField, visual), numericValue(row[yField], yField, visual)]));
      if (xValues.some(value => !byX.has(value))) throw new Error(`Scatter ${visual.visual_id} has incomplete x values in group ${group}`);
      series.push({ name: group, values: xValues.map(value => byX.get(value)) });
    }
    return { chartType: "scatter", pptxType: pptx.ChartType.scatter, rows, series, categoryCount: xValues.length, workbookFields: [xField, yField, ...(groupField ? [groupField] : [])], privacyValidation };
  }

  const categoryField = isTimeSeries ? visualField(visual, "time_field") : (labelField || xField);
  if (!categoryField) throw new Error(`Native chart ${visual.visual_id} requires a category or time field`);
  const categories = [...new Set(rows.map(row => String(row[categoryField])))];
  const groups = groupField ? [...new Set(rows.map(row => String(row[groupField])))] : [analysisName];
  const series = groups.map(group => {
    const subset = groupField ? rows.filter(row => String(row[groupField]) === group) : rows;
    const byCategory = new Map(subset.map(row => [String(row[categoryField]), numericValue(row[yField], yField, visual)]));
    if (categories.some(category => !byCategory.has(category))) throw new Error(`Native chart ${visual.visual_id} has incomplete categories in group ${group}`);
    return { name: group, labels: categories, values: categories.map(category => byCategory.get(category)) };
  });
  const chartType = ["line", "line_chart", "simple_time_series", "time_series"].includes(type)
    ? "line"
    : type === "area_chart" ? "area"
      : (["grouped_bar", "grouped_bar_chart", "grouped_composition"].includes(type) ? "grouped_bar" : "bar");
  const pptxType = chartType === "line" ? pptx.ChartType.line : chartType === "area" ? pptx.ChartType.area : pptx.ChartType.bar;
  return { chartType, pptxType, rows, series, categoryCount: categories.length, workbookFields: [categoryField, yField, ...(groupField ? [groupField] : [])], privacyValidation };
}

function funnelSemanticContract(visual) {
  const payload = visual.presentation_payload || {};
  const semantics = payload.funnel_semantics;
  if (!semantics || typeof semantics !== "object") {
    throw new Error(`Funnel plot ${visual.visual_id} requires an explicit funnel_semantics contract`);
  }
  if (semantics.effect_measure !== "odds_ratio" || semantics.transformed_scale !== "log(OR)") {
    throw new Error(`Funnel plot ${visual.visual_id} supports only the declared odds-ratio/log(OR) semantic contract`);
  }
  if (Number(semantics.null_reference) !== 0 || semantics.reference_kind !== "null_effect_not_pooled") {
    throw new Error(`Funnel plot ${visual.visual_id} must declare log(OR)=0 as a null-effect reference, never a pooled-effect reference`);
  }
  if (semantics.guide_formula !== "log(OR) = 0 ± 1.96 × SE") {
    throw new Error(`Funnel plot ${visual.visual_id} has an unsupported guide formula`);
  }
  const transform = payload.source_transform_contract || {};
  const fields = new Set(transform.fields_used || []);
  if (!["study_label", "odds_ratio", "ci_low", "ci_high"].every(field => fields.has(field))) {
    throw new Error(`Funnel plot ${visual.visual_id} lacks a source-bound OR/CI transform contract`);
  }
  return semantics;
}

function renderFunnelSemanticOverlay(slide, visual, bounds, plotLayout, xLimit, yMaximum, p) {
  const semantics = funnelSemanticContract(visual);
  if (!(xLimit > 0) || !(yMaximum > 0)) throw new Error(`Funnel plot ${visual.visual_id} has an invalid plotting domain`);
  const plot = {
    x: bounds.x + bounds.w * plotLayout.x,
    y: bounds.y + bounds.h * plotLayout.y,
    w: bounds.w * plotLayout.w,
    h: bounds.h * plotLayout.h,
  };
  const xAt = value => plot.x + ((value + xLimit) / (2 * xLimit)) * plot.w;
  const yAt = value => plot.y + (value / yMaximum) * plot.h;
  const nullX = xAt(0);
  const guideSe = Math.min(yMaximum, xLimit / 1.96);
  const guideY = yAt(guideSe);
  const guideOffset = 1.96 * guideSe;
  const annotationGutter = 0.14;
  const annotationLeft = bounds.x + 0.06;
  const annotationRight = bounds.x + bounds.w - 0.06;
  const annotationWidth = annotationRight - annotationLeft;
  const guideLabelWidth = Math.max(2.80, annotationWidth * 0.56);
  const guideLabelBounds = {
    x: annotationLeft,
    y: bounds.y + 0.04,
    w: guideLabelWidth,
    h: 0.24,
  };
  const nullLabelBounds = {
    x: guideLabelBounds.x + guideLabelBounds.w + annotationGutter,
    y: guideLabelBounds.y,
    w: annotationRight - (guideLabelBounds.x + guideLabelBounds.w + annotationGutter),
    h: 0.24,
  };
  if (guideLabelBounds.w < 2.80 || nullLabelBounds.w < 1.90) {
    throw new Error(`Funnel plot ${visual.visual_id} has insufficient horizontal annotation width`);
  }
  if (guideLabelBounds.x + guideLabelBounds.w + annotationGutter > nullLabelBounds.x ||
      nullLabelBounds.x + nullLabelBounds.w > annotationRight + 1e-6) {
    throw new Error(`Funnel plot ${visual.visual_id} has overlapping semantic annotation columns`);
  }
  if (Math.max(guideLabelBounds.y + guideLabelBounds.h, nullLabelBounds.y + nullLabelBounds.h) > plot.y - 0.06) {
    throw new Error(`Funnel plot ${visual.visual_id} has insufficient annotation-to-plot gutter`);
  }
  line(slide, nullX, plot.y, 0, plot.h, p.secondary, 1.4, {
    dash: "dash",
    name: `${visual.visual_id}:null-reference-log-or-zero`,
  });
  line(slide, nullX, plot.y, xAt(-guideOffset) - nullX, guideY - plot.y, SHARED.muted, 1.0, {
    dash: "dash",
    name: `${visual.visual_id}:null-centered-guide-left`,
  });
  line(slide, nullX, plot.y, xAt(guideOffset) - nullX, guideY - plot.y, SHARED.muted, 1.0, {
    dash: "dash",
    name: `${visual.visual_id}:null-centered-guide-right`,
  });
  addText(slide, "零效应 log(OR)=0", nullLabelBounds.x, nullLabelBounds.y, nullLabelBounds.w, nullLabelBounds.h, TYPO.chart, {
    bold: true,
    color: p.secondary,
    name: `${visual.visual_id}:null-reference-label`,
    valign: "top",
    align: "right",
    wrap: false,
  });
  addText(slide, "参考界限 ±1.96×SE｜非合并效应", guideLabelBounds.x, guideLabelBounds.y, guideLabelBounds.w, guideLabelBounds.h, TYPO.chart, {
    color: SHARED.muted,
    name: `${visual.visual_id}:null-centered-guide-label`,
    valign: "top",
    align: "left",
    wrap: false,
  });
  return {
    referenceKind: semantics.reference_kind,
    nullReference: Number(semantics.null_reference),
    guideFormula: semantics.guide_formula,
    interpretation: semantics.guide_interpretation,
    geometryBasis: "manual_native_chart_plot_layout",
    plotLayout,
    guideMaximumStandardError: guideSe,
    pooledEffectUsed: false,
  };
}

function renderNativeChart(slide, visual, spec) {
  const p = profile(spec);
  const geometry = presentationGeometry(spec, visual);
  const bounds = geometry.hero;
  assertSafeContent(bounds, `${visual.visual_id} native chart`);
  const data = nativeChartData(visual);
  const isHistogram = data.chartType === "histogram";
  const isFunnel = data.chartType === "funnel_plot";
  const isGroupedComposition = String(visual.visual_type) === "grouped_composition";
  const isBar = ["bar", "grouped_bar", "histogram"].includes(data.chartType);
  // Native PowerPoint data labels are positioned independently by each
  // renderer.  Dense grouped bars therefore collide even when the chart frame
  // itself is valid.  Keep every aggregate row in the embedded workbook, but
  // suppress direct labels once the grouped-composition label budget is
  // exceeded; the quantitative axis, group legend, and source-bound slide
  // annotation remain the audience-facing reading path.
  const nativeDataLabelCount = data.categoryCount * Math.max(1, data.series.length);
  const groupedCompositionDirectLabelSafe = !isGroupedComposition || nativeDataLabelCount <= 8;
  const showDirectValues = isBar && data.categoryCount <= 10 && groupedCompositionDirectLabelSafe;
  const categoryAxisTitle = isHistogram
    ? "稳定权重"
    : isFunnel ? "log(OR)" : audienceAxisLabel(firstDefined(visual.x_axis_label, visual.time_field, visual.label_field, visual.x_field, ""), "category", visual.visual_type);
  const valueAxisTitle = isHistogram
    ? "频数"
    : isFunnel ? "标准误" : audienceAxisLabel(firstDefined(visual.y_axis_label, visual.unit, visual.y_field, ""), "value", visual.visual_type);
  // Localize only audience-facing category and series labels. The values,
  // workbook-field allowlist, source bindings, and canonical aggregate rows
  // remain unchanged and are still validated before rendering.
  const displaySeries = isGroupedComposition && isChineseDeck()
    ? data.series.map(series => ({
      ...series,
      name: zhLabel(series.name),
      labels: Array.isArray(series.labels) ? series.labels.map(zhLabel) : series.labels,
    }))
    : data.series;
  const funnelPlotLayout = { x: 0.13, y: 0.18, w: 0.78, h: 0.62 };
  const chartOptions = {
    x: bounds.x, y: bounds.y, w: bounds.w, h: bounds.h,
    name: `${visual.visual_id}:native-chart`,
    showTitle: false,
    showLegend: isFunnel ? false : (data.series.length > 1 || data.chartType === "line" || data.chartType === "scatter"),
    legendPos: "b",
    legendFontFace: BODY_FONT,
    legendFontSize: TYPO.chart,
    chartColors: (isHistogram || isFunnel) ? [p.accent] : [p.accent, p.secondary, p.primary, p.negative || SHARED.muted],
    varyColors: false,
    showValue: showDirectValues,
    dataLabelFormatCode: showDirectValues ? "0.0" : undefined,
    showCatName: false,
    showSerName: false,
    dataLabelPosition: isHistogram ? "outEnd" : "outEnd",
    dataLabelColor: SHARED.ink,
    dataLabelFontFace: BODY_FONT,
    dataLabelFontSize: TYPO.chart,
    catAxisLabelColor: SHARED.muted,
    catAxisLabelFontFace: BODY_FONT,
    catAxisLabelFontSize: TYPO.chart,
    valAxisLabelColor: SHARED.muted,
    valAxisLabelFontFace: BODY_FONT,
    valAxisLabelFontSize: TYPO.chart,
    catAxisLineColor: SHARED.line,
    valAxisLineColor: SHARED.line,
    valGridLine: { color: "D9E3EA", width: 1 },
    showCatAxisTitle: Boolean(categoryAxisTitle),
    catAxisTitle: categoryAxisTitle,
    catAxisTitleFontFace: BODY_FONT,
    catAxisTitleFontSize: TYPO.chart,
    showValAxisTitle: Boolean(valueAxisTitle),
    valAxisTitle: valueAxisTitle,
    valAxisTitleFontFace: BODY_FONT,
    valAxisTitleFontSize: TYPO.chart,
    showMarker: data.chartType === "line" || data.chartType === "scatter" || data.chartType === "funnel_plot",
    lineSize: 2,
    showLine: !["scatter", "funnel_plot"].includes(data.chartType) || visual.show_line === true,
    gapWidthPct: isHistogram ? 18 : 55,
    barDir: (isHistogram || data.chartType === "grouped_bar") ? "col" : "bar",
  };
  if (isHistogram) chartOptions.catAxisLabelRotate = -45;
  if (isFunnel) {
    const xValues = data.series[0].values.map(Number);
    const yValues = data.series.slice(1).flatMap(series => series.values.filter(value => value !== null).map(Number));
    const limit = Math.ceil(Math.max(0.5, ...xValues.map(value => Math.abs(value))) * 2) / 2;
    const yMaximum = Math.ceil(Math.max(...yValues) * 20) / 20;
    chartOptions.catAxisMinVal = -limit;
    chartOptions.catAxisMaxVal = limit;
    chartOptions.catAxisMajorUnit = limit <= 1 ? 0.25 : 0.5;
    chartOptions.catLabelFormatCode = "0.00";
    chartOptions.valAxisMinVal = 0;
    chartOptions.valAxisMaxVal = yMaximum;
    chartOptions.valAxisMajorUnit = 0.05;
    chartOptions.valAxisLabelFormatCode = "0.00";
    chartOptions.valAxisOrientation = "maxMin";
    // Pin the native plot area so the mathematically derived editable-shape
    // overlay uses the same coordinate transform in PowerPoint and LibreOffice.
    chartOptions.layout = funnelPlotLayout;
  }
  slide.addChart(data.pptxType, displaySeries, chartOptions);
  let funnelSemanticOverlay = null;
  if (isFunnel) {
    const xValues = data.series[0].values.map(Number);
    const yValues = data.series.slice(1).flatMap(series => series.values.filter(value => value !== null).map(Number));
    const limit = Math.ceil(Math.max(0.5, ...xValues.map(value => Math.abs(value))) * 2) / 2;
    const yMaximum = Math.ceil(Math.max(...yValues) * 20) / 20;
    funnelSemanticOverlay = renderFunnelSemanticOverlay(slide, visual, bounds, funnelPlotLayout, limit, yMaximum, p);
  }
  return {
    renderedRowKeys: [...visual.expected_row_keys],
    objectType: "native_powerpoint_chart",
    editability: "full",
    nativeChart: true,
    chartType: data.chartType,
    chartSeriesCount: ["scatter", "funnel_plot"].includes(data.chartType) ? Math.max(0, data.series.length - 1) : data.series.length,
    chartCategoryCount: data.categoryCount,
    chartWorkbookSource: "visual_spec_aggregate_rows",
    chartWorkbookFields: data.workbookFields,
    directValueLabels: showDirectValues,
    dataLabelPolicy: isGroupedComposition && !showDirectValues ? "axis_legend_plus_annotation" : "native_direct_values_when_safe",
    nativeChartPrivacy: visual.native_chart_privacy || {},
    nativeChartPrivacyValidation: data.privacyValidation,
    funnelSemanticOverlay,
    assetPath: null,
    sourceFigureTreatment: null,
    heroBounds: bounds,
    heroShare: geometry.heroShare,
    contentScale: 1,
    fallbackUsed: false,
    warnings: [...geometry.warnings],
  };
}

function resolveSourceFigureTreatment(visual, spec) {
  const value = String(firstDefined(visual.source_figure_treatment, visual.payload?.source_figure_treatment, presentationFields(spec).sourceFigureTreatment));
  if (!SOURCE_FIGURE_TREATMENTS.has(value)) throw new Error(`Unsupported source-figure treatment ${value} for ${visual.visual_id}`);
  return value;
}

function renderSourceFigure(slide, visual, spec, mode) {
  const p = profile(spec);
  const geometry = presentationGeometry(spec, visual);
  const treatment = resolveSourceFigureTreatment(visual, spec);
  const assetPath = resolveAssetPath(firstDefined(visual.presentation_payload?.svg_path, visual.presentation_payload?.image_path, visual.presentation_payload?.asset_path, visual.payload?.svg_path, visual.payload?.image_path, visual.payload?.asset_path, visual.asset_path));
  const extension = path.extname(assetPath).toLowerCase();
  if (mode === "svg" && extension !== ".svg") throw new Error(`${visual.visual_id} svg render mode received ${extension}`);
  if ((visual.expected_row_keys || []).length !== 1) throw new Error(`Source figure ${visual.visual_id} must bind exactly one renderer-evidence row`);
  let figureBounds = { ...geometry.hero };
  let localCallout = null;
  if (treatment === "figure_with_callout" && visual.payload?.callout && !(visual.annotation_specs || []).length) {
    if (geometry.annotation) {
      localCallout = { ...geometry.annotation };
    } else {
      const gutter = 0.30;
      const calloutWidth = Math.max(2.15, figureBounds.w * 0.25);
      localCallout = { x: figureBounds.x + figureBounds.w - calloutWidth, y: figureBounds.y + 0.32, w: calloutWidth, h: Math.min(3.25, figureBounds.h - 0.64) };
      figureBounds = { ...figureBounds, w: figureBounds.w - calloutWidth - gutter };
    }
  }
  assertSafeContent(figureBounds, `${visual.visual_id} source figure`);
  if (treatment === "cover" || treatment === "full_bleed") imageCover(slide, assetPath, figureBounds.x, figureBounds.y, figureBounds.w, figureBounds.h, rowName(visual, 0));
  else imageContain(slide, assetPath, figureBounds.x, figureBounds.y, figureBounds.w, figureBounds.h, rowName(visual, 0));
  if (localCallout) {
    assertSafeContent(localCallout, `${visual.visual_id} source-figure callout`);
    box(slide, localCallout.x, localCallout.y, localCallout.w, localCallout.h, p.soft, p.accent, 0.03, `${visual.visual_id}:source-callout`);
    const calloutTitle = String(firstDefined(visual.payload?.callout_title, "图示说明"));
    addText(slide, calloutTitle, localCallout.x + 0.20, localCallout.y + 0.18, localCallout.w - 0.40, 0.36, TYPO.sub, { bold: true, color: p.primary });
    addText(slide, visual.payload.callout, localCallout.x + 0.20, localCallout.y + 0.70, localCallout.w - 0.40, localCallout.h - 0.92, TYPO.body, { color: SHARED.ink, valign: "top" });
  }
  return {
    renderedRowKeys: [...visual.expected_row_keys],
    objectType: mode === "svg" ? "source_svg" : "registered_source_figure",
    editability: mode === "svg" ? "vector_image" : "mixed",
    nativeChart: false,
    chartType: null,
    chartSeriesCount: null,
    chartCategoryCount: null,
    chartWorkbookSource: null,
    assetPath,
    sourceFigureTreatment: treatment,
    heroBounds: geometry.hero,
    heroShare: geometry.heroShare,
    contentScale: 1,
    fallbackUsed: false,
    warnings: [...geometry.warnings],
  };
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
  const coverTitle = wrapTitleText(localizeAudienceText(spec.title), 30);
  const coverLines = coverTitle.split("\n");
  if (coverLines.length > 2) throw new Error(`Cover title exceeds the two-line contract on ${spec.slide_id}`);
  const longestLine = Math.max(...coverLines.map(titleDisplayWidth));
  const coverFontSize = longestLine <= 30 ? TYPO.deck : longestLine <= 32 ? 44 : longestLine <= 35 ? 40 : 38;
  addText(slide, coverTitle, 1.28, 1.30, 11.25, 1.88, coverFontSize, { bold: true, color: SHARED.white, name: `${spec.slide_id}:deck-title`, wrap: false, valign: "mid", margin: [0.05, 0, 0.05, 0] });
  const subtitle = { target_trial: "临床方法 / 目标试验模拟", single_cell: "生物信息学 / 单细胞分析", meta_analysis: "证据综合 / Meta分析", protocol: "临床研究方案 / 前瞻性TTE" }[projectKind];
  addText(slide, subtitle, 1.30, 3.35, 8.7, 0.42, TYPO.sub, { color: "CBE4E6" });
  addText(slide, "完全合成回归测试材料｜不构成真实临床证据", 1.30, 5.45, 10.6, 0.34, TYPO.detail, { bold: true, color: "F2BE6A" });
}

function renderHeroQuestion(slide, visual, spec) {
  const p = profile(spec), data = visual.payload;
  addText(slide, data.question, 0.86, 1.52, 11.3, 1.05, 30, { bold: true, color: p.primary, name: rowName(visual, 0) });
  line(slide, 0.90, 2.86, 11.4, 0, p.accent, 2.2);
  const entries = Object.entries(data).filter(([key]) => key !== "question");
  entries.forEach(([key, value], index) => {
    const columnWidth = 11.4 / entries.length;
    const x = 0.92 + index * columnWidth;
    const bodyWidth = columnWidth - 0.14;
    addText(slide, ({ gap: "知识缺口", objective: "目标", population: "人群", exposure: "暴露", outcome: "结局", design: "设计", sites: "站点", boundary: "边界" }[key] || key), x, 3.16, columnWidth - 0.08, 0.30, TYPO.detail, { bold: true, color: p.accent });
    const lineBudget = Math.max(12, Math.floor(bodyWidth * 7.8));
    const tokenSafeWrap = needsTokenSafeBodyWrap(value, lineBudget);
    const bodyText = tokenSafeWrap ? wrapAudienceBodyText(value, lineBudget) : value;
    addText(slide, bodyText, x, 3.62, bodyWidth, 1.34, TYPO.body, { color: SHARED.ink, valign: "top", wrap: tokenSafeWrap ? false : undefined });
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
    const lineBudget = Math.max(12, Math.floor((w - 0.32) * 7.7));
    const tokenSafeWrap = needsTokenSafeBodyWrap(item[1], lineBudget);
    const bodyText = tokenSafeWrap ? wrapAudienceBodyText(item[1], lineBudget) : item[1];
    addText(slide, bodyText, x + 0.16, y + 0.58, w - 0.32, h - 0.70, TYPO.detail, { color: SHARED.muted, valign: "top", wrap: tokenSafeWrap ? false : undefined });
    if (item[2]) {
      const severity = String(item[2]).toLowerCase();
      const badgeColor = severity.includes("high") || severity.includes("高") ? SHARED.warning : profile(spec).accent;
      addText(slide, item[2], x + w - 0.98, y + 0.12, 0.80, 0.24, TYPO.chart, { bold: true, color: badgeColor, align: "right" });
    }
  });
}

function renderBranch(slide, visual, spec) {
  const nodes = visual.payload.nodes || [];
  if (!nodes.length) throw new Error(`Process branch ${visual.visual_id} requires at least one node`);
  const nodeIds = nodes.map((node, index) => String(typeof node === "object" ? firstDefined(node.id, index) : index));
  const nodeIndex = new Map(nodeIds.map((id, index) => [id, index]));
  const rawEdges = visual.payload.edges || visual.payload.branches || [];
  if (!rawEdges.length && nodes.length > 1 && visual.sequential !== true && visual.payload.sequential !== true) {
    throw new Error(`Process branch ${visual.visual_id} requires explicit payload edges; implicit sequencing is prohibited`);
  }
  const edges = (rawEdges.length ? rawEdges : nodes.slice(1).map((_, index) => [index, index + 1])).map((edge, index) => {
    const sourceValue = Array.isArray(edge) ? edge[0] : firstDefined(edge.source, edge.from);
    const targetValue = Array.isArray(edge) ? edge[1] : firstDefined(edge.target, edge.to);
    const source = nodeIndex.has(String(sourceValue)) ? nodeIndex.get(String(sourceValue)) : Number(sourceValue);
    const target = nodeIndex.has(String(targetValue)) ? nodeIndex.get(String(targetValue)) : Number(targetValue);
    if (!Number.isInteger(source) || !Number.isInteger(target) || !nodes[source] || !nodes[target]) {
      throw new Error(`Process branch ${visual.visual_id} edge ${index} references an unknown node`);
    }
    return { source, target, label: Array.isArray(edge) ? "" : String(edge.label || "") };
  });

  // Derive graph depth from the declared edges.  Node array order is retained
  // only as a deterministic tie-breaker; it never creates a semantic edge.
  const outgoing = nodes.map(() => []), indegree = nodes.map(() => 0), depth = nodes.map(() => 0);
  edges.forEach(edge => { outgoing[edge.source].push(edge.target); indegree[edge.target] += 1; });
  const queue = indegree.map((value, index) => ({ value, index })).filter(item => item.value === 0).map(item => item.index);
  const ordered = [];
  while (queue.length) {
    const source = queue.shift();
    ordered.push(source);
    for (const target of outgoing[source]) {
      depth[target] = Math.max(depth[target], depth[source] + 1);
      indegree[target] -= 1;
      if (indegree[target] === 0) queue.push(target);
    }
  }
  if (ordered.length !== nodes.length) throw new Error(`Process branch ${visual.visual_id} must be acyclic`);

  const levelCount = Math.max(...depth) + 1;
  const levelGap = levelCount <= 4 ? 0.72 : 0.48;
  const nodeWidth = Math.max(1.42, Math.min(2.28, (11.70 - levelGap * Math.max(0, levelCount - 1)) / levelCount));
  const nodeHeight = 0.88;
  const xStep = levelCount === 1 ? 0 : (11.70 - nodeWidth) / (levelCount - 1);
  const positions = Array(nodes.length);
  for (let level = 0; level < levelCount; level++) {
    const members = nodes.map((_, index) => index).filter(index => depth[index] === level);
    const top = members.length <= 1 ? 3.02 : 1.72;
    const bottom = members.length <= 1 ? top : 4.56;
    members.forEach((node, lane) => {
      const y = members.length === 1 ? top : top + lane * ((bottom - top) / Math.max(1, members.length - 1));
      positions[node] = { x: 0.82 + level * xStep, y };
    });
  }

  // Connectors are emitted before nodes.  Adjacent levels use an orthogonal
  // gutter so a fork/convergence remains visible in both PowerPoint and LO.
  edges.forEach((edge, index) => {
    const a = positions[edge.source], b = positions[edge.target];
    const startX = a.x + nodeWidth, startY = a.y + nodeHeight / 2;
    const endX = b.x, endY = b.y + nodeHeight / 2;
    const midX = startX + (endX - startX) / 2;
    const edgeName = `${visual.visual_id}:edge:${edge.source}-${edge.target}`;
    if (Math.abs(startY - endY) < 0.02) {
      line(slide, startX, startY, endX - startX, 0, profile(spec).accent, 1.8, { endArrowType: "triangle", name: edgeName });
    } else {
      line(slide, startX, startY, midX - startX, 0, profile(spec).accent, 1.8, { name: `${edgeName}:a` });
      line(slide, midX, startY, 0, endY - startY, profile(spec).accent, 1.8, { name: `${edgeName}:b` });
      line(slide, midX, endY, endX - midX, 0, profile(spec).accent, 1.8, { endArrowType: "triangle", name: `${edgeName}:c` });
    }
    if (edge.label) addText(slide, zhLabel(edge.label), midX - 0.55, Math.min(startY, endY) - 0.30, 1.10, 0.24, TYPO.chart, { align: "center", color: SHARED.muted });
  });
  nodes.forEach((node, index) => {
    const label = typeof node === "object" ? firstDefined(node.label, node.id) : node;
    const pos = positions[index];
    box(slide, pos.x, pos.y, nodeWidth, nodeHeight, depth[index] % 2 ? SHARED.white : profile(spec).soft, profile(spec).accent, 0.03, rowName(visual, index));
    addText(slide, zhLabel(label), pos.x + 0.10, pos.y + 0.10, nodeWidth - 0.20, nodeHeight - 0.20, TYPO.detail, { bold: true, align: "center" });
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
  const confidenceInterval = normalizeScientificTypography(`95% CI ${(Number(differenceRow.ci_low) * 100).toFixed(1)} 至 ${(Number(differenceRow.ci_high) * 100).toFixed(1)} pp`)
    .replace(/^95% CI\s+/, "95% CI\n");
  addText(slide, confidenceInterval, 10.02, 3.48, 2.16, 0.76, TYPO.detail, { color: SHARED.muted, wrap: false, valign: "top" });
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
    addText(slide, `95% CI ${row.ci_low}–${row.ci_high}`, x + 0.28, 3.62, 3.70, 0.36, TYPO.body, { color: SHARED.muted });
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
  const x0 = 4.88, x1 = 9.95, resultX = 10.18, resultW = 2.27;
  const plotTop = 1.57, plotBottom = visual.payload.pooled ? 5.78 : 5.95;
  const compactNumber = value => {
    const numeric = Number(value);
    // Avoid a misleading signed zero after deterministic label rounding.
    // The source value and provenance remain unchanged in VisualSpec/audit.
    const rounded = Math.abs(numeric) < 0.005 ? 0 : numeric;
    return rounded.toFixed(2).replace(/^-/, "−");
  };
  const compactCI = (estimate, low, high) => `${compactNumber(estimate)} [${compactNumber(low)}, ${compactNumber(high)}]`;
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
    addText(slide, labels[index], 0.82, y - 0.02, 3.72, Math.max(0.22, rowH - 0.02), TYPO.chart, { color: SHARED.ink });
    line(slide, scale(lows[index]), y + rowH * 0.42, Math.max(0.02, scale(highs[index]) - scale(lows[index])), 0, p.accent, 1.6);
    ellipse(slide, scale(estimates[index]) - 0.055, y + rowH * 0.42 - 0.055, 0.11, 0.11, p.accent, undefined, rowName(visual, index));
    const ciHeight = isSubgroup ? Math.min(0.28, Math.max(0.22, rowH * 0.44)) : Math.max(0.22, rowH - 0.02);
    addText(slide, compactCI(estimates[index], lows[index], highs[index]), resultX, y - 0.02, resultW, ciHeight, TYPO.chart, { align: "right", color: SHARED.muted, fontFace: "Arial", wrap: false });
    if (isSubgroup) addText(slide, `交互 P=${Number(row.interaction_p).toFixed(2)}`, resultX, y + Math.max(0.20, rowH * 0.46), resultW, Math.min(0.25, rowH * 0.34), TYPO.chart, { align: "right", color: SHARED.muted, fontFace: "Arial", wrap: false });
  });
  if (visual.payload.pooled) {
    const pooled = visual.payload.pooled, e = Number(pooled.estimate), low = Number(pooled.ci_low), high = Number(pooled.ci_high), y = 6.01;
    addText(slide, "随机效应合并", 0.82, y - 0.10, 3.95, 0.30, TYPO.chart, { bold: true, color: p.primary });
    const cx = scale(e), left = scale(low), right = scale(high);
    const diamondWidth = Math.max(0.20, right - left);
    slide.addShape(S.diamond, { x: cx - diamondWidth / 2, y: y - 0.10, w: diamondWidth, h: 0.28, fill: { color: p.secondary }, line: { color: p.secondary }, name: `${visual.visual_id}:pooled-diamond` });
    addText(slide, compactCI(e, low, high), resultX, y - 0.12, resultW, 0.32, TYPO.chart, { bold: true, align: "right", color: p.primary, fontFace: "Arial", wrap: false });
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
  if (!effectField) throw new Error(`faceted_volcano requires an effect x_field: ${visual.visual_id}`);
  if (facets.length > 6) throw new Error(`faceted_volcano supports at most six readable facets per slide; received ${facets.length}`);
  const quantitativeRows = rows.map(item => {
    const effect = Number(item[effectField]);
    const fdr = Number(item.FDR);
    if (!Number.isFinite(effect) || !Number.isFinite(fdr) || fdr <= 0 || fdr > 1) {
      throw new Error(`faceted_volcano requires finite effect and 0 < FDR <= 1 for ${visual.visual_id}`);
    }
    return { item, effect, fdr, score: -Math.log10(fdr) };
  });
  const maxAbsEffect = Math.max(...quantitativeRows.map(row => Math.abs(row.effect)));
  const xLimit = Math.max(1, Math.ceil(maxAbsEffect * 2) / 2);
  const maxScore = Math.max(...quantitativeRows.map(row => row.score));
  const yLimit = Math.max(1, Math.ceil(maxScore / 5) * 5);
  facets.forEach((facet, index) => {
    const col = index % 3, rowIndex = Math.floor(index / 3);
    const x = 0.78 + col * 4.18, y = 1.45 + rowIndex * 2.63, w = 3.82, h = 2.20;
    box(slide, x, y, w, h, SHARED.white, "D6E1E8", 0.02);
    addText(slide, zhLabel(facet), x + 0.14, y + 0.10, w - 0.28, 0.28, TYPO.detail, { bold: true, color: p.primary });
    // Reserve the left gutter exclusively for y tick labels.  Keeping the
    // quantitative title in the top plot band prevents it from intersecting
    // the maximum y tick in PowerPoint's actual text bounds.
    addText(slide, "-log₁₀(FDR)", x + 0.62, y + 0.39, 1.30, 0.20, TYPO.chart, { color: SHARED.muted });
    addText(slide, "log₂FC", x + w - 0.88, y + 0.39, 0.70, 0.20, TYPO.chart, { color: SHARED.muted, align: "right" });
    const subset = quantitativeRows.filter(row => row.item.cell_type === facet);
    const plotLeft = x + 0.50, plotRight = x + w - 0.18;
    const plotTop = y + 0.66, plotBottom = y + h - 0.40;
    const scaleX = linearScale(-xLimit, xLimit, plotLeft, plotRight);
    const scaleY = linearScale(0, yLimit, plotBottom, plotTop);
    line(slide, plotLeft, plotBottom, plotRight - plotLeft, 0, SHARED.muted, 0.8);
    line(slide, plotLeft, plotTop, 0, plotBottom - plotTop, SHARED.muted, 0.8);
    line(slide, scaleX(0), plotTop, 0, plotBottom - plotTop, SHARED.line, 0.8, { dash: "dash" });
    const xTicks = [-xLimit, 0, xLimit];
    xTicks.forEach(tick => {
      const tx = scaleX(tick);
      line(slide, tx, plotBottom - 0.035, 0, 0.07, SHARED.muted, 0.6);
      addText(slide, tick === 0 ? "0" : tick.toFixed(xLimit < 2 ? 1 : 0), tx - 0.26, plotBottom + 0.04, 0.52, 0.20, TYPO.chart, { align: "center", color: SHARED.muted, fontFace: "Arial" });
    });
    const yTicks = [0, yLimit / 2, yLimit];
    yTicks.forEach(tick => {
      const ty = scaleY(tick);
      line(slide, plotLeft - 0.035, ty, 0.07, 0, SHARED.muted, 0.6);
      // At the shared lower-left origin the x-minimum tick sits immediately
      // below the axis.  Move only the y=0 label above the origin, preserving
      // both quantitative labels and the minimum 15 pt chart font.
      const yTickLabelY = tick === 0 ? ty - 0.30 : ty - 0.10;
      addText(slide, Number.isInteger(tick) ? String(tick) : tick.toFixed(1), x + 0.03, yTickLabelY, 0.39, 0.20, TYPO.chart, { align: "right", color: SHARED.muted, fontFace: "Arial" });
    });
    const labelRows = new Set([...subset].sort((a, b) => a.fdr - b.fdr).slice(0, 2).map(row => row.item));
    const placedLabels = [];
    subset.forEach(({ item, effect, score }) => {
      const cx = scaleX(effect), cy = scaleY(score);
      ellipse(slide, cx - 0.055, cy - 0.055, 0.11, 0.11, effect >= 0 ? p.secondary : p.negative || "3B82C4", undefined, rowName(visual, rows.indexOf(item)));
      if (labelRows.has(item)) {
        const labelW = 1.28, labelH = 0.28;
        let labelY = cy - labelH - 0.08;
        if (labelY < plotTop) labelY = cy + 0.08;
        labelY = clamp(labelY, plotTop, plotBottom - labelH);
        const labelX = Math.max(plotLeft + 0.04, Math.min(plotRight - labelW, cx - labelW / 2));
        for (const placed of placedLabels) {
          if (Math.abs(labelX - placed.x) < labelW && Math.abs(labelY - placed.y) < labelH + 0.08) {
            labelY = Math.min(plotBottom - labelH, placed.y + labelH + 0.08);
          }
        }
        placedLabels.push({ x: labelX, y: labelY });
        addText(slide, item.gene, labelX, labelY, labelW, labelH, TYPO.chart, { bold: score > 5, align: "center", color: SHARED.ink, fontFace: "Arial" });
      }
    });
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

function isPredictedInteraction(edge, visual) {
  for (const key of ["predicted", "is_predicted", "prediction_only"]) {
    if (edge[key] === true || String(edge[key]).toLowerCase() === "true") return true;
    if (edge[key] === false || String(edge[key]).toLowerCase() === "false") return false;
  }
  return Boolean(visual.payload?.prediction_badge) || /predict/i.test(String(visual.scientific_role || ""));
}

function clippedEllipseDirectedEdge(source, target, radiusX = 0.55, radiusY = 0.34, gap = 0.06) {
  const dx = target.x - source.x, dy = target.y - source.y;
  const length = Math.hypot(dx, dy);
  if (!(length > 0)) throw new Error("Directed network edge cannot connect a node to itself");
  const boundaryDenominator = Math.sqrt((dx * dx) / (radiusX * radiusX) + (dy * dy) / (radiusY * radiusY));
  if (!(boundaryDenominator > 0)) throw new Error("Directed network edge has invalid ellipse-boundary geometry");
  const boundaryFraction = 1 / boundaryDenominator;
  const ux = dx / length, uy = dy / length;
  const start = {
    x: source.x + dx * boundaryFraction + ux * gap,
    y: source.y + dy * boundaryFraction + uy * gap,
  };
  const end = {
    x: target.x - dx * boundaryFraction - ux * gap,
    y: target.y - dy * boundaryFraction - uy * gap,
  };
  if (Math.hypot(end.x - start.x, end.y - start.y) < 0.18) {
    throw new Error("Directed network edge is too short after node-boundary clipping");
  }
  return { start, end };
}

function renderNetwork(slide, visual, spec) {
  const rows = visual.payload.rows, p = profile(spec);
  if (visual.sequential === true) throw new Error(`Independent communication network ${visual.visual_id} cannot be rendered as a sequence`);
  const names = [...new Set(rows.flatMap(row => [row.sender, row.receiver]))];
  const center = { x: 6.60, y: 3.78 }, radiusX = 4.55, radiusY = 2.20;
  const positions = Object.fromEntries(names.map((name, index) => {
    const angle = -Math.PI / 2 + index / names.length * Math.PI * 2;
    return [name, { x: center.x + Math.cos(angle) * radiusX, y: center.y + Math.sin(angle) * radiusY }];
  }));
  let predictedEdgeCount = 0;
  rows.forEach((edge, index) => {
    const a = positions[edge.sender], b = positions[edge.receiver], probability = Number(edge.communication_probability);
    const predicted = isPredictedInteraction(edge, visual);
    if (predicted) predictedEdgeCount += 1;
    const edgeColor = predicted ? p.prediction : p.accent;
    // End the connector immediately before the receiver ellipse.  Center-to-
    // center arrows place the arrowhead underneath the receiver node in both
    // PowerPoint and LibreOffice, making independent interactions look
    // undirected.  Boundary clipping keeps the visible arrowhead outside the
    // node while retaining one source-bound edge object per canonical row.
    const clipped = clippedEllipseDirectedEdge(a, b);
    line(slide, clipped.start.x, clipped.start.y, clipped.end.x - clipped.start.x, clipped.end.y - clipped.start.y, edgeColor, 0.9 + probability * 2.2, { dash: predicted ? "dash" : undefined, endArrowType: "triangle", name: rowName(visual, index) });
    const dx = b.x - a.x, dy = b.y - a.y, edgeLength = Math.max(0.01, Math.hypot(dx, dy));
    const offset = index % 2 === 0 ? 0.13 : -0.13;
    const mx = (a.x + b.x) / 2 - dy / edgeLength * offset;
    const my = (a.y + b.y) / 2 + dx / edgeLength * offset;
    box(slide, mx - 0.78, my - 0.27, 1.56, 0.54, SHARED.white, edgeColor, 0.02);
    addText(slide, `${edge.ligand_receptor}\nP=${probability.toFixed(2)}`, mx - 0.72, my - 0.22, 1.44, 0.44, TYPO.chart, { align: "center", color: SHARED.ink, margin: 0.01 });
  });
  names.forEach((name, index) => {
    const pos = positions[name], color = CELL_COLORS[name] || (index % 2 ? p.secondary : p.accent);
    ellipse(slide, pos.x - 0.55, pos.y - 0.34, 1.10, 0.68, color, SHARED.white, `${visual.visual_id}:node:${index}`);
    addText(slide, zhLabel(name), pos.x - 0.48, pos.y - 0.21, 0.96, 0.40, TYPO.chart, { bold: true, color: SHARED.white, align: "center" });
  });
  box(slide, 9.72, 6.02, 2.65, 0.52, "FFF4E5", SHARED.warning, 0.02);
  if (predictedEdgeCount) line(slide, 9.02, 6.28, 0.54, 0, p.prediction, 1.8, { dash: "dash", endArrowType: "triangle", name: `${visual.visual_id}:prediction-legend` });
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
  const robLegend = observedRiskOfBiasLegend(rows, visual.visual_id);
  addText(slide, `${rows.length}项研究`, 0.82, 1.50, 2.25, 0.32, TYPO.body, { bold: true, color: p.primary });
  addText(slide, `总样本量 ${Number(visual.payload.total_n).toLocaleString()}｜死亡事件 ${Number(visual.payload.total_events).toLocaleString()}`, 3.05, 1.50, 5.60, 0.32, TYPO.body, { bold: true, color: p.accent });
  addText(slide, "点的位置表示每项研究样本量；颜色表示来源中的偏倚风险概括", 8.28, 1.50, 4.15, 0.32, TYPO.chart, { color: SHARED.muted, align: "right" });
  const legendWidth = 1.40;
  // Keep the compact legend below the numeric summary, away from the
  // renderer-sensitive explanatory sentence at the upper right.
  const legendX = 3.05;
  robLegend.forEach((item, index) => {
    const x = legendX + index * legendWidth;
    ellipse(slide, x, 1.87, 0.14, 0.14, item.color, item.color, `${visual.visual_id}:rob-legend:${item.category}`);
    addText(slide, item.label, x + 0.20, 1.79, 1.12, 0.30, TYPO.chart, {
      color: SHARED.muted,
      name: `${visual.visual_id}:rob-legend-label:${item.category}`,
    });
  });
  const sizes = rows.map(row => Number(row.sample_size));
  const maxSize = Math.max(...sizes), scaleX = linearScale(0, maxSize, 4.55, 11.20);
  const plotTop = 2.17, plotBottom = 6.12, rowH = (plotBottom - plotTop) / rows.length;
  rows.forEach((row, index) => {
    const y = plotTop + index * rowH, x = scaleX(Number(row.sample_size));
    const studyLabel = String(row.study_label || "").trim();
    const year = String(row.year || "").trim();
    const labelAlreadyContainsYear = /\b(?:19|20)\d{2}\b/.test(studyLabel);
    const displayLabel = year && !labelAlreadyContainsYear ? `${studyLabel} · ${year}` : studyLabel;
    addText(slide, displayLabel, 0.82, y - 0.01, 3.48, Math.max(0.23, rowH - 0.01), TYPO.chart, { color: SHARED.ink });
    line(slide, 4.55, y + rowH * 0.42, Math.max(0.03, x - 4.55), 0, "CAD7E1", 1.0);
    const color = riskOfBiasSemantic(row.risk_of_bias, visual.visual_id).color;
    ellipse(slide, x - 0.065, y + rowH * 0.42 - 0.065, 0.13, 0.13, color, undefined, rowName(visual, index));
    addText(slide, `n=${Number(row.sample_size).toLocaleString()}`, 11.35, y - 0.01, 1.05, Math.max(0.23, rowH - 0.01), TYPO.chart, { align: "right", color: SHARED.muted });
  });
  addAxis(slide, 4.55, 6.35, 6.65, [0, maxSize / 2, maxSize], scaleX, "单项研究样本量", p, value => String(Math.round(value)));
  return {
    riskOfBiasLegend: robLegend.map(item => ({
      category: item.category,
      label: item.label,
      color: item.color,
    })),
  };
}

function renderPooledPanel(slide, visual, spec) {
  const p = profile(spec), rows = visual.payload.rows || [];
  if (rows.length !== 3) throw new Error(`Pooled evidence panel requires fixed, random, and prediction rows: ${visual.visual_id}`);
  const entries = rows.map(row => {
    const estimand = String(row.estimand || "");
    if (/prediction interval/i.test(estimand)) return { heading: "预测区间", value: `${row.ci_low}–${row.ci_high}`, detail: "未来研究真实效应的可能范围", prediction: true };
    if (/random/i.test(estimand)) return { heading: "随机效应OR", value: row.estimate, detail: `95% CI ${row.ci_low}–${row.ci_high}；I² ${row.I2_percent}%` };
    if (/fixed/i.test(estimand)) return { heading: "固定效应OR", value: row.estimate, detail: `95% CI ${row.ci_low}–${row.ci_high}` };
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
  const x0 = 4.38, y0 = 1.58, colW = 1.38, rowH = 0.335;
  domains.forEach((domain, index) => addText(slide, zhLabel(domain), x0 + index * colW - 0.08, 1.14, colW, 0.40, TYPO.chart, { bold: true, align: "center", color: SHARED.muted }));
  studies.forEach((study, rowIndex) => {
    const y = y0 + rowIndex * rowH;
    addText(slide, study, 0.82, y - 0.01, 3.32, rowH - 0.01, TYPO.chart, { color: SHARED.ink });
    domains.forEach((domain, colIndex) => {
      const item = rows.find(row => row.study_label === study && row.domain === domain);
      const fill = item ? riskOfBiasSemantic(item.judgment, visual.visual_id).color : "D9E2E8";
      const sourceIndex = item ? rows.indexOf(item) : -1;
      slide.addShape(S.rect, { x: x0 + colIndex * colW, y, w: colW - 0.08, h: rowH - 0.055, fill: { color: fill }, line: { color: SHARED.white, width: 0.6 }, name: sourceIndex >= 0 ? rowName(visual, sourceIndex) : `${visual.visual_id}:missing:${rowIndex}-${colIndex}` });
    });
  });
  const observed = new Set(rows.map(row => String(row.judgment ?? "").trim()));
  const legend = RISK_OF_BIAS_ORDER
    .filter(category => observed.has(category))
    .map(category => riskOfBiasSemantic(category, visual.visual_id));
  legend.forEach((item, index) => {
    slide.addShape(S.rect, { x: 8.05 + index * 1.62, y: 6.48, w: 0.24, h: 0.18, fill: { color: item.color }, line: { color: item.color } });
    addText(slide, item.label, 8.35 + index * 1.62, 6.40, 1.15, 0.30, TYPO.chart, { color: SHARED.muted });
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
  // Same-time assessments require 2D packing: labels may move horizontally
  // within a top/bottom swimlane while a leader preserves the true time point.
  const lanes = [
    { side: "top", visitY: 1.48, assessmentY: 1.82, intervals: [] },
    { side: "bottom", visitY: 3.90, assessmentY: 4.24, intervals: [] },
  ];
  const labelW = 1.92, gutter = 0.10;
  const minLabelX = SAFE_ZONES.content.x;
  const maxLabelX = SAFE_ZONES.content.x + SAFE_ZONES.content.w - labelW;
  const intervalAvailable = (lane, x) => lane.intervals.every(interval => x + labelW + gutter <= interval.left || x >= interval.right + gutter);
  const findTimelineSlot = (lane, idealX) => {
    const candidates = [
      clamp(idealX, minLabelX, maxLabelX), minLabelX, maxLabelX,
      ...lane.intervals.flatMap(interval => [interval.right + gutter, interval.left - gutter - labelW]),
    ].filter(x => x >= minLabelX && x <= maxLabelX && intervalAvailable(lane, x));
    if (!candidates.length) return null;
    return candidates.sort((a, b) => Math.abs(a - idealX) - Math.abs(b - idealX))[0];
  };
  const estimatedLabelHeight = row => {
    const text = `${zhLabel(row.assessment)}\n${zhLabel(row.window)}`;
    const lineBudget = Math.max(9, Math.floor(labelW * 6.2));
    const lines = text.split("\n").reduce((total, part) => total + Math.max(1, Math.ceil([...part].length / lineBudget)), 0);
    return clamp(0.23 * lines + 0.18, 0.64, 1.40);
  };
  [...rows].map((row, originalIndex) => ({ row, originalIndex })).sort((a, b) => Number(a.row.position_hours) - Number(b.row.position_hours)).forEach(({ row, originalIndex }) => {
    const hour = Number(row.position_hours);
    const x = hour <= 72 ? 1.0 + hour / 72 * 8.60 : 11.65;
    const idealX = clamp(x - labelW / 2, minLabelX, maxLabelX);
    const candidatePlacements = lanes.map(lane => ({ lane, labelX: findTimelineSlot(lane, idealX) })).filter(item => item.labelX !== null);
    if (!candidatePlacements.length) throw new Error(`Assessment timeline 2D packing budget exceeded near hour ${hour}`);
    candidatePlacements.sort((a, b) => Math.abs((a.labelX + labelW / 2) - x) - Math.abs((b.labelX + labelW / 2) - x) || a.lane.intervals.length - b.lane.intervals.length);
    const { lane, labelX } = candidatePlacements[0];
    lane.intervals.push({ left: labelX, right: labelX + labelW });
    ellipse(slide, x - 0.10, 3.45, 0.20, 0.20, row.status === "Primary endpoint" ? p.secondary : p.accent, undefined, rowName(visual, originalIndex));
    const labelCenter = labelX + labelW / 2;
    const leaderEndY = lane.side === "top" ? lane.assessmentY + estimatedLabelHeight(row) + 0.04 : lane.visitY - 0.08;
    line(slide, x, 3.45, labelCenter - x, leaderEndY - 3.45, SHARED.line, 0.7);
    addText(slide, zhLabel(row.visit), labelX, lane.visitY, labelW, 0.26, TYPO.chart, { bold: true, align: "center", color: p.primary });
    addText(slide, `${zhLabel(row.assessment)}\n${zhLabel(row.window)}`, labelX, lane.assessmentY, labelW, estimatedLabelHeight(row), TYPO.chart, { align: "center", color: SHARED.ink, valign: "top" });
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
  const p = profile(spec), nodes = visual.payload.nodes || [], edges = visual.payload.edges || [];
  if (!nodes.length || !edges.length) throw new Error(`DAG ${visual.visual_id} requires explicit nodes and edges`);
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
  let routedEdgeCount = 0;
  const routeDAGEdge = (edge, index) => {
    const a = positions[edge.source], b = positions[edge.target];
    if (!a || !b) throw new Error(`DAG edge references missing node: ${edge.source}->${edge.target}`);
    const startX = a.x + nodeWidth, startY = a.y + nodeHeight / 2;
    const endX = b.x, endY = b.y + nodeHeight / 2;
    const strokeWidth = nodes.find(node => node.id === edge.source)?.role === "exposure" ? 2.4 : 1.5;
    const obstructed = nodes.some(node => {
      if (node.id === edge.source || node.id === edge.target) return false;
      const obstacle = positions[node.id];
      if (!obstacle || obstacle.x <= startX || obstacle.x + nodeWidth >= endX) return false;
      const progress = (obstacle.x + nodeWidth / 2 - startX) / Math.max(0.001, endX - startX);
      const lineY = startY + progress * (endY - startY);
      return lineY >= obstacle.y - 0.08 && lineY <= obstacle.y + nodeHeight + 0.08;
    });
    const edgeName = `${visual.visual_id}:edge:${index}`;
    if (!obstructed && endX > startX) {
      line(slide, startX, startY, endX - startX, endY - startY, p.accent, strokeWidth, { endArrowType: "triangle", name: edgeName });
      return;
    }
    // Cross-layer edges are routed above the node field so a confounder-to-
    // outcome path cannot disappear behind exposure/mediator boxes.
    const routeY = 1.40 + Math.min(0.07, routedEdgeCount * 0.035);
    routedEdgeCount += 1;
    const sourceStubX = startX + 0.18, targetStubX = endX - 0.18;
    line(slide, startX, startY, sourceStubX - startX, 0, p.accent, strokeWidth, { name: `${edgeName}:a` });
    line(slide, sourceStubX, startY, 0, routeY - startY, p.accent, strokeWidth, { name: `${edgeName}:b` });
    line(slide, sourceStubX, routeY, targetStubX - sourceStubX, 0, p.accent, strokeWidth, { name: `${edgeName}:c` });
    line(slide, targetStubX, routeY, 0, endY - routeY, p.accent, strokeWidth, { name: `${edgeName}:d` });
    line(slide, targetStubX, endY, endX - targetStubX, 0, p.accent, strokeWidth, { endArrowType: "triangle", name: `${edgeName}:e` });
  };
  edges.forEach(routeDAGEdge);
  nodes.forEach((node, index) => {
    const pos = positions[node.id];
    const fill = node.role === "exposure" ? p.soft : node.role === "outcome" ? p.soft2 : SHARED.white;
    const stroke = node.role === "outcome" ? p.secondary : p.accent;
    box(slide, pos.x, pos.y, nodeWidth, nodeHeight, fill, stroke, 0.03, rowName(visual, index));
    addText(slide, zhLabel(node.label), pos.x + 0.12, pos.y + 0.10, nodeWidth - 0.24, nodeHeight - 0.20, TYPO.chart, { bold: true, align: "center", color: p.primary });
  });
  box(slide, 2.10, 5.55, 9.20, 0.98, "FFF4E5", SHARED.warning, 0.03);
  addText(slide, `预设调整集：${(visual.payload.adjustment_set || []).map(zhLabel).join("、")}`, 2.32, 5.66, 8.76, 0.28, TYPO.chart, { bold: true, color: SHARED.warning, align: "center" });
  if (visual.payload.review_note) addText(slide, localizeAudienceText(visual.payload.review_note), 2.32, 6.02, 8.76, 0.28, TYPO.chart, { color: SHARED.warning, align: "center" });
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
  // Keep the horizontal axis title below the two-line slide-title bound and
  // wide enough to remain a single line in both PowerPoint and LibreOffice.
  addText(slide, "影响 ↑", x0 - 1.08, 1.40, 0.98, 0.28, TYPO.chart, { bold: true, color: SHARED.muted, align: "right" });
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
    addText(slide, localizeAudienceText(`${item.conflicting_value} → ${item.canonical_value}`), 1.02, y + 0.13, 5.32, 0.32, TYPO.detail, { bold: true });
  });
  unresolved.forEach((item, index) => {
    const y = 2.03 + index * 0.86;
    box(slide, 6.90, y, 5.45, 0.65, "FFF4E5", SHARED.warning, 0.02);
    addText(slide, `${item.marker}｜INFORMATION_REQUIRED`, 7.10, y + 0.13, 5.05, 0.32, TYPO.detail, { bold: true, color: SHARED.warning });
  });
  addText(slide, "所有冲突和未解决项均保留在审计产物及人工审核清单中。", 0.84, 6.25, 11.45, 0.34, TYPO.body, { color: SHARED.muted, align: "center" });
}

function fitCreatedObjectsToHero(created, spec, visual) {
  const geometry = presentationGeometry(spec, visual);
  const measurable = created.filter(object => {
    const options = object?.options || {};
    return [options.x, options.y, options.w, options.h].every(Number.isFinite);
  });
  if (!measurable.length) return { contentScale: 1, heroBounds: geometry.hero, heroShare: geometry.heroShare, warnings: [...geometry.warnings] };
  const source = measurable.reduce((accumulator, object) => {
    const options = object.options;
    accumulator.x0 = Math.min(accumulator.x0, options.x);
    accumulator.y0 = Math.min(accumulator.y0, options.y);
    accumulator.x1 = Math.max(accumulator.x1, options.x + Math.max(0.001, options.w));
    accumulator.y1 = Math.max(accumulator.y1, options.y + Math.max(0.001, options.h));
    return accumulator;
  }, { x0: Number.POSITIVE_INFINITY, y0: Number.POSITIVE_INFINITY, x1: Number.NEGATIVE_INFINITY, y1: Number.NEGATIVE_INFINITY });
  const sourceWidth = source.x1 - source.x0;
  const sourceHeight = source.y1 - source.y0;
  if (!(sourceWidth > 0 && sourceHeight > 0)) throw new Error(`Cannot measure editable visual ${visual.visual_id}`);
  const scale = Math.min(1, geometry.hero.w / sourceWidth, geometry.hero.h / sourceHeight);
  const fittedWidth = sourceWidth * scale;
  const fittedHeight = sourceHeight * scale;
  const offsetX = geometry.hero.x + (geometry.hero.w - fittedWidth) / 2;
  const offsetY = geometry.hero.y + (geometry.hero.h - fittedHeight) / 2;
  for (const object of measurable) {
    const options = object.options;
    options.x = offsetX + (options.x - source.x0) * scale;
    options.y = offsetY + (options.y - source.y0) * scale;
    options.w = Math.max(0, options.w * scale);
    options.h = Math.max(0, options.h * scale);
    assertSafeContent({ x: options.x, y: options.y, w: options.w, h: options.h }, `${visual.visual_id} editable object`);
  }
  const warnings = [...geometry.warnings];
  if (scale < 0.78) warnings.push("editable_shape_fit_scale_below_0.78_requires_actual_text_bounds_QA");
  return { contentScale: scale, heroBounds: geometry.hero, heroShare: geometry.heroShare, warnings };
}

function collectSlideAnnotations(spec) {
  const annotations = [];
  for (const visual of spec.visual_specs || []) {
    for (const [index, item] of (visual.annotation_specs || []).entries()) {
      const normalized = normalizeAnnotation(item, index);
      if (normalized) {
        const enriched = {
          ...normalized,
          visual_id: visual.visual_id,
          visual_type: visual.render_visual_type || visual.visual_type,
          source_bindings: visual.source_bindings || [],
        };
        annotations.push({ ...enriched, external_panel: annotationNeedsExternalPanel(enriched) });
      }
    }
  }
  const primary = annotations.filter(item => item.role === "primary");
  const secondary = annotations.filter(item => item.role === "secondary");
  if (primary.length > 1 || secondary.length > 1) throw new Error(`Slide ${spec.slide_id} violates AnnotationSpec 1+1`);
  return { primary: primary[0] || null, secondary: secondary[0] || null, all: annotations };
}

function renderSlideAnnotations(slide, spec) {
  const annotationSet = collectSlideAnnotations(spec);
  if (!annotationSet.all.length) return { primaryCount: 0, secondaryCount: 0, annotations: [], warnings: [] };
  const panelAnnotations = annotationSet.all.filter(item => item.external_panel);
  if (!panelAnnotations.length) {
    return {
      primaryCount: annotationSet.primary ? 1 : 0,
      secondaryCount: annotationSet.secondary ? 1 : 0,
      annotations: annotationSet.all.map(item => ({ annotation_id: item.annotation_id, annotation_type: item.annotation_type, priority: item.role, visual_id: item.visual_id, text: item.text, source_bindings: item.source_bindings, external_panel: false })),
      warnings: [],
    };
  }
  const panelPrimary = panelAnnotations.find(item => item.role === "primary") || null;
  const panelSecondary = panelAnnotations.find(item => item.role === "secondary") || null;
  const geometry = presentationGeometry(spec);
  if (!geometry.annotation) throw new Error(`Slide ${spec.slide_id} has annotations without a reserved annotation zone`);
  const p = profile(spec);
  const zone = geometry.annotation;
  const gap = 0.28;
  const both = Boolean(panelPrimary && panelSecondary);
  const primaryHeight = both ? Math.min(2.55, (zone.h - gap) * 0.56) : zone.h;
  const secondaryHeight = both ? zone.h - primaryHeight - gap : zone.h;
  const renderOne = (annotation, bounds, isPrimary) => {
    assertSafeContent(bounds, `${spec.slide_id} annotation`);
    const fill = isPrimary ? p.soft : "FFF4E5";
    const accent = isPrimary ? p.accent : SHARED.warning;
    box(slide, bounds.x, bounds.y, bounds.w, bounds.h, fill, accent, 0.03, `${spec.slide_id}:annotation:${annotation.annotation_id}`);
    slide.addShape(S.rect, { x: bounds.x, y: bounds.y, w: 0.08, h: bounds.h, fill: { color: accent }, line: { color: accent }, name: `${spec.slide_id}:annotation-rule:${annotation.annotation_id}` });
    addText(slide, annotation.title || (isPrimary ? "关键注释" : "证据边界"), bounds.x + 0.22, bounds.y + 0.18, bounds.w - 0.38, 0.34, TYPO.detail, { bold: true, color: isPrimary ? p.primary : SHARED.warning, valign: "top" });
    addText(slide, annotationAudienceText(annotation), bounds.x + 0.22, bounds.y + 0.66, bounds.w - 0.40, bounds.h - 0.84, isPrimary ? TYPO.body : TYPO.detail, { bold: isPrimary, color: SHARED.ink, valign: "top", name: `${spec.slide_id}:annotation-text:${annotation.annotation_id}` });
  };
  if (panelPrimary) renderOne(panelPrimary, { x: zone.x, y: zone.y, w: zone.w, h: primaryHeight }, true);
  if (panelSecondary) renderOne(panelSecondary, { x: zone.x, y: both ? zone.y + primaryHeight + gap : zone.y, w: zone.w, h: secondaryHeight }, false);
  return {
    primaryCount: annotationSet.primary ? 1 : 0,
    secondaryCount: annotationSet.secondary ? 1 : 0,
    annotations: annotationSet.all.map(item => ({ annotation_id: item.annotation_id, annotation_type: item.annotation_type, priority: item.role, visual_id: item.visual_id, text: item.text, source_bindings: item.source_bindings, external_panel: item.external_panel })),
    warnings: [...geometry.warnings],
  };
}

function renderVisual(slide, visual, spec) {
  const before = (slide._slideObjects || []).length;
  const renderMode = resolveRenderMode(visual, spec);
  const renderVisualType = String(visual.render_visual_type || visual.visual_type);
  let modeEvidence = {};
  if (renderMode === "native_chart") {
    modeEvidence = renderNativeChart(slide, visual, spec);
  } else if (renderMode === "source_figure" || renderMode === "svg") {
    modeEvidence = renderSourceFigure(slide, visual, spec, renderMode);
  } else {
    switch (renderVisualType) {
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
      case "study_characteristics": modeEvidence = renderStudyCharacteristics(slide, visual, spec) || {}; break;
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
      default: throw new Error(`No editable-shape renderer registered for VisualSpec type ${renderVisualType}`);
    }
  }
  const created = (slide._slideObjects || []).slice(before);
  if (renderMode === "editable_shapes") {
    // Text-bearing scientific visuals are authored directly in the shared safe
    // zone. Post-hoc x/y/w/h scaling changes the boxes without scaling fonts or
    // paragraph margins, which makes dense forests and numeric labels wrap and
    // collide. Layout variants and annotations reserve space before rendering;
    // editable objects are never compressed after creation.
    const geometry = presentationGeometry(spec, visual);
    const fitEvidence = {
      contentScale: 1,
      heroBounds: geometry.hero,
      heroShare: geometry.heroShare,
      warnings: [...geometry.warnings, "editable_shapes_intrinsic_safe_zone_layout"],
    };
    const rendererEvidence = modeEvidence;
    modeEvidence = {
      objectType: "native_powerpoint_shapes",
      editability: "full",
      nativeChart: false,
      chartType: null,
      chartSeriesCount: null,
      chartCategoryCount: null,
      chartWorkbookSource: null,
      chartWorkbookFields: [],
      nativeChartPrivacy: {},
      nativeChartPrivacyValidation: null,
      assetPath: null,
      sourceFigureTreatment: null,
      fallbackUsed: false,
      ...fitEvidence,
      ...rendererEvidence,
      warnings: [...new Set([...(fitEvidence.warnings || []), ...(rendererEvidence.warnings || [])])],
    };
  }
  const createdObjects = created.length;
  const prefix = `${visual.visual_id}:row:`;
  const namedRowKeys = [...new Set(created.map(object => object?.options?.name || "").filter(name => name.startsWith(prefix)).map(name => name.slice(prefix.length)))];
  const renderedRowKeys = modeEvidence.renderedRowKeys || namedRowKeys;
  const expectedRowKeys = visual.expected_row_keys || [];
  const missing = expectedRowKeys.filter(key => !renderedRowKeys.includes(key));
  const unexpected = renderedRowKeys.filter(key => !expectedRowKeys.includes(key));
  if (createdObjects <= 0) throw new Error(`Renderer created no PowerPoint objects for ${visual.visual_id}`);
  if (missing.length || unexpected.length || renderedRowKeys.length !== expectedRowKeys.length) {
    throw new Error(`Renderer row evidence mismatch for ${visual.visual_id}: missing=${missing.join(",")} unexpected=${unexpected.join(",")}`);
  }
  const geometry = presentationGeometry(spec, visual);
  return {
    renderedRows: renderedRowKeys.length,
    renderedRowKeys,
    createdObjects,
    renderMode,
    renderVisualType,
    objectType: modeEvidence.objectType,
    editability: modeEvidence.editability,
    nativeChart: Boolean(modeEvidence.nativeChart),
    chartType: modeEvidence.chartType ?? null,
    chartSeriesCount: modeEvidence.chartSeriesCount ?? null,
    chartCategoryCount: modeEvidence.chartCategoryCount ?? null,
    chartWorkbookSource: modeEvidence.chartWorkbookSource ?? null,
    chartWorkbookFields: modeEvidence.chartWorkbookFields || [],
    directValueLabels: modeEvidence.directValueLabels ?? null,
    dataLabelPolicy: modeEvidence.dataLabelPolicy ?? null,
    nativeChartPrivacy: modeEvidence.nativeChartPrivacy || {},
    nativeChartPrivacyValidation: modeEvidence.nativeChartPrivacyValidation || null,
    funnelSemanticOverlay: modeEvidence.funnelSemanticOverlay || null,
    riskOfBiasLegend: modeEvidence.riskOfBiasLegend || null,
    assetPath: modeEvidence.assetPath ?? null,
    sourceFigureTreatment: modeEvidence.sourceFigureTreatment ?? null,
    figureRebuildDecision: visual.figure_rebuild_decision ?? null,
    heroBounds: modeEvidence.heroBounds || geometry.hero,
    heroShare: modeEvidence.heroShare ?? geometry.heroShare,
    contentScale: modeEvidence.contentScale ?? 1,
    fallbackUsed: Boolean(modeEvidence.fallbackUsed),
    warnings: [...new Set([...(modeEvidence.warnings || []), ...geometry.warnings])],
  };
}

const objectManifest = [];
const renderEvidence = {};
const annotationEvidence = {};
const renderContractSlides = [];
const slideRenderStarted = performance.now();
for (const [slideIndex, spec] of graph.slide_specs.entries()) {
  const slide = pptx.addSlide(masterName(spec.layout_family));
  applySlideBackground(slide, spec);
  if (spec.slide_role === "cover") {
    renderCover(slide, spec, graph.kind);
  } else {
    title(slide, spec);
    if (spec.slide_role === "appendix") renderReview(slide, spec);
    else for (const visual of spec.visual_specs || []) renderEvidence[visual.visual_id] = renderVisual(slide, visual, spec);
    annotationEvidence[spec.slide_id] = renderSlideAnnotations(slide, spec);
    footer(slide, spec, slideIndex + 1);
  }
  slide.addNotes(speakerNotes(spec));
  const fields = presentationFields(spec);
  const slideGeometry = presentationGeometry(spec);
  renderContractSlides.push({
    slide_id: spec.slide_id,
    slide_role: spec.slide_role,
    layout_family: spec.layout_family,
    layout_variant: fields.layoutVariant,
    background_variant: fields.backgroundVariant,
    background_treatment: fields.backgroundTreatment,
    visual_intensity: fields.visualIntensity,
    density_target: fields.densityTarget,
    hero_enabled: fields.heroEnabled,
    requested_hero_area_ratio: fields.heroEnabled ? fields.requestedHeroShare : null,
    rendered_hero_area_ratio: fields.heroEnabled ? slideGeometry.heroShare : null,
    hero_bounds: fields.heroEnabled ? slideGeometry.hero : null,
    annotation_primary_count: annotationEvidence[spec.slide_id]?.primaryCount || 0,
    annotation_secondary_count: annotationEvidence[spec.slide_id]?.secondaryCount || 0,
    annotation_specs: annotationEvidence[spec.slide_id]?.annotations || [],
    rhythm_beat: fields.rhythmBeat,
    rhythm_index: fields.rhythmIndex,
    safe_zones: SAFE_ZONES,
    warnings: [...new Set([...(slideGeometry.warnings || []), ...(annotationEvidence[spec.slide_id]?.warnings || [])])],
  });
  for (const visual of spec.visual_specs || []) {
    const evidence = renderEvidence[visual.visual_id] || {};
    objectManifest.push({
      renderer_version: "2.5",
      slide_id: spec.slide_id,
      visual_id: visual.visual_id,
      visual_type: visual.visual_type,
      render_visual_type: evidence.renderVisualType || visual.render_visual_type || visual.visual_type,
      render_mode: evidence.renderMode || visual.render_mode || "editable_shapes",
      object_type: evidence.objectType || "native_powerpoint_shapes",
      editability: evidence.editability || visual.editability_requirement || "full",
      data_contract_id: visual.data_contract_id,
      source_bindings: visual.source_bindings,
      expected_row_count: visual.expected_row_count,
      rendered_row_count: evidence.renderedRows ?? 0,
      rendered_row_keys: evidence.renderedRowKeys ?? [],
      expected_row_keys: visual.expected_row_keys,
      created_object_count: evidence.createdObjects ?? 0,
      layout_family: spec.layout_family,
      layout_variant: fields.layoutVariant,
      background_variant: fields.backgroundVariant,
      visual_intensity: fields.visualIntensity,
      density_target: fields.densityTarget,
      hero_enabled: fields.heroEnabled,
      hero_area_ratio: evidence.heroShare ?? null,
      hero_share: evidence.heroShare ?? null,
      hero_bounds: evidence.heroBounds ?? null,
      content_scale: evidence.contentScale ?? 1,
      native_chart: Boolean(evidence.nativeChart),
      chart_type: evidence.chartType ?? null,
      chart_series_count: evidence.chartSeriesCount ?? null,
      chart_category_count: evidence.chartCategoryCount ?? null,
      chart_workbook_source: evidence.chartWorkbookSource ?? null,
      chart_workbook_fields: evidence.chartWorkbookFields || [],
      direct_value_labels: evidence.directValueLabels ?? null,
      data_label_policy: evidence.dataLabelPolicy ?? null,
      native_chart_privacy: evidence.nativeChartPrivacy || visual.native_chart_privacy || {},
      native_chart_privacy_validation: evidence.nativeChartPrivacyValidation || null,
      funnel_semantic_overlay: evidence.funnelSemanticOverlay || null,
      risk_of_bias_legend: evidence.riskOfBiasLegend || null,
      annotation_specs: (visual.annotation_specs || []).map(item => ({ annotation_type: item.annotation_type || "direct_annotation", priority: item.priority || item.role || "secondary", text: item.text })),
      annotation_count: (visual.annotation_specs || []).length,
      annotation_primary_count: (visual.annotation_specs || []).filter(item => (item.priority || item.role) === "primary").length,
      annotation_secondary_count: (visual.annotation_specs || []).filter(item => (item.priority || item.role) === "secondary").length,
      figure_rebuild_decision: evidence.figureRebuildDecision ?? visual.figure_rebuild_decision ?? null,
      source_figure_treatment: evidence.sourceFigureTreatment ?? null,
      asset_path: evidence.assetPath ?? null,
      artifact_path: evidence.assetPath ?? null,
      fallback_used: Boolean(evidence.fallbackUsed),
      warnings: evidence.warnings || [],
    });
  }
}
const slideRenderMilliseconds = performance.now() - slideRenderStarted;

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
const packageWriteStarted = performance.now();
await pptx.writeFile({ fileName: outputPath });
const packageWriteMilliseconds = performance.now() - packageWriteStarted;
const manifestPath = outputPath.replace(/\.pptx$/i, ".object_manifest.json");
fs.writeFileSync(manifestPath, JSON.stringify(objectManifest, null, 2), "utf8");
const renderContractPath = outputPath.replace(/\.pptx$/i, ".render_contract.json");
fs.writeFileSync(renderContractPath, JSON.stringify({
  renderer_version: "2.5",
  graph_version: graph.graph_version,
  art_direction_id: ART_DIRECTION_SPEC.art_direction_id || null,
  backend: "native_pptxgenjs",
  external_skill_used: false,
  safe_zones: SAFE_ZONES,
  slides: renderContractSlides,
}, null, 2), "utf8");
console.log(`WROTE=${outputPath}`);
console.log(`SLIDES=${graph.slide_specs.length}`);
console.log(`OBJECT_MANIFEST=${manifestPath}`);
console.log(`RENDER_CONTRACT=${renderContractPath}`);
console.log(`SLIDE_RENDER_MS=${slideRenderMilliseconds.toFixed(2)}`);
console.log(`PACKAGE_WRITE_MS=${packageWriteMilliseconds.toFixed(2)}`);
