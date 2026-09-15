// VISUAL_FIDELITY_EXPERIMENT_001 only. No canonical IR, inference, or asset loading.
export const EXPERIMENT_ID = "VISUAL_FIDELITY_EXPERIMENT_001";
export const ARCHETYPES = Object.freeze([
  "PROCESS_HORIZONTAL", "FRAMEWORK_HUB", "SYSTEM_MAP", "EDITORIAL_TAKEAWAY",
]);

const normalize = (s) => s.replace(/\s+/gu, " ").trim();
const bag = (items) => JSON.stringify(items.map(normalize).sort());
function fail(item, plan, reason, code = "INVALID_EXPERIMENTAL_PLAN") {
  throw new Error(`${code}: ${item.slide_id}/${plan.archetype}: ${reason}`);
}

// Presentation geometry only, adapted from scripts/v2_5/render_v2_5_deck.mjs
// clippedEllipseDirectedEdge at stable main 4e288617. No legacy runtime or
// scientific derivation is imported. Endpoints stop outside ellipse boundaries.
export function clippedEllipseEdge(a, b, gap = 0.045) {
  const ax = a.x + a.w / 2, ay = a.y + a.h / 2;
  const bx = b.x + b.w / 2, by = b.y + b.h / 2;
  const dx = bx - ax, dy = by - ay, length = Math.hypot(dx, dy);
  if (!length) throw new Error("Coincident experimental nodes");
  const start = 1 / Math.hypot(dx / (a.w / 2), dy / (a.h / 2));
  const end = 1 / Math.hypot(dx / (b.w / 2), dy / (b.h / 2));
  if (start + end + 2 * gap / length >= 1) throw new Error("Overlapping experimental nodes");
  return { x1: ax + dx * (start + gap / length), y1: ay + dy * (start + gap / length),
    x2: bx - dx * (end + gap / length), y2: by - dy * (end + gap / length) };
}

export function compose(item, plan, contract) {
  if (!ARCHETYPES.includes(plan.archetype)) fail(item, plan, "unknown archetype");
  // This experiment supports the explicitly frozen takeaway-container fixture.
  // Never silently replace a chart, source figure, template, or cover.
  if (item.visual_type !== "takeaway" || item.slide_role === "cover" || item.template_layout_id) {
    fail(item, plan, "unsupported target; scientific/template renderers are protected");
  }
  const entry = (item.planned_geometry || []).find((o) => o.object_id.endsWith(":structure"))
    || (item.planned_geometry || []).find((o) => o.object_id.endsWith(":visual"));
  const frame = entry?.bounds;
  const incompatible = (reason) => fail(item, plan, reason, "EXPERIMENTAL_ARCHETYPE_FRAME_INCOMPATIBLE");
  if (!frame || ![frame.x, frame.y, frame.w, frame.h].every(Number.isFinite)
      || frame.w < 10 || frame.h < 2.8) incompatible("requires a safe frame of at least 10 x 2.8 inches");
  const minFont = Math.max(18, Number(contract?.minimum_font_pt?.body || 18));
  const units = (item.takeaways || []).flatMap((s) => s.split("\n")).filter((s) => s.trim());
  const objects = [];
  const rect = (x, y, w, h) => ({ x: frame.x + x * frame.w, y: frame.y + y * frame.h, w: w * frame.w, h: h * frame.h });
  function shape(kind, id, bounds, tone = "primary", anchor = false) {
    objects.push({ kind, id, bounds, tone, anchor });
  }
  function text(id, value, bounds, size = minFont, hero = false) {
    objects.push({ kind: "text", id, text: value, bounds, font_size_pt: Math.max(size, minFont), hero, anchor: hero });
  }
  if (plan.archetype === "EDITORIAL_TAKEAWAY") {
    if (units.length < 4 || units.length > 6) fail(item, plan, "requires four to six explicit paragraphs");
    shape("rect", "emphasis", rect(0.005, 0.06, 0.009, 0.87));
    text("takeaway-1", units[0], rect(0.04, 0.14, 0.30, 0.72), 28, true);
    const row = 0.94 / (units.length - 1);
    for (let i = 1; i < units.length; i++) {
      const y = 0.03 + (i - 1) * row;
      text(`takeaway-${i + 1}`, units[i], rect(0.41, y, 0.57, row - 0.035));
      if (i < units.length - 1) shape("line", `rule-${i}`, rect(0.41, y + row - 0.009, 0.56, 0), "line");
    }
  } else {
    const nodes = item.diagram_spec?.nodes, edges = item.diagram_spec?.edges;
    if (!Array.isArray(nodes) || !Array.isArray(edges) || nodes.length !== 4) fail(item, plan, "requires four explicit nodes and an edge list");
    if (new Set(nodes.map((n) => n.node_id)).size !== nodes.length || nodes.some((n) => !n.node_id || typeof n.label !== "string")) fail(item, plan, "invalid node IDs/labels");
    if (bag(nodes.map((n) => n.label)) !== bag(units)) fail(item, plan, "node text must exactly preserve the canonical paragraphs");
    const ids = nodes.map((n) => n.node_id), positions = new Map();
    const edgeKeys = edges.map((e) => `${e.source}>${e.target}`);
    if (new Set(edgeKeys).size !== edges.length || edges.some((e) => !ids.includes(e.source) || !ids.includes(e.target) || e.source === e.target || e.label)) {
      fail(item, plan, "invalid/duplicate endpoints or unsupported relationship labels");
    }
    if (plan.archetype === "PROCESS_HORIZONTAL") {
      if (JSON.stringify(plan.node_order) !== JSON.stringify(ids)
          || JSON.stringify(edgeKeys) !== JSON.stringify(ids.slice(1).map((id, i) => `${ids[i]}>${id}`))
          || edges.some((e) => e.directed !== true)) fail(item, plan, "linear order and directed edges must be explicitly frozen");
      nodes.forEach((n, i) => {
        const h = [0.29, 0.36, 0.36, 0.46][i];
        positions.set(n.node_id, rect(0.015 + i * 0.257, 0.50 - h / 2, 0.195, h));
      });
    } else if (plan.archetype === "FRAMEWORK_HUB") {
      if (!ids.includes(plan.central_id) || edges.length !== 3
          || edges.some((e) => e.source !== plan.central_id || e.directed === true)
          || new Set(edges.map((e) => e.target)).size !== 3) fail(item, plan, "hub membership must be explicitly frozen and undirected");
      positions.set(plan.central_id, rect(0.355, 0.29, 0.29, 0.42));
      const surrounding = [rect(0.02, 0.36, 0.235, 0.28), rect(0.73, 0.015, 0.25, 0.28), rect(0.73, 0.705, 0.25, 0.28)];
      nodes.filter((n) => n.node_id !== plan.central_id).forEach((n, i) => positions.set(n.node_id, surrounding[i]));
    } else {
      if (!plan.positions || bag(Object.keys(plan.positions)) !== bag(ids)) fail(item, plan, "system topology needs explicit node positions");
      for (const node of nodes) {
        const p = plan.positions[node.node_id];
        if (!Array.isArray(p) || p.length !== 2 || !p.every(Number.isFinite)) fail(item, plan, "invalid node position");
        positions.set(node.node_id, rect(p[0] - 0.12, p[1] - 0.125, 0.24, 0.25));
      }
    }
    edges.forEach((edge, i) => {
      const e = clippedEllipseEdge(positions.get(edge.source), positions.get(edge.target));
      const bounds = { x: Math.min(e.x1, e.x2), y: Math.min(e.y1, e.y2), w: Math.abs(e.x2 - e.x1), h: Math.abs(e.y2 - e.y1) };
      objects.push({ kind: "line", id: `edge-${i + 1}`, bounds, tone: "muted", edge: { ...edge },
        flipH: e.x2 < e.x1, flipV: e.y2 < e.y1, directed: edge.directed === true });
    });
    nodes.forEach((node, i) => {
      const b = positions.get(node.node_id);
      const hero = plan.archetype === "FRAMEWORK_HUB" ? node.node_id === plan.central_id : i === nodes.length - 1;
      shape("ellipse", `node-${node.node_id}`, b, hero ? "secondary" : "primary", hero);
      text(`label-${node.node_id}`, node.label, { x: b.x + 0.07, y: b.y + 0.03, w: b.w - 0.14, h: b.h - 0.06 }, minFont);
    });
  }
  if (bag(objects.filter((o) => o.kind === "text").map((o) => o.text)) !== bag(units)) fail(item, plan, "scientific text mutation");
  for (const o of objects) {
    const b = o.bounds, footer = contract?.zones?.footer_exclusion;
    if (![b.x, b.y, b.w, b.h].every(Number.isFinite) || b.w < 0 || b.h < 0
        || b.x < frame.x || b.y < frame.y || b.x + b.w > frame.x + frame.w + 1e-8 || b.y + b.h > frame.y + frame.h + 1e-8
        || b.x < 0 || b.y < 0 || b.x + b.w > contract.slide.width_in || b.y + b.h > contract.slide.height_in
        || (footer && b.y + b.h > footer.y)) incompatible(`object ${o.id} exceeds its approved frame`);
  }
  return { slide_id: item.slide_id, archetype_requested: plan.archetype, archetype_rendered: plan.archetype,
    generic_fallback: false, frame, objects };
}

export function prepareExperiment(spec) {
  if (spec.experiment?.id !== EXPERIMENT_ID || spec.experiment?.enabled !== true) throw new Error("EXPERIMENT_NOT_ENABLED");
  const plan = spec.visual_execution_plan;
  if (!plan || typeof plan !== "object" || Array.isArray(plan) || !Object.keys(plan).length) throw new Error("MISSING_EXPERIMENTAL_PLAN");
  const compositions = new Map();
  for (const [id, assignment] of Object.entries(plan)) {
    const item = spec.slides.find((s) => s.slide_id === id);
    if (!item) throw new Error(`Unknown experimental slide: ${id}`);
    compositions.set(id, compose(item, assignment, spec.layout_contract));
  }
  const rendered = [];
  return {
    has: (id) => compositions.has(id),
    render(slide, item, { pptx, C, font }) {
      const start = performance.now(), c = compositions.get(item.slide_id);
      for (const o of c.objects) {
        const common = { ...o.bounds, objectName: `vf001:${item.slide_id}:${o.id}` };
        if (o.kind === "text") {
          slide.addText(o.text, { ...common, fontFace: font, fontSize: o.font_size_pt,
            bold: o.hero || c.archetype_rendered !== "EDITORIAL_TAKEAWAY", color: o.hero ? C.primary : C.ink,
            margin: 0, valign: "mid", align: c.archetype_rendered === "EDITORIAL_TAKEAWAY" ? "left" : "center" });
        } else {
          slide.addShape(pptx.ShapeType[o.kind], { ...common, flipH: o.flipH, flipV: o.flipV,
            fill: { color: o.kind === "rect" ? C.primary : C.white },
            line: { color: C[o.tone], width: o.anchor ? 2.2 : 1.4, ...(o.directed ? { endArrowType: "triangle" } : {}) } });
        }
      }
      rendered.push({ ...c, runtime_seconds: (performance.now() - start) / 1000,
        object_count: c.objects.length, text_count: c.objects.filter((o) => o.kind === "text").length,
        visual_anchor_count: c.objects.filter((o) => o.anchor).length,
        primitive_distribution: c.objects.reduce((a, o) => ({ ...a, [o.kind]: (a[o.kind] || 0) + 1 }), {}) });
    },
    report: () => ({ experiment_id: EXPERIMENT_ID, targeted_slide_count: compositions.size,
      rendered_slide_count: rendered.length, generic_fallback_count: 0, slides: rendered }),
  };
}
