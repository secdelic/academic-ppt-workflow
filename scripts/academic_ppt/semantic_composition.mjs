// Source-bound native semantic composition. No inference or asset loading.
export const COMPOSITION_ID = "SEMANTIC_COMPOSITION";
export const ARCHETYPES = Object.freeze([
  "PROCESS_HORIZONTAL", "FRAMEWORK_HUB", "SYSTEM_MAP", "EDITORIAL_TAKEAWAY",
]);
export const SEMANTIC_FAMILIES = Object.freeze([
  "LAYERED_LOOP_FRAMEWORK", "COMPOSITE_DUAL_FRAMEWORK", "TIMELINE_WITH_BAND",
]);

const normalize = (s) => s.replace(/\s+/gu, " ").trim();
const bag = (items) => JSON.stringify(items.map(normalize).sort());
function fail(item, plan, reason, code = "INVALID_COMPOSITION_PLAN") {
  throw new Error(`${code}: ${item.slide_id}/${plan.archetype}: ${reason}`);
}

// Presentation geometry only, adapted from scripts/v2_5/render_v2_5_deck.mjs
// clippedEllipseDirectedEdge at stable main 4e288617. No legacy runtime or
// scientific derivation is imported. Endpoints stop outside ellipse boundaries.
export function clippedEllipseEdge(a, b, gap = 0.045) {
  const ax = a.x + a.w / 2, ay = a.y + a.h / 2;
  const bx = b.x + b.w / 2, by = b.y + b.h / 2;
  const dx = bx - ax, dy = by - ay, length = Math.hypot(dx, dy);
  if (!length) throw new Error("Coincident composition nodes");
  const start = 1 / Math.hypot(dx / (a.w / 2), dy / (a.h / 2));
  const end = 1 / Math.hypot(dx / (b.w / 2), dy / (b.h / 2));
  if (start + end + 2 * gap / length >= 1) throw new Error("Overlapping composition nodes");
  return { x1: ax + dx * (start + gap / length), y1: ay + dy * (start + gap / length),
    x2: bx - dx * (end + gap / length), y2: by - dy * (end + gap / length) };
}

export function compose(item, plan, contract) {
  if (plan.revision === "SEMANTIC") return composeSemantic(item, plan, contract);
  if (plan.revision === "COMPACT_CJK") return composeCjk(item, plan, contract);
  if (plan.revision) fail(item, plan, "unknown composition revision");
  if (!ARCHETYPES.includes(plan.archetype)) fail(item, plan, "unknown archetype");
  // This compositor requires an explicit native semantic plan.
  // Never silently replace a chart, source figure, template, or cover.
  if (item.visual_type !== "takeaway" || item.slide_role === "cover" || item.template_layout_id) {
    fail(item, plan, "unsupported target; scientific/template renderers are protected");
  }
  const entry = (item.planned_geometry || []).find((o) => o.object_id.endsWith(":structure"))
    || (item.planned_geometry || []).find((o) => o.object_id.endsWith(":visual"));
  const frame = entry?.bounds;
  const incompatible = (reason) => fail(item, plan, reason, "COMPOSITION_FRAME_INCOMPATIBLE");
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

// One explicitly selected CJK calibration. The original B0 path remains an
// immutable-behavior test oracle; this is not a canonical geometry algorithm.
function composeCjk(item, plan, contract) {
  if (!ARCHETYPES.includes(plan.archetype)) fail(item, plan, "unknown archetype");
  if (item.visual_type !== "takeaway" || item.template_layout_id || item.slide_role === "cover") fail(item, plan, "protected renderer");
  if (!plan.source_locator || !["equal", "equal_peripheral"].includes(plan.emphasis)) fail(item, plan, "explicit source and equal-weight semantics required");
  const frame = item.planned_geometry?.find((o) => o.object_id.endsWith(":structure"))?.bounds;
  const incompatible = (reason) => fail(item, plan, reason, "COMPOSITION_FRAME_INCOMPATIBLE");
  if (!frame || frame.w < 10 || frame.h < 2.8) incompatible("safe frame too small");
  const size = Number(plan.font_size_pt || 20), minimum = Math.max(18, contract.minimum_font_pt?.body || 18);
  if (!Number.isFinite(size) || size < minimum) incompatible("font below approved minimum");
  const units = item.takeaways.flatMap((s) => s.split("\n")).filter((s) => s.trim());
  const objects = [], bounds = (x,y,w,h) => ({x:frame.x+x*frame.w,y:frame.y+y*frame.h,w:w*frame.w,h:h*frame.h});
  function label(id, value, b, align = "center") {
    // Conservative fixed-size load check, not a substitute for actual PowerPoint
    // geometry. No shrinking, shortening, silent fallback, or automatic splitting.
    const em = [...value].reduce((n,c) => n + (/[^\x00-\x7F]/u.test(c) ? 1 : 0.6), 0);
    const capacity = b.w * 72 / size;
    const lines = Math.ceil(em / capacity);
    if (lines * size * 1.2 > b.h * 72 + 0.1) incompatible(`text ${id} requires ${lines} lines at ${size} pt; ${b.h.toFixed(3)} inch box`);
    objects.push({kind:"text",id,text:value,bounds:b,font_size_pt:size,align,bold:false,hero:false,anchor:false});
  }
  if (plan.archetype === "EDITORIAL_TAKEAWAY") {
    if (plan.emphasis !== "equal" || units.length < 4 || units.length > 6) fail(item,plan,"four to six equal statements required");
    const rows = Math.ceil(units.length/2), h = 0.94/rows;
    units.forEach((value,i) => {
      const col = Math.floor(i/rows), row = i%rows;
      const b = bounds(0.01+col*0.51,0.03+row*h,0.47,h-0.02);
      objects.push({kind:"roundRect",id:`statement-${i+1}`,bounds:b,tone:"line",anchor:false,line_width:1});
      label(`takeaway-${i+1}`,value,{x:b.x+0.13,y:b.y+0.055,w:b.w-0.26,h:b.h-0.11},"left");
    });
  } else {
    const nodes=item.diagram_spec?.nodes, edges=item.diagram_spec?.edges;
    if (!Array.isArray(nodes) || !Array.isArray(edges) || !nodes.length
        || JSON.stringify(nodes)!==JSON.stringify(plan.nodes) || JSON.stringify(edges)!==JSON.stringify(plan.edges)) fail(item,plan,"nodes and edges must match the frozen hashed plan");
    const ids=nodes.map((n)=>n.node_id), edgeKeys=edges.map((e)=>`${e.source}>${e.target}`);
    if (new Set(ids).size!==ids.length || bag(nodes.map((n)=>n.label))!==bag(units)
        || new Set(edgeKeys).size!==edgeKeys.length || edges.some((e)=>!ids.includes(e.source)||!ids.includes(e.target)||e.source===e.target||e.label)) fail(item,plan,"invalid nodes, text, or edges");
    const positions=new Map();
    if (plan.archetype==="PROCESS_HORIZONTAL") {
      if (nodes.length<2||nodes.length>6) incompatible("process supports two to six stages");
      if (plan.emphasis!=="equal"||JSON.stringify(ids)!==JSON.stringify(plan.node_order)
          ||JSON.stringify(edgeKeys)!==JSON.stringify(ids.slice(1).map((id,i)=>`${ids[i]}>${id}`))
          ||edges.some((e)=>e.directed!==true)||plan.relationship_type!=="explicit_sequence") fail(item,plan,"only explicitly sourced stage order is permitted");
      const gap=0.027,w=(0.98-gap*(nodes.length-1))/nodes.length;
      nodes.forEach((n,i)=>positions.set(n.node_id,bounds(0.01+i*(w+gap),0.12,w,0.76)));
    } else if (plan.archetype==="FRAMEWORK_HUB") {
      if (nodes.length!==4 || plan.emphasis!=="equal_peripheral" || !ids.includes(plan.central_id)
          ||edges.length!==3||edges.some((e)=>e.source!==plan.central_id||e.directed!==false)
          ||new Set(edges.map((e)=>e.target)).size!==3||plan.relationship_type!=="equal_dimensions") fail(item,plan,"one explicit center and three undirected equal dimensions required");
      positions.set(plan.central_id,bounds(0.37,0.31,0.26,0.27));
      const outer=[bounds(0.01,0.02,0.30,0.34),bounds(0.69,0.02,0.30,0.34),bounds(0.35,0.65,0.30,0.34)];
      nodes.filter((n)=>n.node_id!==plan.central_id).forEach((n,i)=>positions.set(n.node_id,outer[i]));
    } else {
      if (nodes.length<4||nodes.length>5) incompatible("system supports four to five nodes");
      if (plan.emphasis!=="equal"||!plan.positions||bag(Object.keys(plan.positions))!==bag(ids)
          ||edges.some((e)=>e.directed!==false)||plan.relationship_type!=="explicit_nondirectional_association") fail(item,plan,"explicit equal, undirected system topology required");
      for (const n of nodes) {
        const p=plan.positions[n.node_id];
        if (!Array.isArray(p)||p.length!==2||!p.every(Number.isFinite)) fail(item,plan,"invalid explicit position");
        positions.set(n.node_id,bounds(p[0]-0.145,p[1]-0.15,0.29,0.30));
      }
    }
    // Rectangle boundary clipping is local presentation logic. Connectors remain
    // independent native lines, not attachment-bound PowerPoint connectors.
    edges.forEach((edge,i)=>{
      const a=positions.get(edge.source),b=positions.get(edge.target);
      const dx=b.x+b.w/2-a.x-a.w/2,dy=b.y+b.h/2-a.y-a.h/2,len=Math.hypot(dx,dy);
      const ta=Math.min(dx ? a.w/2/Math.abs(dx) : Infinity,dy ? a.h/2/Math.abs(dy) : Infinity);
      const tb=Math.min(dx ? b.w/2/Math.abs(dx) : Infinity,dy ? b.h/2/Math.abs(dy) : Infinity);
      if (!len||ta+tb+0.09/len>=1) incompatible(`overlapping edge endpoints ${edge.source}/${edge.target}`);
      const x1=a.x+a.w/2+dx*(ta+0.045/len), y1=a.y+a.h/2+dy*(ta+0.045/len);
      const x2=b.x+b.w/2-dx*(tb+0.045/len), y2=b.y+b.h/2-dy*(tb+0.045/len);
      // Reject routes crossing another node instead of inventing a reroute.
      for (const n of nodes.filter((n)=>![edge.source,edge.target].includes(n.node_id))) {
        const r=positions.get(n.node_id);
        for (let t=0;t<=1;t+=0.01) {
          const x=x1+(x2-x1)*t,y=y1+(y2-y1)*t;
          if(x>r.x&&x<r.x+r.w&&y>r.y&&y<r.y+r.h) incompatible(`edge crosses node ${n.node_id}`);
        }
      }
      objects.push({kind:"line",id:`edge-${i+1}`,bounds:{x:Math.min(x1,x2),y:Math.min(y1,y2),w:Math.abs(x2-x1),h:Math.abs(y2-y1)},tone:"muted",line_width:1.8,
        flipH:x2<x1,flipV:y2<y1,directed:edge.directed,edge:{...edge},relationship_type:plan.relationship_type,source_locator:plan.source_locator});
    });
    nodes.forEach((n)=>{
      const b=positions.get(n.node_id);
      objects.push({kind:"roundRect",id:`node-${n.node_id}`,bounds:b,tone:"primary",line_width:1.4,anchor:false,
        semantic_role:n.node_id===plan.central_id?"center":"equal_member",source_locator:plan.node_locators?.[n.node_id]||plan.source_locator});
      label(`label-${n.node_id}`,n.label,{x:b.x+0.09,y:b.y+0.045,w:b.w-0.18,h:b.h-0.09});
    });
  }
  if(bag(objects.filter((o)=>o.kind==="text").map((o)=>o.text))!==bag(units)) fail(item,plan,"scientific text mutation");
  for(const o of objects){
    const b=o.bounds;
    if(![b.x,b.y,b.w,b.h].every(Number.isFinite)||b.w<0||b.h<0||b.x<frame.x||b.y<frame.y||b.x+b.w>frame.x+frame.w+1e-8||b.y+b.h>frame.y+frame.h+1e-8
        ||b.x<0||b.y<0||b.x+b.w>contract.slide.width_in||b.y+b.h>contract.zones.footer_exclusion.y) incompatible(`object ${o.id} outside approved frame`);
  }
  return {slide_id:item.slide_id,archetype_requested:plan.archetype,archetype_rendered:plan.archetype,generic_fallback:false,revision:"COMPACT_CJK",frame,objects,
    semantics:{relationship_type:plan.relationship_type,emphasis:plan.emphasis,source_locator:plan.source_locator},font_size_pt:size};
}

// SEMANTIC is an explicit, private semantic contract. No classifier, public IR field,
// font shrinking, or alternative dispatch is added. Existing revisions stay above.
function composeSemantic(item, plan, contract) {
  if (!SEMANTIC_FAMILIES.includes(plan.archetype)) fail(item, plan, "unknown SEMANTIC archetype");
  if (item.visual_type !== "takeaway" || item.template_layout_id || item.slide_role === "cover") fail(item, plan, "protected renderer");
  const bad = (reason) => fail(item, plan, reason);
  const capacity = (reason) => fail(item, plan, reason, "COMPOSITION_FRAME_INCOMPATIBLE");
  const frame = item.planned_geometry?.find((o) => o.object_id.endsWith(":structure"))?.bounds;
  if (!frame || frame.w < 10 || frame.h < 2.8) capacity("safe frame too small");
  const size = Number(plan.font_size_pt ?? 20);
  if (!Number.isFinite(size) || size < Math.max(18, contract.minimum_font_pt?.body || 18)) capacity("font below approved minimum");
  const nodes = plan.nodes, edges = plan.edges, groups = plan.groups;
  if (!plan.source_locator || plan.emphasis !== "structural_roles_only" || !Array.isArray(nodes) || !Array.isArray(edges) || !Array.isArray(groups)
      || JSON.stringify(item.diagram_spec?.nodes) !== JSON.stringify(nodes)
      || JSON.stringify(item.diagram_spec?.edges) !== JSON.stringify(edges)) bad("explicit frozen nodes, edges, groups and source required");
  const units = item.takeaways.flatMap((t) => t.split("\n")).filter((t) => t.trim());
  if (bag(nodes.map((n) => n.label)) !== bag(units)) bad("scientific text mutation");
  const ids = nodes.map((n) => n.node_id), groupIds = groups.map((g) => g.group_id);
  if (new Set([...ids, ...groupIds]).size !== ids.length + groupIds.length
      || nodes.some((n) => !n.node_id || !n.role || !n.source_locator || n.emphasis !== "equal_within_role")
      || groups.some((g) => !g.group_id || !g.role || !g.source_locator || !Array.isArray(g.members)
          || new Set(g.members).size !== g.members.length || g.members.some((id) => !ids.includes(id)))) bad("invalid node/group identity or provenance");
  const endpoints = [...ids, ...groupIds];
  if (new Set(edges.map((e) => e.edge_id)).size !== edges.length || edges.some((e) => !e.edge_id || !e.source_locator || !e.relationship_type
      || !endpoints.includes(e.source) || !endpoints.includes(e.target) || e.source === e.target || typeof e.directed !== "boolean")) bad("invalid explicit edge");
  const role = (r) => nodes.filter((n) => n.role === r).map((n) => n.node_id);
  const equal = (a,b) => JSON.stringify(a) === JSON.stringify(b);
  const requireRoles = (counts) => {
    if (Object.values(counts).reduce((a,b) => a+b, 0) !== nodes.length
        || Object.entries(counts).some(([r,n]) => role(r).length !== n)) bad("incorrect semantic roles or capacity");
  };
  const group = (r, members) => {
    const matches = groups.filter((g) => g.role === r);
    if (matches.length !== 1 || !equal(matches[0].members, members)) bad(`explicit ${r} group required`);
    return matches[0].group_id;
  };
  const objects = [], positions = new Map(), containers = new Map();
  const b = (x,y,w,h) => ({x:frame.x+x*frame.w,y:frame.y+y*frame.h,w:w*frame.w,h:h*frame.h});
  const shape = (kind,id,bounds,extra={}) => objects.push({kind,id,bounds,tone:"primary",fill_tone:"white",line_width:1.4,anchor:false,...extra});
  const box = (id, bounds, extra={}) => { positions.set(id,bounds); shape("roundRect",`node-${id}`,bounds,extra); };
  const region = (id,bounds,extra={}) => {containers.set(id,bounds);shape("roundRect",`group-${id}`,bounds,{tone:"line",...extra});};
  const path = (edge, points, dashed=false) => {
    for (let i=1;i<points.length;i++) {
      const [x1,y1]=points[i-1], [x2,y2]=points[i];
      shape("line",`edge-${edge.edge_id}-${i}`,{x:Math.min(x1,x2),y:Math.min(y1,y2),w:Math.abs(x2-x1),h:Math.abs(y2-y1)},
        {tone:"muted",line_width:1.8,flipH:x2<x1,flipV:y2<y1,directed:edge.directed && i===points.length-1,
         ...(i===1?{edge:{...edge}}:{}),logical_edge_id:edge.edge_id,source_locator:edge.source_locator,
         ...(dashed?{dash_type:"dash"}:{})});
    }
  };
  if (plan.archetype === "LAYERED_LOOP_FRAMEWORK") {
    requireRoles({layer:3});
    const layers=role("layer");
    if (groups.length || !equal(plan.layer_order,layers) || edges.length!==3
        || !equal(edges.map((e)=>[e.source,e.target]),[[layers[0],layers[1]],[layers[1],layers[2]],[layers[2],layers[0]]])
        || edges.some((e)=>!e.directed || e.relationship_type!=="teaching_feedback_loop")) bad("three ordered layers and explicit teaching loop required");
    if (plan.external_roles?.length!==1 || plan.external_roles[0].role!=="cross_cutting_principle"
        || plan.external_roles[0].text!==item.single_key_message || !plan.external_roles[0].source_locator) bad("separate source-bound principle required");
    // Equal-width ordered strata; recommendation ranges do not determine area.
    layers.forEach((id,i)=>box(id,b(0.06,0.02+i*0.35,0.78,0.27)));
    edges.slice(0,2).forEach((e)=>{
      const a=positions.get(e.source),z=positions.get(e.target);
      path(e,[[a.x+a.w/2,a.y+a.h+0.045],[z.x+z.w/2,z.y-0.045]]);
    });
    const a=positions.get(layers[2]),z=positions.get(layers[0]),x=frame.x+0.95*frame.w;
    path(edges[2],[[a.x+a.w+0.045,a.y+a.h/2],[x,a.y+a.h/2],[x,z.y+z.h/2],[z.x+z.w+0.045,z.y+z.h/2]]);
  } else if (plan.archetype === "COMPOSITE_DUAL_FRAMEWORK") {
    requireRoles({left_context:1,left_note:1,left_domain:3,right_heading:1,right_core:3,right_plus:1,shared_conclusion:1});
    const left=group("left_module",[...role("left_context"),...role("left_note"),...role("left_domain")]);
    const core=group("right_core",role("right_core"));
    const right=group("right_module",[...role("right_heading"),...role("right_core"),...role("right_plus")]);
    const conclusion=role("shared_conclusion")[0];
    if (groups.length!==3 || edges.length!==2 || !equal(edges.map((e)=>[e.source,e.target]),[[left,conclusion],[right,conclusion]])
        || edges.some((e)=>e.directed || e.relationship_type!=="shared_conclusion")) bad("explicit shared relationship required; no item-to-item edges");
    region(left,b(0.005,0.01,0.475,0.78));
    region(right,b(0.505,0.01,0.49,0.78));
    region(core,b(0.52,0.195,0.46,0.435),{tone:"primary"});
    positions.set(role("left_context")[0],b(0.022,0.045,0.44,0.25));
    role("left_domain").forEach((id,i)=>box(id,b(0.025+i*0.152,0.35,0.137,0.19)));
    positions.set(role("left_note")[0],b(0.025,0.61,0.435,0.13));
    positions.set(role("right_heading")[0],b(0.53,0.045,0.435,0.135));
    role("right_core").forEach((id,i)=>positions.set(id,b(0.535,0.21+i*0.14,0.43,0.13)));
    box(role("right_plus")[0],b(0.52,0.655,0.46,0.13));
    box(conclusion,b(0.005,0.855,0.99,0.14));
    edges.forEach((e)=>{
      const a=containers.get(e.source),z=positions.get(conclusion);
      path(e,[[a.x+a.w/2,a.y+a.h+0.015],[a.x+a.w/2,z.y-0.015]]);
    });
  } else {
    requireRoles({conditional_context:1,support_context:1,persistent_band:1,future_exploration:1,current_stage:4});
    const stages=role("current_stage"),future=role("future_exploration")[0];
    const current=group("current_stages",stages);
    if (groups.length!==1 || !equal(plan.stage_order,stages) || !equal(plan.band_covers,stages)
        || edges.length!==4 || !equal(edges.slice(0,3).map((e)=>[e.source,e.target]),stages.slice(1).map((id,i)=>[stages[i],id]))) bad("explicit stage order and cross-stage coverage required");
    if (edges.slice(0,3).some((e)=>!e.directed || e.relationship_type!=="stage_sequence")
        || edges[3].source!==current || edges[3].target!==future || edges[3].directed
        || edges[3].relationship_type!=="future_exploration_extension") bad("future/current roles must be unambiguous; future is not a routine stage");
    positions.set(role("conditional_context")[0],b(0.015,0.00,0.97,0.125));
    positions.set(role("support_context")[0],b(0.015,0.14,0.97,0.24));
    stages.forEach((id,i)=>box(id,b(0.005+i*0.195,0.43,0.17,0.19)));
    box(role("persistent_band")[0],b(0.005,0.72,0.755,0.27));
    box(future,b(0.815,0.43,0.18,0.56),{dash_type:"dash"});
    containers.set(current,b(0.005,0.43,0.755,0.19));
    edges.slice(0,3).forEach((e)=>{
      const a=positions.get(e.source),z=positions.get(e.target);
      path(e,[[a.x+a.w+0.025,a.y+a.h/2],[z.x-0.025,z.y+z.h/2]]);
    });
    const a=containers.get(current),z=positions.get(future);
    path(edges[3],[[a.x+a.w+0.025,a.y+a.h/2],[z.x-0.025,a.y+a.h/2]],true);
  }
  for (const n of nodes) {
    const p=positions.get(n.node_id);
    const bounds={x:p.x+0.07,y:p.y+0.012,w:p.w-0.14,h:p.h-0.024};
    const lines=n.label.split("\n").reduce((sum,line)=>sum+Math.ceil([...line].reduce((x,c)=>x+(/[^\x00-\x7F]/u.test(c)?1:0.6),0)/(bounds.w*72/size)),0);
    if (lines*size*1.2>bounds.h*72+0.1) capacity(`text ${n.node_id} requires ${lines} lines at ${size} pt; ${bounds.h.toFixed(3)} inch box`);
    objects.push({kind:"text",id:`label-${n.node_id}`,text:n.label,bounds,font_size_pt:size,bold:false,hero:false,anchor:false,
      align:["left_context","left_note","right_heading","right_core","right_plus","conditional_context","support_context"].includes(n.role)?"left":"center",
      semantic_role:n.role,emphasis:n.emphasis,source_locator:n.source_locator});
  }
  for (const o of objects) {
    const p=o.bounds;
    if (![p.x,p.y,p.w,p.h].every(Number.isFinite) || p.w<0 || p.h<0 || p.x<frame.x || p.y<frame.y
        || p.x+p.w>frame.x+frame.w+1e-8 || p.y+p.h>frame.y+frame.h+1e-8
        || p.x<0 || p.y<0 || p.x+p.w>contract.slide.width_in || p.y+p.h>contract.zones.footer_exclusion.y) capacity(`object ${o.id} outside approved frame`);
  }
  return {slide_id:item.slide_id,archetype_requested:plan.archetype,archetype_rendered:plan.archetype,revision:"SEMANTIC",
    generic_fallback:false,frame,objects,font_size_pt:size,semantics:{nodes,groups,edges,external_roles:plan.external_roles||[],
      stage_order:plan.stage_order||[],band_covers:plan.band_covers||[],layer_order:plan.layer_order||[],
      relationship_type:plan.relationship_type,emphasis:plan.emphasis,source_locator:plan.source_locator}};
}

export function prepareComposition(spec) {
  if (spec.composition?.id !== COMPOSITION_ID || spec.composition?.enabled !== true) throw new Error("COMPOSITION_NOT_ENABLED");
  const plan = spec.visual_execution_plan;
  if (!plan || typeof plan !== "object" || Array.isArray(plan) || !Object.keys(plan).length) throw new Error("MISSING_COMPOSITION_PLAN");
  const compositions = new Map();
  for (const [id, assignment] of Object.entries(plan)) {
    const item = spec.slides.find((s) => s.slide_id === id);
    if (!item) throw new Error(`Unknown composition slide: ${id}`);
    compositions.set(id, compose(item, assignment, spec.layout_contract));
  }
  const rendered = [];
  return {
    has: (id) => compositions.has(id),
    render(slide, item, { pptx, C, font }) {
      const start = performance.now(), c = compositions.get(item.slide_id);
      for (const o of c.objects) {
        const common = { ...o.bounds, objectName: `composition:${item.slide_id}:${o.id}` };
        if (o.kind === "text") {
          slide.addText(o.text, { ...common, fontFace: font, fontSize: o.font_size_pt,
            bold: o.bold ?? (o.hero || c.archetype_rendered !== "EDITORIAL_TAKEAWAY"), color: o.hero ? C.primary : C.ink,
            margin: 0, valign: "mid", align: o.align || (c.archetype_rendered === "EDITORIAL_TAKEAWAY" ? "left" : "center") });
        } else {
          slide.addShape(pptx.ShapeType[o.kind], { ...common, flipH: o.flipH, flipV: o.flipV,
            fill: { color: o.fill_tone ? C[o.fill_tone] : o.kind === "rect" ? C.primary : C.white },
            line: { color: C[o.tone], width: o.line_width ?? (o.anchor ? 2.2 : 1.4), ...(o.directed ? { endArrowType: "triangle" } : {}),
              ...(o.dash_type ? {dashType:o.dash_type} : {}) } });
        }
      }
      rendered.push({ ...c, runtime_seconds: (performance.now() - start) / 1000,
        object_count: c.objects.length, text_count: c.objects.filter((o) => o.kind === "text").length,
        visual_anchor_count: c.objects.filter((o) => o.anchor).length,
        primitive_distribution: c.objects.reduce((a, o) => ({ ...a, [o.kind]: (a[o.kind] || 0) + 1 }), {}) });
    },
    report: () => ({ composition_id: COMPOSITION_ID, targeted_slide_count: compositions.size,
      rendered_slide_count: rendered.length, generic_fallback_count: 0, slides: rendered }),
  };
}
