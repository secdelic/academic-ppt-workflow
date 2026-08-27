import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
function loadPptxGenJS() {
  try { return require("pptxgenjs"); } catch (first) {
    const roots = [process.env.PPT_NODE_MODULES, process.env.NODE_PATH].filter(Boolean);
    for (const root of roots) {
      try { return require(path.join(root, "pptxgenjs")); } catch { /* continue */ }
    }
    throw first;
  }
}
const pptxgen = loadPptxGenJS();
const SHAPES = new pptxgen().ShapeType;
const repo = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(?:[A-Za-z]:)/, m => m.slice(1))), "..");
const outRoot = path.resolve(process.argv[2] || path.join(repo, "visual_workspaces"));

const roles = ["cover","section_divider","problem_gap_objective","method_framework","primary_result","secondary_result","hero_result","limitation","conclusion","appendix"];
const configs = {
  clinical_mdt: { label:"Clinical MDT", profile:"clinical_mdt", motif:"system map + decision cards", primary:"17324D", accent:"178A8A", compare:"E76F51", bg:"F6F9FB" },
  clinical_research_protocol: { label:"Clinical Research Protocol", profile:"clinical_protocol", motif:"planned pathway + phase bands", primary:"1D3B46", accent:"3A8D8A", compare:"6C8EBF", bg:"F5F9FA" },
  scientific_results: { label:"Scientific Results", profile:"scientific_results", motif:"effect + uncertainty", primary:"243B53", accent:"2A9D8F", compare:"D95F59", bg:"F7F9FC" },
  thesis_defense: { label:"Thesis Defense", profile:"thesis_defense", motif:"question → evidence → contribution", primary:"253858", accent:"6D78B7", compare:"D08C60", bg:"F8F7FB" },
  bioinformatics: { label:"Bioinformatics", profile:"bioinformatics", motif:"cell identity + molecular state", primary:"213A5C", accent:"2A9D8F", compare:"D95F59", bg:"F6F9FB" },
  evidence_synthesis: { label:"Evidence Synthesis", profile:"evidence_synthesis", motif:"aggregation + certainty", primary:"263B57", accent:"6B7F9E", compare:"C74B50", bg:"F7F8FA" },
};

const q = value => JSON.stringify(value);
const yamlList = values => values.map(v => `  - ${v}`).join("\n");
const write = (file, text) => fs.writeFileSync(file, text.endsWith("\n") ? text : text + "\n", "utf8");
const variants = {
  problem_gap_objective:["asymmetric_flow","central_gap","question_to_objective"],
  method_framework:["horizontal_branch","vertical_stages","framework_with_guardrails"],
  primary_result:["chart_plus_insight","hero_left","full_width"],
  secondary_result:["small_multiples","split_comparison","delta_focus"],
  hero_result:["hero_number_plus_chart","hero_chart_with_boundary"],
  conclusion:["numbered_takeaways","evidence_boundary_summary","hero_number_plus_takeaways"],
};

function addText(slide, text, x,y,w,h, opts={}) { slide.addText(text,{x,y,w,h,fontFace:"Microsoft YaHei",fontSize:opts.size||18,color:opts.color||"17324D",bold:!!opts.bold,margin:0,breakLine:false,fit:"shrink",valign:opts.valign||"mid",align:opts.align||"left"}); }
function addTitle(slide, text, c, dark=false) { addText(slide,text,.72,.35,11.8,.68,{size:34,bold:true,color:dark?"FFFFFF":c.primary}); }
function footer(slide,c,i) { slide.addShape(SHAPES.line,{x:.72,y:7.08,w:11.85,h:0,line:{color:darken(c.accent),width:1}}); addText(slide,"SYNTHETIC WORKSPACE PREVIEW",.72,7.14,4.2,.2,{size:9,color:"627D98"}); addText(slide,String(i),12.2,7.13,.35,.2,{size:9,color:"627D98",align:"right"}); }
function darken(hex){return hex;}
function baseSlide(pptx,c,title,i,dark=false){const s=pptx.addSlide();s.background={color:dark?c.primary:c.bg};if(title)addTitle(s,title,c,dark);footer(s,c,i);return s;}
function pill(slide,text,x,y,w,c,fill){slide.addShape(SHAPES.roundRect,{x,y,w,h:.42,rectRadius:.06,fill:{color:fill},line:{color:fill}});addText(slide,text,x+.12,y+.04,w-.24,.3,{size:14,bold:true,color:c});}
function syntheticChart(slide,c,x,y,w,h,mode="bars"){
  slide.addShape(SHAPES.rect,{x,y,w,h,fill:{color:"FFFFFF",transparency:0},line:{color:"D9E2EC",width:1}});
  slide.addShape(SHAPES.line,{x:x+.55,y:y+h-.55,w:w-1,h:0,line:{color:"9FB3C8",width:1}});
  slide.addShape(SHAPES.line,{x:x+.55,y:y+.45,w:0,h:h-1,line:{color:"9FB3C8",width:1}});
  const vals=mode==="facet"?[.42,.66,.35,.72]:[.28,.48,.7,.56,.8];
  vals.forEach((v,idx)=>{const bw=(w-1.4)/vals.length;slide.addShape(mode==="dots"?SHAPES.ellipse:SHAPES.rect,{x:x+.75+idx*bw,y:mode==="dots"?y+h-.58-v*(h-1.35):y+h-.55-v*(h-1.2),w:mode==="dots"?.16:bw*.55,h:mode==="dots"?.16:v*(h-1.2),fill:{color:idx===vals.length-1?c.compare:c.accent},line:{color:idx===vals.length-1?c.compare:c.accent}})});
}
function notes(slide,role){ if (typeof slide.addNotes === "function") slide.addNotes(`[Sources]\n- SYNTHETIC-WORKSPACE-DATA | ${role} | no real claim or patient data`); }

async function buildWorkspace(id,c){
  const dir=path.join(outRoot,id); fs.mkdirSync(dir,{recursive:true});
  write(path.join(dir,"workspace.yaml"),`schema_version: visual-workspace/1\nworkspace_id: ${id}\nselection:\n  domain_profile: ${c.profile}\n  presentation_types: [academic, clinical, research]\n  project_name_routing: prohibited\n  slide_number_routing: prohibited\nrequired_roles:\n${yamlList(roles)}\n`);
  write(path.join(dir,"art_direction.yaml"),`workspace_id: ${id}\nvisual_tone: rigorous, calm, decision-oriented\nsignature_motif: ${q(c.motif)}\npalette:\n  primary: ${c.primary}\n  accent: ${c.accent}\n  comparison: ${c.compare}\n  canvas: ${c.bg}\ntypography:\n  zh: Microsoft YaHei\n  latin: Aptos\nprohibited_styles: [glow, decorative_gradient, unrelated_photo, entertainment_illustration]\n`);
  write(path.join(dir,"master_roles.yaml"),`roles:\n${roles.map(r=>`  ${r}:\n    master_family: ${r==="cover"?"cover":r==="section_divider"?"section_divider":r==="appendix"?"appendix_audit":"chart_led"}\n    source_footer: ${r==="cover"?"optional":"required"}`).join("\n")}\n`);
  write(path.join(dir,"layout_variants.yaml"),`variants:\n${Object.entries(variants).map(([k,v])=>`  ${k}: [${v.join(", ")}]`).join("\n")}\nmax_same_variant_run: 2\n`);
  write(path.join(dir,"chart_styles.yaml"),`native_chart:\n  font: Aptos\n  axis_color: 627D98\n  series: [${c.accent}, ${c.compare}, 9FB3C8]\neditable_shapes:\n  reference_line: 9FB3C8\n  uncertainty_fill: FFF4E5\n  direct_label: true\n`);
  write(path.join(dir,"annotation_rules.yaml"),`maximum_primary: 1\nmaximum_secondary: 1\nallowed: [direct_label, key_value_callout, delta_annotation, threshold_annotation, confidence_annotation, caution_annotation, evidence_boundary_annotation]\n`);
  write(path.join(dir,"example_storyboard.yaml"),`workspace_id: ${id}\nslides:\n${roles.map((r,i)=>`  - slide_id: SYN-${id.toUpperCase()}-${String(i+1).padStart(2,"0")}\n    role: ${r}\n    layout_variant: ${(variants[r]||["standard"])[i%(variants[r]||["standard"]).length]}\n    synthetic_only: true`).join("\n")}\n`);
  write(path.join(dir,"human_review.csv"),"workspace_id,reviewer,review_date,scientific_clarity,visual_hierarchy,composition,chart_professionalism,consistency,readability,total_score,status,comments\n"+`${id},,,,,,,,,,PENDING,\n`);

  const pptx=new pptxgen(); pptx.layout="LAYOUT_WIDE"; pptx.author="Academic PPT Workflow";pptx.subject=`Synthetic ${c.label} visual workspace`;pptx.title=`${c.label} Workspace`;pptx.company="Local private workflow";pptx.lang="zh-CN";pptx.theme={headFontFace:"Microsoft YaHei",bodyFontFace:"Microsoft YaHei",lang:"zh-CN"};
  let s=baseSlide(pptx,c,"",1,true); addText(s,c.label,.82,1.45,8.6,.9,{size:48,bold:true,color:"FFFFFF"});addText(s,"Canonical Visual Workspace · Synthetic Preview",.84,2.55,7.6,.4,{size:22,color:"D9EAF0"});s.addShape(SHAPES.arc,{x:9.7,y:1.25,w:2.25,h:2.25,adjustPoint:.4,rotate:20,line:{color:c.accent,width:8,transparency:10},fill:{color:c.primary,transparency:100}});notes(s,"cover");
  s=baseSlide(pptx,c,"证据链决定页面节奏",2,true);addText(s,"问题",1.0,2.2,2.2,.6,{size:30,bold:true,color:"FFFFFF"});addText(s,"→",3.35,2.2,.6,.6,{size:30,color:c.accent});addText(s,"证据",4.15,2.2,2.2,.6,{size:30,bold:true,color:"FFFFFF"});addText(s,"→",6.5,2.2,.6,.6,{size:30,color:c.accent});addText(s,"决策",7.3,2.2,2.2,.6,{size:30,bold:true,color:"FFFFFF"});addText(s,"章节页只承担过渡，不堆叠细节",1,4.3,8.8,.45,{size:20,color:"D9EAF0"});notes(s,"section_divider");
  s=baseSlide(pptx,c,"问题、缺口与目标必须形成一条逻辑线",3);["已知","缺口","目标"].forEach((t,i)=>{const x=.9+i*4.0;s.addShape(SHAPES.roundRect,{x,y:1.75,w:3.2,h:3.65,rectRadius:.06,fill:{color:i===1?"FFF4E5":"FFFFFF"},line:{color:i===1?c.compare:"D9E2EC",width:1.2}});pill(s,t,x+.2,2.0,1.0,i===1?"FFFFFF":c.primary,i===1?c.compare:"EAF2F8");addText(s,["观察到稳定但异质的信号","关键决策边界仍未解决","用预先定义的证据回答问题"][i],x+.25,2.75,2.7,1.3,{size:21,bold:true,color:c.primary})});notes(s,"problem_gap_objective");
  s=baseSlide(pptx,c,"方法框架把输入、分析与边界分开",4);const stages=["输入","质量门槛","分析","解释边界"];stages.forEach((t,i)=>{const x=.9+i*3.02;s.addShape(SHAPES.chevron,{x,y:2.25,w:2.55,h:1.55,fill:{color:i%2?c.accent:c.primary},line:{color:"FFFFFF",transparency:100}});addText(s,t,x+.28,2.7,1.85,.45,{size:20,bold:true,color:"FFFFFF",align:"center"})});addText(s,"合成示例不代表真实统计结果",.95,4.55,6.5,.42,{size:18,color:"627D98"});notes(s,"method_framework");
  s=baseSlide(pptx,c,"主要结果需要同时显示效应与比较",5);syntheticChart(s,c,.85,1.55,8.1,4.75);s.addShape(SHAPES.roundRect,{x:9.35,y:2.0,w:2.95,h:2.2,rectRadius:.06,fill:{color:c.primary},line:{color:c.primary}});addText(s,"+12.4",9.65,2.35,2.35,.65,{size:40,bold:true,color:"FFFFFF",align:"center"});addText(s,"合成变化值",9.65,3.15,2.35,.35,{size:17,color:"D9EAF0",align:"center"});notes(s,"primary_result");
  s=baseSlide(pptx,c,"分层结果用小多图保持实体范围",6);[0,1,2].forEach(i=>{addText(s,`分层 ${String.fromCharCode(65+i)}`,1+i*4.1,1.65,3.2,.35,{size:20,bold:true,color:c.primary});syntheticChart(s,c,1+i*4.1,2.15,3.25,3.25,"facet")});notes(s,"secondary_result");
  s=baseSlide(pptx,c,"核心结果只保留一个视觉中心",7,true);addText(s,"0.72",.95,1.75,4.3,1.15,{size:66,bold:true,color:"FFFFFF"});addText(s,"合成效应估计",1.02,3.0,3.8,.38,{size:21,color:"D9EAF0"});syntheticChart(s,c,5.5,1.45,6.5,4.8,"dots");addText(s,"95% 区间跨越预设阈值",.98,4.25,3.95,.55,{size:22,bold:true,color:c.accent});notes(s,"hero_result");
  s=baseSlide(pptx,c,"局限性决定结论可以走多远",8);const limits=["测量误差","残余混杂","外部有效性"];limits.forEach((t,i)=>{const y=1.7+i*1.35;s.addShape(SHAPES.roundRect,{x:1,y,w:10.8,h:.95,rectRadius:.05,fill:{color:i===1?"FFF4E5":"FFFFFF"},line:{color:i===1?c.compare:"D9E2EC"}});addText(s,String(i+1).padStart(2,"0"),1.3,y+.2,.55,.4,{size:20,bold:true,color:c.compare});addText(s,t,2.05,y+.18,2.5,.42,{size:22,bold:true,color:c.primary});addText(s,"需要在正式结论中保持限定语",5.0,y+.2,5.6,.36,{size:18,color:"627D98"})});notes(s,"limitation");
  s=baseSlide(pptx,c,"结论必须把证据、边界与行动连接起来",9);["证据","边界","下一步"].forEach((t,i)=>{const x=1+i*4.05;addText(s,t,x,1.75,2.7,.42,{size:24,bold:true,color:i===2?c.compare:c.primary});s.addShape(SHAPES.line,{x,y:2.35,w:2.75,h:0,line:{color:i===2?c.compare:c.accent,width:4}});addText(s,["合成信号方向一致","不确定性仍需保留","进入人工科学复核"][i],x,2.7,2.8,1.25,{size:21,bold:true,color:c.primary})});notes(s,"conclusion");
  s=baseSlide(pptx,c,"附录：来源、规则与审计对象",10);const rows=[["对象","状态","来源"],["SlideSpec","已验证","SYNTHETIC"],["VisualSpec","已验证","SYNTHETIC"],["人工评分","待完成","human_review.csv"]];rows.forEach((r,i)=>r.forEach((t,j)=>{const x=.9+[0,4.2,7.9][j];const w=[4,3.5,3.9][j];s.addShape(SHAPES.rect,{x,y:1.65+i*.82,w,h:.7,fill:{color:i===0?c.primary:(i%2?"FFFFFF":"EDF2F7")},line:{color:"D9E2EC"}});addText(s,t,x+.15,1.82+i*.82,w-.3,.32,{size:i===0?17:16,bold:i===0,color:i===0?"FFFFFF":c.primary})}));notes(s,"appendix");
  await pptx.writeFile({fileName:path.join(dir,"synthetic_preview.pptx")});
}

fs.mkdirSync(outRoot,{recursive:true});
for (const [id,c] of Object.entries(configs)) await buildWorkspace(id,c);
console.log(JSON.stringify({status:"PASS",workspace_count:Object.keys(configs).length,slide_count:Object.keys(configs).length*10,output:outRoot}));
