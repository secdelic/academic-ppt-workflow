import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const pptxgen = require("pptxgenjs");
const imageSizeModule = require("image-size");
const imageSize = imageSizeModule.imageSize || imageSizeModule.default || imageSizeModule;

const [specPath, outputPath] = process.argv.slice(2);
const deck = JSON.parse(fs.readFileSync(specPath, "utf8"));
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Academic PPT Workflow v2.3";
pptx.subject = "Synthetic evidence-to-visual narrative remediation";
pptx.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos", lang: "en-US" };
pptx.defineSlideMaster({
  title: "V23",
  background: { color: "F6F9FB" },
  objects: [
    { rect: { x: 0, y: 7.16, w: 13.333, h: 0.34, fill: { color: "17324D" }, line: { color: "17324D" } } },
    { line: { x: 0.76, y: 0.48, w: 0.42, h: 0, line: { color: "2A9D8F", width: 4 } } }
  ],
  slideNumber: { x: 12.45, y: 7.18, w: 0.42, h: 0.16, color: "FFFFFF", fontFace: "Arial", fontSize: 9, align: "right" }
});
const S = pptx.ShapeType;
const C = { navy:"17324D", teal:"178A8A", green:"2A9D8F", blue:"4E79A7", coral:"E76F51", amber:"E9A23B", ink:"243B53", muted:"627D98", pale:"EAF1F5", mint:"DDF3F0", white:"FFFFFF", red:"B23A48" };

function title(slide, text, role) {
  slide.addText(text, { x:0.8,y:0.39,w:11.8,h:0.62,fontFace:"Aptos Display",fontSize:35,bold:true,color:C.navy,margin:0,fit:"shrink" });
  slide.addText(role.toUpperCase(), { x:0.82,y:1.08,w:2.4,h:0.2,fontFace:"Arial",fontSize:10,bold:true,color:C.teal,charSpacing:1.3,margin:0 });
}
function footer(slide, contracts) {
  const labels = [...new Set((contracts||[]).map(c=>path.basename(c.source_file||"")))].slice(0,2);
  if(labels.length) slide.addText(`Source: ${labels.join(" · ")}`, { x:0.82,y:7.18,w:10.6,h:0.15,fontSize:9.5,color:C.white,margin:0 });
}
function notes(slide, contracts, conflicts, unresolved) {
  const lines=["[Sources]"];
  for(const c of contracts||[]) lines.push(`${c.claim_id} | ${c.source_id} | ${c.source_file} | ${c.source_location} | ${c.wording_boundary} | conflict=${c.conflict_status} | unresolved=${c.unresolved_status}`);
  lines.push("", "[Conflicts]", ...(conflicts||[]).map(x=>`${x.conflict_id} | ${x.conflicting_value} -> ${x.canonical_value}`));
  lines.push("", "[Unresolved]", ...(unresolved||[]).map(x=>`${x.marker} | ${x.status}`));
  slide.addNotes(lines.join("\n"));
}
function rect(slide,x,y,w,h,fill="FFFFFF",line="C6D5DF") {
  slide.addShape(S.roundRect,{x,y,w,h,rectRadius:0.05,fill:{color:fill},line:{color:line,width:1.1}});
}
function addText(slide,text,x,y,w,h,size=18,opts={}) {
  slide.addText(String(text),{x,y,w,h,fontFace:opts.fontFace||"Aptos",fontSize:size,bold:!!opts.bold,color:opts.color||C.ink,margin:opts.margin??0.04,align:opts.align||"left",valign:opts.valign||"mid",fit:"shrink"});
}
function imageContain(slide,file,x,y,w,h){
  const d=imageSize(fs.readFileSync(file)); const scale=Math.min(w/d.width,h/d.height);
  const iw=d.width*scale,ih=d.height*scale;
  slide.addImage({path:file,x:x+(w-iw)/2,y:y+(h-ih)/2,w:iw,h:ih});
}
function claimMap(){ return Object.fromEntries(deck.claims.map(c=>[c.claim_id,c])); }
const claimsById=claimMap();
function contractsFor(plan){ return (plan.claim_ids||[]).map(id=>claimsById[id]).filter(Boolean); }
function findFigure(pattern){ const files=fs.readdirSync(deck.figures_dir); const f=files.find(n=>n.toLowerCase().includes(pattern)); return f?path.join(deck.figures_dir,f):null; }

function flow(slide,labels){
  const items=labels.slice(0,6), gap=0.34, w=(11.5-gap*(items.length-1))/items.length, y=3.05;
  items.forEach((label,i)=>{
    if(i<items.length-1) slide.addShape(S.line,{x:0.92+i*(w+gap)+w,y:y+0.42,w:gap,h:0,line:{color:C.muted,width:1.8,endArrowType:"triangle"}});
  });
  items.forEach((label,i)=>{const x=0.92+i*(w+gap);rect(slide,x,y,w,0.86,i%2?C.mint:C.white,C.teal);addText(slide,label,x+0.08,y+0.12,w-0.16,0.56,15,{bold:true,align:"center"});});
}
function love(slide,rows){
  const n=Math.min(rows.length,8), x0=5.1,w=6.6,max=.35;
  slide.addShape(S.line,{x:x0+0.1/max*w,y:1.62,w:0,h:4.75,line:{color:C.coral,width:1,dash:"dash"}});
  rows.slice(0,n).forEach((r,i)=>{const y=1.76+i*.58;addText(slide,r.covariate,0.95,y,3.55,.3,15);const a=Number(r.absolute_smd_unweighted),b=Number(r.absolute_smd_weighted);slide.addShape(S.line,{x:x0+Math.min(a,b)/max*w,y:y+.13,w:Math.abs(a-b)/max*w,h:0,line:{color:"AFC4D2",width:1.2}});slide.addShape(S.ellipse,{x:x0+a/max*w-.07,y:y+.06,w:.14,h:.14,fill:{color:C.coral},line:{color:C.coral}});slide.addShape(S.ellipse,{x:x0+b/max*w-.07,y:y+.06,w:.14,h:.14,fill:{color:C.teal},line:{color:C.teal}});});
  addText(slide,"Dashed line: |SMD| = 0.10",5.1,6.45,3,.22,11,{color:C.muted});
}
function forest(slide,rows,labelKey,estimateKey,lowKey,highKey,ref=1){
  const n=Math.min(rows.length,14), x0=5.2,w=5.9;
  const vals=rows.flatMap(r=>[Number(r[lowKey]),Number(r[highKey])]).filter(Number.isFinite);let lo=Math.min(...vals,ref),hi=Math.max(...vals,ref);const pad=(hi-lo)*.08;lo-=pad;hi+=pad;const mx=v=>x0+(v-lo)/(hi-lo)*w;
  slide.addShape(S.line,{x:mx(ref),y:1.46,w:0,h:5.15,line:{color:"91A8B8",width:1,dash:"dash"}});
  rows.slice(0,n).forEach((r,i)=>{const y=1.52+i*(4.85/Math.max(n,1));addText(slide,r[labelKey],0.9,y,3.9,.25,n>10?11:14);const e=Number(r[estimateKey]),l=Number(r[lowKey]),h=Number(r[highKey]);slide.addShape(S.line,{x:mx(l),y:y+.11,w:Math.max(.02,mx(h)-mx(l)),h:0,line:{color:C.blue,width:1.8}});slide.addShape(S.ellipse,{x:mx(e)-.055,y:y+.055,w:.11,h:.11,fill:{color:i===n-1?C.coral:C.teal},line:{color:C.teal}});});
}
function bars(slide,rows,labelKey,valueKeys){
  const n=Math.min(rows.length,10), max=Math.max(1,...rows.flatMap(r=>valueKeys.map(k=>Math.abs(Number(r[k])))).filter(Number.isFinite));
  rows.slice(0,n).forEach((r,i)=>{const y=1.55+i*.5;addText(slide,r[labelKey],.9,y,3.7,.28,13);valueKeys.forEach((k,j)=>{const v=Number(r[k]);const x=4.85;slide.addShape(S.rect,{x,y:y+.05+j*.13,w:Math.max(.03,Math.abs(v)/max*6.5),h:.11,fill:{color:j?C.coral:C.teal},line:{color:"FFFFFF",transparency:100}});});});
}
function cards(slide,items){
  items.slice(0,6).forEach((it,i)=>{const col=i%3,row=Math.floor(i/3),x=.9+col*4.1,y=1.55+row*2.25;rect(slide,x,y,3.75,1.75,row?C.white:C.pale);addText(slide,it.title,x+.22,y+.22,3.3,.48,22,{bold:true,color:C.navy});addText(slide,it.body,x+.22,y+.82,3.3,.62,15,{color:C.muted});});
}
function reviewSlide(slide){
  title(slide,`${deck.conflicts.length} conflicts and ${deck.unresolved.length} items need review`,"appendix");
  let y=1.48;
  deck.conflicts.slice(0,4).forEach(c=>{rect(slide,.85,y,6.0,.65,C.white);addText(slide,`${c.conflicting_value} → ${c.canonical_value}`,1.05,y+.13,5.5,.35,14,{bold:true});y+=.78;});
  y=1.48;
  deck.unresolved.slice(0,4).forEach(u=>{rect(slide,7.0,y,5.4,.65,"FFF5E8",C.amber);addText(slide,u.marker,7.2,y+.13,4.9,.35,14,{bold:true,color:C.red});y+=.78;});
  footer(slide,[]);
  notes(slide,[],deck.conflicts,deck.unresolved);
}

function renderTemplate(slide,plan){
  const t=plan.visual_template_id,spec=deck.spec,contracts=contractsFor(plan);
  title(slide,plan.key_message,plan.slide_role);
  if(t==="target_trial_specification_table") cards(slide,[{title:"Eligibility",body:"Synthetic ICU cohort with prespecified criteria"},{title:"Strategies",body:"Transfuse within 24 h vs no transfusion"},{title:"Outcome",body:"28-day mortality risk"},{title:"Estimands",body:"Risk, RD, RR, and weighted HR"},{title:"Grace period",body:"24 h canonical specification"},{title:"Boundary",body:"Assumption-dependent target-trial estimates"}]);
  else if(t==="clone_censor_weight_flow") flow(slide,["Eligible cohort","Clone strategies","Censor deviations","Estimate weights","Weighted outcome"]);
  else if(t==="covariate_balance_love_plot") love(slide,spec.balance);
  else if(t==="estimand_summary_panel"){
    const eff=spec.effects;cards(slide,[{title:`${(Number(eff[0].estimate)*100).toFixed(1)}%`,body:"28-day risk · transfuse"},{title:`${(Number(eff[1].estimate)*100).toFixed(1)}%`,body:"28-day risk · no transfusion"},{title:`${(Number(eff[2].estimate)*100).toFixed(1)} pp`,body:"Risk difference"},{title:`RR ${Number(eff[3].estimate).toFixed(2)}`,body:`95% CI ${eff[3].ci_low}–${eff[3].ci_high}`},{title:`HR ${Number(eff[4].estimate).toFixed(2)}`,body:`Supportive · 95% CI ${eff[4].ci_low}–${eff[4].ci_high}`}]);
  } else if(t==="sensitivity_forest") forest(slide,spec.sensitivity,"analysis","risk_difference","ci_low","ci_high",0);
  else if(t==="assumptions_failure_modes_matrix") cards(slide,[{title:"Exchangeability",body:"Residual confounding may remain"},{title:"Positivity",body:"Extreme weights signal instability"},{title:"Consistency",body:"Strategies require precise definitions"},{title:"No recommendation",body:"Synthetic estimates are not clinical guidance"}]);
  else if(t==="sc_qc_funnel"){const q=spec.qc;flow(slide,q.map(r=>`${r.stage}\n${r.remaining_cells}`));}
  else if(t==="umap_with_callouts"){const f=findFigure("umap");if(f)imageContain(slide,f,1.0,1.42,8.7,5.35);rect(slide,9.85,1.75,2.55,2.0,C.mint,C.teal);addText(slide,"Cell-type labels",10.1,2.0,2.05,.45,20,{bold:true});addText(slide,"Embedding proximity is descriptive—not lineage or mechanism.",10.1,2.65,2.05,.78,15);}
  else if(t==="composition_grouped_bar_or_delta_plot"){const rows=spec.composition;const types=[...new Set(rows.map(r=>r.cell_type))];const comp=types.map(type=>{const a=rows.find(r=>r.cell_type===type&&r.group==="Control"),b=rows.find(r=>r.cell_type===type&&r.group==="SIMD");return{cell_type:type,Control:a?.percentage||0,SIMD:b?.percentage||0};});bars(slide,comp,"cell_type",["Control","SIMD"]);addText(slide,"Teal: Control   Coral: SIMD",9.1,6.5,3,.22,12,{color:C.muted});}
  else if(t==="pseudobulk_volcano_with_top_gene_labels"){const rows=spec.deg.slice(0,80),effectKey=Object.keys(rows[0]||{}).find(k=>k.toLowerCase().startsWith("log2_fold_change")),fdrKey=Object.keys(rows[0]||{}).find(k=>k.toLowerCase()==="fdr"),xs=rows.map(r=>Number(r[effectKey])),ys=rows.map(r=>-Math.log10(Math.max(Number(r[fdrKey]),1e-20)));const min=-3,max=3;rows.forEach((r,i)=>{const x=1.2+(xs[i]-min)/(max-min)*9.8,y=6.25-Math.min(18,ys[i])/18*4.6;slide.addShape(S.ellipse,{x:x-.035,y:y-.035,w:.07,h:.07,fill:{color:xs[i]>0?C.teal:C.coral},line:{color:"FFFFFF",transparency:100}});});const top=[...rows].sort((a,b)=>Number(a[fdrKey])-Number(b[fdrKey])).slice(0,5);cards(slide,top.map(r=>({title:r.gene,body:`log2FC ${r[effectKey]} · FDR ${r[fdrKey]}`})));}
  else if(t==="pathway_enrichment_bar"){const k=Object.keys(spec.pathway[0]||{}).find(x=>x.toLowerCase().includes("enrichment_score"));bars(slide,spec.pathway,"pathway",[k]);}
  else if(t==="cell_communication_network"){const top=[...spec.chat].sort((a,b)=>Number(b.communication_probability)-Number(a.communication_probability)).slice(0,4);flow(slide,top.map(r=>`${r.sender} → ${r.receiver}\n${r.ligand_receptor}`));addText(slide,"Predicted / prioritized interactions only",4.0,5.35,5.4,.36,20,{bold:true,align:"center",color:C.red});}
  else if(t==="evidence_ladder_computation_to_validation") flow(slide,["Synthetic cells","Pseudobulk signal","Pathway enrichment","Predicted interaction","External validation required"]);
  else if(t==="prisma_flow_native") flow(slide,["Records identified","Screened","Full text","14 included studies"]);
  else if(t==="full_forest_plot_with_pooled_effect") forest(slide,spec.studies,"study_label","odds_ratio","ci_low","ci_high",1);
  else if(t==="heterogeneity_prediction_interval_panel"){const r=spec.pooled.find(x=>x.estimand.toLowerCase().includes("random")),p=spec.pooled.find(x=>x.estimand.toLowerCase().includes("prediction"));cards(slide,[{title:`OR ${r.estimate}`,body:`95% CI ${r.ci_low}–${r.ci_high}`},{title:`I² ${r.I2_percent}%`,body:"Moderate inconsistency"},{title:`${p.ci_low}–${p.ci_high}`,body:"Prediction interval · future settings"}]);}
  else if(t==="funnel_plot_with_caution_label"){const f=findFigure("funnel");if(f)imageContain(slide,f,1.0,1.48,8.5,5.25);rect(slide,9.8,1.85,2.55,2.1,"FFF5E8",C.amber);addText(slide,"Caution",10.05,2.1,2.0,.38,22,{bold:true,color:C.red});addText(slide,"Visual asymmetry cannot prove or exclude publication bias.",10.05,2.75,2.0,.85,15);}
  else if(t==="risk_of_bias_matrix") cards(slide,spec.rob.slice(0,6).map(r=>({title:r.study_label||r.study||"Study",body:Object.entries(r).slice(1,4).map(([k,v])=>`${k}: ${v}`).join(" · ")})));
  else if(t==="grade_summary_table"||t==="evidence_certainty_ladder") cards(slide,spec.grade.map(r=>({title:r.outcome||"Outcome",body:`Certainty: ${r.certainty} · ${r.rationale||""}`})));
  else if(t==="protocol_design_schema") flow(slide,["Enrollment","TTE at 0–6 h","Prespecified covariates","AKI by 72 h","Planned analysis"]);
  else if(t==="eligibility_split_panel") cards(slide,spec.eligibility.map(r=>({title:r.type||r.category||"Criterion",body:r.criterion||r.definition||Object.values(r).join(" · ")})));
  else if(t==="assessment_schedule_timeline"){const rows=spec.schedule;slide.addShape(S.line,{x:1.2,y:3.6,w:10.8,h:0,line:{color:C.teal,width:3}});rows.slice(0,7).forEach((r,i)=>{const x=1.3+i*(10.6/Math.max(1,rows.length-1));slide.addShape(S.ellipse,{x:x-.1,y:3.5,w:.2,h:.2,fill:{color:r.status?.toLowerCase().includes("primary")?C.coral:C.teal},line:{color:C.white}});addText(slide,`${r.assessment}\n${r.window}`,x-.65,i%2?3.9:2.4,1.3,.72,13,{bold:true,align:"center"});});}
  else if(t==="exposure_outcome_definition_panel"){const rows=spec.variables.filter(r=>/primary exposure|primary outcome/i.test(Object.values(r).join(" ")));cards(slide,rows.map(r=>({title:r.variable||r.role||Object.values(r)[0],body:r.definition||Object.values(r).slice(1).join(" · ")})));}
  else if(t==="sample_size_assumption_cards") cards(slide,spec.sample.map(r=>({title:r.parameter||r.assumption||Object.values(r)[0],body:`${r.value||Object.values(r)[1]} ${r.unit||""} · planned`})));
  else if(t==="dag_native_or_fallback_diagram") flow(slide,["Baseline severity","Early TTE phenotype","Planned covariates","72 h AKI outcome"]);
  else if(t==="operational_risk_matrix") cards(slide,spec.risks.map(r=>({title:r.risk||Object.values(r)[0],body:`Probability: ${r.probability||"specified"} · Impact: ${r.impact||"specified"}`})));
  else if(t==="milestone_timeline"){const rows=spec.milestones;flow(slide,rows.map(r=>`${r.milestone||r.activity}\n${r.start_date||r.start}–${r.end_date||r.end}`));}
  else if(t==="no_result_boundary_card") cards(slide,[{title:"Protocol only",body:"No observed event rate, effect estimate, or trend is available."},{title:"Planned values",body:"Enrollment and timing are targets—not achievements."},{title:"Author review",body:"Ethics and funding markers remain unresolved."}]);
  footer(slide,contracts);notes(slide,contracts,[],[]);
}

const cover=pptx.addSlide("V23");
cover.background={color:C.navy};
cover.addShape(S.rect,{x:0,y:0,w:13.333,h:7.5,fill:{color:C.navy},line:{color:C.navy}});
cover.addShape(S.rect,{x:.85,y:1.05,w:.12,h:4.7,fill:{color:C.teal},line:{color:C.teal}});
addText(cover,deck.project_key.replace(/^\d+_/,"").replaceAll("_"," ").toUpperCase(),1.3,1.55,10.5,1.1,48,{bold:true,color:C.white});
addText(cover,"Evidence-to-Visual Narrative Remediation · Arm B2",1.32,3.1,9.8,.45,22,{color:"BFE6E3"});
addText(cover,"COMPLEX SYNTHETIC FIXTURE · NOT REAL CLINICAL EVIDENCE",1.32,5.55,10.5,.38,14,{bold:true,color:C.amber});
cover.addNotes("[Sources]\nSynthetic regression fixture; see project source manifest.");
for(const plan of deck.slides){
  const slide=pptx.addSlide("V23");
  if(plan.slide_role==="appendix") reviewSlide(slide); else renderTemplate(slide,plan);
}
fs.mkdirSync(path.dirname(outputPath),{recursive:true});
await pptx.writeFile({fileName:outputPath});
console.log(`WROTE=${outputPath}`);
console.log(`SLIDES=${deck.slides.length+1}`);
