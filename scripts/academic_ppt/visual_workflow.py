"""Production handoff between source-bound plans, native generation and visual gates."""
from __future__ import annotations
import argparse,copy,json,shutil,zipfile,posixpath
from xml.etree import ElementTree as E
from pathlib import Path
from .utils import write_json,read_csv
from .deck_ir import canonical_json_hash
from .manual_image_handoff import check,sha_file,read_json
from .master_layout import bind_master,ROLE_LAYOUT
from . import visual_quality_floor as floor
from . import visual_ambition as ambition
from . import illustration_requests as illustration


def plan_path(args):
    explicit=getattr(args,'visual_plan',None)
    if explicit:return Path(explicit).resolve()
    project=getattr(args,'project_root',None)
    return Path(project)/'brief'/'visual_plan.json' if project else None


def load_plan(args,deck_ir):
    path=plan_path(args)
    if path is None or not path.exists():return None
    plan=read_json(path)
    check(plan.get('schema_version')=='assisted-visual-plan/1','VISUAL_PLAN_SCHEMA_REQUIRED')
    approval=plan.get('approval',{})
    check(approval.get('status')=='approved' and approval.get('reviewer') and approval.get('reviewed_at'),'APPROVED_VISUAL_PLAN_REQUIRED')
    check(plan.get('deck_content_sha256')==deck_ir['canonical_hash'],'VISUAL_PLAN_SCIENTIFIC_BINDING_CHANGED')
    rows=plan.get('slides',[]);expected=[s['slide_id'] for s in deck_ir['slides']]
    check([s.get('slide_id') for s in rows]==expected and [s.get('slide_index') for s in rows]==list(range(1,len(rows)+1)),'VISUAL_PLAN_PAGE_COVERAGE_INVALID')
    check(all(s.get('page_role') in ROLE_LAYOUT for s in rows),'EXPLICIT_PAGE_ROLE_REQUIRED')
    return plan


def prepare_generation(args,deck_ir,spec_path):
    plan=load_plan(args,deck_ir)
    if plan is None:return None
    assignments={s['slide_id']:s['semantic_composition'] for s in plan['slides'] if s.get('semantic_composition')}
    if assignments:
        spec=read_json(spec_path)
        check(getattr(args,'route','generate')=='generate','SEMANTIC_RECOMPOSITION_REQUIRES_GENERATE_ROUTE')
        spec['composition']={'id':'SEMANTIC_COMPOSITION','enabled':True};spec['visual_execution_plan']=assignments
        write_json(spec_path,spec)
    return plan


def materialize(args,deck_ir,source,staging,plan):
    if plan is None:return None
    if getattr(args,'route','generate')=='enhance-existing':
        check(not plan.get('native_style'),'KEEP_ROUTE_REQUIRES_EXPLICIT_CHANGED_SLIDE_OPERATIONS')
    staging=Path(staging);source=Path(source);current=source
    if plan.get('native_style'):
        from .art_direction import apply_native_art_direction
        style=copy.deepcopy(plan['native_style']);style['source_pptx_sha256']=sha_file(source)
        styled=staging/'native_art_direction.pptx';report=apply_native_art_direction(source,styled,style)
        write_json(staging/'art_direction_execution.json',report);current=styled
    bindings=[]
    for row in plan['slides']:
        bindings.append({**{k:row[k] for k in ('slide_id','slide_index','page_role')},'layout_id':row.get('layout_id',ROLE_LAYOUT[row['page_role']]),
          'binding_reason':row.get('binding_reason','Approved explicit communicative role'),'override_status':row.get('override_status','NONE'),
          'override_reason':row.get('override_reason'),'override_receipt':row.get('override_receipt')})
    target=staging/'master_bound.pptx'
    official=getattr(args,'route','generate')=='fill-template' or bool(getattr(args,'official_template',False))
    keep=getattr(args,'route','generate')=='enhance-existing'
    if keep:check(all(b['override_status']!='NONE' for b in bindings),'KEEP_MASTER_BINDING_REQUIRES_LAYOUT_EXCEPTION_RECEIPTS')
    report=bind_master(current,target,bindings,theme=plan.get('theme'),official_template=official or keep,enabled=True)
    if keep:report['status']='EXISTING_DECK_MASTER_PRESERVED'
    # source is the new run's generated draft, never a registered input or prior formal deck.
    shutil.copyfile(target,source)
    write_json(staging/'master_binding_manifest.json',report)
    return report


def record_core_qa(staging,pptx,scientific_issues,visual_issues,file_issues,layout_report,deck_ir):
    status=lambda issues:'FAIL' if issues else 'PASS'
    geometry=layout_report.get('status','NOT_ASSESSED')
    receipt={'schema_version':'production-visual-qa-receipt/1','deck_sha256':sha_file(pptx),'content_sha256':deck_ir['canonical_hash'],
      'scientific':status(scientific_issues),'semantic':status(scientific_issues),'geometry':'PASS' if geometry in ('PASS','passed') and not visual_issues else 'FAIL' if visual_issues else 'NOT_ASSESSED',
      'typography':'PASS' if geometry in ('PASS','passed') and not visual_issues else 'NOT_ASSESSED','files':status(file_issues),
      'geometry_pages':[{'slide_index':i} for i in range(1,len(deck_ir['slides'])+1)]}
    write_json(Path(staging)/'visual_qa_receipt.json',receipt)
    return receipt


def plan_template(deck_ir):
    roles={'cover':'HERO','closing':'CLOSING','appendix':'APPENDIX','results':'RESULTS','methods':'EVIDENCE','context':'CONTEXT'}
    return {'schema_version':'assisted-visual-plan/1','deck_content_sha256':deck_ir['canonical_hash'],
      'approval':{'status':'pending','reviewer':None,'reviewed_at':None},
      'slides':[{'slide_id':s['slide_id'],'slide_index':i,'page_role':roles.get(s.get('slide_role')),
        'composition_family':s.get('layout_family') or s.get('layout_id'),'semantic_mapping':{'mapping_status':'NOT_ASSESSED'},
        'illustration_decision':None,'ambition_context':None,'illustration_context':None,
        'fallback_reason':None,'fallback_alternatives':[]} for i,s in enumerate(deck_ir['slides'],1)]}


def finalize_run(output,args):
    """Every public validated/full result carries a fail-closed visual delivery gate."""
    output=Path(output);quality=getattr(args,'quality','validated');route=getattr(args,'route','generate')
    if route=='create-style-profile':
        gate={'status':'NOT_APPLICABLE_STYLE_PROFILE_ONLY','quality':quality,'floor_certified':False}
        write_json(output/'visual_delivery_gate.json',gate);return gate
    if quality=='quick':
        gate={'status':'DRAFT_LIMITED_QA','quality':'quick','floor_certified':False,'human_scientific_review_required':True,'human_visual_review_required':True}
        write_json(output/'visual_delivery_gate.json',gate);return gate
    deck_paths=sorted(output.glob('*.pptx'));check(len(deck_paths)==1,'ONE_ASSEMBLED_DECK_REQUIRED')
    pptx=deck_paths[0];stage=Path(getattr(args,'effective_staging_root',None) or output);ir_path=output/'deck_ir.json'
    if not ir_path.exists():ir_path=stage/'deck_ir.json'
    if not ir_path.exists():
        gate={'status':'BLOCK','quality':quality,'floor_certified':False,'reasons':['SOURCE_BOUND_DECK_IR_REQUIRED_FOR_VISUAL_QA'],
          'next_action':'Run full scientific planning for an existing deck before attempting visual-floor certification.'}
        write_json(output/'visual_delivery_gate.json',gate);return gate
    ir=read_json(ir_path);plan=load_plan(args,ir)
    evidence=output/'visual_evidence';evidence.mkdir(exist_ok=False)
    write_json(evidence/'visual_plan_template.json',plan_template(ir))
    qa_path=stage/'visual_qa_receipt.json';qa=read_json(qa_path) if qa_path.exists() else {'deck_sha256':sha_file(pptx),'content_sha256':ir['canonical_hash'],
      'scientific':'NOT_ASSESSED','semantic':'NOT_ASSESSED','geometry':'NOT_ASSESSED','typography':'NOT_ASSESSED','geometry_pages':[]}
    check(qa.get('deck_sha256')==sha_file(pptx),'QA_RECEIPT_DECK_CHANGED')
    write_json(evidence/'qa.json',qa);qa_path=evidence/'qa.json'
    previews=sorted((output/'preview').glob('slide_*.png'));sheet=output/'preview'/'contact_sheet.png';pdf=pptx.with_suffix('.pdf')
    records=[];observations=[]
    execution_path=stage/'artifact_render'/'composition_execution.json';execution=read_json(execution_path) if execution_path.exists() else {'slides':[]}
    composed={x['slide_id']:x for x in execution['slides']}
    package=floor.inspect_pptx_ooxml(pptx)
    with zipfile.ZipFile(pptx) as z:
        for i,slide in enumerate(ir['slides'],1):
            row=plan['slides'][i-1] if plan else {};sid=slide['slide_id'];family=row.get('composition_family')
            family=family.get('selected_family') if isinstance(family,dict) else family
            actual=composed.get(sid)
            # Executed native shapes are checked independently of the requested family.
            xml=E.fromstring(z.read(package['slide_order'][i-1]['slide_part']))
            from .master_layout import NS
            native_count=sum(1 for o in xml.findall('.//p:cNvPr',NS) if o.get('name','').startswith('composition:'))
            observed=actual['archetype_rendered'] if actual and native_count==actual['object_count'] else None
            blocks=[''.join(p.itertext()) for p in xml.findall('.//a:p',NS)]
            longest=max(map(len,blocks),default=0)
            chars=sum(map(len,blocks))
            # Non-recomposed content needs an explicit inspected observation receipt.
            records.append({'slide_id':sid,'slide_index':i,'page_role':row.get('page_role'),
              'semantic_mapping':row.get('semantic_mapping',{'mapping_status':'NOT_ASSESSED'}),
              'composition_family':{'selected_family':family,'generic_fallback':bool(row.get('fallback_reason')),
                'fallback_reason':row.get('fallback_reason'),'fallback_reason_provenance':'APPROVED_VISUAL_PLAN' if row.get('fallback_reason') else None,
                'candidate_evaluation':row.get('fallback_alternatives',[]),
                'selector_inputs':{'content_density':'HIGH' if chars>600 else 'MEDIUM' if chars>300 else 'LOW','max_text_block_length':longest}},
              'art_direction_execution':{'status':'RENDER_REVIEW_REQUIRED'},'illustration_decision':row.get('illustration_decision'), 'visual_QA':[]})
            observations.append({'slide_id':sid,'slide_index':i,'slide_sha256':floor.hashlib.sha256(z.read(package['slide_order'][i-1]['slide_part'])).hexdigest(),'observed_family':observed})
    write_json(evidence/'observations.json',observations)
    visual_brief=getattr(args,'visual_brief_path',None)
    if not visual_brief and getattr(args,'project_root',None):visual_brief=Path(args.project_root)/'brief'/'visual_brief.yaml'
    ref=floor.reference
    manifest={'deck':ref(pptx),'project_id':ir['deck_id'],'deck_content_binding':ref(qa_path,'/content_sha256'),'route':route,'records':records,
      'scientific_inputs':[ref(p) for p in (output/'source_manifest.csv',ir_path) if p.is_file()],
      'observations':[{'slide_index':i,'evidence':ref(evidence/'observations.json',f'/{i-1}'),'preview_sha256':sha_file(previews[i-1]) if len(previews)>=i else None} for i in range(1,len(records)+1)],
      'qa':{k:{'status':ref(qa_path,'/'+k),'deck_binding':ref(qa_path,'/deck_sha256')} for k in ('scientific','semantic','geometry','typography')},
      'render':{'deck_binding':ref(qa_path,'/deck_sha256'),'previews':[{**ref(p),'slide_index':i} for i,p in enumerate(previews,1)],
        'contact_sheet':{**ref(sheet),'slide_indexes':list(range(1,len(records)+1))} if sheet.exists() else {},'pdf':ref(pdf) if pdf.exists() else {}},
      'assets':[],'art_direction':{'brief':ref(visual_brief) if visual_brief and Path(visual_brief).is_file() else {},'fields':[]},'human_reviews':{}}
    manifest['qa']['geometry']['pages']=ref(qa_path,'/geometry_pages')
    supplied=getattr(args,'visual_evidence',None)
    extra={}
    if supplied:
        extra=read_json(supplied);check(extra.get('deck_sha256')==sha_file(pptx),'VISUAL_EVIDENCE_DECK_CHANGED')
        for key in ('observations','assets','art_direction','human_reviews'):
            if key in extra:manifest[key]=extra[key]
    write_json(evidence/'quality_manifest.json',manifest)
    report=floor.evaluate(manifest,mode='enforce');floor.write_outputs(report,evidence/'floor')
    master_path=stage/'master_binding_manifest.json';master=read_json(master_path) if master_path.exists() else None
    if master:write_json(evidence/'master_binding_manifest.json',master)
    master_ok=bool(master and master.get('status') in ('VALID','OFFICIAL_TEMPLATE_PRESERVED','EXISTING_DECK_MASTER_PRESERVED') and
                   master.get('output_sha256')==sha_file(pptx) and len(master.get('bindings',[]))==len(records))
    reasons=[]
    if not plan:reasons.append('APPROVED_VISUAL_PLAN_REQUIRED')
    if not master_ok:reasons.append('MASTER_BINDING_MISSING_OR_UNVERIFIED')
    if report['status']=='VISUAL_QUALITY_BLOCKED':reasons.append('VISUAL_QUALITY_BLOCKED')
    if qa.get('files')=='FAIL':reasons.append('CORE_FILE_QA_FAILED')
    ambition_status='NOT_RUN_VALIDATED_MODE'
    ambition_review_pending=False
    if quality=='full':
        contexts=[x.get('ambition_context') for x in plan['slides']] if plan else []
        if contexts and all(contexts) and visual_brief:
            import yaml
            for context in contexts:
                context['evidence_refs']=[ref(plan_path(args))]
            bundle={'floor_report':report,'floor_records':report['slide_records'],'contexts':contexts,
              'visual_brief':yaml.safe_load(Path(visual_brief).read_text(encoding='utf-8-sig')),'third_party_upload_allowed':False}
            planned=ambition.plan(bundle,enabled=True);ambition.write_outputs(planned,evidence/'ambition');ambition_status=planned['visual_ambition_status']
            expected_ids=[x['slide_id'] for x in planned['records'] if x['is_anchor'] or x['recommended_ambition_level']=='STRETCH']
            review=extra.get('ambition_review',{})
            ambition_review_pending=not (review.get('status')=='REVIEWED' and review.get('reviewer') and review.get('reviewed_at') and
                review.get('deck_sha256')==sha_file(pptx) and review.get('render_set_sha256')==report['render_set_sha256'] and
                review.get('plan_sha256')==floor.digest(planned) and review.get('slide_ids')==expected_ids)
            write_json(evidence/'ambition_review_template.json',{'status':None,'reviewer':None,'reviewed_at':None,
              'deck_sha256':sha_file(pptx),'render_set_sha256':report['render_set_sha256'],'plan_sha256':floor.digest(planned),
              'slide_ids':expected_ids,'decision':None,'comments':None})
            requests=[]
            for row,record in zip(plan['slides'],planned['records']):
                if row.get('illustration_context') is not None:
                    request=illustration.build_request(record,row['illustration_context'],plan['illustration_style'],row['illustration_sources'],enabled=True)
                    requests.append({'slide_index':row['slide_index'],'request':request})
            write_json(evidence/'manual_illustration_records.json',requests)
        else:
            ambition_status='HUMAN_REVIEW_REQUIRED_MISSING_CONTEXT';reasons.append('ANCHOR_OPPORTUNITY_CONTEXT_REQUIRED')
        from scripts.audit_repository_content import _scan_office_package
        privacy=[];parts,external=_scan_office_package(pptx,pptx.name,privacy)
        deep={'status':'FAIL' if privacy or not package['valid'] else 'PASS','privacy_findings':privacy,
          'xml_parts_scanned':parts,'external_relationships':external,'ooxml_valid':package['valid'],
          'scientific_provenance_sha256':report['verified_input_sha256'],'geometry_receipt':qa,
          'native_object_manifest':package['object_manifest'],'patient_content_requires_human_review':True}
        write_json(evidence/'deep_final_audit.json',deep)
        if deep['status']=='FAIL':reasons.append('DEEP_FINAL_AUDIT_FAILED')
    gate={'status':'BLOCK' if reasons else 'HUMAN_REVIEW' if report['status']=='VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED' or ambition_review_pending else 'PASS',
      'quality':quality,'floor_certified':report['deck_floor_acceptance']=='PROCESS_FLOOR_SATISFIED' and not reasons,
      'floor_status':report['status'],'ambition_status':ambition_status,'master_binding_status':master.get('status') if master else 'MISSING',
      'ambition_human_review_pending':ambition_review_pending,
      'deck_sha256':sha_file(pptx),'render_set_sha256':report['render_set_sha256'],'reasons':reasons,
      'second_independent_real_project_validated':False,'aesthetic_quality_guaranteed':False}
    write_json(output/'visual_delivery_gate.json',gate)
    write_json(output/'visual_review_context.json',{'quality':quality,'route':route,'effective_staging_root':str(stage),
      'project_root':str(args.project_root) if getattr(args,'project_root',None) else None,
      'visual_plan':str(plan_path(args)) if plan_path(args) else None,'visual_brief_path':str(visual_brief) if visual_brief else None,
      'deck_sha256':sha_file(pptx)})
    return gate


def review_run(source,output,*,visual_evidence,visual_plan=None):
    """Re-evaluate a reviewed immutable deck; never regenerate the reviewed bytes."""
    source,output=Path(source).resolve(),Path(output).resolve()
    check(source.is_dir() and not output.exists() and source not in output.parents,'NEW_REVIEW_OUTPUT_REQUIRED')
    context=read_json(source/'visual_review_context.json')
    decks=list(source.glob('*.pptx'));check(len(decks)==1 and sha_file(decks[0])==context['deck_sha256'],'REVIEW_DECK_CHANGED')
    context.pop('deck_sha256');context['visual_evidence']=str(visual_evidence)
    if visual_plan is not None:context['visual_plan']=str(visual_plan)
    shutil.copytree(source,output,ignore=shutil.ignore_patterns('visual_evidence','visual_delivery_gate.json','visual_review_context.json'))
    return finalize_run(output,argparse.Namespace(**context))


def audit_assembly(source_run,assembled,output):
    """Render and audit an authorized image-only assembly without re-planning science."""
    from .master_layout import scientific_snapshot,tree,xml,NS
    from .asset_approval import validate_use
    from .rendering import render_pptx,create_contact_sheet
    from .visual_layout_qa import inspect_powerpoint_text_layout,inspect_text_geometry
    source_run,assembled,output=map(lambda p:Path(p).resolve(),(source_run,assembled,output))
    check(not output.exists() and source_run not in output.parents,'NEW_ASSEMBLY_AUDIT_OUTPUT_REQUIRED')
    context=read_json(source_run/'visual_review_context.json');baseline=list(source_run.glob('*.pptx'))
    receipt=read_json(assembled.with_suffix('.assembly.json'))
    check(len(baseline)==1 and sha_file(baseline[0])==context['deck_sha256']==receipt['source_sha256'] and sha_file(assembled)==receipt['output_sha256'],'ASSEMBLY_BASELINE_CHANGED')
    registry=Path(receipt['registry']['path']);check(sha_file(registry)==receipt['registry']['sha256'],'ASSEMBLY_REGISTRY_CHANGED')
    assets=read_json(registry)['assets'];ir=read_json(source_run/'deck_ir.json')
    for asset in assets:validate_use(asset,registry.parent,slide_id=ir['slides'][asset['slide_index']-1]['slide_id'],baseline_sha256=context['deck_sha256'])
    with zipfile.ZipFile(baseline[0]) as z:before={n:z.read(n) for n in z.namelist()}
    with zipfile.ZipFile(assembled) as z:after={n:z.read(n) for n in z.namelist()}
    old=scientific_snapshot(before);new=scientific_snapshot(after)
    check(old['texts']==new['texts'] and all(new['protected_parts'].get(n)==h for n,h in old['protected_parts'].items()),'ASSEMBLY_SCIENTIFIC_INVARIANCE_FAILED')
    old_types=[(n.tag,tuple(sorted(n.attrib.items()))) for n in E.fromstring(before['[Content_Types].xml'])]
    new_types=[(n.tag,tuple(sorted(n.attrib.items()))) for n in E.fromstring(after['[Content_Types].xml'])]
    check(all(n in new_types for n in old_types),'ASSEMBLY_CONTENT_TYPE_CHANGED')
    authorized={a['slide_index']:a for a in assets}
    package=floor.inspect_pptx_ooxml(assembled);check(package['valid'],'ASSEMBLY_OOXML_INVALID')
    mutable={'[Content_Types].xml'}
    for i,page in enumerate(package['slide_order'],1):
        name=page['slide_part'];a,b=E.fromstring(before[name]),E.fromstring(after[name]);old_shapes=list(tree(a));new_shapes=list(tree(b))
        check([xml(s) for s in old_shapes]==[xml(s) for s in new_shapes[:len(old_shapes)]],'ASSEMBLY_NATIVE_OBJECT_CHANGED')
        added=new_shapes[len(old_shapes):];check(len(added)==(1 if i in authorized else 0),'UNAUTHORIZED_ADDED_OBJECT')
        if added:
            check(added[0].tag=='{'+NS['p']+'}pic','ONLY_APPROVED_PICTURE_ALLOWED')
            check(added[0].find('.//p:cNvPr',NS).get('name')=='approved-asset:'+authorized[i]['asset_id'],'ASSEMBLY_ASSET_ID_CHANGED')
            folder,filename=name.rsplit('/',1);rel=folder+'/_rels/'+filename+'.rels';mutable.update((name,rel))
            old_rels={n.get('Id'):n.attrib for n in E.fromstring(before[rel])};new_rels={n.get('Id'):n.attrib for n in E.fromstring(after[rel])}
            check(all(new_rels.get(k)==v for k,v in old_rels.items()),'ASSEMBLY_SOURCE_RELATIONSHIP_CHANGED')
            rid=added[0].find('.//a:blip',NS).get('{'+NS['r']+'}embed');link=new_rels[rid]
            check(set(new_rels)-set(old_rels)=={rid} and link.get('TargetMode')!='External','ASSEMBLY_IMAGE_RELATIONSHIP_INVALID')
            media=posixpath.normpath(posixpath.join(folder,link['Target']))
            check(floor.hashlib.sha256(after[media]).hexdigest()==authorized[i]['sha256'],'ASSEMBLY_IMAGE_BYTES_CHANGED')
    check(all(after.get(n)==data for n,data in before.items() if n not in mutable),'ASSEMBLY_UNAUTHORIZED_PACKAGE_MUTATION')
    check(all(n.startswith('ppt/media/') for n in set(after)-set(before)),'ASSEMBLY_UNAUTHORIZED_NEW_PART')
    old_stage=Path(context['effective_staging_root']);qa=read_json(old_stage/'visual_qa_receipt.json')
    check(qa['deck_sha256']==context['deck_sha256'] and qa['scientific']=='PASS','BASELINE_SCIENTIFIC_QA_REQUIRED')
    shutil.copytree(source_run,output,ignore=shutil.ignore_patterns('*.pptx','*.pdf','preview','visual_evidence','visual_delivery_gate.json','visual_review_context.json'))
    stage=output/'assembly_evidence';stage.mkdir();target=output/'deck.pptx';shutil.copyfile(assembled,target)
    scripts=Path(__file__).resolve().parents[1]
    method,previews,_=render_pptx(target,output/'deck.pdf',output/'preview',scripts,180)
    create_contact_sheet(previews,output/'preview'/'contact_sheet.png')
    layout,_=inspect_powerpoint_text_layout(pptx_path=target,output_json=stage/'powerpoint_layout_manifest.json',script_path=scripts/'inspect_pptx_layout.ps1')
    issues=inspect_text_geometry(layout);record_core_qa(stage,target,[],issues,[],layout,ir)
    master=read_json(old_stage/'master_binding_manifest.json');check(master['output_sha256']==context['deck_sha256'],'BASELINE_MASTER_RECEIPT_CHANGED')
    master.update(source_sha256=context['deck_sha256'],output_sha256=sha_file(target));write_json(stage/'master_binding_manifest.json',master)
    if (old_stage/'artifact_render').exists():shutil.copytree(old_stage/'artifact_render',stage/'artifact_render')
    write_json(stage/'assembly_invariance.json',{'status':'PASS','source_sha256':context['deck_sha256'],'output_sha256':sha_file(target),'renderer':method,'registry':receipt['registry'],'new_human_review_required':True})
    context.pop('deck_sha256');context['effective_staging_root']=str(stage);context['visual_evidence']=None
    return finalize_run(output,argparse.Namespace(**context))


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--review-run',type=Path);parser.add_argument('--output',type=Path)
    parser.add_argument('--visual-evidence',type=Path);parser.add_argument('--visual-plan',type=Path)
    parser.add_argument('--audit-assembly',type=Path,help='New approved-assets PPTX with adjacent assembly receipt')
    parser.add_argument('--source-run',type=Path)
    args=parser.parse_args()
    if args.audit_assembly:
        if args.output is None or args.source_run is None:parser.error('source run and new output are required')
        result=audit_assembly(args.source_run,args.audit_assembly,args.output)
        print(json.dumps(result));raise SystemExit(2 if result['status']=='BLOCK' else 0)
    if args.review_run is None:return
    if args.output is None or args.visual_evidence is None:parser.error('output and visual evidence are required')
    result=review_run(args.review_run,args.output,visual_evidence=args.visual_evidence,visual_plan=args.visual_plan)
    print(json.dumps(result));raise SystemExit(2 if result['status']=='BLOCK' else 0)

if __name__=='__main__':main()
