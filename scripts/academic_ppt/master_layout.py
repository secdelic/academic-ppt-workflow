"""Role-aware real OOXML Master/Layout binding over immutable native scientific content."""
from __future__ import annotations
import copy,hashlib,json,posixpath,re,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
from .ooxml_qa import NS,inspect_pptx_ooxml
from .manual_image_handoff import check,sha_file
from .asset_approval import validate_use
P,A,R,REL,CT=(NS[k] for k in ('p','a','r','pr','ct'))
EMU=914400
LAYOUTS=('HERO_SCENE','CONTEXT_EDITORIAL','EVIDENCE_CLEAN','FRAMEWORK_EDITORIAL','IMPLEMENTATION_EDITORIAL','TIMELINE_CONTINUUM','CLOSING_EDITORIAL')
ROLE_LAYOUT={'HERO':'HERO_SCENE','CONTEXT':'CONTEXT_EDITORIAL','EVIDENCE':'EVIDENCE_CLEAN','RESULTS':'EVIDENCE_CLEAN','COMPARISON':'EVIDENCE_CLEAN','APPENDIX':'EVIDENCE_CLEAN','FRAMEWORK':'FRAMEWORK_EDITORIAL','IMPLEMENTATION':'IMPLEMENTATION_EDITORIAL','TIMELINE':'TIMELINE_CONTINUUM','CLOSING':'CLOSING_EDITORIAL'}
OVERRIDES=('NONE','CUSTOM_SCENE_EXCEPTION','SOURCE_FIGURE_EXCEPTION','TEMPLATE_AUTHORITY_EXCEPTION')
DEFAULT_THEME={'theme_id':'RESTRAINED_ACADEMIC','head_font':'Microsoft YaHei','body_font':'Microsoft YaHei','background':'F6F2EA','ink':'233F39','accent':'668B7D','muted':'7E9293','footer':'Academic presentation'}
SLOT_SPECS={'TITLE':('title',1,[.55,.34,12.2,.85],34),'SUBTITLE':('subTitle',2,[.55,1.3,12.2,.48],20),
 'BODY':('body',3,[.55,1.95,12.2,4.65],22),'FIGURE':('obj',4,[.8,2,11.7,4.4],22),
 'ANNOTATION':('body',5,[.55,6.15,12.2,.45],18),'SOURCE':('body',6,[.55,6.91,11.4,.24],10),
 'FOOTER':('ftr',7,[.55,7.23,10,.18],9),'PAGE_NUMBER':('sldNum',8,[12.2,6.91,.6,.24],10)}


def xml(node):return E.tostring(node,encoding='utf-8',xml_declaration=True)
def tree(root):return root.find('p:cSld/p:spTree',NS)
def texts(root):return [x.text or '' for x in root.findall('.//a:t',NS)]
def shape_name(shape):
    n=shape.find('.//p:cNvPr',NS)
    return n.get('name','') if n is not None else ''
def relroot():return E.Element('{'+REL+'}Relationships')
def addrel(root,rid,kind,target):E.SubElement(root,'{'+REL+'}Relationship',{'Id':rid,'Type':R+'/'+kind,'Target':target})
def part_target(part,target):return posixpath.normpath(posixpath.join(posixpath.dirname(part),target))


def shape_box(shape):
    transform=shape.find('p:spPr/a:xfrm',NS)
    if transform is None:return None
    off,ext=transform.find('a:off',NS),transform.find('a:ext',NS)
    return [int(off.get('x'))/EMU,int(off.get('y'))/EMU,int(ext.get('cx'))/EMU,int(ext.get('cy'))/EMU]


def set_box(shape,box):
    pr=shape.find('p:spPr',NS)
    for old in list(pr):
        if old.tag=='{'+A+'}xfrm':pr.remove(old)
    xf=E.Element('{'+A+'}xfrm');pr.insert(0,xf)
    E.SubElement(xf,'{'+A+'}off',{'x':str(round(box[0]*EMU)),'y':str(round(box[1]*EMU))})
    E.SubElement(xf,'{'+A+'}ext',{'cx':str(round(box[2]*EMU)),'cy':str(round(box[3]*EMU))})


def placeholder(role,typ,idx,box,size,font,ink):
    s=E.Element('{'+P+'}sp');nv=E.SubElement(s,'{'+P+'}nvSpPr')
    E.SubElement(nv,'{'+P+'}cNvPr',{'id':str(100+idx),'name':'Layout:'+role});E.SubElement(nv,'{'+P+'}cNvSpPr')
    props=E.SubElement(nv,'{'+P+'}nvPr');E.SubElement(props,'{'+P+'}ph',{'type':typ,'idx':str(idx)})
    E.SubElement(s,'{'+P+'}spPr');set_box(s,box)
    tx=E.SubElement(s,'{'+P+'}txBody');E.SubElement(tx,'{'+A+'}bodyPr',{'lIns':'0','rIns':'0','tIns':'0','bIns':'0'})
    lst=E.SubElement(tx,'{'+A+'}lstStyle');level=E.SubElement(lst,'{'+A+'}lvl1pPr',{'algn':'l','marL':'0','indent':'0'})
    E.SubElement(level,'{'+A+'}buNone')
    style=E.SubElement(level,'{'+A+'}defRPr',{'sz':str(round(size*100))})
    E.SubElement(E.SubElement(style,'{'+A+'}solidFill'),'{'+A+'}srgbClr',{'val':ink})
    for tag in ('latin','ea','cs'):E.SubElement(style,'{'+A+'}'+tag,{'typeface':font})
    E.SubElement(E.SubElement(tx,'{'+A+'}p'),'{'+A+'}endParaRPr',{'lang':'en-US'})
    return s


def decoration(name,identifier,box,color):
    s=E.Element('{'+P+'}sp');nv=E.SubElement(s,'{'+P+'}nvSpPr');E.SubElement(nv,'{'+P+'}cNvPr',{'id':str(identifier),'name':name})
    E.SubElement(nv,'{'+P+'}cNvSpPr');E.SubElement(nv,'{'+P+'}nvPr');pr=E.SubElement(s,'{'+P+'}spPr');set_box(s,box)
    E.SubElement(E.SubElement(pr,'{'+A+'}prstGeom',{'prst':'rect'}),'{'+A+'}avLst')
    E.SubElement(E.SubElement(pr,'{'+A+'}solidFill'),'{'+A+'}srgbClr',{'val':color})
    E.SubElement(E.SubElement(pr,'{'+A+'}ln'),'{'+A+'}noFill')
    return s


def layout_manifest():
    result={}
    for role in LAYOUTS:
        slots=copy.deepcopy(SLOT_SPECS)
        if role=='HERO_SCENE':slots['TITLE']=('title',1,[.7,1.15,7.1,1.6],42);slots['SUBTITLE']=('subTitle',2,[.7,3.1,7.1,.9],24)
        if role=='CLOSING_EDITORIAL':slots['TITLE']=('title',1,[.7,.65,11.9,1.15],34)
        result[role]={'layout_id':role,'title_safe_zone':slots['TITLE'][2],'content_safe_zone':[.55,1.95,12.2,4.65],
          'figure_safe_zone':slots['FIGURE'][2],'footer_exclusion':[0,6.75,13.333,.75],
          'optional_asset_zone':[8.2,2.1,4.3,4.2] if role=='HERO_SCENE' else ([.55,5.9,12.2,.7] if role=='TIMELINE_CONTINUUM' else None),
          'negative_space_zone':slots['TITLE'][2],'placeholders':slots,'scientific_content_on_master':False}
    return result


def validate_binding(binding):
    check(binding.get('page_role') in ROLE_LAYOUT,'PAGE_ROLE_REQUIRED')
    check(binding.get('layout_id') in LAYOUTS,'UNKNOWN_MASTER_LAYOUT')
    status=binding.get('override_status','NONE');check(status in OVERRIDES,'INVALID_LAYOUT_OVERRIDE')
    if status!='NONE':
        receipt=binding.get('override_receipt',{})
        check(binding.get('override_reason') and receipt.get('approved_by') and receipt.get('reason')
              and receipt.get('slide_id')==binding.get('slide_id'),'LAYOUT_OVERRIDE_RECEIPT_REQUIRED')
        return 'OVERRIDE_WITH_RECEIPT'
    check(binding['layout_id']==ROLE_LAYOUT[binding['page_role']],'ROLE_MISMATCH')
    return 'VALID'


def scientific_snapshot(parts):
    return {'texts':{n:texts(E.fromstring(v)) for n,v in parts.items() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)},
      'protected_parts':{n:hashlib.sha256(v).hexdigest() for n,v in parts.items() if n.startswith(('ppt/notesSlides/','ppt/charts/','ppt/embeddings/','ppt/media/'))}}


def bind_master(source,output,bindings,*,theme=None,official_template=False,scene_by_layout=None,asset_root=None,enabled=False):
    if not enabled:return None
    source,output=Path(source).resolve(),Path(output).resolve()
    check(source.is_file() and not output.exists() and source!=output,'NEW_MASTER_OUTPUT_REQUIRED')
    inspected=inspect_pptx_ooxml(source);check(inspected['valid'],'SOURCE_OOXML_INVALID')
    check(bindings and len({b['slide_index'] for b in bindings})==len(bindings) and len({b['slide_id'] for b in bindings})==len(bindings),'UNIQUE_MASTER_BINDINGS_REQUIRED')
    statuses=[validate_binding(b) for b in bindings];count=len(inspected['slide_order'])
    check(all(type(b['slide_index']) is int and 1<=b['slide_index']<=count for b in bindings),'MASTER_SLIDE_INDEX_INVALID')
    if official_template:
        # Existing institutional layouts remain authoritative; do not graft a replacement master.
        check(not scene_by_layout,'OFFICIAL_TEMPLATE_ASSET_ADAPTATION_REQUIRED')
        output.parent.mkdir(parents=True,exist_ok=True);output.write_bytes(source.read_bytes())
        return {'status':'OFFICIAL_TEMPLATE_PRESERVED','source_sha256':sha_file(source),'output_sha256':sha_file(output),
          'scientific_invariance':True,'bindings':[{**b,'master_binding_status':'OVERRIDE_WITH_RECEIPT','override_status':'TEMPLATE_AUTHORITY_EXCEPTION',
            'override_reason':'Institutional Master/Layout authority preserved','actual_layout_part':inspected['slide_relationship_chains'][b['slide_index']-1]['layout_part']} for b in bindings]}
    theme={**DEFAULT_THEME,**(theme or {})};check(theme['head_font'] in ('Arial','Microsoft YaHei') and theme['body_font'] in ('Arial','Microsoft YaHei'),'VALIDATED_FONT_REQUIRED')
    check(all(re.fullmatch(r'[A-Fa-f0-9]{6}',theme[k]) for k in ('background','ink','accent','muted')),'INVALID_THEME_COLOR')
    check(theme['footer']==DEFAULT_THEME['footer'],'MASTER_SCIENTIFIC_TEXT_FORBIDDEN')
    with zipfile.ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
    before=scientific_snapshot(parts);baseline_sha=sha_file(source);manifest=layout_manifest()
    # PowerPoint uses one global ID space for masters and custom layouts.
    shared_ids=[int(n.get('id')) for name,data in parts.items() if name=='ppt/presentation.xml' or re.fullmatch(r'ppt/slideMasters/slideMaster\d+\.xml',name)
                for n in E.fromstring(data).iter() if n.tag in ('{'+P+'}sldMasterId','{'+P+'}sldLayoutId')]
    next_shared_id=max(shared_ids,default=2147483647)+1
    size=E.fromstring(parts['ppt/presentation.xml']).find('p:sldSz',NS)
    width,height=int(size.get('cx'))/EMU,int(size.get('cy'))/EMU
    check(abs(width-13.333)<.01 and abs(height-7.5)<.01,'WIDE_CANVAS_REQUIRED_FOR_ROLE_LAYOUTS')
    # Some producers reserve content-type names without writing the corresponding
    # part. Allocate beyond both sets so the new package has no duplicate override.
    declared_parts={n.get('PartName','').lstrip('/') for n in E.fromstring(parts['[Content_Types].xml'])}
    reserved_parts=set(parts)|declared_parts
    def next_part(kind):
        return max([int(re.search(r'(\d+)\.xml$',n)[1]) for n in reserved_parts if re.fullmatch(fr'ppt/{kind}s/{kind}\d+\.xml',n)] or [0])+1
    master_num=next_part('slideMaster');start_layout=next_part('slideLayout');master_part=f'ppt/slideMasters/slideMaster{master_num}.xml'
    source_master=inspected['slide_relationship_chains'][0]['master_part'];master=E.fromstring(parts[source_master]);st=tree(master)
    for child in list(st)[2:]:st.remove(child)
    cs=master.find('p:cSld',NS);bg=cs.find('p:bg',NS)
    if bg is not None:cs.remove(bg)
    bg=E.Element('{'+P+'}bg');pr=E.SubElement(bg,'{'+P+'}bgPr');E.SubElement(E.SubElement(pr,'{'+A+'}solidFill'),'{'+A+'}srgbClr',{'val':theme['background']});E.SubElement(pr,'{'+A+'}effectLst');cs.insert(0,bg)
    st.append(decoration('Master:footer-rule',50,[.55,6.8,12.2,.012],theme['muted']))
    layout_ids=master.find('p:sldLayoutIdLst',NS)
    if layout_ids is None:layout_ids=E.SubElement(master,'{'+P+'}sldLayoutIdLst')
    layout_ids.clear();master_rels=relroot()
    source_theme=inspected['slide_relationship_chains'][0]['theme_part'];check(source_theme in parts,'SOURCE_THEME_REQUIRED')
    theme_number=max([int(re.search(r'(\d+)\.xml$',n)[1]) for n in reserved_parts if re.fullmatch(r'ppt/theme/theme\d+\.xml',n)] or [0])+1
    theme_part=f'ppt/theme/theme{theme_number}.xml';theme_xml=E.fromstring(parts[source_theme]);theme_xml.set('name',theme['theme_id'])
    colors=theme_xml.find('a:themeElements/a:clrScheme',NS)
    for key,value in {'dk1':theme['ink'],'lt1':theme['background'],'accent1':theme['accent'],'accent2':theme['muted']}.items():
        entry=colors.find('a:'+key,NS)
        for node in list(entry):entry.remove(node)
        E.SubElement(entry,'{'+A+'}srgbClr',{'val':value})
    for role,font in (('majorFont',theme['head_font']),('minorFont',theme['body_font'])):
        entry=theme_xml.find('a:themeElements/a:fontScheme/a:'+role,NS)
        for tag in ('latin','ea','cs'):entry.find('a:'+tag,NS).set('typeface',font)
    parts[theme_part]=xml(theme_xml)
    addrel(master_rels,'rIdTheme','theme',posixpath.relpath(theme_part,'ppt/slideMasters'))
    layout_parts={};slots_by_layout={};scene_by_layout=scene_by_layout or {};check(set(scene_by_layout)<=set(LAYOUTS),'UNKNOWN_SCENE_LAYOUT')
    for j,layout_id in enumerate(LAYOUTS):
        part=f'ppt/slideLayouts/slideLayout{start_layout+j}.xml';layout_parts[layout_id]=part
        layout=E.Element('{'+P+'}sldLayout',{'type':'cust','preserve':'1','showMasterSp':'1'});c=E.SubElement(layout,'{'+P+'}cSld',{'name':layout_id})
        local=copy.deepcopy(st)
        for child in list(local)[2:]:local.remove(child)
        c.append(local);E.SubElement(E.SubElement(layout,'{'+P+'}clrMapOvr'),'{'+A+'}masterClrMapping')
        if layout_id!='EVIDENCE_CLEAN':
            box=[.55,.23,.52,.045] if layout_id not in ('HERO_SCENE','CLOSING_EDITORIAL') else [.55,.6,.07,4.9 if layout_id=='HERO_SCENE' else 3.1]
            if layout_id=='TIMELINE_CONTINUUM':box=[.55,6.6,12.2,.025]
            local.append(decoration('Layout:role-accent',60,box,theme['accent']))
        slots={}
        for role,(typ,idx,box,font_size) in manifest[layout_id]['placeholders'].items():
            slot=placeholder(role,typ,idx,box,font_size,theme['head_font'] if role=='TITLE' else theme['body_font'],theme['ink']);local.append(slot);slots[role]=slot
        slots_by_layout[layout_id]=slots;rels=relroot();addrel(rels,'rIdMaster','slideMaster',f'../slideMasters/slideMaster{master_num}.xml')
        if layout_id in scene_by_layout:
            scene=scene_by_layout[layout_id];targets=[b for b in bindings if b['layout_id']==layout_id];check(targets,'UNUSED_MASTER_ASSET')
            for b in targets:path=validate_use(scene['asset'],asset_root,slide_id=b['slide_id'],baseline_sha256=baseline_sha)
            box=scene.get('box');check(isinstance(box,list) and len(box)==4 and all(type(v) in (int,float) for v in box),'SCENE_BOX_REQUIRED')
            check(box[0]>=0 and box[1]>=0 and box[2]>0 and box[3]>0 and box[0]+box[2]<=width and box[1]+box[3]<=6.75,'SCENE_OUTSIDE_SAFE_AREA')
            check(layout_id!='EVIDENCE_CLEAN','EVIDENCE_LAYOUT_NO_SCENE')
            media=f'ppt/media/master_{scene["asset"]["sha256"]}{path.suffix.lower()}';parts[media]=path.read_bytes();addrel(rels,'rIdScene','image',posixpath.relpath(media,'ppt/slideLayouts'))
            pic=E.Element('{'+P+'}pic');nv=E.SubElement(pic,'{'+P+'}nvPicPr');E.SubElement(nv,'{'+P+'}cNvPr',{'id':'70','name':'Layout:approved-scene'});E.SubElement(nv,'{'+P+'}cNvPicPr');E.SubElement(nv,'{'+P+'}nvPr')
            fill=E.SubElement(pic,'{'+P+'}blipFill');blip=E.SubElement(fill,'{'+A+'}blip',{'{'+R+'}embed':'rIdScene'});opacity=scene.get('opacity',.15);check(type(opacity) in (int,float) and 0<=opacity<=1,'INVALID_SCENE_OPACITY');E.SubElement(blip,'{'+A+'}alphaModFix',{'amt':str(round(opacity*100000))})
            crop=scene.get('crop',[0,0,0,0])
            check(isinstance(crop,list) and len(crop)==4 and all(type(v) in (float,int) and 0<=v<1 for v in crop) and crop[0]+crop[2]<1 and crop[1]+crop[3]<1,'INVALID_SCENE_CROP')
            E.SubElement(fill,'{'+A+'}srcRect',{k:str(round(v*100000)) for k,v in zip(('l','t','r','b'),crop)})
            E.SubElement(E.SubElement(fill,'{'+A+'}stretch'),'{'+A+'}fillRect');E.SubElement(pic,'{'+P+'}spPr');set_box(pic,box);E.SubElement(E.SubElement(pic.find('p:spPr',NS),'{'+A+'}prstGeom',{'prst':'rect'}),'{'+A+'}avLst');local.insert(2,pic)
        parts[part]=xml(layout)
        parts['ppt/slideLayouts/_rels/'+posixpath.basename(part)+'.rels']=xml(rels)
        rid=f'rIdLayout{j+1}';addrel(master_rels,rid,'slideLayout',f'../slideLayouts/{posixpath.basename(part)}');E.SubElement(layout_ids,'{'+P+'}sldLayoutId',{'id':str(next_shared_id+1+j),'{'+R+'}id':rid})
    parts[master_part]=xml(master);parts['ppt/slideMasters/_rels/'+posixpath.basename(master_part)+'.rels']=xml(master_rels)
    migrations=[]
    for b in bindings:
        part=inspected['slide_order'][b['slide_index']-1]['slide_part'];root=E.fromstring(parts[part]);root.set('showMasterSp','1');cs=root.find('p:cSld',NS)
        for bg in cs.findall('p:bg',NS):cs.remove(bg)
        for shape in tree(root):
            name=shape_name(shape);role='TITLE' if name.endswith(':title') or name.endswith(':cover-title') else ('SOURCE' if name.endswith(':sources') else ('PAGE_NUMBER' if name.endswith(':page-number') else None))
            if role is None or shape.tag!='{'+P+'}sp':continue
            old_text=texts(shape);nv=shape.find('p:nvSpPr/p:nvPr',NS)
            for ph in nv.findall('p:ph',NS):nv.remove(ph)
            nv.append(copy.deepcopy(slots_by_layout[b['layout_id']][role].find('.//p:ph',NS)))
            original_box=shape_box(shape)
            # Preserve scientific composition geometry, inherit layout typography.
            for run in shape.findall('.//a:rPr',NS)+shape.findall('.//a:endParaRPr',NS):
                for key in ('sz','b','i'):run.attrib.pop(key,None)
                for font in list(run):
                    if font.tag in ('{'+A+'}latin','{'+A+'}ea','{'+A+'}cs'):run.remove(font)
            check(texts(shape)==old_text,'SCIENTIFIC_TEXT_MUTATION')
            migrations.append({'slide_id':b['slide_id'],'role':role,'shape_name':name,'geometry_override':original_box,'reason':'Preserve native scientific composition; typography inherited from layout'})
        parts[part]=xml(root);relpart=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels';rels=E.fromstring(parts[relpart]);links=[x for x in rels if x.get('Type','').endswith('/slideLayout')];check(len(links)==1,'SLIDE_LAYOUT_RELATIONSHIP_REQUIRED')
        links[0].set('Target',posixpath.relpath(layout_parts[b['layout_id']],posixpath.dirname(part)));parts[relpart]=xml(rels)
    pres=E.fromstring(parts['ppt/presentation.xml']);ids=pres.find('p:sldMasterIdLst',NS);pr=E.fromstring(parts['ppt/_rels/presentation.xml.rels']);rid='rIdRoleMaster';check(all(x.get('Id')!=rid for x in pr),'MASTER_RELATIONSHIP_COLLISION')
    addrel(pr,rid,'slideMaster',f'slideMasters/slideMaster{master_num}.xml');E.SubElement(ids,'{'+P+'}sldMasterId',{'id':str(next_shared_id),'{'+R+'}id':rid})
    parts['ppt/presentation.xml']=xml(pres);parts['ppt/_rels/presentation.xml.rels']=xml(pr)
    types=E.fromstring(parts['[Content_Types].xml'])
    E.SubElement(types,'{'+CT+'}Override',{'PartName':'/'+theme_part,'ContentType':'application/vnd.openxmlformats-officedocument.theme+xml'})
    for part,kind in [(master_part,'slideMaster'),*[(x,'slideLayout') for x in layout_parts.values()]]:E.SubElement(types,'{'+CT+'}Override',{'PartName':'/'+part,'ContentType':f'application/vnd.openxmlformats-officedocument.presentationml.{kind}+xml'})
    for layout,scene in scene_by_layout.items():
        extension=Path(scene['asset']['file']).suffix.lstrip('.').lower();mime={'png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg'}.get(extension);check(mime,'MASTER_SCENE_FORMAT_UNSUPPORTED')
        if not any(x.get('Extension')==extension for x in types):E.SubElement(types,'{'+CT+'}Default',{'Extension':extension,'ContentType':mime})
    parts['[Content_Types].xml']=xml(types);after=scientific_snapshot(parts)
    check(before['texts']==after['texts'] and all(after['protected_parts'].get(n)==h for n,h in before['protected_parts'].items()),'SCIENTIFIC_INVARIANCE_FAILED')
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'x',zipfile.ZIP_DEFLATED) as archive:
        for name,data in parts.items():archive.writestr(name,data)
    qa=inspect_pptx_ooxml(output);check(qa['valid'],'MASTER_OOXML_INVALID')
    return {'status':'VALID','source_sha256':baseline_sha,'output_sha256':sha_file(output),'master_part':master_part,'layout_parts':layout_parts,
      'scientific_invariance':True,'scientific_content_on_master':False,'theme':theme,'layouts':manifest,'placeholder_bindings':migrations,
      'bindings':[{**b,'master_theme_id':theme['theme_id'],'master_binding_status':status,'actual_layout_part':layout_parts[b['layout_id']]} for b,status in zip(bindings,statuses)]}
