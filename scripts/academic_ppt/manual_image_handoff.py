"""Local manual generation packets and unapproved intake. No network generation."""
from __future__ import annotations
import argparse,copy,csv,hashlib,io,json,re,stat,warnings
from datetime import datetime, timezone
from pathlib import Path
from . import illustration_requests as requests
ROOT=Path(__file__).resolve().parents[2]
REVIEW_FIELDS=['candidate_id','asset_request_id','target_slide_id','approve','decision','comments']
DECISIONS=['APPROVE','APPROVE_WITH_REFINEMENT','REJECT','NO_MEANINGFUL_GAIN','TOO_DECORATIVE','SEMANTIC_CONFUSION_RISK']
BOUNDARY=('This visual is for presentation only. It is not scientific evidence and must not contain '
          'or imitate real patient data, clinical imaging, physiologic monitoring traces, statistical results or diagnostic findings.')
FORMATS={'.png':('PNG','image/png'),'.jpg':('JPEG','image/jpeg'),'.jpeg':('JPEG','image/jpeg'),'.webp':('WEBP','image/webp')}
MAX_BYTES=50*1024*1024
MAX_PIXELS=40_000_000
DIRECTIONS={
 'EDITORIAL_HORIZON':{'manual_role':'HERO_SCENE','type':'FULL_SCENE','elements':'Layered abstract forms with foreground, midground and background; no specific geography.','extra_forbidden':'Readable text, patients, equipment and measurements.','output':'16:9 full scene, preferably 1536x864 or equivalent PNG.'},
 'RESEARCH_PROCESS':{'manual_role':'PROCESS_ICON_FAMILY','type':'ICON_SET','elements':'A coherent icon family for search, selection, consultation, review and document output; equal scale.','extra_forbidden':'Readable text, numerals, data encodings and evidence-grade badges.','output':'Transparent PNG icon family; all scientific labels stay in PowerPoint.'},
 'CONCEPTUAL_FRAMEWORK':{'manual_role':'ABSTRACT_FRAMEWORK','type':'TRANSPARENT_FOREGROUND','elements':'Quiet abstract layers; leave the native framework and categories authoritative.','extra_forbidden':'Anatomical labels, lesions, physiology traces and diagnostic maps.','output':'Transparent PNG with an isolated conceptual illustration.'},
 'IMPLEMENTATION_CUES':{'manual_role':'IMPLEMENTATION_CUES','type':'ICON_SET','elements':'Restrained safety, coordination and check cues.','extra_forbidden':'Protocols, outcome claims, measurements and labels.','output':'Transparent PNG; separable cues with consistent visual weight.'},
 'CONTINUITY_BACKGROUND':{'manual_role':'CONTINUITY_BACKGROUND','type':'BACKGROUND','elements':'Subtle pathway or horizon suggesting visual continuity.','extra_forbidden':'Additional stages, time points, efficacy claims or patient portraits.','output':'16:9 restrained background, preferably 1536x864 or equivalent PNG.'},
}


class BridgeError(ValueError):
    pass

def check(ok, code):
    if not ok:
        raise BridgeError(code)

def sha_bytes(data):
    return hashlib.sha256(data).hexdigest()

def sha_file(path):
    return sha_bytes(Path(path).read_bytes())

def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def write_new(path, data):
    with Path(path).open('x', encoding='utf8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, allow_nan=False)
        f.write('\n')

def now():
    return datetime.now(timezone.utc).isoformat()

def region_text(region):
    # Normalized coordinates, not patient values. No free-form location string accepted.
    requests.rectangle(region)
    return ', '.join(f'{key} {region[key] * 100:g}%' for key in ('x', 'y', 'w', 'h')) + ' of the canvas (origin: upper left)'

def make_packet(request, *, enabled=False):
    if not enabled:
        return None
    try:
        requests.validate_request(request)
    except (ValueError, KeyError, TypeError) as exc:
        raise BridgeError('MANUAL_GENERATION_PACKET_UNSAFE') from exc
    # No-art, unresolved or source-figure requests are not eligible.
    if request['status'] != 'ILLUSTRATION_CANDIDATE_PENDING' or not request['illustration_needed']:
        return None
    check(request['scientific_evidence'] is False and request['presentation_only'] is True, 'MANUAL_GENERATION_PACKET_UNSAFE')
    check(request['subject_class'] in DIRECTIONS, 'MANUAL_GENERATION_PACKET_UNSAFE')
    direction = DIRECTIONS[request['subject_class']]
    composition = request['composition_role'].lower().replace('_', ' ') + '. Keep all native content unobscured; render no labels inside the artwork.'
    negative = 'Leave this region clean and quiet for native slide content: ' + region_text(request['negative_space_region']) + '.'
    focal = 'Place the focal artwork within: ' + region_text(request['focal_region']) + '.'
    depth = request['depth_role'].lower().replace('_', ' ') + '. Visual weight: ' + request['visual_weight'].lower() + '.'
    forbidden = '; '.join(item.lower().replace('_', ' ') for item in request['prohibited_visual_types']) + '. ' + direction['extra_forbidden']
    packet = {
        'schema_version': 'manual-generation-packet/1', 'asset_request_id': request['request_id'],
        'target_slide_id': request['slide_id'], 'page_role': request['page_role'], 'anchor_role': request['anchor_role'],
        'ambition_level': request['ambition_level'], 'illustration_role': request['illustration_role'],
        'manual_trial_role': direction['manual_role'], 'subject_class': request['subject_class'],
        'short_conceptual_subject': request['subject_summary'],
        'style_family': request['style_family'],
        'visual_style': request['style_family'].lower().replace('_', ' '),
        'allowed_visual_types': copy.deepcopy(request['allowed_visual_types']),
        'prohibited_visual_types': copy.deepcopy(request['prohibited_visual_types']),
        'palette': copy.deepcopy(request['palette']), 'aspect_ratio': request['aspect_ratio'],
        'composition_instruction': composition, 'negative_space_instruction': negative,
        'focal_region_instruction': focal, 'depth_instruction': depth,
        'allowed_elements': direction['elements'], 'forbidden_elements': forbidden,
        'scientific_boundary': BOUNDARY, 'preferred_asset_type': direction['type'], 'output_preference': direction['output'],
        'scientific_evidence': False, 'presentation_only': True, 'human_generation_required': True,
        'human_approval_required': True, 'automatic_generation_allowed': False,
        'source_request_sha256': requests.canonical_hash(request),
    }
    sections = [('PURPOSE', 'Create one presentation-only conceptual visual for manual human review. No scientific claim is requested.'),
                ('SUBJECT', packet['short_conceptual_subject']), ('STYLE', packet['visual_style']),
                ('PALETTE', ', '.join(packet['palette'])), ('COMPOSITION', composition + ' ' + focal),
                ('NEGATIVE SPACE', negative), ('VISUAL DEPTH', depth), ('ALLOWED ELEMENTS', packet['allowed_elements']),
                ('FORBIDDEN ELEMENTS', forbidden), ('SCIENTIFIC BOUNDARY', BOUNDARY),
                ('OUTPUT PREFERENCE', packet['output_preference'])]
    packet['prompt'] = '\n\n'.join(title + '\n' + value for title, value in sections)
    return packet

def validate_packet(packet, request):
    expected = make_packet(request, enabled=True)
    check(expected is not None and packet == expected, 'MANUAL_GENERATION_PACKET_UNSAFE')

def packet_markdown(packet, slide_index):
    def readable(value):
        if isinstance(value,bool):return str(value).lower()
        if isinstance(value,list):return '; '.join(value)
        return str(value)
    metadata='\n'.join(f'- **{key}**: {readable(value)}' for key,value in packet.items() if key!='prompt')
    return (f'# S{slide_index:02d} manual generation packet\n\n'
            'Presentation only; human generation and human approval required.\n\n'
            'Copy only the following prompt block into the chosen human-facing interface.\n\n```text\n'
            +packet['prompt']+'\n```\n\n## Local request metadata\n\n'+metadata+'\n')

def registry_schema():
    props = {k: {'type': 'string'} for k in ('candidate_id', 'asset_request_id', 'target_slide_id', 'filename', 'sha256',
             'mime_type', 'imported_at', 'file_creation_time', 'file_modified_time', 'creation_time_basis')}
    props.update({'sha256': {'type':'string','pattern':'^[a-f0-9]{64}$'},
                  'width': {'type':'integer','minimum':1}, 'height': {'type':'integer','minimum':1},
                  'scientific_evidence': {'const':False}, 'presentation_only': {'const':True}, 'human_approved': {'const':False},
                  'approval_status': {'const':'CANDIDATE'}, 'generation_method': {'const':'MANUAL_HUMAN_TRIGGER'},
                  'generator_source': {'const':'USER_DECLARED'}, 'generator_identity': {'type':'null'}})
    return {'$schema':'https://json-schema.org/draft/2020-12/schema', 'type':'object','additionalProperties':False,
            'required':['schema_version','candidates'], 'properties':{'schema_version':{'const':'manual-candidate-registry/1'},
            'candidates':{'type':'array','items':{'type':'object','additionalProperties':False,'required':sorted(props),'properties':props}}}}

def no_reparse(path, *, existing=True):
    path = Path(path).absolute()
    for part in (path, *path.parents):
        if not part.exists() and not part.is_symlink():
            continue
        info = part.lstat()
        check(not stat.S_ISLNK(info.st_mode) and not (getattr(info, 'st_file_attributes', 0) & 0x400), 'CANDIDATE_PATH_ESCAPE')
    check(not existing or path.exists(), 'CANDIDATE_INBOX_MISSING')

def within(path, root):
    no_reparse(root)
    no_reparse(path)
    resolved = Path(path).resolve(strict=True)
    check(resolved.is_relative_to(Path(root).resolve(strict=True)), 'CANDIDATE_PATH_ESCAPE')
    return resolved

def inspect_image(path, root, request_id):
    # Resolve and inspect the path before opening any bytes.
    path=within(path,root); info=path.stat()
    check(path.is_file() and getattr(info,'st_nlink',1)==1,'CANDIDATE_PATH_ESCAPE')
    match=re.fullmatch(re.escape(request_id)+r'__([A-C])(\.[A-Za-z]+)',path.name)
    check(match is not None,'UNKNOWN_ASSET_REQUEST_ID_OR_FILENAME')
    suffix=path.suffix.lower();check(suffix in FORMATS,'UNSUPPORTED_FILE')
    check(0 < info.st_size <= MAX_BYTES,'INVALID_FILE_SIZE')
    data=path.read_bytes()
    check(len(data)==info.st_size,'CANDIDATE_CHANGED_DURING_READ')
    from PIL import Image, UnidentifiedImageError
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('error',Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(data)) as image:
                check(image.format==FORMATS[suffix][0],'MIME_EXTENSION_MISMATCH')
                check(getattr(image,'n_frames',1)==1,'ANIMATED_FILE_UNSUPPORTED')
                width,height=image.size
                check(width>0 and height>0 and width*height<=MAX_PIXELS,'INVALID_IMAGE_DIMENSIONS')
                image.verify()
            with Image.open(io.BytesIO(data)) as image: image.load()
    except (UnidentifiedImageError,OSError,SyntaxError,Image.DecompressionBombError,Image.DecompressionBombWarning) as exc:
        raise BridgeError('CORRUPT_OR_UNSAFE_IMAGE') from exc
    after=path.stat()
    check(info.st_mtime_ns==after.st_mtime_ns and info.st_size==after.st_size,'CANDIDATE_CHANGED_DURING_READ')
    check(sha_file(path)==sha_bytes(data),'CANDIDATE_CHANGED_DURING_READ')
    birth=getattr(info,'st_birthtime',None)
    if birth is None and __import__('os').name=='nt': birth=info.st_ctime
    created=datetime.fromtimestamp(birth,timezone.utc).isoformat() if birth is not None else 'UNAVAILABLE'
    return {'filename':path.name,'sha256':sha_bytes(data),'mime_type':FORMATS[suffix][1], 'width':width,'height':height,
            'file_creation_time':created,'file_modified_time':datetime.fromtimestamp(info.st_mtime,timezone.utc).isoformat(),
            'creation_time_basis':'FILESYSTEM_CREATION_TIME_NOT_GENERATION_TIME' if birth is not None else 'UNAVAILABLE'},data

def discover(inbox, records):
    no_reparse(inbox)
    known={row['asset_request_id']:row for row in records}
    results=[]
    # One directory level only; never follow subdirectories or reparse points.
    for folder in sorted(inbox.iterdir()):
        within(folder,inbox)
        check(folder.is_dir() and folder.name in known,'UNKNOWN_ASSET_REQUEST_ID')
        files=list(folder.iterdir())
        check(len(files)<=3,'TOO_MANY_CANDIDATES')
        slots=set()
        for path in sorted(files):
            meta,data=inspect_image(path,inbox,folder.name)
            slot=path.stem.rsplit('__',1)[-1]
            check(slot not in slots,'DUPLICATE_CANDIDATE_SLOT');slots.add(slot)
            results.append((known[folder.name],meta,data))
    return results

def external_root(path):
    path=Path(path).absolute()
    no_reparse(path,existing=False)
    check(path.resolve()!=ROOT and not path.resolve().is_relative_to(ROOT),'EXTERNAL_ASSET_ROOT_REQUIRED')
    return path.resolve()


def authorized_run(run_id,asset_root):
    check(isinstance(run_id,str) and bool(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,79}',run_id)),'INVALID_RUN_ID')
    return external_root(asset_root)/'candidate_inbox'/run_id


def inbox_schema():
    return {'schema_version':'manual-candidate-inbox/1','run_directory_pattern':'candidate_inbox/<run_id>/<asset_request_id>/',
      'filename_pattern':'<asset_request_id>__[A-C].png|jpg|jpeg|webp','max_candidates_per_request':3,
      'empty_inbox_valid':True,'recursive_scan':False,'links_allowed':False,'max_file_bytes':MAX_BYTES,'max_pixels':MAX_PIXELS}


def prepare(output,run_id,*,records=None,asset_root=None,baseline_pptx=None,enabled=False):
    if not enabled:return None
    output=external_root(output);inbox=authorized_run(run_id,asset_root)
    check(not output.exists() and not inbox.exists(),'OUTPUT_ALREADY_EXISTS')
    check(isinstance(records,list) and bool(records),'CONTROLLED_REQUEST_RECORDS_REQUIRED')
    ids=set();indices=set();packets=[]
    for row in records:
        requests.closed(row,('slide_index','request'),'handoff record')
        i=row['slide_index'];check(type(i) is int and i>0 and i not in indices,'UNIQUE_SLIDE_INDEX_REQUIRED');indices.add(i)
        request=row['request'];requests.validate_request(request)
        check(request['slide_id'] not in ids,'DUPLICATE_SLIDE_ID');ids.add(request['slide_id'])
        packet=make_packet(request,enabled=True)
        if packet:packets.append((i,request,packet))
    baseline=None
    if baseline_pptx is not None:
        baseline_path=Path(baseline_pptx).resolve();check(baseline_path.is_file(),'BASELINE_REQUIRED')
        from .ooxml_qa import inspect_pptx_ooxml
        inspected=inspect_pptx_ooxml(baseline_path)
        check(inspected['valid'] and max(indices)<=len(inspected['slide_order']),'BASELINE_SLIDE_BINDING_INVALID')
        baseline={'path':str(baseline_path),'sha256':sha_file(baseline_path)}
    # Validate the complete batch before writing any output or inbox.
    output.mkdir(parents=True);inbox.mkdir(parents=True);rows=[]
    for i,request,packet in packets:
        folder=output/'generation_packets'/request['slide_id'];folder.mkdir(parents=True)
        (inbox/packet['asset_request_id']).mkdir()
        write_new(folder/'generation_packet.json',packet)
        (folder/'generation_packet.md').write_text(packet_markdown(packet,i),encoding='utf8')
        rows.append({'slide_index':i,'asset_request_id':packet['asset_request_id'],'target_slide_id':packet['target_slide_id'],
          'packet_relative_path':(folder/'generation_packet.json').relative_to(output).as_posix(),
          'packet_sha256':sha_file(folder/'generation_packet.json'),'request':request})
    manifest={'schema_version':'manual-handoff/2','run_id':run_id,'packets':rows,'authorized_inbox':str(inbox),
      'automatic_generation_allowed':False,'network_calls_allowed':False,'human_approval_required':True,'baseline':baseline}
    write_new(output/'manual_handoff_manifest.json',manifest)
    write_new(output/'generation_packet.schema.json',{'oneOf':[{'const':p} for _,_,p in packets]})
    write_new(output/'candidate_inbox_schema.json',inbox_schema());write_new(output/'candidate_asset_registry.schema.json',registry_schema())
    write_new(output/'candidate_asset_registry.json',{'schema_version':'manual-candidate-registry/1','candidates':[]})
    write_new(output/'candidate_intake_report.json',{'status':'NOT_SCANNED','inbox_scanned':False,'approval_inferred':False})
    with (output/'candidate_review_template.csv').open('x',newline='',encoding='utf-8-sig') as stream:csv.DictWriter(stream,fieldnames=REVIEW_FIELDS).writeheader()
    (output/'MANUAL_GENERATION_GUIDE.md').write_text('''# Manual image generation

1. Open your chosen human-facing image-generation interface.
2. Copy only one generation packet prompt block; never upload private project files or notes.
3. Generate 1–3 candidates. Keep scientific labels, numbers and diagrams native in PowerPoint.
4. Save candidates as `<asset_request_id>__A.png`, `__B.png` or `__C.png` in the designated inbox request directory.
5. Run explicit local intake. File presence does not imply approval.
6. Review the local candidate comparison and fill the review CSV yourself.
7. Apply only explicit approvals; assets remain bound to the reviewed deck and target slide.

No automatic image-generation, API, credential handling or network call occurs. Do not place files directly into approved_assets.
''',encoding='utf8')
    return output


def load_package(package,asset_root):
    package=Path(package).resolve();no_reparse(package)
    manifest=read_json(package/'manual_handoff_manifest.json')
    check(manifest.get('schema_version')=='manual-handoff/2' and manifest.get('automatic_generation_allowed') is False
          and manifest.get('network_calls_allowed') is False and manifest.get('human_approval_required') is True,'HANDOFF_AUTHORITY_CHANGED')
    inbox=authorized_run(manifest['run_id'],asset_root)
    check(Path(manifest['authorized_inbox'])==inbox,'CANDIDATE_PATH_ESCAPE')
    ids=set();indices=set()
    for row in manifest['packets']:
        request=row['request'];requests.validate_request(request)
        check(row['slide_index'] not in indices and type(row['slide_index']) is int and row['slide_index']>0,'DUPLICATE_SLIDE_INDEX');indices.add(row['slide_index'])
        check(request['request_id'] not in ids,'DUPLICATE_REQUEST');ids.add(request['request_id'])
        expected=f"generation_packets/{request['slide_id']}/generation_packet.json"
        check(row['packet_relative_path']==expected,'PACKET_PATH_ESCAPE')
        path=within(package/expected,package);check(sha_file(path)==row['packet_sha256'],'PACKET_HASH_MISMATCH')
        packet=read_json(path);validate_packet(packet,request)
        check(row['target_slide_id']==request['slide_id'] and row['asset_request_id']==request['request_id'],'PACKET_BINDING_MISMATCH')
    baseline=manifest.get('baseline')
    if baseline:check(sha_file(baseline['path'])==baseline['sha256'],'BASELINE_HASH_MISMATCH')
    return manifest


def intake(package,output,*,asset_root=None,enabled=False,import_image_candidates=False):
    if not enabled or not import_image_candidates:return None
    manifest=load_package(package,asset_root);output=Path(output).absolute()
    expected=external_root(asset_root)/'candidate_runs'/manifest['run_id']
    check(output.parent==expected and bool(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,79}',output.name)),'AUTHORIZED_EXTERNAL_OUTPUT_REQUIRED')
    no_reparse(output,existing=False);check(not output.exists(),'OUTPUT_ALREADY_EXISTS')
    found=discover(authorized_run(manifest['run_id'],asset_root),manifest['packets'])
    check(not found or manifest.get('baseline') is not None,'BASELINE_REQUIRED_FOR_LOCAL_REVIEW')
    imported_at=now();entries=[]
    for row,meta,data in found:
        entries.append({'candidate_id':'CAN-'+sha_bytes((row['asset_request_id']+'|'+meta['sha256']).encode())[:24].upper(),
         'asset_request_id':row['asset_request_id'],'target_slide_id':row['target_slide_id'],**meta,
         'scientific_evidence':False,'presentation_only':True,'human_approved':False,'approval_status':'CANDIDATE',
         'generation_method':'MANUAL_HUMAN_TRIGGER','generator_source':'USER_DECLARED','generator_identity':None,'imported_at':imported_at})
    check(len({r['candidate_id'] for r in entries})==len(entries),'DUPLICATE_CANDIDATE_CONTENT')
    output.mkdir(parents=True);(output/'assets').mkdir();(output/'render_inputs').mkdir()
    for entry,(_,_,data) in zip(entries,found):
        (output/'assets'/entry['filename']).write_bytes(data)
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:im.convert('RGBA').save(output/'render_inputs'/(entry['candidate_id']+'.png'))
    write_new(output/'candidate_asset_registry.json',{'schema_version':'manual-candidate-registry/1','candidates':entries})
    with (output/'candidate_review_template.csv').open('x',newline='',encoding='utf-8-sig') as stream:
        writer=csv.DictWriter(stream,fieldnames=REVIEW_FIELDS);writer.writeheader()
        for entry in entries:writer.writerow({k:entry[k] for k in REVIEW_FIELDS[:3]})
    write_new(output/'intake_binding.json',{'baseline':manifest.get('baseline'),'packets':manifest['packets'],
      'handoff_sha256':sha_file(Path(package)/'manual_handoff_manifest.json'),'registry_sha256':sha_file(output/'candidate_asset_registry.json')})
    if entries:
        from .candidate_preview import create_preview
        create_preview(output,enabled=True)
    report={'status':'CANDIDATES_IMPORTED_UNAPPROVED' if entries else 'EMPTY_INBOX_VALID','inbox_scanned':True,
      'candidates_registered':len(entries),'approval_inferred':False,'scientific_source_registration':False,'formal_deck_impact':'NONE',
      'pixel_content_safety':'NOT_ASSESSED_HUMAN_REVIEW_REQUIRED','imported_at':imported_at}
    write_new(output/'candidate_intake_report.json',report)
    return output


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--manual-image-handoff',action='store_true')
    parser.add_argument('--import-image-candidates',action='store_true')
    parser.add_argument('--records',type=Path);parser.add_argument('--package',type=Path);parser.add_argument('--asset-root',type=Path)
    parser.add_argument('--output',type=Path);parser.add_argument('--run-id');parser.add_argument('--baseline',type=Path)
    args=parser.parse_args()
    if not args.manual_image_handoff:return
    if not args.asset_root or not args.output:parser.error('--asset-root and --output are required')
    if args.import_image_candidates:
        if not args.package:parser.error('--package is required for intake')
        result=intake(args.package,args.output,asset_root=args.asset_root,enabled=True,import_image_candidates=True)
    else:
        if not args.records or not args.run_id:parser.error('--records and --run-id are required for packets')
        result=prepare(args.output,args.run_id,records=read_json(args.records),asset_root=args.asset_root,baseline_pptx=args.baseline,enabled=True)
    print(str(result))

if __name__=='__main__':main()
