"""Image-only final assembly must preserve exact native scientific package parts."""
import json, shutil, unittest, zipfile
from unittest.mock import patch
from xml.etree import ElementTree as E
from tests import test_asset_approval as approval_tests
from tests.test_manual_image_handoff import synthetic_render
from scripts.academic_ppt import visual_workflow as w
from scripts.academic_ppt import asset_approval as a


class AssemblyAuditTests(unittest.TestCase):
    def setUp(self):
        fixture=approval_tests.AssetApprovalTests();fixture.setUp();self.addCleanup(fixture.doCleanups)
        fixture.decide();fixture.promote();self.root=fixture.root
        self.source=self.root/'source_run';self.source.mkdir();self.stage=self.root/'stage';self.stage.mkdir()
        self.baseline=self.source/'baseline.pptx';shutil.copyfile(fixture.fixture.baseline,self.baseline)
        self.assembled=self.root/'assembled.pptx';self.output=self.root/'audited'
        self.ids=[f'S{i:03d}' for i in range(1,16)]
        self.receipt=a.assemble_approved_assets(self.baseline,self.assembled,fixture.destination/'asset_registry.yaml',self.ids,enabled=True)
        w.write_json(self.assembled.with_suffix('.assembly.json'),self.receipt)
        w.write_json(self.source/'deck_ir.json',{'canonical_hash':'c'*64,'slides':[{'slide_id':sid} for sid in self.ids]})
        w.write_json(self.source/'visual_review_context.json',{'deck_sha256':w.sha_file(self.baseline),'effective_staging_root':str(self.stage),'quality':'validated'})
        w.write_json(self.stage/'visual_qa_receipt.json',{'deck_sha256':w.sha_file(self.baseline),'scientific':'PASS'})
        w.write_json(self.stage/'master_binding_manifest.json',{'output_sha256':w.sha_file(self.baseline),'status':'VALID'})

    def audit(self):return w.audit_assembly(self.source,self.assembled,self.output)

    def mutate(self,part,change):
        with zipfile.ZipFile(self.assembled) as z:parts={n:z.read(n) for n in z.namelist()}
        parts[part]=change(parts[part])
        with zipfile.ZipFile(self.assembled,'w',zipfile.ZIP_DEFLATED) as z:
            for name,data in parts.items():z.writestr(name,data)
        self.receipt['output_sha256']=w.sha_file(self.assembled)
        w.write_json(self.assembled.with_suffix('.assembly.json'),self.receipt)

    def test_new_final_render_and_review_required(self):
        with patch('scripts.academic_ppt.rendering.render_pptx',side_effect=synthetic_render),patch('scripts.academic_ppt.visual_layout_qa.inspect_powerpoint_text_layout',return_value=({'status':'PASS','slides':[]},'')),patch.object(w,'finalize_run',return_value={'status':'HUMAN_REVIEW'}) as gate:
            self.assertEqual(self.audit()['status'],'HUMAN_REVIEW')
        self.assertIsNone(gate.call_args.args[1].visual_evidence)
        self.assertEqual(w.sha_file(self.output/'deck.pptx'),w.sha_file(self.assembled))  # self-containment: generated-temp
        qa=w.read_json(self.output/'assembly_evidence/visual_qa_receipt.json')  # self-containment: generated-temp
        self.assertEqual(qa['scientific'],'PASS');self.assertEqual(qa['deck_sha256'],w.sha_file(self.assembled))

    def test_changed_scientific_text_rejected_even_with_new_outer_hash(self):
        self.mutate('ppt/slides/slide1.xml',lambda b:b.replace(b'Synthetic source 1',b'Altered source 1'))
        with self.assertRaisesRegex(ValueError,'SCIENTIFIC_INVARIANCE'):self.audit()
        self.assertFalse(self.output.exists())

    def test_changed_native_shape_rejected(self):
        def change(data):
            root=E.fromstring(data)
            from scripts.academic_ppt.master_layout import NS,xml
            root.find('.//a:xfrm/a:off',NS).set('x','999')
            return xml(root)
        self.mutate('ppt/slides/slide1.xml',change)
        with self.assertRaisesRegex(ValueError,'NATIVE_OBJECT_CHANGED'):self.audit()

    def test_content_type_mutation_rejected(self):
        self.mutate('[Content_Types].xml',lambda b:b.replace(b'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',b'application/x-synthetic'))
        with self.assertRaisesRegex(ValueError,'CONTENT_TYPE_CHANGED'):self.audit()

    def test_unapproved_package_mutation_rejected(self):
        self.mutate('docProps/core.xml',lambda b:b.replace(b'</cp:coreProperties>',b'<!-- changed --></cp:coreProperties>'))
        with self.assertRaisesRegex(ValueError,'UNAUTHORIZED_PACKAGE_MUTATION'):self.audit()

    def test_prior_scientific_qa_not_assumed(self):
        w.write_json(self.stage/'visual_qa_receipt.json',{'deck_sha256':w.sha_file(self.baseline),'scientific':'NOT_ASSESSED'})
        with self.assertRaisesRegex(ValueError,'BASELINE_SCIENTIFIC_QA_REQUIRED'):self.audit()


if __name__=='__main__':unittest.main()
