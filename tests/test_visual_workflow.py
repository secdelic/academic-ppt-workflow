"""Production visual gates over explicit synthetic, source-bound receipts."""
import argparse,copy,json,shutil,unittest
from pathlib import Path
from tests import test_visual_quality_floor as floor_tests
from tests.test_visual_ambition import fixture as ambition_fixture
from scripts.academic_ppt import visual_workflow as w
from scripts.academic_ppt.utils import write_json


class VisualWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.fixture=floor_tests.QualityFloorTests();self.fixture.setUp();self.addCleanup(self.fixture.doCleanups)
        self.root=self.fixture.root;self.out=self.root/'run';self.out.mkdir();self.stage=self.root/'stage';self.stage.mkdir()
        self.manifest=self.fixture.manifest;shutil.copyfile(self.root/'synthetic.pptx',self.out/'deck.pptx');shutil.copyfile(self.root/'synthetic.pdf',self.out/'deck.pdf')
        (self.out/'preview').mkdir()
        for i,ref in enumerate(self.manifest['render']['previews'],1):shutil.copyfile(ref['path'],self.out/'preview'/f'slide_{i:03d}.png')
        shutil.copyfile(self.root/'sheet.png',self.out/'preview'/'contact_sheet.png')
        shutil.copyfile(self.root/'qa.json',self.stage/'visual_qa_receipt.json')
        self.ir={'canonical_hash':'c'*64,'deck_id':'SYNTHETIC_ONLY','slides':[{'slide_id':r['slide_id']} for r in self.manifest['records']]}
        write_json(self.out/'deck_ir.json',self.ir)
        self.plan={'schema_version':'assisted-visual-plan/1','deck_content_sha256':'c'*64,'approval':{'status':'approved','reviewer':'SYNTHETIC_ONLY','reviewed_at':'2026-01-01'},'slides':copy.deepcopy(self.manifest['records'])}
        self.planfile=self.root/'plan.json';write_json(self.planfile,self.plan)
        self.extra={'deck_sha256':w.sha_file(self.out/'deck.pptx'),'observations':self.manifest['observations']}
        self.extra_file=self.root/'extra.json';write_json(self.extra_file,self.extra)
        write_json(self.stage/'master_binding_manifest.json',{'status':'VALID','output_sha256':self.extra['deck_sha256'],'bindings':self.plan['slides']})
        self.args=argparse.Namespace(quality='validated',route='generate',visual_plan=self.planfile,visual_evidence=self.extra_file,visual_brief_path=self.root/'brief.json',effective_staging_root=self.stage)

    def run_gate(self):return w.finalize_run(self.out,self.args)
    def save_plan(self):write_json(self.planfile,self.plan)
    def test_quick_not_certified(self):
        self.args.quality='quick';self.assertEqual(self.run_gate()['status'],'DRAFT_LIMITED_QA');self.assertFalse((self.out/'visual_evidence').exists())
    def test_template_profile_not_deck_certified(self):
        self.args.route='create-style-profile';self.assertFalse(self.run_gate()['floor_certified'])
    def test_unsigned_human_reviews_remain_pending(self):self.assertEqual(self.run_gate()['status'],'HUMAN_REVIEW')
    def test_missing_plan_blocks(self):
        self.planfile.unlink();self.assertIn('APPROVED_VISUAL_PLAN_REQUIRED',self.run_gate()['reasons'])
    def test_unapproved_plan_rejected(self):
        self.plan['approval']['status']='pending';self.save_plan()
        with self.assertRaisesRegex(ValueError,'APPROVED_VISUAL_PLAN'):self.run_gate()
    def test_scientific_binding_changed_rejected(self):
        self.plan['deck_content_sha256']='a'*64;self.save_plan()
        with self.assertRaisesRegex(ValueError,'SCIENTIFIC_BINDING_CHANGED'):self.run_gate()
    def test_missing_master_blocks(self):
        (self.stage/'master_binding_manifest.json').unlink();self.assertIn('MASTER_BINDING_MISSING_OR_UNVERIFIED',self.run_gate()['reasons'])
    def test_master_receipt_wrong_deck_blocks(self):
        p=self.stage/'master_binding_manifest.json';d=w.read_json(p);d['output_sha256']='a'*64;write_json(p,d)
        self.assertIn('MASTER_BINDING_MISSING_OR_UNVERIFIED',self.run_gate()['reasons'])
    def test_mere_plan_does_not_prove_rendered_composition(self):
        self.args.visual_evidence=None;self.assertEqual(self.run_gate()['status'],'BLOCK')
    def test_receipt_wrong_deck_rejected(self):
        self.extra['deck_sha256']='a'*64;write_json(self.extra_file,self.extra)
        with self.assertRaisesRegex(ValueError,'VISUAL_EVIDENCE_DECK_CHANGED'):self.run_gate()
    def test_render_bytes_changed_blocks(self):
        (self.out/'preview'/'slide_001.png').write_bytes(b'corrupt');self.assertEqual(self.run_gate()['status'],'BLOCK')
    def test_same_deck_signed_review_can_pass_without_regeneration(self):
        initial=self.run_gate();self.assertEqual(initial['status'],'HUMAN_REVIEW')
        self.fixture.approve_synthetic_human_reviews();self.extra['human_reviews']=self.manifest['human_reviews'];write_json(self.extra_file,self.extra)
        target=self.root/'reviewed';gate=w.review_run(self.out,target,visual_evidence=self.extra_file)
        self.assertEqual(gate['status'],'PASS');self.assertEqual(w.sha_file(target/'deck.pptx'),w.sha_file(self.out/'deck.pptx'))
    def test_existing_review_not_overwritten(self):
        self.run_gate()
        with self.assertRaisesRegex(ValueError,'NEW_REVIEW_OUTPUT'):w.review_run(self.out,self.out,visual_evidence=self.extra_file)
    def test_changed_review_deck_rejected(self):
        self.run_gate();(self.out/'deck.pptx').write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError,'REVIEW_DECK_CHANGED'):w.review_run(self.out,self.root/'reviewed',visual_evidence=self.extra_file)
    def test_full_missing_context_blocks(self):
        self.args.quality='full';gate=self.run_gate();self.assertIn('ANCHOR_OPPORTUNITY_CONTEXT_REQUIRED',gate['reasons'])
        self.assertTrue((self.out/'visual_evidence'/'deep_final_audit.json').is_file())
    def test_full_ambition_requires_independent_human_review(self):
        self.args.quality='full'
        contexts=ambition_fixture(4)['contexts']
        for row,context in zip(self.plan['slides'],contexts):row['ambition_context']=context
        self.save_plan();brief=w.read_json(self.root/'brief.json');brief['direction']['direction_name']='SYNTHETIC';write_json(self.root/'brief.json',brief)
        self.fixture.approve_synthetic_human_reviews();self.extra['human_reviews']=self.manifest['human_reviews'];write_json(self.extra_file,self.extra)
        gate=self.run_gate();self.assertEqual(gate['status'],'HUMAN_REVIEW');self.assertTrue(gate['ambition_human_review_pending'])
        self.assertTrue((self.out/'visual_evidence'/'ambition_review_template.json').is_file())
    def test_plan_template_never_approves(self):
        template=w.plan_template(self.ir);self.assertEqual(template['approval']['status'],'pending');self.assertIsNone(template['approval']['reviewer'])

if __name__=='__main__':unittest.main()
