"""Explicit human approval and native insertion governance on synthetic data."""
import copy,csv,json,unittest
from pathlib import Path
from tests import test_manual_image_handoff as handoff_tests
from scripts.academic_ppt import asset_approval as a
from scripts.academic_ppt import manual_image_handoff as m


class AssetApprovalTests(unittest.TestCase):
    def setUp(self):
        self.fixture=handoff_tests.ManualHandoffTests();self.fixture.setUp();self.addCleanup(self.fixture.doCleanups)
        self.fixture.candidate();self.fixture.intake();self.root=self.fixture.root;self.imported=self.fixture.out
        self.csv=self.imported/'candidate_review_template.csv';self.destination=self.root/'approved'
    def decide(self,decision='APPROVE',approve='true'):
        with self.csv.open(encoding='utf-8-sig',newline='') as stream:rows=list(csv.DictReader(stream))
        rows[0].update(decision=decision,approve=approve)
        with self.csv.open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=m.REVIEW_FIELDS);writer.writeheader();writer.writerows(rows)
    def promote(self):return a.apply_approvals(self.imported,self.csv,self.destination,reviewer='SYNTHETIC_REVIEWER',enabled=True)
    def test_default_off_no_inputs(self):self.assertIsNone(a.apply_approvals(None,None,None))
    def test_blank_review_never_approves(self):
        self.assertEqual(self.promote()['approved_count'],0);self.assertFalse(self.destination.exists())
    def test_file_presence_not_approval(self):
        self.decide('', 'true');self.assertEqual(self.promote()['approved_count'],0)
    def test_decision_without_approve_not_approval(self):
        self.decide('APPROVE','');self.assertEqual(self.promote()['approved_count'],0)
    def test_refinement_is_not_final_approval(self):
        self.decide('APPROVE_WITH_REFINEMENT');self.assertEqual(self.promote()['approved_count'],0)
    def test_rejection_is_not_approval(self):
        self.decide('REJECT');self.assertEqual(self.promote()['approved_count'],0)
    def test_human_identity_required(self):
        self.decide()
        with self.assertRaisesRegex(ValueError,'EXPLICIT_HUMAN_REVIEWER'):a.apply_approvals(self.imported,self.csv,self.destination,enabled=True)
    def test_exact_human_approval_preserves_source_registry(self):
        before=m.sha_file(self.imported/'candidate_asset_registry.json');self.decide();report=self.promote();self.assertEqual(report['approved_count'],1)
        self.assertEqual(before,m.sha_file(self.imported/'candidate_asset_registry.json'))
        asset=m.read_json(self.destination/'asset_registry.yaml')['assets'][0];self.assertTrue(asset['human_approved']);self.assertFalse(asset['scientific_evidence'])
        self.assertEqual(asset['slide_ids'],['S001']);self.assertEqual(asset['sha256'],m.sha_file(self.destination/asset['file']))
    def test_changed_preview_cannot_be_approved(self):
        self.decide();review=m.read_json(self.imported/'review_manifest.json');(self.imported/'review'/review['entries'][0]['preview_filename']).write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError,'REVIEW_PREVIEW_CHANGED'):self.promote()
    def test_changed_asset_cannot_be_approved(self):
        self.decide();c=m.read_json(self.imported/'candidate_asset_registry.json')['candidates'][0];(self.imported/'assets'/c['filename']).write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError,'CANDIDATE_HASH_MISMATCH'):self.promote()
    def test_wrong_slide_use_rejected(self):
        self.decide();self.promote();asset=m.read_json(self.destination/'asset_registry.yaml')['assets'][0]
        with self.assertRaisesRegex(ValueError,'SCOPE_MISMATCH'):a.validate_use(asset,self.destination,slide_id='S002',baseline_sha256=asset['baseline_sha256'])
    def test_changed_deck_use_rejected(self):
        self.decide();self.promote();asset=m.read_json(self.destination/'asset_registry.yaml')['assets'][0]
        with self.assertRaisesRegex(ValueError,'SCOPE_MISMATCH'):a.validate_use(asset,self.destination,slide_id='S001',baseline_sha256='0'*64)
    def test_approved_bytes_changed_rejected(self):
        self.decide();self.promote();asset=m.read_json(self.destination/'asset_registry.yaml')['assets'][0];(self.destination/asset['file']).write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError,'APPROVED_ASSET_HASH_MISMATCH'):a.validate_use(asset,self.destination,slide_id='S001',baseline_sha256=asset['baseline_sha256'])
    def test_existing_approved_batch_not_overwritten(self):
        self.decide();self.promote()
        with self.assertRaisesRegex(ValueError,'APPROVED_OUTPUT_ALREADY_EXISTS'):self.promote()
    def test_assembly_default_off(self):self.assertIsNone(a.assemble_approved_assets(None,None,None,None))
    def test_native_science_preserved_after_insertion(self):
        self.decide();self.promote();baseline=self.fixture.baseline;before=m.sha_file(baseline)
        report=a.assemble_approved_assets(baseline,self.root/'assembled.pptx',self.destination/'asset_registry.yaml',[f'S{i:03d}' for i in range(1,16)],enabled=True)
        self.assertTrue(report['scientific_invariance']);self.assertEqual(before,m.sha_file(baseline))
        import zipfile
        with zipfile.ZipFile(self.root/'assembled.pptx') as z:
            self.assertIn(b"encoding='utf-8'",z.read('[Content_Types].xml').splitlines()[0])
    def test_asset_cannot_cover_native_scientific_content(self):
        self.decide();self.promote();reg=self.destination/'asset_registry.yaml';doc=m.read_json(reg);doc['assets'][0]['focal_region']={'x':.01,'y':.01,'w':.5,'h':.5};reg.write_text(json.dumps(doc))
        with self.assertRaisesRegex(ValueError,'AUTHORIZATION_CHANGED'):a.assemble_approved_assets(self.fixture.baseline,self.root/'blocked.pptx',reg,[f'S{i:03d}' for i in range(1,16)],enabled=True)

if __name__=='__main__':unittest.main()
