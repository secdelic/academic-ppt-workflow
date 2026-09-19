"""The shipped workflow has one current manual and no removed runtime imports."""
import re, unittest
from pathlib import Path
from urllib.parse import unquote
from scripts.build_release_bundle import selected, match, EXCLUDE

ROOT=Path(__file__).resolve().parents[1]


class ConsolidatedReleaseTests(unittest.TestCase):
    def test_local_markdown_links_exist_in_bundle(self):
        candidates=selected(ROOT);included={p.resolve() for p in candidates}
        for path in candidates:
            if path.suffix.lower()!='.md':continue
            for raw in re.findall(r'!?\[[^\]]*\]\(([^)]+)\)',path.read_text(encoding='utf-8')):
                target=raw.strip().strip('<>').split('#',1)[0]
                if not target or re.match(r'[a-z][a-z0-9+.-]*:',target,re.I):continue
                resolved=(path.parent/unquote(target)).resolve()
                self.assertTrue(resolved.exists(),f'{path.relative_to(ROOT)} -> {target}')
                if resolved.is_file():self.assertIn(resolved,included,f'Link target missing from release: {target}')

    def test_current_documentation_authority(self):
        self.assertTrue((ROOT/'docs/中文操作说明书.md').is_file())
        self.assertIn('docs/中文操作说明书.md',(ROOT/'README.md').read_text(encoding='utf-8'))
        self.assertFalse((ROOT/'README_FIRST_INSTALL.md').exists())

    def test_runtime_has_no_experiment_import(self):
        for path in selected(ROOT):
            if not path.relative_to(ROOT).as_posix().startswith('scripts/') or path.suffix not in ('.py','.mjs'):continue
            text=path.read_text(encoding='utf-8')
            self.assertNotRegex(text,r'(?:from|import|require).*experiments[/.]')

    def test_private_runtime_directories_excluded(self):
        for path in ('output/run/deck.pptx','staging/run/data.json','audit/run/review.json','candidate_inbox/run/image.png','approved_assets/image.png','.env','config/local.yaml'):  # self-containment: documentation-only
            self.assertTrue(match(path,EXCLUDE),path)


if __name__=='__main__':unittest.main()
