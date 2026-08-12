from __future__ import annotations

import csv
import hashlib
import io
import json
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

import yaml

from scripts.validate_release_bundle import contains_private_absolute_path
from scripts.build_release_bundle import VERSION

ROOT = Path(__file__).resolve().parents[1]

class VisualWorkspaceTests(unittest.TestCase):
    def test_six_workspaces_have_required_contracts(self):
        root=ROOT/"visual_workspaces"; dirs=sorted(p for p in root.iterdir() if p.is_dir())
        self.assertEqual(len(dirs),6)
        required={"workspace.yaml","art_direction.yaml","master_roles.yaml","layout_variants.yaml","chart_styles.yaml","annotation_rules.yaml","example_storyboard.yaml","synthetic_preview.pptx","human_review.csv"}
        for directory in dirs:
            self.assertTrue(required <= {p.name for p in directory.iterdir()},directory)
            with zipfile.ZipFile(directory/"synthetic_preview.pptx") as zf:
                slides=[n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
                self.assertEqual(len(slides),10)

    def test_human_review_is_not_auto_filled(self):
        for path in (ROOT/"visual_workspaces").glob("*/human_review.csv"):
            with path.open(encoding="utf-8") as handle:
                row=next(csv.DictReader(handle))
            self.assertEqual(row["status"],"PENDING")
            self.assertEqual(row["total_score"],"")

class PortabilityTests(unittest.TestCase):
    def test_release_version_contract_is_consistent(self):
        expected = VERSION.removeprefix("v")
        package = json.loads((ROOT / "package.json").read_text(encoding="utf-8"))
        package_lock = json.loads((ROOT / "package-lock.json").read_text(encoding="utf-8"))
        dependency_lock = json.loads((ROOT / "config/dependency_lock.json").read_text(encoding="utf-8"))
        self.assertEqual(package["version"], expected)
        self.assertEqual(package_lock["version"], expected)
        self.assertEqual(package_lock["packages"][""]["version"], expected)
        self.assertEqual(dependency_lock["schema_version"], expected)
        self.assertEqual(
            dependency_lock["node"]["default_backend"]["package_json_sha256"].lower(),
            hashlib.sha256((ROOT / "package.json").read_bytes()).hexdigest(),
        )

    def test_release_scope_contract_lists_runtime_dependencies_and_private_exclusions(self):
        rules = yaml.safe_load((ROOT / "release_exclusion_rules.yaml").read_text(encoding="utf-8"))
        self.assertTrue(
            {"extensions/**", "prompts/**", "visual_templates/**", "templates/**"}
            <= set(rules["include"])
        )
        self.assertTrue(
            {"audit/**", "benchmark/**", "private/**", "release/**", "docs/*.docx"}
            <= set(rules["exclude"])
        )

    def test_local_config_is_ignored(self):
        result=subprocess.run(["git","check-ignore","config/local.yaml"],cwd=ROOT,text=True,capture_output=True)
        self.assertEqual(result.returncode,0,result.stderr)

    def test_production_scope_has_no_fixed_repository_drive(self):
        paths=[ROOT/"run_ppt_workflow.py",ROOT/"scripts",ROOT/"config"]
        bad=[]
        for base in paths:
            candidates=[base] if base.is_file() else base.rglob("*")
            for path in candidates:
                if path.is_file() and path.suffix.lower() in {".py",".ps1",".mjs",".json",".yaml",".yml"}:
                    text=path.read_text(encoding="utf-8",errors="replace").lower()
                    if ("".join(("e:", "/ppt")) in text or "".join(("e:", "\\ppt")) in text): bad.append(str(path.relative_to(ROOT)))
        self.assertEqual(bad,[])

    def test_release_builder_excludes_private_roots(self):
        with tempfile.TemporaryDirectory() as raw:
            out=Path(raw)
            subprocess.run([__import__('sys').executable,str(ROOT/"scripts/build_release_bundle.py"),"--output-root",str(out)],cwd=ROOT,check=True)
            bundle=out/"academic-ppt-workflow-v2.7.0-rc1.zip"
            with zipfile.ZipFile(bundle) as zf:
                names=[n.split("/",1)[1] for n in zf.namelist() if "/" in n]
            self.assertFalse(any(n.startswith(("input/","output/","staging/","audit/","benchmark/","private/","release/",".cache/")) for n in names))
            self.assertFalse(any(n.startswith("docs/") and n.endswith(".docx") for n in names))
            self.assertTrue(any(n.startswith("extensions/") for n in names))
            self.assertTrue(any(n.startswith("prompts/") for n in names))
            self.assertTrue(any(n.startswith("visual_templates/") for n in names))
            self.assertTrue(any(n.startswith("templates/") for n in names))

    def test_bundle_validator_detects_json_escaped_private_path(self):
        separator = chr(92)
        private_path = separator.join(("E:", "PPT", "input", "figure.png"))
        payload = ('{"visual_asset_path":"' + private_path.replace(separator, separator * 2) + '"}').encode()
        self.assertTrue(contains_private_absolute_path(payload, ".json"))

    def test_bundle_validator_detects_private_path_inside_office_package(self):
        separator = chr(92)
        private_path = separator.join(("E:", "PPT", "input", "figure.png"))
        payload = io.BytesIO()
        with zipfile.ZipFile(payload, "w") as package:
            package.writestr("ppt/slides/slide1.xml", f'<p:cNvPr descr="{private_path}"/>')
        self.assertTrue(contains_private_absolute_path(payload.getvalue(), ".pptx"))

if __name__=="__main__": unittest.main()
