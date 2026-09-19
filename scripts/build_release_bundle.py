from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import os
import shutil
import subprocess
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import yaml

# One checked-in selection authority, also used by privacy auditing.
RULES_PATH = Path(__file__).resolve().parents[1] / "release_exclusion_rules.yaml"
RULES = yaml.safe_load(RULES_PATH.read_text(encoding="utf-8"))
INCLUDE, EXCLUDE = RULES["include"], RULES["exclude"]
VERSION = "v" + json.loads((RULES_PATH.parent / "package.json").read_text(encoding="utf-8"))["version"]


def match(path: str, patterns: list[str]) -> bool:
    candidates=[variant for pattern in patterns for variant in ([pattern,pattern[3:]] if pattern.startswith('**/') else [pattern])]
    return any(fnmatch.fnmatch(path, p) or (p.endswith("/**") and path.startswith(p[:-3] + "/")) for p in candidates)

def selected(repo: Path) -> list[Path]:
    rows=[]
    for folder,dirs,files in os.walk(repo,followlinks=False):
        base=Path(folder)
        dirs[:]=[name for name in dirs if not match((base/name).relative_to(repo).as_posix()+"/",EXCLUDE)]
        for name in files:
            path=base/name;rel=path.relative_to(repo).as_posix()
            if not match(rel,EXCLUDE) and match(rel,INCLUDE):rows.append(path)
    return sorted(rows, key=lambda p:p.relative_to(repo).as_posix().lower())

def digest(path: Path) -> str:
    h=hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda:f.read(1024*1024),b""):h.update(block)
    return h.hexdigest()

def git_value(repo: Path, *arguments: str) -> str:
    completed=subprocess.run(
        ["git", *arguments], cwd=repo, capture_output=True, text=True, check=False
    )
    if completed.returncode != 0 or not completed.stdout.strip():
        return "NOT_AVAILABLE"
    return completed.stdout.strip()

def build(repo: Path, output_root: Path) -> dict:
    if repo == output_root or repo in output_root.parents:
        raise ValueError("Release dry-run output must be outside the repository")
    output_root.mkdir(parents=True,exist_ok=True)
    files=selected(repo); bundle=output_root/f"academic-ppt-workflow-{VERSION}.zip"; prefix=f"academic-ppt-workflow-{VERSION}"
    manifest_files=[]
    with zipfile.ZipFile(bundle,"w",compression=zipfile.ZIP_DEFLATED,compresslevel=9) as zf:
        for path in files:
            rel=path.relative_to(repo).as_posix(); zf.write(path,f"{prefix}/{rel}"); manifest_files.append({"path":rel,"size":path.stat().st_size,"sha256":digest(path)})
    bootstrap=output_root/"bootstrap_windows.ps1"; shutil.copy2(repo/"scripts/bootstrap_windows.ps1",bootstrap)
    package_lock=repo/"package-lock.json"; python_lock=repo/"requirements-lock.txt"
    manifest={"schema_version":"academic-ppt-release-manifest/2","version":VERSION,"release_type":"release_candidate","status":"RELEASE_CANDIDATE_NOT_PUBLISHED","repository":"secdelic/academic-ppt-workflow","candidate_branch":git_value(repo,"branch","--show-current"),"candidate_commit":git_value(repo,"rev-parse","HEAD"),"build_timestamp":datetime.now(timezone.utc).isoformat(),"python_lock_hash":digest(python_lock),"npm_lock_hash":digest(package_lock),"workflow_schema_version":"2.7","cache_schema_version":"2","default_backend":"native_pptxgenjs","external_skill_status":"NO_GO_EXTERNAL_SKILL","file_count":len(files),"files":manifest_files,"archive_sha256":digest(bundle),"bundle":{"name":bundle.name,"size":bundle.stat().st_size,"sha256":digest(bundle)},"gates":{"case_b":"NOT_RUN","case_c":"NOT_RUN","second_device":"PENDING_RC4_CLEAN_INSTALL","human_workspace_scores":"PENDING","github_actions":"NOT_RUN"}}
    (output_root/"release_manifest.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    checksum_rows=[]
    for path in sorted([bundle,bootstrap,output_root/"release_manifest.json"]):checksum_rows.append(f"{digest(path)}  {path.name}")
    (output_root/"SHA256SUMS.txt").write_text("\n".join(checksum_rows)+"\n",encoding="utf-8")
    shutil.copy2(repo/"docs/中文操作说明书.md",output_root/"中文操作说明书.md"); shutil.copy2(repo/"CHANGELOG.md",output_root/"CHANGELOG.md")
    return manifest

def main()->int:
    p=argparse.ArgumentParser();p.add_argument("--repo",default=Path(__file__).resolve().parents[1]);p.add_argument("--output-root",required=True);a=p.parse_args();result=build(Path(a.repo).resolve(),Path(a.output_root).resolve());print(json.dumps(result["bundle"],indent=2));return 0
if __name__=="__main__":raise SystemExit(main())
