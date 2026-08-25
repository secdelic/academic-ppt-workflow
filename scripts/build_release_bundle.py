from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path

VERSION = "v2.7.0-rc2"
INCLUDE = [
    "run_ppt_workflow.py", "pyproject.toml", "package.json", "package-lock.json", "requirements-lock.txt",
    "README.md", "README_使用说明.md", "README_中文使用说明.md",
    "README_FIRST_INSTALL.md", "CHANGELOG.md", ".gitignore", ".gitattributes",
    "privacy_allowlist.yaml", "release_exclusion_rules.yaml", "config/**", "scripts/**",
    "docs/**", "design_system/**", "visual_workspaces/**", "visual_templates/**",
    "extensions/**", "prompts/**", "templates/**", "tests/**", "regression/**",
    "governance/**", ".github/**", "AGENTS.md",
]
EXCLUDE = [
    "config/local.yaml", "**/__pycache__/**", "**/*.pyc", "**/~$*", "**/*.tmp",
    "input/**", "output/**", "staging/**", "audit/**", "benchmark/**", "archive/**",
    "private/**", "runtime_logs/**", "install_receipts/**", "docs/*.docx",
    ".cache/**", ".venv/**", "node_modules/**", ".git/**", "release/**",
]

def match(path: str, patterns: list[str]) -> bool:
    return any(fnmatch.fnmatch(path, p) or (p.endswith("/**") and path.startswith(p[:-3] + "/")) for p in patterns)

def selected(repo: Path) -> list[Path]:
    rows=[]
    for path in repo.rglob("*"):
        if not path.is_file(): continue
        rel=path.relative_to(repo).as_posix()
        if match(rel, EXCLUDE): continue
        if match(rel, INCLUDE): rows.append(path)
    return sorted(rows, key=lambda p:p.relative_to(repo).as_posix().lower())

def digest(path: Path) -> str:
    h=hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda:f.read(1024*1024),b""):h.update(block)
    return h.hexdigest()

def build(repo: Path, output_root: Path) -> dict:
    output_root.mkdir(parents=True,exist_ok=True)
    files=selected(repo); bundle=output_root/f"academic-ppt-workflow-{VERSION}.zip"; prefix=f"academic-ppt-workflow-{VERSION}"
    manifest_files=[]
    with zipfile.ZipFile(bundle,"w",compression=zipfile.ZIP_DEFLATED,compresslevel=9) as zf:
        for path in files:
            rel=path.relative_to(repo).as_posix(); zf.write(path,f"{prefix}/{rel}"); manifest_files.append({"path":rel,"size":path.stat().st_size,"sha256":digest(path)})
    bootstrap=output_root/"bootstrap_windows.ps1"; shutil.copy2(repo/"scripts/bootstrap_windows.ps1",bootstrap)
    manifest={"schema_version":"academic-ppt-release-manifest/1","version":VERSION,"status":"RELEASE_CANDIDATE_NOT_PUBLISHED","created_at":datetime.now(timezone.utc).isoformat(),"canonical_backend":"native_pptxgenjs","external_skill_used":False,"file_count":len(files),"files":manifest_files,"bundle":{"name":bundle.name,"size":bundle.stat().st_size,"sha256":digest(bundle)},"gates":{"case_b":"NOT_RUN","case_c":"NOT_RUN","second_device":"NOT_RUN","human_workspace_scores":"PENDING","github_actions":"NOT_RUN"}}
    (output_root/"release_manifest.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    checksum_rows=[]
    for path in sorted([bundle,bootstrap,output_root/"release_manifest.json"]):checksum_rows.append(f"{digest(path)}  {path.name}")
    (output_root/"SHA256SUMS.txt").write_text("\n".join(checksum_rows)+"\n",encoding="utf-8")
    shutil.copy2(repo/"README_FIRST_INSTALL.md",output_root/"README_FIRST_INSTALL.md"); shutil.copy2(repo/"CHANGELOG.md",output_root/"CHANGELOG.md")
    return manifest

def main()->int:
    p=argparse.ArgumentParser();p.add_argument("--repo",default=Path(__file__).resolve().parents[1]);p.add_argument("--output-root",default="release");a=p.parse_args();result=build(Path(a.repo).resolve(),Path(a.output_root).resolve());print(json.dumps(result["bundle"],indent=2));return 0
if __name__=="__main__":raise SystemExit(main())
