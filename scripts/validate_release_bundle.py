from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import zipfile
from pathlib import Path

try:
    from audit_repository_content import _has_private_path
except ModuleNotFoundError:  # imported as ``scripts.validate_release_bundle``
    from scripts.audit_repository_content import _has_private_path


PROHIBITED = re.compile(
    r"(?i)(?:^|/)(?:input|output|staging|audit|benchmark|archive|private|"
    r"runtime_logs|install_receipts|release|\.cache|\.venv|node_modules|\.git)"
    r"(?:/|$)|(?:patient|real_world_mdt|~\$)"
)
REQUIRED = {
    "run_ppt_workflow.py",
    "package-lock.json",
    "requirements-lock.txt",
    "README_FIRST_INSTALL.md",
    "CHANGELOG.md",
    "scripts/bootstrap_windows.ps1",
    "extensions/visual_backends/native_pptxgenjs.py",
    "prompts/Codex_PPT_Production.md",
    "prompts/GPT_Art_Director.md",
    "visual_templates/library.json",
    "templates/README.md",
}
TEXT_SUFFIXES = {
    ".py", ".ps1", ".mjs", ".js", ".json", ".yaml", ".yml", ".md",
    ".csv", ".txt", ".xml", ".rels",
}
OFFICE_SUFFIXES = {".pptx", ".potx", ".docx", ".xlsx", ".xlsm"}


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def payload_has_private_absolute_path(payload: bytes) -> bool:
    return _has_private_path(payload.decode("utf-8", errors="replace"))


def contains_private_absolute_path(payload: bytes, suffix: str) -> bool:
    if payload_has_private_absolute_path(payload):
        return True
    if suffix not in OFFICE_SUFFIXES:
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as package:
            for member in package.infolist():
                if member.is_dir() or Path(member.filename).suffix.lower() not in TEXT_SUFFIXES:
                    continue
                if payload_has_private_absolute_path(package.read(member)):
                    return True
    except (OSError, zipfile.BadZipFile):
        return False
    return False


def validate(zip_path: Path) -> dict:
    issues=[]
    with zipfile.ZipFile(zip_path) as zf:
        names=[n for n in zf.namelist() if not n.endswith("/")]
        roots={n.split("/",1)[1] if "/" in n else n for n in names}
        missing=sorted(REQUIRED-roots)
        if missing: issues.append(f"missing required files: {missing}")
        for name in names:
            rel=name.split("/",1)[1] if "/" in name else name
            if PROHIBITED.search(rel): issues.append(f"prohibited path: {rel}")
            suffix=Path(rel).suffix.lower()
            if suffix in TEXT_SUFFIXES | OFFICE_SUFFIXES:
                if contains_private_absolute_path(zf.read(name),suffix):
                    issues.append(f"private absolute path: {rel}")
    return {"status":"PASS" if not issues else "BLOCKED","bundle":zip_path.name,"sha256":sha(zip_path),"file_count":len(names),"issues":issues}


def main()->int:
    p=argparse.ArgumentParser();p.add_argument("bundle");p.add_argument("--output");a=p.parse_args();result=validate(Path(a.bundle));payload=json.dumps(result,indent=2)+"\n";print(payload,end="")
    if a.output: Path(a.output).write_text(payload,encoding="utf-8")
    return 0 if result["status"]=="PASS" else 2
if __name__=="__main__": raise SystemExit(main())
