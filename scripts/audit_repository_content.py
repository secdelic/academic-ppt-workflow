from __future__ import annotations

import argparse
import fnmatch
import json
import re
import zipfile
from pathlib import Path


# Match ordinary and JSON-escaped Windows drive paths.  The expression requires
# a concrete drive letter so the scanner's own regular-expression source is not
# a match.
WINDOWS_ABSOLUTE_PATH = re.compile(r"(?i)(?<![A-Za-z0-9_])[A-Z]:(?:\\+|/+)[^\r\n\"'<>]*")
# A private UNC file path must include server, share, and at least one concrete
# path component. Requiring an alphanumeric segment start avoids treating
# escaped regular expressions such as ``\\begin\\s*`` as network paths.
UNC_PATH = re.compile(
    r"(?i)(?<!\\)\\\\"
    r"[A-Za-z0-9][A-Za-z0-9._-]*\\+"
    r"[A-Za-z0-9$][A-Za-z0-9$._-]*"
    r"(?:\\+[A-Za-z0-9_][^\\\r\n\"'<>]*)+"
)
SECRET = re.compile(
    r"(?i)(?:"
    r"sk-[A-Za-z0-9_-]{16,}|"
    r"gh[pousr]_[A-Za-z0-9]{16,}|"
    r"AKIA[0-9A-Z]{16}|"
    r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----|"
    r"(?:api[_-]?key|access[_-]?token|refresh[_-]?token|password|secret)"
    r"\s*[:=]\s*['\"][^'\"\r\n]{8,}['\"]"
    r")"
)
PROHIBITED_NAMES = re.compile(r"(?i)(?:patient|real[_-]?world[_-]?mdt|clinical[_-]?output|~\$)")

TEXT_SUFFIXES = {
    "",
    ".cjs",
    ".cfg",
    ".csv",
    ".css",
    ".gitignore",
    ".gitattributes",
    ".html",
    ".ini",
    ".js",
    ".json",
    ".lock",
    ".md",
    ".mjs",
    ".ps1",
    ".py",
    ".rels",
    ".toml",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}
OFFICE_SUFFIXES = {".docx", ".pptx", ".potx", ".xlsx"}
FONT_SUFFIXES = {".eot", ".otf", ".ttc", ".ttf", ".woff", ".woff2"}
KNOWN_BINARY_SUFFIXES = OFFICE_SUFFIXES | FONT_SUFFIXES | {
    ".avif",
    ".bmp",
    ".emf",
    ".gif",
    ".ico",
    ".jpeg",
    ".jpg",
    ".pdf",
    ".png",
    ".svgz",
    ".tif",
    ".tiff",
    ".webp",
    ".wmf",
}
TEMP_SUFFIXES = {".bak", ".old", ".orig", ".swp", ".swo", ".temp", ".tmp"}
LOCAL_CONFIG_NAMES = {
    ".env",
    ".npmrc",
    ".pypirc",
    "credentials.json",
    "credentials.yaml",
    "credentials.yml",
    "local.json",
    "local.yaml",
    "local.yml",
    "secrets.json",
    "secrets.yaml",
    "secrets.yml",
}
PROHIBITED_PATH_PARTS = {
    ".cache",
    ".git",
    ".venv",
    "archive",
    "audit",
    "benchmark",
    "cache",
    "input",
    "install_receipts",
    "node_modules",
    "output",
    "private",
    "runtime_logs",
    "staging",
}
BINARY_ALLOWLIST = (
    "regression/fixtures/synthetic*/**",
    "regression/fixtures/synthetic*/*",
    "templates/authorized_sanitized/**",
    "visual_workspaces/*/synthetic_preview.pptx",
)
MAX_RELEASE_FILE_BYTES = 10 * 1024 * 1024


def candidate_files(repo: Path) -> list[str]:
    """Return the exact release-bundle candidate set.

    Keeping this import local prevents the audit policy from drifting away from
    the bundle builder while still allowing the module to be imported in tests.
    """

    try:
        from build_release_bundle import selected
    except ModuleNotFoundError:  # imported as ``scripts.audit_repository_content``
        from scripts.build_release_bundle import selected
    return [path.relative_to(repo).as_posix() for path in selected(repo)]


def _issue(issues: list[dict[str, str]], rel: str, issue: str, *, entry: str | None = None) -> None:
    row = {"severity": "CRITICAL", "file": rel, "issue": issue}
    if entry:
        # OOXML part names are archive-relative and safe to report; never include
        # the matched secret/path value itself.
        row["entry"] = entry
    if row not in issues:
        issues.append(row)


def _has_private_path(text: str) -> bool:
    return bool(WINDOWS_ABSOLUTE_PATH.search(text) or UNC_PATH.search(text))


def _scan_text(text: str, rel: str, issues: list[dict[str, str]], *, entry: str | None = None) -> None:
    if _has_private_path(text):
        _issue(issues, rel, "absolute private path", entry=entry)
    if SECRET.search(text):
        _issue(issues, rel, "secret-like value", entry=entry)


def _scan_office_package(path: Path, rel: str, issues: list[dict[str, str]]) -> tuple[int, int]:
    scanned_parts = 0
    external_relationships = 0
    try:
        with zipfile.ZipFile(path) as package:
            for info in package.infolist():
                name = info.filename.replace("\\", "/")
                lower = name.lower()
                if not (lower.endswith(".xml") or lower.endswith(".rels") or lower == "[content_types].xml"):
                    continue
                scanned_parts += 1
                text = package.read(info).decode("utf-8", errors="replace")
                _scan_text(text, rel, issues, entry=name)
                if lower.endswith(".rels") and re.search(
                    r"(?i)TargetMode\s*=\s*['\"]External['\"]", text
                ):
                    external_relationships += 1
                    _issue(issues, rel, "OOXML external relationship", entry=name)
    except (OSError, zipfile.BadZipFile, RuntimeError):
        _issue(issues, rel, "invalid or unreadable OOXML package")
    return scanned_parts, external_relationships


def _binary_is_allowed(rel: str) -> bool:
    return any(fnmatch.fnmatch(rel, pattern) for pattern in BINARY_ALLOWLIST)


def _is_local_config(rel: str) -> bool:
    name = Path(rel).name.lower()
    if name in LOCAL_CONFIG_NAMES:
        return True
    return name.startswith(".env.") and not name.endswith((".example", ".sample", ".template"))


def _is_temp_file(rel: str) -> bool:
    path = Path(rel)
    name = path.name.lower()
    return name.startswith("~$") or path.suffix.lower() in TEMP_SUFFIXES


def audit(repo: Path) -> dict:
    repo = repo.resolve()
    files = candidate_files(repo)
    issues: list[dict[str, str]] = []
    office_packages_scanned = 0
    office_xml_parts_scanned = 0
    external_relationships = 0
    binary_files = 0
    large_files = 0
    font_files = 0

    for rel in files:
        normalized = Path(rel).as_posix()
        parts = {part.lower() for part in Path(normalized).parts}
        path = repo / normalized
        suffix = path.suffix.lower()

        if parts & PROHIBITED_PATH_PARTS:
            _issue(issues, normalized, "prohibited release path class")
        if PROHIBITED_NAMES.search(normalized):
            _issue(issues, normalized, "prohibited filename class")
        if _is_local_config(normalized):
            _issue(issues, normalized, "local configuration file")
        if _is_temp_file(normalized):
            _issue(issues, normalized, "temporary or backup file")
        if path.is_symlink():
            _issue(issues, normalized, "symbolic link in release candidate")
        if not path.is_file():
            _issue(issues, normalized, "candidate file missing or unreadable")
            continue
        if path.stat().st_size > MAX_RELEASE_FILE_BYTES:
            large_files += 1
            _issue(issues, normalized, "large file requires explicit authorization")
        if suffix in FONT_SUFFIXES:
            font_files += 1
            _issue(issues, normalized, "embedded font file is prohibited")

        if suffix in OFFICE_SUFFIXES:
            binary_files += 1
            if not _binary_is_allowed(normalized):
                _issue(issues, normalized, "binary asset is outside the release allowlist")
            office_packages_scanned += 1
            part_count, external_count = _scan_office_package(path, normalized, issues)
            office_xml_parts_scanned += part_count
            external_relationships += external_count
            continue

        if suffix in KNOWN_BINARY_SUFFIXES:
            binary_files += 1
            if not _binary_is_allowed(normalized):
                _issue(issues, normalized, "binary asset is outside the release allowlist")
            continue

        if suffix not in TEXT_SUFFIXES:
            _issue(issues, normalized, "unrecognized file type in release candidate")
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        _scan_text(text, normalized, issues)

    prohibited_name_count = sum(issue["issue"] == "prohibited filename class" for issue in issues)
    return {
        "status": "PASS" if not issues else "BLOCKED",
        "release_candidate_file_count": len(files),
        "critical_count": len(issues),
        "real_patient_files_in_git": 0 if not prohibited_name_count else None,
        "binary_files_checked": binary_files,
        "office_packages_scanned": office_packages_scanned,
        "office_xml_parts_scanned": office_xml_parts_scanned,
        "external_relationship_count": external_relationships,
        "large_file_count": large_files,
        "font_file_count": font_files,
        "issues": issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output")
    args = parser.parse_args()
    result = audit(Path(args.repo).resolve())
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(payload, encoding="utf-8")
    print(payload, end="")
    return 0 if result["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
