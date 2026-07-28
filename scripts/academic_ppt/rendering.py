from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageOps


def _run(command: list[str], timeout: int, env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        check=False,
        text=True,
        capture_output=True,
        timeout=timeout,
        encoding="utf-8",
        errors="replace",
        env=env,
    )


def _find_executable(explicit: str | None, names: list[str]) -> str | None:
    if explicit:
        candidate = Path(explicit)
        if candidate.exists():
            return str(candidate)
        located = shutil.which(explicit)
        if located:
            return located
    for name in names:
        located = shutil.which(name)
        if located:
            return located
    return None


def normalize_preview_names(preview_dir: Path) -> list[Path]:
    candidates = sorted(
        [p for p in preview_dir.iterdir() if p.suffix.lower() == ".png"],
        key=lambda p: int("".join(filter(str.isdigit, p.stem)) or "0"),
    )
    normalized: list[Path] = []
    for index, path in enumerate(candidates, start=1):
        target = preview_dir / f"slide_{index:03d}.png"
        if path != target:
            if target.exists():
                target.unlink()
            path.rename(target)
        normalized.append(target)
    return normalized


def render_pptx(
    pptx_path: Path,
    pdf_path: Path,
    preview_dir: Path,
    script_root: Path,
    timeout: int,
    soffice: str | None = None,
    pdftoppm: str | None = None,
) -> tuple[str, list[Path], str]:
    preview_dir.mkdir(parents=True, exist_ok=True)
    powershell = shutil.which("powershell") or shutil.which("powershell.exe")
    if powershell:
        command = [
            powershell,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_root / "render_pptx.ps1"),
            "-InputPptx",
            str(pptx_path),
            "-OutputPdf",
            str(pdf_path),
            "-PreviewDir",
            str(preview_dir),
        ]
        result = _run(command, timeout)
        if result.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0:
            previews = normalize_preview_names(preview_dir)
            if previews:
                return "PowerPoint COM", previews, result.stdout.strip()
        com_error = (result.stdout + "\n" + result.stderr).strip()
    else:
        com_error = "PowerShell unavailable"

    soffice_path = _find_executable(soffice, ["soffice.com", "soffice", "libreoffice"])
    if not soffice_path:
        raise RuntimeError(f"Render unavailable. PowerPoint error: {com_error}")
    with tempfile.TemporaryDirectory(prefix="academic-ppt-lo-") as profile:
        profile_uri = Path(profile).resolve().as_uri()
        result = _run(
            [
                soffice_path,
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_path.parent),
                str(pptx_path),
            ],
            timeout,
        )
    generated = pdf_path.parent / f"{pptx_path.stem}.pdf"
    if generated.exists() and generated != pdf_path:
        generated.replace(pdf_path)
    if result.returncode != 0 or not pdf_path.exists() or pdf_path.stat().st_size == 0:
        raise RuntimeError(
            "LibreOffice render failed after PowerPoint failure. "
            f"PowerPoint: {com_error}; LibreOffice: {result.stdout} {result.stderr}"
        )
    pdftoppm_path = _find_executable(pdftoppm, ["pdftoppm"])
    if not pdftoppm_path:
        raise RuntimeError("PDF created by LibreOffice, but pdftoppm is unavailable for slide previews")
    prefix = preview_dir / "slide"
    raster = _run([pdftoppm_path, "-png", "-r", "144", str(pdf_path), str(prefix)], timeout)
    if raster.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {raster.stdout} {raster.stderr}")
    previews = normalize_preview_names(preview_dir)
    if not previews:
        raise RuntimeError("Renderer produced no slide previews")
    return "LibreOffice + Poppler", previews, (result.stdout + raster.stdout).strip()


def create_contact_sheet(previews: list[Path], output_path: Path) -> None:
    if not previews:
        raise ValueError("No previews for contact sheet")
    thumb_w, thumb_h = 400, 225
    cols = 3
    rows = (len(previews) + cols - 1) // cols
    canvas = Image.new("RGB", (cols * thumb_w, rows * thumb_h), "white")
    for index, path in enumerate(previews):
        with Image.open(path) as image:
            thumb = ImageOps.contain(image.convert("RGB"), (thumb_w, thumb_h))
            x = (index % cols) * thumb_w + (thumb_w - thumb.width) // 2
            y = (index // cols) * thumb_h + (thumb_h - thumb.height) // 2
            canvas.paste(thumb, (x, y))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path)
