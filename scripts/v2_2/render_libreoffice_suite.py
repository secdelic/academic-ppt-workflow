from __future__ import annotations

import argparse
import csv
import json
import subprocess
import time
from pathlib import Path

import pypdfium2 as pdfium


def run() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--deck-index", type=Path, required=True)
    parser.add_argument("--soffice", type=Path, required=True)
    parser.add_argument("--profile-root", type=Path, required=True)
    parser.add_argument("--render-index", type=Path, required=True)
    args = parser.parse_args()
    decks = list(csv.DictReader(args.deck_index.open("r", encoding="utf-8-sig", newline="")))
    results = []
    for row in decks:
        start = time.perf_counter()
        pptx = Path(row["pptx_path"]).resolve()
        output_dir = pptx.parent / "libreoffice_compat"
        profile = args.profile_root / row["project_key"] / f"arm_{row['arm']}"
        if output_dir.exists():
            raise RuntimeError(f"Refusing to overwrite {output_dir}")
        output_dir.mkdir(parents=True)
        profile.mkdir(parents=True)
        profile_uri = "file:///" + str(profile.resolve()).replace("\\", "/")
        completed = subprocess.run(
            [
                str(args.soffice),
                "--headless",
                f"-env:UserInstallation={profile_uri}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_dir),
                str(pptx),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
        pdf_path = output_dir / f"{pptx.stem}.pdf"
        if completed.returncode != 0 or not pdf_path.is_file() or pdf_path.stat().st_size == 0:
            raise RuntimeError(
                f"LibreOffice export failed for {pptx}: {completed.stderr}"
            )
        preview = output_dir / "preview"
        preview.mkdir()
        document = pdfium.PdfDocument(str(pdf_path))
        try:
            for index in range(len(document)):
                page = document[index]
                image = page.render(scale=1.5).to_pil()
                image.save(preview / f"slide_{index + 1:03d}.png")
        finally:
            document.close()
        results.append(
            {
                "project_key": row["project_key"],
                "arm": row["arm"],
                "pptx": str(pptx),
                "pdf": str(pdf_path),
                "preview": str(preview),
                "pdf_page_count": len(list(preview.glob("slide_*.png"))),
                "exit_code": completed.returncode,
                "runtime_seconds": round(time.perf_counter() - start, 3),
                "stdout": completed.stdout.strip(),
                "stderr": completed.stderr.strip(),
            }
        )
    args.render_index.parent.mkdir(parents=True, exist_ok=True)
    args.render_index.write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"LIBREOFFICE_DECKS={len(results)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
