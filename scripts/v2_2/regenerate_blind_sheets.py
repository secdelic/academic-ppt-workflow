from __future__ import annotations

import argparse
import json
from pathlib import Path

from analyze_benchmark import contact_sheet


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-root", type=Path, required=True)
    parser.add_argument("--audit-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    render_rows = json.loads(
        (args.run_root / "powerpoint_render_index.json").read_text(
            encoding="utf-8-sig"
        )
    )
    render_map = {
        (row["project_key"], row["arm"]): row for row in render_rows
    }
    blind_map = json.loads(
        (args.audit_root / "blind_code_map.json").read_text(encoding="utf-8")
    )
    count = 0
    for project, arms in blind_map.items():
        for arm, code in arms.items():
            preview = Path(render_map[(project, arm)]["preview"])
            output = (
                args.output_root
                / f"{project}_{code.replace(' ', '_')}.png"
            )
            contact_sheet(
                sorted(preview.glob("slide_*.png")),
                output,
                f"{project} · {code}",
            )
            count += 1
    print(f"ANONYMOUS_CONTACT_SHEETS={count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
