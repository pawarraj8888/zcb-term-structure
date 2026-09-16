#!/usr/bin/env python3
"""Assemble the GitHub Pages site: docs/data.js (results + methodology for the
static page) and copies of the downloadable deliverables under docs/files/."""
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

from write_methodology import build_paragraphs

DOWNLOADS = [
    ("excel/Miniproject2_ZCB_Term_Structure.xlsx", "Miniproject2_ZCB_Term_Structure.xlsx"),
    ("writeup/Miniproject2_ZCB_Term_Structure_Writeup.pdf", "Miniproject2_ZCB_Term_Structure_Writeup.pdf"),
    ("output/zero_curve_payment_dates.csv", "zero_curve_payment_dates.csv"),
    ("output/bond_fit.csv", "bond_fit.csv"),
    ("data/Treasury_data_090426.xlsx", "Treasury_data_090426.xlsx"),
    ("output/figures/zero_curve.png", "zero_curve.png"),
]


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default="..", help="repository root")
    args = parser.parse_args(argv)
    root = Path(args.root).resolve()
    docs = root / "docs"
    files = docs / "files"
    files.mkdir(parents=True, exist_ok=True)

    results = json.loads((root / "output" / "results.json").read_text())
    results["methodology"] = [{"heading": h, "text": t} for h, t in build_paragraphs(results)]
    (docs / "data.js").write_text("window.ZCB_DATA = " + json.dumps(results, separators=(",", ":")) + ";\n")
    stale = docs / "data.json"
    if stale.exists():
        stale.unlink()

    for source, name in DOWNLOADS:
        src = root / source
        if not src.exists():
            raise SystemExit(f"missing deliverable: {src}")
        shutil.copy2(src, files / name)
    (docs / ".nojekyll").touch()
    print(f"Site data written to {docs} ({len(DOWNLOADS)} downloads copied)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
