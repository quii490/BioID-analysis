#!/usr/bin/env python3
"""Download the official CRAPome control files and rebuild the frequency table.

The CRAPome site now redirects to REPRINT (reprint-apms.org), which serves the
same v1 flat files over a valid certificate. Run this when the table needs to be
rebuilt or refreshed:

    python3 scripts/fetch_crapome.py --out assets/crapome_frequency.tsv

Network access to reprint-apms.org is required. If it is unreachable, the
analysis still runs; the CRAPome column is then reported as "not applied"
rather than silently omitted.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

BASE = "https://reprint-apms.org/?q=system/files/"
FILES = {
    "matrix": "crap_db_v1_flat_file_human.xlsx",
    "annotation": "CRAPome_annotation_30MAR2013_0.xlsx",
}
HERE = Path(__file__).resolve().parent


def download(name, target):
    url = BASE + name
    print(f"downloading {url}")
    with urllib.request.urlopen(url, timeout=180) as response, open(target, "wb") as fh:
        while True:
            chunk = response.read(1 << 20)
            if not chunk:
                break
            fh.write(chunk)
    size = target.stat().st_size
    if size < 10_000 or target.read_bytes()[:2] != b"PK":
        raise SystemExit(f"{name} did not look like an xlsx workbook ({size} bytes)")
    print(f"  {name}: {size / 1e6:.1f} MB")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=str(HERE.parent / "assets" / "crapome_frequency.tsv"))
    parser.add_argument("--keep-downloads", action="store_true")
    args = parser.parse_args()

    target = Path(args.out).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    workdir = Path(tempfile.mkdtemp(prefix="crapome_")) if not args.keep_downloads else target.parent / "crapome_raw"
    workdir.mkdir(parents=True, exist_ok=True)

    paths = {}
    for key, name in FILES.items():
        paths[key] = workdir / name
        download(name, paths[key])

    builder = HERE / "build_crapome_table.py"
    cmd = [sys.executable, str(builder), str(paths["matrix"]), str(paths["annotation"]), str(target)]
    print("building frequency table")
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(result.stdout.strip() or result.stderr.strip())
    if result.returncode != 0:
        raise SystemExit("table build failed")
    print(f"written: {target} ({target.stat().st_size / 1e3:.0f} kB)")
    if args.keep_downloads:
        print(f"raw downloads kept in {workdir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
