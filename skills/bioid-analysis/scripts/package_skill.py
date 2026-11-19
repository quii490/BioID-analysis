#!/usr/bin/env python3
"""Package the skill, cache included, into one carryable archive.

The archive is self-contained: another machine needs only Python with numpy and
openpyxl, plus Chrome if PDF export is wanted. Nothing is fetched at run time —
the CRAPome control frequencies and the UniProt annotation cache travel inside.

    python3 scripts/package_skill.py [--out ~/Desktop] [--verify]

--verify extracts the archive to a temporary directory and runs the pipeline from
the extracted copy against a batch, which is the only way to be sure the
packaged skill does not depend on the paths it was built from.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tarfile
import tempfile
import zipfile
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SKILL_NAME = SKILL.name
EXCLUDE_DIRS = {"__pycache__", ".git", "dist", ".pytest_cache"}
EXCLUDE_SUFFIXES = {".pyc", ".pyo", ".DS_Store"}


def collect(root):
    files = []
    for path in sorted(root.rglob("*")):
        if any(part in EXCLUDE_DIRS for part in path.parts):
            continue
        if path.is_dir():
            continue
        if path.suffix in EXCLUDE_SUFFIXES or path.name.startswith(".~"):
            continue
        if path.name == "MEMORY.md":
            continue
        files.append(path)
    return files


def build_zip(out_dir, version):
    out_dir.mkdir(parents=True, exist_ok=True)
    archive = out_dir / f"{SKILL_NAME}-{version}.zip"
    files = collect(SKILL)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for path in files:
            zf.write(path, Path(SKILL_NAME) / path.relative_to(SKILL))
    return archive, files


def build_tarball(out_dir, version):
    archive = out_dir / f"{SKILL_NAME}-{version}.tar.gz"
    files = collect(SKILL)
    with tarfile.open(archive, "w:gz") as tf:
        for path in files:
            tf.add(path, arcname=str(Path(SKILL_NAME) / path.relative_to(SKILL)))
    return archive, files


def verify(archive, batch):
    """Extract and run from the extracted copy, so path assumptions surface."""
    with tempfile.TemporaryDirectory(prefix="skill_verify_") as work:
        work = Path(work)
        with zipfile.ZipFile(archive) as zf:
            zf.extractall(work)
        extracted = work / SKILL_NAME
        script = extracted / "scripts" / "bioid_analysis.py"
        if not script.exists():
            print("  FAIL: pipeline script missing from the archive")
            return False
        print(f"  extracted to {extracted}")
        print(f"  contents: {len(collect(extracted))} files, "
              f"{(sum(p.stat().st_size for p in collect(extracted)) / 1e6):.1f} MB")
        out = work / "out"
        result = subprocess.run(
            [sys.executable, str(script), str(batch), "--out", str(out), "--plotly", "cdn"],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print("  FAIL: pipeline failed from the extracted copy")
            print(result.stdout[-1500:])
            print(result.stderr[-1500:])
            return False
        lines = [l for l in result.stdout.splitlines()
                 if "annotations" in l or "report:" in l or "conditions analysed" in l]
        for line in lines:
            print("   ", line)
        report = out / "BioID_analysis_report.html"
        summary = out / "BioID_analysis_summary.xlsx"
        ok = report.exists() and summary.exists() and report.stat().st_size > 100_000
        print(f"  report written: {report.exists()} ({report.stat().st_size / 1e6:.1f} MB)"
              if report.exists() else "  FAIL: no report written")
        return ok


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=str(Path.home() / "Desktop"),
                        help="where to write the archive (default: ~/Desktop)")
    parser.add_argument("--version", default=date.today().isoformat())
    parser.add_argument("--batch", default=str(Path.home() / "Desktop" / "Candidates" / "BioID-3nd"),
                        help="batch used by --verify")
    parser.add_argument("--verify", action="store_true",
                        help="extract the archive and run the pipeline from the extracted copy")
    parser.add_argument("--no-tar", action="store_true", help="write only the .zip")
    args = parser.parse_args()

    out_dir = Path(args.out).expanduser().resolve()
    archive, files = build_zip(out_dir, args.version)
    print(f"packaged {len(files)} files")
    print(f"  {archive}  ({archive.stat().st_size / 1e6:.1f} MB)")
    if not args.no_tar:
        tarball, _ = build_tarball(out_dir, args.version)
        print(f"  {tarball}  ({tarball.stat().st_size / 1e6:.1f} MB)")

    biggest = sorted(files, key=lambda p: -p.stat().st_size)[:5]
    print("  largest members:")
    for path in biggest:
        print(f"    {path.relative_to(SKILL)}  {path.stat().st_size / 1e6:.1f} MB")

    if args.verify:
        batch = Path(args.batch).expanduser()
        if not batch.is_dir():
            print(f"  skipping verification: {batch} is not a folder")
            return 0
        print(f"verifying against {batch}")
        if not verify(archive, batch):
            return 1
        print("  verification passed: the packaged skill runs standalone")
    return 0


if __name__ == "__main__":
    sys.exit(main())
