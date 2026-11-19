#!/usr/bin/env python3
"""Objective layout checks on a PDF printed from the report.

The pipeline no longer writes a multi-page PDF of its own (that feature was removed as
unnecessary: the HTML is the report, and a per-figure PDF button covers the "export one
plot" need). A PDF still exists whenever someone prints the report — the browser's
Cmd-P with "Save as PDF" — and these checks are what to run on it before sending
the document anywhere. They were written against the pipeline-generated PDF, so a few of
them (the caveat-attribution check in particular) simply skip pages they cannot identify.

The visual gate for a report is a human (or a judge model) looking at the pages,
but three of the failures this report has actually shipped are measurable, and a
measurement catches them on every batch instead of on the one page someone
happened to look at:

  1. text drawn on text — two labels whose boxes intersect. Comparing only spans
     of the same size keeps the tiny side-bearing overlaps between neighbouring
     glyph runs out of the count, and the area threshold keeps rotated axis tick
     labels (whose axis-aligned boxes always intersect when they are adjacent)
     out of it as well unless the glyphs genuinely share pixels.
  2. text past the page edge — a table column that runs off the paper. The
     first version of this report clipped the last column to "CRAPo" and the
     values to "0." with nothing in the extracted text to show for it.
  3. a near-blank page — the tail of a document that broke badly.

    python3 tests/check_pdf_layout.py BioID_analysis_report.pdf [--quiet]

Exit code 0 when every check passes, 1 otherwise.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF

# A real collision. Two gene names printed on each other cover tens of square
# points; two adjacent glyphs in one line share a square point or two of side
# bearing. Rotated tick labels sit at 40-80 points of separation and overlap by
# ~35 pt^2 in axis-aligned terms while their glyphs stay clear, so the threshold
# is set above that and the size filter keeps most of them out anyway.
COLLISION_AREA = 60.0
SIZE_TOLERANCE = 0.6
EDGE_MARGIN = 6.0
BLANK_CHARS = 60


def spans(page):
    out = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                if span["text"].strip():
                    out.append((fitz.Rect(span["bbox"]), span["text"], span["size"]))
    return out


def check_banner_attribution(doc, payload: Path) -> list:
    """Every caveat box must state its own condition's load correction.

    The box carries the condition's warnings, and the first of them quotes the load
    correction that was applied. When the box was printed at the top of its panel and
    pagination pushed the figure to the next page, it landed at the foot of the
    *previous* condition's table and read as a footnote to that condition's numbers.
    Nothing about that is visible in a single page's layout, which is why it is
    checked here against the payload: the caption's candidate count identifies the
    page's condition, and the quoted correction must belong to it.
    """
    data = json.loads(payload.read_text())
    by_count, by_corr = {}, {}
    for cond in data["conditions"]:
        n = sum(1 for row in cond["points"] if row[data["pointFields"].index("pass")])
        by_count.setdefault(n, []).append(cond)
        by_corr[round(cond["diagnostics"]["correction_log2"], 2)] = cond["condition"]

    wrong = []
    for n, page in enumerate(doc, 1):
        text = page.get_text()
        counts = [int(x) for x in re.findall(r"(\d+) CANDIDATES", text)]
        corrs = [float(x) for x in re.findall(r"correction of ([+-]?\d+\.\d+) log2", text)]
        if not counts or not corrs:
            continue
        owners = by_count.get(counts[0], [])
        if len(owners) != 1:
            continue
        expected = round(owners[0]["diagnostics"]["correction_log2"], 2)
        if all(abs(c - expected) > 0.02 for c in corrs):
            wrong.append((n, owners[0]["condition"], expected, corrs[0]))
    return wrong


def check(path: Path, quiet: bool = False, payload: Path | None = None) -> bool:
    doc = fitz.open(path)
    collisions, offpage, blanks = [], [], []
    for n, page in enumerate(doc, 1):
        items = spans(page)
        for i in range(len(items)):
            r1, t1, s1 = items[i]
            for j in range(i + 1, len(items)):
                r2, t2, s2 = items[j]
                if abs(s1 - s2) > SIZE_TOLERANCE:
                    continue
                inter = r1 & r2
                if not inter.is_empty and inter.get_area() >= COLLISION_AREA:
                    collisions.append((n, round(inter.get_area(), 1), t1[:28], t2[:28]))
        width = page.rect.width
        for rect, text, _ in items:
            if rect.x1 > width - EDGE_MARGIN:
                offpage.append((n, round(rect.x1 - width, 1), text[:34]))
        if len(page.get_text().strip()) < BLANK_CHARS and not page.get_images():
            drawings = page.get_drawings()
            if len(drawings) < 50:
                blanks.append(n)

    # A clipped table column does not draw past the page edge — Chrome crops it — so
    # the columns are checked by name instead: wherever a candidate table header is
    # present, every one of its column headings must be readable in full.
    truncated = []
    for n, page in enumerate(doc, 1):
        text = page.get_text()
        if "Gene" not in text or "Accession" not in text:
            continue
        for head in ("Accession", "CRAPome"):
            if head not in text:
                truncated.append((n, head))

    def report(name, rows, detail):
        status = "PASS" if not rows else "FAIL"
        print(f"  [{status}] {name}{'' if not rows else ' — ' + detail}")
        if rows and not quiet:
            for row in rows[:8]:
                print(f"        {row}")

    print(f"layout check: {path.name} ({doc.page_count} pages)")
    report("no text is drawn on other text", collisions,
           f"{len(collisions)} overlapping pair(s)")
    report("no text runs off the page edge", offpage,
           f"{len(offpage)} span(s) past the edge")
    report("no near-blank page", blanks, f"pages {blanks}")
    report("every candidate-table column heading is complete", truncated,
           f"{len(truncated)} missing heading(s)")
    misattributed = check_banner_attribution(doc, payload) if payload else []
    if payload:
        report("every caveat box belongs to its page's condition", misattributed,
               f"{len(misattributed)} page(s) — (page, condition, expected, shown)")
    failed = bool(collisions or offpage or blanks or truncated or misattributed)
    print("  " + ("ALL LAYOUT CHECKS PASSED" if not failed else "LAYOUT CHECKS FAILED"))
    return not failed


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--quiet", action="store_true", help="counts only, no examples")
    parser.add_argument("--payload", type=Path,
                        help="analysis_payload.json next to the report; enables the "
                             "caveat-box attribution check")
    args = parser.parse_args(argv)
    if not args.pdf.exists():
        raise SystemExit(f"no such PDF: {args.pdf}")
    if args.payload and not args.payload.exists():
        raise SystemExit(f"no such payload: {args.payload}")
    return 0 if check(args.pdf, args.quiet, args.payload) else 1


if __name__ == "__main__":
    sys.exit(main())
