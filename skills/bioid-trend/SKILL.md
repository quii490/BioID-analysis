---
name: bioid-trend
description: "Plot one or a few proteins' normalized log2FC across several BioID workbooks (bioid / bioid-batch-qc / bioid-analysis outputs) as a line plot: one solid line per protein, colour per protein, marker shape per group, a protein not detected in a BioID drawn as a hollow marker at 0 with the segments touching it dashed, optional background bands that group the x axis, and an optional mean line across the plotted proteins. Use when asked how a protein or paralog family trends across constructs, mutants, baits, cell lines or batches, or to reproduce the ATOSB-PP2A style BioID line figure as a PDF."
---

# BioID trend / cross-workbook line plot

`scripts/bioid_trend.py` reads N workbooks (one per x-axis point) plus a protein list and writes a
vector PDF, a PNG preview, a values TSV and a sources TSV. Inputs are opened read-only.

The design follows the lab's ATOSB-PP2A figure: colour = protein, marker shape = group
(A scaffold = circle, C catalytic = square, B55 regulatory = triangle), legend label `group: gene`.
Every protein is a solid line; the only dashed segments are the ones touching a protein that was not
detected in that BioID (see below). Line style does not encode paralogs.

## Quick start

```bash
bash scripts/doctor.sh          # prints a python that actually has matplotlib
PY="${PYTHON:-/usr/bin/python3}"
SKILL="${SKILL_DIR:-$HOME/.dsh/skills/bioid-trend}"

# 1) see the matrix first; writes nothing
"$PY" "$SKILL/scripts/bioid_trend.py" --samples samples.tsv --proteins proteins.tsv \
    --check --xlabel "ATOSB construct"

# 2) render after confirming
"$PY" "$SKILL/scripts/bioid_trend.py" --samples samples.tsv --proteins proteins.tsv --out out/ \
    --name atosb_pp2a --xlabel "ATOSB construct" \
    --title "ATOSB full-length and truncations: BioID signal of PP2A subunits" \
    --legend-title "PP2A groups / proteins"
```

`--check` is the confirmation gate: it prints the x order, the value column used per workbook, the
protein × BioID matrix (`▲0` = missing), the per-point provenance and every warning. Get an explicit
yes before rendering when the figure is going into a report.

## The two input tables

Tab- or comma-separated, `#` comments allowed, header row optional. **Row order is the x order.**

`samples.tsv` — one x point per line:

```text
label	file	sheet	column	xgroup
FL	/…/BioID-2nd/flagged/21FAM214B_flagged.xlsx			FL
N	/…/bioid/flagged/ATOSB-N_processed_flagged.xlsx			truncation
C	/…/bioid/flagged/ATOSB-C_processed_flagged.xlsx			truncation
F405K	/…/bioid/flagged/ATOSB-F405K_processed_flagged.xlsx			mutation
```

- `file` may be relative to the samples table; `label` is the tick text.
- `sheet` is required for `bioid-analysis`'s `BioID_analysis_summary.xlsx` (one sheet per condition;
  `condition` is accepted as an alias). Without it the script uses `Proteins`, and refuses a workbook
  that has several condition sheets instead of silently picking one.
- `column` pins one exact value-column name for that file (overrides `--value`).
- `xgroup` (aliases `block`, `section`, `xclass`) groups consecutive x points; the group is painted
  as a background band and labelled above the axes. Consecutive rows with the same name form one band,
  a different name starts a new one, and an empty cell leaves that slot unbanded. Without a header row
  the positional order is `label, file, sheet, column, xgroup`.
- `--no-xgroup-bands` keeps the column but skips the bands; `--no-xgroup-labels` paints the bands
  without the names above the axis.

`proteins.tsv` — one line per protein:

```text
# gene	group	label	style	color	accession
PPP2R1A	A scaffold
PPP2R1B	A scaffold
PPP2CA	C catalytic
PPP2CB	C catalytic
PPP2R2A	B55 regulatory
PPP2R2D	B55 regulatory
```

Only `gene` is required. `label` changes the legend text, `style` (`solid`/`dashed`/`dashdot`/`dotted`)
and `color` override the automatic assignment, `accession` matches by UniProt instead of symbol. For a
one-off without groups use `--genes PPP2R1A,PPP2CA`, and for 1-3 points use
`--point FL=/path/a.xlsx --point N=/path/b.xlsx` instead of a samples table.

Alias matching (e.g. `ATOSB` vs a workbook's `FAM214B`) needs `--aliases aliases.tsv`: any two columns,
everything in one connected component is treated as the same protein, so direction does not matter.

## Which number is plotted

Per workbook, `--value auto` (default) takes the first column present:

| priority | column | produced by |
|---|---|---|
| 1 | `Norm_LOG2FC` | `bioid-batch-qc` (`*_flagged.xlsx`) |
| 2 | `log2FC (EV-corrected)` | `bioid-analysis` summary, per-condition sheet |
| 3 | `LOG2FC` | `bioid` (`*_processed.xlsx`) |

The column actually used is printed for every workbook and recorded in `<stem>_sources.tsv`; the y-axis
label follows it. When the workbooks disagree the caption says so and the run warns — do not compare
absolute heights across different columns. `bioid` products carry `GN`/`LOG2FC` twice in the header;
the first occurrence is used (the duplicates carry identical values).

`LOG2FC` is *not* interchangeable with `Norm_LOG2FC`: the former is uncorrected, the latter removes the
batch's control/load offset. Use `--value raw` only when no normalized column exists anywhere.

## Missing proteins: 0 is a position, not a measurement

A protein counts as missing when the gene is absent from the workbook, or the row exists but the value
cell is empty (typically control not detected). Both are drawn as a **hollow marker at y = 0**, joined
by a **dashed segment** on each side (the rest of that protein's line stays solid), and listed in the
caption and the `na_samples` column. The values TSV keeps `NA`; only the plot moves it to 0.

Do not read a 0 as "no binding": it means "not quantified in this run".

A protein missing from *every* BioID is a different case: it is not plotted (a flat line of hollow
zeros would read as a measurement), it stays in the values TSV with `plotted=no`, the caption names it,
and the run warns. That is usually a wrong gene symbol. `--keep-all-na` draws it anyway if the absence
itself is the point.

## Background bands and the mean line

`xgroup` turns the x axis into labelled blocks (e.g. full-length / truncation / mutation). Bands are
pale tints that stay behind the grid; alternating group names get different tints so adjacent blocks
are distinguishable, and a narrow band gets its label rotated instead of overlapping its neighbour.
Bands are decoration only — the data stays on the normal categorical axis.

`--mean` draws one thick dark line, the **mean of the plotted proteins at each BioID**. Per the
project's choice, a protein that was not detected there contributes its plotted value, i.e. **0**.
That means the mean mixes abundance with detection: a BioID where half the panel is missing will show
a low mean even if the proteins that were measured are strongly enriched. The caption says so
explicitly; do not present the mean as "average binding" without that caveat. The mean is also written
to the values TSV as a `MEAN` row (and printed in `--check`), so it can be recomputed by hand.

## Conventions and overrides

- Colour order is the protein table order (tab10). Marker shape follows the group's first appearance.
  Lines are solid; a segment with a not-detected endpoint is dashed, so the dash reads as "this end is
  a 0 by convention", not as a different protein class.
- Markers carry a thin white edge and are drawn above the lines, so two points that land almost on
  top of each other still show both outlines. Measured points sit above the hollow markers.
- Several proteins missing at the *same* BioID share one position (y = 0). They are drawn as
  concentric hollow markers of increasing size, largest first, so three missing proteins read as three
  outlines rather than one. The legend carries both keys: `not detected (drawn at 0)` (hollow marker)
  and `segment joining a not-detected point` (dashed line).
- Legend text is `group: gene`; without groups it is just the gene and every marker is a circle.
- A `style` column still overrides the line style per protein, but it is better left empty: a manual
  dashed line would be indistinguishable from the not-detected convention.
- `--fig-width/--fig-height` (default 6.0 × 4.5 in), `--ylim "0,9.5"`, `--yticks "0,2,4,6,8"`.
- `--rotate-labels` defaults to `auto`: the x labels are turned 25° only when a label would not fit
  its slot (long construct names). Pass `0` to force horizontal labels, or an angle to force rotation.
- `--no-caption` drops the generated figure legend; `--caption-text "…"` replaces it entirely
  (use it to name the truncation boundaries, e.g. N = 1-347, C = 348-538).
- Fonts: DejaVu Sans by default (the same face as the reference figure). **The PDF backend cannot fall
  back per glyph**, so when any visible text contains non-ASCII the script switches the whole figure to
  a CJK-capable face (Arial Unicode MS / PingFang HK / Hiragino Sans GB, whichever exists) and prints
  which one — otherwise Chinese labels silently render blank even though the text layer still carries
  the characters. Type 42 embedding keeps text selectable and editable in Illustrator.

## Outputs

One folder per figure, figures and tables apart (default since 2026-10-28):

```text
<out>/<stem>/
    <stem>.pdf            vector figure
    <stem>.png            300 dpi preview
    data/<stem>_values.tsv   gene × BioID matrix, `NA` preserved, `na_count`, `na_samples`, `plotted`
    data/<stem>_sources.tsv  label / file / sheet / column / rows
```

`--flat` restores the old behaviour (everything directly in `--out`, no subfolder). `--name` sets the
figure name and therefore the folder name; without it the stem lists up to four gene names.
Default `--out` is `<samples dir>/bioid_trend/`.

## Before delivering

1. Open the PNG/PDF: hollow markers must sit at 0, every segment touching them must be dashed, all other
   segments solid, and no line may silently stop at a missing point.
2. Cross-check the printed matrix against the caption's missing list.
3. Confirm the PDF is still vector: `pdfimages -list <stem>.pdf` shows no rows, and `pdftotext` finds
   the axis labels and the caption.
4. Read the batch warning. Values are normalized *within* each BioID, so a curve mixing batches
   (e.g. full-length from `BioID-2nd` with mutants from `BioID-5th`) shows relative change per protein,
   not comparable absolute enrichment; say so if the figure is presented as absolute signal.

## Environment and tests

Needs Python ≥3.9 with `matplotlib`, `openpyxl`, `numpy`. On this Mac `python3` (Framework 3.13) has no
matplotlib while `/usr/bin/python3` does — run `scripts/doctor.sh` instead of guessing. `pdftoppm` /
`pdfimages` come from poppler and are only used for self-checks.

```bash
/usr/bin/python3 -m unittest discover -s tests -v
```

The tests build synthetic `bioid` / `bioid-batch-qc` / `bioid-analysis` workbooks and assert the
column fallback, NA handling (row present vs row absent), alias matching, relative paths and that the
PDF stays image-free.
