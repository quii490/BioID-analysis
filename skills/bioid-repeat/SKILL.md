---
name: bioid-repeat
description: "Compare the same bait's prey enrichment between two independent BioID experiments (different cell lines, batches or runs) and render a correlation scatter report per input pair: matched proteins, Spearman/Pearson r, fitted slope, censored (lower-bound) handling, and the proteins detected in only one side. Use when asked whether a BioID result reproduces across cell lines or batches, to plot one bait's two experiments against each other, or to list cell-line-specific prey."
---

# BioID repeat / cross-experiment comparison

`scripts/bioid_repeat.py` takes **pairs** of protein-level workbooks (the same bait measured twice) and
writes one self-contained HTML report: one section per pair with a correlation scatter, a summary, and
the proteins that only one side detected.

Typical inputs are `bioid-batch-qc` outputs (`<batch>/flagged/*_flagged.xlsx`, which carry
`Norm_LOG2FC` and the batch-QC columns) or `bioid` outputs (`*_processed.xlsx`). Both are detected
automatically. Sources are never modified.

## Required inputs

- Two workbooks per comparison, one per experiment. Any pairing is allowed: different batches
  (`BioID-2nd/flagged/21FAM214B_flagged.xlsx` vs `BioID-4th/.../9AKAP8L_processed_flagged.xlsx`),
  different cell lines, or two runs of the same line.
- A value column: `Norm_LOG2FC` (batch-QC, default), `LOG2FC`, or a linear `FC` column. The script
  reports which column it used per file.
- The **same bait** on both sides. The tool cannot verify this — it warns when the two file labels
  differ, but the check is the user's (see the confirmation step below).

## Workflow

1. Write the pairing table (one line per comparison, tab or comma separated, `#` comments allowed):

   ```text
   # 标签	文件A	文件B
   ATOSB: 293 vs MCF7	/path/BioID-2nd/flagged/21FAM214B_flagged.xlsx	/path/.../10ATOSB_processed_flagged.xlsx
   AKAP8L: 1st vs 4th	/path/BioID-1st/flagged/1AKAP8L_flagged.xlsx	/path/.../9AKAP8L_processed_flagged.xlsx
   ```

   The label is what the report shows and is also used to find the bait (text before `vs`, minus any
   leading order prefix) so it can be drawn as a purple diamond. If a pair has no label, one is derived
   from the file names, using the upper directory level when both sides reduce to the same gene.

2. **Show the pairing to the user before rendering** (same rule as `bioid-analysis`):

   ```bash
   # python 需要 numpy + openpyxl（python3 -c "import numpy, openpyxl" 自检）
   PY="${PYTHON:-python3}"
   SKILL="${SKILL_DIR:-$HOME/.dsh/skills/bioid-repeat}"
   "$PY" "$SKILL/scripts/bioid_repeat.py" --pairs pairs.tsv --list
   ```

   `--list` reads both files, prints how many proteins each holds, how many match by accession, and
   warns when the two labels differ. Nothing is written. Get an explicit yes, then re-run without
   `--list`.

3. One-off pair without a table: `--pair FILE_A FILE_B` (repeatable).

4. Read the report, then optionally hand the same workbooks to `bioid-analysis` for the full hit list.

Useful flags: `--value norm_log2fc|log2fc|fc` (default `norm_log2fc`), `--label-top N` (initial label
budget, default 8), `--out DIR` (default `<folder of first file>/bioid_repeat/`),
`--plotly inline|cdn` (default `inline`, offline-readable).

## The report

- **Figure PDF**: the plotly modebar carries a **vector** PDF button (same implementation as
  `bioid-analysis`: `Plotly.toImage('svg')` → jsPDF + svg2pdf.js, with an inlined Noto Sans SC subset
  for CJK) that downloads the figure alone as `<pair>_scatter.pdf` — lines stay paths, labels stay
  selectable text, no embedded bitmap. The raster path is kept as a fallback if the libraries are
  missing.
- **Layout**: the page is not width-capped and the scatter is sized `clamp(520px, 72vh, 880px)` tall,
  so it reads near-square on a normal window instead of as a flat strip.
- **One tab per pair** (the same components as `bioid-analysis`: a tab bar over per-pair panels). When
  printed or exported to PDF each pair starts on a new page and takes two pages; **all tabs are
  rendered before printing**, otherwise unopened tabs would print blank.
- **Mark proteins, multi-select**: type a gene or accession and click **several** suggestions; each
  one queues in the "待加入…" chip row (click × to drop one). Pick the colour with the toolbar's 颜色
  swatch, then "加入标记" commits the whole queue as one mark group — the same flow as `bioid-analysis`
  (pending list → colour → Add mark). Rows selected in the table join the same queue with "把选中的行
  加入待加入", and **the row selection is cleared when committed**, so a second batch never carries the
  first batch's rows. Marks are stored per pair in localStorage, drawn as large coloured dots, and
  their names are labelled first. A mark on a protein that is not plotted says so in its chip instead
  of producing a label without a point.
- **Label 数量**: how many automatic names to draw (0 = none, default from `--label-top`). Names that
  cannot be placed without overlapping are dropped, and the toolbar says how many were drawn of how
  many were wanted.
- **Point classes**: filled = measured on both sides (these carry the correlation); hollow circles =
  one side's control was never detected, so the ratio is a floored lower bound; hollow squares pinned
  to the axis floor = the protein is present in both files but one side has no value (the square's
  position is a convention, not a number); the purple diamond is the bait.
- **One table, not four**: every row of the comparison (matched, one-side-only, unquantified), with
  columns 蛋白 / Accession / A 值 / B 值 / Δ (A−B) / 类别 / unique 肽 / 覆盖 % / 备注. Sort by clicking a
  header, filter with the dropdown (all / both measured / one-side unquantified / only-one-side /
  batch-QC flagged), scroll inside the box, click a row to pick it. It holds all rows on screen and the
  first 18 when printed, with the caption saying so.

## What the numbers mean

- **r is computed on measured pairs only.** Prey whose control was never detected carry a floored
  (lower-bound) ratio; they are drawn as hollow markers and excluded from r, from the slope, and from
  the "only in one side" counts. They are still shown, because "in the bait, absent from the EV" is
  the strongest single-run evidence a BioID design offers.
- **Spearman r** is the headline: it is rank-based, so it survives the different dynamic ranges of two
  experiments. **Pearson r** and the **slope/intercept** of the y-on-x fit are reported next to it; a
  perfect repeat has slope 1 and intercept 0.
- **Read r as reproducibility, not as a p-value.** Two single runs per bait give no replication
  structure: r ≈ 0.6 across cell lines means "the same proteins tend to come up", not "the difference
  is significant".
- **Never compare the absolute fold changes across batches.** Each batch is load-corrected on its own;
  `Norm_LOG2FC` is additionally centered within its own workbook. The plot is about relative
  agreement.
- **Labels are placed by "walking outward" from the dot, not by an ink grid.** The sibling skill
  scores 56×56 ink cells; that grid is ~30×11 px per cell here, so a box lying on dots could still
  score 0 and names ended up floating on the cloud with long leaders. This skill instead:
  1. *Direction order* — `CALLOUT_DIRS` tries the directions perpendicular to the band (up-left,
     down-right) first, then horizontal, then the ones that run along the band. Both axes are the
     same quantity here, so the data is one tight diagonal band and only the perpendicular directions
     leave it.
  2. *Search* — `CALLOUT_STEPS` walks outward in eleven steps (1 … 19 × about 11 px) and takes the
     cheapest candidate, so a name lands at the nearest free spot instead of at a far one chosen by a
     scoring artefact.
  3. *Cost* — dots within 5 px of the box (exact count, ×12 — one dot is worth 12 px of extra leader),
     leader length (1.0/px: 1 px of leader costs as much as one dot under the text), a leader crossing
     another leader or passing through another name's box (80), a leader tip landing within 18 px of
     another tip (60), a leader crossing a guide line (30).
  4. *Hard rules* — the box stays inside the axes, keeps 11 px / 7 px clear of another name's box,
     never contains its own dot, never contains a dot that already carries a name (bait or marked), and
     the leader starts at the dot's edge (5 px) and stops 3 px short of the text, so every line visibly
     touches the dot it belongs to. (The sibling skill starts its leader 22 % of the way along, which
     looks detached as soon as a name is pushed out.)
  Automatic names refuse a box with more than 3 dots within 5 px or a leader longer than **45 px**;
  bait and marked proteins are always drawn (they were asked for) and may go to **60 px**, falling back
  to a longer leader rather than losing the name entirely. Raising "Label 数量" really does draw more
  names — 0 means none (the bait keeps its name), and the note prints how many of the wanted names
  actually fit. Label halos are semi-transparent (`rgba(255,255,255,.6)`) so the cloud shows through.
- **Re-plan the names whenever the view changes.** Leader length is measured in pixels, so it is only
  meaningful for the layout it was planned in: a figure exported while zoomed 1.7× rendered every
  leader 1.7× longer than intended (the user's complaint "引线还是有太长的问题" at 34 names). Labels are
  therefore re-planned on `plotly_relayout` (zoom/pan, debounced 220 ms) and on window resize (350 ms),
  against the *visible* axis range, with the current range written back into the layout so the user's
  zoom survives. Two traps when touching this: bind the handler only after the first `Plotly.newPlot`
  (plotly attaches `div.on` in newPlot — calling `gd.on` earlier throws `gd.on is not a function` and
  the plot never appears), and never let the re-plan reset the axis range.
- **Verify a callout change on the exported figure, not the screen.** The PDF button exports exactly
  what is on screen (`Plotly.toImage` at `_fullLayout.width × height`, scale 2), and the report writes
  the same figure as page 1 when printed — render that image and look at it before believing the
  screen render; in a crowded pair (18 marks + 15 automatic names) it is the only view where long
  leaders and crossed lines are obvious. Three measured baselines to compare against: the default
  8-name layout has median leader 4–7 px / max 43 px, the extreme (250 names wanted) has median 23 px
  with nothing over 43 px, and the user's marked scene (18 marks + 15 names, exported at 1.7× zoom)
  has max ~55 px versus ~145 px before the zoom fix.
- **Check the accounting when a point seems missing.** Every matched protein lands in exactly one of:
  plotted, one-side-unquantified, or both-unquantified — the note under the plot prints this identity
  (`可匹配 N = 图上点 + 一侧未定量 + 两侧均未定量`) so nothing can quietly disappear.
- **"Only in A / only in B" is depth-sensitive.** A file with more quantified proteins contributes
  many more one-sided hits (1,063 vs 121 in the AKAP8L example) — check the per-file protein counts
  printed by `--list` before reading that list as cell-line-specific biology.

## Interpretation rules

- Say plainly whether the pair repeats: r ≥ 0.8 good, 0.5–0.8 moderate, < 0.5 weak. Quote the number
  and the counts it rests on (`n` used, censored excluded).
- When the two experiments disagree on a protein the user cares about, check the hover values and the
  censored flag before calling it a difference — a lower bound can move a point a long way.
- Cell-line-specific prey (present in one side only) are worth reporting per line, but rule out depth
  and the batch-QC background flag first.
- The bait itself is drawn as a diamond; if it is missing from a side, say so — that is a bait-recovery
  problem (see `bioid-analysis`), not a repeat problem.

## Bundled resources

- `scripts/bioid_repeat.py` — the tool.
- `assets/report_template.html` — the report template (plotly is inlined from `assets/plotly.min.js`).
- `assets/plotly.min.js` — vendored so reports stay readable offline.
- `assets/uniprot_annotations.tsv` — the same UniProt annotation cache `bioid-analysis` uses
  (37.7k accessions: name / gene / aliases / function); the report embeds only the accessions that
  appear in the pairs, and the hover card shows the protein name, "also known as" aliases and the
  function paragraph. Override with `--annotations <tsv>`; a different cache in the same column format
  works.
