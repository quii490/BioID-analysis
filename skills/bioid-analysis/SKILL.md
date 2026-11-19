---
name: bioid-analysis
description: "Score protein-level BioID/AP-MS workbooks into a reviewable interactor list: CRAPome contaminant filtering, a SAINT-style Bayesian score with a calibrated q-value, a 1.5x enrichment gate, a >=2 unique peptide gate, bait/control load correction, and a self-contained interactive HTML scatter report. Use when the user asks to score BioID hits, filter common contaminants, rank candidate interactors, or produce a visual enrichment report from proteomics workbooks."
---

# BioID analysis

Turn one protein-level workbook per bait into a ranked, filtered candidate list plus an interactive
report. `scripts/bioid_analysis.py` does the work; the source workbooks are never modified.

Use this skill for **scoring and review of a hit list**. Use the sibling skill `bioid-batch-qc` for
batch-level QC (background frequency across baits, sample-similarity and carryover screening). They
read the same workbooks and complement each other; quoting the across-bait background evidence
needs `bioid-batch-qc`.

## Required inputs

- One `.xlsx` per bait in a folder, each with a `Proteins` (or `+5%`) sheet.
- Columns: `Accession` (UniProt), `GN`, `Description`, `Coverage [%]`, `# Peptides`, `# PSMs`,
  `# Unique Peptides`, and two raw `Abundance:` columns — **experimental first, paired control
  second** (do not use `Abundances (Normalized):`).
- The three columns `bioid-batch-qc` appends — `Combined_score`, `Batch_BG_flag`,
  `Cross_sample_QC`. They are carried into the candidate table, the hover card and the scatter
  (background-flagged prey are drawn in a separate colour that the legend can hide). **Run
  `bioid-batch-qc` first**: without its output these three columns are empty, and the report is only
  a scoring sheet, not a QC'd hit list.
- Proteome Discoverer reports `# Unique Peptides` for the experiment rather than per channel. The
  gate therefore uses it as a proxy; say so when reporting.
- **Read the raw columns, not a pre-processed export.** The pipeline takes the two raw
  `Abundance:` columns and ignores anything downstream of them (`Abundances (Normalized):`,
  `LOG2FC`). It *does* use `bioid-batch-qc`'s output when it exists: point it at the folder that
  contains `flagged/` (the default), or at the QC output itself with `--from-qc`; the `_flagged`
  suffix is stripped so condition names match. Use `--no-qc` to force reading the folder's own
  workbooks instead.
- **No QC output is a hard error by default** (the old behaviour silently produced a report without
  Background flag). Two escape hatches: `--allow-no-qc` (accept a report with the three QC columns
  empty) and `--allow-partial-qc` (accept a `flagged/` that covers only some workbooks). If the
  `flagged/` set is missing conditions that exist in the folder, the run stops and names them.
- **Who imputed the control channel?** Point the skill at raw workbooks and it handles the censored
  case itself: an empty EV cell gets this skill's own 5th-percentile floor and is labelled
  `FC_lower_bound; Control_missing`, excluded from σ and from every p/q. But if an *upstream*
  step already substituted a floor, the cell holds a number and the skill cannot tell it from a
  measurement — those rows keep `FC_type = Measured` and do take part in σ and the p/q values,
  so its own labelling never fires. The run prints a WARNING when it sees the signature (one
  control value repeated across many rows), which is the check to read before comparing this
  report with another tool's. Ask the user which convention their preprocessing uses.

## Workflow

0. **Confirm the bait with the user before analysing.** If the file-name label is a construct,
   mutant or legacy name rather than a gene symbol (`ATOSB-C`, `FAM118B-L32R`, `FAM214B`), also ask
   the user for the `<label> <gene>` mapping and pass it as `--bait-alias`; without it the bait is
   neither excluded nor marked. `bioid_run.py` builds and asks for that table at its gate. The bait is inferred from the filename, and a
   wrong inference is silent: the real bait is never excluded, and the wrong gene is dropped from the
   candidates (this happened with `MCF1-1`, inferred as `NCOA3` when the construct is `AKAP8L`). Run
   `--list-baits` first: it prints, per workbook and without writing anything, the inferred label, the
   channels, whether that label is actually detected (with peptide evidence), and the top
   bait-exclusive/≥5x-enriched proteins as evidence. Show that table to the user and get an explicit
   yes. To correct a label, either rename the workbook to `<order><GENE>_processed.xlsx` or pass
   `--bait-map <tsv>` (lines `<filename or condition><TAB><GENE>`) — the map is per run, so the
   correction is recorded rather than hidden in a filename. When the label is a **legacy or construct
   name** that the database now reports differently (`FAM214A`→`ATOSA`, `FAM214B`→`ATOSB`,
   `mKIAA1614`→`KIAA1614`, `SPANXA2`→`SPANXA1`), use `--bait-alias <tsv>` instead: rows are
   `<label><TAB><database symbol>`, the alias only feeds bait exclusion and the bait marker, and your
   own condition names stay unchanged.
0b. Run this stage **after** `bioid` (preprocessing) and, for the QC columns, after `bioid-batch-qc`.
   If `<folder>/flagged/` exists, analysis auto-detects the `*_flagged.xlsx` workbooks there and says
   which input it used; `--from-qc` demands them and `--no-qc` ignores them. The `_flagged` suffix is
   stripped when naming conditions, so a condition keeps the same name across both inputs.
1. Inspect one workbook to confirm the sheet name, the column names above, and which channel is
   which. If a workbook has no raw `Abundance:` columns, the script skips it and lists it at the
   end of the run rather than failing the batch.
2. Run the pipeline. **The whole chain in one command** (preprocess -> batch QC -> scoring ->
   verify) is `bioid/scripts/bioid_run.py`; it runs the bait check for you and refuses to produce a
   report without QC:

   ```bash
   python3 ~/.dsh/skills/bioid/scripts/bioid_run.py "/absolute/path/to/raw_batch" \
       --out "/absolute/path/to/run" --yes
   ```

   Or run this stage alone (needs numpy + openpyxl; `python3 -c "import numpy, openpyxl"`):

   ```bash
   PY="${PYTHON:-python3}"
   SKILL="${SKILL_DIR:-$HOME/.dsh/skills/bioid-analysis}"
   "$PY" "$SKILL/scripts/bioid_analysis.py" "/absolute/path/to/BioID_batch"
   ```

   Outputs land in `BioID_batch/bioid_analysis/` (with `--smoke` and no `--out`:
   `BioID_batch/bioid_analysis_smoke/`, so a template-iteration run never overwrites the real
   report). Flags: `--from-qc` (read the batch-QC output
   `flagged/*_flagged.xlsx` instead of the folder's own workbooks), `--no-qc` (ignore a `flagged/`
   folder even when it exists), `--allow-no-qc` / `--allow-partial-qc` (see above), `--min-fold 2`,
   `--min-unique-peptides 2`, `--normalize total|median|robust|none`, `--out <dir>`,
   `--sheet <name>` (default `Proteins`, falls back to `+5%`),
   `--plotly inline|cdn` (default `inline`: self-contained and offline),
   and for the contaminant table and the hover annotations, all of which ship inside the skill:
   `--crapome-table <path>`, `--crapome-frequency 0.2`, `--crapome-min-sc 2`,
   `--crapome-cell HEK293|all`, `--annotations <path>`.

   **Iterating on the report template? Use `--smoke`**: it analyses the first condition only,
   seconds per run instead of minutes, which is enough for every template change. Verify with the
   smoke output (`node tests/verify_report_ui.js "file://<out>/BioID_analysis_report.html#1" out.png`),
   and only then re-run the whole batch. Regenerating a document for every template tweak is the
   mistake to avoid.

   **A run writes no PDF, and needs no browser.** The HTML is the report. A `--pdf` flag used to
   drive headless Chrome to print every condition into a multi-page document; it was removed as an
   unnecessary slow step — it was the most expensive part of a run, and the need it was built for
   ("export a plot as PDF") is served by the PDF button in each figure's modebar, which needs no
   flag. If someone wants a document, they print the open panel from the browser (Cmd-P → "Save as
   PDF"); that path is kept working and looks right.
3. Read the run output before opening anything: per condition it prints the load ratio and
   correction, the candidate count under each estimator, the CRAPome count, the top candidate, and
   every warning. A warning means the number next to it is not what it looks like.
4. Open `BioID_analysis_report.html`. It is self-contained (plotly.js embedded): check the
   Overview, then per condition the scatter, the candidate table, and the **ratio noise σ** and
   **share above gate** cards. Hovering a point shows the gene, UniProt accession, full protein
   name, fold changes, peptide evidence, score/q and why the protein passed or failed; the scatter
   also carries marginal abundance histograms and the candidate table lists the protein name. Hovering shows a
   GeneCards-style entry — gene, accession, full protein name, aliases and a short function summary from the
   cached UniProt annotation (see `references/annotations.md`).

   The plot area is sized near-square (a log-log scatter is read along its diagonal), markers are small and
   semi-transparent so a few thousand prey do not merge, and candidate names are lifted off their markers so
   text and dots do not collide.

   Every drawn name — candidate or marked — is a callout: a thin diagonal leader from the dot to the name,
   with a gap before the text, directions rotating through four diagonals so a cluster of labels fans out
   instead of stacking. Marking proteins: type into the search box for autocomplete suggestions, or pick
   rows in the candidate table — "Add mark" recolours those dots in the chosen colour and their names
   appear once, in that colour (a marked prey is dropped from the candidate labels, so no double names).
   Marks are gate-independent, so a contaminant or a below-cutoff protein can be followed too, and they
   persist across page reloads (kept in the browser's local storage per batch and condition).

   Zoom: the mouse wheel scrolls the page by default (so reading does not hijack the plot); drag a box to
   zoom, double-click or "Reset zoom" to undo, and a "Wheel" selector switches the wheel between scrolling
   and zooming if wheel-zoom is preferred.

   Label by: the gates decide who may be a candidate, and this selector decides the order their names are
   drawn **on the plot** when more pass than the label limit (the candidate table below is always ordered
   by normalized fold change, and the workbook matches it). `evidence tier` (default) puts prey the control never
   detected first — their evidence is a censored observation, ordered by their own floor — then everything
   measured, by the SAINT-style score when the mixture fitted and by the calibrated q-value when it did not.
   `SAINT-style score`, `q value`, `fold change` and `abundance` rank on that single quantity instead. The
   note under the toolbar names the criterion in words, and says so when the top names are one big tie
   (a condition with no signal), because then the order is arbitrary and the names mean nothing.

   Mark colour: the swatch on each mark chip is a colour input — click it to recolour an existing mark's dot,
   its name and its leader line, without deleting and re-adding the mark.

   `Methods & caveats → How the numbers are computed` holds the definition of every term (raw FC vs normalized
   FC, the load correction, the candidate criteria, and where the annotation comes from); the on-page prose is
   deliberately kept terse.
5. Validate before handoff: source workbooks unchanged, output confined to `bioid_analysis/`
   (or `bioid_analysis_smoke/` for `--smoke` without `--out`),
   candidate counts consistent with the estimator note, no claim that the 1.5× list is significant.

## What the numbers mean

- **Load correction.** The bait and control channels are compared by total protein abundance;
  `median` (default) removes the scale difference robustly, `total` matches summed signal (right
  when the channels are known to have carried different amounts), `robust` restricts the median to
  well-quantified proteins. All three, plus the top-10 signal concentration and the candidate count
  under each, are reported — when they disagree, the shortlist depends on the choice, and the report
  says so.
- **Fold-change gate ≥ 1.5×** and **≥ 2 unique peptides** are the two hard gates, plus CRAPome
  exclusion and bait exclusion. The baited protein itself is never a candidate, is drawn as its own
  purple diamond, and is never counted in the candidate card — its own-channel enrichment is the assay
  working, not an interaction.
- **Prey the control never detected are kept, as lower bounds.** "In the bait, absent from the EV" is
  the strongest single-run evidence a BioID design offers, and dropping those rows hides the very class
  the experiment is looking for. Each such prey is given the 5th-percentile control floor of its own
  workbook (the same convention as `bioid-batch-qc`), and its fold change is reported as
  `FC_lower_bound; Control_missing` — a floor, never a measurement. Because the gate is applied to the
  lower bound, passing it is a conservative statement. These prey take **no part** in the background
  spread σ, in the mixture fit, or in any p/q value: a censored ratio is not a test, and letting one in
  would both inflate σ and consume a test slot in the BH correction. They are drawn as hollow teal
  triangles at the floor they were compared against, listed as `EV: not detected` in the table, and
  counted on their own card.
- **Guide lines are drawn in the corrected frame.** A guide at k× sits at the raw ratio
  k × 2<sup>correction</sup>, which is the ratio the gate was actually applied to, so the drawn line
  bounds exactly the prey whose colour it decides; the 2σ/3σ null lines are placed the same way. A faint
  `1:1 raw` diagonal is kept only to show how far the load correction moved the decision lines. Without
  this, an uncorrected 1.5× line can look like it contradicts the green dots below it — in this batch up
  to 93% of the passing prey sit under an uncorrected 1.5× line.
- **SAINT-style score** is the posterior probability of belonging to an enriched mixture component,
  fitted per condition on normalized log2 ratios with the null taken from the negative half of the
  distribution and a sparsity prior on the mixing proportion. **BFDR** follows Newton et al. (2004).
- **q-value** is the calibrated companion: a one-sided p-value from the same empirical null with
  Benjamini–Hochberg across quantified prey. This is the number that answers "is this surprising
  given the noise", and it is usually far stricter than 1.5×.
- **One ordering everywhere.** The workbook's `Candidates` sheet, the per-condition sheets and the
  report's candidate table are the same list in the same order: **normalized fold change,
  descending, with EV-undetected prey mixed in** (their FC is a floor, shown with ≥; raw FC and
  abundance break ties). The user asked for exactly this — no separate tier at the top of the table.
  The SAINT-style score stays a column, not the sort key: on batches where no mixture fits, score
  ordering degenerates to source-file order while FC is always meaningful. The plot's *labels* are a
  separate choice governed by the "Label by" selector. If a change touches one of these orderings,
  change all of them — a list that reads as two different rankings between deliverables is worse
  than either ranking.

## Interpretation rules

- **A 1.5× cutoff is a filter, not evidence.** On these single-run batches the ratio spread is
  σ ≈ 1.3–2.1 log2, so 20–43% of proteins clear 1.5× by chance. Always read the "share above gate"
  and "min fold at q = 0.05" cards with the candidate list, and never present the list as
  significant.
- **This is not the official SAINTq.** SAINTq scores peptide/fragment intensities; protein-level
  aggregation is the SAINTexpress-equivalent path its own paper reports as FDR-underestimating.
  One bait run plus one shared control gives no replicates, so scores are prioritization aids.
- Where the mixture did not fit, the report says so; do not quote a score for that condition.
- Say what would change the conclusion: bait replicates, independent EV runs, or a peptide-level
  export (which also makes the unique-peptide gate channel-specific and allows real SAINTq).
- Do not delete anything. Contaminants and failed gates are annotated, and the report keeps them.

## Bundled resources

- `scripts/bioid_analysis.py` — the pipeline (needs numpy + openpyxl only).
- `assets/report_template.html` — report shell; `assets/plotly.min.js` is vendored so reports open
  offline (MIT licensed).
- `assets/crapome_frequency.tsv` — compact CRAPome control frequencies; see `references/crapome.md`.
- `assets/uniprot_annotations.tsv` — protein names, gene synonyms and function summaries for the hover cards;
  see `references/annotations.md`. Refresh with `scripts/fetch_uniprot_annotations.py`.
  It ships the reviewed **human** and **mouse** proteomes plus every accession seen in the batches,
  so a run on either species needs no network.
- `scripts/fetch_crapome.py`, `scripts/build_crapome_table.py` — rebuild that table from the
  official files.
- `scripts/fetch_uniprot_annotations.py` — extend the annotation cache when new batches arrive
  (fetches only what is missing).
  (The two browser-driven test helpers below need Chrome; the pipeline itself does not.)
- `tests/capture_report_hover.js` — opens a report in headless Chrome via the DevTools Protocol,
  hovers the top candidate with a real mouse event and screenshots the page, so the tooltip content
  can be verified without a display or a Playwright install.
- **PDF output is per figure, by design — and it is a real vector PDF.** Each scatter carries a
  **vector PDF** button in its modebar (it replaces plotly's PNG camera). One click downloads that
  figure alone as a vector PDF: `Plotly.toImage('svg')` → jsPDF + svg2pdf.js → PDF, so lines stay
  paths and labels stay selectable text (0 embedded bitmaps; opens editable in Illustrator). Both
  libraries (MIT) are inlined in the report, so this needs no network and no Chrome. CJK labels work
  because a **Noto Sans SC subset** (SIL OFL, ~520 KB, built from the lab's own Noto TTF) is inlined
  and registered with jsPDF, and the SVG's `font-family` is rewritten onto it — without that the
  text silently falls back to Helvetica and Chinese comes out as garbage. The old
  rasterise-then-wrap path (`figurePdfBytes`, a hand-written one-page PDF holding a PNG) is kept as a
  fallback for reports whose libraries were stripped. A whole-report document still comes from
  printing the browser view, not from the pipeline. When someone asks for a report PDF, say that the
  HTML is the primary artifact and the workbook list is always complete.
- `tests/verify_report_ui.js` — 33 runtime checks on the interactive report (autocomplete, table
  picking, mark colouring, recolouring and labels, hover annotation, zoom-aware relabelling,
  gate-driven recounts, the label-by selector, plot geometry), driving a real browser and
  screenshotting the result. Three of them are the invariants that are easy to break and hard to see:
  the drawn gate line must sit at `minFold × 2^correction`, no prey coloured as passing may sit below
  it while no prey coloured as failing may sit above it, and the candidate card must equal the
  pipeline's own candidate count (bucket classes must not count rows the gates excluded). Run it after
  touching the template; it catches failures a static render cannot.
- `tests/check_pdf_layout.py` — objective layout checks on a PDF *printed from* the report
  (Cmd-P → Save as PDF; the pipeline no longer writes one): no text drawn
  on other text (with the thresholds tuned so rotated tick labels and glyph side bearings
  do not count), no text past the page edge, no near-blank page, and every candidate-table
  column heading complete, and — given `--payload analysis_payload.json` — every caveat
  box states its own page's condition, not the next one's (the box identifies its
  condition by the candidate count in the caption and checks the load correction the
  bullets quote against the payload). These are the defects this report has actually
  shipped, and each one is measurable, so the check runs on every batch instead of only
  on the page a reader happens to look at. Run it with `python3 tests/check_pdf_layout.py
  <printed.pdf> --payload <analysis_payload.json>` whenever a PDF is produced from the
  report, and after any change to the print CSS.
- `tests/make_synthetic_and_check.py` — ground-truth test, four cases: planted interactors with
  contaminants and one-peptide proteins; a 4× channel-scale difference that the load correction must
  absorb; a genuinely null dataset that must yield no discoveries; and prey the control never detected,
  which must be kept, labelled `Control_missing`, compared against the workbook's own floor, left out of
  every p/q value, and must not inflate the background spread. Run it after any change to the pipeline.
- `references/methodology.md` — model definitions, the measured specificity problem, checklist.
