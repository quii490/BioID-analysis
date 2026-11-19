---
name: bioid
description: "Preprocess raw protein-level mass-spec exports (Proteome Discoverer / MaxQuant style .xlsx) into BioID workbooks: explicit bait-to-EV channel mapping, protein-level and contaminant filtering, log2 enrichment, a 5th-percentile control floor for prey the EV never detected, and the +5%/Proteins/Raw sheets. Use when asked to process a BioID batch from raw exports, add samples to a batch, rerun the preprocessing step, or when a downstream BioID analysis is missing its input workbooks."
---

# BioID batch preprocessing

`scripts/bioid_process.R` turns one raw protein-level export per bait into one processed workbook per
**bait × EV** comparison. Source workbooks are never modified.

Pipeline position: raw export → **this skill** → `bioid-analysis` (scoring, hit list, report) and
`bioid-batch-qc` (cross-bait background and mixing QC). Both of those read the `Proteins` / `+5%`
sheets this skill writes, so run this stage first.

## Required inputs

- One `.xlsx` per bait, either a folder or a single file. A raw single-sheet export is expected; a
  workbook that already contains a `LOG2FC` column is treated as an already-processed product and
  skipped (override with `--force` only when you really intend to reprocess it).
- A `channel_map.tsv` in the input folder (or passed with `--map`), one rule per bait channel. **This
  is not optional and not inferable** — the script never guesses which column is the bait.
- R with `readxl` and `openxlsx`.

## The channel map is the one thing you must get right

```text
<bait channel>   <EV channel>[|<EV channel>...]   [file-name glob]
F4               F1|F2
F8               F6
F12              F11
```

Channel names are the `F<n>` token in headers like `Abundance: F4: Sample`. Write one rule per bait.
Writing several EVs makes the script emit one workbook per EV, because each comparison needs its own
control column and its own 5th-percentile floor.

**`bioid_run.py` can draft this file for you.** When neither `--channel-map` nor
`<input>/channel_map.tsv` exists, it reads every workbook's `Abundance:` columns, looks up the
filename-derived gene, and writes `<out>/00_channel_map.tsv` — one rule per file, with the evidence
(bait abundance and its ratio to the control) in the comment above each rule — then **stops for
confirmation**. Confidence is `high` when the label matches a GN exactly, `medium` when only a prefix
matches (construct names like `ATOSB-C`), `low` when no gene name could be derived (it then falls back
to "the channel that appears in the fewest files", which is the shared control's opposite). If the
gene is higher in a channel that recurs across files (the shared control) than in the file-unique one,
the draft assumes the columns are reversed: it picks the unique channel and says so in the evidence. The draft
is never used silently: rerun with `--yes`, or edit the file first (edits are reused). A hand-written
`channel_map.tsv` in the input folder, or `--channel-map`, always wins over the draft.

**Construct names must be confirmed, not guessed.** When a file-name label is not an exact gene
symbol (`ATOSB-C`, `FAM118B-L32R`, `EEF1AKMT4-Y39W41A+F204206A`, a legacy name), the same gate also
writes `<out>/00_bait_aliases.tsv` — one `label<TAB>gene` per line, taken from the protein that
carries the bait enrichment — and prints it under "需要你确认的 构造名 → 基因名". That table is what QC
(bait recovery, bait transfer, direction check) and analysis (bait exclusion/marking) need; without
it, QC reports every construct as "Bait recovery QC failure". **Ask the user to confirm or correct
that mapping before running with `--yes`**; the confirmed table is then passed to both downstream
stages automatically. Exact gene-symbol filenames produce no alias table and no question.

`assets/channel_map.tsv` currently holds one batch's map as a worked example. Treat it as an example,
not a default: for a new batch, derive the map from the sample sheet, or — when only the files exist —
from the data itself. The reliable check is that **the shared control appears in many files while the
bait channel is unique to its file**: correlate log2 abundance per channel across the batch, confirm
that a channel used as the experimental column in one file is not the same physical sample as another
file's bait channel, and where possible confirm the bait protein is detected only in its own channel.
Getting this wrong is silent — a wrong map still produces tidy workbooks — so confirm the printed
bait→EV table against the sample sheet, and treat "channel_map 未覆盖" skips as a signal that the
file's channels need a rule.

## Workflow

1. Inspect one input workbook: sheet name, the `Abundance: <channel>: Sample` headers, and whether
   peptide-level rows are stacked below the protein rows (Proteome Discoverer exports often are).
2. Write `channel_map.tsv` for the batch.
3. Run. **The whole chain in one command** — preprocessing, batch QC, scoring, and a final
   verification — is `scripts/bioid_run.py`. It is the recommended entry point because the two
   downstream stages are otherwise easy to skip (a report without batch QC has no Background flag):

   ```bash
   python3 ~/.dsh/skills/bioid/scripts/bioid_run.py \
     "/absolute/path/to/raw_batch" --out "/absolute/path/to/run" --yes
   # layout: 01_processed/  02_qc/*_flagged.xlsx + BioID_batch_QC.xlsx  03_analysis/  logs/
   #   --resume   skip a stage whose outputs already exist (analysis only if the last run
   #              passed verification AND the analysis arguments/inputs are unchanged)
   #   --dry-run  print the commands only
   #   --skip-preprocess / --skip-qc / --skip-analysis  stop the chain where the debugging needs it
   #              (--skip-qc requires --skip-analysis: a QC-less report is not produced by the runner)
   #   exit: 0 ok, 2 usage/preflight, 3 bait confirmation needed (rerun with --yes), 4 stage failed, 5 verification failed
   ```

   Pass `--bait-alias <tsv>` when the file names are constructs (`FAM118B-L32R`) rather than
gene symbols: the runner forwards it to **both** batch QC (bait recovery, bait transfer, direction
check) and analysis (bait exclusion and marking). Without it, QC reports every construct as
"Bait recovery QC failure" even when the bait is clearly detected.

Without `--yes` the runner prints the inferred bait per workbook (the `--list-baits` evidence) and
   stops before scoring: a wrong bait label is silent, so it must be confirmed once per batch.

   This stage alone (R needs `openxlsx` + `readxl`; `Rscript -e 'library(openxlsx); library(readxl)'`):

   ```bash
   SKILL="${SKILL_DIR:-$HOME/.dsh/skills/bioid}"
   Rscript "$SKILL/scripts/bioid_process.R" \
     "/absolute/path/to/BioID_batch" "/absolute/path/to/BioID_batch/processed"
   ```

4. Read the run output before opening anything: the bait→EV table, per-file kept/removed counts, and
   every ⚠. Exit codes: `0` clean, `2` some files were skipped (listed with reasons), `3` a structural
   problem in the output package.
5. Sanity-check one output workbook, then run `bioid-batch-qc` **before** `bioid-analysis`
   (`bioid_run.py` does this order automatically): QC appends the three columns the analysis report
   needs, and the analysis stage now refuses to run without them unless `--allow-no-qc` is given.

## What happens to each row

- Rows that are not protein-level are dropped explicitly (`Master` must be `Master Protein`; when
  present, `Reverse` and `Potential contaminant` flags are also honoured), so a stacked peptide table
  never reaches the metrics.
- Protein rows are dropped for an empty `Description`, a contaminant match, or a gene name that
  cannot be extracted. Every dropped row is written to `Removed` with its reason and its input row
  number — deletions are never silent.
- `GN` is taken from `Description` (`GN=...` up to the next space or semicolon).
- `LOG2FC = log2(bait / EV)`. Prey detected in the bait but not in the EV are `Unique`: they get the
  5th percentile of that workbook's positive EV values, are filled light green, and are the reason the
  `+5%` sheet exists. Rows with no bait value are kept but grouped and hidden.
- The abundance columns are reordered to **bait first, then this workbook's EV**, then the rest. This
  matters: `bioid-analysis` and `bioid-batch-qc` both take the first two `Abundance:` columns by
  position, so the reordered pair is what keeps them consistent with this skill.

## Outputs

One workbook per comparison: `<stem>_processed.xlsx`, or `<stem>_vs<EV>_processed.xlsx` when the bait
has several EVs (the trailing `_processed` is stripped by `bioid-analysis` when naming conditions).

| Sheet | Contents |
| --- | --- |
| `+5%` | Visible rows sorted by the floored LOG2FC; `Unique` rows carry the imputed EV value |
| `Proteins` | Visible rows: `Unique` first by bait abundance, then the rest by LOG2FC |
| `Raw` | The input sheet verbatim — every row, original column order. Recover anything filtered from here |
| `Removed` | One row per deleted row: `RawSheetRow`, `Accession`, `Description`, `GN`, `Reason` |
| `Methods` | The channel pair used, the rule that selected it, and every count |

A `bioid_run_log.txt` next to the workbooks records the map file, per-file counts and the R session.

The downstream tools read the **bait's identity from the file name**, and a wrong label removes the
wrong gene from the candidates silently. Name outputs after the gene (`<order><GENE>_processed.xlsx`),
or have `bioid-analysis` confirm the bait with the user (`--list-baits`) and override it with
`--bait-map` before the real run.

In `+5%` and `Proteins`, columns whose retained rows are numeric are written as numbers — the bait/EV
pair, `Coverage [%]`, `# Peptides`, `# Unique Peptides`, `# PSMs` — because these raw exports regularly
mix annotation text into those columns, and readxl then types the whole column as text, which Excel
cannot sort. The verbatim form is preserved in `Raw`.

Each run also reports both channel totals and their ratio per workbook, in the console and in the
`Methods` sheet, flagging a 2-fold or larger imbalance — that ratio is exactly what the downstream load
correction has to absorb.

## Caveats

- **`LOG2FC` is unnormalised.** No load correction is applied here; the bait and control channels are
  frequently loaded several-fold apart. `bioid-analysis` applies the correction and reports the load
  ratio. Never present this column as a normalised effect size.
- **The 5th-percentile value is a floor, not a measurement.** Treating `Unique` prey as passing an
  enrichment gate is a conservative statement; the downstream tools label them `Control_missing`.
- **Contaminant filtering here is only keratin** plus whatever `--contaminants` adds. The systematic
  filter is the CRAPome step inside `bioid-analysis`; do not report "contaminant-free" from this stage.
- **openxlsx writes two defects that break openpyxl**, which both downstream tools use (a relationship
  pointing at a drawing part that is never written, and a worksheet `dimension` of `A1` that makes
  read-only readers see a single column). The script patches both and verifies the saved package; if it
  reports structure problems (exit `3`), re-save the workbooks in Excel or rewrite them with
  `openxlsx2` before analysing.
- **One bait with several EVs is not a replicate.** `bioid-batch-qc` will flag such a pair as highly
  similar (`rho` near 1) because it compares the same bait channel; that is expected, not sample
  mixing. Shared EVs are one control sample, not one replicate per bait.
- Isoforms and legacy gene-name aliases: keep one file per bait and follow `bioid-batch-qc` for alias
  and isoform handling rather than editing gene names here.
