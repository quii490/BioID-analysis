---
name: bioid-batch-qc
description: "Reusable BioID mass-spectrometry batch analysis for protein-level Excel workbooks: compute normalized log2 enrichment and combined abundance/enrichment scores, collapse duplicate GN records for batch background frequency, color-code high/review background suspects, and screen cross-sample bait carryover or mixing. Use whenever the user asks to analyze a BioID batch, compare bait workbooks against an EV control, flag likely false positives, identify possible sample swaps/carryover, or produce marked Excel copies and a QC report."
---

# BioID Batch QC

Run a reproducible, protein-level BioID batch workflow that preserves source files and produces auditable marked copies. Use the bundled `scripts/bioid_batch_review.py` rather than rewriting the scoring logic.

## Required inputs

- A folder containing one `.xlsx` workbook per bait condition.
- Each workbook should contain `+5%`, `Proteins`, and `Raw` sheets, a `GN` column, `Accession` where available, and at least two raw columns whose headers begin with `Abundance:`.
- The first raw abundance column is the experimental bait channel; the second is its paired EV/negative-control channel. Do not use `Abundances (Normalized):` columns for the primary calculation.
- If the EV is shared (for example F12 or F14), treat it as one control sample, not as one independent replicate per bait. If there are multiple independent EVs, document the mapping before changing the method.

If the folder contains explicitly named isoforms or legacy GN aliases (for example, FAM172BP = ARB2BP), supply or confirm a family/alias mapping. Pass `--bait-alias <tsv>` (rows `<label> <gene>`, confirmed with the user) whenever the file-name label is a construct, not a gene symbol — e.g. `FAM118B-L32R → FAM118B`: without it, bait recovery reads "Bait recovery QC failure" for every construct and the bait-transfer matrix cannot be computed. The same table also feeds the bait-direction check. Treat isoforms as separate conditions for scoring and transfer analysis, but collapse them to one family only for background-frequency evidence. Use known aliases only for bait-recovery and bait-transfer QC; preserve the original GN labels. Do not make isoform-specific claims from GN-level rows without accession and isoform-specific peptide evidence.

Filename inference is automatic but never silent. Files that share a bait label (for example `18-2FAM227A.xlsx` with `18-10FAM227A.xlsx`, or `C19orf67.xlsx` with `C19orf67_processed.xlsx`) become separate `<gene>_isoform_<n>` conditions inside one family, and every run prints the inferred multi-condition families before anything else. Confirm that mapping before reading family-level flags or an isoform sheet. If two files would resolve to the same condition name, the script stops with an explicit error instead of letting one file's results overwrite the other's; keep one file per condition, or rename them with distinct `<order>-<isoform><gene>` prefixes.

## Workflow

1. Inspect the workbook structure and render at least one source workbook before editing. Confirm sheet names, raw abundance columns, hidden columns, existing styles, and whether bait proteins are recoverable.
2. Run the bundled script with the target folder as its argument:

   ```bash
   # python 需要 numpy + openpyxl；SKILL 指向本 skill 的安装目录
   PY="${PYTHON:-python3}"
   SKILL="${SKILL_DIR:-$HOME/.dsh/skills/bioid-batch-qc}"
   "$PY" "$SKILL/scripts/bioid_batch_review.py" "/absolute/path/to/BioID_batch" \
     [--out DIR] [--bait-alias aliases.tsv] [--exempt-genes "GENE1,GENE2"]
   ```

   The script creates `BioID_batch/flagged/`, never overwrites the source workbooks, and writes one `<stem>_flagged.xlsx` per input plus `BioID_batch_QC.xlsx`. The folder argument is required: conditions, bait families and bait genes are all inferred from the filenames, so there is no built-in default directory.

3. In each flagged `+5%` and `Proteins` sheet, keep the original columns and append the analysis columns. Show only `Norm_LOG2FC`, `Combined_score`, `Batch_BG_flag`, and `Cross_sample_QC`; hide and group the helper columns. Keep `Raw` unchanged. Add `QC_Methods` to each flagged workbook.
4. Review the QC workbook sheets: `Batch Overview`, `Background Candidates`, `Similarity Matrix`, `Pairwise Similarity`, `Bait Transfer`, one `<family> Isoforms` comparison sheet per family that holds more than one condition, `Family Notes` when no family does, and `Methods`. Read the `Methods` row `Threshold reachability` before reporting.
5. Validate all final workbooks: source values and hidden states are unchanged, output sheets open, expected appended columns exist, only the four requested analysis columns are visible, no formula-error strings occur, and a focused visual render is legible.

## Scoring rules

- For positive bait and EV values, calculate raw `log2FC = log2(bait / EV)`.
- Normalize within each workbook by subtracting the median raw log2FC among measured proteins: `Norm_LOG2FC = raw log2FC − median(raw log2FC)`. This is the primary effect-size comparison across baits.
- If bait is positive and EV is missing, use the positive EV 5th percentile only as a lower bound: `FC_lower_bound = log2(bait / EV_q05) − median shift`; label `Control_missing`. Never present this as a measured FC or p value.
- Rank positive experimental abundance and normalized enrichment within each condition, with the largest value at percentile 1. Use `Combined_score = sqrt(Abundance_pct × Enrichment_pct)` as a descriptive prioritization score, not a probability, p value, or FDR.
- Aggregate repeated accession rows at GN level for batch counts using the strongest valid row, then propagate the GN-level label to all matching original rows. Preserve the original rows rather than deleting duplicates.

## Batch background flags

Count both condition-level and family-level evidence. Family-level ranking uses the maximum score within each family, so two isoforms do not artificially double the frequency. Apply:

- Orange `High background suspect`: family combined top 2% in at least 4 bait families.
- Yellow `Review background suspect`: not orange, and either family combined top 5% in at least 4 families or family enrichment top 5% in at least 5 families.

These are review priorities, not automatic false-positive calls. Shared localization and genuine common neighbors can create repeated hits.

## Cross-sample QC

Do not call mixing from correlations of log2FC values that share the same EV denominator. Instead:

- Center raw experimental `log2(abundance)` for each prey by its median across bait conditions.
- Report pairwise prey-centered Spearman correlation and Top-100 specificity Jaccard.
- Pairs that share the same experimental channel (one bait measured against two different EVs) are reported as `Expected: same bait channel, different control` and are kept out of the threshold calibration — comparing them is tautological, and leaving them in inflates the similarity cutoffs used for every other pair.
- Build a bait-transfer matrix with `target abundance / source abundance` for each source bait gene. Use F-order adjacency only as `order-dependent provisional` evidence. When no ratio can be computed — no bait gene resolvable from the filenames, or the bait was never detected — the sheet title says so instead of presenting an empty matrix.
- Workbooks written by older `openxlsx` versions carry a dangling drawing relationship and a worksheet dimension of `A1`, which makes openpyxl either crash or see a single column. The script repairs such files on a temporary copy and says so (`缺陷包…已修复`); a file that still cannot be read is listed as `Skipped workbooks` at the end rather than aborting the batch.
- Classify low-level signals as unresolved, adjacent low-ratio signals as possible carryover, and larger nonadjacent or multi-source transfer as possible mixing/mislabeling or true proximity. Do not conflate these with batch background flags.
- Mark a condition `Bait recovery QC failure` when its bait gene is not detected in its own protein-level result. A missing bait is not proof of a sample swap; confirm with bait-specific peptides, tags, reporter/channel metadata, blanks, and injection order.

## Interpretation and reporting

Report the shared EV identity, number of conditions and families, orange/yellow counts, bait recovery failures, highest bait-transfer ratios, and highest non-family profile similarities. Explicitly label conclusions as provisional when acquisition order is assumed. State that protein-level Excel results cannot prove carryover or distinguish shared peptides from isoform-specific interactions.

The fixed background rules need at least 4 families holding a qualifying gene, and the enrichment route needs 5. A batch smaller than that cannot flag anything, so the run prints a WARNING and records the unreachable rules in the QC workbook: report that caveat with the results and never present an empty flag list as a clean batch. The same applies to the profile-similarity cutoffs — when they are not positive at the current batch size the run reports them as not estimable instead of labeling a negative correlation as elevated similarity.

For future replicated peptide-level experiments, recommend MSstatsTMT, SAINTexpress, or MiST only when their replicate/control assumptions are met; do not manufacture p values from a shared EV copied across bait files.

## Bundled resources

- `scripts/bioid_batch_review.py`: deterministic openpyxl/numpy implementation with a folder-path CLI. The folder argument is required; condition metadata is inferred from each workbook's first raw `Abundance:` column and from the filenames (see the script's `--help` epilog for the `<order><gene>[_isoform]` convention).
- `references/methodology.md`: compact rationale, thresholds, assumptions, and validation checklist to consult when a new batch has unusual controls, isoforms, or file layouts.
