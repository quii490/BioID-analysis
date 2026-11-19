# Changelog

## v1.0.0 — 2026-11-20

First release of the unified BioID analysis suite: five skills plus a one-command orchestrator.

### Added

- `bioid_run.py` (in `bioid/scripts/`): preprocess → batch QC → bait confirmation → scoring →
  verification, with `--resume`, `--dry-run`, `run_manifest.json`, per-stage logs and exit codes
  `0/2/3/4/5`.
- Automatic channel-map draft when no `channel_map.tsv` is supplied: per file it reads the
  `Abundance:` columns, finds the bait protein by the file-name label and writes
  `<out>/00_channel_map.tsv` with the evidence, then stops for confirmation. Hand-written maps and
  `--channel-map` always win.
- Construct-name gate: when file-name labels are not gene symbols (`FAM118B-L32R`,
  `ATOSB-C`, legacy names), the runner writes `<out>/00_bait_aliases.tsv` (`label<TAB>gene`) and asks
  the user to confirm it; the confirmed table is passed to both batch QC and analysis.
- `bioid-batch-qc --bait-alias` plus case-insensitive gene lookup, so construct-named batches get
  bait recovery, bait transfer and the bait-direction check instead of "Bait recovery QC failure".
- `bioid-trend`: grouped x-axis bands (`xgroup`), optional mean line (`--mean`, missing counted as 0),
  `--keep-all-na`, automatic CJK font switch, vector PDF + PNG + values/sources TSV.
- `bioid-repeat` and `bioid-trend` documented end to end (pair tables, sample/protein tables).

### Changed

- `bioid-analysis` now **requires batch-QC output** by default (`--allow-no-qc` opts out;
  `--allow-partial-qc` accepts a partial `flagged/`); the console always prints the input source and a
  warning when the three QC columns will be empty.
- Missing bait labels are no longer silent: a diagnostic warning (with the most enriched candidates)
  goes to the console and the Overview sheet; verification fails the run.
- `--smoke` writes to `bioid_analysis_smoke/` unless `--out` is given, so template iteration cannot
  overwrite a real report.
- Console `top=` now uses the same fold-change ordering as the tables and report.

### Fixed

- `bioid-batch-qc`: the Bait Transfer detail header was hardcoded at row 15, overwriting the last
  matrix row when a batch had ≥13 conditions (verified on BioID-8th: the FAM118B_F122R row was lost);
  it is now placed after the matrix.
- `bioid-batch-qc`: flag-reachability now uses the maximum per-gene family-hit count instead of the
  total across genes; a workbook whose control channel has no positive values is skipped with a
  warning instead of crashing the batch; `Abundance` column matching no longer accepts
  `Abundances (Normalized)`; rerunning into the input folder no longer picks up the tool's own outputs.
- `bioid-analysis`: two Python 3.12-only f-strings made the script unparseable on Python 3.9–3.11.
- `bioid-analysis`: missing `--bait-map`/`--bait-alias` files raise a clean error; `--list-baits`
  no longer creates an output directory; the bait colour in the report is defined.
- `bioid-trend`: missing `bait_self` colour, Python 3.9 compatibility, CJK labels rendering blank in
  PDF (the PDF backend does not fall back per glyph).

### Verification

- `bioid-analysis/tests/make_synthetic_and_check.py` — planted-truth checks (recovery, FDR, load
  correction, censored ratios, null data): all pass.
- `bioid/tests/test_bioid_run.py` — channel-map inference (exact/prefix/unique-channel/reversed),
  alias-draft gate, dry-run: 8 tests.
- `bioid-batch-qc/tests/test_batch_qc.py` — Bait Transfer layout at 13/2 conditions, alias-driven bait
  recovery: 4 tests.
- Real-data run: BioID-10th (8 construct-named files) end to end, 8/8 baits detected, report verified.
