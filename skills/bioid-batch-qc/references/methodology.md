# BioID batch methodology

## Inputs and controls

Use the first two raw `Abundance:` columns in `Proteins` as bait and paired EV. A shared EV is one negative-control sample; it is not a set of independent replicates. Keep any acquisition-order mapping separate from biological condition labels.

## Metrics

For a protein with positive bait and EV abundance:

```text
raw_log2FC = log2(bait_abundance / EV_abundance)
Norm_LOG2FC = raw_log2FC - median(raw_log2FC among measured proteins in that workbook)
```

The median subtraction removes channel-scale shifts so absolute log2FC values from different bait files are comparable. It does not create replication or statistical significance.

For EV-missing proteins, use the positive EV 5th percentile only as a censoring bound and label the result `FC_lower_bound; Control_missing`. Keep original abundance values visible in the source columns.

Within each condition, rank positive experimental abundance and normalized enrichment separately from 0 to 1, with the largest value equal to 1. Combine them as:

```text
Combined_score = sqrt(Abundance_pct * Enrichment_pct)
```

This geometric mean rewards proteins that are both abundant and specifically enriched. A high score is not a probability, p value, FDR, or proof of interaction.

## Background frequency

Collapse duplicate rows by GN for frequency calculations while retaining all original rows in the flagged workbook. A family is a deliberately defined same-bait or isoform group; family ranking takes the maximum member score, while condition counts retain every condition. If no family metadata is available, use each bait as its own family and state that assumption.

Apply these fixed review thresholds:

| Label | Rule |
|---|---|
| High background suspect (orange) | Family combined top 2% in ≥4 families |
| Review background suspect (yellow) | Not orange, and family combined top 5% in ≥4 families OR family enrichment top 5% in ≥5 families |

Never delete these proteins automatically. Repeated localization or common cellular-neighborhood biology can produce genuine shared BioID neighbors.

Minimum batch size: every rule needs at least 4 families holding a qualifying gene, and the enrichment route needs 5. The counts are also limited by family size, since a family resolves its top 2% only from >=50 detected genes and its top 5% from >=20. When a rule cannot fire, the run prints a WARNING and records the unreachable rules in the `Threshold reachability` row of the QC `Methods` sheet; an empty flag list at that size is not evidence of a clean batch. The thresholds themselves stay fixed rather than being rescaled per batch.

## Cross-sample screening

Build raw experimental `log2(abundance)` profiles per prey, subtracting each prey's across-condition median. Use this prey-centered matrix for Spearman correlations and Top-100 specificity Jaccard; do not use shared-EV log2FC correlations to call swaps.

The Spearman cutoffs are applied only when both are positive. If the non-family distribution at this batch size is not positive, pairs are reported as `not estimable` instead of being labeled elevated similarity, which a non-positive cutoff would otherwise do for uncorrelated or anti-correlated profiles.

For each source bait gene, calculate `target_abundance / source_abundance`. Adjacent low-ratio transfer in the assumed F-order is `Possible carryover (order-dependent provisional)`. Larger nonadjacent transfer is `Cross-sample signal / unresolved` or `Possible mixing/mislabeling or true proximity`, depending on ratio. Similarity without bait transfer is `Biological overlap / unresolved`, not proof of contamination.

If a bait is absent from its own result, mark `Bait recovery QC failure` and review bait-specific peptides, affinity tags, reporter/channel metadata, blanks, chromatographic peaks, and the actual injection order.

## Validation checklist

1. Confirm source workbooks are not overwritten.
2. Confirm every output has `+5%`, `Proteins`, `Raw`, and `QC_Methods`.
3. Confirm `Raw` values and original source columns are unchanged.
4. Confirm appended columns and the four-column visibility rule.
5. Confirm EV-missing rows are labeled lower bounds, not measured FC.
6. Confirm isoforms are separate conditions and family counts are not double-weighted.
7. Scan for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, and `#N/A`.
8. Render a focused preview of a flagged workbook and every QC report sheet before handoff.
9. Confirm the inferred multi-condition family mapping printed at the start of the run matches the intended design.
10. Confirm the `Threshold reachability` row: if any rule cannot fire at this batch size, report that with the results.
