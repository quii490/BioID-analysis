# bioid-analysis methodology

## Input contract

One `.xlsx` workbook per bait condition, each with a `Proteins` (or `+5%`) sheet holding:

- `Accession` (UniProt), `GN`, `Description`, `Coverage [%]`
- `# Peptides`, `# PSMs`, `# Unique Peptides`
- two raw `Abundance:` columns: the **first** is the experimental channel, the **second** is the paired control (EV)

Workbooks are read-only. Everything is written to a new `bioid_analysis/` folder.

## Load correction

A bait and an EV channel that received different amounts of material shift every fold change in
the same direction. Three estimators are computed for every condition, because they disagree when
a few proteins dominate the signal:

| Estimator | Definition | Behaviour |
|---|---|---|
| `total` | `log2(sum control / sum experimental)` | Matches summed signal, which is what a genuine loading difference looks like. Fragile: in these batches the 10 most abundant proteins carry 47–85% of the signal, so a few proteins decide the correction. |
| `median` | median of `log2(experimental / control)` over all paired proteins | Standard label-free normalization; assumes the typical protein is not enriched. Default. |
| `robust` | the same median restricted to the better-quantified half | Same idea with less influence from the noisy detection limit. |

`--normalize` selects which one is applied to the folded changes and the gates.

Two properties decide how much this choice matters:

- **Within one condition the correction cannot reorder anything.** Subtracting the same constant from every
  protein's log2 ratio leaves the ranking untouched, so which proteins stand out is decided by the ratio spread,
  not by this choice.
- **It decides where a fixed cutoff falls, and it feeds the q-values.** That is the point of doing it: without
  it a 1.5x gate means something different in every condition, and in these batches the offsets reach 4.8 log2
  (28x).

The assumption to check is that the typical quantified prey is not an interactor. If a bait genuinely
co-purifies a large abundant complex, the median moves with real signal and the correction over-corrects. That is
why the candidate counts are printed for all three estimators rather than only the applied one.

`raw FC` is not a new quantity: it equals the `LOG2FC` column already present in the source workbook, verified
identical to machine precision. `normalized FC = raw FC x 2^(-correction)`, and it is the number the gates use. All three values,
the concentration of signal in the top 10 proteins, and the candidate count under each estimator
are reported, so the sensitivity of the shortlist to this choice is visible.

In a TMT experiment all channels are combined before acquisition, so a per-channel "load"
difference is really a difference in how much protein each pulldown delivered. When it is known
to be a yield or loading artefact, `total` is the right correction; when the bait itself may have
brought down more protein, a robust estimator is safer — which is why the default is `median`.

## Gates

| Gate | Default | Meaning |
|---|---|---|
| Fold change | ≥ 1.5× normalized | Literature-style cutoff for a potential interactor |
| Unique peptides | ≥ 2 | Minimum peptide evidence in the experiment |
| CRAPome | < 20% of controls at ≥ 2 spectral counts | Common contaminant |
| Bait self-hit | excluded | The bait is always detected in its own run; drawn as its own colour and never counted as a candidate |

`# Unique Peptides` in a Proteome Discoverer protein table is counted for the **experiment**, not
for one channel. It is a proxy for per-channel peptide evidence; an exact per-channel count needs
a peptide-level export. This is stated in the report and in the Methods sheet.

### Prey the control never detected (censored ratios)

A prey that the control channel never detected is not a missing value with no information — it is a
censored one. Its true ratio is *at least* `bait / floor`, where the floor is what a detectable
control value would have been at worst. Two consequences shape the pipeline:

- **Dropping the row is a bias, not a neutral default.** These prey are precisely the "seen in the
  bait, absent from the EV" class, which is the strongest single-run evidence a BioID design offers.
  Dropping them (the naive `fc_raw is None → skip`) removes the top of the ranking.
- **Feeding them in as ordinary ratios is also wrong.** A floor ratio is systematically large, so it
  widens the empirical null (inflating σ and suppressing real hits) and, once in the BH ranking, takes
  a test slot while carrying a p-value the data does not support.

The pipeline therefore keeps them and labels them:

| Step | Treatment |
|---|---|
| Floor | 5th percentile of the positive control values in the same workbook (identical to `bioid-batch-qc`) |
| Ratio | `log2(bait / floor) − correction`, written as `FC_lower_bound; Control_missing` |
| Fold gate | applied to the lower bound, so passing it means the real ratio passes too (conservative) |
| σ, mixture fit, p, q | excluded — a censored ratio is not a measurement and not a test |
| Report | hollow teal triangles at the floor they were compared against, `EV: not detected` in the table, counted on their own card |

The displayed FC always carries a `≥` for these rows, in the report and in the workbook, so a floor
can never be read as a measurement. `tests/make_synthetic_and_check.py` case 4 asserts all five
properties.

## Scoring

**SAINT-style score.** A two-component Gaussian mixture is fitted per condition by EM on the
normalized log2 ratios. The no-change component is centred on zero with a width taken from the
**negative half** of the ratio distribution (enrichment only pushes ratios up, so the negative half
is unaffected by the very hits being sought). The mixing proportion carries a `Beta(1, 20)` prior,
as the SAINT family does; without it a dataset with no interactors still yields a non-empty
"enriched" component fitted to the positive tail. The reported score is the posterior probability
of belonging to the enriched component. If the fitted shift is not separated from the background
spread the model reports "not fitted" instead of inventing a cluster.

**q-value (calibrated).** A one-sided p-value from the same empirical null
(`z = normalized log2FC / sigma0`) followed by Benjamini–Hochberg over all quantified prey. This is
the multiple-testing-aware companion: it needs only that the null be roughly symmetric, not that
the enriched preys form an identifiable second cluster.

**BFDR** is the Newton et al. (2004) estimator over the posterior probabilities, reported alongside.

### What these numbers are not

- **Not the official SAINTq.** SAINTq scores peptide or fragment intensities; its paper describes
  protein-level aggregation as "SAINTexpress-equivalent" and reports that it underestimates the
  FDR. This skill implements the same model shape at protein level and says so in the report.
- **Not calibrated to a real FDR** in a design with one run per bait and one shared control.
  Replicates are what make SAINT-family scores calibrated.

## The specificity problem (measured, not assumed)

On the batches in use here, the corrected log2 ratio has a spread of **σ ≈ 1.3–2.1 log2**
(2.5–4× ). Consequences that the report states explicitly:

- **20–43% of proteins clear a 1.5× cutoff.** The cutoff is a filter for manual review, never
  evidence of an interaction.
- The fold change at which the ratio survives BH at q ≤ 0.05 is far larger (tens of fold), which is
  why the calibrated layer usually finds few or no proteins on a single run.
- The two channels correlate at r ≈ 0.7–0.83 across proteins, so per-protein ratios carry a large
  measurement component. Evidence beyond the ratio — peptide support, CRAPome exclusion, and
  especially **repetition across baits** — is what makes a hit believable. For the across-bait
  view use the sibling skill `bioid-batch-qc`, which counts how often a prey is a top hit in
  different baits and screens cross-sample carryover.

## Annotation (display only)

Hover cards are built from `assets/uniprot_annotations.tsv` (UniProt name, gene synonyms, trimmed
FUNCTION comment), shared per accession across conditions and embedded only for the accessions the
batch actually uses. It is display-only: it cannot change a score, a gate or a ranking, and a run
without the cache still works, falling back to the workbook's own `Description`. See
`references/annotations.md`.

## Outputs

- `BioID_analysis_summary.xlsx`: `Overview` (per-condition QC, load estimates, candidate counts per
  estimator, noise scale), `Candidates` (every protein passing all gates, ranked by score), one
  sheet per condition (all metrics, both reference normalizations), `Methods`.
- `BioID_analysis_report.html`: self-contained interactive report with plotly.js embedded, so it
  opens offline and can be sent to a collaborator as one file.
- `analysis_payload.json`: the same numbers, machine-readable.

## Validation checklist

1. Source workbooks untouched; output confined to `bioid_analysis/`.
2. `Overview` load ratio and corrections plausible per condition; warnings read.
3. Ratio noise σ and "share above gate" read before interpreting any candidate list.
4. Candidates checked against CRAPome status, unique peptides and PSMs.
5. Where the mixture did not fit, say so rather than quoting a score.
6. Report written to the user without presenting the 1.5× list as statistically significant.
7. `EV not detected` rows read as lower bounds; their count reported, never mixed into the measured
   candidate evidence.
8. The drawn gate line checked against the coloured points (the report's own regression test does
   this): the line belongs at `minFold × 2^correction`, not at `minFold`.

## References

- Mellacheruvu et al. *Nat Methods* 2013 — CRAPome contaminant repository.
- Choi et al. *Nat Methods* 2011 — SAINT spectral-count mixture model.
- Teo et al. *Proteomics* 2016 — SAINTq, peptide/fragment-level scoring.
- Newton et al. *Biostatistics* 2004 — FDR from posterior probabilities.
- Benjamini & Hochberg 1995 — false discovery rate control.
