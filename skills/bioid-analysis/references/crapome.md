# CRAPome reference table

## Provenance

`assets/crapome_frequency.tsv` is derived from the official CRAPome control data:

| Item | Value |
|---|---|
| Source | `crap_db_v1_flat_file_human.xlsx` (CRAPome v1 flat file, April 2016 release) |
| Annotation | `CRAPome_annotation_30MAR2013_0.xlsx` |
| Host | retrieved from `reprint-apms.org`, the successor site that `crapome.org` now redirects to |
| Controls | 343 human experiments, of which 283 are HEK293 |
| Proteins | 8,861 gene symbols |
| Reference | Mellacheruvu et al., *Nat Methods* 2013 |

Nothing is fetched at analysis time. The table ships with the skill so runs are offline and
reproducible. `crapome.org` currently serves an **expired TLS certificate** and redirects to
REPRINT; the downloader therefore uses the REPRINT URL, which has a valid certificate.

## Columns

| Column | Meaning |
|---|---|
| `gene` | gene symbol as used by CRAPome |
| `uniprot` | UniProt accession (join key; the analysis falls back to the gene symbol) |
| `n_ctrl_ge1/ge2/ge4` | human controls detecting the protein at ≥1 / ≥2 / ≥4 spectral counts |
| `n_ctrl_hek_ge1/ge2/ge4` | the same restricted to HEK293 controls |
| `n_ctrl_total`, `n_ctrl_hek_total` | denominators (343 and 283) |

## How the analysis uses it

- A prey is labelled `contaminant` when its detection frequency in the chosen control subset
  reaches `--crapome-frequency` (default 20%) at `--crapome-min-sc` spectral counts (default 2).
  This mirrors the CRAPome default filter.
- `--crapome-cell HEK293` (default) restricts to HEK293 controls, the closest match to these
  samples; `--crapome-cell all` uses every human control.
- Preys absent from CRAPome get frequency 0, which is not the same as "clean" — it means the
  protein was never seen in these control runs.
- Contaminants are **annotated, never deleted**. They are excluded from the candidate list and
  drawn in orange in the report.

## Caveats

- CRAPome controls are **affinity-purification** runs. These samples are BioID, whose background
  differs (endogenously biotinylated carboxylases such as PC, PCCA, PCCB, MCCC1/2 appear here as
  top hits and are also frequent in CRAPome — the two references agree on those, but not perfectly
  in general).
- The table is CRAPome v1 (April 2016). It is not a live resource; nothing newer is published.
- Frequency is not a probability of contamination for a specific experiment; it is a prior.

## Rebuilding

```bash
python3 scripts/fetch_crapome.py --out assets/crapome_frequency.tsv
```

Downloads both official files, verifies them, and rewrites the compact table.
`scripts/build_crapome_table.py` can be run directly to convert matrix files already on disk
(including the yeast matrix, if ever needed).
