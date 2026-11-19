# Protein annotation (hover cards)

The report's hover card reads like a GeneCards entry: gene symbol, UniProt accession, reviewed
protein name, gene synonyms, and a short function summary. That text comes from a cache that ships
with the skill, not from a live lookup.

## Where the data comes from

| Item | Value |
|---|---|
| Source | UniProt REST `uniprotkb/search`, fields `accession,gene_names,protein_name,cc_function` |
| Cache | `assets/uniprot_annotations.tsv` |
| Coverage | the reviewed human proteome (20,432) + the reviewed mouse proteome (17,283) + every accession seen in the BioID-1st/2nd/3nd and xy workbooks — 37,715 entries, none missing |
| Species | human and mouse, because BioID series are run in both |
| Columns | `accession`, `name`, `gene`, `aliases`, `function` |

`function` is the UniProt `FUNCTION` comment with the `FUNCTION:` prefix removed, ECO evidence tags
stripped, and trimmed to the first sentence or two (≤ ~420 characters). `aliases` is the gene
synonym list minus the primary symbol.

Reports embed only the accessions they use — BioID-3nd uses 3,736 of the 5,834 — so a report opens,
travels and can be sent to a collaborator with no network access.

## Refreshing or extending

```bash
# extend for a new batch (fetches only what is missing)
python3 scripts/fetch_uniprot_annotations.py --folder /path/to/BioID-batch [--folder ...] [--refresh]

# refetch a whole base layer; the stream is gzipped and retried, because a long
# stream can end in an IncompleteRead
python3 scripts/fetch_uniprot_annotations.py --proteome "organism_id:10090 AND reviewed:true"
```

Reads the `Accession` column of every workbook, fetches only what is missing from the cache, saves
after each batch (so an interrupted run keeps its progress), and merges into the existing table. Run
it once when a new batch is added; the analysis itself never touches the network.

Notes learned while building this:

- **Batch size matters.** The search endpoint rejects long queries with `HTTP 400`; 100 accessions
  per request works, 120 does not. The default is 100 and `--batch` can lower it.
- **Python's SSL trust store is not configured on this machine** (`CERTIFICATE_VERIFY_FAILED` even
  though `curl` succeeds). The script builds its context from `certifi` (installed) or
  `/etc/ssl/cert.pem`, so verification stays on — nothing is fetched insecurely.
- **Withdrawn entries** carry placeholder names such as `demerged`; the parser drops those and lets
  the source workbook's own description be shown instead.

## What the analysis does with it

Nothing analytical. Annotation is display-only: the hover card joins by accession and falls back to
the workbook's `Description` column when an accession has no entry. It cannot change any score, gate
or ranking.
