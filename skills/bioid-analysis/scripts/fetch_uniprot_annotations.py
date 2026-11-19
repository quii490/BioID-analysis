#!/usr/bin/env python3
"""Fetch GeneCards-style annotation for a set of UniProt accessions.

Produces the compact table the report uses for its hover cards: protein name,
gene symbol, aliases (gene synonyms) and a short function summary. The table is
cached in assets/ so report runs stay offline and reproducible.

    python3 scripts/fetch_uniprot_annotations.py --folder /path/to/batch \
        [--out assets/uniprot_annotations.tsv] [--refresh]

Accessions are read from every workbook in the folder (and any --extra folders),
queried in batches against the UniProt REST API, and merged into the existing
cached table, so repeated runs only fetch what is missing.
"""

from __future__ import annotations

import argparse
import csv
import gzip
import re
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
DEFAULT_OUT = HERE.parent / "assets" / "uniprot_annotations.tsv"
API = "https://rest.uniprot.org/uniprotkb/search"
FIELDS = "accession,gene_names,protein_name,cc_function"
BATCH = 100
MAX_RETRIES = 4
# Real UniProt accessions mix letters and digits freely (P, Q and O included);
# only the O-G-O pattern is reserved for obsolete entries.
ACCESSION_RE = re.compile(r"^[A-Z][A-Z0-9]{5,9}$")


def accessions_from_workbook(path):
    found = set()
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = "Proteins" if "Proteins" in wb.sheetnames else wb.sheetnames[0]
        rows = wb[sheet].iter_rows(values_only=True)
        header = list(next(rows))
        try:
            index = header.index("Accession")
        except ValueError:
            return found
        for row in rows:
            if index >= len(row) or not row[index]:
                continue
            raw = str(row[index])
            for part in re.split(r"[;,]", raw):
                candidate = part.strip().split("-")[0].strip()
                if ACCESSION_RE.match(candidate):
                    found.add(candidate)
    finally:
        wb.close()
    return found


def collect(folders):
    accessions = set()
    for folder in folders:
        for path in sorted(Path(folder).expanduser().glob("*.xlsx")):
            if path.name.startswith(("~$", ".~")):
                continue
            accessions |= accessions_from_workbook(path)
    return accessions


# UniProt marks entries that were withdrawn or merged; their name fields carry
# placeholder text that must never reach a hover card.
PLACEHOLDER_NAMES = {"demerged", "deleted", "obsolete", "merged"}


def shorten_function(text, limit=420):
    """First sentences of the FUNCTION comment, which is what a hover card wants."""
    if not text:
        return ""
    text = re.sub(r"^\s*FUNCTION:\s*", "", text.strip())
    clean = re.sub(r"\s*\{ECO:[^}]*\}", "", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    if len(clean) <= limit:
        return clean
    cut = clean[:limit]
    for stop in (". ", "; "):
        position = cut.rfind(stop)
        if position > limit * 0.5:
            return cut[: position + 1].strip()
    return cut.rstrip() + "…"


def clean_names(text):
    """'TUBB4B (TUBB4) {ECO:...}' -> ['TUBB4B', 'TUBB4']."""
    if not text:
        return []
    clean = re.sub(r"\{ECO:[^}]*\}", "", text)
    clean = clean.replace("(", " ").replace(")", " ")
    parts = re.split(r"[;,\s]+", clean)
    seen, out = set(), []
    for part in parts:
        name = part.strip()
        if name and name.upper() not in seen:
            seen.add(name.upper())
            out.append(name)
    return out


def ssl_context():
    """A verifying SSL context.

    This machine's python.org build has no default CA file configured, so the
    standard library refuses every HTTPS request with CERTIFICATE_VERIFY_FAILED
    while curl succeeds. certifi (already installed) supplies the bundle; the
    macOS system bundle is the fallback. Verification stays on either way.
    """
    for candidate in _ca_candidates():
        if candidate and Path(candidate).exists():
            return ssl.create_default_context(cafile=candidate)
    return ssl.create_default_context()


def _ca_candidates():
    candidates = ["/etc/ssl/cert.pem"]
    try:
        import certifi
        candidates.insert(0, certifi.where())
    except Exception:
        pass
    return candidates


def fetch_batch(accessions, retries=MAX_RETRIES):
    query = " OR ".join(f"accession:{a}" for a in accessions)
    url = (f"{API}?query={urllib.parse.quote(query)}&fields={FIELDS}"
           f"&format=tsv&size={len(accessions)}")
    last = None
    for attempt in range(1, retries + 1):
        try:
            request = urllib.request.Request(url, headers={"User-Agent": "bioid-analysis/1.0"})
            with urllib.request.urlopen(request, timeout=120, context=ssl_context()) as response:
                return response.read().decode("utf-8", "replace")
        except Exception as exc:  # network here is flaky; retry with backoff
            last = exc
            time.sleep(2 * attempt)
    raise RuntimeError(f"UniProt request failed after {retries} attempts: {last}")


def parse_tsv(text):
    """Turn a UniProt TSV response into {accession: (name, symbol, aliases, function)}."""
    out = {}
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        return out
    header = lines[0].split("\t")
    try:
        i_acc = header.index("Entry")
        i_gene = header.index("Gene Names")
        i_name = header.index("Protein names")
        i_func = header.index("Function [CC]")
    except ValueError as exc:
        raise RuntimeError(f"unexpected UniProt columns {header}: {exc}")
    for line in lines[1:]:
        cells = line.split("\t")
        if len(cells) <= max(i_acc, i_gene, i_name, i_func):
            continue
        accession = cells[i_acc].strip()
        if not accession:
            continue
        aliases = clean_names(cells[i_gene])
        if cells[i_name].strip().startswith("["):
            # A subunit entry: '[Histone H2A]: ...' — the name lives after the colon
            cells = cells[:]
            cells[i_name] = cells[i_name].split(":", 1)[-1]
        name = re.sub(r"\s*\(.*?\)\s*$", "", cells[i_name]).strip()
        name = re.sub(r"\s*\{ECO:.*$", "", name).strip()
        if name.strip().lower() in PLACEHOLDER_NAMES:
            name = ""
        out[accession] = (
            name,
            aliases[0] if aliases else "",
            "; ".join(aliases[1:]) if len(aliases) > 1 else "",
            shorten_function(cells[i_func]),
        )
    return out


def load_existing(path):
    table = {}
    if not path.exists():
        return table
    with open(path, newline="") as fh:
        lines = [line for line in fh if not line.startswith("#")]
    for row in csv.DictReader(lines, delimiter="\t"):
        table[row["accession"]] = (row["name"], row["gene"], row["aliases"], row["function"])
    return table


def write_table(path, table):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="") as fh:
        fh.write("# UniProt annotation cache for report hover cards\n")
        fh.write("# source: rest.uniprot.org /uniprotkb/search (fields: "
                 "accession,gene_names,protein_name,cc_function)\n")
        fh.write("# regenerate or extend with scripts/fetch_uniprot_annotations.py\n")
        writer = csv.writer(fh, delimiter="\t")
        writer.writerow(["accession", "name", "gene", "aliases", "function"])
        for accession in sorted(table):
            writer.writerow([accession, *table[accession]])


STREAM = "https://rest.uniprot.org/uniprotkb/stream"


def fetch_stream(query, fields=FIELDS, retries=3, compressed=True):
    """One streamed request for a whole query (used for base layers).

    The gzipped form is markedly faster for a whole proteome, so it is the
    default; the response is decompressed here.
    """
    url = f"{STREAM}?query={urllib.parse.quote(query)}&fields={fields}&format=tsv"
    if compressed:
        url += "&compressed=true"
    last = None
    for attempt in range(1, retries + 1):
        try:
            request = urllib.request.Request(url, headers={"User-Agent": "bioid-analysis/1.0"})
            with urllib.request.urlopen(request, timeout=1800, context=ssl_context()) as response:
                payload = response.read()
            if compressed:
                payload = gzip.decompress(payload)
            return payload.decode("utf-8", "replace")
        except Exception as exc:
            last = exc
            print(f"  stream attempt {attempt} failed: {exc}", file=sys.stderr)
            time.sleep(5 * attempt)
    raise RuntimeError(f"UniProt stream failed: {last}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", action="append", default=[],
                        help="folder of workbooks; repeat to add more")
    parser.add_argument("--proteome", action="append", default=[],
                        help="fetch a whole base layer, e.g. 'organism_id:9606 AND reviewed:true'")
    parser.add_argument("--out", default=str(DEFAULT_OUT))
    parser.add_argument("--refresh", action="store_true",
                        help="re-fetch accessions already in the cache")
    parser.add_argument("--batch", type=int, default=BATCH)
    args = parser.parse_args()

    if not args.folder and not args.proteome:
        parser.error("give at least one --folder or --proteome")

    out = Path(args.out).expanduser().resolve()
    table = {} if args.refresh else load_existing(out)
    print(f"cached annotations: {len(table)}")

    for query in args.proteome:
        print(f"streaming base layer: {query}")
        parsed = parse_tsv(fetch_stream(query))
        added = sum(1 for a in parsed if a not in table)
        table.update(parsed)
        print(f"  {len(parsed)} records, {added} new (running total {len(table)})")
        write_table(out, table)

    accessions = collect(args.folder) if args.folder else set()
    print(f"accessions found in workbooks: {len(accessions)}")
    todo = sorted(a for a in accessions if a not in table)
    print(f"to fetch: {len(todo)}")
    if not todo:
        print("cache is complete for these inputs")
        return 0

    fetched = 0
    for start in range(0, len(todo), args.batch):
        chunk = todo[start:start + args.batch]
        try:
            parsed = parse_tsv(fetch_batch(chunk))
        except Exception as exc:
            print(f"  batch {start // args.batch + 1}: FAILED — {exc}", file=sys.stderr)
            continue
        table.update(parsed)
        fetched += len(parsed)
        print(f"  batch {start // args.batch + 1}: "
              f"{len(parsed)}/{len(chunk)} annotated "
              f"(running total {len(table)})")
        write_table(out, table)          # save as we go, so a failure keeps progress
        time.sleep(0.4)

    write_table(out, table)
    missing = [a for a in accessions if a not in table]
    print(f"annotations written: {len(table)} -> {out}")
    print(f"accessions with no UniProt record: {len(missing)}")
    if missing[:10]:
        print("  e.g.", ", ".join(missing[:10]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
