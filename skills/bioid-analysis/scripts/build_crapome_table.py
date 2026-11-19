#!/usr/bin/env python3
"""Derive a compact CRAPome frequency table from the official flat files."""
import sys, csv
from collections import defaultdict
from openpyxl import load_workbook

matrix_path, annot_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]

# 1) control experiments and their cell line from the annotation file
meta = {}
wb = load_workbook(annot_path, read_only=True, data_only=True)
ws = wb["annotation_29MAR2013"]
rows = ws.iter_rows(values_only=True)
hdr = [str(v or "").strip() for v in next(rows)]
idx = {name: i for i, name in enumerate(hdr)}
for r in rows:
    if not r or not r[0]:
        continue
    cc = str(r[0]).strip()
    meta[cc] = {
        "cell": str(r[idx.get("Cell/tissue type", 4)] or "").strip(),
        "protocol": str(r[idx.get("Protocol Name", 2)] or "").strip(),
    }
wb.close()

# 2) spectral count matrix
wb = load_workbook(matrix_path, read_only=True, data_only=True)
ws = wb["Sheet1"]
rows = ws.iter_rows(values_only=True)
hdr = list(next(rows))
cc_cols = [(i, str(v).strip()) for i, v in enumerate(hdr) if str(v or "").strip().startswith("CC")]
print(f"control columns in matrix: {len(cc_cols)}; annotated experiments: {len(meta)}")
hek_cols = [cc for _, cc in cc_cols if meta.get(cc, {}).get("cell", "").lower().startswith("hek")]
print(f"HEK293 controls: {len(hek_cols)}")

def to_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0

out = []
for r in rows:
    if not r or not r[0]:
        continue
    gene = str(r[0]).strip()
    uniprot = str(r[2] or "").strip()
    n_all = n_all2 = n_all4 = 0
    n_hek = n_hek2 = n_hek4 = 0
    for i, cc in cc_cols:
        sc = to_int(r[i]) if i < len(r) else 0
        if sc > 0:
            n_all += 1
        if sc >= 2:
            n_all2 += 1
        if sc >= 4:
            n_all4 += 1
        if cc in hek_cols:
            if sc > 0: n_hek += 1
            if sc >= 2: n_hek2 += 1
            if sc >= 4: n_hek4 += 1
    out.append((gene, uniprot, n_all, n_all2, n_all4, n_hek, n_hek2, n_hek4))
wb.close()

with open(out_path, "w", newline="") as fh:
    fh.write("# CRAPome frequency table derived from the official flat file\n")
    fh.write(f"# source matrix: crap_db_v1_flat_file_human.xlsx (CRAPome v1, Apr 2016 release)\n")
    fh.write(f"# source annotation: CRAPome_annotation_30MAR2013_0.xlsx\n")
    fh.write(f"# controls: {len(cc_cols)} human, {len(hek_cols)} HEK293; counts are experiments with SC >= threshold\n")
    fh.write("# n_ctrl_ge1/n_ctrl_ge2/n_ctrl_ge4 = human controls with SC>=1/2/4; the *_hek columns are the HEK293 subset\n")
    w = csv.writer(fh, delimiter="\t")
    w.writerow(["gene", "uniprot", "n_ctrl_ge1", "n_ctrl_ge2", "n_ctrl_ge4",
                "n_ctrl_hek_ge1", "n_ctrl_hek_ge2", "n_ctrl_hek_ge4",
                "n_ctrl_total", "n_ctrl_hek_total"])
    for row in out:
        w.writerow(list(row) + [len(cc_cols), len(hek_cols)])
print(f"genes written: {len(out)}")
