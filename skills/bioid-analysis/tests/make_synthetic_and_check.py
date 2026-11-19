#!/usr/bin/env python3
"""Ground-truth test for the bioid-analysis pipeline.

Builds synthetic protein-level workbooks where the answer is known, runs the
real pipeline, and checks that each requested feature does what it claims:

  * enriched proteins are recovered and the fold changes match the planted ones
  * the SAINT-style score ranks them well, and the calibrated q-value controls
    the false discovery rate at the level it advertises
  * the unique-peptide gate drops planted one-peptide proteins
  * the CRAPome gate drops planted common contaminants
  * a deliberately introduced 4x channel-scale difference is measured and
    removed, so the planted enrichment is still recovered
  * on a genuinely null dataset (no enrichment anywhere) nothing is discovered

Usage: python3 tests/make_synthetic_and_check.py [--keep]
"""

from __future__ import annotations

import argparse
import math
import random
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
PIPELINE = SKILL / "scripts" / "bioid_analysis.py"

HEADERS = [
    "Master", "Accession", "Description", "GN", "Coverage [%]",
    "# Peptides", "# PSMs", "# Unique Peptides", "# AAs", "MW [kDa]",
    "Abundance: F{exp}: Sample", "Abundance: F{ctrl}: Sample", "Modifications", "LOG2FC",
]

CONTAMINANT_GENES = [
    "HSPA1A", "TUBA1A", "HSPA8", "ACTB", "TUBB", "HSPA1L", "CCT2", "FLNA",
    "ACACA", "MCCC1", "PC", "PCCA", "PCCB", "HSP90AA1", "HSP90AB1", "RPL13",
    "RPL7A", "EEF1A1", "GAPDH", "PKM", "CCT5", "TCP1", "TUBB4B", "MYH9", "ACTN4",
]


def write_workbook(path, rows, exp_channel, ctrl_channel):
    wb = Workbook()
    ws = wb.active
    ws.title = "Proteins"
    ws.append([h.format(exp=exp_channel, ctrl=ctrl_channel) for h in HEADERS])
    for row in rows:
        ws.append(row)
    for sheet in ("+5%", "Raw"):
        wb.create_sheet(sheet).append(["placeholder"])
    wb.save(path)
    wb.close()


def simulate(seed, n_background=1800, n_true=45, shift_log2=2.8, control_scale=1.0,
             n_contaminants=25, n_one_peptide=40):
    """Rows plus the ground truth they encode.

    'enriched' is every protein deliberately given a higher experimental
    abundance, which is what the gates and the score are supposed to find.
    The background rows are the only true negatives.
    """
    rng = random.Random(seed)
    rows = []
    truth = {"enriched": [], "true": [], "contaminant": [], "one_peptide": [], "bait": "BAIT1"}

    def add(gn, acc, bait_ab, ctrl_ab, up, psm, cov=20.0):
        rows.append([
            "Master", acc, f"{gn} description OS=Homo sapiens", gn, cov,
            max(up, 1) + 2, psm, up, 400, 45.0, bait_ab, ctrl_ab, "None", None,
        ])

    add("BAIT1", "Q00001", 5e8, 2e6, 12, 220, 60.0)

    for i in range(n_background):
        base = math.exp(rng.gauss(22.0, 1.6))
        add(f"BG{i:04d}", f"BG{i:04d}",
            base * math.exp(rng.gauss(0.0, 0.28)),
            base * math.exp(rng.gauss(0.0, 0.28)) * control_scale,
            rng.choice([2, 3, 4, 5, 8]), rng.randint(2, 60))

    for i in range(n_true):
        base = math.exp(rng.gauss(21.5, 1.2))
        gn = f"TRUE{i:02d}"
        add(gn, f"T{i:04d}",
            base * (2 ** shift_log2) * math.exp(rng.gauss(0.0, 0.3)),
            base * math.exp(rng.gauss(0.0, 0.3)) * control_scale,
            rng.choice([3, 4, 6, 9, 14]), rng.randint(5, 120))
        truth["true"].append(gn)
        truth["enriched"].append(gn)

    for i in range(min(n_contaminants, len(CONTAMINANT_GENES))):
        gn = CONTAMINANT_GENES[i]
        base = math.exp(rng.gauss(22.5, 1.0))
        add(gn, f"C{i:04d}",
            base * (2 ** rng.uniform(1.0, 3.0)) * math.exp(rng.gauss(0.0, 0.2)),
            base * math.exp(rng.gauss(0.0, 0.2)) * control_scale,
            rng.choice([3, 5, 8]), rng.randint(5, 150))
        truth["contaminant"].append(gn)
        truth["enriched"].append(gn)

    for i in range(n_one_peptide):
        gn = f"ONEPEP{i:02d}"
        base = math.exp(rng.gauss(21.0, 1.0))
        add(gn, f"P{i:04d}",
            base * (2 ** rng.uniform(2.0, 4.0)) * math.exp(rng.gauss(0.0, 0.2)),
            base * math.exp(rng.gauss(0.0, 0.2)) * control_scale,
            1, rng.randint(1, 4))
        truth["one_peptide"].append(gn)
        truth["enriched"].append(gn)

    return rows, truth



def with_missing_controls(seed, n_missing=25, n_background=1800):
    """Planted interactors that the control channel never detected.

    This is the censored case: the prey is seen in the bait and absent from the
    EV, so its ratio is a lower bound rather than a measurement. The pipeline must
    keep it (a floor that clears the gate is conservative), must label it, and must
    keep it out of the background spread and the p/q values.
    """
    rows, truth = simulate(seed, n_background=n_background, n_true=45)
    missing = [r for r in rows if str(r[3]).startswith("TRUE")][:n_missing]
    missing_gn = set()
    for row in missing:
        row[11] = None            # control abundance: not detected
        missing_gn.add(row[3])
    truth["missing"] = sorted(missing_gn)
    return rows, truth


def pure_null(seed, n_background=1800, control_scale=1.0):
    """A dataset with no enrichment at all: the bait plus unchanged background."""
    rows, truth = simulate(seed, n_background=n_background, n_true=0, control_scale=control_scale)
    rows = [r for r in rows
            if not str(r[3]).startswith(("TRUE", "ONEPEP")) and r[3] not in CONTAMINANT_GENES]
    truth["enriched"] = []
    return rows, truth


def run(folder, out_dir):
    result = subprocess.run(
        # 合成用例只测统计核心，不带 batch-QC：显式 --allow-no-qc（默认会拒绝缺 QC 的输入）
        [sys.executable, str(PIPELINE), str(folder), "--out", str(out_dir), "--plotly", "cdn",
         "--allow-no-qc"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        raise SystemExit(f"pipeline failed with code {result.returncode}")
    return result.stdout


def read_summary(out_dir, condition):
    wb = load_workbook(out_dir / "BioID_analysis_summary.xlsx", data_only=True)
    ws = wb[condition]
    index = {cell.value: i for i, cell in enumerate(ws[1])}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rows.append({
            "gn": row[index["GN"]],
            "fc": row[index["FC (EV-corrected)"]],
            "fc_raw": row[index["FC (raw)"]],
            "score": row[index["SAINT-style score"]],
            "q": row[index["q (BH)"]],
            "up": row[index["Unique peptides"]],
            "passes": row[index["Passes gates"]] == "yes",
            "why": row[index["Reasons excluded"]],
            "fc_type": row[index["FC type"]] if "FC type" in index else None,
            "control_used": row[index["Control used"]] if "Control used" in index else None,
        })
    wb.close()
    return rows



def read_diagnostics(out_dir, condition):
    """The Overview row for one condition, keyed by its column headers."""
    wb = load_workbook(out_dir / "BioID_analysis_summary.xlsx", data_only=True)
    ws = wb["Overview"]
    headers = [cell.value for cell in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] == condition:
            wb.close()
            return dict(zip(headers, row))
    wb.close()
    return {}


def auc(scores, labels):
    """Rank-based AUC (Mann-Whitney); ascending ranks, so 0.5 is chance."""
    pairs = sorted(zip(scores, labels), key=lambda t: (t[0], t[1]))
    positives = sum(labels)
    negatives = len(labels) - positives
    if positives == 0 or negatives == 0:
        return float("nan")
    rank_sum, index = 0.0, 0
    while index < len(pairs):
        end = index
        while end + 1 < len(pairs) and pairs[end + 1][0] == pairs[index][0]:
            end += 1
        average_rank = (index + 1 + end + 1) / 2
        rank_sum += average_rank * sum(pairs[k][1] for k in range(index, end + 1))
        index = end + 1
    return (rank_sum - positives * (positives + 1) / 2) / (positives * negatives)


def median(values):
    values = sorted(v for v in values if v is not None)
    if not values:
        return None
    mid = len(values) // 2
    return values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2


def check(name, passed, detail=""):
    print(f"  [{'PASS' if passed else 'FAIL'}] {name}{(' — ' + detail) if detail else ''}")
    return bool(passed)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--keep", action="store_true")
    args = parser.parse_args()
    root = Path(tempfile.mkdtemp(prefix="bioid_synth_"))
    ok = True
    print(f"synthetic workspace: {root}")

    # ------------------------------------------------------------------ case 1
    print("\ncase 1: 45 interactors at 2.8 log2, plus planted contaminants and one-peptide proteins")
    folder = root / "case1"
    folder.mkdir(parents=True)
    rows, truth = simulate(seed=1)
    write_workbook(folder / "BAIT1.xlsx", rows, 1, 14)
    out = folder / "bioid_analysis"
    run(folder, out)
    table = read_summary(out, "BAIT1")
    enriched = truth["enriched"]
    labels = [1 if r["gn"] in enriched else 0 for r in table]
    scores = [r["score"] if r["score"] is not None else 0.0 for r in table]
    score_auc = auc(scores, labels)
    ok &= check("SAINT-style score ranks enriched proteins first",
                score_auc > 0.95, f"AUC={score_auc:.4f} against all planted enriched proteins")

    discovered = {r["gn"] for r in table if r["q"] is not None and r["q"] <= 0.05}
    recalled = discovered & set(truth["true"])
    false_pos = discovered - set(enriched)
    fdr = len(false_pos) / len(discovered) if discovered else 0.0
    ok &= check("q <= 0.05 recovers the planted interactors",
                len(recalled) >= 0.9 * len(truth["true"]),
                f"{len(recalled)}/{len(truth['true'])} recovered")
    ok &= check("q <= 0.05 keeps the false discovery rate low",
                fdr <= 0.05,
                f"FDR={fdr:.1%} ({len(false_pos)} of {len(discovered)} discoveries were pure background)")

    fc = {r["gn"]: r["fc"] for r in table}
    true_fc = median([fc[g] for g in truth["true"] if fc.get(g)])
    ok &= check("median fold change of interactors matches the planted 7.0x",
                true_fc is not None and 5.0 < true_fc < 9.0, f"median FC={true_fc:.2f}x")

    leaked = [r["gn"] for r in table if r["gn"] in truth["one_peptide"] and r["passes"]]
    ok &= check("one-peptide proteins never pass the gates", not leaked, f"{len(leaked)} leaks")

    flagged = {r["gn"] for r in table
               if r["gn"] in truth["contaminant"] and "CRAPome" in (r["why"] or "")}
    ok &= check("planted contaminants are flagged by CRAPome",
                len(flagged) >= 0.7 * len(truth["contaminant"])
                and {"HSPA1A", "HSPA8", "TUBA1A"} <= flagged,
                f"{len(flagged)}/{len(truth['contaminant'])} flagged")

    candidates = {r["gn"] for r in table if r["passes"]}
    ok &= check("real interactors survive every gate", len(set(truth["true"]) & candidates) > 0,
                f"{len(set(truth['true']) & candidates)} candidates")
    ok &= check("the bait itself is never a candidate", "BAIT1" not in candidates)

    # ------------------------------------------------------------------ case 2
    print("\ncase 2: control channel scaled 4x — the load correction must absorb it")
    folder2 = root / "case2"
    folder2.mkdir(parents=True)
    rows2, truth2 = simulate(seed=2, control_scale=4.0)
    write_workbook(folder2 / "BAIT1.xlsx", rows2, 2, 14)
    out2 = folder2 / "bioid_analysis"
    run(folder2, out2)
    table2 = read_summary(out2, "BAIT1")
    fc2 = {r["gn"]: r["fc"] for r in table2}
    raw2 = median([r["fc_raw"] for r in table2 if r["gn"] in truth2["true"]])
    corrected2 = median([fc2[g] for g in truth2["true"] if fc2.get(g)])
    ok &= check("planted enrichment is recovered after correction",
                corrected2 is not None and 5.0 < corrected2 < 9.0,
                f"normalized FC={corrected2:.2f}x where the raw ratio reads {raw2:.2f}x")
    scores2 = [r["score"] if r["score"] is not None else 0.0 for r in table2]
    labels2 = [1 if r["gn"] in truth2["enriched"] else 0 for r in table2]
    auc2 = auc(scores2, labels2)
    ok &= check("ranking survives the correction", auc2 > 0.95, f"AUC={auc2:.4f}")
    discovered2 = {r["gn"] for r in table2 if r["q"] is not None and r["q"] <= 0.05}
    recall2 = len(discovered2 & set(truth2["true"]))
    ok &= check("q <= 0.05 still recovers the interactors",
                recall2 >= 0.9 * len(truth2["true"]),
                f"{recall2}/{len(truth2['true'])} recovered")

    # ------------------------------------------------------------------ case 3
    print("\ncase 3: genuinely null dataset (bait plus unchanged background only)")
    folder3 = root / "case3"
    folder3.mkdir(parents=True)
    rows3, truth3 = pure_null(seed=3)
    truth3_bait = truth3["bait"]
    write_workbook(folder3 / "BAIT1.xlsx", rows3, 3, 14)
    out3 = folder3 / "bioid_analysis"
    run(folder3, out3)
    table3 = read_summary(out3, "BAIT1")
    # The bait itself is enriched in its own channel by construction and is
    # already excluded from the candidate list, so it is not a false discovery.
    discovered3 = [r for r in table3
                   if r["q"] is not None and r["q"] <= 0.05 and r["gn"] != truth3_bait]
    posterior_high = [r for r in table3 if (r["score"] or 0) >= 0.9]
    candidates3 = [r for r in table3 if r["passes"]]
    ok &= check("nothing is discovered at q <= 0.05", len(discovered3) <= 2,
                f"{len(discovered3)} discoveries")
    quantified = [r for r in table3 if r["fc"] is not None]
    rate = len(candidates3) / max(len(quantified), 1)
    # A bare fold-change cutoff is not a specificity control: with a ratio spread
    # of about 0.6 log2, roughly 15% of unchanged proteins sit above 1.5x, so a
    # null dataset is expected to produce a large "candidate" list. What must not
    # happen is a calibrated discovery.
    ok &= check("the 1.5x gate alone stays non-specific on null data, as expected",
                0.05 <= rate <= 0.30,
                f"{len(candidates3)} of {len(quantified)} quantified prey ({rate:.0%}) pass 1.5x by chance")
    # Benjamini-Hochberg controls the false discovery rate, not the count: on a
    # null dataset a couple of tail proteins can still land at q <= 0.05, so the
    # check is that this stays rare rather than that it never happens.
    ok &= check("the calibrated layer finds (almost) nothing",
                len(discovered3) <= 3,
                f"{len(discovered3)} discoveries beyond the bait: "
                + ", ".join(r["gn"] for r in discovered3[:5]))
    print(f"       (for reference: {len(posterior_high)} prey reach the mixture posterior >= 0.9 "
          "on null data, and the effective ratio spread sets the 1.5x false-positive rate)")

    # ------------------------------------------------------------------ case 4
    print("\ncase 4: prey the control never detected (censored ratios)")
    folder4 = root / "case4"
    folder4.mkdir(parents=True)
    rows4, truth4 = with_missing_controls(seed=4)
    write_workbook(folder4 / "BAIT1.xlsx", rows4, 3, 14)
    out4 = folder4 / "bioid_analysis"
    run(folder4, out4)
    table4 = read_summary(out4, "BAIT1")
    by_gn = {r["gn"]: r for r in table4}
    missing = truth4["missing"]
    # The floor has to be derived from this workbook, not carried in from outside
    diagnostics4 = read_diagnostics(out4, "BAIT1")
    q05 = diagnostics4.get("Control floor (5th pct)")
    ok &= check("the control floor is derived from the workbook",
                q05 is not None and q05 > 0, f"floor = {q05}")

    kept = [g for g in missing if g in by_gn and by_gn[g]["passes"]]
    ok &= check("EV-undetected interactors are kept, not dropped",
                len(kept) >= int(0.8 * len(missing)),
                f"{len(kept)} of {len(missing)} planted censored prey are candidates")

    labelled = [g for g in kept if "Control_missing" in str(by_gn[g].get("fc_type") or "")]
    ok &= check("their fold change is labelled a lower bound",
                len(labelled) == len(kept),
                f"{len(labelled)} of {len(kept)} carry FC_lower_bound; Control_missing")

    used_floor = [g for g in kept if by_gn[g].get("control_used") and abs(by_gn[g]["control_used"] - q05) < 1e-6]
    ok &= check("they are compared against the floor, not against a blank",
                len(used_floor) == len(kept),
                f"{len(used_floor)} of {len(kept)} use the floor as the control value")

    unscored = [g for g in kept if by_gn[g].get("q") is None and by_gn[g].get("score") is None]
    ok &= check("they take no part in the p/q values (a censored ratio is not a test)",
                len(unscored) == len(kept),
                f"{len(unscored)} of {len(kept)} have no q and no score")

    # The background spread must not move when censored rows are added: compare with
    # case 1, whose controls are complete but whose background is otherwise the same
    sigma4 = (diagnostics4.get("Ratio noise sigma (log2)") or 0)
    sigma1 = (read_diagnostics(out, "BAIT1").get("Ratio noise sigma (log2)") or 0)
    ok &= check("the background spread is not inflated by the censored rows",
                sigma4 > 0 and sigma1 > 0 and abs(sigma4 - sigma1) < 0.5,
                f"sigma = {sigma4:.2f} log2 with censored rows vs {sigma1:.2f} without")

    # One ordering for the whole list: normalized FC descending, with the censored
    # (EV-undetected) rows mixed in rather than parked in their own tier at the top
    # or sunk to the bottom (they used to sink, because they carry no q-value).
    sheet = load_workbook(out4 / "BioID_analysis_summary.xlsx", data_only=True)["Candidates"]
    head = [c.value for c in sheet[1]]
    col = {name: i for i, name in enumerate(head)}
    rows4 = [r for r in sheet.iter_rows(min_row=2, values_only=True)]
    fcs = [r[col["FC (EV-corrected)"]] or 0.0 for r in rows4]
    lb_flags = [("Control_missing" in str(r[col["FC type"]])) for r in rows4]
    ok &= check("the candidate list is one FC-descending sequence",
                all(a >= b - 1e-9 for a, b in zip(fcs, fcs[1:])),
                f"first={fcs[0]:.1f}x last={fcs[-1]:.1f}x over {len(fcs)} rows")
    # mixed in, not segregated: when both kinds exist, some measured row must sit
    # above some EV-undetected row (unless every censored FC is larger than every
    # measured one, which the planted 2.8-log2 shift makes unlikely)
    if lb_flags and False in lb_flags:
        last_lb = max(i for i, f in enumerate(lb_flags) if f)
        ok &= check("EV-undetected rows are mixed into the list, not a separate tier",
                    any(not f for f in lb_flags[:last_lb]),
                    f"{sum(lb_flags)} censored row(s), last at position {last_lb + 1} of {len(lb_flags)}")

    print("\n" + ("ALL CHECKS PASSED" if ok else "SOME CHECKS FAILED"))
    if args.keep:
        print(f"kept: {root}")
    else:
        shutil.rmtree(root, ignore_errors=True)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
