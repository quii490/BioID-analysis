#!/usr/bin/env python3
"""BioID analysis: CRAPome annotation, SAINT-style scoring, interactive report.

Reads one protein-level Excel workbook per bait condition (same input contract
as bioid-batch-qc), then per bait:

1. corrects the bait/control load difference when the two channels differ in
   total protein abundance (SAINTq's control-normalization idea),
2. annotates every prey with CRAPome control frequency,
3. applies the two hard gates requested for manual review: >=2 unique peptides
   in the experimental channel and >=1.5-fold enrichment after normalization,
4. scores every prey against the control background with a SAINT-style
   two-component mixture model, and
5. writes an Excel summary plus a self-contained interactive HTML report.

The HTML is the report: this script reads and writes no PDF, and needs no browser.
Each figure in the report carries its own "download as PDF" button.

Source workbooks are never modified; all output goes to a new subfolder.

This is NOT the official SAINTq implementation. SAINTq scores peptide or
fragment intensity data; its own paper describes protein-level aggregation as
"SAINTexpress-equivalent" and notes that it underestimates the FDR. With one
bait run and one shared control per condition there are no replicates either,
so the score here is a screening priority, not a calibrated FDR. See
references/methodology.md.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import shutil
import re
import statistics
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SKILL_DIR = Path(__file__).resolve().parent.parent
CRAPOME_TABLE = SKILL_DIR / "assets" / "crapome_frequency.tsv"
ANNOTATION_TABLE = SKILL_DIR / "assets" / "uniprot_annotations.tsv"
PLOTLY_JS = SKILL_DIR / "assets" / "plotly.min.js"
# Vector figure export (MIT): jsPDF + svg2pdf.js, inlined into the report so it works offline
JSPDF_JS = SKILL_DIR / "assets" / "jspdf.umd.min.js"
SVG2PDF_JS = SKILL_DIR / "assets" / "svg2pdf.umd.min.js"
# CJK glyphs for the vector PDF (Noto Sans SC subset, SIL OFL): the jsPDF core fonts have none
NOTO_TTF = SKILL_DIR / "assets" / "noto_sans_sc_subset.ttf"
REPORT_TEMPLATE = SKILL_DIR / "assets" / "report_template.html"

DEFAULTS = {
    "min_unique_peptides": 2,
    "min_fold": 1.5,
    "crapome_frequency": 0.20,
    "crapome_min_sc": 2,
    "crapome_cell": "HEK293",
    "normalize": "median",
    "load_warning_log2": 1.0,
    "plotly_mode": "inline",
}

REPORT_COLORS = {
    "pass": "#157a48",
    "lower_bound": "#0b7285",
    "below_fold": "#93b8dd",
    "crapome": "#d9822b",
    "few_peptides": "#b9bec6",
    "not_quantified": "#e4e7ec",
    "bait": "#8e44ad",
    "bait_self": "#8e44ad",
    # 批次 QC 的「疑似背景」统一用这一种颜色（在散点图上描一圈提醒）
    "qc_bg": "#c92a2a",
}


# --------------------------------------------------------------------------
# small numeric helpers
# --------------------------------------------------------------------------
def to_num(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def to_int(value, default=0):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def median(values):
    values = [v for v in values if v is not None]
    return statistics.median(values) if values else None


def mad_sigma(values):
    """Robust standard-deviation estimate from the median absolute deviation."""
    values = [v for v in values if v is not None]
    if len(values) < 2:
        return None
    med = statistics.median(values)
    return 1.4826 * statistics.median([abs(v - med) for v in values])


def norm_pdf(x, mu, sigma):
    if sigma <= 0:
        sigma = 1e-6
    z = (x - mu) / sigma
    return math.exp(-0.5 * z * z) / (sigma * math.sqrt(2 * math.pi))


def log2_safe(value):
    return math.log2(value) if value and value > 0 else None


def clean_description(text):
    """Trim the UniProt annotation tail so the name reads as a name.

    'Tubulin beta-4B chain OS=Homo sapiens OX=9606 GN=TUBB4B PE=1 SV=2'
    becomes 'Tubulin beta-4B chain'.
    """
    if not text:
        return ""
    return re.split(r"\s+(?:OS|OX|GN|PE|SV)=", text.strip(), maxsplit=1)[0].strip()


# --------------------------------------------------------------------------
# reading source workbooks
# --------------------------------------------------------------------------
def _norm_header(text):
    return re.sub(r"\s+", " ", str(text or "").strip().lower())


def locate_columns(header):
    """Map the analysis columns we need onto 1-based column indexes."""
    found = {}
    for index, raw in enumerate(header, 1):
        name = _norm_header(raw)
        if not name:
            continue
        if name == "accession" and "accession" not in found:
            found["accession"] = index
        elif name == "gn" and "gn" not in found:
            found["gn"] = index
        elif "description" in name and "description" not in found:
            found["description"] = index
        elif name == "# unique peptides" and "unique_peptides" not in found:
            found["unique_peptides"] = index
        elif name == "# peptides" and "peptides" not in found:
            found["peptides"] = index
        elif name in ("# psms", "# psm") and "psms" not in found:
            found["psms"] = index
        elif name.startswith("coverage") and "coverage" not in found:
            found["coverage"] = index
        elif name == "combined_score" and "qc_combined" not in found:
            found["qc_combined"] = index          # batch-QC：abundance/enrichment 百分位的几何平均
        elif name == "batch_bg_flag" and "qc_bg_flag" not in found:
            found["qc_bg_flag"] = index            # batch-QC：High / Review background suspect
        elif name == "cross_sample_qc" and "qc_cross" not in found:
            found["qc_cross"] = index              # batch-QC：交叉样本筛查结论
    raw_abundance = [
        (index, str(raw))
        for index, raw in enumerate(header, 1)
        if isinstance(raw, str)
        # 只认 "Abundance:" / "Abundance " 开头的原始通道，避免把
        # "Abundances (Normalized):" 当成原始列（与 bioid-batch-qc 的模式一致）。
        and re.match(r"abundance[:\s]", raw.strip().lower())
        and "normalized" not in raw.lower()
    ]
    if len(raw_abundance) < 2:
        raise ValueError(
            "At least two raw 'Abundance:' columns (experimental then control) are required; "
            "0-1 found. If this workbook came from an old openxlsx version, its dimension may be "
            "written as A1 (only one column is visible) — re-run the 'bioid' preprocessing stage "
            "to rebuild it"
        )
    found["bait_abundance"] = raw_abundance[0][0]
    found["control_abundance"] = raw_abundance[1][0]
    found["bait_header"] = raw_abundance[0][1]
    found["control_header"] = raw_abundance[1][1]
    return found


def read_bait_map(path):
    """读人工诱饵映射表：<文件名/条件名>  <基因名>，可用 # 注释。

    文件名推断有时不可靠（例如 MCF1-1 的推断写成 NCOA3，实际诱饵是 AKAP8L），
    这张表用于在不改名的情况下覆盖。
    """
    mapping = {}
    if path is None:
        return mapping
    if not Path(path).expanduser().exists():
        raise SystemExit(f"bait map 文件不存在: {path}")
    for line in Path(path).expanduser().read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split()
        if len(parts) < 2:
            raise SystemExit(f"bait map 行无法解析（需要：<文件名或条件名> <基因>）: {line}")
        mapping[parts[0]] = parts[1]
    return mapping


def bait_override(mapping, path):
    """按文件名、去后缀名或目录中的准条件名匹配人工映射。"""
    if not mapping:
        return None
    stem = path.stem
    stripped = re.sub(r"(_processed|_flagged)+$", "", stem, flags=re.IGNORECASE)
    without_flagged = re.sub(r"_flagged$", "", stem, flags=re.IGNORECASE)
    for key in (path.name, stem, stripped, without_flagged):
        if key in mapping:
            return mapping[key]
    return None


def read_workbook(path, sheet="Proteins", bait_map=None):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet in wb.sheetnames:
            ws = wb[sheet]
        elif "+5%" in wb.sheetnames:
            ws = wb["+5%"]
        else:
            raise ValueError(f"{path.name}: no '{sheet}' or '+5%' sheet")
        rows = ws.iter_rows(values_only=True)
        header = list(next(rows))
        cols = locate_columns(header)
        proteins = []
        for row in rows:
            if row is None:
                continue
            def cell(key):
                index = cols.get(key)
                if index is None or index > len(row):
                    return None
                return row[index - 1]
            gn = str(cell("gn") or "").strip()
            accession = str(cell("accession") or "").strip()
            if not gn and not accession:
                continue
            proteins.append(
                {
                    "accession": accession,
                    "uniprot": accession.split(";")[0].split("-")[0].strip(),
                    "gn": gn,
                    "description": str(cell("description") or "").strip(),
                    "coverage": to_num(cell("coverage")),
                    "n_peptides": to_int(cell("peptides")),
                    "n_unique_peptides": to_int(cell("unique_peptides")),
                    "n_psms": to_int(cell("psms")),
                    "bait_abundance": to_num(cell("bait_abundance")),
                    "control_abundance": to_num(cell("control_abundance")),
                    "qc_combined": to_num(cell("qc_combined")),
                    "qc_bg_flag": str(cell("qc_bg_flag") or "").strip(),
                    "qc_cross": str(cell("qc_cross") or "").strip(),
                }
            )
    finally:
        wb.close()
    bait, isoform = split_bait_name(path.stem)
    override = bait_override(bait_map or {}, path)
    if override:
        bait, isoform = override, None
    match = re.search(r"\bF(\d+)\b", cols["bait_header"])
    return {
        "file": path.name,
        "bait": bait,
        "isoform": isoform,
        "order": int(match.group(1)) if match else 0,
        "bait_header": cols["bait_header"],
        "control_header": cols["control_header"],
        "bait_channel": channel_id(cols["bait_header"]),
        "control_channel": channel_id(cols["control_header"]),
        "proteins": proteins,
    }


def channel_id(header):
    """'Abundance: F12: Sample' -> 'F12'."""
    match = re.search(r"\bF(\d+)\b", str(header or ""))
    return f"F{match.group(1)}" if match else str(header or "")


SIBLING_QC = Path(__file__).resolve().parents[2] / "bioid-batch-qc" / "scripts" / "bioid_batch_review.py"


def _condition_key(path):
    """工作簿 -> 条件名：去掉 _processed / _flagged 后缀。"""
    return re.sub(r"(_processed|_flagged)+$", "", path.stem, flags=re.IGNORECASE)


def qc_coverage_gap(folder, qc_files):
    """QC 产物是否覆盖了目录里全部 *_processed 工作簿（按条件名比）。"""
    plain = {_condition_key(p) for p in folder.glob("*_processed.xlsx")
             if not p.name.startswith(("~$", ".~"))}
    if not plain:
        return []
    return sorted(plain - {_condition_key(p) for p in qc_files})


def discover_workbooks(folder, from_qc=False, no_qc=False, allow_no_qc=False,
                       list_only=False):
    """决定读哪一批工作簿，并返回一句说明。

    bioid-batch-qc 把结果写在 <folder>/flagged/*_flagged.xlsx：只有这些工作簿带
    Combined_score / Batch_BG_flag / Cross_sample_QC。默认会优先用它们（并说明），
    --from-qc 强制要求它们存在，--no-qc 则忽略。
    """
    def flagged(path):
        return sorted(p for p in path.glob("*_flagged.xlsx") if not p.name.startswith(("~$", ".~")))

    def plain(path):
        return sorted(p for p in path.glob("*.xlsx")
                      if not p.name.startswith(("~$", ".~")) and not p.name.endswith("_flagged.xlsx"))

    nested = flagged(folder / "flagged")
    inline = flagged(folder)
    if from_qc:
        if nested:
            return nested, f"batch-QC 输出 {folder / 'flagged'}（{len(nested)} 份 *_flagged.xlsx）"
        if inline:
            return inline, f"batch-QC 输出（本目录 {len(inline)} 份 *_flagged.xlsx）"
        raise SystemExit(
            f"--from-qc: {folder}/flagged/ 下没有 *_flagged.xlsx；请先跑 bioid-batch-qc"
        )
    if no_qc:
        return plain(folder), "本目录原始工作簿（--no-qc：明确忽略 batch-QC 输出）"
    if nested:
        return nested, (f"检测到 batch-QC 输出 {folder / 'flagged'}（{len(nested)} 份），本次改用它；"
                        "要忽略请加 --no-qc")
    if inline:
        return inline, (f"本目录含 batch-QC 输出（{len(inline)} 份 *_flagged.xlsx），本次只用它们；"
                        "要改用同目录其他工作簿请加 --no-qc")
    if allow_no_qc or list_only:
        return plain(folder), None
    raise SystemExit(
        "没有找到 batch-QC 输出：<folder>/flagged/ 下没有 *_flagged.xlsx。\n"
        "  Background flag / Cross-sample QC / Combined score 三列只有在跑过 bioid-batch-qc 之后才存在，\n"
        "  没有它们的结果会被当成完整结果读。\n"
        "  二选一：\n"
        f"    1) 先跑 batch-QC：python3 {SIBLING_QC} <folder>\n"
        "       （或直接用 bioid_run.py 一条命令跑完预处理→QC→分析）\n"
        "    2) 明确接受缺 QC：加 --allow-no-qc（三项 QC 列会留空，报告顶部会标注）"
    )


def split_bait_name(stem):
    """Filename stem -> (bait label, isoform number).

    A leading numeric order prefix is dropped, so 2PM20D2 -> PM20D2, and the
    ``<order>-<isoform><bait>`` convention (18-2FAM227A) also yields the isoform
    number, which is what keeps two files for one gene apart as conditions.
    """
    # 同时剥掉预处理（_processed）与批次 QC（_flagged）后缀，两条链路的条件名保持一致
    stem = re.sub(r"(_processed|_flagged)+$", "", stem, flags=re.IGNORECASE)
    isoform = re.match(r"^[0-9]+-([0-9]+)(?=[A-Za-z])", stem)
    bait = re.sub(r"^[0-9-]+", "", stem)
    return bait or stem, (int(isoform.group(1)) if isoform else None)


def assign_conditions(workbooks):
    """Give every workbook a unique condition label and a shared family.

    Files that resolve to the same gene become separate isoform conditions but
    one family, so isoform rows are never treated as two independent baits and
    two different files can never end up with the same label.
    """
    grouped = {}
    for item in workbooks:
        grouped.setdefault(item["bait"], []).append(item)
    used = set()
    for bait, members in grouped.items():
        members.sort(key=lambda x: (x["order"], x["file"]))
        for index, item in enumerate(members, 1):
            if len(members) == 1:
                label = bait
            else:
                number = item["isoform"] or item["order"] or index
                label = f"{bait}_isoform_{number}"
            if label in used:
                # 一个诱饵对多个 EV 时 F 号相同（同一诱饵通道），用对照通道把条件名
                # 区分开；family 仍是同一个，因此背景频率不会被同一个诱饵数两次。
                token = item.get("control_channel") or f"ev{index}"
                label = f"{label}_{token}"
            while label in used:
                label = f"{label}_x"
            used.add(label)
            item["condition"] = label
            item["family"] = bait
    labels = [item["condition"] for item in workbooks]
    duplicates = {label for label in labels if labels.count(label) > 1}
    if duplicates:
        detail = ", ".join(sorted(duplicates))
        raise ValueError(
            f"condition labels collide ({detail}); rename the files as "
            "<order>-<isoform><gene> with distinct prefixes"
        )
    return workbooks


# --------------------------------------------------------------------------
# load normalization
# --------------------------------------------------------------------------
def normalize_loads(proteins, mode, warning_log2):
    """Scale the control channel onto the experimental channel.

    A bait and an EV channel that were loaded differently make every fold
    change shift in the same direction, which is exactly what a global
    correction removes. 'total' matches the summed signal of the two channels
    (the SAINTq control-normalization idea); 'median' uses the median per
    protein ratio, which is less sensitive to a few very abundant proteins.
    """
    diagnoses = _load_estimates(proteins)
    diagnostics = {
        "mode": mode,
        "bait_total": diagnoses["bait_total"],
        "control_total": diagnoses["control_total"],
        "load_ratio_total": diagnoses["load_ratio_total"],
        "load_ratio_median": diagnoses["load_ratio_median"],
        "correction_total": diagnoses["correction_total"],
        "correction_median": diagnoses["correction_median"],
        "correction_robust": diagnoses["correction_robust"],
        "concentration_top10": diagnoses["concentration_top10"],
        "ctrl_q05": diagnoses["ctrl_q05"],
        "repeated_control": diagnoses["repeated_control"],
        "correction_log2": 0.0,
        "warnings": [],
    }
    by_mode = {
        "total": diagnoses["correction_total"],
        "median": diagnoses["correction_median"],
        "robust": diagnoses["correction_robust"],
        "none": 0.0,
    }
    diagnostics["correction_log2"] = by_mode[mode] or 0.0

    applied = diagnostics["correction_log2"]
    if mode != "none" and abs(applied) >= warning_log2 / 2:
        diagnostics["warnings"].append(
            f"a {mode} load correction of {applied:+.2f} log2 ({2 ** abs(applied):.1f}-fold) "
            f"was subtracted from every fold change before the gates were applied"
        )
    spread = max(by_mode.values()) - min(by_mode.values())
    if mode != "none" and spread >= 1.0:
        diagnostics["warnings"].append(
            f"load estimators disagree over a {2 ** spread:.1f}-fold range "
            f"(total {by_mode['total']:+.2f}, median {by_mode['median']:+.2f}, "
            f"robust {by_mode['robust']:+.2f} log2), so the candidate list depends on this choice; "
            "the overview sheet reports counts for each estimator"
        )
    if diagnoses["repeated_control"]:
        value, count, total = diagnoses["repeated_control"]
        diagnostics["warnings"].append(
            f"the control channel repeats the value {value:.6g} in {count} of {total} "
            "quantified prey, which is what an upstream floor imputation looks like. If "
            "your preprocessing already substituted a floor, those rows are read here as "
            "measured controls: they keep FC_type 'Measured', take part in the background "
            "spread and in the p/q values, and this skill's own 'Control_missing' "
            "labelling never fires. Say which convention applies before comparing with "
            "another tool."
        )
    if diagnoses["concentration_top10"] and diagnoses["concentration_top10"] >= 0.5:
        diagnostics["warnings"].append(
            f"the 10 most abundant proteins carry {diagnoses['concentration_top10']:.0%} of the "
            "experimental signal, so a total-sum correction is unreliable here"
        )
    return diagnostics


def _load_estimates(proteins):
    """Three independent estimates of the bait/control scale difference.

    'total' matches summed signal, which is what a loading difference would do,
    but a handful of very abundant proteins usually dominate the sums. 'median'
    takes the median log2 ratio over every paired protein, and 'robust' restricts
    that median to the better-quantified half of the proteome. The spread between
    them shows how much the choice matters for the fold-change gate.
    """
    paired = [
        (p["bait_abundance"], p["control_abundance"])
        for p in proteins
        if p["bait_abundance"] and p["bait_abundance"] > 0
        and p["control_abundance"] and p["control_abundance"] > 0
    ]
    bait_total = sum(b for b, _ in paired)
    control_total = sum(c for _, c in paired)
    ratios = [math.log2(b / c) for b, c in paired]
    ranked = sorted(paired, key=lambda bc: math.sqrt(bc[0] * bc[1]), reverse=True)
    well_quantified = ranked[: max(20, len(ranked) // 2)]
    top10 = sorted(paired, key=lambda bc: bc[0], reverse=True)[:10]
    controls = [
        p["control_abundance"] for p in proteins
        if p["control_abundance"] and p["control_abundance"] > 0
    ]
    # A real abundance column has ~no repeated values; a control channel with one value
    # repeating across many rows is the signature of an upstream floor imputation. The
    # pipeline cannot tell an imputed floor from a measurement, so it says so: if the
    # user's own workflow already substituted floors, the report's 'Measured' rows for
    # those prey are lower bounds in disguise, and this skill's own 'Control_missing'
    # labelling never fires because the cell is not empty.
    repeated = None
    if len(controls) >= 100:
        value, count = Counter(controls).most_common(1)[0]
        if count >= max(20, 0.02 * len(controls)):
            repeated = (value, count, len(controls))
    return {
        "repeated_control": repeated,
        "bait_total": bait_total or None,
        "control_total": control_total or None,
        "ctrl_q05": float(np.quantile(controls, 0.05)) if controls else None,
        "load_ratio_total": (control_total / bait_total) if bait_total > 0 else None,
        "load_ratio_median": (2 ** statistics.median(ratios)) if ratios else None,
        "correction_total": math.log2(control_total / bait_total) if bait_total > 0 and control_total > 0 else None,
        "correction_median": statistics.median(ratios) if ratios else None,
        "correction_robust": (
            statistics.median([math.log2(b / c) for b, c in well_quantified])
            if well_quantified else None
        ),
        "concentration_top10": (sum(b for b, _ in top10) / bait_total) if bait_total > 0 else None,
    }


def impute_missing_controls(proteins, diagnostics):
    """Give prey that the control never detected a control floor, not a blank.

    "Detected in the bait, absent from the EV" is the strongest single-run
    evidence a BioID experiment offers, and dropping those rows hides exactly
    the class the experiment is looking for. The control value is censored, not
    missing: the ratio is at least bait/q05, never less. The batch QC skill uses
    the same floor (the 5th percentile of the positive control values in the same
    workbook), so the two skills agree on what a lower-bound fold change means --
    a floor, never a measurement, and never a p-value input.
    """
    q05 = diagnostics.get("ctrl_q05")
    imputed = 0
    for p in proteins:
        p["control_used"] = p["control_abundance"]
        p["control_imputed"] = False
        p["fc_is_lower_bound"] = False
        p["fc_type"] = ""
        if p["bait_abundance"] and p["bait_abundance"] > 0 and not (
                p["control_abundance"] and p["control_abundance"] > 0) and q05:
            p["control_used"] = q05
            p["control_imputed"] = True
            p["fc_is_lower_bound"] = True
            imputed += 1
    diagnostics["n_control_imputed"] = imputed
    if imputed:
        diagnostics["warnings"].append(
            f"{imputed} prey were detected in the experimental channel but not in the "
            f"control; they were given the 5th-percentile control floor "
            f"({q05:.3g}) and their fold changes are reported as lower bounds "
            "('Control_missing'), not as measurements. They are excluded from the "
            "background spread and from every p/q value."
        )
    return proteins


def add_fold_changes(proteins, diagnostics):
    """Attach raw and normalized fold changes, including both reference modes."""
    correction = diagnostics["correction_log2"]
    for p in proteins:
        bait, control = p["bait_abundance"], p.get("control_used")
        if bait and bait > 0 and control and control > 0:
            p["fc_raw"] = bait / control
            p["log2fc_raw"] = log2_safe(p["fc_raw"])
            p["log2fc_norm"] = p["log2fc_raw"] - correction
            p["fc_norm"] = 2 ** p["log2fc_norm"]
            p["fc_type"] = "FC_lower_bound; Control_missing" if p.get("fc_is_lower_bound") else "Measured"
            for name in ("median", "total"):
                reference = diagnostics.get(f"correction_{name}")
                p[f"fc_norm_{name}"] = 2 ** (p["log2fc_raw"] - reference) if reference is not None else None
        else:
            p["fc_raw"] = None
            p["log2fc_raw"] = None
            p["log2fc_norm"] = None
            p["fc_norm"] = None
            p["fc_norm_median"] = None
            p["fc_norm_total"] = None
    return proteins


def candidate_counts_by_mode(proteins, diagnostics, min_fold, min_unique_peptides):
    """How many prey would pass the gates under each load estimator."""
    counts = {}
    for mode in ("median", "robust", "total"):
        correction = diagnostics.get(f"correction_{mode}")
        if correction is None:
            continue
        passing = 0
        for p in proteins:
            if not (p["bait_abundance"] and p["bait_abundance"] > 0):
                continue
            if p["n_unique_peptides"] < min_unique_peptides or p["crapome_flag"] or p.get("is_bait"):
                continue
            if p["log2fc_raw"] is None:
                continue
            if 2 ** (p["log2fc_raw"] - correction) < min_fold:
                continue
            passing += 1
        counts[mode] = passing
    return counts


# --------------------------------------------------------------------------
# CRAPome
# --------------------------------------------------------------------------
def load_crapome(path, cell="HEK293"):
    """Load the compact CRAPome frequency table (see references/crapome.md)."""
    if not path or not Path(path).exists():
        return None
    table = {"by_uniprot": {}, "by_gene": {}, "cell": cell, "source": str(path), "meta": []}
    with open(path, newline="") as fh:
        lines = [line for line in fh if not line.startswith("#")]
    reader = csv.DictReader(lines, delimiter="\t")
    runs_all = runs_hek = 0
    for row in reader:
        gene = (row.get("gene") or "").strip().upper()
        uniprot = (row.get("uniprot") or "").strip().upper()
        runs_all = to_int(row.get("n_ctrl_total"), runs_all)
        runs_hek = to_int(row.get("n_ctrl_hek_total"), runs_hek)
        row["_hits_all"] = to_int(row.get("n_ctrl_ge2"))
        row["_hits_hek"] = to_int(row.get("n_ctrl_hek_ge2"))
        if uniprot and ";" not in uniprot:
            table["by_uniprot"].setdefault(uniprot, row)
        if gene:
            table["by_gene"].setdefault(gene, row)
    table["runs_all"] = runs_all
    table["runs_hek"] = runs_hek
    return table


def annotate_crapome(proteins, table, cell, frequency_threshold, min_sc):
    """Attach control-run frequency; never drop a protein, only label it."""
    if table is None:
        for p in proteins:
            p.update(
                crapome_status="not applied (table unavailable)",
                crapome_frequency=None,
                crapome_hits=None,
                crapome_runs=None,
                crapome_flag="",
            )
        return {"applied": False, "n_flagged": 0}
    use_hek = bool(cell) and str(cell).lower().startswith("hek")
    runs = table["runs_hek"] if use_hek else table["runs_all"]
    hits_key = "_hits_hek" if use_hek else "_hits_all"
    subset = cell if use_hek else "all human"
    for p in proteins:
        row = table["by_uniprot"].get(p["uniprot"].upper()) or table["by_gene"].get(p["gn"].upper())
        if row is None:
            p.update(
                crapome_status="not in CRAPome",
                crapome_frequency=0.0,
                crapome_hits=0,
                crapome_runs=runs,
                crapome_flag="",
            )
            continue
        hits = to_int(row.get(hits_key))
        frequency = (hits / runs) if runs else None
        p.update(
            crapome_status=f"{subset} controls",
            crapome_frequency=frequency,
            crapome_hits=hits,
            crapome_runs=runs,
            crapome_flag="contaminant"
            if (frequency is not None and frequency >= frequency_threshold and hits >= min_sc)
            else "",
        )
    return {"applied": True, "n_flagged": sum(1 for p in proteins if p["crapome_flag"])}


# --------------------------------------------------------------------------
# protein annotation (hover cards)
# --------------------------------------------------------------------------
def load_annotations(path):
    """Load the UniProt annotation cache: accession -> (name, gene, aliases, function)."""
    if not path or not Path(path).exists():
        return {}
    table = {}
    with open(path, newline="") as fh:
        lines = [line for line in fh if not line.startswith("#")]
    for row in csv.DictReader(lines, delimiter="\t"):
        accession = (row.get("accession") or "").strip()
        if accession:
            table[accession] = [
                (row.get("name") or "").strip(),
                (row.get("gene") or "").strip(),
                (row.get("aliases") or "").strip(),
                (row.get("function") or "").strip(),
            ]
    return table


def build_annotation_payload(results, table):
    """Share one entry per accession instead of repeating it in every condition."""
    shared = {}
    for result in results:
        for protein in result["proteins"]:
            accession = protein["uniprot"]
            if not accession or accession in shared:
                continue
            entry = table.get(accession)
            if entry:
                shared[accession] = entry
    return shared


# --------------------------------------------------------------------------
# SAINT-style mixture model
# --------------------------------------------------------------------------
def fit_mixture(values, mu0, sigma0, max_iter=300, tol=1e-7, pi_prior=(1.0, 20.0)):
    """EM for a two-component Gaussian mixture with the null fixed at zero.

    Component 0 (no change) is the background: centred on the corrected ratio
    zero with a width taken from the negative half of the distribution, which
    enrichment cannot inflate. Component 1 (enrichment) is fitted by EM.

    The mixing proportion carries a Beta prior, as the SAINT family does. Without
    it a dataset with no interactors still yields a non-empty "enriched"
    component fitted to the positive tail of the noise, which is how a method
    invents hits. The default prior mean is about 5% of preys enriched.
    """
    x = np.asarray([v for v in values if v is not None], dtype=float)
    if x.size < 10 or sigma0 <= 0:
        return None
    pi = 0.05
    mu1 = float(np.median(x)) + 1.0
    sigma1 = max(float(np.std(x)), 0.5)
    for _ in range(max_iter):
        f1 = pi * _normal_pdf_arr(x, mu1, sigma1)
        f0 = (1.0 - pi) * _normal_pdf_arr(x, mu0, sigma0)
        denom = f1 + f0
        denom[denom <= 0] = 1e-300
        gamma = f1 / denom
        weight = gamma.sum()
        # MAP update of the mixing proportion under Beta(a, b).
        prior_a, prior_b = pi_prior
        new_pi = float(np.clip(
            (gamma.sum() + prior_a - 1.0) / (gamma.size + prior_a + prior_b - 2.0),
            1e-5, 0.5,
        ))
        if weight < 1e-6:
            break
        new_mu1 = float((gamma * x).sum() / weight)
        new_sigma1 = max(float(math.sqrt((gamma * (x - new_mu1) ** 2).sum() / weight)), 0.25)
        if (
            abs(new_pi - pi) < tol
            and abs(new_mu1 - mu1) < tol
            and abs(new_sigma1 - sigma1) < tol
        ):
            pi, mu1, sigma1 = new_pi, new_mu1, new_sigma1
            break
        pi, mu1, sigma1 = new_pi, new_mu1, new_sigma1
    return {"pi": pi, "mu1": mu1, "sigma1": sigma1, "mu0": mu0, "sigma0": sigma0}


def _normal_pdf_arr(x, mu, sigma):
    sigma = max(float(sigma), 1e-6)
    z = (x - mu) / sigma
    return np.exp(-0.5 * z * z) / (sigma * math.sqrt(2 * math.pi))


def _norm_isf(p):
    """Inverse standard-normal CDF by bisection on erfc (scipy is not required)."""
    if p <= 0:
        return float("inf")
    if p >= 1:
        return float("-inf")
    lo, hi = -40.0, 40.0
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if 0.5 * math.erfc(mid / math.sqrt(2)) > p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2.0


def noise_summary(proteins, diagnostics, min_fold, min_unique_peptides):
    """How noisy the ratio distribution is, and what threshold that implies.

    The spread of the corrected log2 ratio sets the fold change a protein needs
    before it is unusual at all; on single-run data this is often far above the
    literature-style 1.5x cutoff, which is why the cutoff alone cannot be read as
    evidence. The q=0.05 threshold is the smallest fold change whose one-sided
    p-value survives Benjamini-Hochberg over the quantified prey.
    """
    corrected = [
        p["log2fc_norm"] for p in proteins
        if p["log2fc_norm"] is not None and not p.get("fc_is_lower_bound")
    ]
    if len(corrected) < 20:
        return {}
    negative = [abs(v) for v in corrected if v <= 0]
    sigma0 = 1.4826 * statistics.median(negative) if len(negative) >= 20 else None
    if not sigma0 or sigma0 <= 0:
        return {}
    gate = math.log2(min_fold)
    above = sum(1 for v in corrected if v >= gate) / len(corrected)
    z_q05 = _norm_isf(0.05 / len(corrected))
    # The strictest BH threshold (rank 1) is 2^(z*sigma); a prey further down the
    # ranking can be significant at a smaller fold change. Quoting only the strict
    # number invited the reader to compare it with a table row that passed at a
    # lower FC, so the smallest FC among the prey that actually reach q <= 0.05 is
    # reported alongside it.
    # Restricted to prey that pass the gates, so the number can be checked against
    # the candidate table. A significant protein that fails a gate (too few
    # peptides, a contaminant) is real but does not belong on that list, and
    # quoting it here read as a contradiction.
    significant = [
        p["fc_norm"] for p in proteins
        if p.get("passes") and p.get("q_value") is not None
        and p["q_value"] <= 0.05 and p.get("fc_norm")
    ]
    lower_bound_gate = [
        p for p in proteins
        if p.get("passes") and p.get("fc_is_lower_bound")
    ]
    return {
        "sigma0": sigma0,
        "share_above_gate": above,
        "fold_at_q05": 2 ** (z_q05 * sigma0),
        "fold_at_q05_raw": 2 ** (z_q05 * sigma0 + diagnostics["correction_log2"]),
        "z_q05": z_q05,
        "n_significant": len(significant),
        "min_fc_significant": min(significant) if significant else None,
        "n_lower_bound_gate": len(lower_bound_gate),
    }


def saint_scores(proteins, diagnostics):
    """Score every prey of one condition and add a Bayesian FDR.

    The modelled quantity is the normalized log2 ratio of each prey against its
    own control, not the absolute abundance: abundance spans orders of magnitude
    across the proteome, and that spread would swamp the enrichment signal. After
    the load correction the no-change background is centred on zero, and its
    width is taken from the robust spread of the ratio distribution, so the
    mixture separates "unchanged" from "enriched" the way the SAINT family does.
    """
    for p in proteins:
        p["saint_score"] = None
        p["saint_bfdr"] = None
        p["p_one_sided"] = None
        p["q_value"] = None

    # A lower-bound ratio is a censored observation, not a measurement: including
    # it would widen the null (inflating sigma0) and would also take a test slot
    # in the BH correction while carrying no p-value of its own. Those prey keep
    # q = None and are reported as their own class instead.
    paired = [
        p for p in proteins
        if p["log2fc_raw"] is not None and not p.get("fc_is_lower_bound")
    ]
    if len(paired) < 20:
        return {"fitted": False, "reason": "fewer than 20 prey quantified in both channels"}

    correction = diagnostics["correction_log2"]
    centered = [p["log2fc_raw"] - correction for p in paired]
    mu0 = 0.0
    # Enrichment only pushes ratios up, so the spread of the null can be read
    # from the negative half alone; using the whole distribution would let the
    # very hits we are looking for widen the background and hide themselves.
    negative_half = [abs(v) for v in centered if v <= mu0]
    if len(negative_half) >= 20:
        sigma0 = 1.4826 * float(np.median(negative_half))
    else:
        sigma0 = float(mad_sigma(centered) or 0.0)
    if sigma0 <= 0:
        sigma0 = float(np.std(centered))
    if sigma0 <= 0:
        return {"fitted": False, "reason": "the ratio distribution has no usable spread"}

    # Calibrated companion: a one-sided p-value from the null estimated on the
    # negative half, then Benjamini-Hochberg. The mixture posterior answers
    # "does this look like the enriched cluster", which needs the cluster to be
    # identifiable; the q-value answers "is this ratio surprising given how
    # ratios behave when nothing is enriched", which one run can support.
    n_tests = len(centered)
    pvals = {}
    for p in paired:
        x = p["log2fc_raw"] - correction
        z = x / sigma0 if sigma0 > 0 else 0.0
        pvals[id(p)] = 0.5 * math.erfc(z / math.sqrt(2))
    ordered = sorted(paired, key=lambda p: pvals[id(p)])
    running_min = 1.0
    for rank, p in enumerate(reversed(ordered), 1):
        index = n_tests - rank + 1
        q = pvals[id(p)] * n_tests / index
        running_min = min(running_min, q)
        p["q_value"] = min(running_min, 1.0)
        p["p_one_sided"] = pvals[id(p)]

    model = fit_mixture(centered, mu0, sigma0)
    if model is None:
        return {
            "fitted": False,
            "reason": "mixture model did not converge",
            "calibrated": True,
        }

    # Refuse to report an "enriched" component that is not separated from the
    # no-change background, which is what a dataset without interactors yields.
    separation = model["mu1"] / sigma0 if sigma0 > 0 else 0.0
    model["separation_sigma"] = separation
    model["sigma0"] = sigma0
    if separation < 1.0 or model["pi"] < 0.002:
        return {
            "fitted": False,
            "reason": (
                f"no enriched component separated from background: the fitted shift is "
                f"{model['mu1']:.2f} log2 against a background spread of {sigma0:.2f} log2 "
                f"({separation:.2f} sigma), so a mixture cannot separate the two"
            ),
            "model": model,
            "calibrated": True,
        }

    for p in proteins:
        # A censored ratio is a floor, not a measurement: scoring it would turn
        # "the control never saw this protein" into a posterior near 1 and put a
        # number on it that the data does not support.
        if p["log2fc_raw"] is None or p.get("fc_is_lower_bound"):
            continue
        value = p["log2fc_raw"] - correction
        f1 = model["pi"] * norm_pdf(value, model["mu1"], model["sigma1"])
        f0 = (1 - model["pi"]) * norm_pdf(value, model["mu0"], model["sigma0"])
        p["saint_score"] = f1 / (f1 + f0) if (f1 + f0) > 0 else 0.0

    scored = sorted(
        (p for p in proteins if p["saint_score"] is not None),
        key=lambda p: p["saint_score"],
        reverse=True,
    )
    n = len(scored)
    running = 0.0
    for rank, p in enumerate(scored, 1):
        running += 1.0 - p["saint_score"]
        p["saint_bfdr"] = running / rank if rank else None
    # enforce monotonicity from the most stringent threshold downwards
    floor = 1.0
    for p in reversed(scored):
        floor = min(floor, p["saint_bfdr"])
        p["saint_bfdr"] = floor
    model["n_scored"] = n
    model["sigma0"] = sigma0
    return {"fitted": True, "model": model}


# --------------------------------------------------------------------------
# gates and status
# --------------------------------------------------------------------------
def tag_bait(proteins, condition, aliases=None):
    """Mark the bait's own row so it is never counted as a candidate.

    `aliases` 处理"构建名与数据库符号不一致"的情况：文件/标签写作 FAM214A、mKIAA1614、
    SPANXA2，而当前 UniProt 把同一蛋白列为 ATOSA、KIAA1614、SPANXA1。别名只影响排除与
    诱饵标记，条件名不变（保持你自己的命名）。
    """
    wanted = {str(condition).upper()} | {str(a).upper() for a in (aliases or [])}
    for p in proteins:
        p["is_bait"] = bool(p["gn"]) and p["gn"].upper() in wanted
    return proteins


def read_bait_aliases(path):
    """读诱饵别名表：每行 <标签> <数据库符号>[,<符号>…]，可用 # 注释。"""
    aliases = {}
    if not path:
        return aliases
    if not Path(path).expanduser().exists():
        raise SystemExit(f"bait alias 文件不存在: {path}")
    for line in Path(path).expanduser().read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.replace(",", " ").split()
        if len(parts) < 2:
            raise SystemExit(f"bait alias 行无法解析（需要：<标签> <数据库符号>）: {line}")
        aliases.setdefault(parts[0], set()).update(parts[1:])
    return aliases


def apply_gates(proteins, args, condition):
    for p in proteins:
        reasons = []
        notes = []
        quantified = bool(p["bait_abundance"] and p["bait_abundance"] > 0)
        if not quantified:
            reasons.append("not quantified in experimental channel")
        if p["n_unique_peptides"] < args.min_unique_peptides:
            reasons.append(
                f"only {p['n_unique_peptides']} unique peptide(s) in the experiment "
                f"(need {args.min_unique_peptides})"
            )
        if p["fc_norm"] is None:
            reasons.append("no control value and no control floor to compare against")
        elif p.get("fc_is_lower_bound"):
            # The gate is applied to the lower bound, so passing it is a statement
            # about the worst case: the real ratio can only be larger.
            if p["fc_norm"] < args.min_fold:
                reasons.append(
                    f"enrichment {p['fc_norm']:.2f}x below {args.min_fold}x "
                    "(lower bound: the control was not detected)"
                )
            else:
                notes.append(
                    f"lower bound: the control was not detected, so the "
                    f"{p['fc_norm']:.2f}x enrichment is a floor, not a measurement"
                )
        elif p["fc_norm"] < args.min_fold:
            reasons.append(f"enrichment {p['fc_norm']:.2f}x below {args.min_fold}x")
        if p["crapome_flag"]:
            reasons.append("CRAPome contaminant")
        if p.get("is_bait"):
            reasons.append("bait self-hit")
        p["pass_reasons"] = reasons
        p["pass_notes"] = notes
        p["passes"] = not reasons
    return proteins


# --------------------------------------------------------------------------
# outputs
# --------------------------------------------------------------------------
def _top_fc_label(top):
    """Overview 的 top FC 文本；≥ 前缀表示对照未检出（下界）。"""
    if not top:
        return None
    prefix = "≥" if top.get("fc_is_lower_bound") else ""
    return f"{prefix}{top['fc_norm']:.1f}x"


def _candidate_order(protein):
    """Order the candidate list by fold change, with the censored rows mixed in.

    The user-facing question this list answers is "what is most enriched", and a prey
    the control never detected answers that just as directly as a measured one — its
    fold change is a floor, displayed with a >= prefix, not a different quantity. So
    there is one ordering for everything: normalized FC descending, then raw FC, then
    experimental abundance, all as tiebreakers. (An earlier revision led with the
    EV-undetected rows as a separate tier, which read as a second list; the user asked
    for them to be sorted together.) The SAINT-style score stays a column, not the sort
    key: on batches where no mixture fits, ordering by score degenerates to source-file
    order, while FC is always meaningful.
    """
    fc = protein.get("fc_norm")
    if fc is None:
        return (1, 0.0, 0.0, 0.0)
    return (0, -fc, -(protein.get("fc_raw") or 0.0), -(protein.get("bait_abundance") or 0.0))


CANDIDATE_COLUMNS = [
    ("GN", "gn"),
    ("Accession", "accession"),
    ("Description", "description"),
    ("FC (EV-corrected)", "fc_norm"),
    ("log2FC (EV-corrected)", "log2fc_norm"),
    ("FC (raw)", "fc_raw"),
    ("SAINT-style score", "saint_score"),
    ("BFDR (approx)", "saint_bfdr"),
    ("p (one-sided)", "p_one_sided"),
    ("q (BH)", "q_value"),
    ("Unique peptides", "n_unique_peptides"),
    ("Peptides", "n_peptides"),
    ("PSMs", "n_psms"),
    ("Coverage [%]", "coverage"),
    ("Experimental abundance", "bait_abundance"),
    ("Control abundance", "control_abundance"),
    ("Control used", "control_used"),
    ("FC type", "fc_type"),
    ("Gate notes", "pass_notes"),
    ("CRAPome frequency", "crapome_frequency"),
    ("CRAPome hits >=2SC", "crapome_hits"),
    ("CRAPome control runs", "crapome_runs"),
    ("CRAPome status", "crapome_status"),
    # 批次 QC（bioid-batch-qc 的输出才有；没有时这三列为空）
    ("Combined score", "qc_combined"),
    ("Background flag", "qc_bg_flag"),
    ("Cross-sample QC", "qc_cross"),
]


def write_excel(results, path, args, crapome_info, crapome_source):
    wb = Workbook()
    overview = wb.active
    overview.title = "Overview"
    headers = [
        "Order", "Condition", "File", "Experimental channel", "Control channel",
        "Quantified (exp)", "Quantified (control)", "Paired (measured)",
        "Control not detected (floored)",
        "Control floor (5th pct)", "Lower-bound candidates",
        "Experimental total", "Control total", "Summed signal control/experimental ratio",
        "Load correction applied (log2)", "Normalization mode", "Fold scale applied (x)",
        "Correction total / median / robust (log2)",
        "Candidates by estimator (median/robust/total)", "Top-10 signal share",
        "Ratio noise sigma (log2)", "Share above the fold gate", "Min normalized FC at q=0.05",
        "Unique peptides >= threshold", "CRAPome contaminants",
        "FC >= gate (measured or floored)", "Pass all gates", "Bait recovered",
        "Mixture fitted", "Model pi", "Model mu1", "Model mu0/sigma0",
        "Top candidate", "Top candidate FC", "Warnings",
    ]
    overview.append(headers)
    for index, result in enumerate(results, 1):
        proteins = result["proteins"]
        diagnostics = result["diagnostics"]
        model = result.get("saint", {}).get("model", {}) or {}
        candidates = [p for p in proteins if p["passes"]]
        # Same order as the report and the Candidates sheet. Taking the maximum score
        # returned an arbitrary row whenever no mixture fitted (every score is None).
        top = min(candidates, key=_candidate_order, default=None)
        overview.append([
            index,
            result["condition"],
            result["file"],
            result["bait_header"],
            result["control_header"],
            sum(1 for p in proteins if p["bait_abundance"] and p["bait_abundance"] > 0),
            sum(1 for p in proteins if p["control_abundance"] and p["control_abundance"] > 0),
            # measured in both channels; the floored rows are counted in their own column,
            # so this must not absorb them or "Paired" overstates the evidence
            sum(1 for p in proteins
                if p["fc_norm"] is not None and not p.get("fc_is_lower_bound")),
            diagnostics.get("n_control_imputed", 0),
            diagnostics.get("ctrl_q05"),
            result.get("noise", {}).get("n_lower_bound_gate", 0),
            diagnostics["bait_total"],
            diagnostics["control_total"],
            diagnostics["load_ratio_total"],
            diagnostics["correction_log2"],
            diagnostics["mode"],
            2 ** -diagnostics["correction_log2"],
            " / ".join(f"{diagnostics.get('correction_' + m):+.2f}" if diagnostics.get("correction_" + m) is not None else "n/a"
                       for m in ("total", "median", "robust")),
            " / ".join(str(result["mode_counts"].get(m, "n/a")) for m in ("median", "robust", "total")),
            diagnostics["concentration_top10"],
            result.get("noise", {}).get("sigma0"),
            result.get("noise", {}).get("share_above_gate"),
            result.get("noise", {}).get("fold_at_q05"),
            sum(1 for p in proteins if p["n_unique_peptides"] >= args.min_unique_peptides),
            sum(1 for p in proteins if p["crapome_flag"]),
            sum(1 for p in proteins if p["fc_norm"] is not None and p["fc_norm"] >= args.min_fold),
            len(candidates),
            "Detected" if any(p["is_bait"] for p in proteins) else "Bait recovery QC failure",
            "yes" if result.get("saint", {}).get("fitted") else
            f"no ({result.get('saint', {}).get('reason', 'n/a')})",
            model.get("pi"),
            model.get("mu1"),
            f"{model.get('mu0'):.3f}/{model.get('sigma0'):.3f}" if model else None,
            f"{top['gn']} ({top['accession']})" if top else None,
            # the fold change that put it on top (>= marks a censored ratio)
            _top_fc_label(top),
            " | ".join(diagnostics["warnings"]) if diagnostics["warnings"] else "",
        ])
    _style_header(overview)
    column_of = {name: index + 1 for index, name in enumerate(headers)}
    formats = {
        "Summed signal control/experimental ratio": "0.000",
        "Fold scale applied (x)": "0.000",  # factor each fold change was multiplied by
        "Load correction applied (log2)": "+0.000;-0.000",
        "Top-10 signal share": "0.0%",
        "Ratio noise sigma (log2)": "0.00",
        "Share above the fold gate": "0.0%",
        "Min normalized FC at q=0.05": "0.0",
        "Model pi": "0.000",
        # the value is a string like "63.3x" / ">=1750.9x"; a number format on text is
        # ignored by Excel, the key just has to match the header
        "Top candidate FC": "0.0000",
    }
    for cell_name, number_format in formats.items():
        index = column_of.get(cell_name)
        if not index:
            continue
        for row in range(2, overview.max_row + 1):
            overview.cell(row, index).number_format = number_format
    _autosize(overview, max_width=52)

    candidates_sheet = wb.create_sheet("Candidates")
    candidates_sheet.append(
        ["Condition", "Rank in condition", "In CRAPome table id"] + [name for name, _ in CANDIDATE_COLUMNS]
    )
    total = 0
    for result in results:
        rows = sorted((p for p in result["proteins"] if p["passes"]), key=_candidate_order)
        for rank, p in enumerate(rows, 1):
            candidates_sheet.append(
                [result["condition"], rank, p.get("accession")]
                + [_cell_value(p, key) for _, key in CANDIDATE_COLUMNS]
            )
        total += len(rows)
    _style_header(candidates_sheet)
    _format_columns(candidates_sheet, CANDIDATE_COLUMNS)
    _autosize(candidates_sheet, max_width=48)

    for result in results:
        title = _safe_sheet_title(result["condition"], wb.sheetnames)
        ws = wb.create_sheet(title)
        ws.append(["Condition", "GN", "Accession", "Description", "FC (raw)", "FC (EV-corrected)",
                   "log2FC (EV-corrected)", "FC (median-normalized)", "FC (total-normalized)",
                   "FC type", "SAINT-style score", "BFDR (approx)", "p (one-sided)", "q (BH)",
                   "Unique peptides",
                   "Peptides", "PSMs", "Coverage [%]", "Experimental abundance", "Control abundance",
                   "Control used", "CRAPome frequency", "CRAPome hits >=2SC", "CRAPome runs",
                   "Passes gates", "Reasons excluded", "Gate notes",
                   "Combined score", "Background flag", "Cross-sample QC"])
        rows = sorted(result["proteins"], key=lambda p: (not p["passes"], _candidate_order(p)))
        for p in rows:
            ws.append([
                result["condition"], p["gn"], p["accession"], p["description"], p["fc_raw"],
                p["fc_norm"], p["log2fc_norm"], p["fc_norm_median"], p["fc_norm_total"],
                p["fc_type"], p["saint_score"], p["saint_bfdr"], p["p_one_sided"], p["q_value"],
                p["n_unique_peptides"], p["n_peptides"], p["n_psms"], p["coverage"],
                p["bait_abundance"], p["control_abundance"],
                p.get("control_used"), p["crapome_frequency"],
                p["crapome_hits"], p["crapome_runs"],
                "yes" if p["passes"] else "no", "; ".join(p["pass_reasons"]),
                "; ".join(p.get("pass_notes") or []),
                p.get("qc_combined"), p.get("qc_bg_flag") or "", p.get("qc_cross") or "",
            ])
        _style_header(ws)
        numeric = {
            "FC (raw)", "FC (EV-corrected)", "log2FC (EV-corrected)",
            "FC (median-normalized)", "FC (total-normalized)", "Combined score",
            "SAINT-style score", "BFDR (approx)", "p (one-sided)", "q (BH)",
            "CRAPome frequency",
        }
        header_row = [cell.value for cell in ws[1]]
        for index, name in enumerate(header_row, 1):
            if name in numeric:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, index).number_format = "0.0000"
        _autosize(ws, max_width=44)

    methods = wb.create_sheet("Methods")
    rows = [
        ["Item", "Value", "Note"],
        ["Run date", date.today().isoformat(), ""],
        ["Conditions", len(results), ""],
        ["Fold-change gate", f">= {args.min_fold}x (normalized)",
         "Literature commonly uses >= 1.5x for potential interactors. The overview sheet also reports the "
         "ratio noise sigma and the share of proteins that clear this gate, because on single-run data that "
         "share is the gate's false-positive rate and is often 20-40%"],
        ["Unique peptide gate", f">= {args.min_unique_peptides} unique peptides",
         "PD reports this count experiment-wide, not per channel; a per-channel count needs a peptide-level export"],
        ["Load normalization", args.normalize,
         "Removes the bait/control scale difference before any fold change is read. Three estimators are "
         "reported because they disagree when a few proteins dominate the totals: total matches summed signal, "
         "median uses the median log2 ratio, robust uses the better-quantified half. The reference columns keep "
         "the median- and total-normalized fold change side by side."],
        ["CRAPome table", crapome_source or "not available", "Official control matrix, human, HEK293 subset preferred"],
        ["CRAPome contaminant rule",
         f">= {args.crapome_frequency:.0%} of controls with >= {args.crapome_min_sc} spectral counts",
         "Mirrors the CRAPome default filter; annotated, never auto-deleted"],
        ["SAINT-style score",
         "Posterior probability from a two-component Gaussian mixture",
         "Controls fix the null; the enrichment component is fitted per condition by EM"],
        ["SAINT-style score limits",
         "Not the official SAINTq; not a calibrated FDR",
         "SAINTq scores peptide/fragment intensities; protein-level aggregation underestimates the FDR, "
         "and one bait run plus one shared control gives no replicates"],
        ["BFDR (approx)", "Newton et al. 2004 estimator over the posterior probabilities",
         "Approximate under this design; treat as a ranking aid"],
        ["p / q (calibrated)", "One-sided p from a null estimated on the negative half of the "
         "ratio distribution; q is Benjamini-Hochberg over all quantified prey",
         "This is the multiple-testing-aware companion: it needs only that the null be roughly "
         "symmetric, not that the enriched preys form an identifiable second cluster"],
        ["Control not detected",
         "5th percentile of the positive control values in the same workbook",
         "A prey the control never detected has a censored ratio, not a missing one: the ratio is at least "
         "bait/floor and can only be larger. Those fold changes are labelled 'FC_lower_bound; Control_missing' "
         "(the same convention as the batch QC skill) and are excluded from the background spread and from "
         "every p/q value, so they cannot inflate the null or take a test slot in the BH correction"],
        ["Paired (measured)", "Quantified in both channels",
         "Prey the control never detected are NOT counted here; they are counted in 'Control not "
         "detected (floored)' and their fold changes are lower bounds"],
        ["Never modified", "Source workbooks are read-only", "All output written under the run folder"],
    ]
    for row in rows:
        methods.append(row)
    _style_header(methods)
    _autosize(methods, max_width=80)
    wb.save(path)
    return total


def _cell_value(protein, key):
    value = protein.get(key)
    if isinstance(value, float):
        return round(value, 6)
    if isinstance(value, list):
        return "; ".join(value)
    return value


def _format_columns(ws, columns):
    for row in ws.iter_rows(min_row=2):
        for index, (_, key) in enumerate(columns, start=4):
            if key in ("fc_norm", "fc_raw", "log2fc_norm", "saint_score", "saint_bfdr",
                       "p_one_sided", "q_value", "crapome_frequency", "control_used"):
                row[index - 1].number_format = "0.0000"


def _style_header(ws):
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"


def _autosize(ws, max_width=40):
    for column in range(1, ws.max_column + 1):
        letter = get_column_letter(column)
        longest = 0
        for row in range(1, min(ws.max_row, 200) + 1):
            value = ws.cell(row, column).value
            if value is None:
                continue
            longest = max(longest, len(str(value)))
        ws.column_dimensions[letter].width = min(max(10, longest + 2), max_width)


def _safe_sheet_title(name, existing):
    title = re.sub(r"[\[\]:*?/\\]", "_", name)[:31] or "condition"
    base, counter = title, 2
    while title in existing:
        suffix = f"_{counter}"
        title = base[: 31 - len(suffix)] + suffix
        counter += 1
    return title


# --------------------------------------------------------------------------
# report payload
# --------------------------------------------------------------------------
POINT_FIELDS = [
    "gn", "acc", "name", "bait", "ctrl", "fc", "fc_raw", "score", "bfdr", "q", "up",
    "pep", "psm", "cov", "cfreq", "chits", "cruns", "crapome", "bait_hit",
    "pass", "why", "cuse", "lb",
    # batch-QC 追加的三列（只有跑过 bioid-batch-qc 的输入才有；缺省为空）
    "qc_combined", "qc_bg_flag", "qc_cross",
    # 线性 FC 与 log2 同时给前端，表里两列并列显示
    "log2fc",
]


def build_payload(results, args, crapome_info, crapome_source, output_dir, annotations=None):
    conditions = []
    for result in results:
        proteins = result["proteins"]
        points = []
        for p in proteins:
            points.append([
                p["gn"] or p["accession"],
                p["accession"],
                clean_description(p.get("description")),
                p["bait_abundance"],
                p["control_abundance"],
                p["fc_norm"],
                p["fc_raw"],
                p["saint_score"],
                p["saint_bfdr"],
                p["q_value"],
                p["n_unique_peptides"],
                p["n_peptides"],
                p["n_psms"],
                p["coverage"],
                p["crapome_frequency"],
                p["crapome_hits"],
                p["crapome_runs"],
                1 if p["crapome_flag"] else 0,
                1 if p["is_bait"] else 0,
                1 if p["passes"] else 0,
                "; ".join(p["pass_reasons"]),
                p.get("control_used"),
                1 if p.get("fc_is_lower_bound") else 0,
                p.get("qc_combined"),
                p.get("qc_bg_flag") or "",
                p.get("qc_cross") or "",
                p.get("log2fc_norm"),
            ])
        conditions.append({
            "condition": result["condition"],
            "file": result["file"],
            "expChannel": result["bait_channel"],
            "ctrlChannel": result["control_channel"],
            "expHeader": result["bait_header"],
            "ctrlHeader": result["control_header"],
            "diagnostics": result["diagnostics"],
            "modeCounts": result.get("mode_counts", {}),
            "noise": result.get("noise", {}),
            "saint": {k: v for k, v in result["saint"].items() if k != "model"},
            "model": result["saint"].get("model"),
            "points": points,
        })
    def round_floats(node):
        if isinstance(node, float):
            return round(node, 4)
        if isinstance(node, dict):
            return {key: round_floats(value) for key, value in node.items()}
        if isinstance(node, list):
            return [round_floats(value) for value in node]
        return node

    payload = round_floats({
        "generated": date.today().isoformat(),
        "folder": str(output_dir),
        "thresholds": {
            "minFold": args.min_fold,
            "minUniquePeptides": args.min_unique_peptides,
            "crapomeFrequency": args.crapome_frequency,
            "crapomeMinSc": args.crapome_min_sc,
            "crapomeCell": args.crapome_cell,
            "normalize": args.normalize,
        },
        "crapome": {
            "applied": crapome_info.get("applied", False),
            "source": crapome_source,
        },
        "pointFields": POINT_FIELDS,
        # 报告据此决定提示「未接入 QC」还是画背景标记
        "qcAttached": any(
            (p.get("qc_bg_flag") or p.get("qc_cross") or p.get("qc_combined") is not None)
            for result in results for p in result["proteins"]
        ),
        "annotations": annotations or {},
        "colors": REPORT_COLORS,
        "conditions": conditions,
    })
    return payload


def write_report(payload, path, plotly_mode):
    template = REPORT_TEMPLATE.read_text()
    if plotly_mode == "inline":
        libs = []
        for lib in (PLOTLY_JS, JSPDF_JS, SVG2PDF_JS):
            if not lib.exists():
                raise FileNotFoundError(
                    f"{lib} is missing; re-run with --plotly cdn or restore the vendored library"
                )
            libs.append(f"<script>{lib.read_text()}</script>")
        if NOTO_TTF.exists():
            import base64 as _b64
            libs.append('<script>window.NOTO_FONT_B64="'
                        + _b64.b64encode(NOTO_TTF.read_bytes()).decode() + '";</script>')
        lib_tag = "".join(libs)
    else:
        lib_tag = (
            '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js" charset="utf-8"></script>'
            '<script src="https://unpkg.com/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>'
            '<script src="https://unpkg.com/svg2pdf.js@2.2.3/dist/svg2pdf.umd.min.js"></script>'
        )
    html = (
        template.replace("<!--PLOTLY_LIB-->", lib_tag)
        .replace("/*DATA*/null", json.dumps(payload, allow_nan=False, default=str))
    )
    path.write_text(html, encoding="utf-8")
    return path.stat().st_size


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def parse_args(argv):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("folder", help="folder holding one .xlsx workbook per bait condition")
    parser.add_argument("--out", default=None, help="output folder (default: <folder>/bioid_analysis)")
    parser.add_argument("--list-baits", action="store_true",
                        help="只列出每个文件推断出的诱饵（附证据）供人工确认，不做分析")
    parser.add_argument("--bait-alias", default=None,
                        help="诱饵别名表：每行 <标签> <数据库符号>（如 FAM214A ATOSA），只用于排除诱饵，不改条件名")
    parser.add_argument("--bait-map", default=None,
                        help="人工诱饵映射表：每行 <工作簿文件名> <基因名>，覆盖文件名推断"
                             "（匹配文件名或去掉 _processed/_flagged 后缀的文件名）")
    parser.add_argument("--from-qc", action="store_true",
                        help="read the batch-QC output (flagged/*_flagged.xlsx) instead of the folder's own workbooks")
    parser.add_argument("--no-qc", action="store_true",
                        help="ignore a batch-QC flagged/ folder even when it exists")
    parser.add_argument("--allow-no-qc", action="store_true",
                        help="没有 batch-QC 输出时仍然分析（QC 三列留空）；默认拒绝，"
                             "避免把缺 QC 的结果当成完整结果")
    parser.add_argument("--allow-partial-qc", action="store_true",
                        help="batch-QC 只覆盖部分工作簿时，仍然只分析已 QC 的那部分")
    parser.add_argument("--sheet", default="Proteins", help="sheet to read (default: Proteins)")
    parser.add_argument("--min-fold", type=float, default=DEFAULTS["min_fold"],
                        help="normalized fold-change gate (default 1.5)")
    parser.add_argument("--min-unique-peptides", type=int, default=DEFAULTS["min_unique_peptides"],
                        help="unique peptide gate in the experimental channel (default 2)")
    parser.add_argument("--normalize", choices=["median", "robust", "total", "none"],
                        default=DEFAULTS["normalize"],
                        help="bait/control load correction: median (default, robust to a few dominant "
                             "proteins), robust (median over the better-quantified half), total "
                             "(match summed signal; use when channels are known to differ in amount), none")
    parser.add_argument("--crapome-table", default=str(CRAPOME_TABLE), help="CRAPome frequency table")
    parser.add_argument("--annotations", default=str(ANNOTATION_TABLE),
                        help="UniProt annotation cache for the hover cards; built by "
                             "scripts/fetch_uniprot_annotations.py")
    parser.add_argument("--crapome-frequency", type=float, default=DEFAULTS["crapome_frequency"],
                        help="control frequency above which a prey is a contaminant (default 0.20)")
    parser.add_argument("--crapome-min-sc", type=int, default=DEFAULTS["crapome_min_sc"],
                        help="spectral counts required in a control run (default 2)")
    parser.add_argument("--crapome-cell", default=DEFAULTS["crapome_cell"],
                        help="control cell-line subset; use 'all' for every human control")
    parser.add_argument("--plotly", choices=["inline", "cdn"], default=DEFAULTS["plotly_mode"],
                        help="embed plotly.js (offline, default) or reference the CDN")
    parser.add_argument("--smoke", action="store_true",
                        help="analyse only the first condition: a fast iteration loop while "
                             "changing the report template (seconds instead of minutes). "
                             "Not for results -- the numbers it prints describe one condition only")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv if argv is not None else sys.argv[1:])
    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        raise SystemExit(f"not a folder: {folder}")
    if args.out:
        output_dir = Path(args.out).expanduser().resolve()
    elif args.smoke:
        # smoke 是模板迭代用的，单独一个目录，避免把正式 summary/report 覆盖成单条件版本
        output_dir = folder / "bioid_analysis_smoke"
    else:
        output_dir = folder / "bioid_analysis"

    files, source_note = discover_workbooks(folder, args.from_qc, args.no_qc,
                                            args.allow_no_qc, args.list_baits)
    if source_note:
        print(f"input: {source_note}")
    else:
        print(f"input: {folder} 的原始工作簿（未接入 batch-QC）")
    if (args.no_qc or not source_note) and not args.list_baits:
        why = "--no-qc 按你的要求跳过" if args.no_qc else "没有找到 batch-QC 输出"
        print(f"WARNING {why}：候选表不会有 Combined score / Background flag / "
              "Cross-sample QC 三列；需要它们请先跑 bioid-batch-qc，或用 bioid_run.py 一条命令跑完")

    gap = (qc_coverage_gap(folder, files)
           if source_note and "batch-QC" in source_note and not args.list_baits else [])
    if gap:
        message = ("batch-QC 输出缺这些条件（原始工作簿里有）：" + ", ".join(gap)
                   + "；请重跑 bioid-batch-qc，或加 --allow-partial-qc 只分析已 QC 的那部分")
        if not args.allow_partial_qc:
            raise SystemExit(message)
        print(f"WARNING {message}")
    if not files:
        raise SystemExit(f"no .xlsx workbooks in {folder}")
    if args.smoke:
        files = files[:1]
        print(f"smoke mode: analysing only {files[0].name} — for checking the report, "
              "not for results; re-run without --smoke for the batch")

    bait_map = read_bait_map(args.bait_map) if args.bait_map else {}
    bait_aliases = read_bait_aliases(args.bait_alias)
    if bait_aliases:
        alias_preview = ", ".join(
            key + "->" + "/".join(sorted(value))
            for key, value in list(bait_aliases.items())[:6])
        print(f"诱饵别名: {len(bait_aliases)} 条（{alias_preview}）")
    if bait_map:
        print(f"人工诱饵映射: {len(bait_map)} 条（{', '.join(f'{k}->{v}' for k, v in list(bait_map.items())[:6])}"
              + (" …" if len(bait_map) > 6 else "") + "）")

    if args.list_baits:
        # 人工确认用：只读、只打印，不写任何输出。诱饵蛋白通常是诱饵通道里最亮的蛋白之一，
        # 因此同时给出该通道丰度最高的几个蛋白作为判断依据。
        print("\n每份工作簿推断出的诱饵（请人工核对后再正式跑）：\n")
        for path in files:
            try:
                info = read_workbook(path, args.sheet, bait_map)
            except Exception as error:
                print(f"  {path.name}: 读取失败 {type(error).__name__}: {error}")
                continue
            label = info["bait"]
            overridden = bait_override(bait_map, path) is not None
            hit = next((p for p in info["proteins"]
                        if p["gn"] and p["gn"].upper() == str(label).upper()), None)
            recovered = (f"检出：{hit['bait_abundance']:.3g}，"
                         f"{hit['n_unique_peptides']} unique 肽，覆盖 {hit['coverage']}%"
                         if hit and hit.get("bait_abundance") else "未检出（不能排除库/物种问题）")
            print(f"  {path.name}")
            print(f"      诱饵标签: {label}" + ("（来自 --bait-map 人工指定）" if overridden else "（由文件名推断）"))
            print(f"      通道: 实验 {info['bait_channel']} ← 对照 {info['control_channel']}")
            print(f"      该标签的检出情况: {recovered}")
            # 证据：诱饵应当是"诱饵通道专属、且足够亮"的蛋白。生物素化羧化酶总是最亮
            # （内源生物素），对判断诱饵没有信息量，所以先按"对照未检出或 raw FC ≥ 5"过滤。
            def is_evidence(protein):
                if not protein.get("bait_abundance"):
                    return False
                ctrl = protein.get("control_abundance")
                return (not ctrl) or (protein["bait_abundance"] / ctrl >= 5)

            top = sorted((p for p in info["proteins"] if is_evidence(p)),
                         key=lambda p: -(p["bait_abundance"] or 0))[:5]
            print("      候选证据（诱饵通道专属或 ≥5× 富集，按诱饵丰度排序）:")
            for rank, protein in enumerate(top, 1):
                ctrl = protein.get("control_abundance")
                relation = ("对照未检出" if not ctrl
                            else f"{protein['bait_abundance'] / ctrl:.0f}× vs 对照")
                print(f"        {rank}. {protein['gn'] or protein['accession']:<12} "
                      f"诱饵丰度 {protein['bait_abundance']:.3g} · {relation}"
                      f" · {protein.get('n_unique_peptides')} unique 肽")
            print()
        print("确认方式：诱饵正确则直接跑；需要改就用 --bait-map（每行 <文件名或条件名> <基因>），"
              "或把工作簿重命名为 <序号><基因>_processed.xlsx。")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    crapome = load_crapome(args.crapome_table, args.crapome_cell)
    crapome_source = None
    if crapome is not None:
        crapome_source = f"{Path(args.crapome_table).name} ({args.crapome_cell} controls)"

    results, skipped, pending = [], [], []
    for path in files:
        try:
            condition = read_workbook(path, args.sheet, bait_map)
        except Exception as exc:  # keep going: one bad export must not stop the batch
            skipped.append(f"{path.name}: {exc}")
            continue
        if not condition["proteins"]:
            skipped.append(f"{path.name}: no protein rows")
            continue
        pending.append(condition)
    try:
        assign_conditions(pending)
    except ValueError as exc:
        raise SystemExit(str(exc))
    for condition in pending:
        diagnostics = normalize_loads(condition["proteins"], args.normalize, DEFAULTS["load_warning_log2"])
        impute_missing_controls(condition["proteins"], diagnostics)
        add_fold_changes(condition["proteins"], diagnostics)
        crapome_info = annotate_crapome(
            condition["proteins"], crapome, args.crapome_cell,
            args.crapome_frequency, args.crapome_min_sc,
        )
        tag_bait(condition["proteins"], condition["family"],
                 bait_aliases.get(condition["family"], set()))
        if not any(p.get("is_bait") for p in condition["proteins"]):
            top_hits = sorted(
                (p for p in condition["proteins"] if p["fc_norm"] is not None),
                key=lambda p: -p["fc_norm"])[:3]
            hint = ", ".join(
                f"{p['gn'] or p['accession']}"
                f"({'≥' if p.get('fc_is_lower_bound') else ''}{p['fc_norm']:.0f}x)"
                for p in top_hits)
            diagnostics["warnings"].append(
                f"bait '{condition['family']}' was not found among this workbook's proteins, so it "
                f"was neither excluded nor marked; the most enriched proteins here are {hint or 'n/a'}"
                " — check the filename inference with --list-baits or pass --bait-map")
        saint = saint_scores(condition["proteins"], diagnostics)
        apply_gates(condition["proteins"], args, condition["condition"])
        condition["mode_counts"] = candidate_counts_by_mode(
            condition["proteins"], diagnostics, args.min_fold, args.min_unique_peptides
        )
        condition["noise"] = noise_summary(
            condition["proteins"], diagnostics, args.min_fold, args.min_unique_peptides
        )
        condition["diagnostics"] = diagnostics
        condition["crapome_info"] = crapome_info
        condition["saint"] = saint
        results.append(condition)

    if not results:
        raise SystemExit("no readable workbooks; " + ("; ".join(skipped) if skipped else ""))

    excel_path = output_dir / "BioID_analysis_summary.xlsx"
    n_candidates = write_excel(results, excel_path, args, {}, crapome_source)

    annotations = build_annotation_payload(results, load_annotations(args.annotations))
    payload = build_payload(results, args, {"applied": crapome is not None}, crapome_source,
                            output_dir, annotations)
    report_path = output_dir / "BioID_analysis_report.html"
    report_size = write_report(payload, report_path, args.plotly)
    (output_dir / "analysis_payload.json").write_text(
        json.dumps(payload, allow_nan=False, default=str), encoding="utf-8"
    )

    print(f"conditions analysed: {len(results)}")
    for skipped_file in skipped:
        print(f"SKIPPED  {skipped_file}")
    print(f"summary workbook: {excel_path}")
    print(f"interactive report: {report_path} ({report_size / 1e6:.1f} MB, plotly={args.plotly})")
    print(f"protein annotations: {len(annotations)} accession(s) with a name/alias/function entry")
    print(
        f"gates: >={args.min_fold}x normalized, >={args.min_unique_peptides} unique peptides"
        f"{'' if crapome is not None else '; CRAPome table NOT available (annotation skipped)'}"
    )
    families = {}
    for result in results:
        families.setdefault(result["family"], []).append(result["condition"])
    merged = {f: c for f, c in families.items() if len(c) > 1}
    if merged:
        print("multi-condition families (scored separately, one family for bait identity):")
        for family in sorted(merged):
            print(f"  {family}: {', '.join(sorted(merged[family]))}")
    print(f"candidates passing all gates: {n_candidates}")
    for result in results:
        diagnostics = result["diagnostics"]
        top = min((p for p in result["proteins"] if p["passes"]),
                  key=_candidate_order, default=None)
        counts = result["mode_counts"]
        line = (
            f"  {result['condition']:22} load x{diagnostics['load_ratio_total'] or float('nan'):.2f} "
            f"corr {diagnostics['correction_log2']:+.2f} "
            f"pass={sum(1 for p in result['proteins'] if p['passes']):4d} "
            f"(median/robust/total: {counts.get('median','-')}/{counts.get('robust','-')}/{counts.get('total','-')}) "
            f"crapome={sum(1 for p in result['proteins'] if p['crapome_flag']):4d} "
            f"top={top['gn'] if top else '-'}"
        )
        print(line)
        for warning in diagnostics["warnings"]:
            print(f"      WARNING {warning}")
        if not result["saint"].get("fitted"):
            print(f"      WARNING SAINT-style score not fitted: {result['saint'].get('reason')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
