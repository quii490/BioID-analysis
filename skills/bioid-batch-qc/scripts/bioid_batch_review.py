#!/usr/bin/env python3
"""Batch-level BioID scoring, background flagging, and cross-sample QC.

This script never overwrites the source workbooks. It creates flagged copies
and a separate batch QC workbook.
"""

from __future__ import annotations

import argparse
import atexit
import math
import os
import re
import statistics
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from copy import copy
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# 目录、输出路径与条件元数据都在 main() 里按传入目录推断；这里不再内置任何批次的数据。
ROOT = OUT_DIR = QC_PATH = FILE_META = None

ORANGE = "F4B183"
YELLOW = "FFF2CC"
PINK = "F4CCCC"
BLUE_HEADER = "A7CDF0"
LIGHT_BLUE = "DDEBF7"
GREEN = "C6EFCE"
NAVY = "1F4E78"
WHITE = "FFFFFF"
LIGHT_BORDER = "B4C6E7"
GRAY = "E7E6E6"

# Canonical bait names can have legacy aliases in protein annotations. Add
# batch-specific aliases here when a construct is known under another GN.
# They are used for bait recovery and transfer QC, while original GN rows are
# preserved unchanged in the output workbooks.
KNOWN_BAIT_ALIASES = {
    "ARB2BP": {"ARB2BP", "FAM172BP"},
}


def to_num(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def fmt(value, spec=".4f"):
    """Render a console-summary number, or '-' when it is None."""
    return "-" if value is None else format(value, spec)


def gene_metric(gene_metrics, name):
    """按基因名取 gene_metrics：先精确，再大小写不敏感（构造名/别名大小写常不同）。"""
    if not name:
        return None
    metric = gene_metrics.get(name)
    if metric is not None:
        return metric
    upper = str(name).upper()
    for gn, gm in gene_metrics.items():
        if str(gn).upper() == upper:
            return gm
    return None


def median(values):
    values = [x for x in values if x is not None and math.isfinite(x)]
    return statistics.median(values) if values else None


def q05(values):
    values = [x for x in values if x is not None and x > 0]
    return float(np.quantile(values, 0.05)) if values else None


def high_percentiles(values_by_id):
    """Return 0..1 percentile where 1 is the largest value, with average ties."""
    valid = [(key, value) for key, value in values_by_id.items() if value is not None]
    if not valid:
        return {}
    valid.sort(key=lambda item: item[1], reverse=True)
    n = len(valid)
    output = {}
    i = 0
    while i < n:
        j = i + 1
        while j < n and valid[j][1] == valid[i][1]:
            j += 1
        average_rank = ((i + 1) + j) / 2
        percentile = 1.0 if n == 1 else 1 - (average_rank - 1) / (n - 1)
        for k in range(i, j):
            output[valid[k][0]] = percentile
        i = j
    return output


def top_rank_fraction(values_by_id):
    """Return 1/n..1 rank fraction where smaller means closer to the top."""
    valid = [(key, value) for key, value in values_by_id.items() if value is not None]
    valid.sort(key=lambda item: item[1], reverse=True)
    n = len(valid)
    return {key: rank / n for rank, (key, _) in enumerate(valid, 1)} if n else {}


def average_ranks(values):
    order = sorted(range(len(values)), key=lambda i: values[i])
    ranks = [0.0] * len(values)
    i = 0
    while i < len(order):
        j = i + 1
        while j < len(order) and values[order[j]] == values[order[i]]:
            j += 1
        avg = ((i + 1) + j) / 2
        for k in range(i, j):
            ranks[order[k]] = avg
        i = j
    return ranks


def pearson(x, y):
    if len(x) < 3:
        return None
    mx, my = statistics.mean(x), statistics.mean(y)
    sx = sum((v - mx) ** 2 for v in x)
    sy = sum((v - my) ** 2 for v in y)
    if sx == 0 or sy == 0:
        return None
    return sum((a - mx) * (b - my) for a, b in zip(x, y)) / math.sqrt(sx * sy)


def spearman(x, y):
    return pearson(average_ranks(x), average_ranks(y))


# --------------------------------------------------------------------------
# 与 openxlsx 的互操作
# --------------------------------------------------------------------------
# openxlsx 写出的工作簿有两处会让 openpyxl 读不了或读错：
#   ① 工作表关系指向从不写出的 drawing 部件 -> openpyxl 抛
#      KeyError: "There is no item named 'xl/drawings/drawing1.xml'"；
#   ② 工作表 dimension 被写成 "A1" -> read_only 模式只看到一列，
#      于是「至少两个 Abundance 列」的检查必然失败、整批中止。
# 这里在临时副本上修掉这两处；文件本身正常时直接用原文件，不拷贝。
_REPAIR_DIR = None
_ANNOUNCED_REPAIRS = set()
SKIPPED_SOURCES = []


def _col_to_int(letters):
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value


def _int_to_col(value):
    letters = ""
    while value:
        value, rem = divmod(value - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters or "A"


def _dangling_rels(members):
    """返回 (rels 名, 悬空的 Target 列表)，只针对包内部关系。"""
    found = []
    for name in list(members):
        if not (name.endswith(".rels") and "/_rels/" in name):
            continue
        body = members[name].decode("utf-8", "ignore")
        base = name.replace("/_rels/", "/").rsplit("/", 1)[0] + "/"
        missing = []
        for target in re.findall(r'Target="([^"]+)"', body):
            if target.startswith(("http://", "https://", "/")):
                continue
            if os.path.normpath(base + target) not in members:
                missing.append(target)
        if missing:
            found.append((name, missing))
    return found


def _degenerate_dimension(members):
    """dimension 缺失或只覆盖 A1（但工作表里其实有更多列）的工作表部件。"""
    bad = []
    for name in list(members):
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
            continue
        text = members[name].decode("utf-8", "ignore")
        match = re.search(r'<dimension ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"', text)
        if match and match.group(3) is not None:
            continue
        bad.append(name)
    return bad


def repair_openxlsx_package(path):
    """需要时返回修好的副本路径，否则返回原路径。"""
    global _REPAIR_DIR
    try:
        with zipfile.ZipFile(path) as archive:
            members = {name: archive.read(name) for name in archive.namelist()}
    except zipfile.BadZipFile:
        return path
    dangling = _dangling_rels(members)
    bad_dims = _degenerate_dimension(members)
    if not dangling and not bad_dims:
        return path

    for name, missing in dangling:
        body = members[name].decode("utf-8", "ignore")
        for target in missing:
            body = re.sub(r'<Relationship [^>]*Target="%s"[^>]*/>' % re.escape(target), "", body)
        members[name] = body.encode("utf-8")
    for name in bad_dims:
        text = members[name].decode("utf-8", "ignore")
        refs = re.findall(r'<c r="([A-Z]+)(\d+)"', text)
        if not refs:
            continue
        max_col = max(_col_to_int(col) for col, _ in refs)
        max_row = max(int(row) for _, row in refs)
        element = '<dimension ref="A1:%s%d"/>' % (_int_to_col(max_col), max_row)
        if re.search(r"<dimension[^>]*/>", text):
            text = re.sub(r"<dimension[^>]*/>", element, text, count=1)
        else:
            text = text.replace("<sheetViews>", element + "<sheetViews>", 1)
        members[name] = text.encode("utf-8")

    if _REPAIR_DIR is None:
        _REPAIR_DIR = tempfile.mkdtemp(prefix="bioid_repair_")
        atexit.register(lambda: __import__("shutil").rmtree(_REPAIR_DIR, ignore_errors=True))
    out = os.path.join(_REPAIR_DIR, os.path.basename(path))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    # 同一个文件会被多个阶段打开，只提示一次，避免刷屏。
    announced = os.path.basename(path)
    if announced not in _ANNOUNCED_REPAIRS:
        _ANNOUNCED_REPAIRS.add(announced)
        print(f"  注意: {announced} 是 openxlsx 写出的缺陷包（悬空关系/维度写成 A1），已在临时副本上修复后读取")
    return out


def load_workbook_compat(path, **kwargs):
    """读 openxlsx 写出的工作簿：必要时在修复过的副本上打开。"""
    return load_workbook(repair_openxlsx_package(path), **kwargs)


def headers(ws):
    return [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]


def find_abundance_columns(header):
    idx = [
        i
        for i, value in enumerate(header)
        if isinstance(value, str)
        and re.match(r"abundance[:\s]", value.strip().lower())
        and "normalized" not in value.lower()
    ]
    if len(idx) < 2:
        raise ValueError("At least two Abundance columns are required")
    return idx[0] + 1, idx[1] + 1


def record_key(accession, description, gn):
    return (str(accession or ""), str(description or ""), str(gn or "").strip())


def inspect_source(path, meta):
    wb = load_workbook_compat(path, data_only=False, read_only=False)
    ws = wb["Proteins"]
    hdr = headers(ws)
    exp_col, ctrl_col = find_abundance_columns(hdr)
    idx = {name: i + 1 for i, name in enumerate(hdr)}
    gn_col = idx["GN"]
    acc_col = idx.get("Accession")
    desc_col = next(
        (i + 1 for i, x in enumerate(hdr) if isinstance(x, str) and "Description" in x),
        None,
    )

    rows = []
    ctrl_positive = []
    for row in range(2, ws.max_row + 1):
        gn = str(ws.cell(row, gn_col).value or "").strip()
        exp = to_num(ws.cell(row, exp_col).value)
        ctrl = to_num(ws.cell(row, ctrl_col).value)
        accession = ws.cell(row, acc_col).value if acc_col else ""
        description = ws.cell(row, desc_col).value if desc_col else ""
        if ctrl is not None and ctrl > 0:
            ctrl_positive.append(ctrl)
        rows.append(
            {
                "row": row,
                "key": record_key(accession, description, gn),
                "gn": gn,
                "accession": accession,
                "description": description,
                "exp": exp,
                "ctrl": ctrl,
            }
        )

    ctrl_q05 = q05(ctrl_positive)
    measured_fc = []
    for item in rows:
        exp, ctrl = item["exp"], item["ctrl"]
        if exp is not None and exp > 0 and ctrl is not None and ctrl > 0:
            fc = math.log2(exp / ctrl)
            item["raw_fc"] = fc
            item["fc_type"] = "Measured"
            measured_fc.append(fc)
        elif exp is not None and exp > 0 and ctrl_q05 and ctrl_q05 > 0:
            item["raw_fc"] = math.log2(exp / ctrl_q05)
            item["fc_type"] = "FC_lower_bound; Control_missing"
        else:
            item["raw_fc"] = None
            item["fc_type"] = ""

    shift = median(measured_fc) or 0.0
    visible = [item for item in rows if item["exp"] is not None and item["exp"] > 0]
    abundance_pct = high_percentiles({i: item["exp"] for i, item in enumerate(visible)})
    enrichment_pct = high_percentiles(
        {
            i: item["raw_fc"] - shift
            for i, item in enumerate(visible)
            if item["raw_fc"] is not None
        }
    )

    for i, item in enumerate(visible):
        item["abundance_pct"] = abundance_pct[i]
        if item["raw_fc"] is None:
            # 对照通道整列没有正值（ctrl_q05 为 None）：没有可算的富集值，留空而不是
            # 让 None 参与减法/sqrt；整本工作簿随后会在 main() 里跳过并告警。
            item["norm_fc"] = None
            item["enrichment_pct"] = None
            item["combined"] = None
            continue
        item["norm_fc"] = item["raw_fc"] - shift
        item["enrichment_pct"] = enrichment_pct[i]
        item["combined"] = math.sqrt(abundance_pct[i] * enrichment_pct[i])

    gene_metrics = {}
    for item in visible:
        gn = item["gn"]
        if not gn:
            continue
        gm = gene_metrics.setdefault(
            gn,
            {
                "combined": None,
                "enrichment_pct": None,
                "abundance_pct": None,
                "max_exp": None,
                "norm_fc": None,
                "accessions": set(),
            },
        )
        for key in ("combined", "enrichment_pct", "abundance_pct", "max_exp", "norm_fc"):
            source = item["exp"] if key == "max_exp" else item[key]
            if source is not None and (gm[key] is None or source > gm[key]):
                gm[key] = source
        if item["accession"]:
            gm["accessions"].add(str(item["accession"]))

    combined_rank = top_rank_fraction(
        {gn: gm["combined"] for gn, gm in gene_metrics.items()}
    )
    enrichment_rank = top_rank_fraction(
        {gn: gm["enrichment_pct"] for gn, gm in gene_metrics.items()}
    )
    for gn, gm in gene_metrics.items():
        gm["combined_rank_pct"] = combined_rank.get(gn)
        gm["enrichment_rank_pct"] = enrichment_rank.get(gn)

    row_metrics = {}
    for item in rows:
        if item["exp"] is not None and item["exp"] > 0:
            row_metrics[item["key"]] = item

    exp_detected = sum(1 for x in rows if x["exp"] is not None and x["exp"] > 0)
    ctrl_detected = sum(1 for x in rows if x["ctrl"] is not None and x["ctrl"] > 0)
    shared = sum(
        1
        for x in rows
        if x["exp"] is not None
        and x["exp"] > 0
        and x["ctrl"] is not None
        and x["ctrl"] > 0
    )
    control_missing = exp_detected - shared
    bait_metric = next(
        (
            metric
            for metric in (gene_metric(gene_metrics, alias)
                           for alias in meta.get("bait_aliases", [meta["bait_gene"]]))
            if metric is not None
        ),
        None,
    )

    return {
        **meta,
        "file": path.name,
        "path": path,
        "exp_header": hdr[exp_col - 1],
        "ctrl_header": hdr[ctrl_col - 1],
        "exp_col": exp_col,
        "ctrl_col": ctrl_col,
        "ctrl_q05": ctrl_q05,
        "median_shift": shift,
        "rows": rows,
        "row_metrics": row_metrics,
        "gene_metrics": gene_metrics,
        "exp_detected": exp_detected,
        "ctrl_detected": ctrl_detected,
        "shared": shared,
        "control_missing": control_missing,
        "bait_metric": bait_metric,
    }


def read_exempt_genes(spec):
    """读用户指定的「豁免基因」——不计入批次背景旗。

    spec 可以是文件路径（每行一个符号，# 开头为注释，行内只取第一列）或逗号/分号分隔的符号串。
    用途是两类**结构性复现**：共用诱饵自身，以及已知的真实共同邻居（它们本来就会在 ≥4 个 family
    里同时排进前百分之几）。豁免只清空背景旗：规则结果（rule_flag）与条件/家族命中计数照常写进
    工作簿，豁免名单与其原因也写进每个工作簿的 QC_Methods 与 QC 工作簿的 Methods，所以不是静默的。
    """
    if not spec:
        return set()
    path = Path(str(spec)).expanduser()
    text = path.read_text(encoding="utf-8") if path.exists() else str(spec)
    genes = set()
    for line in text.replace(";", ",").replace("\t", ",").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        for token in line.split(","):
            token = token.strip()
            if token:
                genes.add(token.upper())
    return genes


def calculate_batch_flags(conditions, exempt=None):
    exempt = exempt or set()
    condition_detect = Counter()
    condition_top2 = Counter()
    condition_top5 = Counter()
    condition_enrich5 = Counter()
    for condition in conditions:
        for gn, gm in condition["gene_metrics"].items():
            condition_detect[gn] += 1
            if gm["combined_rank_pct"] <= 0.02:
                condition_top2[gn] += 1
            if gm["combined_rank_pct"] <= 0.05:
                condition_top5[gn] += 1
            if gm["enrichment_rank_pct"] <= 0.05:
                condition_enrich5[gn] += 1

    families = defaultdict(list)
    for condition in conditions:
        families[condition["family"]].append(condition)

    family_metrics = {}
    for family, members in families.items():
        genes = set().union(*(set(x["gene_metrics"]) for x in members))
        combined = {}
        enrichment = {}
        detected = set()
        for gn in genes:
            vals = [
                x["gene_metrics"][gn]["combined"]
                for x in members
                if gn in x["gene_metrics"]
            ]
            evals = [
                x["gene_metrics"][gn]["enrichment_pct"]
                for x in members
                if gn in x["gene_metrics"]
            ]
            if vals:
                combined[gn] = max(vals)
                enrichment[gn] = max(evals)
                detected.add(gn)
        cr = top_rank_fraction(combined)
        er = top_rank_fraction(enrichment)
        family_metrics[family] = {
            gn: {
                "combined_rank_pct": cr[gn],
                "enrichment_rank_pct": er[gn],
            }
            for gn in detected
        }

    family_detect = Counter()
    family_top2 = Counter()
    family_top5 = Counter()
    family_enrich5 = Counter()
    for metrics in family_metrics.values():
        for gn, gm in metrics.items():
            family_detect[gn] += 1
            if gm["combined_rank_pct"] <= 0.02:
                family_top2[gn] += 1
            if gm["combined_rank_pct"] <= 0.05:
                family_top5[gn] += 1
            if gm["enrichment_rank_pct"] <= 0.05:
                family_enrich5[gn] += 1

    all_genes = set().union(*(set(c["gene_metrics"]) for c in conditions))
    flags = {}
    for gn in all_genes:
        if family_top2[gn] >= 4:
            flag = "High background suspect"
        elif family_top5[gn] >= 4 or family_enrich5[gn] >= 5:
            flag = "Review background suspect"
        else:
            flag = ""
        exempted = gn.upper() in exempt
        flags[gn] = {
            # 豁免基因：旗清空，但规则结果与全部计数保留（可追溯）
            "flag": "" if exempted else flag,
            "rule_flag": flag,
            "exempted": exempted,
            "condition_detect": condition_detect[gn],
            "condition_top2": condition_top2[gn],
            "condition_top5": condition_top5[gn],
            "condition_enrich5": condition_enrich5[gn],
            "family_detect": family_detect[gn],
            "family_top2": family_top2[gn],
            "family_top5": family_top5[gn],
            "family_enrich5": family_enrich5[gn],
        }
    return flags, family_metrics


def check_flag_reachability(family_metrics):
    """Report fixed background-flag rules that cannot fire for this batch size.

    The thresholds were fixed on the original 10-family batch. In a smaller
    batch the required family counts can be unreachable, and the batch would
    look clean only because nothing could be flagged. A rule can only fire if
    ONE gene accumulates the required number of family hits, so each route is
    checked against the maximum per-gene family-hit count, not the batch total
    summed over all genes; a family reaches its own top 2% only from >=50
    detected genes (top 5%: >=20).
    """
    n_families = len(family_metrics)
    top2_counts = Counter()
    top5_counts = Counter()
    enrich5_counts = Counter()
    for metrics in family_metrics.values():
        for gn, gm in metrics.items():
            combined = gm.get("combined_rank_pct")
            enrichment = gm.get("enrichment_rank_pct")
            if combined is not None and combined <= 0.02:
                top2_counts[gn] += 1
            if combined is not None and combined <= 0.05:
                top5_counts[gn] += 1
            if enrichment is not None and enrichment <= 0.05:
                enrich5_counts[gn] += 1
    top2_hits = max(top2_counts.values(), default=0)
    top5_hits = max(top5_counts.values(), default=0)
    enrich5_hits = max(enrich5_counts.values(), default=0)

    routes = [
        ("High background suspect (orange)", "family combined top 2%", 4, top2_hits),
        ("Review background suspect (yellow)", "family combined top 5%", 4, top5_hits),
        ("Review background suspect (yellow)", "family enrichment top 5%", 5, enrich5_hits),
    ]
    warnings = []
    if n_families < 4:
        warnings.append(
            f"this batch has {n_families} bait families, but every fixed rule needs at "
            "least 4 families holding a qualifying gene (the thresholds were set on the "
            "original 10-family batch), so no protein can be flagged and an empty flag "
            "list is not evidence of a clean batch"
        )
        return warnings

    for label, rule, needed, hits in routes:
        # One gene can accumulate at most one hit per family, so the best single
        # gene's family-hit count (not the sum across genes) decides reachability.
        if n_families >= needed and hits >= needed:
            continue
        percent = int(re.search(r"top (\d+)%", rule).group(1))
        min_genes = math.ceil(100 / percent)
        if n_families < needed:
            reason = (
                f"it needs {rule} in >= {needed} families, but this batch has only "
                f"{n_families} families"
            )
        else:
            reason = (
                f"it needs {rule} in >= {needed} families, but only {hits} family hit(s) "
                f"qualify; a family needs >= {min_genes} detected genes to resolve its "
                f"top {percent}% at all"
            )
        warnings.append(f"{label} cannot fire in this batch: {reason}")
    return warnings


def check_bait_direction(conditions):
    """Warn when a bait is >2-fold higher in the paired control channel.

    The analysis assumes the first raw Abundance column is the experimental
    bait channel and the second is its paired EV/negative control. A bait
    whose own normalized log2FC is below -1 is >2-fold higher in the control
    channel, the signature of swapped Abundance columns.
    """
    warnings = []
    for condition in conditions:
        aliases = condition.get("bait_aliases") or [condition["bait_gene"]]
        for gene in aliases:
            gm = gene_metric(condition["gene_metrics"], gene)
            if gm is None or gm.get("norm_fc") is None or gm["norm_fc"] >= -1:
                continue
            warnings.append(
                f"{condition['condition']}: bait {gene} is {abs(gm['norm_fc']):.1f} "
                "log2 higher in the control channel — the two Abundance columns may be reversed"
            )
            break
    return warnings


def calculate_similarity(conditions):
    all_genes = set().union(*(set(c["gene_metrics"]) for c in conditions))
    prey_median = {}
    for gn in all_genes:
        values = []
        for condition in conditions:
            gm = condition["gene_metrics"].get(gn)
            if gm and gm["max_exp"] and gm["max_exp"] > 0:
                values.append(math.log2(gm["max_exp"]))
        if values:
            prey_median[gn] = statistics.median(values)

    residual = {}
    for condition in conditions:
        residual[condition["condition"]] = {
            gn: math.log2(gm["max_exp"]) - prey_median[gn]
            for gn, gm in condition["gene_metrics"].items()
            if gm["max_exp"] and gm["max_exp"] > 0 and gn in prey_median
        }

    pairs = []
    for i, left in enumerate(conditions):
        for right in conditions[i + 1 :]:
            a, b = left["condition"], right["condition"]
            shared = sorted(set(residual[a]) & set(residual[b]))
            rho = spearman(
                [residual[a][gn] for gn in shared],
                [residual[b][gn] for gn in shared],
            )
            top_a = set(sorted(residual[a], key=residual[a].get, reverse=True)[:100])
            top_b = set(sorted(residual[b], key=residual[b].get, reverse=True)[:100])
            union = top_a | top_b
            jaccard = len(top_a & top_b) / len(union) if union else None
            pairs.append(
                {
                    "left": a,
                    "right": b,
                    "left_family": left["family"],
                    "right_family": right["family"],
                    # 同一个诱饵通道、只是对照不同（一个诱饵对多个 EV）：这不是两个
                    # 独立样本的相似性证据，比较它们是同义反复，必须单独归类。
                    "same_bait_channel": bool(left.get("exp_header"))
                    and left.get("exp_header") == right.get("exp_header"),
                    "shared": len(shared),
                    "rho": rho,
                    "jaccard": jaccard,
                    "top100_overlap": len(top_a & top_b),
                }
            )

    non_family = [
        p["rho"]
        for p in pairs
        if p["rho"] is not None
        and p["left_family"] != p["right_family"]
        and not p["same_bait_channel"]
    ]
    if len(non_family) < 3:
        for pair in pairs:
            if pair["same_bait_channel"]:
                pair["status"] = "Expected: same bait channel, different control"
            elif pair["left_family"] == pair["right_family"]:
                pair["status"] = "Expected isoform similarity"
            else:
                pair["status"] = (
                    "Insufficient non-family comparisons for profile-similarity thresholds"
                )
        return residual, pairs, None, None
    q1, q3 = np.quantile(non_family, [0.25, 0.75])
    outlier_threshold = float(q3 + 1.5 * (q3 - q1))
    review_threshold = float(np.quantile(non_family, 0.90))
    if outlier_threshold <= 0 or review_threshold <= 0:
        # A non-positive cutoff would label uncorrelated or anti-correlated
        # profiles as elevated similarity, so say it is not estimable instead.
        for pair in pairs:
            if pair["same_bait_channel"]:
                pair["status"] = "Expected: same bait channel, different control"
            elif pair["left_family"] == pair["right_family"]:
                pair["status"] = "Expected isoform similarity"
            else:
                pair["status"] = (
                    "Profile similarity not estimable at this batch size: the "
                    f"non-family distribution is not positive "
                    f"({len(non_family)} comparisons)"
                )
        return residual, pairs, None, None
    for pair in pairs:
        if pair["same_bait_channel"]:
            pair["status"] = "Expected: same bait channel, different control"
        elif pair["left_family"] == pair["right_family"]:
            pair["status"] = "Expected isoform similarity"
        elif pair["rho"] is not None and pair["rho"] >= outlier_threshold:
            pair["status"] = "High profile similarity; review mixing/biology"
        elif pair["rho"] is not None and pair["rho"] >= review_threshold:
            pair["status"] = "Elevated profile similarity; review"
        else:
            pair["status"] = ""
    return residual, pairs, outlier_threshold, review_threshold


def calculate_bait_transfer(conditions):
    by_name = {c["condition"]: c for c in conditions}
    annotations = defaultdict(list)
    records = []
    ratio_matrix = {}

    for source in conditions:
        source_gene = source["bait_gene"]
        source_gm = next(
            (
                metric
                for metric in (gene_metric(source["gene_metrics"], alias)
                               for alias in source.get("bait_aliases", [source_gene]))
                if metric is not None
            ),
            None,
        )
        source_abundance = source_gm["max_exp"] if source_gm else None
        ratio_matrix[source["condition"]] = {}
        for target in conditions:
            target_gm = next(
                (
                    metric
                    for metric in (gene_metric(target["gene_metrics"], alias)
                                   for alias in source.get("bait_aliases", [source_gene]))
                    if metric is not None
                ),
                None,
            )
            target_abundance = target_gm["max_exp"] if target_gm else None
            ratio = (
                target_abundance / source_abundance
                if source_abundance
                and source_abundance > 0
                and target_abundance
                and target_abundance > 0
                else None
            )
            ratio_matrix[source["condition"]][target["condition"]] = ratio

            if source["condition"] == target["condition"]:
                classification = (
                    "Source bait detected"
                    if source_abundance
                    else "Bait recovery QC failure"
                )
            elif source["family"] == target["family"]:
                classification = (
                    "Isoform shared-peptide / biological overlap; carryover unresolved"
                    if target_abundance
                    else ""
                )
            elif source_abundance is None:
                classification = "Not assessable: source bait not detected"
            elif target_abundance is None:
                classification = ""
            elif target["order"] == source["order"] + 1 and ratio <= 0.05:
                classification = "Possible carryover (order-dependent provisional)"
            elif ratio >= 0.05:
                classification = "Possible mixing/mislabeling or true proximity"
            elif ratio >= 0.001:
                classification = "Cross-sample signal / unresolved"
            else:
                classification = "Low-level cross-sample signal / unresolved"

            records.append(
                {
                    "source": source["condition"],
                    "source_family": source["family"],
                    "bait_gene": source_gene,
                    "source_order": source["order"],
                    "target": target["condition"],
                    "target_order": target["order"],
                    "source_abundance": source_abundance,
                    "target_abundance": target_abundance,
                    "ratio": ratio,
                    "classification": classification,
                }
            )
            if (
                source["condition"] != target["condition"]
                and target_abundance
                and classification
            ):
                annotations[(target["condition"], source_gene)].append(
                    f"{classification}; source={source['condition']}; "
                    f"cross/source={ratio:.4%}" if ratio is not None else classification
                )

    return ratio_matrix, records, annotations, by_name


def style_new_columns(ws, start_col, end_col):
    thin = Side(style="thin", color=LIGHT_BORDER)
    for col in range(start_col, end_col + 1):
        source_header = ws.cell(1, start_col - 1)
        cell = ws.cell(1, col)
        if source_header.has_style:
            cell._style = copy(source_header._style)
        else:
            cell.fill = PatternFill("solid", fgColor=BLUE_HEADER)
            cell.font = Font(bold=True)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in range(2, ws.max_row + 1):
            source = ws.cell(row, start_col - 1)
            target = ws.cell(row, col)
            if source.has_style:
                target._style = copy(source._style)
            else:
                target.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
                target.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            target.font = copy(source.font) if source.has_style else Font(name="Arial", size=10)


def add_methods_sheet(wb, condition, flags):
    name = "QC_Methods"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    rows = [
        ["Parameter", "Value", "Interpretation"],
        [
            "Condition",
            condition["condition"],
            "One bait condition; the paired EV abundance column is the shared negative control",
        ],
        [
            "Bait family",
            condition["family"],
            "Conditions in the same explicitly defined family are collapsed only for background frequency; scoring remains condition-specific",
        ],
        [
            "Bait aliases",
            ", ".join(condition.get("bait_aliases", [condition["bait_gene"]])),
            "Known GN aliases are used only for bait recovery/transfer QC; original GN values remain unchanged",
        ],
        ["Experiment column", condition["exp_header"], "Original abundance retained"],
        ["EV column", condition["ctrl_header"], "Same EV sample; not independent replicates"],
        ["Median log2FC shift", condition["median_shift"], "Subtracted from measured/bound FC"],
        ["EV 5th percentile", condition["ctrl_q05"], "Only used as lower bound when EV is missing"],
        [
            "Combined_score",
            "SQRT(Abundance_pct × Enrichment_pct)",
            "Descriptive score; not a probability or p value",
        ],
        [
            "High background suspect",
            "Family top 2% count ≥ 4",
            "Orange; review rather than delete",
        ],
        [
            "Review background suspect",
            "Family top 5% count ≥ 4 OR family enrichment top 5% count ≥ 5",
            "Yellow; review rather than delete",
        ],
        [
            "用户豁免（不计入背景旗）",
            ", ".join(
                sorted(gn for gn, data in flags.items()
                       if data.get("exempted") and gn in condition["gene_metrics"])
            ) or "（本条件无）",
            "User-exempted genes keep their rule result and counts but carry no Batch_BG_flag/colour",
        ],
        [
            "Cross-sample QC",
            "Experimental abundance profiles and bait transfer ratios",
            "Shared EV log2FC correlations are not used to call mixing",
        ],
        [
            "Order-dependent",
            "F-number order treated provisionally as acquisition/sample order",
            "Confirm with the actual injection/channel map",
        ],
        [
            "Important limitation",
            "Protein-level tables only",
            "Isoform-specific conclusions require isoform-specific peptides",
        ],
        [
            "MSstatsTMT",
            "https://pmc.ncbi.nlm.nih.gov/articles/PMC8015007/",
            "For future TMT peptide-level replicated analysis",
        ],
        [
            "SAINTexpress",
            "https://pmc.ncbi.nlm.nih.gov/articles/PMC4102138/",
            "For multiple independent controls and bait replicates",
        ],
        [
            "MiST",
            "https://pubmed.ncbi.nlm.nih.gov/25754993/",
            "Abundance, reproducibility and specificity framework",
        ],
    ]
    for row in rows:
        ws.append(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 65
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.sheet_view.showGridLines = False


def write_flagged_workbook(condition, flags, annotations):
    source = condition["path"]
    output = OUT_DIR / f"{source.stem}_flagged.xlsx"
    wb = load_workbook_compat(source, data_only=False, read_only=False)
    for sheet_name in ("+5%", "Proteins"):
        ws = wb[sheet_name]
        hdr = headers(ws)
        idx = {name: i + 1 for i, name in enumerate(hdr)}
        gn_col = idx["GN"]
        acc_col = idx.get("Accession")
        desc_col = next(
            (
                i + 1
                for i, value in enumerate(hdr)
                if isinstance(value, str) and "Description" in value
            ),
            None,
        )
        start = ws.max_column + 1
        new_headers = [
            "Norm_LOG2FC",
            "FC_type",
            "Abundance_pct",
            "Enrichment_pct",
            "Combined_score",
            "Combined_rank_pct",
            "Condition_top2_count",
            "Condition_top5_count",
            "Family_top2_count",
            "Family_top5_count",
            "Family_enrich_top5_count",
            "Batch_BG_flag",
            "Cross_sample_QC",
            "Flag_reason",
        ]
        for offset, value in enumerate(new_headers):
            ws.cell(1, start + offset, value)
        style_new_columns(ws, start, start + len(new_headers) - 1)

        comments = {
            "Norm_LOG2FC": "Measured log2FC or lower-bound FC minus the within-workbook median measured log2FC.",
            "FC_type": "Measured when both bait and EV are positive; otherwise lower-bound when EV is missing.",
            "Combined_score": "Geometric mean of abundance and enrichment percentiles; descriptive, not a p value.",
            "Batch_BG_flag": "Cross-bait background review priority; never an automatic deletion decision.",
            "Cross_sample_QC": "Protein-level carryover/mixing signal. Confirmation requires injection order and raw peptide evidence.",
        }
        for offset, value in enumerate(new_headers):
            if value in comments:
                ws.cell(1, start + offset).comment = Comment(comments[value], "Codex")

        for row in range(2, ws.max_row + 1):
            gn = str(ws.cell(row, gn_col).value or "").strip()
            accession = ws.cell(row, acc_col).value if acc_col else ""
            description = ws.cell(row, desc_col).value if desc_col else ""
            metric = condition["row_metrics"].get(record_key(accession, description, gn))
            gm = condition["gene_metrics"].get(gn)
            flag = flags.get(gn, {})
            cross = " | ".join(annotations.get((condition["condition"], gn), []))
            flag_text = flag.get("flag", "")
            reason_parts = []
            if flag.get("exempted"):
                reason_parts.append(
                    "用户豁免（不标背景旗）；规则原本为 "
                    f"{flag.get('rule_flag', '')}: family top2={flag['family_top2']}; "
                    f"family top5={flag['family_top5']}; "
                    f"family enrich top5={flag['family_enrich5']}"
                )
            if flag_text:
                reason_parts.append(
                    f"family top2={flag['family_top2']}; "
                    f"family top5={flag['family_top5']}; "
                    f"family enrich top5={flag['family_enrich5']}; "
                    f"condition top5={flag['condition_top5']}"
                )
            if cross:
                reason_parts.append(cross)

            values = [
                metric.get("norm_fc") if metric else None,
                metric.get("fc_type") if metric else "",
                metric.get("abundance_pct") if metric else None,
                metric.get("enrichment_pct") if metric else None,
                metric.get("combined") if metric else None,
                gm.get("combined_rank_pct") if gm else None,
                flag.get("condition_top2"),
                flag.get("condition_top5"),
                flag.get("family_top2"),
                flag.get("family_top5"),
                flag.get("family_enrich5"),
                flag_text,
                cross,
                " | ".join(reason_parts),
            ]
            for offset, value in enumerate(values):
                ws.cell(row, start + offset, value)

            if flag_text == "High background suspect":
                fill = PatternFill("solid", fgColor=ORANGE)
                ws.cell(row, gn_col).fill = fill
                ws.cell(row, start + 11).fill = copy(fill)
                ws.cell(row, start + 11).font = Font(bold=True, color="9C0006")
            elif flag_text == "Review background suspect":
                fill = PatternFill("solid", fgColor=YELLOW)
                ws.cell(row, gn_col).fill = fill
                ws.cell(row, start + 11).fill = copy(fill)
                ws.cell(row, start + 11).font = Font(bold=True, color="7F6000")
            if cross:
                ws.cell(row, start + 12).fill = PatternFill("solid", fgColor=PINK)
                ws.cell(row, start + 12).font = Font(color="9C0006")

        for col in (start, start + 2, start + 3, start + 4, start + 5):
            for row in range(2, ws.max_row + 1):
                ws.cell(row, col).number_format = "0.000"
        for col in range(start + 6, start + 11):
            for row in range(2, ws.max_row + 1):
                ws.cell(row, col).number_format = "0"

        widths = [14, 28, 14, 14, 14, 17, 18, 18, 16, 16, 22, 28, 52, 65]
        visible_analysis_headers = {
            "Norm_LOG2FC",
            "Combined_score",
            "Batch_BG_flag",
            "Cross_sample_QC",
        }
        for offset, width in enumerate(widths):
            column = get_column_letter(start + offset)
            dimension = ws.column_dimensions[column]
            dimension.width = width
            dimension.hidden = new_headers[offset] not in visible_analysis_headers
            dimension.outlineLevel = 1 if dimension.hidden else 0
            dimension.collapsed = False
        # Mark the ends of the three hidden helper-column groups so Excel/Calc
        # presents them as collapsible sections.
        for offset in (3, 10, 13):
            ws.column_dimensions[get_column_letter(start + offset)].collapsed = True
        ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 15, 34)

    add_methods_sheet(wb, condition, flags)
    # Some vendor-exported workbooks omit calcPr entirely.  Preserve the
    # source in that case rather than failing after the analysis columns are
    # written.
    if wb.calculation is not None:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    wb.save(output)
    # A second serialization pass makes the hidden/outline state explicit in
    # the OOXML for workbooks whose source files contain grouped columns.
    # This avoids Calc/Excel inheriting a visible state for an occasional
    # helper column while preserving the requested four visible outputs.
    saved = load_workbook_compat(output, data_only=False, read_only=False)
    for sheet_name in ("+5%", "Proteins"):
        ws = saved[sheet_name]
        start = ws.max_column - len(new_headers) + 1
        for offset, header in enumerate(new_headers):
            dimension = ws.column_dimensions[get_column_letter(start + offset)]
            dimension.hidden = header not in visible_analysis_headers
            dimension.outlineLevel = 1 if dimension.hidden else 0
    saved.save(output)
    return output


def qc_header(ws, row=1):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(color=WHITE, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def finalize_sheet(ws, widths=None, freeze="A2", filter_ref=None):
    ws.sheet_view.showGridLines = False
    if freeze:
        ws.freeze_panes = freeze
    if filter_ref:
        ws.auto_filter.ref = filter_ref
    if widths:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.row > 1:
                cell.alignment = Alignment(vertical="top", wrap_text=True)


def family_display_label(family, condition):
    """Short label for a condition inside its family, e.g. FAM227A_isoform_2 -> Isoform 2."""
    label = re.sub(r"^" + re.escape(family) + r"[_\-\s]*", "", condition)
    label = label.replace("_", " ").strip()
    if not label:
        return condition
    return label[:1].upper() + label[1:]


def family_sheet_title(family, used_titles):
    base = f"{family} Isoforms"
    title = base[:31]
    suffix = 2
    while title in used_titles:
        tail = f" ({suffix})"
        title = base[: 31 - len(tail)] + tail
        suffix += 1
    return title


def write_family_comparison(wb, family, members, used_titles):
    """Write the same-family comparison sheet for the two leading conditions.

    Conditions within one family stay separate for scoring and transfer QC;
    this sheet only documents how similar their GN-level profiles are.
    """
    left, right = members[0], members[1]
    left_label = family_display_label(family, left["condition"])
    right_label = family_display_label(family, right["condition"])
    title = family_sheet_title(family, used_titles)
    used_titles.add(title)
    iso = wb.create_sheet(title)
    shared = sorted(set(left["gene_metrics"]) & set(right["gene_metrics"]))
    abundance_rho = spearman(
        [math.log2(left["gene_metrics"][g]["max_exp"]) for g in shared],
        [math.log2(right["gene_metrics"][g]["max_exp"]) for g in shared],
    )
    normfc_rho = spearman(
        [left["gene_metrics"][g]["norm_fc"] for g in shared],
        [right["gene_metrics"][g]["norm_fc"] for g in shared],
    )
    if len(members) == 2:
        members_note = f"Compared conditions: {left['condition']} vs {right['condition']}"
    else:
        members_note = (
            f"Compared conditions: {left['condition']} vs {right['condition']}; "
            f"other family members ({', '.join(m['condition'] for m in members[2:])}) "
            "remain separate in this workbook"
        )
    metrics = [
        ["Metric", "Value", "Interpretation"],
        ["Family conditions", ", ".join(m["condition"] for m in members), members_note],
        [f"{left_label} detected genes", len(left["gene_metrics"]), ""],
        [f"{right_label} detected genes", len(right["gene_metrics"]), ""],
        ["Shared genes", len(shared), "GN-level; shared peptides are unresolved"],
        ["log2 abundance Spearman", abundance_rho, "Expected to be related, not replicates"],
        ["Normalized FC Spearman", normfc_rho, "GN-level similarity"],
    ]
    for n in (20, 50, 100, 200):
        top_left = set(
            sorted(
                left["gene_metrics"],
                key=lambda g: left["gene_metrics"][g]["norm_fc"],
                reverse=True,
            )[:n]
        )
        top_right = set(
            sorted(
                right["gene_metrics"],
                key=lambda g: right["gene_metrics"][g]["norm_fc"],
                reverse=True,
            )[:n]
        )
        union = top_left | top_right
        metrics.append(
            [
                f"Top {n} overlap",
                len(top_left & top_right),
                f"Jaccard={len(top_left & top_right)/len(union):.3f}"
                if union
                else "Jaccard=n/a: no detected genes",
            ]
        )
    for row in metrics:
        iso.append(row)
    qc_header(iso)
    title_row = iso.max_row + 3
    iso.cell(title_row, 1, "Largest isoform differences among shared GN")
    iso.cell(title_row, 1).font = Font(bold=True, size=13, color=NAVY)
    header_row = title_row + 1
    difference_headers = [
        "GN",
        f"{left_label} normFC",
        f"{right_label} normFC",
        "Absolute difference",
        f"{left_label} abundance",
        f"{right_label} abundance",
        "Caution",
    ]
    for col, value in enumerate(difference_headers, 1):
        iso.cell(header_row, col, value)
    differences = sorted(
        shared,
        key=lambda g: abs(
            left["gene_metrics"][g]["norm_fc"] - right["gene_metrics"][g]["norm_fc"]
        ),
        reverse=True,
    )
    for gn in differences[:200]:
        a, b = left["gene_metrics"][gn], right["gene_metrics"][gn]
        iso.append(
            [
                gn,
                a["norm_fc"],
                b["norm_fc"],
                abs(a["norm_fc"] - b["norm_fc"]),
                a["max_exp"],
                b["max_exp"],
                "GN-level only; confirm with isoform-specific peptides",
            ]
        )
    for cell in iso[header_row]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color=WHITE, bold=True)
    for row in range(2, iso.max_row + 1):
        for col in (2, 3, 4):
            iso.cell(row, col).number_format = "0.000"
        for col in (5, 6):
            iso.cell(row, col).number_format = "#,##0.0"
    iso.auto_filter.ref = f"A{header_row}:G{iso.max_row}"
    iso.freeze_panes = f"A{header_row + 1}"
    finalize_sheet(
        iso,
        widths={
            "A": 28,
            "B": 21,
            "C": 23,
            "D": 20,
            "E": 22,
            "F": 24,
            "G": 60,
        },
        freeze=None,
    )


def build_qc_workbook(
    conditions,
    flags,
    residual,
    pairs,
    outlier_threshold,
    review_threshold,
    ratio_matrix,
    transfer_records,
    flag_warnings=None,
    exempt=None,
):
    wb = Workbook()
    overview = wb.active
    overview.title = "Batch Overview"
    overview_headers = [
        "Order",
        "Condition",
        "Bait family",
        "Bait gene",
        "File",
        "Experiment column",
        "EV column",
        "Experiment detected",
        "EV detected",
        "Shared quantified",
        "EV missing",
        "EV missing %",
        "Median raw log2FC shift",
        "EV q05",
        "Bait abundance",
        "Bait normFC/bound",
        "Bait recovery status",
        "Sample QC",
    ]
    overview.append(overview_headers)
    for c in conditions:
        bait = c["bait_metric"]
        bait_status = "Detected" if bait else "Bait recovery QC failure"
        sample_qc = (
            "Review: bait not detected in protein-level result"
            if bait is None
            else "Pass at protein-level bait recovery check"
        )
        overview.append(
            [
                c["order"],
                c["condition"],
                c["family"],
                c["bait_gene"],
                c["file"],
                c["exp_header"],
                c["ctrl_header"],
                c["exp_detected"],
                c["ctrl_detected"],
                c["shared"],
                c["control_missing"],
                c["control_missing"] / c["exp_detected"] if c["exp_detected"] else None,
                c["median_shift"],
                c["ctrl_q05"],
                bait["max_exp"] if bait else None,
                bait["norm_fc"] if bait else None,
                bait_status,
                sample_qc,
            ]
        )
        if bait is None:
            for cell in overview[overview.max_row]:
                cell.fill = PatternFill("solid", fgColor=PINK)
    qc_header(overview)
    for row in range(2, overview.max_row + 1):
        overview.cell(row, 12).number_format = "0.0%"
        overview.cell(row, 13).number_format = "0.000"
        overview.cell(row, 14).number_format = "#,##0.0"
        overview.cell(row, 15).number_format = "#,##0.0"
        overview.cell(row, 16).number_format = "0.000"
    finalize_sheet(
        overview,
        widths={
            "A": 8,
            "B": 24,
            "C": 18,
            "D": 15,
            "E": 24,
            "F": 28,
            "G": 28,
            "H": 18,
            "I": 14,
            "J": 18,
            "K": 14,
            "L": 14,
            "M": 23,
            "N": 16,
            "O": 18,
            "P": 18,
            "Q": 26,
            "R": 45,
        },
        filter_ref=f"A1:R{overview.max_row}",
    )

    bg = wb.create_sheet("Background Candidates")
    bg_headers = [
        "GN",
        "Batch_BG_flag",
        "Condition detected",
        "Family detected",
        "Condition top2",
        "Condition top5",
        "Condition enrich top5",
        "Family top2",
        "Family top5",
        "Family enrich top5",
        "Top5 conditions",
        "Detected conditions",
        "Reason",
    ]
    bg.append(bg_headers)
    candidates = [
        (gn, data)
        for gn, data in flags.items()
        if data["flag"] in ("High background suspect", "Review background suspect")
        or data.get("exempted")   # 豁免的也列出来（灰底、旗留空），不静默
    ]
    priority = {"High background suspect": 0, "Review background suspect": 1}
    candidates.sort(
        key=lambda x: (
            priority.get(x[1]["flag"], 2),   # 豁免行（flag 为空）排最后
            -x[1]["family_top2"],
            -x[1]["family_top5"],
            x[0],
        )
    )
    for gn, data in candidates:
        top5_conditions = [
            c["condition"]
            for c in conditions
            if gn in c["gene_metrics"]
            and c["gene_metrics"][gn]["combined_rank_pct"] <= 0.05
        ]
        detected_conditions = [
            c["condition"] for c in conditions if gn in c["gene_metrics"]
        ]
        reason = (
            f"Family top2={data['family_top2']}; "
            f"family top5={data['family_top5']}; "
            f"family enrichment top5={data['family_enrich5']}"
        )
        if data.get("exempted"):
            reason = (
                f"用户豁免，不标背景旗（规则原本为 {data.get('rule_flag', '')}）；" + reason
            )
        bg.append(
            [
                gn,
                data["flag"] or ("（用户豁免，不标）" if data.get("exempted") else ""),
                data["condition_detect"],
                data["family_detect"],
                data["condition_top2"],
                data["condition_top5"],
                data["condition_enrich5"],
                data["family_top2"],
                data["family_top5"],
                data["family_enrich5"],
                ", ".join(top5_conditions),
                ", ".join(detected_conditions),
                reason,
            ]
        )
        if data.get("exempted"):
            fill = "F2F2F2"          # 灰底：豁免行，不按橙/黄
        else:
            fill = ORANGE if data["flag"].startswith("High") else YELLOW
        bg.cell(bg.max_row, 1).fill = PatternFill("solid", fgColor=fill)
        bg.cell(bg.max_row, 2).fill = PatternFill("solid", fgColor=fill)
        bg.cell(bg.max_row, 2).font = Font(bold=True)
    qc_header(bg)
    finalize_sheet(
        bg,
        widths={
            "A": 18,
            "B": 29,
            "C": 18,
            "D": 16,
            "E": 15,
            "F": 15,
            "G": 21,
            "H": 13,
            "I": 13,
            "J": 20,
            "K": 65,
            "L": 80,
            "M": 55,
        },
        filter_ref=f"A1:M{bg.max_row}",
    )

    names = [c["condition"] for c in conditions]
    sim = wb.create_sheet("Similarity Matrix")
    sim["A1"] = "Prey-centered abundance Spearman correlation"
    sim["A1"].font = Font(bold=True, size=13, color=NAVY)
    sim.append([""] + names)
    pair_map = {}
    for pair in pairs:
        pair_map[frozenset((pair["left"], pair["right"]))] = pair
    for name in names:
        row = [name]
        for other in names:
            if name == other:
                row.append(1.0)
            else:
                row.append(pair_map[frozenset((name, other))]["rho"])
        sim.append(row)
    matrix_end = 2 + len(names)
    sim.conditional_formatting.add(
        f"B3:{get_column_letter(1 + len(names))}{matrix_end}",
        ColorScaleRule(
            start_type="min",
            start_color="5B9BD5",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFFFFF",
            end_type="max",
            end_color="F8696B",
        ),
    )
    for row in range(3, matrix_end + 1):
        for col in range(2, 2 + len(names)):
            sim.cell(row, col).number_format = "0.000"

    start_j = matrix_end + 3
    sim.cell(start_j, 1, "Top-100 specificity Jaccard")
    sim.cell(start_j, 1).font = Font(bold=True, size=13, color=NAVY)
    sim.append([])
    sim.cell(start_j + 1, 1, "")
    for col, name in enumerate(names, 2):
        sim.cell(start_j + 1, col, name)
    for r_idx, name in enumerate(names, start_j + 2):
        sim.cell(r_idx, 1, name)
        for c_idx, other in enumerate(names, 2):
            value = (
                1.0
                if name == other
                else pair_map[frozenset((name, other))]["jaccard"]
            )
            sim.cell(r_idx, c_idx, value)
            sim.cell(r_idx, c_idx).number_format = "0.000"
    sim.conditional_formatting.add(
        f"B{start_j + 2}:{get_column_letter(1 + len(names))}{start_j + 1 + len(names)}",
        ColorScaleRule(
            start_type="min",
            start_color="FFFFFF",
            end_type="max",
            end_color="70AD47",
        ),
    )
    for row in (2, start_j + 1):
        for cell in sim[row]:
            if cell.column <= len(names) + 1:
                cell.fill = PatternFill("solid", fgColor=NAVY)
                cell.font = Font(color=WHITE, bold=True)
                cell.alignment = Alignment(text_rotation=45, wrap_text=True)
    sim.column_dimensions["A"].width = 26
    for col in range(2, len(names) + 2):
        sim.column_dimensions[get_column_letter(col)].width = 16
    sim.sheet_view.showGridLines = False
    sim.freeze_panes = "B3"

    pair_ws = wb.create_sheet("Pairwise Similarity")
    pair_ws.append(
        [
            "Condition A",
            "Condition B",
            "Family A",
            "Family B",
            "Shared genes",
            "Prey-centered Spearman",
            "Top100 overlap",
            "Top100 Jaccard",
            "Status",
        ]
    )
    for pair in sorted(pairs, key=lambda x: x["rho"] or -9, reverse=True):
        pair_ws.append(
            [
                pair["left"],
                pair["right"],
                pair["left_family"],
                pair["right_family"],
                pair["shared"],
                pair["rho"],
                pair["top100_overlap"],
                pair["jaccard"],
                pair["status"],
            ]
        )
        if pair["status"].startswith("High"):
            for cell in pair_ws[pair_ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=PINK)
        elif pair["status"].startswith("Elevated"):
            for cell in pair_ws[pair_ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=YELLOW)
        elif pair["status"].startswith("Expected"):
            for cell in pair_ws[pair_ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=GREEN)
    qc_header(pair_ws)
    for row in range(2, pair_ws.max_row + 1):
        pair_ws.cell(row, 6).number_format = "0.000"
        pair_ws.cell(row, 8).number_format = "0.000"
    finalize_sheet(
        pair_ws,
        widths={
            "A": 24,
            "B": 24,
            "C": 18,
            "D": 18,
            "E": 14,
            "F": 24,
            "G": 16,
            "H": 18,
            "I": 48,
        },
        filter_ref=f"A1:I{pair_ws.max_row}",
    )

    transfer = wb.create_sheet("Bait Transfer")
    matrix_values = [v for src in names for v in (ratio_matrix.get(src) or {}).values()]
    if any(v is not None for v in matrix_values):
        transfer["A1"] = "Cross/source bait abundance ratio"
    else:
        # 空表必须说明原因，否则读的人会把「算不出来」误读成「没有转运」。
        transfer["A1"] = (
            "Cross/source bait abundance ratio — 本批次无法计算：条件名里解析不出诱饵基因"
            "（文件名不含基因名），或在任何文件里都没检出该诱饵蛋白。"
            "把工作簿命名为 <序号><基因>（如 1SPO16.xlsx）后重跑即可启用此表。"
        )
    transfer["A1"].font = Font(bold=True, size=13, color=NAVY)
    transfer.append(["Source bait condition"] + names)
    for source in names:
        transfer.append(
            [source] + [ratio_matrix[source].get(target) for target in names]
        )
    for row in range(3, 3 + len(names)):
        for col in range(2, 2 + len(names)):
            transfer.cell(row, col).number_format = "0.000%"
    transfer.conditional_formatting.add(
        f"B3:{get_column_letter(1 + len(names))}{2 + len(names)}",
        ColorScaleRule(
            start_type="min",
            start_color="FFFFFF",
            mid_type="percentile",
            mid_value=90,
            mid_color=YELLOW,
            end_type="max",
            end_color="F8696B",
        ),
    )
    for cell in transfer[2]:
        if cell.column <= len(names) + 1:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(color=WHITE, bold=True)
            cell.alignment = Alignment(text_rotation=45, wrap_text=True)

    # 矩阵占第 3..(2+len(names)) 行；空一行再写明细表头，避免条件数 >=13 时
    # 明细表头覆盖矩阵最后一行（条件越多，固定行号越早失效）。
    detail_start = 4 + len(names)
    detail_headers = [
        "Source",
        "Bait gene",
        "Source order",
        "Target",
        "Target order",
        "Source abundance",
        "Target abundance",
        "Cross/source ratio",
        "Classification",
    ]
    for col, value in enumerate(detail_headers, 1):
        transfer.cell(detail_start, col, value)
    for rec in transfer_records:
        if (
            rec["source"] == rec["target"]
            or rec["target_abundance"] is not None
            or rec["classification"].startswith("Not assessable")
        ):
            transfer.append(
                [
                    rec["source"],
                    rec["bait_gene"],
                    rec["source_order"],
                    rec["target"],
                    rec["target_order"],
                    rec["source_abundance"],
                    rec["target_abundance"],
                    rec["ratio"],
                    rec["classification"],
                ]
            )
            current = transfer.max_row
            if rec["classification"].startswith(("Possible", "Bait recovery")):
                for cell in transfer[current]:
                    cell.fill = PatternFill("solid", fgColor=PINK)
            elif "unresolved" in rec["classification"]:
                for cell in transfer[current]:
                    cell.fill = PatternFill("solid", fgColor=YELLOW)
            for col in (6, 7):
                transfer.cell(current, col).number_format = "#,##0.0"
            transfer.cell(current, 8).number_format = "0.000%"
    for cell in transfer[detail_start]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color=WHITE, bold=True)
    transfer.auto_filter.ref = f"A{detail_start}:I{transfer.max_row}"
    transfer.freeze_panes = "B3"
    transfer.column_dimensions["A"].width = 26
    for col in range(2, 2 + len(names)):
        transfer.column_dimensions[get_column_letter(col)].width = 18
    transfer.column_dimensions["F"].width = 18
    transfer.column_dimensions["G"].width = 18
    transfer.column_dimensions["H"].width = 18
    transfer.column_dimensions["I"].width = 58
    transfer.sheet_view.showGridLines = False

    family_members = defaultdict(list)
    for c in conditions:
        family_members[c["family"]].append(c)
    multi_families = {
        family: sorted(members, key=lambda x: (x["order"], x["condition"]))
        for family, members in family_members.items()
        if len(members) >= 2
    }
    if multi_families:
        # Which families hold more than one isoform condition comes from the
        # metadata, so no gene name is hardcoded here.
        used_titles = set()
        for family in sorted(multi_families):
            write_family_comparison(wb, family, multi_families[family], used_titles)
    else:
        # For batches without an explicitly defined isoform pair, report the
        # family mapping without inventing an isoform comparison.
        family_ws = wb.create_sheet("Family Notes")
        family_ws.append(
            ["Bait family", "Conditions", "Condition count", "Bait genes", "Note"]
        )
        grouped = defaultdict(list)
        for c in conditions:
            grouped[c["family"]].append(c)
        for family in sorted(grouped):
            members = sorted(grouped[family], key=lambda x: x["order"])
            family_ws.append(
                [
                    family,
                    ", ".join(c["condition"] for c in members),
                    len(members),
                    ", ".join(sorted({c["bait_gene"] for c in members})),
                    "No explicit isoform pair; conditions remain independent for scoring and transfer",
                ]
            )
        qc_header(family_ws)
        finalize_sheet(
            family_ws,
            widths={"A": 22, "B": 70, "C": 16, "D": 28, "E": 75},
            filter_ref=f"A1:E{family_ws.max_row}",
        )

    methods = wb.create_sheet("Methods")
    ev_headers = sorted({c["ctrl_header"] for c in conditions})
    ev_description = (
        f"Shared paired EV column(s): {', '.join(ev_headers)}"
        if ev_headers
        else "Shared paired EV column"
    )
    similarity_cutoffs = (
        f"High outlier cutoff={outlier_threshold:.3f}; review cutoff={review_threshold:.3f}"
        if outlier_threshold is not None and review_threshold is not None
        else "Not estimated: too few non-family pairwise comparisons, or their distribution is not positive"
    )
    reachability_note = (
        "All fixed flag rules are reachable at this batch size"
        if not flag_warnings
        else " | ".join(flag_warnings)
    )
    method_rows = [
        ["Topic", "Rule / value", "Notes / source"],
        ["EV", ev_description, "Not treated as independent controls or replicates"],
        [
            "Normalized log2FC",
            "Measured log2FC − median measured log2FC within each workbook",
            "Main descriptive effect size",
        ],
        [
            "EV missing",
            "log2(bait / EV q05) − median shift",
            "Reported as FC_lower_bound, not measured FC",
        ],
        [
            "Abundance percentile",
            "Within-condition percentile; highest=1",
            "Absolute signal-strength rank",
        ],
        [
            "Enrichment percentile",
            "Within-condition normalized-FC percentile; highest=1",
            "Relative enrichment rank",
        ],
        [
            "Combined_score",
            "sqrt(Abundance_pct × Enrichment_pct)",
            "Descriptive only; not a probability",
        ],
        [
            "High background suspect",
            "Family combined top2% count ≥ 4",
            "Orange; never auto-delete",
        ],
        [
            "Review background suspect",
            "Family combined top5% count ≥ 4 OR family enrichment top5% count ≥ 5",
            "Yellow; never auto-delete",
        ],
        [
            "用户豁免（不计入背景旗）",
            "、".join(sorted(exempt)) if exempt else "（无）",
            "--exempt-genes 指定；豁免的基因仍列在 Background Candidates 里（灰底、Batch_BG_flag 留空），"
            "规则计数照常保留。用于共用诱饵自身与已知真实共同邻居这类结构性复现",
        ],
        [
            "Threshold reachability",
            reachability_note,
            "A rule that cannot fire is not evidence of a clean batch",
        ],
        [
            "Bait family handling",
            "Family-level counts collapse explicitly defined same-bait/isoform groups; all conditions remain separate for scoring",
            "Isoform specificity needs accession and isoform-specific peptides; no forced family collapse is inferred here",
        ],
        [
            "Sample similarity",
            "Spearman of prey-median-centered experimental log2 abundance",
            similarity_cutoffs,
        ],
        [
            "Carryover",
            "Bait transfer ratio plus provisional F-order adjacency",
            "Confirm with actual injection/channel order, blanks and peptide chromatograms",
        ],
        [
            "MSstatsTMT",
            "https://pmc.ncbi.nlm.nih.gov/articles/PMC8015007/",
            "Future peptide-level replicated TMT analysis",
        ],
        [
            "SAINTexpress",
            "https://pmc.ncbi.nlm.nih.gov/articles/PMC4102138/",
            "Future analysis with independent controls and bait replicates",
        ],
        [
            "MiST",
            "https://pubmed.ncbi.nlm.nih.gov/25754993/",
            "Abundance, reproducibility and specificity framework",
        ],
        [
            "Human Cell Map",
            "https://www.nature.com/articles/s41586-021-03592-2",
            "Shared BioID patterns can reflect localization rather than contamination",
        ],
    ]
    for row in method_rows:
        methods.append(row)
    if flag_warnings:
        for row in range(2, methods.max_row + 1):
            if methods.cell(row, 1).value == "Threshold reachability":
                for col in (1, 2, 3):
                    cell = methods.cell(row, col)
                    cell.fill = PatternFill("solid", fgColor=YELLOW)
                    cell.font = Font(bold=True)
    qc_header(methods)
    finalize_sheet(
        methods,
        widths={"A": 30, "B": 72, "C": 75},
        filter_ref=f"A1:C{methods.max_row}",
    )

    if wb.calculation is not None:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    wb.save(QC_PATH)


def split_bait_name(stem):
    """Split a workbook filename stem into a bait label and isoform number.

    A leading numeric order prefix is dropped, so 2PM20D2 -> PM20D2. The
    ``<order>-<isoform><bait>`` convention (18-2FAM227A, 293-23mKIAA1614) also
    yields the isoform number, which is what keeps two files for the same gene
    apart as separate conditions.
    """
    stem = re.sub(r"_processed$", "", stem, flags=re.IGNORECASE)
    isoform = re.match(r"^[0-9]+-([0-9]+)(?=[A-Za-z])", stem)
    bait = re.sub(r"^[0-9-]+", "", stem)
    return bait or stem, (int(isoform.group(1)) if isoform else None)


def read_bait_alias_table(path):
    """读 <标签> <数据库符号> 的别名表（每行一条，# 注释）；键大写。"""
    mapping = {}
    if not path:
        return mapping
    alias_path = Path(path).expanduser()
    if not alias_path.exists():
        raise SystemExit(f"bait alias 文件不存在: {path}")
    for line in alias_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.replace(",", " ").split()
        if len(parts) < 2:
            raise SystemExit(f"bait alias 行无法解析（需要 <标签> <基因>）: {line}")
        mapping[parts[0].upper()] = parts[1]
    return mapping


def build_meta_from_folder(root, label_map=None):
    """Infer condition metadata from a folder of TE-style workbooks.

    The first raw Abundance column identifies the experimental F-number. The
    filename's leading numeric/order prefix is removed to obtain the bait
    label. Files that share a bait label are treated as isoforms of one gene:
    they stay distinct conditions for scoring and transfer QC, but collapse to
    a single family for background frequency, matching the explicit isoform
    convention. Distinct bait labels remain distinct families.
    """
    entries = []
    for path in sorted(root.glob("*.xlsx")):
        if path.name.startswith(("~$", ".~")):
            continue
        # Rerun safety: when --out points at the input folder (or a previous
        # run's folder is reused as input), skip this script's own outputs so
        # they are not analyzed as if they were bait conditions.
        if path.name.endswith("_flagged.xlsx") or path.name == "BioID_batch_QC.xlsx":
            continue
        try:
            wb = load_workbook_compat(path, read_only=True, data_only=True)
            try:
                ws = wb["Proteins"]
                hdr = headers(ws)
                raw_abundance = [
                    value
                    for value in hdr
                    if isinstance(value, str)
                    and value.lower().startswith("abundance:")
                    and not value.lower().startswith("abundances (normalized)")
                ]
                if len(raw_abundance) < 2:
                    raise ValueError("少于两个原始 Abundance: 列（或文件结构读不到表头）")
                match = re.search(r"\bF(\d+)\b", raw_abundance[0])
                if not match:
                    raise ValueError(f"无法从 {raw_abundance[0]!r} 推断 F 通道号")
                ctrl_match = re.search(r"\bF(\d+)\b", raw_abundance[1])
                ctrl_channel = f"F{ctrl_match.group(1)}" if ctrl_match else None
            finally:
                wb.close()
        except Exception as error:
            # 单个文件读不了不该让整批中止：记录下来，结尾统一报告。
            SKIPPED_SOURCES.append((path.name, f"{type(error).__name__}: {error}"))
            continue
        bait, isoform = split_bait_name(path.stem)
        entries.append(
            {
                "file": path.name,
                "bait": bait,
                "isoform": isoform,
                "exp_header": raw_abundance[0],
                "ctrl_header": raw_abundance[1],
                "ctrl_channel": ctrl_channel,
                "order": int(match.group(1)),
            }
        )

    if not entries:
        raise FileNotFoundError(f"No .xlsx workbooks found in {root}")

    grouped = defaultdict(list)
    for entry in entries:
        grouped[entry["bait"]].append(entry)

    meta = {}
    conditions = defaultdict(list)
    used = set()
    for bait, members in grouped.items():
        members.sort(key=lambda x: (x["order"], x["file"]))
        mapped = (label_map or {}).get(str(bait).upper())
        bait_gene = mapped or bait
        bait_aliases = sorted({bait_gene, bait} | KNOWN_BAIT_ALIASES.get(bait, set()))
        for index, entry in enumerate(members, 1):
            if len(members) == 1:
                condition = bait
            else:
                number = entry["isoform"] or entry["order"] or index
                condition = f"{bait}_isoform_{number}"
            if condition in used:
                # 一个诱饵对多个 EV 时 F 号（= 诱饵通道）相同，靠对照通道区分条件名；
                # family 仍是同一个，背景频率不会把同一个诱饵数两次。
                token = entry.get("ctrl_channel") or entry.get("control_channel") or f"ev{index}"
                condition = f"{condition}_{token}"
            while condition in used:
                condition = f"{condition}_x"
            used.add(condition)
            conditions[condition].append(entry["file"])
            meta[entry["file"]] = {
                "condition": condition,
                "family": bait,
                "bait_gene": bait_gene,
                "bait_aliases": bait_aliases,
                "order": entry["order"],
            }

    # Two different files must never share one condition name: downstream
    # similarity and bait-transfer results are keyed by it and would otherwise
    # overwrite each other silently.
    clashes = {
        condition: files
        for condition, files in conditions.items()
        if len(files) > 1
    }
    if clashes:
        detail = "; ".join(
            f"{condition} <- {', '.join(files)}" for condition, files in sorted(clashes.items())
        )
        raise ValueError(
            "Condition names inferred from filenames collide, so results would be "
            f"overwritten: {detail}. If these files are two exports of the same "
            "condition, keep only the one you want to analyze; if they really are "
            "separate conditions, rename them as <order>-<isoform><gene> with distinct "
            "prefixes or supply explicit metadata."
        )
    return meta


def main(argv=None):
    global ROOT, OUT_DIR, QC_PATH, FILE_META
    parser = argparse.ArgumentParser(
        prog="bioid_batch_review.py",
        description=(
            "BioID 批次 QC：从一批蛋白级工作簿（bioid 的产物）推断条件与诱饵家族，"
            "计算跨批次背景频率与相似度，输出逐条件标记工作簿和一份 QC 工作簿。"
        ),
        epilog=(
            "条件、诱饵家族与诱饵基因都从工作簿文件名推断，所以目录必须显式给出；"
            "文件名形如 <顺序号><基因>[_isoform].xlsx。源文件不会被修改。"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("folder", help="包含工作簿的目录（每个条件一个 .xlsx）")
    parser.add_argument("--out", default=None,
                        help="标记工作簿与 QC 工作簿的输出目录（默认 <folder>/flagged）")
    parser.add_argument("--version", action="version", version="bioid-batch-qc 1.0.0")
    parser.add_argument("--bait-alias", default=None,
                        help="构造名→数据库符号的别名表（每行 <标签> <基因>）；"
                             "用于诱饵回收、诱饵转运与方向自检")
    parser.add_argument("--exempt-genes", default=None,
                        help="豁免基因（不参与批次背景旗）：逗号/分号分隔的符号，或一个文件路径"
                             "（每行一个符号，# 注释）；常用于共用诱饵与已知真实共同邻居")
    args = parser.parse_args(argv)
    exempt = read_exempt_genes(args.exempt_genes)
    ROOT = Path(args.folder).expanduser()
    if not ROOT.is_dir():
        parser.error(f"目录不存在：{ROOT}")
    OUT_DIR = Path(args.out).expanduser() if args.out else ROOT / "flagged"
    QC_PATH = OUT_DIR / "BioID_batch_QC.xlsx"
    alias_table = read_bait_alias_table(args.bait_alias) if args.bait_alias else {}
    if alias_table:
        print(f"诱饵别名: {len(alias_table)} 条（"
              + ", ".join(f"{k}->{v}" for k, v in list(alias_table.items())[:6])
              + (" …" if len(alias_table) > 6 else "") + "）")
    FILE_META = build_meta_from_folder(ROOT, alias_table)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    conditions = []
    for filename, meta in FILE_META.items():
        path = ROOT / filename
        if not path.exists():
            raise FileNotFoundError(path)
        try:
            condition = inspect_source(path, meta)
        except Exception as error:
            # 单个工作簿读不了不该让整批中止（与 build_meta_from_folder 一致）。
            SKIPPED_SOURCES.append((filename, f"{type(error).__name__}: {error}"))
            print(f"  WARNING: 跳过 {filename}: {type(error).__name__}: {error}")
            continue
        if condition["ctrl_q05"] is None:
            # 对照通道没有任何正值：无法计算 log2FC/富集百分位，保留它只会让
            # 下游的 None 比较出错；跳过并告警，批次里其余工作簿照常处理。
            reason = "对照通道没有任何正值，无法计算 log2FC / 富集百分位"
            SKIPPED_SOURCES.append((filename, reason))
            print(
                f"  WARNING: 跳过 {filename}: 对照通道 {condition['ctrl_header']} "
                "没有任何正值，无法计算 log2FC"
            )
            continue
        conditions.append(condition)
    conditions.sort(key=lambda x: x["order"])
    if not conditions:
        print(
            "WARNING: no usable workbook: every input workbook was skipped "
            "(see the per-file warnings above); no QC workbook was written."
        )
        raise SystemExit(1)

    flags, family_metrics = calculate_batch_flags(conditions, exempt)
    residual, pairs, outlier, review = calculate_similarity(conditions)
    ratio_matrix, transfer_records, annotations, _ = calculate_bait_transfer(conditions)
    reachability_warnings = check_flag_reachability(family_metrics)
    direction_warnings = check_bait_direction(conditions)
    flag_warnings = reachability_warnings + direction_warnings

    outputs = []
    for condition in conditions:
        outputs.append(write_flagged_workbook(condition, flags, annotations))
    build_qc_workbook(
        conditions,
        flags,
        residual,
        pairs,
        outlier,
        review,
        ratio_matrix,
        transfer_records,
        flag_warnings,
        exempt,
    )

    high = sorted(gn for gn, data in flags.items() if data["flag"].startswith("High"))
    review_genes = sorted(
        gn for gn, data in flags.items() if data["flag"].startswith("Review")
    )
    if exempt:
        hit = sorted(gn for gn in flags if gn.upper() in exempt)
        missing = sorted(g for g in exempt if g not in {gn.upper() for gn in flags})
        print(f"豁免背景旗（--exempt-genes）: {len(hit)} 个命中 —— " + "、".join(hit))
        if missing:
            print("  名单里未出现在本批的符号（检查拼写）: " + "、".join(missing))
    print(f"Flagged workbooks: {len(outputs)}")
    print(f"QC workbook: {QC_PATH}")
    if SKIPPED_SOURCES:
        print(f"Skipped workbooks: {len(SKIPPED_SOURCES)}（未进入本次 QC）")
        for name, why in SKIPPED_SOURCES:
            print(f"  - {name}: {why}")
    # 条件与家族一律从文件名推断，所以这段提示总是适用：让用户先核对家族映射。
    grouped = defaultdict(list)
    for condition in conditions:
        grouped[condition["family"]].append(condition["condition"])
    merged = {f: c for f, c in grouped.items() if len(c) > 1}
    if merged:
        print(
            "Inferred multi-condition families (families are collapsed for "
            "background frequency; confirm the mapping):"
        )
        for family in sorted(merged):
            print(f"  {family}: {', '.join(sorted(merged[family]))}")
    print(f"High background suspect: {len(high)}")
    print(", ".join(high))
    print(f"Review background suspect: {len(review_genes)}")
    print(", ".join(review_genes))
    if reachability_warnings:
        print(
            "WARNING: fixed background-flag thresholds are not reachable in this batch:"
        )
        for warning in reachability_warnings:
            print(f"  - {warning}")
        print(
            "  Report this with the results; do not present an empty flag list as a clean batch."
        )
    for warning in direction_warnings:
        print(f"WARNING: {warning}")
    if outlier is None or review is None:
        print(
            "Similarity thresholds: not estimated (too few non-family comparisons "
            "or a non-positive distribution)"
        )
    else:
        print(f"Similarity high cutoff: {outlier:.4f}; review cutoff: {review:.4f}")
    for pair in sorted(pairs, key=lambda x: x["rho"] or -9, reverse=True)[:8]:
        print(
            f"{pair['left']} vs {pair['right']}: "
            f"rho={fmt(pair['rho'])}, Jaccard={fmt(pair['jaccard'])}, "
            f"{pair['status']}"
        )


if __name__ == "__main__":
    main()
