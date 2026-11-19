#!/usr/bin/env python3
"""bioid-repeat —— 同一个诱饵在不同细胞系/批次之间的重复性散点图。

比较"同一诱饵、同一批 prey 在两个独立实验里的富集程度"，用于判断结果能否重复。
输入是成对的蛋白级工作簿（bioid-batch-qc 的 `*_flagged.xlsx`，或 bioid 的 `*_processed.xlsx`），
输出一个自包含 HTML：每对一节，含散点图（横纵坐标为选定值）、汇总统计、以及只在一边检出的蛋白。

用法:
  # 一对（标签自动取自文件名）
  bioid_repeat.py --pair A.xlsx B.xlsx

  # 多对，带标签（每行 <标签> <文件A> <文件B>，制表符或空白分隔，# 开头为注释）
  bioid_repeat.py --pairs pairs.tsv

  # 先只看解析结果（不写文件，用于人工确认输入对不对）
  bioid_repeat.py --pairs pairs.tsv --list

常用参数:
  --value norm_log2fc|log2fc|fc   纵/横坐标用哪个量（默认 norm_log2fc）
  --out DIR                       输出目录（默认 <第一个文件所在目录>/bioid_repeat）
  --label-top N                   每张图最多标几个名字（默认 8）
  --plotly inline|cdn             默认 inline（自包含、离线可看）
"""

from __future__ import annotations

import argparse
import atexit
import json
import math
import os
import re
import statistics
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

SKILL_DIR = Path(__file__).resolve().parent.parent
PLOTLY_JS = SKILL_DIR / "assets" / "plotly.min.js"
# 矢量 PDF 导出用（MIT）：jsPDF + svg2pdf.js，随包内联，运行时不联网
JSPDF_JS = SKILL_DIR / "assets" / "jspdf.umd.min.js"
SVG2PDF_JS = SKILL_DIR / "assets" / "svg2pdf.umd.min.js"
# PDF 里的中文字形（Noto Sans SC 子集，SIL OFL）：jsPDF 自带字体没有 CJK，不内嵌就只有乱码
NOTO_TTF = SKILL_DIR / "assets" / "noto_sans_sc_subset.ttf"
TEMPLATE = SKILL_DIR / "assets" / "report_template.html"
# 与 bioid-analysis 共用同一份 UniProt 注释缓存（accession -> 名称/基因/别名/功能）
ANNOTATION_TABLE = SKILL_DIR / "assets" / "uniprot_annotations.tsv"
# CRAPome 频率表同样与 bioid-analysis 共用（同套件、同目录下的兄弟 skill）；
# 找不到就只是不标污染物，不影响其它输出。
CRAPOME_TABLE = SKILL_DIR.parent / "bioid-analysis" / "assets" / "crapome_frequency.tsv"
CRAPOME_DEFAULTS = {"frequency": 0.20, "min_sc": 2, "cell": "HEK293"}

COLORS = {
    "point": "#2b6cb0",
    "censored": "#0b7285",
    "bait": "#8e44ad",
    "bg": "#c92a2a",
    # 与 bioid-analysis 同一套语义色：批次 QC 背景可疑 / CRAPome 污染物
    "qc_bg": "#c92a2a",
    "crapome": "#b26a00",
}
CONTEXT_COLOR = {"bait": "#8e44ad", "censored": "#0b7285", "bg": "#c92a2a"}


# --------------------------------------------------------------------------
# 与 openxlsx 的互操作（老批次的工作簿可能是 openxlsx 写的，openpyxl 直接读不了）
# --------------------------------------------------------------------------
_REPAIR_DIR = None


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


def repair_openxlsx_package(path):
    """必要时在临时副本上修掉悬空关系与 dimension="A1" 两个缺陷，返回可读路径。"""
    global _REPAIR_DIR
    try:
        with zipfile.ZipFile(path) as archive:
            members = {name: archive.read(name) for name in archive.namelist()}
    except zipfile.BadZipFile:
        return path
    changed = False
    for name in list(members):
        if name.endswith(".rels") and "/_rels/" in name:
            body = members[name].decode("utf-8", "ignore")
            base = name.replace("/_rels/", "/").rsplit("/", 1)[0] + "/"
            for target in re.findall(r'Target="([^"]+)"', body):
                if target.startswith(("http://", "https://", "/")):
                    continue
                if os.path.normpath(base + target) not in members:
                    body = re.sub(r'<Relationship [^>]*Target="%s"[^>]*/>' % re.escape(target), "", body)
                    changed = True
            members[name] = body.encode("utf-8")
        elif re.match(r"xl/worksheets/sheet\d*\.xml$", name):
            text = members[name].decode("utf-8", "ignore")
            match = re.search(r'<dimension ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"', text)
            if match and match.group(3) is not None:
                continue
            refs = re.findall(r'<c r="([A-Z]+)(\d+)"', text)
            if not refs:
                continue
            element = '<dimension ref="A1:%s%d"/>' % (
                _int_to_col(max(_col_to_int(c) for c, _ in refs)), max(int(r) for _, r in refs))
            text = re.sub(r"<dimension[^>]*/>", element, text, count=1) if "<dimension" in text \
                else text.replace("<sheetViews>", element + "<sheetViews>", 1)
            members[name] = text.encode("utf-8")
            changed = True
    if not changed:
        return path
    if _REPAIR_DIR is None:
        _REPAIR_DIR = tempfile.mkdtemp(prefix="bioid_repeat_")
        atexit.register(lambda: __import__("shutil").rmtree(_REPAIR_DIR, ignore_errors=True))
    out = os.path.join(_REPAIR_DIR, os.path.basename(path))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    return out


# --------------------------------------------------------------------------
# 读取与列定位
# --------------------------------------------------------------------------
def _norm(text):
    return re.sub(r"\s+", " ", str(text or "").strip()).lower()


def column_index(header):
    index = {}
    for i, raw in enumerate(header, 1):
        name = _norm(raw)
        if not name or name in index:
            continue
        index[name] = i
    return index


def to_num(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def to_int(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


VALUE_SOURCES = {
    # --value 取值 -> 依次尝试的表头（规范化后）
    "norm_log2fc": ["norm_log2fc"],
    "log2fc": ["log2fc", "log2fc (ev-corrected)", "log2fc (normalized)", "log2fc (raw)"],
    "fc": ["log2fc", "log2fc (ev-corrected)", "log2fc (normalized)", "fc (ev-corrected)", "fc (raw)"],
}


def load_pair_member(path, value_mode="norm_log2fc"):
    """读一份工作簿，返回 {原文件, 标签, 诱饵, 值列, 蛋白{key: {...}}, 计数}。"""
    path = Path(path).expanduser()
    if not path.exists():
        raise SystemExit(f"文件不存在: {path}")
    book_path = repair_openxlsx_package(str(path))
    wb = load_workbook(book_path, read_only=True, data_only=True)
    sheet = "Proteins" if "Proteins" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    index = column_index(header)

    def col(*names):
        for name in names:
            if name in index:
                return index[name]
        return None

    i_gn, i_acc = col("gn"), col("accession")
    i_desc, i_up, i_cov = col("description"), col("# unique peptides"), col("coverage [%]")
    i_type, i_bg, i_comb = col("fc type"), col("batch_bg_flag"), col("combined_score")
    ab = [i for i, raw in enumerate(header, 1)
          if isinstance(raw, str) and re.match(r"abundance[:\s]", raw.strip().lower())
          and "normalized" not in raw.lower()]
    i_ev = ab[1] if len(ab) > 1 else None

    # 值列：log2 优先；要求 fc 时若只有线性列则直接用
    value_col = None
    for name in VALUE_SOURCES.get(value_mode, []):
        if name in index:
            value_col = (index[name], name)
            break
    if value_col is None:
        raise SystemExit(
            f"{path.name}: 找不到 {value_mode} 需要的列（可用表头: "
            + ", ".join(sorted(k for k in index if k))[:180] + "）")
    vcol, vname = value_col

    proteins = {}
    n_rows = 0
    for row in rows:
        if row is None:
            continue
        gn = str(row[i_gn - 1] or "").strip() if i_gn and i_gn <= len(row) else ""
        acc = str(row[i_acc - 1] or "").strip() if i_acc and i_acc <= len(row) else ""
        if not gn and not acc:
            continue
        n_rows += 1
        value = to_num(row[vcol - 1]) if vcol <= len(row) else None
        if value_mode == "fc" and vname.startswith("log2fc") and value is not None:
            value = 2 ** value
        ev_value = to_num(row[i_ev - 1]) if i_ev and i_ev <= len(row) else None
        kind = str(row[i_type - 1] or "").lower() if i_type and i_type <= len(row) else ""
        censored = "lower" in kind or (i_ev is not None and ev_value is None)
        key = re.split(r"[;\-]", acc.split(";")[0])[0].strip() or gn.upper()
        proteins[key] = {
            "gn": gn, "acc": acc.split(";")[0].strip(),
            "desc": (str(row[i_desc - 1] or "").strip()[:120] if i_desc and i_desc <= len(row) else ""),
            "value": value,
            "up": to_int(row[i_up - 1]) if i_up and i_up <= len(row) else None,
            "cov": to_num(row[i_cov - 1]) if i_cov and i_cov <= len(row) else None,
            "censored": bool(censored),
            "bg": str(row[i_bg - 1] or "").strip() if i_bg and i_bg <= len(row) else "",
            "combined": to_num(row[i_comb - 1]) if i_comb and i_comb <= len(row) else None,
        }
    wb.close()
    label = re.sub(r"(_processed|_flagged)+$", "", path.stem, flags=re.IGNORECASE)
    return {
        "file": str(path), "name": path.name, "label": label,
        "value_column": vname, "n_rows": n_rows, "proteins": proteins,
    }


# --------------------------------------------------------------------------
# 统计
# --------------------------------------------------------------------------
def ranks(values):
    order = sorted(range(len(values)), key=lambda i: values[i])
    out = [0.0] * len(values)
    i = 0
    while i < len(order):
        j = i + 1
        while j < len(order) and values[order[j]] == values[order[i]]:
            j += 1
        avg = (i + 1 + j) / 2
        for k in range(i, j):
            out[order[k]] = avg
        i = j
    return out


def pearson(x, y):
    if len(x) < 3:
        return None
    mx, my = statistics.mean(x), statistics.mean(y)
    sx = sum((a - mx) ** 2 for a in x)
    sy = sum((b - my) ** 2 for b in y)
    if sx == 0 or sy == 0:
        return None
    return sum((a - mx) * (b - my) for a, b in zip(x, y)) / math.sqrt(sx * sy)


def spearman(x, y):
    return pearson(ranks(x), ranks(y))


def ols(x, y):
    n = len(x)
    if n < 3:
        return None, None
    mx, my = statistics.mean(x), statistics.mean(y)
    sxx = sum((a - mx) ** 2 for a in x)
    if sxx == 0:
        return None, None
    slope = sum((a - mx) * (b - my) for a, b in zip(x, y)) / sxx
    return slope, my - slope * mx


# --------------------------------------------------------------------------
# 组装
# --------------------------------------------------------------------------
def clean_label(text):
    """去掉文件名里的序号前缀，得到基因名样式（1AKAP8L -> AKAP8L）。"""
    return re.sub(r"^[0-9\-]+", "", str(text)).strip() or str(text)


def bait_gene_from_label(label):
    """从配对标签里取诱饵基因：取 "vs"/":"/空白 之前的部分，再去掉序号前缀。"""
    token = re.split(r"\s+vs\s+|[:：\s]", str(label), maxsplit=1)[0]
    return clean_label(token).upper()


def distinguishing_path(path_a, path_b, depth=4):
    """从文件往上找第一段不同的目录，用它区分两个同名样本（如 BioID-1st/flagged）。"""
    a, b = Path(path_a).parent.parts, Path(path_b).parent.parts   # 只比目录，忽略文件名
    for k in range(1, min(len(a), len(b), depth) + 1):
        if a[-k] != b[-k]:
            return "/".join(a[-k:])
    return a[-1]


def auto_pair_label(left, right):
    la, lb = clean_label(left["label"]), clean_label(right["label"])
    if la.upper() == lb.upper():
        return (f"{la} · {distinguishing_path(left['file'], right['file'])}"
                f" vs {distinguishing_path(right['file'], left['file'])}")
    return f"{left['label']} vs {right['label']}"


def split_pair_label(label):
    """"ATOSB: 2nd vs 4th" -> ("ATOSB: 2nd", "4th")；没有 vs 时返回 (None, None)。"""
    parts = re.split(r"\s+vs\.?\s+", str(label), maxsplit=1, flags=re.IGNORECASE)
    if len(parts) == 2 and parts[0].strip() and parts[1].strip():
        return parts[0].strip(), parts[1].strip()
    return None, None


def side_label(member, peer):
    """坐标轴上的名字：两侧同名时补上层目录以区分（如 AKAP8L · BioID-1st/flagged）。"""
    mine, theirs = clean_label(member["label"]), clean_label(peer["label"])
    if mine.upper() == theirs.upper():
        return f"{mine} · {distinguishing_path(member['file'], peer['file'])}"
    return mine


def build_pair(left, right, pair_label, value_mode, label_top):
    """把两份工作簿合成一张统一的行表（前端只用这一份数据画图和列表）。

    每行的 kind 取值：
      both      —— 两侧都定量到（含 censored 标记：某侧对照未检出，比值为下限）
      unquant_a —— A 侧有值，B 侧在该文件里没有数值
      unquant_b —— B 侧有值，A 侧没有数值
      none      —— 两侧都没有数值（bioid 的隐藏行里本来就有这类）
      only_a    —— 只在 A 文件里出现
      only_b    —— 只在 B 文件里出现
    """
    rows = []

    def base_of(a, b):
        return {
            "gn": a["gn"] or b["gn"], "acc": a["acc"] or b["acc"],
            "desc": a["desc"] or b["desc"],
            "up": min([v for v in (a["up"], b["up"]) if v is not None] or [1]),
            "cov": max([v for v in (a["cov"], b["cov"]) if v is not None] or [0]),
            "bg": a["bg"] or b["bg"],
            # CRAPome：两侧取较强的那个（一侧被判污染物就标出来）
            "cfreq": max([v for v in (a.get("cfreq"), b.get("cfreq")) if v is not None] or [0.0]),
            "crapome": bool(a.get("crapome") or b.get("crapome")),
        }

    for key in sorted(set(left["proteins"]) & set(right["proteins"])):
        a, b = left["proteins"][key], right["proteins"][key]
        row = base_of(a, b)
        row["gn"] = row["gn"] or key
        censored = bool(a["censored"] or b["censored"])
        if a["value"] is not None and b["value"] is not None:
            rows.append(dict(row, kind="both", x=a["value"], y=b["value"], censored=censored))
        elif a["value"] is not None:
            rows.append(dict(row, kind="unquant_b", x=a["value"], y=None, censored=censored))
        elif b["value"] is not None:
            rows.append(dict(row, kind="unquant_a", x=None, y=b["value"], censored=censored))
        else:
            rows.append(dict(row, kind="none", x=None, y=None, censored=censored))
    for key in sorted(set(left["proteins"]) - set(right["proteins"])):
        a = left["proteins"][key]
        row = base_of(a, a)
        row["gn"] = row["gn"] or key
        rows.append(dict(row, kind="only_a", x=a["value"], y=None, censored=bool(a["censored"])))
    for key in sorted(set(right["proteins"]) - set(left["proteins"])):
        b = right["proteins"][key]
        row = base_of(b, b)
        row["gn"] = row["gn"] or key
        rows.append(dict(row, kind="only_b", x=None, y=b["value"], censored=bool(b["censored"])))

    both = [r for r in rows if r["kind"] == "both"]
    measured = [r for r in both if not r["censored"]]
    xs = [r["x"] for r in measured]
    ys = [r["y"] for r in measured]
    rho, r = spearman(xs, ys), pearson(xs, ys)
    slope, intercept = ols(xs, ys)

    label_left, label_right = split_pair_label(pair_label)
    left_axis = label_left or side_label(left, right)
    right_axis = label_right or side_label(right, left)
    bait_key = bait_gene_from_label(pair_label)

    n = lambda kind: sum(1 for r in rows if r["kind"] == kind)
    verdict = "—"
    if rho is not None:
        verdict = ("重复性好（Spearman r ≥ 0.8）" if rho >= 0.8 else
                   "中等重复（0.5 ≤ r < 0.8）" if rho >= 0.5 else
                   "弱重复（0.2 ≤ r < 0.5）" if rho >= 0.2 else "基本不重复（r < 0.2）")

    # 标注候选：只标"图上真的画出来"的点（曾经把未定量的点也拿去标注，于是出现
    # 只有文字、没有点的孤儿标签）。诱饵与手工标记由前端优先处理。
    ranked = sorted(both, key=lambda r: -((r["x"] + r["y"]) / 2))

    return {
        "label": pair_label,
        "left": {k: left[k] for k in ("file", "name", "label", "value_column", "n_rows")},
        "right": {k: right[k] for k in ("file", "name", "label", "value_column", "n_rows")},
        "value_label": "FC" if value_mode == "fc" else "log2FC",
        "value_column": left["value_column"] if left["value_column"] == right["value_column"]
                        else f'{left["value_column"]} / {right["value_column"]}',
        "bait_gene": bait_key,
        "left_axis": left_axis, "right_axis": right_axis,
        "n_rows": len(rows),
        "n_both": len(both), "n_measured": len(measured),
        "n_censored": len(both) - len(measured),
        "n_unquant_a": n("unquant_a"), "n_unquant_b": n("unquant_b"),
        "n_none": n("none"), "n_only_a": n("only_a"), "n_only_b": n("only_b"),
        "rho": rho, "r": r, "slope": slope, "intercept": intercept,
        "delta": math.log2(1.5) if value_mode != "fc" else 1.5,
        "verdict": verdict,
        "unit": "×" if value_mode == "fc" else "",
        "rows": rows,
        "label_top": label_top,
        "mark_key": f"{pair_label}|{left['name']}|{right['name']}",
        "ranked_gn": [r["gn"] for r in ranked[:400]],
    }


def load_crapome(path, cell):
    """读紧凑版 CRAPome 频率表（与 bioid-analysis 同一份资产，格式见其 references/crapome.md）。"""
    if not path:
        return None
    target = Path(path).expanduser()
    if not target.exists():
        return None
    import csv
    use_hek = bool(cell) and str(cell).lower().startswith("hek")
    lines = [line for line in target.read_text(encoding="utf-8").splitlines()
             if not line.startswith("#")]
    table = {"by_uniprot": {}, "by_gene": {}, "runs": 0, "source": str(target),
             "cell": cell if use_hek else "all human",
             "hits_key": "_hits_hek" if use_hek else "_hits_all"}
    for row in csv.DictReader(lines, delimiter="\t"):
        gene = (row.get("gene") or "").strip().upper()
        uniprot = (row.get("uniprot") or "").strip().upper()
        row["_hits_all"] = to_int(row.get("n_ctrl_ge2"))
        row["_hits_hek"] = to_int(row.get("n_ctrl_hek_ge2"))
        if uniprot and ";" not in uniprot:
            table["by_uniprot"].setdefault(uniprot, row)
        if gene:
            table["by_gene"].setdefault(gene, row)
        table["runs"] = to_int(row.get("n_ctrl_hek_total") if use_hek else row.get("n_ctrl_total")) or table["runs"]
    return table


def annotate_crapome(workbook, table, threshold, min_sc):
    """给工作簿的每个蛋白附上 CRAPome 控制组频率与污染物标记（只标注，不删行）。"""
    if workbook is None:
        return 0
    flagged = 0
    for protein in workbook["proteins"].values():
        if table is None:
            protein["cfreq"], protein["crapome"] = None, False
            continue
        row = (table["by_uniprot"].get((protein["acc"] or "").upper())
               or table["by_gene"].get((protein["gn"] or "").upper()))
        hits = to_int(row.get(table["hits_key"])) if row else 0
        runs = table["runs"] or 0
        frequency = (hits / runs) if runs else (0.0 if row is None else None)
        protein["cfreq"] = frequency
        protein["crapome"] = bool(frequency is not None
                                  and frequency >= threshold and hits >= min_sc)
        flagged += 1 if protein["crapome"] else 0
    return flagged


def load_annotations(path):
    """读注释缓存：accession -> (name, gene, aliases, function)（与 bioid-analysis 同格式）。"""
    import csv
    if not path or not Path(path).exists():
        return {}
    table = {}
    with open(path, newline="", encoding="utf-8") as fh:
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


def build_annotation_payload(pairs, table):
    """每个 accession 只存一份，避免在多个配对里重复膨胀。"""
    shared = {}
    for pair in pairs:
        for row in pair["rows"]:
            key = row.get("acc") or ""
            key = key.split(";")[0].split("-")[0].strip()
            if not key or key in shared:
                continue
            entry = table.get(key)
            if entry:
                shared[key] = entry
    return shared


def render(payload, out_path, plotly_mode="inline"):
    template = TEMPLATE.read_text(encoding="utf-8")

    def lib_tag(path, cdn_url):
        if plotly_mode == "cdn":
            return f'<script src="{cdn_url}"></script>'
        if not path.exists():
            raise SystemExit(f"缺少 {path}；用 --plotly cdn 或恢复随 skill 附带的库文件")
        return "<script>" + path.read_text(encoding="utf-8", errors="ignore") + "</script>"

    plotly = lib_tag(PLOTLY_JS, "https://cdn.plot.ly/plotly-2.35.2.min.js")
    jspdf = lib_tag(JSPDF_JS, "https://unpkg.com/jspdf@2.5.2/dist/jspdf.umd.min.js")
    svg2pdf = lib_tag(SVG2PDF_JS, "https://unpkg.com/svg2pdf.js@2.2.3/dist/svg2pdf.umd.min.js")
    if NOTO_TTF.exists():
        import base64 as _b64
        font_tag = '<script>window.NOTO_FONT_B64="' + _b64.b64encode(NOTO_TTF.read_bytes()).decode() + '";</script>'
    else:
        font_tag = "<script>window.NOTO_FONT_B64=null;</script>"
    html = (template
            .replace("/*__PLOTLY__*/", plotly)
            .replace("/*__JSPDF__*/", jspdf)
            .replace("/*__SVG2PDF__*/", svg2pdf)
            .replace("/*__FONTB64__*/", font_tag)
            .replace("/*__PAYLOAD__*/", json.dumps(payload, ensure_ascii=False)))
    out_path.write_text(html, encoding="utf-8")


# --------------------------------------------------------------------------
# 输入解析
# --------------------------------------------------------------------------
def read_pairs_file(path):
    pairs = []
    for lineno, line in enumerate(Path(path).expanduser().read_text(encoding="utf-8").splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p for p in re.split(r"[\t,]+", line) if p.strip()]
        if len(parts) == 2:
            label = None
            left, right = parts
        elif len(parts) >= 3:
            label, left, right = parts[0], parts[1], parts[2]
        else:
            raise SystemExit(f"{path}:{lineno} 无法解析（需要 <标签> <文件A> <文件B>）: {line}")
        pairs.append((label, left.strip(), right.strip()))
    if not pairs:
        raise SystemExit(f"{path} 里没有可用的配对")
    return pairs


def main():
    parser = argparse.ArgumentParser(description="同一诱饵在不同细胞系/批次之间的相关性散点图")
    parser.add_argument("--pairs", help="配对表：每行 <标签> <文件A> <文件B>")
    parser.add_argument("--pair", nargs=2, action="append", metavar=("FILE_A", "FILE_B"),
                        help="单对输入（可重复；标签取自文件名）")
    parser.add_argument("--value", default="norm_log2fc", choices=["norm_log2fc", "log2fc", "fc"],
                        help="坐标用哪个量（默认 norm_log2fc，即 batch-QC 的批内中心化 log2FC）")
    parser.add_argument("--label-top", type=int, default=8, help="每张图最多标几个名字")
    parser.add_argument("--out", default=None, help="输出目录（默认 <首个文件所在目录>/bioid_repeat）")
    parser.add_argument("--annotations", default=str(ANNOTATION_TABLE),
                        help="UniProt 注释缓存（默认用 skill 自带的那份，与 bioid-analysis 共用）")
    parser.add_argument("--crapome-table", default=str(CRAPOME_TABLE),
                        help="CRAPome 频率表（默认用兄弟 skill bioid-analysis 里那份）")
    parser.add_argument("--no-crapome", action="store_true", help="不做 CRAPome 污染物标注")
    parser.add_argument("--crapome-frequency", type=float, default=CRAPOME_DEFAULTS["frequency"],
                        help="判为污染物的控制组检出频率阈值（默认 0.20）")
    parser.add_argument("--crapome-min-sc", type=int, default=CRAPOME_DEFAULTS["min_sc"],
                        help="判为污染物所需的最少控制组谱图数（默认 2）")
    parser.add_argument("--crapome-cell", default=CRAPOME_DEFAULTS["cell"],
                        help="用哪套控制组频率（默认 HEK293；换别的名字则退回全部人类对照）")
    parser.add_argument("--plotly", default="inline", choices=["inline", "cdn"])
    parser.add_argument("--list", action="store_true", help="只列出解析到的配对，不生成报告")
    args = parser.parse_args()

    crapome = None if args.no_crapome else load_crapome(args.crapome_table, args.crapome_cell)

    pairs = []
    if args.pairs:
        pairs += read_pairs_file(args.pairs)
    if args.pair:
        pairs += [(None, a, b) for a, b in args.pair]
    if not pairs:
        raise SystemExit("需要 --pairs <tsv> 或 --pair FILE_A FILE_B（可用 --list 先看解析结果）")

    resolved = []
    n_crapome = 0
    for label, left, right in pairs:
        l = load_pair_member(left, args.value)
        r = load_pair_member(right, args.value)
        for member in (l, r):
            n_crapome += annotate_crapome(member, crapome,
                                          args.crapome_frequency, args.crapome_min_sc)
        name = label or auto_pair_label(l, r)
        resolved.append((name, l, r))

    payload_pairs = [build_pair(l, r, name, args.value, args.label_top) for name, l, r in resolved]

    if args.list:
        print(f"共 {len(resolved)} 对（--value {args.value}）：\n")
        if crapome:
            print(f"CRAPome: {Path(crapome['source']).name} · {crapome['cell']} 对照 {crapome['runs']} 例 · "
                  f"阈值 ≥{args.crapome_frequency:.0%} 且 ≥{args.crapome_min_sc} SC · 两侧合计标了 {n_crapome} 行\n")
        else:
            print("CRAPome: 未应用（--no-crapome 或找不到频率表）\n")
        for name, l, r in resolved:
            print(f"  {name}")
            print(f"    A: {l['name']}")
            print(f"       {len(l['proteins'])} 个蛋白（{l['n_rows']} 行），用列: {l['value_column']}")
            print(f"    B: {r['name']}")
            print(f"       {len(r['proteins'])} 个蛋白（{r['n_rows']} 行），用列: {r['value_column']}")
            shared = len(set(l["proteins"]) & set(r["proteins"]))
            print(f"    可匹配: {shared} 个（按 Accession 归并）")
            if l["label"].upper() != r["label"].upper():
                print(f"    注意: 两侧文件名标签不同（{l['label']} vs {r['label']}）—— 若确实是同一个诱饵请忽略")
            print()
        print("确认无误后去掉 --list 正式生成报告。")
        return

    annotation_table = load_annotations(args.annotations)
    payload = {
        "generated": date.today().isoformat(),
        "annotations": build_annotation_payload(payload_pairs, annotation_table) if annotation_table else {},
        "value_mode": args.value,
        "colors": COLORS,
        "crapome": {
            "applied": bool(crapome),
            "cell": crapome["cell"] if crapome else None,
            "runs": crapome["runs"] if crapome else None,
            "frequency": args.crapome_frequency,
            "min_sc": args.crapome_min_sc,
            "source": Path(crapome["source"]).name if crapome else None,
            "n_rows_flagged": n_crapome,
        },
        "pairs": payload_pairs,
    }
    out_dir = Path(args.out).expanduser() if args.out else \
        Path(resolved[0][1]["file"]).parent / "bioid_repeat"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "bioid_repeat_report.html"
    render(payload, out_path, args.plotly)

    print(f"{len(payload['pairs'])} 对比较 → {out_path}")
    if crapome:
        print(f"  CRAPome 标注: {Path(crapome['source']).name} · {crapome['cell']} 对照 {crapome['runs']} 例 · "
              f"≥{args.crapome_frequency:.0%} 且 ≥{args.crapome_min_sc} SC 判为污染物；两侧合计标了 {n_crapome} 行")
    else:
        print("  CRAPome 标注: 未应用（--no-crapome 或找不到频率表）")
    for pair in payload["pairs"]:
        rho = pair["rho"]
        matched = pair["n_both"] + pair["n_unquant_a"] + pair["n_unquant_b"] + pair["n_none"]
        print(f"  {pair['label']:<28} 行 {pair['n_rows']:>5} · 匹配 {matched:>5} · 可算相关 {pair['n_measured']:>5} · "
              f"截尾 {pair['n_censored']:>4} · 只在单侧 {pair['n_only_a'] + pair['n_only_b']:>4} · Spearman r = "
              f"{'—' if rho is None else format(rho, '.3f')} · {pair['verdict']}")


if __name__ == "__main__":
    main()
