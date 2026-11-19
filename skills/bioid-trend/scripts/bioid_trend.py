#!/usr/bin/env python3
"""bioid-trend -- 把同一个（组）蛋白在多个 BioID 工作簿里的标准化 log2FC 画成折线图。

输入两个小表（TSV/CSV，允许 # 注释）：
  * 样本表：一个 x 轴点一行，`label` + `file`（可选 `sheet` / `column`）
  * 蛋白表：一个折线一行，`gene`（可选 `group` / `label` / `style` / `color` / `accession`）

缺省按 batch-QC 的 `Norm_LOG2FC` 取值；没有就退回 `LOG2FC`（bioid 产物）或
`log2FC (EV-corrected)`（bioid-analysis 产物）。每个工作簿用了哪一列都打印出来。

某蛋白在某次 BioID 里没检出（整行都没有，或该行有但数值为空）时：
画在 y=0、空心 marker，线照常连过去 —— 0 是绘图位置，不是测量值。

只读输入；不修改任何工作簿。
"""
from __future__ import annotations

import argparse
import math
import re
import sys
import textwrap
from collections import OrderedDict
from pathlib import Path

# 与参考图一致的 tab10 顺序；蛋白按表内顺序取色
PALETTE = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
           "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]
# marker 形状按 group 首次出现顺序分配（参考图：A scaffold=圆、C catalytic=方、B55=三角）
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "<", ">", "h"]
# 所有蛋白主线一律实线；虚线只用于“有端点未检出”的线段（见 segment_runs）
# x 轴背景分组：浅到不盖住线，但相邻组能一眼分开
# 2026-11-17 评审：原色带对白底 ΔE 只有 ~8、相邻带亮度几乎相同（黑白打印会糊成一片），
# 这里把饱和度提高一档，并让相邻带的亮度拉开 ~10-17 个灰阶，灰度下也能分出组
BAND_TINTS = ["#D6E4F3", "#FBEAD3", "#C7E3C7", "#F2E4EF", "#E3E3E3"]
MEAN_COLOR = "#1a1a1a"

STYLE_ALIASES = {
    "solid": "-", "": "-", "dash": "--", "dashed": "--", "dotdash": "-.",
    "dashdot": "-.", "dotted": ":", "none": "None", "false": "None",
}

# 值列优先级：本项目的三种产物各有一列
VALUE_COLUMNS = OrderedDict([
    ("norm", ("Norm_LOG2FC",)),
    ("ev", ("log2FC (EV-corrected)",)),
    ("raw", ("LOG2FC", "Log2FC", "log2fc")),
])
COLUMN_LABELS = {
    "Norm_LOG2FC": "normalized log2FC (batch-QC)",
    "LOG2FC": "log2FC (bioid, uncorrected)",
    "log2FC (EV-corrected)": "EV-corrected log2FC (bioid-analysis)",
}
YLABEL_DEFAULT = {
    "Norm_LOG2FC": "Normalized Log2FC (norm_log2fc)",
    "LOG2FC": "Log2FC (LOG2FC)",
    "log2FC (EV-corrected)": "Normalized Log2FC (EV-corrected)",
}

GN_KEYS = ("gn", "gene", "genename", "genenames", "genesymbol", "symbol")
ACC_KEYS = ("accession", "proteinid", "uniprot", "uniprotid", "majorityproteinids")


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------
def norm_key(text):
    """表头/字段名归一化：小写、去掉非字母数字。"""
    return re.sub(r"[^a-z0-9]+", "", str(text or "").lower())


def parse_number(value):
    """把单元格变成 float；空、NA、文本都返回 None。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if (isinstance(value, float) and math.isnan(value)) else float(value)
    text = str(value).strip()
    if not text or text.upper() in {"NA", "N/A", "NAN", "NONE", "NULL", "-"}:
        return None
    if text.startswith("<"):
        return None
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return None


def slugify(text, limit=40):
    slug = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text)).strip("_")
    return (slug or "bioid_trend")[:limit].rstrip("_")


def split_symbols(text):
    """GN 可能是 'A;B'、'A, B' 或带空格的复合值。"""
    if not text:
        return []
    parts = re.split(r"[;,/|]", str(text))
    out = []
    for part in parts:
        part = part.strip()
        if part and part.upper() not in {"NA", "NONE"}:
            out.append(part)
    return out


# ---------------------------------------------------------------------------
# 表格读取
# ---------------------------------------------------------------------------
def read_delimited(path):
    """读 TSV/CSV：自动判分隔符，跳过空行与 # 注释。"""
    raw = Path(path).read_text(encoding="utf-8-sig")
    lines = []
    for line in raw.splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        lines.append(line)
    if not lines:
        raise ValueError(f"{path}: 文件为空（或只有注释）")
    delim = "\t" if lines[0].count("\t") >= lines[0].count(",") else ","
    rows = [[cell.strip() for cell in line.split(delim)] for line in lines]
    return rows


def column_map(headers):
    """表头 -> 列号；同名列只认第一次出现（bioid 产物里 GN/LOG2FC 各出现两次）。"""
    mapping = {}
    duplicates = {}
    for i, head in enumerate(headers):
        key = norm_key(head)
        if not key:
            continue
        if key in mapping:
            duplicates.setdefault(key, [mapping[key]]).append(i)
        else:
            mapping[key] = i
    return mapping, duplicates


def _iter_sheet_rows(ws):
    """读整张表；老 openxlsx 产物把 dimension 写成 A1，read-only 模式会只读一列，
    这时 reset_dimensions() 后重读。"""
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    return [list(row) for row in ws.iter_rows(values_only=True)]


class WorkbookTable:
    """一个工作簿（一张表）= 一个 BioID 条件。"""

    def __init__(self, path, sheet, column):
        from openpyxl import load_workbook  # 延迟导入，--inspect 也能跑

        self.path = Path(path)
        wb = load_workbook(self.path, read_only=True, data_only=True)
        try:
            names = list(wb.sheetnames)
            if sheet:
                if sheet not in names:
                    raise ValueError(
                        f"{self.path.name}: 没有工作表 {sheet!r}；可选 {names}")
                ws = wb[sheet]
            elif "Proteins" in names:
                ws = wb["Proteins"]
            else:
                # bioid-analysis 的 summary：一个条件一张表，必须点名，否则会静默画错条件
                candidates = []
                for name in names:
                    if name in ("Overview", "Candidates", "Methods"):
                        continue
                    head = next(wb[name].iter_rows(values_only=True), None)
                    if head and any(norm_key(h) in GN_KEYS for h in head):
                        candidates.append(name)
                if not candidates:
                    raise ValueError(
                        f"{self.path.name}: 找不到带 GN 列的表（工作表 {names}）；"
                        "bioid-analysis 的 summary 请用 sheet 列指定条件名")
                if len(candidates) > 1:
                    raise ValueError(
                        f"{self.path.name}: 有 {len(candidates)} 张条件表 {candidates}，"
                        "请在样本表 sheet 列写明用哪一张（或 --inspect --sheet 先看）")
                ws = wb[candidates[0]]
            self.sheet = ws.title
            rows = _iter_sheet_rows(ws)
        finally:
            wb.close()

        if not rows:
            raise ValueError(f"{self.path.name}/{self.sheet}: 空表")
        header_index = None
        for i, row in enumerate(rows[:10]):
            if any(norm_key(h) in GN_KEYS + ACC_KEYS for h in row):
                header_index = i
                break
        if header_index is None:
            raise ValueError(f"{self.path.name}/{self.sheet}: 前 10 行里没有 GN/Accession 表头")
        headers = [str(h) if h is not None else "" for h in rows[header_index]]
        columns, duplicates = column_map(headers)
        self.headers = headers
        self.duplicates = duplicates

        gn_i = next((columns[k] for k in GN_KEYS if k in columns), None)
        acc_i = next((columns[k] for k in ACC_KEYS if k in columns), None)
        pep_i = columns.get("uniquepeptides")
        if pep_i is None:
            pep_i = columns.get("peptides")

        self.column = self._resolve_column(columns, column)
        value_i = columns[norm_key(self.column)]

        self.records = []
        for offset, row in enumerate(rows[header_index + 1:]):
            if not row or len(row) <= value_i:
                row = list(row) + [None] * (value_i + 1 - len(row))
            gn = str(row[gn_i]).strip() if gn_i is not None and row[gn_i] is not None else ""
            acc = str(row[acc_i]).strip() if acc_i is not None and row[acc_i] is not None else ""
            if not gn and not acc:
                continue
            self.records.append({
                "gn": gn,
                "accession": acc,
                "value": parse_number(row[value_i]),
                "peptides": parse_number(row[pep_i]) if pep_i is not None else None,
                "line": header_index + 2 + offset,
                "raw": row[value_i],
            })
        self._index()

    def _resolve_column(self, columns, requested):
        """`--value auto` 按 Norm -> EV -> LOG2FC 退；指定了就必须存在。"""
        available = [name for names in VALUE_COLUMNS.values() for name in names
                     if norm_key(name) in columns]
        def real_header(index):
            return str(self.headers[index]).strip()

        if requested and requested not in ("auto", "norm", "ev", "raw"):
            if norm_key(requested) in columns:
                return real_header(columns[norm_key(requested)])
            raise ValueError(
                f"{self.path.name}/{self.sheet}: 没有列 {requested!r}；"
                f"可用数值列 {available or '（无）'}")
        modes = ["norm", "ev", "raw"] if requested in (None, "auto") else [requested]
        for mode in modes:
            for name in VALUE_COLUMNS[mode]:
                if norm_key(name) in columns:
                    return real_header(columns[norm_key(name)])
        raise ValueError(
            f"{self.path.name}/{self.sheet}: 找不到 Norm_LOG2FC / LOG2FC / "
            f"log2FC (EV-corrected)；可用数值列 {available or '（无）'}")

    def _index(self):
        self.by_symbol = {}
        self.by_accession = {}
        for rec in self.records:
            for symbol in split_symbols(rec["gn"]):
                self.by_symbol.setdefault(symbol.upper(), []).append(rec)
            if rec["accession"]:
                self.by_accession.setdefault(rec["accession"].upper(), []).append(rec)

    def count_symbols(self):
        return len(self.by_symbol)


# ---------------------------------------------------------------------------
# 别名表（任意列的连通分量：两个方向都能查）
# ---------------------------------------------------------------------------
class AliasMap:
    def __init__(self):
        self.parent = {}

    def load(self, path):
        try:
            rows = read_delimited(path)
        except Exception as exc:
            raise ValueError(f"{path}: 读不了别名表（{exc}）") from exc
        for row in rows:
            cells = [c.strip() for c in row if c.strip()]
            if len(cells) < 2:
                continue
            first = cells[0]
            for other in cells[1:]:
                self._union(first, other)
        return self

    def _find(self, item):
        item = item.upper()
        self.parent.setdefault(item, item)
        root = item
        while self.parent[root] != root:
            root = self.parent[root]
        while self.parent[item] != root:
            self.parent[item], item = root, self.parent[item]
        return root

    def _union(self, a, b):
        ra, rb = self._find(a), self._find(b)
        if ra != rb:
            self.parent[rb] = ra

    def expand(self, symbol):
        root = self._find(symbol)
        return {name for name, parent in self.parent.items() if parent == root or name == root}


# ---------------------------------------------------------------------------
# 样本表 / 蛋白表
# ---------------------------------------------------------------------------
def load_samples(path):
    rows = read_delimited(path)
    header = None
    if rows and norm_key(rows[0][0]) in {"label", "sample", "name", "condition", "x", "bioid"}:
        header = [norm_key(c) for c in rows[0]]
        rows = rows[1:]
    samples = []
    for row in rows:
        if header:
            record = {key: (row[i] if i < len(row) else "") for i, key in enumerate(header)}
            label = record.get("label") or record.get("sample") or record.get("name") or record.get("condition") or ""
            file = record.get("file") or record.get("path") or record.get("workbook") or ""
            sheet = record.get("sheet") or record.get("worksheet") or record.get("condition") or ""
            column = record.get("column") or record.get("value") or record.get("valuecol") or ""
            xgroup = (record.get("xgroup") or record.get("block") or record.get("section")
                      or record.get("xclass") or "")
        else:
            cells = list(row) + [""] * 5
            label, file, sheet, column, xgroup = (cells[0], cells[1], cells[2],
                                                  cells[3], cells[4])
        if not label and not file:
            continue
        if not file:
            raise ValueError(f"{path}: 样本 {label!r} 没有 file 列")
        samples.append({"label": label or Path(file).stem, "file": file,
                        "sheet": sheet or None, "column": column or None,
                        "xgroup": (xgroup or "").strip()})
    if not samples:
        raise ValueError(f"{path}: 没有读到样本行")
    return samples


def load_proteins(path=None, genes=None):
    specs = []
    if path:
        rows = read_delimited(path)
        header = None
        if rows and norm_key(rows[0][0]) in {"gene", "genes", "protein", "symbol", "proteinid"}:
            header = [norm_key(c) for c in rows[0]]
            rows = rows[1:]
        for row in rows:
            if header:
                rec = {key: (row[i] if i < len(row) else "") for i, key in enumerate(header)}
                gene = rec.get("gene") or rec.get("genes") or rec.get("protein") or rec.get("symbol") or ""
                spec = {
                    "gene": gene.strip(),
                    "group": (rec.get("group") or rec.get("class") or rec.get("family") or "").strip(),
                    "label": (rec.get("label") or rec.get("display") or "").strip(),
                    "style": (rec.get("style") or rec.get("linestyle") or "").strip(),
                    "color": (rec.get("color") or rec.get("colour") or "").strip(),
                    "accession": (rec.get("accession") or rec.get("uniprot") or "").strip(),
                }
            else:
                cells = list(row) + [""] * 5
                spec = {"gene": cells[0].strip(), "group": cells[1].strip(),
                        "label": cells[2].strip(), "style": cells[3].strip(),
                        "color": cells[4].strip(), "accession": ""}
            if spec["gene"]:
                specs.append(spec)
    if genes:
        for gene in re.split(r"[,\s]+", genes):
            if gene.strip():
                specs.append({"gene": gene.strip(), "group": "", "label": "",
                              "style": "", "color": "", "accession": ""})
    seen = set()
    unique = []
    for spec in specs:
        key = (spec["gene"].upper(), spec["accession"].upper())
        if key in seen:
            continue
        seen.add(key)
        unique.append(spec)
    if not unique:
        raise ValueError("蛋白表为空")
    return unique


# ---------------------------------------------------------------------------
# 取数
# ---------------------------------------------------------------------------
def pick_record(records):
    """同基因多行（不同 accession / isoform）时：unique peptides 多者优先，再取数值最大者。"""
    if not records:
        return None, []
    ordered = sorted(
        records,
        key=lambda rec: (
            -1 if rec["peptides"] is None else rec["peptides"],
            -math.inf if rec["value"] is None else rec["value"],
        ),
        reverse=True,
    )
    return ordered[0], ordered[1:]


def lookup(table, spec, aliases):
    """返回 (record|None, how, alternatives)。先 accession，再 exact symbol，再别名。"""
    if spec["accession"]:
        hits = table.by_accession.get(spec["accession"].upper())
        if hits:
            rec, alts = pick_record(hits)
            return rec, "accession", alts
        return None, "accession-miss", []

    names = [spec["gene"]] + sorted(aliases.expand(spec["gene"]) - {spec["gene"].upper()})
    for rank, name in enumerate(names):
        hits = table.by_symbol.get(name.upper())
        if not hits:
            continue
        rec, alts = pick_record(hits)
        how = "symbol" if rank == 0 else f"alias:{name}"
        return rec, how, alts
    return None, "absent", []


def build_series(samples, proteins, args, aliases, base_dir):
    x_labels, tables, series = [], [], []
    warnings = []
    for sample in samples:
        file_path = Path(sample["file"])
        if not file_path.is_absolute():
            file_path = (Path(base_dir) / file_path).resolve()
        table = WorkbookTable(file_path, sample["sheet"], args.value if not sample["column"] else sample["column"])
        tables.append(table)
        sample["resolved_file"] = str(file_path)
        sample["sheet_used"] = table.sheet
        sample["column_used"] = table.column
        x_labels.append(sample["label"])

    for spec in proteins:
        values, na_flags, hits = [], [], []
        for table in tables:
            rec, how, alternatives = lookup(table, spec, aliases)
            if rec is None:
                values.append(0.0)
                na_flags.append(True)
                hits.append(f"{table.path.stem}:{how}")
                continue
            if rec["value"] is None:
                values.append(0.0)
                na_flags.append(True)
                hits.append(f"{table.path.stem}:no_value(line {rec['line']}, raw={rec['raw']!r})")
                continue
            values.append(rec["value"])
            na_flags.append(False)
            hits.append(f"{table.path.stem}:{how}")
            if alternatives:
                desc = "; ".join(
                    f"acc={alt['accession'] or '?'} value={alt['raw']!r} pep={alt['peptides']}"
                    for alt in alternatives)
                warnings.append(
                    f"{spec['gene']} 在 {table.path.name} 有 {len(alternatives) + 1} 行匹配，"
                    f"取了 acc={rec['accession'] or '?'} value={rec['raw']!r}；其余 {desc}")
        series.append({
            "gene": spec["gene"],
            "group": spec["group"],
            "label": spec["label"] or spec["gene"],
            "style": spec["style"],
            "color": spec["color"],
            "values": values,
            "na": na_flags,
            "hits": hits,
        })
    return x_labels, tables, series, warnings


# ---------------------------------------------------------------------------
# 样式分配：颜色按蛋白顺序；marker 按 group 首次出现；线型按组内序号
# ---------------------------------------------------------------------------
def assign_styles(series, args):
    groups = OrderedDict()
    for item in series:
        groups.setdefault(item["group"] or "", 0)
    group_seq = list(groups)
    for item in series:
        group = item["group"] or ""
        item["marker"] = MARKERS[group_seq.index(group) % len(MARKERS)]
    for i, item in enumerate(series):
        item["color"] = item["color"] or PALETTE[i % len(PALETTE)]
    for item in series:
        requested = STYLE_ALIASES.get(item["style"].lower(), item["style"])
        item["linestyle"] = requested if item["style"] else "-"
    return group_seq


def segment_runs(na):
    """把一条折线切成连续段 [(i, j, dashed)]。

    dashed=True 表示这一段里至少有一个端点未检出（那一端的值是按 0 画的位置，
    不是测量值）；其余段是实线。段的端点共享，画出来不会断开。
    """
    n = len(na)
    runs = []
    i = 0
    while i < n - 1:
        dashed = bool(na[i] or na[i + 1])
        j = i
        while j < n - 1 and bool(na[j] or na[j + 1]) == dashed:
            j += 1
        runs.append((i, j, dashed))
        i = j
    return runs


def xgroup_runs(x_groups):
    """把 x 轴上的分组名切成连续段 [(start, end, name)]，空的格子不画背景。"""
    runs = []
    index = 0
    total = len(x_groups)
    while index < total:
        name = (x_groups[index] or "").strip()
        if not name:
            index += 1
            continue
        end = index
        while end + 1 < total and (x_groups[end + 1] or "").strip() == name:
            end += 1
        runs.append((index, end, name))
        index = end + 1
    return runs


def compute_means(series, n_points):
    """每个 x 上跨蛋白求平均；values 里 NA 已经是画图的 0，所以未检出按 0 参与。"""
    if not series:
        return [0.0] * n_points
    return [sum(item["values"][i] for item in series) / len(series) for i in range(n_points)]


def legend_label(item):
    if item["group"]:
        return f"{item['group']}: {item['label']}"
    return item["label"]


# ---------------------------------------------------------------------------
# 打印
# ---------------------------------------------------------------------------
def batch_dir(path):
    """工作簿所属批次：向上找 BioID-<n> 这一层，找不到就退回父目录。"""
    for parent in path.resolve().parents:
        if re.fullmatch(r"BioID[-_ ]?\d+[a-z]*", parent.name, flags=re.IGNORECASE):
            return parent
    return path.parent


def describe_inputs(samples, tables):
    lines = []
    sources = set()
    for i, (sample, table) in enumerate(zip(samples, tables), 1):
        sources.add(str(batch_dir(table.path)))
        group = f"  《{sample['xgroup']}》" if sample.get("xgroup") else ""
        lines.append(
            f"  {i:>2}. {sample['label']:<12} {table.path.name}  "
            f"[{table.sheet}]  列={table.column}  蛋白={len(table.records)}{group}")
    return lines, sources


def print_matrix(series, x_labels, means=None):
    width = max([len(item["label"]) for item in series] + ([4] if means is not None else []))
    col = max(10, max(len(label) for label in x_labels) + 1)
    header = " " * (width + 2) + "".join(f"{label:>{col}}" for label in x_labels)
    print(header)
    for item in series:
        cells = []
        for value, na in zip(item["values"], item["na"]):
            text = "▲0" if na else f"{value:.2f}"
            cells.append(f"{text:>{col}}")
        print(f"  {item['label']:<{width}}" + "".join(cells))
    if means is not None:
        cells = [f"{value:{col}.2f}" for value in means]
        print(f"  {'MEAN':<{width}}" + "".join(cells))
    print("  ▲ = 未检出（画在 0，空心点）" + ("；MEAN = 跨蛋白平均（未检出按 0 算）" if means is not None else ""))


def collect_warnings(series, samples, tables, sources):
    warnings = []
    if len(sources) > 1:
        warnings.append(
            f"样本来自 {len(sources)} 个批次目录：" + "、".join(sorted(Path(s).name for s in sources))
            + "。Norm_LOG2FC 是批内中心化，跨批次的绝对高度不可直接比较，"
              "只能读同一蛋白在各构造间的相对变化。")
    columns = {table.column for table in tables}
    if len(columns) > 1:
        warnings.append("各工作簿用的值列不一致：" + "、".join(sorted(columns))
                        + "；图注里会写明，跨列的高度差要谨慎。")
    for item in series:
        if any(item["na"]):
            miss = [label for label, na in zip([s["label"] for s in samples], item["na"]) if na]
            warnings.append(f"{item['label']} 在 {', '.join(miss)} 未检出（画在 0）")
    return warnings


# ---------------------------------------------------------------------------
# 出图
# ---------------------------------------------------------------------------
def build_caption(args, series, samples, tables, x_labels, dropped=(), x_groups=None, means=None):
    if args.caption_text:
        return args.caption_text
    if args.no_caption:
        return None
    columns = sorted({table.column for table in tables})
    groups = OrderedDict()
    for item in series:
        if item["group"]:
            groups.setdefault(item["group"], []).append(item["label"])
    bits = ["Figure legend:"]
    x_desc = f"X axis: {args.xlabel} ({', '.join(x_labels)})."
    bits.append(x_desc)
    bands = [] if args.no_xgroup_bands else xgroup_runs(x_groups or [])
    if bands:
        bits.append("Background bands group the x axis: "
                    + "; ".join(f"{name} ({', '.join(x_labels[start:end + 1])})"
                                for start, end, name in bands) + ".")
    if len(columns) == 1:
        bits.append(f"Y axis: {COLUMN_LABELS.get(columns[0], columns[0])}.")
    else:
        bits.append("Y axis: normalized log2FC (mixed columns: "
                    + ", ".join(COLUMN_LABELS.get(c, c) for c in columns) + ").")
    if groups:
        bits.append("Marker shape shows the group ("
                    + "; ".join(f"{group}: {'/'.join(genes)}" for group, genes in groups.items())
                    + ").")
    else:
        bits.append("Colour identifies the protein.")
    na_total = sum(sum(1 for flag in item["na"] if flag) for item in series)
    if na_total:
        missing = []
        for item in series:
            miss = [label for label, flag in zip(x_labels, item["na"]) if flag]
            if miss:
                missing.append(f"{item['label']} ({', '.join(miss)})")
        bits.append("Lines are solid except a segment that touches a not-detected point, "
                    "which is dashed.")
        bits.append(f"Hollow markers are proteins not detected or not quantified in that BioID "
                    f"({'; '.join(missing)}); they are drawn at 0 (a plotting position, not a "
                    f"measurement) and joined by dashed segments.")
    if len({table.path.parent.name for table in tables}) > 1:
        bits.append("Values are normalized within each BioID batch; only relative "
                    "changes within a protein should be compared.")
    if means is not None:
        bits.append(f"The thick dark line is the mean across the {len(series)} plotted proteins "
                    "at each BioID; a protein not detected there counts as the plotted 0, "
                    "so the mean mixes abundance with detection.")
    if dropped:
        bits.append("Not plotted because they were absent from every BioID: "
                    + ", ".join(item["label"] for item in dropped) + ".")
    text = " ".join(bits)
    return textwrap.fill(text, width=max(70, int(args.fig_width * 18)))


def choose_cjk_font(texts, family):
    """标题/标签/图例里一旦出现非 ASCII，PDF backend 不会逐字回退，
    必须把一款能覆盖中文的字体提到最前面，否则中文会渲染成空白。"""
    joined = " ".join(str(text) for text in texts if text)
    if not any(ord(char) > 127 for char in joined):
        return None
    try:
        from matplotlib import font_manager
        available = {font.name for font in font_manager.fontManager.ttflist}
    except Exception:
        return None
    for candidate in ("Arial Unicode MS", "PingFang HK", "Hiragino Sans GB",
                      "STHeiti", "Heiti TC", "Songti SC"):
        if candidate in available:
            family.insert(0, candidate)
            return f"图中有中文，正文字体改用 {candidate}（PDF 需要单字体覆盖中文）"
    return ("图中有中文，但没找到可用的中文字体；中文可能显示为空白。"
            "请安装 Arial Unicode MS / PingFang，或把标签改成英文")


def auto_rotate(x_labels, tick_size, args):
    """刻度文字按约 0.5em/字符估宽；放不下就转 25°，用户给了数字则照用户的。"""
    requested = str(args.rotate_labels).strip().lower()
    if requested not in ("", "auto", "none", "0"):
        return float(requested)
    if requested in ("none", "0"):
        return 0.0
    axes_width_pt = args.fig_width * (0.70 - 0.115) * 72
    slot = axes_width_pt / max(1, len(x_labels))
    widest = max(len(label) * tick_size * 0.62 for label in x_labels)
    return 25.0 if widest > slot * 0.85 else 0.0


def render(args, x_labels, x_groups, series, caption, means, out_pdf, out_png):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D

    family = ["DejaVu Sans", "Helvetica", "Arial"]
    font_note = choose_cjk_font([args.title, args.xlabel, args.ylabel, caption] + x_labels
                                + [item["label"] for item in series], family)
    matplotlib.rcParams.update({
        "pdf.fonttype": 42,          # TrueType 内嵌，Illustrator/AI 里文字可编辑
        "ps.fonttype": 42,
        "font.family": "sans-serif",
        "font.sans-serif": family,
        "axes.unicode_minus": False,
        "svg.fonttype": "none",
    })
    if font_note:
        print(f"注意: {font_note}")

    fig_size = (args.fig_width, args.fig_height)
    fig, ax = plt.subplots(figsize=fig_size)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.set_axisbelow(True)
    ax.grid(True, which="major", axis="both", color="#dcdcdc", lw=0.6)
    ax.set_xlim(-0.35, len(x_labels) - 0.65)
    ax.set_xticks(range(len(x_labels)))
    tick_size = 9.5
    ax.set_xticklabels(x_labels, fontsize=tick_size)
    bands = [] if args.no_xgroup_bands else xgroup_runs(x_groups or [])
    band_index = {}
    for start, end, name in bands:
        band_index.setdefault(name, len(band_index))
        ax.axvspan(start - 0.5, end + 0.5, color=BAND_TINTS[band_index[name] % len(BAND_TINTS)],
                   zorder=0, lw=0)
    for start, end, name in bands[:-1]:
        ax.axvline(end + 0.5, color="white", lw=1.4, zorder=0.6)
    rotate = auto_rotate(x_labels, tick_size, args)
    if rotate:
        plt.setp(ax.get_xticklabels(), rotation=rotate, ha="right", rotation_mode="anchor")

    handles = []
    x = list(range(len(x_labels)))
    # 1) 折线先画，避免后画的线压住先画点的 marker；
    #    有端点未检出的那一段改虚线，其余实线（线型不再承担 paralog 的含义）
    for item in series:
        base = item["linestyle"]
        for start, end, dashed in segment_runs(item["na"]):
            ax.plot(x[start:end + 1], item["values"][start:end + 1],
                    linestyle="--" if dashed else base, color=item["color"],
                    lw=1.9, zorder=2, solid_capstyle="round", dash_capstyle="round")
    # 2) 实测点：白边让两个贴在一起的点仍能看出各自的轮廓
    for item in series:
        measured = [i for i, na in enumerate(item["na"]) if not na]
        if measured:
            ax.plot([x[i] for i in measured], [item["values"][i] for i in measured],
                    linestyle="none", marker=item["marker"], ms=5.6,
                    mfc=item["color"], mec="white", mew=0.9, zorder=4)
    # 3) 未检出的空心点最后画，永远在最上层（它们是需要被看见的“0 是位置”标记）。
    #    同一个 x 上有多个未检出时，按出现顺序递增 marker 尺寸叠成同心环：
    #    位置不动（y 仍是 0），但每个轮廓都看得见，不会看起来只有一个点。
    na_calls = []
    for item in series:
        for i, na in enumerate(item["na"]):
            if na:
                na_calls.append({"x": x[i], "marker": item["marker"],
                                 "color": item["color"], "ms": 6.4})
    stacks = {}
    for call in na_calls:
        stacks.setdefault(call["x"], []).append(call)
    for group in stacks.values():
        # 只有一个未检出时用和实测点一样的尺寸（避免"缺的点看起来更大"这种假语义）；
        # 同一位置叠了多个才逐级放大，让轮廓都能看见
        for rank, call in enumerate(group):
            call["ms"] = 5.6 if len(group) == 1 else 6.2 + 1.2 * rank
    for call in sorted(na_calls, key=lambda item: -item["ms"]):
        ax.plot([call["x"]], [0.0], linestyle="none", marker=call["marker"],
                ms=call["ms"], mfc="white", mec=call["color"], mew=1.25, zorder=3)
    for item in series:
        handles.append(Line2D([0], [0], color=item["color"], linestyle=item["linestyle"], lw=1.9,
                              marker=item["marker"], ms=5.6, mfc=item["color"], mec="white",
                              mew=0.9, label=legend_label(item)))

    if means is not None:
        ax.plot(x, means, color=MEAN_COLOR, lw=2.6, zorder=2.5, solid_capstyle="round")
        handles.append(Line2D([0], [0], color=MEAN_COLOR, lw=2.6,
                              label=f"mean (n={len(series)})"))

    values = [v for item in series for v, na in zip(item["values"], item["na"]) if not na]
    low = min([0.0] + values)
    high = max([0.0] + values)
    span = max(high - low, 1.0)
    ylim = (low - 0.08 * span, high + 0.10 * span)
    if args.ylim:
        ylim = tuple(float(v) for v in args.ylim.split(","))
    ax.set_ylim(*ylim)
    if args.yticks:
        ax.set_yticks([float(v) for v in args.yticks.split(",")])

    rotated_band = False
    if bands and not args.no_xgroup_labels:
        axes_width_pt = args.fig_width * (0.70 - 0.115) * 72
        slot_pt = axes_width_pt / max(1, len(x_labels))
        for start, end, name in bands:
            width_pt = slot_pt * (end - start + 1)
            if len(name) * 8.8 * 0.55 > width_pt * 0.95:
                ax.text((start + end) / 2.0, 1.012, name, transform=ax.get_xaxis_transform(),
                        ha="center", va="bottom", fontsize=8.4, color="#555555",
                        rotation=25, rotation_mode="anchor")
                rotated_band = True
            else:
                ax.text((start + end) / 2.0, 1.012, name, transform=ax.get_xaxis_transform(),
                        ha="center", va="bottom", fontsize=8.8, color="#555555")
    if args.title:
        pad = 9
        if bands and not args.no_xgroup_labels:
            pad = 34 if rotated_band else 26
        ax.set_title(args.title, loc="left", fontsize=11.5, pad=pad)
    rotate_now = auto_rotate(x_labels, 9.5, args)
    if rotate_now:
        # 旋转刻度时 matplotlib 不会把 xlabel 自动推到旋转后的文字下方，
        # 字一长就会压到 xlabel 上：按最长标签的竖向投影给 labelpad 和底边距。
        extent_pt = max(len(label) for label in x_labels) * 9.5 * 0.55 \
            * math.sin(math.radians(rotate_now))
        xlabel_pad = 6 + extent_pt * 0.55
        bottom = min(0.55, max(0.22, (extent_pt * 0.55 + 46) / (args.fig_height * 72)))
    else:
        xlabel_pad = 5
        bottom = 0.145
    ax.set_xlabel(args.xlabel, fontsize=10, labelpad=xlabel_pad)
    ax.set_ylabel(args.ylabel or "Log2FC", fontsize=10)

    for side in ("top", "right", "left", "bottom"):
        ax.spines[side].set_visible(True)
        ax.spines[side].set_linewidth(0.9)
        ax.spines[side].set_color("#333333")
    ax.tick_params(direction="out", length=3, width=0.9, colors="#333333",
                   labelsize=9.5, color="#333333")

    if any(any(item["na"]) for item in series):
        handles.append(Line2D([0], [0], linestyle="none", marker="o", ms=5.6,
                              mfc="white", mec="#444444", mew=1.2,
                              label="not detected (drawn at 0)"))
        handles.append(Line2D([0], [0], linestyle="--", color="#444444", lw=1.9,
                              label="segment joining a not-detected point"))
    legend = ax.legend(handles=handles, title=args.legend_title or None,
                       loc="upper left", bbox_to_anchor=(1.02, 1.0),
                       frameon=True, framealpha=1.0, edgecolor="#cccccc",
                       fontsize=8.6, title_fontsize=9.0, borderpad=0.6,
                       labelspacing=0.45, handlelength=2.2)
    legend.get_frame().set_linewidth(0.7)

    # 固定边距而不是 tight_layout：图例放在坐标轴外，tight_layout 会把绘图区压窄；
    # bottom 由上面按旋转标签的竖向投影算好
    fig.subplots_adjust(left=0.115, right=0.70, bottom=bottom, top=0.905)
    if caption:
        # 旋转刻度时 xlabel 会被 labelpad 推得很低，caption 不能再写死 -0.018，
        # 否则会压在 xlabel 上：画一次，取 xlabel 的实际底边再往下放。
        fig.canvas.draw()
        try:
            renderer = fig.canvas.get_renderer()
            label_bottom = fig.transFigure.inverted().transform(
                (0, ax.xaxis.label.get_window_extent(renderer=renderer).y0))[1]
            caption_y = min(-0.018, label_bottom - 0.02)
        except Exception:
            caption_y = -0.018
        fig.text(0.01, caption_y, caption, ha="left", va="top", fontsize=7.0,
                 color="#333333", linespacing=1.35)

    fig.savefig(out_pdf, format="pdf", bbox_inches="tight", pad_inches=0.12)
    if out_png:
        fig.savefig(out_png, format="png", dpi=args.dpi, bbox_inches="tight",
                    pad_inches=0.12)
    plt.close(fig)


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------
def write_values(path, x_labels, series, samples, tables, dropped=(), means=None):
    dropped_genes = {id(item) for item in dropped}
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("gene\tlabel\tgroup\t" + "\t".join(x_labels)
                     + "\tna_count\tna_samples\tplotted\n")
        for item in list(series) + list(dropped):
            cells = ["NA" if na else f"{value:.6g}" for value, na in zip(item["values"], item["na"])]
            missing = [label for label, na in zip(x_labels, item["na"]) if na]
            handle.write("\t".join([item["gene"], item["label"], item["group"]] + cells
                                   + [str(len(missing)), ",".join(missing),
                                      "no" if id(item) in dropped_genes else "yes"]) + "\n")
        if means is not None:
            na_total = sum(1 for item in series for flag in item["na"] if flag)
            handle.write("\t".join(["MEAN", "mean", ""] + [f"{v:.6g}" for v in means]
                                   + [str(na_total), "", "yes"]) + "\n")


def write_sources(path, samples, tables):
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("label\tfile\tsheet\tcolumn\tn_proteins\n")
        for sample, table in zip(samples, tables):
            handle.write("\t".join([sample["label"], str(table.path), table.sheet,
                                    table.column, str(len(table.records))]) + "\n")


def inspect_file(path, args):
    table = WorkbookTable(path, args.sheet, args.value)
    print(f"文件: {table.path}")
    print(f"工作表: {table.sheet}")
    print(f"数值列: {table.column}")
    print(f"蛋白行: {len(table.records)}  基因数: {table.count_symbols()}")
    header_line = "  ".join(f"[{i}] {h}" for i, h in enumerate(table.headers) if h)
    print("表头: " + header_line)
    if table.duplicates:
        print("同名列（取第一次出现的）: " + ", ".join(sorted(table.duplicates)))
    return 0


# ---------------------------------------------------------------------------
def main(argv=None):
    parser = argparse.ArgumentParser(
        description="把若干蛋白在多个 BioID 工作簿里的标准化 log2FC 画成折线图（PDF+PNG+数值表）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            示例：
              python3 bioid_trend.py --check --samples samples.tsv --proteins proteins.tsv \\
                  --xlabel "ATOSB construct"
              python3 bioid_trend.py --samples samples.tsv --proteins proteins.tsv \\
                  --title "ATOSB constructs: PP2A subunits" --xlabel "ATOSB construct" \\
                  --legend-title "PP2A groups / proteins" --out out/
            """))
    parser.add_argument("--samples", help="样本表：label <TAB> file [sheet] [column]")
    parser.add_argument("--point", action="append", metavar="LABEL=FILE",
                        help="快速指定一个 x 点（可重复，按给定顺序）；与 --samples 可混用")
    parser.add_argument("--proteins", help="蛋白表：gene [group] [label] [style] [color] [accession]")
    parser.add_argument("--genes", help="快速指定蛋白（逗号/空格分隔），不带分组")
    parser.add_argument("--out", help="输出目录（默认 <samples 所在目录>/bioid_trend/）")
    parser.add_argument("--name", help="图名（同时作为 PDF/PNG 前缀与子目录名；默认按蛋白名自动生成）")
    parser.add_argument("--flat", action="store_true",
                        help="旧布局：所有产物平铺在 --out 目录里，而不是每张图一个子目录")
    parser.add_argument("--value", default="auto",
                        help="auto|norm|ev|raw 或具体列名（默认 auto：Norm_LOG2FC -> EV -> LOG2FC）")
    parser.add_argument("--aliases", help="别名表：任意两列，同一连通分量视为同一蛋白")
    parser.add_argument("--title", default="")
    parser.add_argument("--xlabel", default="BioID")
    parser.add_argument("--ylabel", default=None)
    parser.add_argument("--legend-title", default="")
    parser.add_argument("--caption-text", default=None, help="替换整段图注（Latin-1 之外的字符也可）")
    parser.add_argument("--no-caption", action="store_true")
    parser.add_argument("--ylim", help="如 -0.5,9.5")
    parser.add_argument("--yticks", help="如 0,2,4,6,8,10")
    parser.add_argument("--fig-width", type=float, default=6.0)
    parser.add_argument("--fig-height", type=float, default=4.5)
    parser.add_argument("--rotate-labels", default="auto",
                        help="auto（默认，放不下才旋转）或角度数值，0 强制不转")
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--keep-all-na", action="store_true",
                        help="连所有工作簿里都找不到的蛋白也画成贴 0 的空心线（默认跳过并报警）")
    parser.add_argument("--mean", action="store_true",
                        help="加一条跨蛋白平均的粗线（未检出按 0 参与平均）")
    parser.add_argument("--no-xgroup-bands", action="store_true",
                        help="样本表有 xgroup 列也不画背景分组")
    parser.add_argument("--no-xgroup-labels", action="store_true",
                        help="只画背景色，不在轴上方写分组名")
    parser.add_argument("--check", action="store_true", help="只打印将要画的数据，不写文件")
    parser.add_argument("--inspect", help="打印某个工作簿的工作表/表头/将用的数值列")
    parser.add_argument("--sheet", help="配合 --inspect 指定工作表")
    args = parser.parse_args(argv)

    if args.inspect:
        try:
            return inspect_file(args.inspect, args)
        except (ValueError, OSError) as exc:
            print(f"错误: {exc}", file=sys.stderr)
            return 2

    if not args.samples and not args.point:
        parser.error("需要 --samples 或 --point（或用 --inspect 先看文件）")
    if not args.proteins and not args.genes:
        parser.error("需要 --proteins 或 --genes")

    base_dir = Path(args.samples).resolve().parent if args.samples else Path.cwd()
    try:
        aliases = AliasMap().load(args.aliases) if args.aliases else AliasMap()
        samples = load_samples(args.samples) if args.samples else []
        for spec in args.point or []:
            if "=" not in spec:
                raise ValueError(f"--point 要写成 LABEL=FILE，收到 {spec!r}")
            label, file = spec.split("=", 1)
            samples.append({"label": label.strip(), "file": file.strip(),
                            "sheet": None, "column": None, "xgroup": ""})
        if not samples:
            raise ValueError("没有读到样本")
        proteins = load_proteins(args.proteins, args.genes)
        x_labels, tables, series, warnings = build_series(samples, proteins, args, aliases, base_dir)
    except (ValueError, OSError) as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    dropped = []
    if not args.keep_all_na:
        dropped = [item for item in series if all(item["na"])]
        series = [item for item in series if not all(item["na"])]
    if not series:
        print("错误: 没有任何可画的点 —— 指定的蛋白在所有工作簿里都找不到", file=sys.stderr)
        return 2
    groups = assign_styles(series, args)
    x_groups = [sample.get("xgroup", "") for sample in samples]
    means = compute_means(series, len(x_labels)) if args.mean else None
    if args.mean and len(series) == 1:
        warnings.append("--mean 只选了一个蛋白，平均线就是它本身")

    print("=" * 78)
    print(f"bioid-trend  {len(x_labels)} 个 BioID × {len(series)} 个蛋白")
    print("=" * 78)
    input_lines, sources = describe_inputs(samples, tables)
    print("输入（x 轴顺序）：")
    for line in input_lines:
        print(line)
    if groups and any(groups):
        print("分组：" + "; ".join(
            f"{group}: {', '.join(item['label'] for item in series if item['group'] == group)}"
            for group in groups))
    print()
    print("数值矩阵（NA 画在 0）：")
    print_matrix(series + dropped, x_labels, means=means)
    print()
    print("每点来源：")
    for item in series:
        print(f"  {item['label']:<12} " + " | ".join(item["hits"]))
    print()

    for item in dropped:
        warnings.append(f"{item['label']} 在所有工作簿里都没有可用数值，未画进图"
                        "（数值表里仍列出）；先确认基因名，或用 --keep-all-na 强制画")
    all_warnings = collect_warnings(series, samples, tables, sources) + warnings
    if all_warnings:
        print("注意：")
        for note in all_warnings:
            print(f"  ! {note}")
        print()

    if args.check:
        print("--check：没有写任何文件。确认数据和分组后去掉 --check 再跑。")
        return 0

    import matplotlib  # noqa: F401  (提前给出清楚的缺依赖报错)
    out_dir = Path(args.out) if args.out else base_dir / "bioid_trend"
    if args.name:
        stem = args.name
    else:
        stem = "bioid_trend_" + "_".join(slugify(item["gene"], 18) for item in series[:4])
        if len(series) > 4:
            stem += f"_and{len(series) - 4}"
    # 默认每张图一个子目录，图和表格分开：<out>/<stem>/<stem>.pdf|png + <out>/<stem>/data/<stem>_*.tsv
    if args.flat:
        fig_dir = data_dir = out_dir
    else:
        fig_dir = out_dir / stem
        data_dir = fig_dir / "data"
    fig_dir.mkdir(parents=True, exist_ok=True)
    data_dir.mkdir(parents=True, exist_ok=True)
    out_pdf = fig_dir / f"{stem}.pdf"
    out_png = fig_dir / f"{stem}.png"
    out_values = data_dir / f"{stem}_values.tsv"
    out_sources = data_dir / f"{stem}_sources.tsv"

    if args.ylabel is None:
        columns = {table.column for table in tables}
        if len(columns) == 1:
            args.ylabel = YLABEL_DEFAULT.get(next(iter(columns)), "Log2FC")
        else:
            args.ylabel = "Log2FC (mixed columns; see caption)"
    caption = build_caption(args, series, samples, tables, x_labels, dropped,
                            x_groups=x_groups, means=means)
    try:
        render(args, x_labels, x_groups, series, caption, means, out_pdf, out_png)
    except Exception as exc:  # matplotlib 缺失等
        print(f"错误: 出图失败（{type(exc).__name__}: {exc}）", file=sys.stderr)
        return 2
    write_values(out_values, x_labels, series, samples, tables, dropped, means)
    write_sources(out_sources, samples, tables)

    print("已写出：")
    for path in (out_pdf, out_png, out_values, out_sources):
        if path.exists():
            print(f"  {path}  ({path.stat().st_size / 1024:.1f} KB)")
    print("\n交付前请打开 PDF 看一眼：空心点位置、图例分组、图注里的 NA 列表是否都对。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
