#!/usr/bin/env python3
"""bioid_run.py -- run the BioID chain in one command.

预处理（bioid，R）-> 批次 QC（bioid-batch-qc）-> 打分与报告（bioid-analysis），
每一步的日志落在 <out>/logs/，最后做一次验收：条件数是否一致、QC 三列是否接上、
每个条件有没有标到诱饵。任何一步失败就停在那里，退出码说明卡点。

只读输入；产物全部写在 --out 指定的运行目录里。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve()
SKILLS_ROOT = HERE.parents[2]                      # .../skills
R_SCRIPT = HERE.parent / "bioid_process.R"
QC_SCRIPT = Path(os.environ.get(
    "BIOID_QC_SCRIPT", SKILLS_ROOT / "bioid-batch-qc" / "scripts" / "bioid_batch_review.py"))
ANALYSIS_SCRIPT = Path(os.environ.get(
    "BIOID_ANALYSIS_SCRIPT", SKILLS_ROOT / "bioid-analysis" / "scripts" / "bioid_analysis.py"))

EXIT_USAGE = 2
EXIT_CONFIRM = 3
EXIT_STAGE = 4
EXIT_VERIFY = 5


def say(message=""):
    print(message, flush=True)


def workbook_count(folder, suffix="_processed.xlsx"):
    return sorted(p for p in Path(folder).glob(f"*{suffix}")
                  if not p.name.startswith(("~$", ".~")))


def run_stage(label, command, log_path, dry_run=False):
    """跑一个阶段：同时写日志与屏幕；返回 (exit code, 耗时秒)。"""
    say(f"\n[{label}] {' '.join(str(c) for c in command)}")
    if dry_run:
        return 0, 0.0
    log_path.parent.mkdir(parents=True, exist_ok=True)
    started = time.time()
    with open(log_path, "w", encoding="utf-8") as log:
        log.write("# " + " ".join(str(c) for c in command) + "\n")
        process = subprocess.Popen(command, stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT, text=True)
        assert process.stdout is not None
        for line in process.stdout:
            sys.stdout.write(line)
            sys.stdout.flush()
            log.write(line)
        code = process.wait()
    elapsed = time.time() - started
    say(f"[{label}] exit={code}  {elapsed:.1f}s  log={log_path}")
    return code, elapsed


def check_python(python):
    probe = subprocess.run([python, "-c", "import numpy, openpyxl, sys; print(sys.version.split()[0])"],
                           capture_output=True, text=True)
    if probe.returncode != 0:
        raise SystemExit(
            f"--python {python} 缺少 numpy/openpyxl：{probe.stderr.strip()[:200]}\n"
            "  换一个解释器（例如 python3），或先 pip install numpy openpyxl")
    return probe.stdout.strip()


def preflight(args):
    input_dir = Path(args.input).expanduser().resolve()
    out_dir = Path(args.out).expanduser().resolve()
    if not input_dir.is_dir():
        raise SystemExit(f"输入目录不存在：{input_dir}")
    if out_dir == input_dir:
        raise SystemExit(f"--out 不能就是输入目录（{out_dir}）；建一个子目录，例如 <batch>/analysis")
    # 允许 --out 是输入目录的子目录（<batch>/analysis 是最顺手的布局）：R 的 list.files 和
    # 草稿推断的 glob 都是非递归的，看不到子目录里的产物。
    sources = [p for p in input_dir.iterdir()
               if p.suffix.lower() == ".xlsx" and not p.name.startswith(("~$", ".~"))
               and not p.name.endswith("_flagged.xlsx")]
    if not sources and not args.skip_preprocess:
        raise SystemExit(f"{input_dir} 里没有 .xlsx 工作簿")
    if not args.skip_preprocess:
        if shutil.which(args.rscript) is None:
            raise SystemExit(f"找不到 Rscript（{args.rscript}）；预处理阶段需要 R + openxlsx/readxl")
        if args.channel_map and not Path(args.channel_map).expanduser().exists():
            raise SystemExit(f"--channel-map 指向的文件不存在：{args.channel_map}")
    if not QC_SCRIPT.exists():
        raise SystemExit(f"找不到 batch-QC 脚本：{QC_SCRIPT}（可用 BIOID_QC_SCRIPT 指定）")
    if not ANALYSIS_SCRIPT.exists():
        raise SystemExit(f"找不到 analysis 脚本：{ANALYSIS_SCRIPT}（可用 BIOID_ANALYSIS_SCRIPT 指定）")
    version = check_python(args.python)
    say(f"input: {input_dir}")
    say(f"out:   {out_dir}")
    say(f"python: {args.python} ({version})   Rscript: {args.rscript}")
    return input_dir, out_dir, sources


def stage_done(paths):
    return all(Path(p).exists() for p in paths)


def digest(*parts):
    import hashlib
    handle = hashlib.sha256()
    for part in parts:
        handle.update(str(part).encode("utf-8"))
    return handle.hexdigest()[:16]


ABUNDANCE_RE = re.compile(r"abundance[:\s]", re.IGNORECASE)
F_TOKEN_RE = re.compile(r"\bF(\d+)\b")


def split_label(stem):
    """文件名 -> 诱饵标签：去掉 _processed/_flagged 与开头的序号。"""
    stem = re.sub(r"(_processed|_flagged)+$", "", stem, flags=re.IGNORECASE)
    return re.sub(r"^[0-9-]+", "", stem) or stem


def _to_num(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def probe_workbook(path, sheet=None):
    """读一份原始导出：丰度通道、每个 GN 在各通道的丰度。读不了返回 None。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = wb.sheetnames
        if sheet and sheet in names:
            ws = wb[sheet]
        elif "Proteins" in names:
            ws = wb["Proteins"]
        else:
            ws = wb[names[0]]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        rows = ws.iter_rows(values_only=True)
        header = None
        for _ in range(6):
            row = next(rows, None)
            if row is None:
                break
            if any(isinstance(cell, str) and ABUNDANCE_RE.match(cell.strip().lower())
                   for cell in row):
                header = row
                break
        if header is None:
            return None
        channels = []
        for index, cell in enumerate(header):
            if not isinstance(cell, str):
                continue
            text = cell.strip().lower()
            if not ABUNDANCE_RE.match(text) or "normalized" in text:
                continue
            match = F_TOKEN_RE.search(cell)
            token = f"F{match.group(1)}" if match else cell.strip()
            channels.append((token, index, cell.strip()))
        if len(channels) < 2:
            return None
        gn_index = None
        desc_index = None
        for index, cell in enumerate(header):
            if not isinstance(cell, str):
                continue
            name = re.sub(r"[^a-z0-9]", "", cell.lower())
            if name in ("gn", "gene", "genename") and gn_index is None:
                gn_index = index
            elif name == "description" and desc_index is None:
                desc_index = index
        data = {}
        for row in rows:
            if row is None:
                continue
            gn = ""
            if gn_index is not None and gn_index < len(row) and row[gn_index] is not None:
                gn = str(row[gn_index]).strip()
            if not gn and desc_index is not None and desc_index < len(row):
                # 公司导出通常没有 GN 列，只能从 Description 的 GN=... 取（与 bioid_process.R 一致）
                match = re.search(r"GN=([^ ;]+)", str(row[desc_index] or ""))
                if match:
                    gn = match.group(1).strip()
            if not gn:
                continue
            values = {}
            for token, index, _ in channels:
                if index < len(row):
                    values[token] = _to_num(row[index])
            if any(value is not None and value > 0 for value in values.values()):
                data.setdefault(gn.upper(), values)
        return {"path": path, "channels": [(token, header) for token, _, header in channels],
                "rows": data}
    except Exception:
        return None
    finally:
        wb.close()


def infer_channel_map(input_dir, sheet=None):
    """从原始工作簿推断 <诱饵通道> <EV通道> <文件名>，每条带置信度和证据。

    这只是草稿：真正的判断（哪个通道是诱饵）仍要人确认。规则：
    1) 文件名里的标签在工作簿 GN 里精确命中 -> 该蛋白丰度最高的通道是诱饵通道（high）；
    2) 只命中前缀（如 ATOSB-N vs ATOSB）-> 同上（medium）；
    3) 都不命中 -> 取"出现文件数最少"的通道，再比该通道的最大丰度（low）。
    """
    files = sorted(
        path for path in Path(input_dir).glob("*.xlsx")
        if not path.name.startswith(("~$", ".~"))
        and not re.search(r"(_processed|_flagged)$", path.stem, re.IGNORECASE))
    probes = [(path, probe_workbook(path, sheet)) for path in files]
    probes = [(path, probe) for path, probe in probes if probe]
    probes = [(path, probe) for path, probe in probes if probe["rows"]]
    if not probes:
        return []
    file_count = {}
    for _, probe in probes:
        for token, _ in probe["channels"]:
            file_count[token] = file_count.get(token, 0) + 1

    records = []
    for path, probe in probes:
        label = split_label(path.stem)
        tokens = [token for token, _ in probe["channels"]]
        matched, confidence, note = None, "low", ""
        if label.upper() in probe["rows"]:
            matched, confidence = label.upper(), "high"
        else:
            candidates = [gn for gn in probe["rows"]
                          if gn and (label.upper().startswith(gn) or gn.startswith(label.upper()))]
            if candidates:
                matched = sorted(candidates, key=len)[0]
                confidence = "medium"
                note = f"标签 {label!r} 未精确匹配，按前缀取 {matched}"
        shared = {token for token, count in file_count.items()
                  if count > 1}                        # 跨文件重复出现：像共用的对照通道
        if matched:
            values = probe["rows"][matched]
            best = max(tokens, key=lambda token: (values.get(token) or 0))
            if best in shared and any(token not in shared for token in tokens):
                # 标签蛋白在共用通道里更高：多半是诱饵/对照列反了，按"独有通道"取并告警
                unique = [token for token in tokens if token not in shared]
                bait = max(unique, key=lambda token: (values.get(token) or 0))
                confidence = "low"
                note = (f"标签 {label!r} 在 {best} 更高，但 {best} 出现在 "
                        f"{file_count[best]} 个文件里（像共用对照）；按独有通道 {bait} 取，"
                        "请确认是不是诱饵/对照列反了")
            else:
                bait = best
        else:
            bait = min(tokens, key=lambda token: (
                file_count.get(token, 0),
                -max((probe["rows"][gn].get(token) or 0) for gn in probe["rows"])),)
            confidence = "low"
            note = "文件名里没有可用的基因名，按“通道出现文件数最少”推断"

        evs = [token for token in tokens if token != bait]
        evidence_gn = matched
        if evidence_gn is None:
            evidence_gn = max(
                probe["rows"],
                key=lambda gn: (probe["rows"][gn].get(bait) or 0))
        values = probe["rows"].get(evidence_gn, {})
        ratio_bits = []
        for ev in evs:
            top = values.get(bait)
            ctrl = values.get(ev)
            if top and ctrl and ctrl > 0:
                ratio_bits.append(f"{top / ctrl:.0f}x vs {ev}")
            elif top and not ctrl:
                ratio_bits.append(f"{ev} 未检出")
            else:
                ratio_bits.append(f"{ev} 无数据")
        records.append({
            "file": path.name, "bait": bait, "evs": evs, "label": label,
            "matched": matched, "confidence": confidence, "note": note,
            "evidence_gn": evidence_gn,
            "evidence": f"{evidence_gn} " + "; ".join(ratio_bits),
        })
    return records


def write_channel_map(path, records):
    lines = [
        "# bioid_run.py 自动推断的通道映射（草稿）。格式：<诱饵通道>  <EV通道>[|EV2...]  <文件名>",
        "# 每行上方的注释是推断依据。可以直接编辑本文件；确认后加 --yes 重跑，编辑内容会被沿用。",
    ]
    for record in records:
        lines.append(
            f"# {record['file']}  置信={record['confidence']}  "
            f"标签={record['label']}  证据={record['evidence']}"
            + (f"  ({record['note']})" if record["note"] else ""))
        lines.append("\t".join([record["bait"], "|".join(record["evs"]), record["file"]]))
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def proposed_bait_aliases(records):
    """文件名标签不是精确基因名时，给出 <标签> <基因> 的候选映射。"""
    pairs = []
    for record in records:
        gene = (record.get("matched") or "").strip()
        if not gene:
            continue
        if gene.upper() != str(record["label"]).upper():
            pairs.append((record["label"], gene))
    # 同一标签只留一条
    seen, unique = set(), []
    for label, gene in pairs:
        if label.upper() in seen:
            continue
        seen.add(label.upper())
        unique.append((label, gene))
    return unique


def write_bait_alias_draft(path, pairs):
    lines = [
        "# 构造/旧名 -> 数据库符号（bioid_run.py 推断；确认或修改后随 --yes 使用）",
        "# 每行：<标签> <基因>。bioid-batch-qc 与 bioid-analysis 都会收到这张表。",
    ]
    lines += [f"{label}\t{gene}" for label, gene in pairs]
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def print_alias_gate(pairs, alias_path, written=True):
    say("\n需要你确认的 构造名 → 基因名（这些文件名不是标准基因符号）：")
    for label, gene in pairs:
        say(f"  {label:<34} -> {gene}")
    if written:
        say(f"\n已写到 {alias_path}：确认后加 --yes 会直接用它（QC 的诱饵回收/转运与分析都会收到）；"
            "要改就编辑这个文件，每行 <标签> <基因>。")
    else:
        say(f"\n（--dry-run：没有写盘；去掉 --dry-run 会写到 {alias_path}）")


def print_channel_map_gate(records, map_path, written=True):
    say("\n推断的通道映射（草稿，尚未运行）：")
    say(f"  {'文件':<32}{'诱饵':<5}{'EV':<8}{'置信':<7}证据（诱饵通道里该蛋白的丰度 / 与对照比）")
    for record in records:
        say(f"  {record['file'][:31]:<32}{record['bait']:<5}"
            f"{'|'.join(record['evs']):<8}{record['confidence']:<7}{record['evidence']}")
    weak = [record for record in records if record["confidence"] != "high"]
    if weak:
        say("\n需要留意：")
        for record in weak:
            reason = record["note"] or "标签没有精确命中工作簿里的 GN"
            say(f"  {record['file']}: {reason}；证据 {record['evidence']}")
        say("  如果标签是构造名/旧名：下面的别名草稿（若有）会给出 标签→基因 的候选，"
            "确认或修改后再用 --yes；也可以用 --bait-alias 指定自己的表。")
    if written:
        say(f"\n草稿已写到 {map_path}（可直接编辑；确认无误后加 --yes 重跑，编辑内容会被沿用）。")
    else:
        say("\n（--dry-run：没有写盘；去掉 --dry-run 会把它写到 "
            f"{map_path}）。")


def read_payload(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return None


def verify(out_dir, processed_dir, n_expected, args):
    """三件事：条件数一致、QC 接上、诱饵标到。返回 (ok, 摘要行列表)。"""
    lines, ok = [], True
    processed = workbook_count(processed_dir)
    flagged = workbook_count(out_dir / "02_qc", "_flagged.xlsx")
    lines.append(f"processed={len(processed)} flagged={len(flagged)} expected={n_expected}")
    if len(flagged) != len(processed):
        ok = False
        lines.append(f"  FAIL flagged 数与 processed 不一致（{len(flagged)} vs {len(processed)}）")
    if n_expected and len(processed) < n_expected:
        lines.append(f"  WARN processed 少于输入（{len(processed)} vs {n_expected}）"
                     "——可能有工作簿被跳过，看 logs/01_preprocess.log")
    elif n_expected and len(processed) > n_expected:
        lines.append(f"  INFO processed 比输入多（{len(processed)} vs {n_expected}）——"
                     "一个诱饵配多个 EV 时会这样，正常")

    payload = read_payload(out_dir / "03_analysis" / "analysis_payload.json")
    if payload is None:
        ok = False
        lines.append("  FAIL 没有 analysis_payload.json，分析阶段没跑完")
        return ok, lines
    if not payload.get("qcAttached"):
        ok = False
        lines.append("  FAIL 报告里 QC 未接入（qcAttached=false），Background flag 不会出现")
    fields = payload.get("pointFields") or []
    conditions = payload.get("conditions") or []
    lines.append(f"analysis conditions={len(conditions)} qcAttached={payload.get('qcAttached')}")
    if len(conditions) != len(flagged):
        ok = False
        lines.append(f"  FAIL 分析只覆盖 {len(conditions)} 个条件，但 flagged 有 {len(flagged)} 份"
                     "——有工作簿没进分析，看 logs/04_analysis.log")
    no_bait = []
    for condition in conditions:
        points = [dict(zip(fields, row)) for row in condition.get("points", [])]
        if not any(point.get("bait_hit") for point in points):
            no_bait.append(condition.get("condition") or condition.get("file"))
    if no_bait:
        ok = False
        lines.append("  FAIL 这些条件没有标到诱饵：" + ", ".join(map(str, no_bait))
                     + "（用 --list-baits 核对文件名，或加 --bait-map）")
    return ok, lines


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=("一条命令跑完 BioID 预处理 -> 批次 QC -> 打分报告，并在最后验收。\n"
                     "没有 channel_map 时会自动推断草稿（<out>/00_channel_map.tsv，可直接编辑），"
                     "与诱饵确认一起停下等你 --yes。"),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
示例：
  Rscript 阶段 + QC + 分析：
    python3 bioid_run.py ~/raw/BioID-9th --out ~/runs/BioID-9th --yes
  已经有处理好的工作簿，只跑 QC + 分析：
    python3 bioid_run.py ~/runs/BioID-9th/01_processed --out ~/runs/BioID-9th --skip-preprocess
""")
    parser.add_argument("input", help="原始导出目录；配 --skip-preprocess 时是已处理工作簿目录")
    parser.add_argument("--out", required=True, help="运行目录（会创建 01_processed/02_qc/03_analysis/logs）")
    parser.add_argument("--channel-map", help="bioid 的通道映射（默认 <input>/channel_map.tsv）")
    parser.add_argument("--bait-map", help="analysis 的人工诱饵映射 <文件名> <基因>")
    parser.add_argument("--bait-alias", help="analysis 的诱饵别名表 <标签> <数据库符号>")
    parser.add_argument("--exempt-genes", help="batch-qc 的豁免基因（不参与背景旗）")
    parser.add_argument("--sheet", help="bioid/analysis 读取的工作表名")
    parser.add_argument("--force", action="store_true", help="预处理阶段允许覆盖已有 LOG2FC 的工作表")
    parser.add_argument("--skip-preprocess", action="store_true", help="输入已是 bioid 处理产物")
    parser.add_argument("--skip-qc", action="store_true", help="不跑批次 QC（分析会缺三列，仅调试用）")
    parser.add_argument("--skip-analysis", action="store_true", help="只跑到 QC")
    parser.add_argument("--resume", action="store_true", help="已有阶段产物就跳过")
    parser.add_argument("--dry-run", action="store_true", help="只打印命令")
    parser.add_argument("--yes", action="store_true", help="跳过诱饵确认闸门（确认过文件名推断正确再用）")
    parser.add_argument("--python", default=sys.executable, help="跑 Python 阶段的解释器")
    parser.add_argument("--rscript", default=os.environ.get("RSCRIPT", "Rscript"))
    args = parser.parse_args(argv)

    try:
        input_dir, out_dir, sources = preflight(args)
    except SystemExit as error:
        say(f"错误: {error}")
        return EXIT_USAGE
    dirs = {
        "processed": out_dir / "01_processed",
        "qc": out_dir / "02_qc",
        "analysis": out_dir / "03_analysis",
        "logs": out_dir / "logs",
    }
    if not args.dry_run:
        for path in dirs.values():
            path.mkdir(parents=True, exist_ok=True)
    manifest = {"started": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "input": str(input_dir), "out": str(out_dir), "stages": []}

    def record(stage, command, code=None, elapsed=None, skipped=False):
        manifest["stages"].append({
            "stage": stage, "command": [str(c) for c in command],
            "exit": code, "seconds": round(elapsed, 1) if elapsed is not None else None,
            "skipped": skipped,
        })

    # 1) 通道映射：显式文件 > 输入目录的 channel_map.tsv > 自动推断草稿（<out>/00_channel_map.tsv）
    map_path = None
    if not args.skip_preprocess:
        explicit = Path(args.channel_map).expanduser() if args.channel_map else None
        input_map = input_dir / "channel_map.tsv"
        if explicit:
            map_path = explicit
        elif input_map.exists():
            map_path = input_map
        else:
            map_path = out_dir / "00_channel_map.tsv"
            fresh = not map_path.exists()
            records = infer_channel_map(input_dir, args.sheet) if fresh else None
            if fresh and not records:
                say("错误：没能从原始工作簿里推断出通道（找不到两列 Abundance:）。"
                    "请手写 channel_map.tsv 并用 --channel-map 指定。")
                return EXIT_USAGE
            if fresh and not args.dry_run:
                write_channel_map(map_path, records)
            say(f"\n[channel-map] {'自动推断草稿' if fresh else '沿用已有草稿'}：{map_path}")
            alias_path = out_dir / "00_bait_aliases.tsv"
            pairs = []
            if fresh:
                print_channel_map_gate(records, map_path, written=not args.dry_run)
                pairs = proposed_bait_aliases(records)
                if pairs and not args.bait_alias:
                    if not args.dry_run and not alias_path.exists():
                        write_bait_alias_draft(alias_path, pairs)
                    args.bait_alias = str(alias_path)
                    print_alias_gate(pairs, alias_path, written=not args.dry_run)
            else:
                say(map_path.read_text(encoding="utf-8"))
                if not args.bait_alias and alias_path.exists():
                    args.bait_alias = str(alias_path)
                    say(f"[bait-alias] 沿用 {alias_path}")
            manifest["channel_map"] = str(map_path)
            if not args.dry_run and not args.yes:
                # 同一轮里已经确认过（有 manifest）的 --resume 不再拦
                already = args.resume and (out_dir / "run_manifest.json").exists()
                if not already:
                    say("\n确认无误后加 --yes 重跑；要改就编辑上面的文件"
                        "（通道映射 / 诱饵别名都可以直接编辑；也可以 --bait-alias 指定自己的表）。")
                    return EXIT_CONFIRM
        manifest["channel_map"] = str(map_path)
        say(f"[channel-map] 使用 {map_path}")

    # 2) 预处理
    processed_dir = input_dir if args.skip_preprocess else dirs["processed"]
    if args.skip_preprocess:
        n_expected = len(workbook_count(input_dir))
    else:
        n_expected = len(sources)
    if not args.skip_preprocess:
        resume_ready = args.resume and len(workbook_count(dirs["processed"])) >= n_expected > 0
        if resume_ready:
            say(f"\n[preprocess] --resume：已有 {len(workbook_count(dirs['processed']))} 份产物，跳过")
            record("preprocess", [], skipped=True)
        else:
            command = [args.rscript, str(R_SCRIPT), str(input_dir), str(dirs["processed"]),
                       "--map", str(map_path)]
            if args.force:
                command.append("--force")
            if args.sheet:
                command += ["--sheet", args.sheet]
            code, elapsed = run_stage("preprocess", command, dirs["logs"] / "01_preprocess.log",
                                      args.dry_run)
            record("preprocess", command, code, elapsed)
            if code != 0:
                say(f"\n预处理失败（exit={code}）；日志 {dirs['logs'] / '01_preprocess.log'}")
                return EXIT_STAGE

    # 3) 批次 QC
    if args.skip_qc:
        if not args.skip_analysis:
            say("\n错误：--skip-qc 时不能自动跑分析（会缺 QC 三列）。"
                "要么去掉 --skip-qc，要么加 --skip-analysis 只跑到 QC。")
            return EXIT_USAGE
        say("\n[qc] --skip-qc：跳过")
        record("qc", [], skipped=True)
        qc_folder = None
    else:
        qc_ready = args.resume and len(workbook_count(dirs["qc"], "_flagged.xlsx")) >= n_expected > 0
        if qc_ready:
            say(f"\n[qc] --resume：已有 {len(workbook_count(dirs['qc'], '_flagged.xlsx'))} 份 flagged 工作簿，跳过")
            record("qc", [], skipped=True)
        else:
            command = [args.python, str(QC_SCRIPT), str(processed_dir), "--out", str(dirs["qc"])]
            if args.exempt_genes:
                command += ["--exempt-genes", args.exempt_genes]
            if args.bait_alias:
                # 构造名（如 FAM118B-L32R）要让 QC 也能解析回基因，否则诱饵回收/转运全是 "failure"
                command += ["--bait-alias", args.bait_alias]
            code, elapsed = run_stage("qc", command, dirs["logs"] / "02_qc.log", args.dry_run)
            record("qc", command, code, elapsed)
            if code != 0:
                say(f"\n批次 QC 失败（exit={code}）；日志 {dirs['logs'] / '02_qc.log'}")
                return EXIT_STAGE
        qc_folder = dirs["qc"]

    if args.skip_analysis:
        say("\n[analysis] --skip-analysis：停在这里")
        return 0

    # 4) 诱饵复核：处理后的工作簿上再看一遍证据（正常路径不拦，验收阶段兜底）
    if qc_folder is None:
        say("\n错误：跳过 QC 时分析阶段无法自动运行（缺 QC 输入）。请去掉 --skip-qc。")
        return EXIT_USAGE
    bait_cmd = [args.python, str(ANALYSIS_SCRIPT), str(qc_folder), "--list-baits"]
    if args.bait_map:
        bait_cmd += ["--bait-map", args.bait_map]
    if args.bait_alias:
        bait_cmd += ["--bait-alias", args.bait_alias]
    if args.sheet:
        bait_cmd += ["--sheet", args.sheet]
    code, elapsed = run_stage("bait-check", bait_cmd, dirs["logs"] / "03_bait_check.log", args.dry_run)
    record("bait-check", bait_cmd, code, elapsed)
    if code != 0:
        say(f"\n诱饵核对失败（exit={code}）；日志 {dirs['logs'] / '03_bait_check.log'}")
        return EXIT_STAGE
    if (not args.dry_run and args.skip_preprocess and not args.yes
            and not args.bait_map):
        say("\n提示：上面是每份工作簿推断出的诱饵与证据。确认无误后加 --yes 重跑"
            "（或提供 --bait-map），已完成的阶段用 --resume 跳过。")
        return EXIT_CONFIRM

    # 5) 打分 + 报告
    # --resume 只在上一轮验收通过时才跳过分析：验收失败（如诱饵没标到）后，
    # 用户改了 --bait-alias/--bait-map 重跑，必须让分析阶段真的重跑。
    previous = read_payload(out_dir / "run_manifest.json") or {}
    previous_ok = bool((previous or {}).get("verify", {}).get("ok"))
    flagged_files = workbook_count(dirs["qc"], "_flagged.xlsx")
    signature = digest(
        str(qc_folder), args.bait_map or "", args.bait_alias or "", args.sheet or "",
        ANALYSIS_SCRIPT.stat().st_mtime if ANALYSIS_SCRIPT.exists() else 0,
        max((p.stat().st_mtime for p in flagged_files), default=0),
    )
    analysis_ready = (args.resume and previous_ok
                      and previous.get("analysis_signature") == signature
                      and stage_done([dirs["analysis"] / "BioID_analysis_report.html"]))
    if analysis_ready:
        say("\n[analysis] --resume：报告已存在，跳过")
        record("analysis", [], skipped=True)
    else:
        command = [args.python, str(ANALYSIS_SCRIPT), str(qc_folder),
                   "--from-qc", "--out", str(dirs["analysis"])]
        if args.bait_map:
            command += ["--bait-map", args.bait_map]
        if args.bait_alias:
            command += ["--bait-alias", args.bait_alias]
        if args.sheet:
            command += ["--sheet", args.sheet]
        code, elapsed = run_stage("analysis", command, dirs["logs"] / "04_analysis.log", args.dry_run)
        record("analysis", command, code, elapsed)
        if code != 0:
            say(f"\n分析失败（exit={code}）；日志 {dirs['logs'] / '04_analysis.log'}")
            return EXIT_STAGE

    if args.dry_run:
        say("\n--dry-run：没有执行任何阶段")
        return 0

    # 5) 验收
    ok, lines = verify(out_dir, processed_dir, n_expected, args)
    say("\n验收：")
    for line in lines:
        say("  " + line)
    manifest["verify"] = {"ok": ok, "lines": lines}
    manifest["analysis_signature"] = signature
    manifest["finished"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    (out_dir / "run_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    say(f"\n产物目录：{out_dir}")
    if args.skip_preprocess:
        say(f"  01_processed/  （--skip-preprocess：输入目录 {len(workbook_count(processed_dir))} 份）")
    else:
        say(f"  01_processed/  {len(workbook_count(dirs['processed']))} 份")
    say(f"  02_qc/         {len(workbook_count(dirs['qc'], '_flagged.xlsx'))} 份 flagged + BioID_batch_QC.xlsx")
    say(f"  03_analysis/   BioID_analysis_report.html / .xlsx / payload")
    say(f"  run_manifest.json 记录了每步命令与耗时")
    if not ok:
        say("\n验收未通过：见上面的 FAIL 行与 logs/。")
        return EXIT_VERIFY
    say("\n验收通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
