#!/usr/bin/env python3
"""bioid-trend 的行为测试：用合成工作簿验证取值列、NA=0、别名、输出文件。

跑法（需要一个带 openpyxl + numpy + matplotlib 的 python）：
    /usr/bin/python3 -m unittest discover -s tests -v
"""
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

SKILL = Path(__file__).resolve().parent.parent
SCRIPT = SKILL / "scripts" / "bioid_trend.py"
PYTHON = sys.executable

HEADER = ["Checked", "Master", "Accession", "Description", "GN", "GN", "Coverage [%]",
          "# Peptides", "# PSMs", "# Unique Peptides", "Abundance: F1: Sample",
          "Abundance: F9: Control", "LOG2FC", "LOG2FC", "Norm_LOG2FC"]


def write_workbook(path, rows, sheet="Proteins", header=HEADER, extra_sheets=()):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(header)
    for row in rows:
        ws.append(row)
    for name in extra_sheets:
        wb.create_sheet(name)
    wb.save(path)


def row(gn, acc, value, raw=None, peptides=5, columns=len(HEADER)):
    """按 HEADER 的位置填一行（最后三列是 LOG2FC/LOG2FC/Norm_LOG2FC）。"""
    cells = [None] * columns
    cells[2] = acc
    cells[3] = f"{gn} test protein"
    cells[4] = gn
    cells[5] = gn
    cells[9] = peptides
    cells[12] = raw
    cells[13] = raw
    cells[14] = value
    return cells


def row_raw(gn, acc, raw, peptides=5):
    """只有 bioid 产物的 LOG2FC 列（无 Norm_LOG2FC）。"""
    cells = [None] * 13
    cells[2] = acc
    cells[3] = f"{gn} test protein"
    cells[4] = gn
    cells[5] = gn
    cells[9] = peptides
    cells[12] = raw
    return cells


class BioidTrendTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="bioid_trend_test_"))
        self.work = self.tmp / "work"
        self.work.mkdir()

        # A: batch-QC 产物，PPP2R1B 有行但 Norm_LOG2FC 为空 -> no_value
        write_workbook(self.work / "A_processed_flagged.xlsx", [
            row("PPP2R1A", "P30153", 1.0, 0.5),
            row("PPP2R1B", "P30154", None, 0.5),
            row("PPP2CA", "P67775", 3.0, 2.0),
        ])
        # B: 只有 bioid 的 LOG2FC
        write_workbook(self.work / "B_processed.xlsx", [
            row_raw("PPP2R1A", "P30153", 2.0),
            row_raw("PPP2R1B", "P30154", 2.5),
        ], header=[h for h in HEADER[:13]])
        # C: bioid-analysis 的 summary，两个条件页，PPP2CA 整行缺失 -> absent
        summary = self.work / "BioID_analysis_summary.xlsx"
        wb = Workbook()
        wb.active.title = "Overview"
        ws = wb.create_sheet("COND-A")
        ws.append(["Condition", "GN", "Accession", "Description", "FC (EV-corrected)",
                   "log2FC (EV-corrected)"])
        ws.append(["COND-A", "PPP2R1A", "P30153", "x", 1.19, 0.25])
        ws.append(["COND-A", "PPP2R1B", "P30154", "x", 1.68, 0.75])
        wb.save(summary)
        self.summary = summary

        self.samples = self.tmp / "samples.tsv"
        self.samples.write_text(
            "# label\tfile\tsheet\n"
            "cond-A\twork/A_processed_flagged.xlsx\t\n"
            "cond-B\twork/B_processed.xlsx\t\n"
            "cond-C\twork/BioID_analysis_summary.xlsx\tCOND-A\n", encoding="utf-8")
        self.proteins = self.tmp / "proteins.tsv"
        self.proteins.write_text(
            "# gene\tgroup\n"
            "PPP2R1A\tA scaffold\n"
            "PPP2R1B\tA scaffold\n"
            "PPP2CA\tC catalytic\n", encoding="utf-8")
        self.out = self.tmp / "out"

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_cli(self, *extra, check=True):
        cmd = [PYTHON, str(SCRIPT), "--samples", str(self.samples),
               "--proteins", str(self.proteins), *extra]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if check:
            self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        return result

    # -- 取数与 NA ---------------------------------------------------------
    def test_check_reads_each_format_and_flags_na(self):
        result = self.run_cli("--check", "--xlabel", "construct")
        text = result.stdout
        self.assertIn("Norm_LOG2FC", text)       # cond-A
        self.assertIn("LOG2FC", text)            # cond-B 回退
        self.assertIn("log2FC (EV-corrected)", text)   # cond-C summary
        self.assertIn("PPP2R1A", text)
        # 三个条件里 PPP2R1A 依次 1.00 / 2.00 / 0.25
        self.assertIn("1.00", text)
        self.assertIn("2.00", text)
        self.assertIn("0.25", text)
        # NA 一律显示 ▲0，并有 no_value / absent 的来源说明
        self.assertIn("▲0", text)
        self.assertIn("no_value", text)
        self.assertIn("absent", text)
        self.assertIn("没有写任何文件", text)
        self.assertFalse(self.out.exists() and any(self.out.iterdir()))

    def test_relative_paths_resolve_against_samples_file(self):
        # samples.tsv 位于 self.tmp，file 写的是相对路径 work/...
        result = self.run_cli("--check")
        self.assertIn("A_processed_flagged.xlsx", result.stdout)
        self.assertNotIn("读不了", result.stdout)

    # -- 出图与产物 ---------------------------------------------------------
    def test_render_outputs_pdf_png_tsv_and_vector_pdf(self):
        result = self.run_cli("--out", str(self.out), "--title", "test title",
                              "--xlabel", "construct", "--legend-title", "groups")
        stem_candidates = sorted(self.out.glob("*/*.pdf"))
        self.assertEqual(len(stem_candidates), 1, msg=result.stdout)
        pdf = stem_candidates[0]
        png = pdf.with_suffix(".png")
        # 默认布局：<out>/<stem>/<stem>.pdf|png + <out>/<stem>/data/<stem>_values.tsv|_sources.tsv
        self.assertEqual(pdf.parent.parent, self.out)
        self.assertEqual(pdf.stem, pdf.parent.name)
        self.assertEqual(png.parent, pdf.parent)
        data = pdf.parent / "data"
        self.assertTrue(data.is_dir(), msg="data/ 子目录没建")
        values = data / (pdf.stem + "_values.tsv")
        sources = data / (pdf.stem + "_sources.tsv")
        self.assertEqual(list(self.out.glob("*.pdf")), [], msg="默认布局不应在 --out 根目录留 PDF")
        for path, minimum in ((pdf, 5000), (png, 5000), (values, 100), (sources, 100)):
            self.assertTrue(path.exists(), msg=f"{path} 没生成")
            self.assertGreater(path.stat().st_size, minimum, msg=f"{path} 太小")

        # 数值表：NA 原样写 NA（画图位置 0 只在图上）
        table = values.read_text(encoding="utf-8").splitlines()
        self.assertTrue(table[0].startswith("gene"))
        rows = {line.split("\t")[0]: line.split("\t") for line in table[1:]}
        self.assertEqual(rows["PPP2R1A"][3:6], ["1", "2", "0.25"])
        self.assertEqual(rows["PPP2R1B"][3:6], ["NA", "2.5", "0.75"])
        self.assertEqual(rows["PPP2R1B"][6], "1")          # na_count
        self.assertEqual(rows["PPP2CA"][4:6], ["NA", "NA"])  # B、C 里没有这行
        self.assertEqual(rows["PPP2CA"][6], "2")

        # 来源表：逐文件记了实际用的列
        src = sources.read_text(encoding="utf-8").splitlines()
        self.assertIn("Norm_LOG2FC", src[1])
        self.assertIn("LOG2FC", src[2])
        self.assertIn("log2FC (EV-corrected)", src[3])

        # 矢量 PDF：内嵌的是文字/路径，没有整页位图
        if shutil.which("pdfimages"):
            listing = subprocess.run(["pdfimages", "-list", str(pdf)],
                                     capture_output=True, text=True).stdout
            images = [line for line in listing.splitlines()[2:] if line.strip()]
            self.assertEqual(images, [], msg="PDF 里出现了位图对象：" + listing)

    def test_flat_layout_keeps_everything_in_out_dir(self):
        result = self.run_cli("--out", str(self.out), "--name", "flat_case", "--flat")
        pdf = self.out / "flat_case.pdf"
        self.assertTrue(pdf.exists(), msg=result.stdout)
        self.assertTrue((self.out / "flat_case.png").exists())
        self.assertTrue((self.out / "flat_case_values.tsv").exists())
        self.assertTrue((self.out / "flat_case_sources.tsv").exists())
        self.assertFalse((self.out / "flat_case").exists(), msg="--flat 不应再建子目录")

    def test_caption_mentions_missing_points(self):
        self.run_cli("--out", str(self.out))
        pdf = sorted(self.out.glob("*/*.pdf"))[0]
        self.assertIsNotNone(pdf)
        if shutil.which("pdftotext"):
            text = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                                  capture_output=True, text=True).stdout
            self.assertIn("Hollow markers", text)
            self.assertIn("PPP2R1B", text)
            self.assertIn("PPP2CA", text)

    # -- 其它入口 -----------------------------------------------------------
    def test_alias_table_matches_workbook_symbol(self):
        write_workbook(self.work / "D_alias_flagged.xlsx", [
            row("FAM214B", "Q7L5A3", 4.0, 3.0),
        ])
        samples = self.tmp / "samples_alias.tsv"
        samples.write_text("label\tfile\t\nFL\twork/D_alias_flagged.xlsx\t\n", encoding="utf-8")
        aliases = self.tmp / "aliases.tsv"
        aliases.write_text("ATOSB\tFAM214B\n", encoding="utf-8")
        proteins = self.tmp / "proteins_alias.tsv"
        proteins.write_text("ATOSB\n", encoding="utf-8")
        result = subprocess.run(
            [PYTHON, str(SCRIPT), "--samples", str(samples), "--proteins", str(proteins),
             "--aliases", str(aliases), "--check"],
            capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn("alias:FAM214B", result.stdout)
        self.assertIn("4.00", result.stdout)

    @staticmethod
    def load_module():
        import importlib.util
        spec = importlib.util.spec_from_file_location("bioid_trend_mod", SCRIPT)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def test_segment_runs_dashes_only_na_segments(self):
        module = self.load_module()
        # 没有未检出：整条实线
        self.assertEqual(module.segment_runs([False, False, False]), [(0, 2, False)])
        # 中间一个未检出：两侧两段都算"接触未检出"，合并成一段虚线
        self.assertEqual(module.segment_runs([False, True, False]), [(0, 2, True)])
        # 只有最后一段接触未检出
        self.assertEqual(module.segment_runs([False, False, True]), [(0, 1, False), (1, 2, True)])
        # 全未检出
        self.assertEqual(module.segment_runs([True, True, True]), [(0, 2, True)])
        # 单点没有线段
        self.assertEqual(module.segment_runs([False]), [])

    def test_all_proteins_are_solid_by_default(self):
        module = self.load_module()
        series = [
            {"gene": "PPP2R1A", "group": "A scaffold", "style": "", "color": ""},
            {"gene": "PPP2R1B", "group": "A scaffold", "style": "", "color": ""},
            {"gene": "PPP2CA", "group": "C catalytic", "style": "", "color": ""},
        ]
        module.assign_styles(series, None)
        self.assertEqual([item["linestyle"] for item in series], ["-", "-", "-"])
        # 显式 style 仍然覆盖
        series[1]["style"] = "dotted"
        module.assign_styles(series, None)
        self.assertEqual(series[1]["linestyle"], ":")

    def test_xgroup_runs_merges_contiguous_names(self):
        module = self.load_module()
        self.assertEqual(module.xgroup_runs(["", "", ""]), [])
        self.assertEqual(module.xgroup_runs(["FL", "trunc", "trunc"]),
                         [(0, 0, "FL"), (1, 2, "trunc")])
        # 中间空一格：不跨过去
        self.assertEqual(module.xgroup_runs(["A", "", "A"]),
                         [(0, 0, "A"), (2, 2, "A")])
        self.assertEqual(module.xgroup_runs(["B", "B", "C", "B"]),
                         [(0, 1, "B"), (2, 2, "C"), (3, 3, "B")])

    def test_mean_counts_missing_as_zero(self):
        module = self.load_module()
        series = [
            {"values": [1.0, 0.0, 3.0], "na": [False, True, False]},
            {"values": [3.0, 2.0, 0.0], "na": [False, False, True]},
        ]
        self.assertEqual(module.compute_means(series, 3), [2.0, 1.0, 1.5])

    def test_groups_and_mean_in_plot_and_tables(self):
        samples = self.tmp / "samples_groups.tsv"
        samples.write_text(
            "label\tfile\txgroup\n"
            "cond-A\twork/A_processed_flagged.xlsx\tFL\n"
            "cond-B\twork/B_processed.xlsx\ttrunc\n"
            "cond-C\twork/BioID_analysis_summary.xlsx\ttrunc\n", encoding="utf-8")
        result = subprocess.run(
            [PYTHON, str(SCRIPT), "--samples", str(samples), "--proteins", str(self.proteins),
             "--mean", "--check", "--xlabel", "construct"],
            capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn("《FL》", result.stdout)
        self.assertIn("MEAN", result.stdout)
        self.assertIn("1.33", result.stdout)      # (1 + 0 + 3) / 3

        out = self.tmp / "grouped"
        subprocess.run(
            [PYTHON, str(SCRIPT), "--samples", str(samples), "--proteins", str(self.proteins),
             "--mean", "--out", str(out), "--name", "grp"],
            capture_output=True, text=True, check=True)
        values = (out / "grp" / "data" / "grp_values.tsv").read_text(encoding="utf-8")
        mean_line = [line for line in values.splitlines() if line.startswith("MEAN")][0]
        self.assertEqual(mean_line.split("\t")[3:6], ["1.33333", "1.5", "0.333333"])
        if shutil.which("pdftotext"):
            text = subprocess.run(["pdftotext", "-layout", str(out / "grp" / "grp.pdf"), "-"],
                                  capture_output=True, text=True).stdout
            self.assertIn("Background bands group the x axis", text)
            self.assertIn("FL (cond-A)", text)
            self.assertIn("mean across the 3 plotted proteins", text)

    def test_caption_drops_bands_when_disabled(self):
        samples = self.tmp / "samples_groups2.tsv"
        samples.write_text(
            "label\tfile\txgroup\n"
            "cond-A\twork/A_processed_flagged.xlsx\tFL\n"
            "cond-B\twork/B_processed.xlsx\ttrunc\n", encoding="utf-8")
        out = self.tmp / "nobands"
        subprocess.run(
            [PYTHON, str(SCRIPT), "--samples", str(samples), "--proteins", str(self.proteins),
             "--no-xgroup-bands", "--out", str(out), "--name", "nobands"],
            capture_output=True, text=True, check=True)
        if shutil.which("pdftotext"):
            text = subprocess.run(["pdftotext", "-layout", str(out / "nobands" / "nobands.pdf"), "-"],
                                  capture_output=True, text=True).stdout
            self.assertNotIn("Background bands", text)

    def test_cjk_text_switches_font(self):
        # PDF 后端不会逐字回退：图里有中文时必须把 CJK 字体提到最前，否则中文是空白
        module = self.load_module()
        family = ["DejaVu Sans"]
        note = module.choose_cjk_font(["全长", "ATOSB"], family)
        if note is None:
            self.skipTest("这台机器没有 matplotlib/font_manager")
        if "改用" in note:
            self.assertNotEqual(family[0], "DejaVu Sans", msg=note)
        english = ["DejaVu Sans"]
        self.assertIsNone(module.choose_cjk_font(["FL", "N", "PPP2R1A"], english))
        self.assertEqual(english, ["DejaVu Sans"])

    def test_inspect_reports_sheets_and_column(self):
        result = subprocess.run(
            [PYTHON, str(SCRIPT), "--inspect", str(self.summary), "--sheet", "COND-A"],
            capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn("COND-A", result.stdout)
        self.assertIn("log2FC (EV-corrected)", result.stdout)

    def test_summary_without_sheet_is_refused(self):
        write_workbook(self.work / "E_multi_flagged.xlsx", [row("PPP2R1A", "P30153", 1.0, 0.5)])
        # 造一个多条件 summary：不点名 sheet 时必须报错而不是随便挑一张
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.title = "Overview"
        for cond in ("COND-A", "COND-B"):
            ws = wb.create_sheet(cond)
            ws.append(["Condition", "GN", "log2FC (EV-corrected)"])
            ws.append([cond, "PPP2R1A", 1.0])
        path = self.tmp / "multi_summary.xlsx"
        wb.save(path)
        result = subprocess.run(
            [PYTHON, str(SCRIPT), "--inspect", str(path)],
            capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("COND-A", result.stderr)
        self.assertIn("sheet", result.stderr)

    def test_missing_value_column_is_an_error(self):
        result = self.run_cli("--value", "norm", "--check", check=False)
        # cond-A 有 Norm_LOG2FC，cond-B 没有 -> 应在 cond-B 上明确报错
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Norm_LOG2FC", result.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
