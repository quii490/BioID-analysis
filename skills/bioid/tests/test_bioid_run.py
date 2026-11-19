#!/usr/bin/env python3
"""bioid_run.py 的通道映射自动推断：合成工作簿验证三种置信度与通道选择。

跑法： python3 tests/test_bioid_run.py
"""
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

SKILL = Path(__file__).resolve().parent.parent
SCRIPT = SKILL / "scripts" / "bioid_run.py"
HEADER = ["Accession", "Description", "GN", "# Unique Peptides", "Coverage [%]",
          "Abundance: F1: Sample", "Abundance: F9: Control"]


def write_workbook(path, gn, bait_value, control_value, header=None, extra=None):
    header = header or HEADER
    wb = Workbook()
    ws = wb.active
    ws.title = "Proteins"
    ws.append(header)
    ws.append(["P00001", f"{gn} protein", gn, 10, 40, bait_value, control_value])
    for row in extra or []:
        ws.append(row)
    wb.save(path)


class ChannelMapInferenceTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="bioid_run_test_"))

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_dry(self, input_dir):
        result = subprocess.run(
            [sys.executable, str(SCRIPT), str(input_dir), "--out", str(self.tmp / "run"),
             "--dry-run"], capture_output=True, text=True)
        return result

    def test_exact_label_picks_matching_channel(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        write_workbook(raw / "AKAP8L.xlsx", "AKAP8L", 1e8, 2e5)
        write_workbook(raw / "NCOA3.xlsx", "NCOA3", 9e7, 5e3,
                       header=["Accession", "Description", "GN", "# Unique Peptides",
                               "Coverage [%]", "Abundance: F2: Sample",
                               "Abundance: F9: Control"])
        result = self.run_dry(raw)
        self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        self.assertIn("AKAP8L 500x vs F9", result.stdout)
        # 两个文件都用文件名里的 GN 精确命中 -> high 置信
        self.assertIn("high", result.stdout)
        self.assertIn("F1", result.stdout)
        self.assertNotIn("low", result.stdout.split("推断的通道映射")[1].split("草稿已写到")[0])

    def test_bait_high_in_shared_control_channel_is_flagged(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        # AKAP8L：F1 独有、F9 共用，正常；NCOA3：诱饵在 F9（共用）更高 -> 疑似列反
        write_workbook(raw / "AKAP8L.xlsx", "AKAP8L", 1e8, 2e5)
        write_workbook(raw / "NCOA3.xlsx", "NCOA3", 5e3, 9e7,
                       header=["Accession", "Description", "GN", "# Unique Peptides",
                               "Coverage [%]", "Abundance: F2: Sample",
                               "Abundance: F9: Control"])
        result = self.run_dry(raw)
        self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        self.assertIn("列反", result.stdout)
        self.assertIn("low", result.stdout)

    def test_construct_name_falls_back_to_prefix(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        write_workbook(raw / "ATOSB-C.xlsx", "ATOSB", 2e8, 1e5)
        result = self.run_dry(raw)
        self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        self.assertIn("medium", result.stdout)
        self.assertIn("按前缀取 ATOSB", result.stdout)

    def test_no_gene_name_uses_unique_channel(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        # 文件名里没有基因名；F1 只在本文件出现，F9 是共用对照
        write_workbook(raw / "X1.xlsx", "SOMEPROTEIN", 3e7, 1e7)
        write_workbook(raw / "X2.xlsx", "OTHERPROTEIN", 4e7, 1e7,
                       header=["Accession", "Description", "GN", "# Unique Peptides",
                               "Coverage [%]", "Abundance: F2: Sample",
                               "Abundance: F9: Control"])
        result = self.run_dry(raw)
        self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        self.assertIn("low", result.stdout)
        self.assertIn("出现文件数最少", result.stdout)
        # F1 / F2 是各自文件独有的通道，F9 是共用对照
        self.assertIn("F1", result.stdout)
        self.assertIn("F2", result.stdout)

    def test_construct_names_produce_alias_draft_and_ask(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        write_workbook(raw / "ATOSB-C.xlsx", "ATOSB", 2e8, 1e5)
        result = subprocess.run(
            [sys.executable, str(SCRIPT), str(raw), "--out", str(self.tmp / "run")],
            capture_output=True, text=True)
        self.assertEqual(result.returncode, 3, msg=result.stdout + result.stderr)
        self.assertIn("需要你确认的 构造名", result.stdout)
        alias = self.tmp / "run" / "00_bait_aliases.tsv"
        self.assertTrue(alias.exists(), "别名草稿没有生成")
        rows = [line.split() for line in alias.read_text(encoding="utf-8").splitlines()
                if line and not line.startswith("#")]
        self.assertEqual(rows, [["ATOSB-C", "ATOSB"]])

    def test_exact_gene_names_do_not_create_alias_draft(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        write_workbook(raw / "MSANTD2.xlsx", "MSANTD2", 2e8, 1e5)
        result = subprocess.run(
            [sys.executable, str(SCRIPT), str(raw), "--out", str(self.tmp / "run")],
            capture_output=True, text=True)
        self.assertEqual(result.returncode, 3, msg=result.stdout + result.stderr)
        self.assertNotIn("需要你确认的 构造名", result.stdout)
        self.assertFalse((self.tmp / "run" / "00_bait_aliases.tsv").exists())

    def test_existing_input_map_skips_draft(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        write_workbook(raw / "AKAP8L.xlsx", "AKAP8L", 1e8, 2e5)
        (raw / "channel_map.tsv").write_text("F1\tF9\tAKAP8L.xlsx\n", encoding="utf-8")
        result = self.run_dry(raw)
        self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        self.assertIn("使用", result.stdout)
        self.assertNotIn("推断的通道映射", result.stdout)

    def test_dry_run_writes_nothing(self):
        raw = self.tmp / "raw"
        raw.mkdir()
        write_workbook(raw / "AKAP8L.xlsx", "AKAP8L", 1e8, 2e5)
        result = self.run_dry(raw)
        self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
        self.assertFalse((self.tmp / "run").exists(), "dry-run 不应写任何文件")


if __name__ == "__main__":
    unittest.main(verbosity=2)
