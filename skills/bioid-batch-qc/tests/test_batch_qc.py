#!/usr/bin/env python3
"""Regression tests for the Bait Transfer sheet layout in bioid_batch_review.py.

The bug: build_qc_workbook() hardcoded the detail-block start at row 15, so a
batch with >= 13 conditions got its matrix's last row overwritten by the detail
header. The matrix occupies rows 3..(2 + len(names)); the detail header must
come after it with one blank row: row 4 + len(names).

The test builds minimal workbooks whose header names are copied from a real
BioID analysis workbook, runs the real script as a subprocess, and inspects the
generated QC workbook.

Run: python3 tests/test_batch_qc.py
"""

import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "bioid_batch_review.py"
PYTHON = sys.executable
CONTROL_F = 9


def build_workbook(path, order, bait):
    """Write one minimal BioID-style workbook (Proteins / +5% / Raw)."""
    # Names copied from
    # BioID-5th/analysis/bioid/ATOSB-N_processed.xlsx ("Abundance: F1: Sample",
    # "Abundance: F9: Control"); the experimental F number varies per workbook.
    header = [
        "Checked",
        "Master",
        "Accession",
        "Description",
        "GN",
        "Coverage [%]",
        "# Unique Peptides",
        "Abundance: F{}: Sample".format(order),
        "Abundance: F{}: Control".format(CONTROL_F),
    ]
    proteins = [
        [None, "Yes", "Q{:05d}".format(order), bait + " bait", bait, 21.5, 6,
         500000.0 + order * 1000, 20000.0],
        [None, "Yes", "P{:05d}".format(order), "Shared prey one", "PREY1", 12.0, 3,
         100000.0 + order * 500, 15000.0],
        [None, "Yes", "R{:05d}".format(order), "Shared prey two", "PREY2", 8.5, 2,
         80000.0 - order * 300, 12000.0],
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "Proteins"
    ws.append(header)
    for row in proteins:
        ws.append(row)
    for sheet_name in ("+5%", "Raw"):
        ws = wb.create_sheet(sheet_name)
        ws.append(header)
        for row in proteins:
            ws.append(row)
    wb.save(path)


def run_batch(folder):
    result = subprocess.run(
        [PYTHON, str(SCRIPT), str(folder)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        universal_newlines=True,
    )
    if result.returncode != 0:
        raise AssertionError(
            "bioid_batch_review.py failed (exit {}):\nSTDOUT:\n{}\nSTDERR:\n{}".format(
                result.returncode, result.stdout, result.stderr
            )
        )
    return result


def read_transfer(qc_path):
    wb = load_workbook(qc_path, data_only=True)
    try:
        ws = wb["Bait Transfer"]
        rows = [
            [ws.cell(row, col).value for col in (1, 2, 3)]
            for row in range(1, ws.max_row + 1)
        ]
        return rows, ws.auto_filter.ref
    finally:
        wb.close()


class BaitTransferLayoutTest(unittest.TestCase):
    def _build_and_run(self, root, baits):
        folder = Path(root) / "batch"
        folder.mkdir()
        for order, bait in enumerate(baits, 1):
            build_workbook(folder / "{}{}.xlsx".format(order, bait), order, bait)
        run_batch(folder)
        qc_path = folder / "flagged" / "BioID_batch_QC.xlsx"
        self.assertTrue(qc_path.exists(), "QC workbook was not written")
        return read_transfer(qc_path)

    def test_13_conditions_keep_matrix_and_place_detail_below(self):
        baits = ["GENE" + chr(ord("A") + i) for i in range(13)]
        with tempfile.TemporaryDirectory(prefix="bioid_qc_13_") as root:
            rows, filter_ref = self._build_and_run(root, baits)

        matrix_end = 2 + len(baits)  # 15: matrix rows are 3..matrix_end
        detail_start = 4 + len(baits)  # 17: one blank row before the detail header

        # Every condition keeps its matrix row; the old code overwrote the last one.
        self.assertEqual(
            [rows[row - 1][0] for row in range(3, matrix_end + 1)], baits
        )
        self.assertNotIn(
            "Source", [rows[row - 1][0] for row in range(3, matrix_end + 1)]
        )
        # One blank spacer row, then the detail header.
        self.assertIsNone(rows[matrix_end][0])
        self.assertEqual(rows[detail_start - 1][0], "Source")
        self.assertGreater(detail_start, matrix_end)

        # auto_filter must cover only the detail block, not the matrix header.
        self.assertIsNotNone(filter_ref)
        match = re.match(r"^A(\d+):I(\d+)$", filter_ref)
        self.assertIsNotNone(match, "unexpected auto_filter.ref: {!r}".format(filter_ref))
        self.assertEqual(int(match.group(1)), detail_start)
        self.assertEqual(int(match.group(2)), len(rows))

    def test_2_conditions_matrix_and_detail_do_not_overlap(self):
        baits = ["GENEA", "GENEB"]
        with tempfile.TemporaryDirectory(prefix="bioid_qc_2_") as root:
            rows, filter_ref = self._build_and_run(root, baits)

        matrix_end = 2 + len(baits)  # 4
        detail_start = 4 + len(baits)  # 6

        self.assertEqual(
            [rows[row - 1][0] for row in range(3, matrix_end + 1)], baits
        )
        self.assertIsNone(rows[matrix_end][0])
        self.assertEqual(rows[detail_start - 1][0], "Source")
        match = re.match(r"^A(\d+):I(\d+)$", filter_ref)
        self.assertIsNotNone(match, "unexpected auto_filter.ref: {!r}".format(filter_ref))
        self.assertEqual(int(match.group(1)), detail_start)


def build_named_workbook(path, order, gn):
    """Like build_workbook, but the file name and the GN are given separately.

    This is the construct-name case: FAM118B-L32R.xlsx contains GN=FAM118B, so
    the filename-derived bait label is not a gene symbol.
    """
    build_workbook(Path(path), order, gn)


def read_overview(qc_path):
    wb = load_workbook(qc_path, data_only=True)
    try:
        ws = wb["Batch Overview"]
        header = [cell.value for cell in ws[1]]
        index = {name: i for i, name in enumerate(header)}
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            rows.append({
                "condition": row[index["Condition"]],
                "bait_gene": row[index["Bait gene"]],
                "status": row[index["Bait recovery status"]],
            })
        return rows
    finally:
        wb.close()


class BaitAliasTest(unittest.TestCase):
    """Construct-named files must resolve their bait through --bait-alias."""

    def _make_batch(self, root):
        folder = Path(root) / "batch"
        folder.mkdir()
        build_named_workbook(folder / "FAM118B-L32R.xlsx", 1, "FAM118B")
        build_named_workbook(folder / "FAM118B-L32R+F122R.xlsx", 2, "FAM118B")
        return folder

    def test_without_alias_bait_recovery_fails(self):
        with tempfile.TemporaryDirectory(prefix="bioid_qc_alias_") as root:
            folder = self._make_batch(root)
            run_batch(folder)
            rows = read_overview(folder / "flagged" / "BioID_batch_QC.xlsx")
        self.assertTrue(rows)
        self.assertTrue(all("failure" in str(r["status"]).lower() for r in rows),
                        "construct labels should not resolve without --bait-alias: {}".format(rows))

    def test_with_alias_bait_recovery_is_detected(self):
        with tempfile.TemporaryDirectory(prefix="bioid_qc_alias_") as root:
            folder = self._make_batch(root)
            alias = Path(root) / "aliases.tsv"
            alias.write_text("# 标签\t基因\nFAM118B-L32R\tFAM118B\n"
                             "FAM118B-L32R+F122R\tFAM118B\n", encoding="utf-8")
            result = subprocess.run(
                [PYTHON, str(SCRIPT), str(folder), "--bait-alias", str(alias)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            self.assertEqual(result.returncode, 0, msg=result.stdout + result.stderr)
            rows = read_overview(folder / "flagged" / "BioID_batch_QC.xlsx")
        self.assertTrue(rows)
        self.assertTrue(all(r["status"] == "Detected" for r in rows),
                        "alias table should resolve the bait: {}".format(rows))
        self.assertTrue(all(str(r["bait_gene"]).upper() == "FAM118B" for r in rows),
                        "bait gene should come from the alias table: {}".format(rows))


if __name__ == "__main__":
    unittest.main()


if __name__ == "__main__":
    unittest.main()
