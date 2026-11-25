import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from compare_excel import compare


class CompareExcelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def workbook(self, name, headers, rows):
        book = Workbook()
        sheet = book.active
        sheet.title = "uniform"
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        path = self.folder / name
        book.save(path)
        book.close()
        return path

    def run_compare(self, first, second):
        output = io.StringIO()
        with patch("compare_excel.get_configured_path", return_value=self.folder):
            with contextlib.redirect_stdout(output):
                counts = compare(first, second)
        return counts, output.getvalue()

    def test_reordered_rows_and_non_ok_additions(self):
        headers = ["Kundennummer", "Kunde", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [
            [1, "A", "OK"], [2, "B", None],
        ])
        second = self.workbook("second.xlsx", headers, [
            [3, "new", None], [2, "B", None], [1, "A", " ok "],
        ])
        counts, output = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 2, "different": 1, "warnings": 0})
        self.assertNotIn("WARNUNG", output)

    def test_duplicate_count_mismatch_warns_with_or_without_ok(self):
        headers = ["Kunde", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [
            ["A", "OK"], ["A", None], ["B", None],
        ])
        second = self.workbook("second.xlsx", headers, [
            ["B", None], ["A", None], ["B", None], ["B", None],
        ])
        counts, output = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 2, "different": 3, "warnings": 3})
        self.assertIn("Datei 1: 2", output)
        self.assertIn("Datei 2: 3", output)

    def test_missing_ok_row_in_either_workbook_warns(self):
        headers = ["Kunde", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [["old", "OK"]])
        second = self.workbook("second.xlsx", headers, [["new", "OK"]])
        counts, _ = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 0, "different": 2, "warnings": 2})

    def test_equal_duplicate_counts_and_blank_rows(self):
        headers = ["Kunde", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [
            ["A", "OK"], [None, None], ["A", None],
        ])
        second = self.workbook("second.xlsx", headers, [
            ["A", None], ["A", "OK"],
        ])
        counts, output = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 2, "different": 0, "warnings": 0})
        self.assertNotIn("WARNUNG", output)

    def test_header_order_and_boolean_values(self):
        first = self.workbook("first.xlsx", ["Kunde", "Wert", "Anmerkungen"], [
            ["A", True, "OK"], ["B", 2, None],
        ])
        second = self.workbook("second.xlsx", ["Anmerkungen", "Wert", "Kunde"], [
            [None, 2, "B"], ["OK", 1, "A"],
        ])
        counts, _ = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 1, "different": 2, "warnings": 2})


if __name__ == "__main__":
    unittest.main()
