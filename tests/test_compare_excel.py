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
        self.assertIn("Erste Datei: 2", output)
        self.assertIn("Zweite Datei: 3", output)

    def test_log_labels_and_changed_columns_for_moved_rows(self):
        headers = ["Kundennummer", "Kunde", "Typ", "Menge", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [
            [1, "A", "X", 10, "OK"], [2, "B", "Y", 1, None],
        ])
        second = self.workbook("second.xlsx", headers, [
            [2, "B", "Y", 1, None], [1, "A", "X", 12, "OK"],
        ])
        counts, output = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 1, "different": 2, "warnings": 2})
        self.assertIn(f"Erste Datei: {first}", output)
        self.assertIn(f"Zweite Datei: {second}", output)
        self.assertIn("Menge: 10 -> 12", output)
        self.assertNotIn("Typ:", output)
        logs = list(self.folder.glob("compare_excel_*.log"))
        self.assertEqual(len(logs), 1)
        self.assertEqual(logs[0].read_text(encoding="utf-8"), output)

    def test_ambiguous_closest_match_is_not_paired(self):
        headers = ["Kundennummer", "Kunde", "Wert", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [[1, "A", 10, "OK"]])
        second = self.workbook("second.xlsx", headers, [
            [1, "A", 11, "OK"], [1, "A", 12, "OK"],
        ])
        counts, output = self.run_compare(first, second)
        self.assertEqual(counts, {"equal": 0, "different": 3, "warnings": 3})
        self.assertNotIn("Wert: 10 ->", output)
        self.assertIn("Kein eindeutiger Vergleichspartner", output)

    def test_unrelated_customers_are_not_paired(self):
        headers = ["Kundennummer", "Kunde", "Wert", "Anmerkungen"]
        first = self.workbook("first.xlsx", headers, [[1, "A", 10, None]])
        second = self.workbook("second.xlsx", headers, [[2, "B", 11, None]])
        _, output = self.run_compare(first, second)
        self.assertNotIn("Wert: 10 -> 11", output)
        self.assertIn("Kein eindeutiger Vergleichspartner", output)

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
