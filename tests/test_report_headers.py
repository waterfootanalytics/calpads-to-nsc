import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from calpads_to_nsc.generate import ValidationError
from calpads_to_nsc.reports import HEADER_ROW_LIMIT, folder_inputs, report_records
from calpads_to_nsc.supplement import read_supplement


FIELDS = ("School Name", "SSID", "Student Legal Name", "Birth Date", "Enrollment Exit Date", "School Completion Status")


class ReportHeadingTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)

    def test_delimiters_and_encodings_after_titles(self):
        for delimiter, extension in ((",", ".csv"), ("\t", ".tsv"), ("|", ".txt"), ("^", ".csv")):
            for encoding in ("utf-8-sig", "utf-16"):
                with self.subTest(delimiter=delimiter, encoding=encoding):
                    path = self.folder / ("report" + extension)
                    with path.open("w", encoding=encoding, newline="") as handle:
                        handle.write("CALPADS 1.23, Graduates | Completers\r\n")
                        writer = csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_ALL)
                        writer.writerow([])
                        writer.writerow(["School Name", "SSID"])
                        writer.writerow(FIELDS)
                    self.assertEqual(list(report_records([path], headers_only=True)), [])
                    path.unlink()

    def test_excel_searches_to_limit_and_mixed_folder_roles_are_preserved(self):
        path = self.folder / "AccountReference-export.xlsx"
        book = Workbook()
        book.active.append(["Organization Name:"])
        book.active.append(["School Name", "SSID"])
        for row in range(3, HEADER_ROW_LIMIT):
            book.active.cell(row, 1, "Filter")
        book.active.append(list(FIELDS))
        book.save(path)
        book.close()
        account = self.folder / "account.xlsx"
        book = Workbook()
        book.active.append(["Organization Name:"])
        book.save(account)
        book.close()
        supplement = self.folder / "supplement.csv"
        supplement.write_text("SSID,GraduationDate,FERPABlock\n", encoding="utf-8")
        roles = folder_inputs(self.folder)
        self.assertEqual(roles, {"reports": [path], "reference": [account], "dates": [supplement]})
        self.assertEqual(list(report_records([path])), [])

    def test_missing_header_and_malformed_header_keep_their_locations(self):
        path = self.folder / "report.csv"
        for fields, expected in ((FIELDS + ("SSID",), "Ambiguous"), (FIELDS[:-1], "Missing")):
            with self.subTest(fields=fields):
                path.write_text("Title\nSchool Name,SSID\n" + ",".join(fields) + "\n", encoding="utf-8")
                with self.assertRaises(ValidationError) as raised:
                    list(report_records([path], headers_only=True))
                self.assertIn(f"{path.name}\nRow 3 · {expected}", str(raised.exception))
        path.write_text("Title\n" * HEADER_ROW_LIMIT + ",".join(FIELDS), encoding="utf-8")
        with self.assertRaisesRegex(ValidationError, f"first {HEADER_ROW_LIMIT} rows"):
            list(report_records([path], headers_only=True))

    def test_first_data_error_reports_its_physical_line(self):
        path = self.folder / "report.csv"
        with path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["CALPADS 1.23"])
            writer.writerow([])
            writer.writerow(FIELDS)
            writer.writerow(["", "", "", "\n", "", ""])
            writer.writerow(["incomplete"])
        with self.assertRaisesRegex(ValidationError, "Row 6 · Incomplete report row"):
            list(report_records([path]))

    def test_excel_keeps_worksheet_row_numbers_and_requires_one_sheet(self):
        path = self.folder / "report.xlsx"
        book = Workbook()
        book.active.append(["Title"])
        book.active.append(list(FIELDS) + ["SSID"])
        book.save(path)
        with self.assertRaisesRegex(ValidationError, "Row 2 · Ambiguous"):
            list(report_records([path], headers_only=True))
        book.create_sheet("Other")
        book.save(path)
        book.close()
        with self.assertRaisesRegex(ValidationError, "one flat report worksheet"):
            list(report_records([path], headers_only=True))

    def test_supplement_keeps_header_on_row_one(self):
        path = self.folder / "supplement.csv"
        path.write_text("Title,Filters\nSSID,GraduationDate\n", encoding="utf-8")
        with self.assertRaisesRegex(ValidationError, "Row 1"):
            read_supplement(path)

    def test_header_only_does_not_validate_records_below_it(self):
        path = self.folder / "report.csv"
        path.write_text("Title\n" + ",".join(FIELDS) + '\n"unclosed', encoding="utf-8")
        self.assertEqual(list(report_records([path], headers_only=True)), [])
        with self.assertRaisesRegex(ValidationError, "Row 3 · Cannot read this row"):
            list(report_records([path]))
