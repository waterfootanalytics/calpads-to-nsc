import csv
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch, sentinel

from openpyxl import Workbook

from calpads_to_nsc.generate import ValidationError, detail_row
from calpads_to_nsc.reference_sheet import read_reference
from calpads_to_nsc.reports import columns, cycle_for_exit
from calpads_to_nsc.supplement import read_supplement


REPORT_COLUMNS = (
    "School Name",
    "SSID",
    "Student Legal Name",
    "Local Student ID",
    "Birth Date",
    "Enrollment Exit Date",
    "Student Exit Reason Code",
    "School Completion Status",
)


class FileContractsTest(unittest.TestCase):
    def test_certified_headers_pass_through_cli(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.csv"
            path.write_text(",".join(REPORT_COLUMNS) + "\n", encoding="utf-8")
            command = [sys.executable, "-m", "calpads_to_nsc", str(path), "--check-headers"]
            result = subprocess.run(command, capture_output=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn(b"Report headers passed", result.stdout)

    def test_required_and_ambiguous_report_headers(self):
        with self.assertRaisesRegex(ValidationError, "Missing"):
            columns(REPORT_COLUMNS[1:])
        with self.assertRaisesRegex(ValidationError, "Ambiguous"):
            columns([*REPORT_COLUMNS, "State Student ID"])

    def test_optional_block_column(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "supplement.csv"
            for block_column in (False, True):
                with self.subTest(block_column=block_column):
                    headings = "SSID,GraduationDate" + (",FERPABlock" if block_column else "")
                    path.write_text(headings + "\n", encoding="utf-8")
                    data = read_supplement(path)
                    self.assertEqual(data.has_block_column, block_column)
                    self.assertFalse(data.dates)
                    self.assertFalse(data.blocked)

    def test_reporting_cycle_boundary(self):
        self.assertEqual(cycle_for_exit(date(2030, 8, 15)), 2030)
        self.assertEqual(cycle_for_exit(date(2030, 8, 16)), 2031)

    def test_pdf_dependency_and_unsupported_word_extensions(self):
        with patch("subprocess.run", side_effect=FileNotFoundError), self.assertRaisesRegex(ValidationError, "Poppler"):
            read_reference(Path("Account Reference Sheet.pdf"))
        for extension in (".doc", ".docx"):
            with self.subTest(extension=extension), self.assertRaisesRegex(ValidationError, "xls, .xlsx or PDF"):
                read_reference(Path("Account Reference Sheet" + extension))

    def test_nsc_detail_fields_and_order(self):
        student_fields = {
            key: getattr(sentinel, key) for key in ("first_name", "middle_name", "last_name", "birth_date")
        }
        award_fields = {key: getattr(sentinel, key) for key in ("diploma_type", "graduation_date", "ferpa_block")}
        school_fields = {key: getattr(sentinel, key) for key in ("nsc_school_name", "act_code")}
        with patch("calpads_to_nsc.generate.name", side_effect=lambda value, *_: value):
            row = detail_row(student_fields, sentinel.student_id, award_fields, school_fields)
        self.assertEqual(
            row,
            [
                "PD3",
                "NO SSN",
                sentinel.first_name,
                sentinel.middle_name,
                sentinel.last_name,
                "",
                "",
                "",
                sentinel.birth_date,
                sentinel.student_id,
                sentinel.diploma_type,
                sentinel.graduation_date,
                sentinel.ferpa_block,
                sentinel.nsc_school_name,
                sentinel.act_code,
                *([""] * 12),
                "ED",
            ],
        )


class PublishedLayoutTest(unittest.TestCase):
    def test_certified_report_headings_and_complete_graduates_layout(self):
        with tempfile.TemporaryDirectory() as scratch:
            root = Path(scratch)
            report = root / "Certified 1.23.csv"
            headings = [
                "School Code",
                "School Name",
                "SSID",
                "Student Legal Name",
                "Local Student ID",
                "Birth Date",
                "Enrollment Exit Date",
                "Student Exit Reason Code",
                "School Completion Status",
            ]
            with report.open("x", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(headings)
                writer.writerow(
                    [
                        "0000000",
                        "Practice High School",
                        "0000000001",
                        "Sample, River S",
                        "0004",
                        "2006-01-15",
                        "2024-06-12",
                        "E230",
                        "100 - HS Diploma",
                    ]
                )
            reference = root / "Account.xlsx"
            book = Workbook()
            for row in (
                ("Organization Name:", "Practice District"),
                ("Service Agreement/Account Number:", "00001234"),
                ("Total High School Count:", 1),
                ("ACT Code", "HS Operating Status", "High School Name"),
                ("000123", "OPEN", "PRACTICE HIGH SCHOOL"),
            ):
                book.active.append(row)
            book.save(reference)
            book.close()
            workbook = root / "School review.xlsx"
            command = [
                sys.executable,
                "-B",
                "-m",
                "calpads_to_nsc",
                str(report),
                "--reference-sheet",
                str(reference),
                "--date",
                "20240701",
                "--no-prompt",
            ]
            subprocess.run([*command, "--create-crosswalk", str(workbook)], check=True, capture_output=True)
            expected = (
                b"PH3\t00001234\tPractice District\tP\t20240701\t2024\r\n"
                b"PD3\tNO SSN\tRiver\tS\tSample\t\t\t\t20060115\t0004\tRegular Diploma\t20240612\tN\t"
                b"PRACTICE HIGH SCHOOL\t000123" + b"\t" * 13 + b"ED\r\nPT3\t3\r\n"
            )
            for mode in ("--test-output", "--output"):
                with self.subTest(mode=mode):
                    target = root / (mode.strip("-") + ".txt")
                    result = subprocess.run(
                        [*command, "--crosswalk", str(workbook), "--confirm-no-ferpa-blocks", mode, str(target)],
                        capture_output=True,
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    self.assertEqual(target.read_bytes(), expected)
