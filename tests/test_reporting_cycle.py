import csv
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from calpads_to_nsc.graduates import eligible_graduates
from calpads_to_nsc.reports import reporting_year
from calpads_to_nsc.supplement import read_supplement
from calpads_to_nsc.generate import ValidationError


class ReportingCycleAndRowTest(unittest.TestCase):
    headers = (
        "SchoolName",
        "SSID",
        "FirstName",
        "LastName",
        "BirthDate",
        "EnrollmentExitDate",
        "StudentCompletionStatus",
        "LocalStudentID",
        "ExitReason",
    )

    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.early = self.root / "early.csv"
        self.late = self.root / "late.csv"
        self.example = [
            "Example High",
            "0000000101",
            "PrivateGiven",
            "PrivateFamily",
            "2008-09-01",
            "2026-07-12",
            "100",
            "PRIVATE01",
            "E230",
        ]

    def write(self, path, *rows, headers=None):
        with path.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(self.headers if headers is None else headers)
            writer.writerows(rows)

    def check_reason(self, caught, filename, row, reason):
        message = str(caught.exception)
        self.assertIn(filename, message)
        self.assertIn(f"Row {row}", message)
        self.assertIn(reason, message)
        for sensitive in (
            "PrivateGiven",
            "PrivateFamily",
            "PRIVATE01",
            "0000000101",
            "2008-09-01",
            "2026-07-12",
            "bad-private",
        ):
            self.assertNotIn(sensitive, message)

    def test_boundary_is_based_on_each_exit_date(self):
        row = self.example.copy()
        for raw, expected in (("20260815", 2026), ("2026-08-16", 2027), ("07/12/2026", 2026)):
            with self.subTest(raw=raw):
                row[5] = raw
                self.write(self.early, row)
                self.assertEqual(reporting_year([self.early]), expected)

    def test_mixed_rows_reference_the_first_and_conflicting_positions(self):
        self.write(self.early, [], self.example)
        second = self.example.copy()
        second[5], second[6] = "20260816", "120"
        self.write(self.late, second)
        with self.assertRaises(ValidationError) as caught:
            reporting_year([self.early, self.late])
        self.check_reason(caught, "late.csv", 2, "different reporting cycle")
        self.assertIn("early.csv, Row 3", str(caught.exception))

    def test_no_student_rows_allows_header_check_only(self):
        self.write(self.early)
        completed = subprocess.run(
            [sys.executable, "-m", "calpads_to_nsc", str(self.early), "--check-headers"], capture_output=True, text=True
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        with self.assertRaisesRegex(ValidationError, "No CALPADS student rows"):
            reporting_year([self.early])

    def test_explicit_year_validates_the_roster(self):
        self.write(self.early, [], self.example)
        with self.assertRaises(ValidationError) as caught:
            reporting_year([self.early], 2025)
        self.check_reason(caught, "early.csv", 3, "outside the selected CALPADS cycle")

    def test_missing_or_malformed_exit_dates_and_identifiers_have_distinct_messages(self):
        for field, value, reason in (
            (5, "", "Missing enrollment exit date"),
            (5, "202661", "Invalid enrollment exit date"),
            (1, "", "Missing SSID"),
            (1, "bad-private", "Malformed SSID"),
            (8, "bad-private", "Unexpected exit reason"),
            (6, "bad-private", "Completion status"),
        ):
            with self.subTest(field=field, reason=reason):
                row = self.example.copy()
                row[field] = value
                self.write(self.early, [], row)
                with self.assertRaises(ValidationError) as caught:
                    list(eligible_graduates([self.early], 2026, {"100": "Regular Diploma"}, set()))
                self.check_reason(caught, "early.csv", 3, reason)

    def test_duplicate_report_identifiers_show_both_files_and_rows(self):
        self.write(self.early, [], self.example)
        self.write(self.late, self.example)
        with self.assertRaises(ValidationError) as caught:
            list(eligible_graduates([self.early, self.late], 2026, {"100": "Regular Diploma"}, set()))
        self.check_reason(caught, "late.csv", 2, "Duplicate SSID")
        self.assertIn("early.csv, Row 3", str(caught.exception))

    def test_supplement_errors_are_located_without_repeating_private_fields(self):
        for row, reason in (
            (["", "", "Y"], "Missing SSID"),
            (["bad-private", "", "Y"], "Malformed SSID"),
            (["0000000101", "", "bad-private"], "FERPABlock"),
            (["0000000101", "", "Y", "bad-private"], "number of fields"),
        ):
            with self.subTest(reason=reason):
                self.write(self.late, [], row, headers=("SSID", "GraduationDate", "FERPABlock"))
                with self.assertRaises(ValidationError) as caught:
                    read_supplement(self.late)
                self.check_reason(caught, "late.csv", 3, reason)

    def test_truncated_records_and_invalid_csv_quoting_are_located(self):
        self.write(self.early, [], self.example[:-1])
        with self.assertRaises(ValidationError) as caught:
            reporting_year([self.early])
        self.check_reason(caught, "early.csv", 3, "Incomplete report row")
        self.write(self.early, self.example)
        with self.early.open("a", encoding="utf-8") as stream:
            stream.write('"bad-private')
        with self.assertRaises(ValidationError) as caught:
            reporting_year([self.early])
        self.check_reason(caught, "early.csv", 3, "Cannot read this row")

    def test_invalid_utf8_does_not_echo_data_or_misidentify_the_header(self):
        self.write(self.early, self.example)
        with self.early.open("ab") as stream:
            stream.write(b"bad-private,\xff\n")
        with self.assertRaises(ValidationError) as caught:
            reporting_year([self.early])
        self.check_reason(caught, "early.csv", 3, "Unreadable character")
