import csv
import tempfile
import unittest
from pathlib import Path

from calpads_to_nsc.generate import ValidationError
from calpads_to_nsc.programs import PROGRAM_FIELDS
from calpads_to_nsc.supplement import is_supplement_header, read_supplement


class ProgramSupplementTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name, "Supplement.csv")
        self.ssid = "0000000123"
        self.other = "0000000456"

    def read(self, headers, *rows):
        with self.path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(headers)
            writer.writerows(rows)
        return read_supplement(self.path)

    def test_all_seven_fields_land_in_the_specified_nsc_positions(self):
        headers = ["SSID", *PROGRAM_FIELDS]
        self.assertTrue(is_supplement_header(headers))
        data = self.read(headers, [self.ssid, "1", "2", "3", "1", "08", "yes", "AVID.AP_FAFSA"])
        self.assertEqual(data.program_columns, tuple(PROGRAM_FIELDS))
        self.assertEqual(
            data.program_fields(self.ssid), {18: "1", 19: "2", 20: "3", 21: "1", 23: "8", 24: "Y", 26: "AVID.AP_FAFSA"}
        )
        self.assertEqual(data.program_fields(self.other), {})
        self.assertFalse(data.has_date_column)
        self.assertFalse(data.has_block_column)
        self.assertFalse(data.dates)

    def test_blanks_stay_blank_and_zero_is_preserved(self):
        data = self.read(
            ["SSID", *PROGRAM_FIELDS],
            [self.ssid, "", "", "", "", "0", "no", ""],
            [self.other, "", "", "", "", "", "", ""],
        )
        self.assertEqual(data.program_fields(self.ssid), {23: "0", 24: "N"})
        self.assertEqual(data.program_fields(self.other), {})

    def test_program_code_preserves_freeform_label_and_supported_separators(self):
        for code in ("Arts & Sciences.AP_CTE", "A" * 50):
            with self.subTest(code=code):
                data = self.read(["SSID", "ProgramCode"], [self.ssid, code])
                self.assertEqual(data.program_fields(self.ssid), {26: code})

    def test_program_code_rejects_placeholder_without_printing_the_value(self):
        for placeholder in ("N/A", "n/a", " N/a "):
            with self.subTest(placeholder=placeholder):
                data = self.read(["SSID", "ProgramCode"], [self.ssid, placeholder])
                with self.assertRaises(ValidationError) as result:
                    data.program_fields(self.ssid)
                message = str(result.exception)
                for label in (self.path.name, "Row 2", "ProgramCode", "Leave this cell blank"):
                    self.assertIn(label, message)
                self.assertNotIn("n/a", message.casefold())
                self.assertNotIn(self.ssid, message)

    def test_fieldwise_merge_of_sparse_or_equivalent_rows_keeps_known_blocks_and_dates(self):
        data = self.read(
            ["SSID", "GraduationDate", "FERPABlock", "MathSemesters", "ProgramCode", "DualEnrollment"],
            [self.ssid, "20260610", "", "08", "", "Yes"],
            [self.ssid, "", "Y", "8", "AVID", "Y"],
        )
        self.assertEqual(data.program_fields(self.ssid), {23: "8", 24: "Y", 26: "AVID"})
        self.assertTrue(data.has_date_column)
        self.assertTrue(data.has_block_column)
        self.assertEqual(data.blocked, {self.ssid})
        self.assertIn((2, "20260610", "Compared"), data.dates[self.ssid])

    def test_conflicting_nonblank_program_entries_report_both_rows_without_private_values(self):
        data = self.read(
            ["SSID", "ProgramCode"],
            [self.ssid, "FIRST"],
            [self.other, "UNRELATED"],
            [self.ssid, "SECOND"],
        )
        with self.assertRaises(ValidationError) as result:
            data.program_fields(self.ssid)
        message = str(result.exception)
        for label in (self.path.name, "Row 2", "Row 4", "ProgramCode"):
            self.assertIn(label, message)
        for raw in (self.ssid, self.other, "FIRST", "SECOND", "UNRELATED"):
            self.assertNotIn(raw, message)

    def test_invalid_values_are_checked_only_for_requested_academic_records(self):
        for header, invalid in (
            ("Grade8MathAssessment", "4"),
            ("Grade8ELAAssessment", "0"),
            ("HighSchoolMathAssessment", "2.5"),
            ("HighSchoolELAAssessment", "N/A"),
            ("MathSemesters", "100"),
            ("MathSemesters", "-1"),
            ("DualEnrollment", "maybe"),
            ("ProgramCode", "A" * 51),
            ("ProgramCode", "bad-code"),
            ("ProgramCode", "bad—code"),
            ("ProgramCode", 'bad"code'),
            ("ProgramCode", "bad\tcode"),
            ("ProgramCode", "badcode\n"),
            ("ProgramCode", "=SUM(2,3)"),
        ):
            with self.subTest(header=header, invalid=invalid):
                data = self.read(["SSID", header], [self.ssid, ""], [self.other, invalid])
                self.assertEqual(data.program_fields(self.ssid), {})
                with self.assertRaises(ValidationError) as result:
                    data.program_fields(self.other)
                message = str(result.exception)
                for label in (self.path.name, "Row 3", header):
                    self.assertIn(label, message)
                self.assertNotIn(self.other, message)
                if len(invalid) > 3:
                    self.assertNotIn(invalid, message)

    def test_new_and_legacy_headers_use_the_same_rules(self):
        for headers in (["ssid", "program code"], ["SSID", "FERPABlock"], ["GraduationDate", "SSID"]):
            self.assertTrue(is_supplement_header(headers))
            self.read(headers)
        for headers in (["SSID"], ["SSID", "UnknownPrivateField"], ["SSID", "Program Code", "ProgramCode"]):
            with self.subTest(headers=headers), self.assertRaises(ValidationError) as result:
                self.read(headers)
            self.assertIn("Row 1", str(result.exception))
            self.assertNotIn("UnknownPrivateField", str(result.exception))
        self.assertFalse(is_supplement_header(["SchoolName", "SSID", "GraduationDate", "ProgramCode"]))
