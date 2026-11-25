import csv
import tempfile
import unittest
from pathlib import Path

from calpads_to_nsc.demographics import DEMOGRAPHIC_LABELS, demographic_values
from calpads_to_nsc.generate import DetailColumn
from calpads_to_nsc.reports import demographic_headers, report_records


BASE = ("School Name", "SSID", "Student Legal Name", "Birth Date", "Enrollment Exit Date", "School Completion Status")
DETAIL = ("Example High", "0000000001", "Sample, Alex", "2008-01-01", "2026-06-12", "100")


class DemographicsTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)

    def report(self, name, extra_headers=(), data=()):
        path = self.root / name
        with path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow(["Certified CALPADS 1.23"])
            writer.writerow((*BASE, *extra_headers))
            writer.writerow((*DETAIL, *data))
        return path

    def test_fields_map_from_report_headings_in_any_order(self):
        path = self.report(
            "report.csv",
            ("ELAS Status", "Ethnicity/Race", "Special Education Status", "Socio-Economically Disadvantaged", "Gender"),
            ("EL - English Learner", "Hispanic", "Eligible and Participating", "No", "Female"),
        )
        _, row = next(iter(report_records([path])))
        notices = set()
        values = demographic_values(row, include_disability=True, notices=notices)
        self.assertEqual(values, {15: "F", 16: "HI", 17: "N", 22: "Y", 25: "Y"})
        self.assertFalse(notices)

    def test_combined_race_categories(self):
        for raw, code in (
            ("American Indian/Alaska Native", "AM"),
            ("Am Ind/Alaska Nat", "AM"),
            ("Am Indian/Alskn Nat", "AM"),
            ("Asian", "AS"),
            ("Black / African American", "BL"),
            ("Black/African Am", "BL"),
            ("Hispanic or Latino", "HI"),
            ("Hawaiian/Pacific Islander", "PI"),
            ("Nat Hawaii/Pac Isl", "PI"),
            ("Nat Hwiin/Other Pac Islndr", "PI"),
            ("White", "WH"),
            ("Multiple", "MU"),
            ("Multi", "MU"),
        ):
            with self.subTest(raw=raw):
                values = demographic_values({"ethnicity": raw}, include_disability=False, notices=set())
                self.assertEqual(values, {DetailColumn.ETHNICITY: code})

    def test_missing_unknown_and_nonbinary_values_stay_blank_without_disclosing_values(self):
        row = {
            "gender": "Nonbinary",
            "ethnicity": "a-private-cell-value",
            "english_learner": "TBD",
            "economic_disadvantage": "",
            "disability": None,
        }
        notices = set()
        self.assertEqual(demographic_values(row, include_disability=True, notices=notices), {})
        self.assertEqual(len(notices), 3)
        for text in ("Nonbinary", "a-private-cell-value", "TBD"):
            self.assertNotIn(text, " ".join(notices))

    def test_known_english_status_and_explicit_special_education_choice(self):
        for elas, code in (
            ("English Learner (EL)", "Y"),
            ("IFEP", "N"),
            ("RFEP - Reclassified Fluent English Proficient", "N"),
            ("EO", "N"),
        ):
            with self.subTest(elas=elas):
                row = {"english_learner": elas, "disability": "No"}
                values = demographic_values(row, include_disability=False, notices=set())
                self.assertEqual(values, {DetailColumn.ENGLISH_LEARNER: code})
                values = demographic_values(row, include_disability=True, notices=set())
                self.assertEqual(values[DetailColumn.DISABILITY], "N")
        for unknown in ("", "2", "Unknown", "Eligible and Not Participating"):
            selected = demographic_values({"disability": unknown}, include_disability=True, notices=set())
            self.assertNotIn(DetailColumn.DISABILITY, selected)

    def test_discovery_reads_only_headers_and_combines_available_fields(self):
        first = self.report("first.csv", ("GenderCode", "EthnicityRace"))
        second = self.report("second.csv", ("ELAS", "SpecialEd", "SED"))
        with second.open("a", encoding="utf-8") as handle:
            handle.write('"unterminated body')
        self.assertEqual(demographic_headers([first, second]), set(DEMOGRAPHIC_LABELS))

    def test_repeated_headers_and_missing_trailing_values_leave_only_those_fields_blank(self):
        path = self.report(
            "report.csv", ("Gender", "Student Gender Code", "ELAS", "GenderCode", "SED"), ("M", "F", "EL")
        )
        self.assertEqual(demographic_headers([path]), {"english_learner", "economic_disadvantage"})
        _, row = next(iter(report_records([path])))
        self.assertIsNone(row["economic_disadvantage"])
        notices = set()
        self.assertEqual(demographic_values(row, include_disability=False, notices=notices), {22: "Y"})
        self.assertIn("missing or repeated Gender heading", " ".join(notices))
