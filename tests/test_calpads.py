import csv
import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from calpads_to_nsc.calpads import main
from calpads_to_nsc.crosswalk import create
from calpads_to_nsc.reference_sheet import read_reference


class CalpadsTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.reference = self.root / "reference.xlsx"
        self.review = self.root / "schools.xlsx"
        self.campuses = {"Sample HS": ("012345", "SAMPLE HIGH")}
        self.prepare_school_inputs()
        self.report = self.root / "extract.txt"
        self.out = self.root / "graduates.txt"
        self.headers = [
            "SchoolName",
            "SSID",
            "StudentName",
            "LocalID",
            "BirthDate",
            "StudentEnrollmentExitDate",
            "StudentExitReason",
            "StudentCompletionStatus",
        ]
        self.record = [
            "Sample HS",
            "0123456789",
            "Example-Smith, Renée M.",
            "0001",
            "2008-04-03",
            "06/03/2026",
            "E230",
            "100 - HS Diploma",
        ]

    def prepare_school_inputs(self, *, pending=(), rehearsal=False):
        book = Workbook()
        for row in (
            ("Organization Name:", "Example District"),
            ("Service Agreement/Account Number:", "0012"),
            ("Total High School Count:", len(self.campuses)),
            ("ACT Code", "HS Operating Status", "High School Name"),
        ):
            book.active.append(row)
        for code, name in self.campuses.values():
            book.active.append([code, "OPEN", name])
        book.save(self.reference)
        book.close()
        self.review.unlink(missing_ok=True)
        digest = hashlib.sha256(self.reference.read_bytes()).hexdigest()
        create(self.review, set(self.campuses) | set(pending), read_reference(self.reference), digest, rehearsal)
        book = load_workbook(self.review)
        for row in book["Match schools"].iter_rows(min_row=10, min_col=2, max_col=4):
            if row[0].value in self.campuses:
                code, name = self.campuses[row[0].value]
                row[2].value = f"{name} ({code})"
        book.save(self.review)
        book.close()

    def write(self, *rows):
        with self.report.open("w", newline="", encoding="utf-8") as handle:
            csv.writer(handle, delimiter="^").writerows([self.headers, *rows])

    def run_cli(self, *opts, school_inputs=True, confirm_blocks=True):
        common = [str(self.report), "--year", "2026", "--date", "20260630", "--no-prompt"]
        if confirm_blocks and "--output" in opts:
            common.append("--confirm-no-ferpa-blocks")
        if school_inputs and "--check-headers" not in opts:
            common.extend(["--reference-sheet", str(self.reference), "--crosswalk", str(self.review)])
        return subprocess.run([sys.executable, "-B", "-m", "calpads_to_nsc", *common, *opts], capture_output=True)

    def review_decisions(self, answers, *opts):
        args = [
            "prepare-nsc-graduates",
            str(self.report),
            "--reference-sheet",
            str(self.reference),
            "--crosswalk",
            str(self.review),
            "--date",
            "20260630",
            "--review-decisions",
            "--dry-run",
            *opts,
        ]
        output = io.StringIO()
        with (
            redirect_stdout(output),
            patch.object(sys, "argv", args),
            patch.object(sys.stdin, "isatty", return_value=True),
        ):
            with (
                patch.object(output, "isatty", return_value=True),
                patch("builtins.input", side_effect=answers) as prompt,
            ):
                main()
        self.assertEqual(prompt.call_count, len(answers))
        return output.getvalue()

    def test_completion_mapping_requires_a_decision_for_unrecognized_codes(self):
        self.record[-1] = "102"
        self.write(self.record)
        result = self.run_cli("--dry-run")
        self.assertIn(b"Completion code 102", result.stderr)
        settings = self.root / "completion-policy.json"
        settings.write_text(json.dumps({"completion_types": {"102": "Regular Diploma"}}), encoding="utf-8")
        result = self.run_cli("--settings", str(settings), "--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_review_shows_blank_ferpa_and_calpads_fallback(self):
        self.write(self.record)
        supplement = self.root / "supplement.csv"
        for value, blank_notice in (("", True), ("N", False), ("Y", False)):
            with self.subTest(value=value):
                supplement.write_text(f"SSID,GraduationDate,FERPABlock\n{self.record[1]},,{value}\n")
                result = self.run_cli("--supplement", str(supplement), "--dry-run")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(b"no Y or N entries" in result.stdout, blank_notice)
                self.assertIn(b"Kept CALPADS exit dates", result.stdout)
                self.assertIn(b"Completion code 100", result.stdout)
                self.assertNotIn(self.record[1].encode(), result.stdout)

    def test_text_and_trailer(self):
        excluded = self.record.copy()
        excluded[1], excluded[-1] = "1111111111", "120"
        self.write(self.record, excluded)
        self.assertEqual(self.run_cli("--dry-run").returncode, 0)
        self.assertFalse(self.out.exists())
        result = self.run_cli("--output", str(self.out))
        self.assertEqual(result.returncode, 0, result.stderr)
        expected = (
            "PH3\t0012\tExample District\tP\t20260630\t2026\r\n"
            "PD3\tNO SSN\tRenee\tM\tExample-Smith\t\t\t\t20080403\t0001\tRegular Diploma\t"
            "20260603\tN\tSAMPLE HIGH\t012345" + "\t" * 13 + "ED\r\nPT3\t3\r\n"
        )
        self.assertEqual(self.out.read_bytes(), expected.encode())
        self.assertNotEqual(self.run_cli("--output", str(self.out)).returncode, 0)

    def test_optional_fields_have_the_same_28_column_layout_in_test_and_full_files(self):
        self.headers.extend(
            ["Gender", "Ethnicity/Race", "Socio-Economically Disadvantaged", "ELAS", "Special Education Status"]
        )
        self.write([*self.record, "Female", "Black or African American", "No", "EL", "Eligible and Participating"])
        supplement = self.root / "supplement.csv"
        supplement.write_text(
            "SSID,GraduationDate,FERPABlock,Grade8MathAssessment,Grade8ELAAssessment,HighSchoolMathAssessment,"
            "HighSchoolELAAssessment,MathSemesters,DualEnrollment,ProgramCode\n"
            f"{self.record[1]},2026-06-02,Y,1,2,3,1,08,No,AVID.AP\n",
            encoding="utf-8",
        )
        expected = (
            "PD3\tNO SSN\tRenee\tM\tExample-Smith\t\t\t\t20080403\t0001\tRegular Diploma\t"
            "20260602\tY\tSAMPLE HIGH\t012345\tF\tBL\tN\t1\t2\t3\t1\tY\t8\tN\tY\tAVID.AP\tED"
        ).split("\t")
        for mode in ("--test-output", "--output"):
            with self.subTest(mode=mode):
                output = self.root / f"{mode}.txt"
                result = self.run_cli(
                    "--supplement",
                    str(supplement),
                    "--include-demographics",
                    "--use-special-education-for-disability",
                    mode,
                    str(output),
                    confirm_blocks=False,
                )
                self.assertEqual(result.returncode, 0, result.stderr)
                with output.open(encoding="utf-8") as handle:
                    header, detail, trailer = csv.reader(handle, delimiter="\t")
                self.assertEqual(header, ["PH3", "0012", "Example District", "P", "20260630", "2026"])
                self.assertEqual(len(detail), 28)
                self.assertEqual(detail, expected)
                self.assertEqual(trailer, ["PT3", "3"])

    def test_program_only_supplement_works_with_or_without_calpads_demographics(self):
        self.headers.extend(["Gender", "Ethnicity/Race", "Special Education Status"])
        self.write([*self.record, "M", "PRIVATE_UNMAPPED_VALUE", "Y"])
        supplement = self.root / "programs.csv"
        supplement.write_text(
            f"SSID,DualEnrollment,ProgramCode\n{self.record[1]},Y,\n9999999999,N,OTHER_PROGRAM\n", encoding="utf-8"
        )
        for with_demographics in (False, True):
            with self.subTest(with_demographics=with_demographics):
                output = self.root / f"programs-{with_demographics}.txt"
                options = ["--include-demographics"] if with_demographics else []
                result = self.run_cli(*options, "--supplement", str(supplement), "--output", str(output))
                self.assertEqual(result.returncode, 0, result.stderr)
                messages = result.stdout + result.stderr
                self.assertNotIn(b"PRIVATE_UNMAPPED_VALUE", messages)
                if with_demographics:
                    self.assertIn(b"Race/ethnicity", messages)
                    self.assertIn(b"English learner", messages)
                with output.open(encoding="utf-8") as handle:
                    _, detail, trailer = csv.reader(handle, delimiter="\t")
                self.assertEqual(detail[11:13], ["20260603", "N"])
                self.assertEqual(detail[15:], ["M" if with_demographics else "", *([""] * 8), "Y", "", "", "ED"])
                self.assertEqual(trailer, ["PT3", "3"])
        comparison = self.root / "comparison.csv"
        result = self.run_cli("--supplement", str(supplement), "--compare-dates", str(comparison))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b"GraduationDate", result.stderr)
        self.assertFalse(comparison.exists())

    def test_workbook_maps_by_act_and_uses_the_reference_name(self):
        self.write(self.record)
        for size in (50, 51):
            with self.subTest(size=size):
                official = "W" * size
                self.campuses["Sample HS"] = ("012345", official)
                self.prepare_school_inputs()
                output = self.root / f"file_{size}.txt"
                result = self.run_cli("--output", str(output))
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(
                    output.read_text(encoding="utf-8").splitlines()[1].split("\t")[13:15], [official, "012345"]
                )
                self.assertEqual(b"Warning" in result.stderr, size > 50)
        book = load_workbook(self.reference)
        book.active["A5"] = "012346"
        book.save(self.reference)
        book.close()
        result = self.run_cli("--dry-run")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b"reference sheet changed", result.stderr)

    def test_generation_requires_a_reviewed_workbook_and_real_reference(self):
        self.write(self.record)
        for mode in (["--dry-run"], ["--test-output", str(self.out)], ["--output", str(self.out)]):
            with self.subTest(mode=mode[0]):
                result = self.run_cli("--reference-sheet", str(self.reference), *mode, school_inputs=False)
                self.assertNotEqual(result.returncode, 0)
                self.assertIn(b"--crosswalk", result.stderr)
                result = self.run_cli("--crosswalk", str(self.review), *mode, school_inputs=False)
                self.assertNotEqual(result.returncode, 0)
                self.assertIn(b"--reference-sheet", result.stderr)
                self.assertFalse(self.out.exists())

    def test_settings_accept_only_district_completion_decisions(self):
        self.write(self.record)
        supplement = self.root / "supplement.csv"
        supplement.write_text("SSID,GraduationDate\n", encoding="utf-8")
        policy = self.root / "district.json"
        audit = self.root / "comparison.csv"
        policy.write_text(json.dumps({"schools": {}, "completion_types": {}}), encoding="utf-8")
        for mode in (["--output", str(self.out)], ["--supplement", str(supplement), "--compare-dates", str(audit)]):
            result = self.run_cli("--settings", str(policy), *mode)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn(b"only the completion_types", result.stderr)
        self.assertFalse(self.out.exists())
        self.assertFalse(audit.exists())

    def test_portal_sample_includes_each_school_once_and_rebuilds_period_and_trailer(self):
        self.campuses["Second HS"] = ("012346", "SECOND HIGH")
        self.prepare_school_inputs()
        earlier = self.record.copy()
        earlier[1], earlier[3], earlier[5] = "1111111111", "0002", "09/01/2025"
        second = self.record.copy()
        second[0], second[1], second[3] = "Second HS", "2222222222", "0003"
        self.write(self.record, earlier, second)
        portal = self.root / "sample.txt"
        result = self.run_cli("--output", str(self.out))
        self.assertEqual(result.returncode, 0, result.stderr)
        result = self.run_cli("--test-output", str(portal))
        self.assertEqual(result.returncode, 0, result.stderr)
        with self.out.open(encoding="utf-8") as handle:
            rows = list(csv.reader(handle, delimiter="\t"))
        with portal.open(encoding="utf-8") as handle:
            sample = list(csv.reader(handle, delimiter="\t"))
        self.assertEqual(rows[0][-1], "2025-2026")
        self.assertEqual(sample[0][-1], "2026")
        self.assertEqual(sample[1:-1], [rows[1], rows[3]])
        self.assertEqual(sample[-1], ["PT3", "4"])

    def test_excel_dates_and_block(self):
        self.report = self.root / "example.xlsx"
        book = Workbook()
        book.active.append(self.headers + ["GraduationDate"])
        self.record[4] = datetime(2008, 4, 3)
        book.active.append(self.record + [datetime(2026, 6, 2)])
        book.save(self.report)
        book.close()
        supplement = self.root / "supplement.csv"
        supplement.write_text(f"SSID,GraduationDate,FERPABlock\n{self.record[1]},,Y\n", encoding="utf-8")
        self.assertEqual(self.run_cli("--supplement", str(supplement), "--output", str(self.out)).returncode, 0)
        self.assertIn(b"20260603\tY", self.out.read_bytes())

    def test_reject_incomplete_or_ambiguous(self):
        for index, value in (
            (0, "Missing school"),
            (2, "Unstructured Name"),
            (5, "8/16/2026"),
            (6, "E155"),
            (7, "102"),
            (1, ""),
        ):
            row = self.record.copy()
            row[index] = value
            self.write(row)
            result = self.run_cli("--output", str(self.out))
            self.assertNotEqual(result.returncode, 0)
            self.assertNotIn(self.record[1].encode(), result.stderr)
            self.assertFalse(self.out.exists())
        self.write(self.record, self.record)
        self.assertNotEqual(self.run_cli("--dry-run").returncode, 0)
        self.write(self.record)
        self.assertEqual(self.run_cli("--dry-run").returncode, 0)

    def test_duplicate_local_id_reports_both_rows_without_printing_the_id(self):
        duplicate = self.record.copy()
        duplicate[1] = "1123456789"
        self.write(self.record, [], duplicate)
        result = self.run_cli("--output", str(self.out))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b"extract.txt", result.stderr)
        self.assertIn(b"Row 4", result.stderr)
        self.assertIn(b"Duplicate local student ID", result.stderr)
        self.assertIn(b"extract.txt, Row 2", result.stderr)
        self.assertNotIn(b"0001", result.stderr)
        self.assertFalse(self.out.exists())

    def test_malformed_birth_date_and_local_id_identify_the_field_and_location(self):
        for col, reason in ((4, b"Invalid birth date"), (3, b"Malformed local student ID")):
            with self.subTest(column=col):
                row = self.record.copy()
                row[col] = "PRIVATE VALUE"
                self.write([], row)
                result = self.run_cli("--dry-run")
                self.assertNotEqual(result.returncode, 0)
                self.assertIn(self.report.name.encode(), result.stderr)
                self.assertIn(b"Row 3", result.stderr)
                self.assertIn(reason, result.stderr)
                self.assertNotIn(b"PRIVATE VALUE", result.stdout + result.stderr)
                self.assertNotIn(self.record[1].encode(), result.stdout + result.stderr)

    def test_comparison_rejects_a_fallback_date_before_birth_with_a_located_error(self):
        self.record[4] = "2027-01-03"
        self.write([], self.record)
        dates = self.root / "supplement.csv"
        dates.write_text("SSID,GraduationDate\n", encoding="utf-8")
        target = self.root / "comparison.csv"
        result = self.run_cli("--compare-dates", str(target), "--supplement", str(dates))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(target.exists())
        self.assertIn(self.report.name.encode(), result.stderr)
        self.assertIn(b"Row 3", result.stderr)
        self.assertNotIn(b"2027-01-03", result.stderr)

    def test_enriched_report_prompts_for_a_supplement_to_preserve_known_blocks(self):
        self.headers.append("FERPA Block")
        self.write(self.record + ["Y"])
        result = self.run_cli("--output", str(self.out))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b"district supplement", result.stderr)
        self.assertNotIn(self.record[1].encode(), result.stderr)
        self.assertFalse(self.out.exists())

    def test_bad_name_points_to_source_row_without_disclosing_student_data(self):
        self.record[2] = "9@_#, PrivateGiven"
        self.write([], self.record)
        result = self.run_cli("--output", str(self.out))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(self.report.name.encode(), result.stderr)
        self.assertIn(b"Row 3", result.stderr)
        self.assertIn(b"last name", result.stderr)
        self.assertNotIn(b"Private", result.stderr)
        self.assertNotIn(self.record[1].encode(), result.stderr)
        self.assertFalse(self.out.exists())

    def test_combined_names_cleaned_automatically_in_generated_file_only(self):
        self.record[2] = "O’Exámple9_Sm!ith‑Lee， Gr3ace A."
        self.write(self.record)
        original = self.report.read_bytes()
        for opts in (["--dry-run"], ["--output", str(self.out)]):
            result = self.run_cli(*opts)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(result.stdout.count(b"formatted automatically"), 1)
            self.assertNotIn(self.record[1].encode(), result.stdout + result.stderr)
            self.assertNotIn(b"Exampl", result.stdout + result.stderr)
            self.assertEqual(self.report.read_bytes(), original)
        row = self.out.read_text(encoding="utf-8").splitlines()[1].split("\t")
        self.assertEqual(row[2:5], ["Grace", "A", "O'ExampleSmith-Lee"])

    def test_combined_supplement_merges_dates_and_blocks_for_calpads_graduates(self):
        records = []
        for n in range(1, 8):
            row = self.record.copy()
            row[1], row[3] = f"{n:010d}", str(n)
            records.append(row)
        self.write(*records)
        supplement = self.root / "supplement.csv"
        supplement.write_text(
            "SSID,GraduationDate,FERPABlock\n"
            "0000000001,20260601,Y\n0000000001,20260602,N\n"
            "0000000002,20260601,Y\n0000000003,20260601,Y\n"
            "0000000004,,\n0000000005,20260601,N\n0000000006,,Y\n"
            "0000000007,20260601,N\n0000000007,20260601,Y\n9999999999,20260601,Y\n",
            encoding="utf-8",
        )
        unresolved = self.run_cli("--supplement", str(supplement), "--output", str(self.out))
        self.assertNotEqual(unresolved.returncode, 0)
        self.assertIn(b"Rows 2, 3", unresolved.stderr)
        self.assertFalse(self.out.exists())
        self.review_decisions(["1"], "--supplement", str(supplement))
        result = self.run_cli("--supplement", str(supplement), "--output", str(self.out))
        self.assertEqual(result.returncode, 0, result.stderr)
        rows = [line.split("\t") for line in self.out.read_text(encoding="utf-8").splitlines()[1:-1]]
        self.assertEqual([row[9] for row in rows], list("1234567"))
        self.assertEqual(
            [(row[11], row[12]) for row in rows],
            [
                ("20260603", "Y"),
                ("20260601", "Y"),
                ("20260601", "Y"),
                ("20260603", "N"),
                ("20260601", "N"),
                ("20260603", "Y"),
                ("20260601", "Y"),
            ],
        )
        comparison = self.root / "comparison.csv"
        result = self.run_cli("--supplement", str(supplement), "--compare-dates", str(comparison))
        self.assertEqual(result.returncode, 0, result.stderr)
        with comparison.open(encoding="utf-8") as handle:
            groups = list(csv.DictReader(handle))
        self.assertEqual(
            {g["Status"]: int(g["Students"]) for g in groups},
            {
                "Confirmed CALPADS date for conflicting supplemental diploma dates": 1,
                "Blank diploma date in district supplement": 2,
                "Compared": 4,
            },
        )

    def test_invalid_supplemental_block_explains_allowed_values_without_student_data(self):
        self.write(self.record)
        supplement = self.root / "supplement.csv"
        supplement.write_text("FERPABlock,SSID,GraduationDate\nPRIVATE_VALUE,0123456789,20260601\n", encoding="utf-8")
        result = self.run_cli("--supplement", str(supplement), "--output", str(self.out))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(self.out.exists())
        self.assertIn(b"FERPA", result.stderr)
        self.assertIn(b"Y", result.stderr)
        self.assertIn(b"N", result.stderr)
        self.assertNotIn(b"0123456789", result.stdout + result.stderr)
        self.assertNotIn(b"PRIVATE_VALUE", result.stdout + result.stderr)

    def test_misspelled_or_duplicate_optional_block_header_cannot_silently_drop_a_block(self):
        self.write(self.record)
        supplement = self.root / "supplement.csv"
        for columns, values in (("FERPABlok", "Y"), ("FERPABlock,FERPA Block", "N,Y")):
            with self.subTest(columns=columns):
                supplement.write_text(
                    f"SSID,GraduationDate,{columns}\n0123456789,20260601,{values}\n", encoding="utf-8"
                )
                result = self.run_cli("--supplement", str(supplement), "--output", str(self.out))
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse(self.out.exists())
                self.assertNotIn(b"0123456789", result.stdout + result.stderr)

    def test_cannot_ignore_supplemental_block_with_a_missing_or_invalid_ssid(self):
        self.write(self.record)
        supplement = self.root / "supplement.csv"
        for invalid in ("", "BAD_PRIVATE_ID"):
            with self.subTest(ssid_shape="missing" if not invalid else "invalid"):
                supplement.write_text(f"SSID,GraduationDate,FERPABlock\n{invalid},,Y\n", encoding="utf-8")
                result = self.run_cli("--supplement", str(supplement), "--output", str(self.out))
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse(self.out.exists())
                self.assertIn(b"SSID", result.stderr)
                self.assertNotIn(b"BAD_PRIVATE_ID", result.stdout + result.stderr)

    def test_saved_diploma_exception_is_scoped_to_evidence_and_revisable_across_modes(self):
        self.write(self.record)
        supp = self.root / "district.csv"

        def write_supplement(award):
            supp.write_text(f"SSID,GraduationDate,FERPABlock\n{self.record[1]},{award},Y\n", encoding="utf-8")

        write_supplement("2025-08-15")
        opts = ["--supplement", str(supp)]
        sample = self.root / "test.txt"
        compare = self.root / "compare.csv"
        for mode in [("--dry-run",), ("--test-output", str(sample)), ("--output", str(self.out))]:
            result = self.run_cli(*opts, *mode)
            self.assertNotEqual(result.returncode, 0)
            for private in (self.record[1], "2025-08-15", self.record[2].split(",")[0]):
                self.assertNotIn(private.encode("ascii"), result.stderr)
            self.assertIn(b"extract.txt", result.stderr)
            self.assertIn(b"district.csv", result.stderr)
        self.assertFalse(any(p.exists() for p in (sample, compare, self.out)))
        screen = self.review_decisions(["2"], *opts)
        self.assertIn("outside", screen)
        db = self.review.with_suffix(".decisions.json")
        if os.name == "posix":
            self.assertEqual(db.stat().st_mode & 0o777, 0o600)
        for private in (self.record[1], self.record[2], "20250815", "2008-04-03"):
            self.assertNotIn(private.encode(), db.read_bytes())
        for mode, path in [("--test-output", sample), ("--output", self.out)]:
            result = self.run_cli(*opts, mode, str(path))
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn(b"20250815\tY", path.read_bytes())
        self.review_decisions(["1"], *opts)
        result = self.run_cli(*opts, "--compare-dates", str(compare))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("CALPADS,2026-06-03", compare.read_text(encoding="utf-8"))
        write_supplement("2025-08-14")
        self.assertNotEqual(self.run_cli(*opts, "--dry-run").returncode, 0)
        write_supplement("2025-08-16")
        self.assertEqual(self.run_cli(*opts, "--dry-run").returncode, 0)

    def test_completion_decisions_require_confirmation_and_apply_to_just_the_selected_cycle(self):
        adult = self.record.copy()
        adult[1], adult[3], adult[-1] = "0000000001", "002", "250 - Adult Ed."
        self.write(self.record, adult)
        screen = self.review_decisions(["9", "no", "9", "EXCLUDE"])
        self.assertIn("Adult Education High School Diploma", screen)
        self.assertIn("Selection unchanged", screen)
        self.assertEqual(self.run_cli("--output", str(self.out)).returncode, 0)
        self.assertEqual(len(self.out.read_bytes().splitlines()), 3)
        self.review_decisions(["7", "APPLY"])
        sample = self.root / "another.txt"
        self.assertEqual(self.run_cli("--output", str(sample)).returncode, 0)
        self.assertEqual(len(sample.read_bytes().splitlines()), 4)
        self.record[5] = adult[5] = "06/05/2027"
        self.write(self.record, adult)
        result = self.run_cli("--year", "2027", "--date", "20271117", "--dry-run")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b"Completion code 250", result.stderr)

    def test_use_calpads_exit_dates_still_applies_supplemental_ferpa_blocks(self):
        self.headers.extend(["Gender", "Special Education Status"])
        self.write([*self.record, "Female", "Y"])
        supplement = self.root / "supplement.csv"
        supplement.write_text(
            "FERPABlock,GraduationDate,SSID,DualEnrollment,ProgramCode\nY,20260601,0123456789,Y,AVID\n",
            encoding="utf-8",
        )
        result = self.run_cli(
            "--supplement", str(supplement), "--calpads-exit-dates", "--include-demographics", "--output", str(self.out)
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        details = self.out.read_text(encoding="utf-8").splitlines()[1].split("\t")
        self.assertEqual(details[11:13], ["20260603", "Y"])
        self.assertEqual(details[15:], ["F", *([""] * 11), "ED"])

    def test_comparison_matches_the_supplemental_dates_used_in_column_l(self):
        self.headers.append("DiplomaDate")
        self.record.append("2026-05-30")
        rows = []
        for n, exit_date, completion in [
            (1, "06/03/2026", "100"),
            (2, "06/03/2026", "100"),
            (3, "06/02/2026", "100"),
            (4, "06/04/2026", "100"),
            (5, "06/03/2026", "120"),
        ]:
            row = self.record.copy()
            row[1], row[3], row[5], row[7] = f"{n:010d}", str(n), exit_date, completion
            rows.append(row)
        self.write(*rows)
        dates = self.root / "dates.tsv"
        dates.write_text(
            "GRADUATION_DATE\tSSID\n2026-06-01\t0000000001\n6/1/2026\t0000000002\n"
            "20260604\t0000000003\n2026-06-04\t0000000004\n\t9999999999\n",
            encoding="utf-8-sig",
        )
        report = self.root / "comparison.csv"
        opts = ["--supplement", str(dates)]
        audit = self.run_cli(*opts, "--compare-dates", str(report), school_inputs=False)
        self.assertEqual(audit.returncode, 0, audit.stderr)
        with report.open(encoding="utf-8") as handle:
            pairs = list(csv.DictReader(handle))
        self.assertEqual(
            {(p["CALPADSExitDate"], p["SupplementGraduationDate"], p["DaysDifference"], p["Students"]) for p in pairs},
            {
                ("2026-06-03", "2026-06-01", "2", "2"),
                ("2026-06-02", "2026-06-04", "-2", "1"),
                ("2026-06-04", "2026-06-04", "0", "1"),
            },
        )
        self.assertTrue(all(p["Status"] == "Compared" for p in pairs))
        if os.name == "posix":
            self.assertEqual(report.stat().st_mode & 0o777, 0o600)
        self.assertNotIn("0000000001", report.read_text(encoding="utf-8"))
        self.assertNotEqual(self.run_cli(*opts, "--compare-dates", str(report)).returncode, 0)
        self.assertEqual(self.run_cli(*opts, "--dry-run").returncode, 0)
        built = self.run_cli(*opts, "--output", str(self.out))
        self.assertEqual(built.returncode, 0, built.stderr)
        details = [line.split("\t") for line in self.out.read_text(encoding="utf-8").splitlines()[1:-1]]
        self.assertEqual([d[11] for d in details], ["20260601", "20260601", "20260604", "20260604"])
        self.assertEqual([d[12] for d in details], ["N"] * 4)

    def test_optional_dates_choose_one_distinct_date_and_fall_back_locally(self):
        self.headers.append("GraduationDate")
        self.record.append("2026-05-31")
        rows = []
        for n in range(1, 12):
            row = self.record.copy()
            row[1], row[3] = f"{n:010d}", str(n)
            rows.append(row)
        self.write(*rows)
        dates = self.root / "dates.csv"
        dates.write_text(
            "SSID,GraduationDate\n0000000001,\n0000000002,2026-02-31\n0000000003,2026061\n"
            "0000000004,2026-06-01\n0000000004,2026-06-01\n"
            "0000000006,06/01/26\n0000000007,UNREADABLE_SOURCE_DATE\n"
            "0000000008,\n0000000008,06/01/2026\n0000000009,20260601\n0000000009,20260602\n"
            "0000000010,20070403\n0000000011,20260701\n9999999999,20260601\n",
            encoding="utf-8",
        )
        report = self.root / "comparison.csv"
        opts = ["--supplement", str(dates)]
        self.review_decisions(["1"], *opts)
        audit = self.run_cli(*opts, "--compare-dates", str(report))
        self.assertEqual(audit.returncode, 0, audit.stderr)
        with report.open(encoding="utf-8") as handle:
            problems = list(csv.DictReader(handle))
        self.assertEqual(
            {p["Status"]: int(p["Students"]) for p in problems},
            {
                "Blank diploma date in district supplement": 1,
                "Invalid date in diploma export": 3,
                "Compared": 2,
                "SSID missing from district supplement": 1,
                "Confirmed CALPADS date for conflicting supplemental diploma dates": 1,
                "Two-digit year in diploma export": 1,
                "Supplemental date outside birth-to-transmission range": 2,
            },
        )
        for p in problems:
            if p["Status"] != "Compared":
                self.assertEqual((p["DateSource"], p["NSCGraduationDate"]), ("CALPADS", "2026-06-03"))
        result = self.run_cli(*opts, "--output", str(self.out))
        self.assertEqual(result.returncode, 0, result.stderr)
        details = [line.split("\t") for line in self.out.read_text(encoding="utf-8").splitlines()[1:-1]]
        self.assertEqual([r[11] for r in details], ["20260601" if i in (4, 8) else "20260603" for i in range(1, 12)])

    def test_partial_rehearsal_checks_included_graduates_and_supplement(self):
        pending = self.record.copy()
        pending[0], pending[1], pending[-1] = "Pending Campus", "0123456788", "unknown"
        self.write(pending, self.record)
        self.prepare_school_inputs(pending=["Pending Campus"], rehearsal=True)
        supplement = self.root / "supplement.csv"
        for ssid, expected_block in [(pending[1], "N"), (self.record[1], "Y"), ("0000000000", "N")]:
            with self.subTest(block=expected_block):
                supplement.write_text(f"SSID,GraduationDate,FERPABlock\n{ssid},,Y\n", encoding="utf-8")
                output = self.root / f"test_{ssid}.txt"
                result = self.run_cli(
                    "--defer-school", "Pending Campus", "--supplement", str(supplement), "--test-output", str(output)
                )
                self.assertEqual(result.returncode, 0, result.stderr)
                with output.open(encoding="utf-8") as handle:
                    rows = list(csv.reader(handle, delimiter="\t"))
                self.assertEqual([row[0] for row in rows], ["PH3", "PD3", "PT3"])
                self.assertEqual(rows[1][12:15], [expected_block, "SAMPLE HIGH", "012345"])
                self.assertEqual(rows[-1], ["PT3", "3"])

    def test_folder_combines_and_validates_across_schools(self):
        batch = self.root / "reports"
        batch.mkdir()
        self.report = batch / "a.txt"
        self.write(self.record)
        second = self.record.copy()
        second[0:2] = ["Another HS", "0123456788"]
        second[3] = "0002"
        book = Workbook()
        book.active.append(self.headers)
        book.active.append(second)
        book.save(batch / "b.xlsx")
        self.report = batch
        self.campuses["Another HS"] = ("012346", "ANOTHER HIGH")
        self.prepare_school_inputs()
        self.assertEqual(self.run_cli("--check-headers").returncode, 0)
        fresh = self.root / "new.xlsx"
        generate = ("--reference-sheet", str(self.reference), "--create-crosswalk", str(fresh))
        result = self.run_cli(*generate, school_inputs=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        workbook = load_workbook(fresh)
        school_rows = list(workbook["Match schools"].values)
        self.assertEqual({row[1] for row in school_rows if row[1] in self.campuses}, set(self.campuses))
        self.assertNotIn(self.record[1], str(school_rows))
        workbook.close()
        self.assertNotEqual(self.run_cli(*generate, school_inputs=False).returncode, 0)
        self.assertEqual(self.run_cli("--output", str(self.out)).returncode, 0)
        self.assertEqual(
            [line.split("\t")[0] for line in self.out.read_text(encoding="utf-8").splitlines()],
            ["PH3", "PD3", "PD3", "PT3"],
        )
        self.assertTrue(self.out.read_bytes().endswith(b"PT3\t4\r\n"))
        book.active["B2"] = self.record[1]
        book.save(batch / "b.xlsx")
        self.assertNotEqual(self.run_cli("--dry-run").returncode, 0)

    def test_folder_preflight_accepts_only_reports_and_checks_later_headers(self):
        folder = self.root / "inputs"
        folder.mkdir()
        self.report = folder / "first.txt"
        self.write(self.record)
        self.report = folder
        (folder / "z.txt").write_text("wrong^header\n", encoding="utf-8")
        result = self.run_cli("--check-headers")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b"z.txt", result.stderr)
        self.assertIn(b"Row 1", result.stderr)
        (folder / "archive.zip").touch()
        self.assertNotEqual(self.run_cli("--check-headers").returncode, 0)

    def test_preview_groups_pending_date_and_completion_without_prompting_or_persisting(self):
        normal = self.record.copy()
        uncertain = self.record.copy()
        uncertain[1], uncertain[3], uncertain[-1] = "0000000001", "2", "250"
        conflict = self.record.copy()
        conflict[1], conflict[3] = "0000000002", "3"
        self.write(normal, uncertain, conflict)
        dates = self.root / "dates.csv"
        dates.write_text("SSID,GraduationDate\n0000000001,20250815\n0000000002,20260615\n0000000002,20260616\n")
        target = self.root / "preview.csv"
        argv = [
            "prepare-nsc-graduates",
            str(self.report),
            "--reference-sheet",
            str(self.reference),
            "--crosswalk",
            str(self.review),
            "--date",
            "20260630",
            "--supplement",
            str(dates),
            "--guide-session",
            "--compare-dates",
            str(target),
        ]
        screen = io.StringIO()
        with (
            patch.object(sys, "argv", argv),
            patch.object(sys.stdin, "isatty", return_value=True),
            redirect_stdout(screen),
            patch.object(screen, "isatty", return_value=True),
            patch("builtins.input", side_effect=AssertionError("comparison prompted")),
        ):
            main()
        with target.open() as handle:
            rows = list(csv.DictReader(handle))
        self.assertEqual(sum(int(r["Students"]) for r in rows), 3)
        pending = {row["Status"]: row for row in rows if not row["NSCGraduationDate"]}
        double = pending["Needs decision: diploma date outside reporting cycle"]
        self.assertEqual(double["Eligibility"], "Needs decision: completion code 250")
        self.assertEqual(double["SupplementGraduationDate"], "2025-08-15")
        self.assertEqual(pending["Needs decision: conflicting supplemental diploma dates"]["Eligibility"], "Eligible")
        self.assertTrue(all(not row["DateSource"] for row in pending.values()))
        self.assertFalse(list(self.root.glob("*.decisions.*")))

    def test_completion_mapping_cannot_be_saved_by_a_single_choice_and_settings_are_reported(self):
        from calpads_to_nsc.terminal_input import Finished

        self.record[-1] = "250"
        self.write(self.record)
        with self.assertRaises(Finished):
            self.review_decisions(["1", "q"])
        self.assertFalse(self.review.with_suffix(".decisions.json").exists())
        self.review_decisions(["1", "APPLY"])
        result = self.run_cli("--dry-run")
        self.assertIn(b"Completion code 250", result.stdout)
        self.assertIn(b"Endorsed/Advanced Diploma", result.stdout)
        policy = self.root / "policy.json"
        policy.write_text(json.dumps({"completion_types": {"250": "Regular Diploma"}}))
        override = self.run_cli("--settings", str(policy), "--dry-run")
        self.assertIn(b"Completion code 250", override.stdout)
        self.assertIn(b"Regular Diploma", override.stdout)
        self.assertNotIn(b"Endorsed", override.stdout)

    def test_id_policy_only_validates_selected_local_ids_and_keeps_ssid_joins(self):
        self.record[3] = "invalid local !"
        self.write(self.record)
        dates = self.root / "dates.csv"
        dates.write_text(f"SSID,GraduationDate,FERPABlock\n{self.record[1]},20260601,Y\n")
        for option, expected in (("local", None), ("ssid", self.record[1]), ("blank", "")):
            path = self.root / f"{option}.txt"
            result = self.run_cli(
                "--student-id", option, "--supplement", str(dates), "--output", str(path), confirm_blocks=False
            )
            if expected is None:
                self.assertIn(b"Malformed local student ID", result.stderr)
                self.assertFalse(path.exists())
            else:
                self.assertEqual(result.returncode, 0, result.stderr)
                detail = path.read_text().splitlines()[1].split("\t")
                self.assertEqual([detail[i] for i in (9, 11, 12)], [expected, "20260601", "Y"])

    def test_block_confirmation_only_for_full_output_without_block_column(self):
        self.write(self.record)
        dates = self.root / "dates.csv"
        dates.write_text("SSID,GraduationDate\n")
        for extra in ([], ["--supplement", str(dates)]):
            result = self.run_cli(*extra, "--output", str(self.out), confirm_blocks=False)
            self.assertIn(b"--confirm-no-ferpa-blocks", result.stderr)
            self.assertFalse(self.out.exists())
        result = self.run_cli("--test-output", str(self.root / "practice.txt"), confirm_blocks=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        result = self.run_cli("--supplement", str(dates), "--output", str(self.out))
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_skipped_supplement_rows_have_one_notice_and_validate_the_flag_first(self):
        self.write(self.record)
        dates = self.root / "dates.csv"
        dates.write_text("SSID,GraduationDate,FERPABlock\n,,N\nunknown-id,,\n0123456789,20260601,Y\n")
        result = self.run_cli("--supplement", str(dates), "--output", str(self.out), confirm_blocks=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.count(b"Skipped district supplement"), 1)
        self.assertNotIn(b"unknown-id", result.stdout)
        self.assertIn(b"20260601\tY", self.out.read_bytes())
        dates.write_text("SSID,GraduationDate,FERPABlock\nunknown-id,,maybe\n")
        result = self.run_cli("--supplement", str(dates), "--dry-run")
        self.assertIn(b"Row 2", result.stderr)
        self.assertIn(b"FERPABlock", result.stderr)
        self.assertNotIn(b"unknown-id", result.stderr)

    def test_future_transmission_date_stops_both_outputs_and_allows_modeling(self):
        self.write(self.record)
        for mode in ("--output", "--test-output"):
            result = self.run_cli("--date", "99991231", mode, str(self.out))
            self.assertIn(b"transmission date is in the future", result.stderr)
            self.assertFalse(self.out.exists())
        result = self.run_cli("--date", "99991231", "--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_debug_diagnostic_includes_a_program_location_and_omits_sensitive_values(self):
        self.record[1] = "private-invalid"
        self.write(self.record)
        result = self.run_cli("--debug", "--dry-run")
        diagnostic = next(line for line in result.stderr.splitlines() if line.startswith(b"Diagnostic"))
        self.assertIn(b"ValidationError", diagnostic)
        self.assertIn(b"calpads_to_nsc.", diagnostic)
        self.assertNotIn(str(self.root).encode(), diagnostic)
        self.assertNotIn(b"private-invalid", result.stderr)
