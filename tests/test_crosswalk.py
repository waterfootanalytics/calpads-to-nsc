import csv
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from calpads_to_nsc.crosswalk import EXCLUDE, create, reviewed_config, suggestion
from calpads_to_nsc.generate import ValidationError


class CrosswalkTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.path = Path(temp.name) / "review.xlsx"
        self.ref = dict(
            account_number="001", organization_name="Example", schools={"012345": "PARK HS", "012346": "LAKE HS"}
        )
        self.labels = {"Park High School", "Former Lake Academy"}
        create(self.path, self.labels, self.ref, "fingerprint", rehearsal=True)

    def validate(self, **changes):
        inputs = dict(path=self.path, labels=self.labels, reference=self.ref, digest="fingerprint", writing=False)
        return reviewed_config(**(inputs | changes))

    def test_acronym_and_generic_grade_span(self):
        reference = {"schools": {"012345": "OAS 6-12 COMPLEX", "012346": "OAK AVENUE HS"}}
        self.assertEqual(suggestion("Oak Arts and Sciences 6-12 Complex", reference), "OAS 6-12 COMPLEX (012345)")
        self.assertEqual(suggestion("Unrelated 6-12 Complex", reference), "No clear match")

    def test_different_campus_numbers_require_review(self):
        reference = {"schools": {"012345": "CENTRAL ACADEMY 2"}}
        self.assertEqual(suggestion("Central High School 1", reference), "No clear match")
        self.assertEqual(suggestion("Central High School", reference), "No clear match")

    def test_workbook_writes_reference_text_as_literals(self):
        label = "=Example()"
        reference = dict(self.ref, organization_name=label, schools={"012345": label})
        path = self.path.with_name("literal.xlsx")
        create(path, {label}, reference, "fingerprint")
        book = load_workbook(path)
        try:
            for sheet, cells in (
                ("Match schools", ["B3", "B13", "C13", "D13"]),
                ("Look up a school", ["B10"]),
                ("Source", ["B3"]),
                ("NSC choices", ["A3"]),
            ):
                for cell in cells:
                    self.assertEqual(book[sheet][cell].data_type, "s")
        finally:
            book.close()

    def test_shared_account_prefix_does_not_identify_a_campus(self):
        reference = {
            "organization_name": "Harbor Learning Cooperative",
            "schools": {"012345": "HARBOR OAS 6-12 COMPLEX", "012346": "HARBOR OAK AVENUE HS"},
        }
        self.assertEqual(
            suggestion("Harbor Oak Arts and Sciences 6-12 Complex", reference), "HARBOR OAS 6-12 COMPLEX (012345)"
        )
        self.assertEqual(suggestion("Harbor Unrelated 6-12 Complex", reference), "No clear match")

    def test_decisions_and_stale_inputs(self):
        book = load_workbook(self.path)
        sheet = book["Match schools"]
        self.assertEqual([s.title for s in book if s.sheet_state == "visible"], ["Match schools", "Look up a school"])
        self.assertFalse(sheet.column_dimensions["C"].hidden)
        self.assertEqual(len(sheet.data_validations.dataValidation), 1)
        self.assertEqual(str(sheet.data_validations.dataValidation[-1].sqref), "D13:D14")
        self.assertEqual(sheet.data_validations.dataValidation[0].formula1, "SchoolChoices")
        self.assertEqual(list(book.defined_names["SchoolChoices"].destinations), [("NSC choices", "$A$2:$A$5")])
        self.assertEqual(book["NSC choices"]["A2"].value, "Flag for review")
        self.assertEqual(book["NSC choices"]["A5"].value, EXCLUDE)
        self.assertEqual(sheet["C14"].value, "PARK HS (012345)")
        self.assertEqual(sheet["D14"].value, sheet["C14"].value)
        self.assertEqual(sheet["C13"].value, "LAKE HS (012346)")
        self.assertIsNone(sheet["D13"].value)
        with self.assertRaisesRegex(ValidationError, "D13  Former Lake Academy"):
            self.validate()
        sheet["D13"] = "Flag for review"
        sheet["D14"] = sheet["C14"].value
        sheet["E13"] = "Confirmed"
        book.save(self.path)
        with self.assertRaisesRegex(ValidationError, "D13  Former Lake Academy — Flagged"):
            self.validate()
        sheet["D13"] = "LAKE HS (012346)"
        book.save(self.path)
        self.assertEqual(self.validate()["schools"]["Former Lake Academy"]["act_code"], "012346")
        for changes in (dict(writing=True), dict(labels={"Park High School"}), dict(digest="new")):
            with self.assertRaises(ValidationError):
                self.validate(**changes)
        sheet["D13"] = sheet["D14"].value
        book.save(self.path)
        with self.assertRaises(ValidationError):
            self.validate()
        sheet["D13"] = "MADE UP (012346)"
        book.save(self.path)
        with self.assertRaises(ValidationError):
            self.validate()
        book.close()

    def test_rebuild_preserves_saved_blanks_flags_and_manual_choices(self):
        labels = self.labels | {"Coastal Charter", "Unrelated Campus"}
        old = self.path.with_name("earlier.xlsx")
        create(old, labels, self.ref, "fingerprint", rehearsal=True)
        book = load_workbook(old)
        sheet = book["Match schools"]
        rows = {sheet.cell(n, 2).value: n for n in range(13, 17)}
        for label, pick in (
            ("Coastal Charter", "Flag for review"),
            ("Former Lake Academy", "PARK HS (012345)"),
            ("Park High School", None),
        ):
            sheet.cell(rows[label], 4).value = pick
        book.save(old)
        book.close()
        new = self.path.with_name("filled.xlsx")
        create(new, labels, self.ref, "fingerprint", previous=old)
        book = load_workbook(new)
        sheet = book["Match schools"]
        self.assertIsNone(sheet.cell(rows["Park High School"], 4).value)
        self.assertEqual(sheet.cell(rows["Former Lake Academy"], 4).value, "PARK HS (012345)")
        self.assertEqual(sheet.cell(rows["Coastal Charter"], 4).value, "Flag for review")
        self.assertIsNone(sheet.cell(rows["Unrelated Campus"], 4).value)
        self.assertEqual(dict(book["Source"].values)["Use"], "Rehearsal")
        book.close()

    def test_defer_only_named_unresolved_rehearsal_school(self):
        book = load_workbook(self.path)
        sheet = book["Match schools"]
        pending = {"Former Lake Academy"}
        sheet["D13"] = "LAKE HS (012346)"
        book.save(self.path)
        with self.assertRaisesRegex(ValidationError, "already has a choice"):
            self.validate(deferred=pending)
        for value in (None, "Flag for review"):
            sheet["D13"] = value
            book.save(self.path)
            self.assertEqual(set(self.validate(deferred=pending)["schools"]), {"Park High School"})
        with self.assertRaisesRegex(ValidationError, "D13  Former Lake Academy"):
            self.validate()
        for excluded in ({"Typo"}, self.labels):
            with self.assertRaises(ValidationError):
                self.validate(deferred=excluded)
        sheet["D14"] = None
        book.save(self.path)
        with self.assertRaisesRegex(ValidationError, "D14  Park High School"):
            self.validate(deferred=pending)
        book["Source"]["B2"] = "Submission"
        book.save(self.path)
        with self.assertRaisesRegex(ValidationError, "rehearsal workbook"):
            self.validate(deferred=pending)
        book.close()

    def test_unsupported_workbook_requires_fresh_review_before_reading_or_reusing_choices(self):
        from calpads_to_nsc.crosswalk import SELECTION_POLICY, SchoolEdit
        from calpads_to_nsc.terminal_schools import save_choice

        for problem in (
            "missing selection policy",
            "unknown selection policy",
            "missing exclusion column",
            "missing school list",
        ):
            with self.subTest(problem=problem):
                path = self.path.with_name(problem + ".xlsx")
                book = load_workbook(self.path)
                source = book["Source"]
                policy_row = next(cell.row for cell in source["A"] if cell.value == SELECTION_POLICY)
                if problem == "missing selection policy":
                    source.delete_rows(policy_row)
                elif problem == "unknown selection policy":
                    source.cell(policy_row, 2).value = 2
                elif problem == "missing exclusion column":
                    book["Match schools"]["G12"] = None
                else:
                    del book["NSC choices"]
                book["Match schools"]["D13"] = book["Match schools"]["C13"].value
                book.save(path)
                book.close()
                original = path.read_bytes()
                rebuilt = self.path.with_name("rebuilt.xlsx")
                edit = SchoolEdit("Match schools", "D13", "Former Lake Academy")
                for operation in (
                    lambda: self.validate(path=path, writing=True),
                    lambda: create(rebuilt, self.labels, self.ref, "fingerprint", previous=path),
                    lambda: save_choice(path, self.ref, edit, path.stat(), "LAKE HS (012346)"),
                ):
                    with self.assertRaisesRegex(ValidationError, "Create a new school review workbook"):
                        operation()
                self.assertEqual(path.read_bytes(), original)
                self.assertFalse(rebuilt.exists())

    def test_registered_name_always_comes_from_reference_sheet(self):
        registered = "EXAMPLE REGIONAL ACADEMY FOR ARTS, SCIENCES & TECHNOLOGY HIGH SCHOOL"
        self.ref["schools"]["012347"] = registered
        self.path = self.path.with_name("new.xlsx")
        self.labels = {"Example School"}
        create(self.path, self.labels, self.ref, "fingerprint")
        book = load_workbook(self.path)
        book["Match schools"]["D13"] = f"{registered} (012347)"
        book.save(self.path)
        book.close()
        before = self.path.read_bytes()
        expected = {"act_code": "012347", "nsc_school_name": registered}
        for mode in ({}, {"test_file": True}, {"writing": True}):
            with self.subTest(mode=mode):
                self.assertEqual(self.validate(**mode)["schools"]["Example School"], expected)
        self.assertEqual(self.path.read_bytes(), before)
        fresh = self.path.with_name("refreshed.xlsx")
        create(fresh, self.labels, self.ref, "fingerprint", previous=self.path)
        self.assertEqual(self.validate(path=fresh, writing=True)["schools"]["Example School"], expected)

    def test_rehearsal_uses_registered_name(self):
        registered = "NORTHEAST REGIONAL ACADEMY OF ARTS AND SCIENCE High School"
        ref = dict(self.ref, schools={"012345": registered})
        path = self.path.with_name("practice.xlsx")
        create(path, {registered}, ref, "fingerprint", rehearsal=True)
        inputs = dict(path=path, reference=ref, labels={registered})
        for mode in ({}, {"test_file": True}):
            self.assertEqual(self.validate(**inputs, **mode)["schools"][registered]["nsc_school_name"], registered)
        with self.assertRaisesRegex(ValidationError, "production output"):
            self.validate(**inputs, writing=True)

    def test_confirmed_exclusions_are_bound_to_reporting_year_and_agreement(self):
        path = self.path.with_name("submission.xlsx")
        create(path, self.labels, self.ref, "fingerprint", year=2026)
        book = load_workbook(path)
        self.addCleanup(book.close)
        tab = book["Match schools"]
        tab["D13"] = EXCLUDE
        for reason, confirmation in (
            ("", "EXCLUDE"),
            ("Outside this agreement", "Y"),
            (" " * 5, "EXCLUDE"),
            ("x" * 241, "EXCLUDE"),
        ):
            tab["F13"], tab["G13"] = reason, confirmation
            book.save(path)
            with self.assertRaisesRegex(ValidationError, "F13.*G13"):
                self.validate(path=path, writing=True, year=2026)
        tab["F13"], tab["G13"] = "Campus has a separate agreement", "EXCLUDE"
        book.save(path)
        decisions = self.validate(path=path, writing=True, year=2026)
        self.assertEqual(decisions["excluded_schools"], {"Former Lake Academy": "Campus has a separate agreement"})
        self.assertEqual(set(decisions["schools"]), {"Park High School"})
        for kwargs in (dict(year=2027), dict(labels=self.labels | {"New campus"}), dict(digest="changed")):
            with self.assertRaises(ValidationError):
                self.validate(path=path, writing=True, **kwargs)
        rebuilt = path.with_name("copy.xlsx")
        create(rebuilt, self.labels, self.ref, "fingerprint", previous=path, year=2026)
        self.assertEqual(self.validate(path=rebuilt, writing=True, year=2026), decisions)
        with self.assertRaisesRegex(ValidationError, "fresh school review"):
            create(path.with_name("other.xlsx"), self.labels, self.ref, "fingerprint", previous=path, year=2027)
        tab["D14"], tab["F14"], tab["G14"] = EXCLUDE, "Another agreement", "EXCLUDE"
        book.save(path)
        with self.assertRaisesRegex(ValidationError, "at least one school"):
            self.validate(path=path, writing=True, year=2026)
        tab["D13"] = "LAKE HS (012346)"
        book.save(path)
        with self.assertRaisesRegex(ValidationError, "clear F13 and G13"):
            self.validate(path=path, writing=True, year=2026)
        tab["F13"], tab["G13"] = None, None
        book.save(path)
        self.assertEqual(set(self.validate(path=path, year=2026)["schools"]), {"Former Lake Academy"})

    def test_all_submission_commands_report_and_apply_school_exclusions(self):
        root = self.path.parent
        sheet = root / "agreement.xlsx"
        book = Workbook()
        for row in (
            ("Organization Name:", "Practice District"),
            ("Service Agreement/Account Number:", "001"),
            ("Total High School Count:", 2),
            ("ACT Code", "High School Name", "HS Operating Status"),
            ("012345", "PARK HS", "OPEN"),
            ("012346", "LAKE HS", "OPEN"),
        ):
            book.active.append(row)
        book.save(sheet)
        book.close()
        report = root / "roster.csv"
        report.write_text(
            "SchoolName,SSID,FirstName,LastName,BirthDate,"
            "StudentEnrollmentExitDate,StudentCompletionStatus\n"
            "Park High School,0000000001,Fictional,Example,2008-01-01,2026-06-01,100\n"
            "Lake High School,0000000002,Unused,Example,2008-01-01,2026-06-02,250\n",
            encoding="utf-8",
        )
        supplement = root / "supplement.csv"
        supplement.write_text("SSID,GraduationDate,FERPABlock\n0000000001,20260601,Y\n", encoding="utf-8")
        review = root / "school-review.xlsx"
        args = [
            sys.executable,
            "-m",
            "calpads_to_nsc",
            str(report),
            "--no-prompt",
            "--reference-sheet",
            str(sheet),
            "--year",
            "2026",
            "--date",
            "20260630",
        ]

        def run(*arguments):
            return subprocess.run([*args, *map(str, arguments)], capture_output=True, text=True)

        result = run("--create-crosswalk", review)
        self.assertEqual(result.returncode, 0, result.stderr)
        args += ["--crosswalk", str(review)]
        book = load_workbook(review)
        tab = book["Match schools"]
        self.assertEqual(tab["B13"].value, "Lake High School")
        tab["D13"], tab["F13"], tab["G13"] = EXCLUDE, "Separate agreement", "EXCLUDE"
        book.save(review)
        book.close()
        for mode in (
            ("--check-crosswalk",),
            ("--dry-run",),
            ("--test-output", root / "sample.txt"),
            ("--output", root / "full.txt", "--confirm-no-ferpa-blocks"),
            ("--supplement", supplement, "--compare-dates", root / "dates.csv"),
        ):
            with self.subTest(mode=mode[0]):
                result = run(*mode)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIn("Excluded from this submission:\n  Lake High School", result.stdout)
                self.assertNotIn("Separate agreement", result.stdout + result.stderr)
        for name in ("sample.txt", "full.txt"):
            lines = (root / name).read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(lines), 3)
            self.assertEqual(lines[1].split("\t")[14], "012345")
        with (root / "dates.csv").open(encoding="utf-8", newline="") as handle:
            groups = list(csv.DictReader(handle))
        self.assertEqual([row["Students"] for row in groups], ["1"])
