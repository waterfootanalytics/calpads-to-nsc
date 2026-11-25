import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from calpads_to_nsc import terminal_guide
from calpads_to_nsc import native_picker
from calpads_to_nsc.native_picker import DialogUnavailable


def write_saved_graduates(path):
    fields = [
        "PD3",
        "NO SSN",
        "River",
        "",
        "Example",
        "",
        "",
        "",
        "20000101",
        "",
        "Regular Diploma",
        "20240601",
        "N",
        "Example High",
        "012345",
        *([""] * 12),
        "ED",
    ]
    path.write_text("PH3\t123\tExample District\tP\t20240630\t2024\n" + "\t".join(fields) + "\nPT3\t3\n")


class TerminalGuideTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.launch = self.root / "Launch folder"
        self.launch.mkdir()
        cwd = patch.object(Path, "cwd", return_value=self.launch)
        cwd.start()
        self.addCleanup(cwd.stop)
        self.outputs = self.root / "outputs"
        self.reports = self.root / "Certified reports"
        self.reports.mkdir()
        self.report = self.reports / "report.csv"
        self.report.write_text(
            "SchoolName,SSID,FirstName,LastName,BirthDate,EnrollmentExitDate,StudentCompletionStatus\n"
            "Sample High School,0000000001,Example,Sample,2008-01-01,2026-06-12,100\n",
            encoding="utf-8",
        )
        self.reference = self.root / "Account reference.xlsx"
        self.reference.write_bytes(b"synthetic account sheet")
        self.review = self.root / "School review.xlsx"
        self.prefs = self.root / "Settings" / "locations.json"
        self.paths = dict(reports=str(self.reports), reference=str(self.reference), review=str(self.review))
        settings = patch.object(terminal_guide, "PREFERENCES", self.prefs)
        settings.start()
        self.addCleanup(settings.stop)
        self.real_converter = terminal_guide.run_converter
        converter = patch.object(terminal_guide, "run_converter", return_value=0)
        self.converter = converter.start()
        self.addCleanup(converter.stop)

        def select_test_path(*args, **kwargs):
            selection = input().strip()
            return None if selection.casefold() == "q" else Path(selection)

        picker = patch.object(terminal_guide, "pick_path", side_effect=select_test_path)
        self.picker = picker.start()
        self.addCleanup(picker.stop)

    def seed(self, **extra):
        self.review.write_bytes(b"synthetic review workbook")
        self.prefs.parent.mkdir(exist_ok=True)
        self.prefs.write_text(json.dumps({**self.paths, **extra}), encoding="utf-8")

    def session(self, *responses):
        entries = "".join(str(response) + "\n" for response in responses)
        output = io.StringIO()
        with patch.object(sys, "stdin", io.StringIO(entries)), redirect_stdout(output), redirect_stderr(output):
            status = terminal_guide.main()
        self.assertEqual(status, 0, output.getvalue())
        return output.getvalue()

    def child_args(self):
        self.converter.assert_called_once()
        return list(map(str, self.converter.call_args.args[0]))

    def write_reference(self, path):
        book = Workbook()
        for row in (
            ("Organization Name:", "Sample District"),
            ("Service Agreement/Account Number:", "123"),
            ("Total High School Count:", 1),
            ("ACT Code", "HS Operating Status", "High School Name"),
            ("012345", "OPEN", "SAMPLE HIGH SCHOOL"),
        ):
            book.active.append(row)
        book.save(path)
        book.close()

    def test_first_run_explains_required_and_optional_inputs_before_any_action(self):
        output = self.session("q")
        self.assertIn("CALPADS", output)
        self.assertIn("1.23", output)
        self.assertIn("account reference sheet", output.casefold())
        self.assertIn("optional", output.casefold())
        self.assertIn("District supplement CSV", output)
        self.assertIn("district records", output)
        self.assertNotIn("Make the full graduates file", output)
        self.converter.assert_not_called()

    def test_preflight_prevents_generation_with_empty_or_invalid_reports(self):
        self.seed()
        self.report.unlink()
        (self.reports / ".DS_Store").write_bytes(b"synthetic")
        (self.reports / "~$report.xlsx").write_bytes(b"synthetic")
        nested = self.reports / "nested"
        for kind in ("hidden files only", "empty export", "mixed exports and subfolders"):
            with self.subTest(kind=kind):
                if kind == "empty export":
                    self.report.touch()
                if kind == "mixed exports and subfolders":
                    self.report.write_bytes(b"synthetic")
                    nested.mkdir()
                    (nested / "report.csv").write_bytes(b"synthetic")
                output = self.session("q")
                self.assertIn("CALPADS", output)
                self.assertIn("1.23", output)
                self.assertNotIn("Make the full graduates file", output)
                self.converter.assert_not_called()

    def test_separate_folder_preflight_reads_only_reports_to_infer_the_year(self):
        dates = self.root / "Dates.csv"
        dates.write_bytes(b"synthetic")
        self.seed(dates=str(dates))
        protected = {self.reference, self.review, dates}
        reports_read = []
        original_open = io.open

        def check_open(path, *args, **kwargs):
            if isinstance(path, (str, os.PathLike)) and Path(path) in protected:
                self.fail("Cycle inference opened an unrelated district file.")
            if isinstance(path, (str, os.PathLike)) and Path(path) == self.report:
                reports_read.append(path)
            return original_open(path, *args, **kwargs)

        with patch("io.open", side_effect=check_open), patch("builtins.open", side_effect=check_open):
            output = self.session("q")
        self.assertTrue(reports_read)
        self.assertIn("CALPADS completion reporting year · 2025–26", output)
        self.assertIn("Review school matches", output)
        self.converter.assert_not_called()

    def test_reference_workbook_in_report_folder_is_not_counted_as_a_calpads_export(self):
        self.report.unlink()
        wrong_location = self.reports / self.reference.name
        self.reference.rename(wrong_location)
        self.seed(reference=str(wrong_location))
        output = self.session("q")
        self.assertIn("Move Account reference.xlsx out of the CALPADS folder", output)
        self.assertNotIn("Make the full graduates file", output)
        self.converter.assert_not_called()

    def test_first_run_remembers_the_reports_when_user_needs_to_obtain_the_reference_sheet(self):
        self.session(self.reports, "q")
        self.assertEqual(json.loads(self.prefs.read_text())["reports"], str(self.reports))
        transcript = self.session("q")
        self.assertNotIn("Folder of certified CALPADS 1.23 exports", transcript)
        self.assertIn("Clearinghouse Account Reference Sheet", transcript)

    def test_missing_optional_dates_still_allow_generation_with_calpads(self):
        self.seed(dates=str(self.root / "Missing dates.csv"))
        output = self.session("3", "", "q")
        self.assertIn("Add a district supplement?", output)
        self.assertIn("N for no block or unknown status", output)
        args = self.child_args()
        self.assertNotIn("--supplement", args)
        self.assertNotIn("--calpads-exit-dates", args)
        self.assertNotIn("--include-demographics", args)
        self.assertIn("No demographic columns were found", output)
        self.assertEqual(args[args.index("--output") + 1], str(self.outputs))

    def add_report_columns(self, headings, fields):
        header, student = self.report.read_text(encoding="utf-8").splitlines()
        self.report.write_text(f"{header},{headings}\n{student},{fields}\n", encoding="utf-8")

    def test_disability_defaults_to_included_after_previous_and_current_file_exclusions(self):
        self.add_report_columns("Gender,Special Education Status", "UnknownTestGender,Yes")
        self.seed(demographics="include")
        screen = self.session("2", "", "", "", "3", "", "", "2", "2", "", "", "", "3", "", "", "1", "q")
        self.assertNotIn("UnknownTestGender", screen)
        self.assertIn("Available in the report headings · Gender, Disability", screen)
        self.assertRegex(screen, r"1 +[^\n]*Special Education Status")
        self.assertRegex(screen, r"2 +[^\n]*without disability")
        calls = [call.args[0] for call in self.converter.call_args_list]
        self.assertEqual(
            ["--use-special-education-for-disability" in args for args in calls], [True, False, True, True]
        )
        self.assertTrue(all("--include-demographics" in args for args in calls))
        self.assertNotIn("demographics", terminal_guide.load_locations())

    def test_excluding_demographics_skips_disability_even_with_saved_inclusion(self):
        self.add_report_columns("Gender,Special Education Status", "F,Yes")
        self.seed(demographics="include_disability")
        screen = self.session("3", "", "2", "q")
        self.assertNotIn("Use CALPADS Special Education Status for NSC disability?", screen)
        args = self.child_args()
        self.assertNotIn("--include-demographics", args)
        self.assertNotIn("--use-special-education-for-disability", args)

    def test_default_includes_demographics_after_saved_or_current_file_exclusion(self):
        self.add_report_columns("Gender", "F")
        self.seed(demographics="omit")
        screen = self.session("2", "", "", "3", "", "2", "2", "", "", "q")
        self.assertRegex(screen, r"1\s+Include available demographics")
        self.assertRegex(screen, r"2\s+Exclude demographics")
        self.assertEqual(
            ["--include-demographics" in call.args[0] for call in self.converter.call_args_list],
            [True, False, True],
        )

    def test_optional_disability_question_needs_its_header_and_accepts_global_exit(self):
        self.add_report_columns("Gender", "F")
        self.seed()
        screen = self.session("2", "", "1", "q")
        self.assertNotIn("Use CALPADS Special Education Status for NSC disability?", screen)
        self.assertIn("--include-demographics", self.child_args())
        self.converter.reset_mock()
        self.add_report_columns("Special Education Status", "Yes")
        saved = self.prefs.read_bytes()
        self.assertIn("Guide closed.", self.session("3", "", "1", "Q"))
        self.assertEqual(self.prefs.read_bytes(), saved)
        self.converter.assert_not_called()

    def test_replace_an_unavailable_optional_date_file_when_generating(self):
        stale = self.root / "Stale dates.csv"
        self.seed(dates=str(stale))
        replacement = self.root / "Current dates.csv"
        replacement.write_text("SSID,GraduationDate\n0000000001,20260612\n", encoding="utf-8")
        self.session("2", "y", replacement, "q")
        args = self.child_args()
        self.assertNotIn(str(stale), args)
        self.assertEqual(args[args.index("--supplement") + 1], str(replacement))
        self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(replacement))

    def test_repair_only_the_moved_account_sheet_before_showing_the_menu(self):
        self.seed()
        replacement = self.root / "Current account.xls"
        self.reference.rename(replacement)
        wrong_format = self.root / "Wrong file.csv"
        wrong_format.write_bytes(b"synthetic")
        self.session(wrong_format, replacement, "", "q")
        args = self.child_args()
        self.assertEqual(args[args.index("--reference-sheet") + 1], str(replacement))
        self.assertEqual(args[args.index("--crosswalk") + 1], str(self.review))
        self.assertNotIn("--create-crosswalk", args)
        self.assertEqual(json.loads(self.prefs.read_text())["reference"], str(replacement))

    def test_missing_review_workbook_can_be_replaced_or_located(self):
        self.seed()
        self.review.unlink()
        existing = self.root / "Located review.xlsx"
        existing.write_bytes(b"existing choices")
        for choice, target, expected in (("2", existing, "--crosswalk"), ("", self.review, "--create-crosswalk")):
            with self.subTest(action=expected):
                self.prefs.write_text(json.dumps(self.paths))
                self.converter.reset_mock()
                self.converter.return_value = 3
                self.session(choice, target, "q")
                args = self.child_args()
                self.assertEqual(args[args.index(expected) + 1], str(target))
                self.assertEqual(args[args.index("--reference-sheet") + 1], str(self.reference))
                self.assertEqual(existing.read_bytes(), b"existing choices")

    def test_return_resumes_original_workbook_without_opening_a_desktop_app(self):
        self.seed()
        self.review.write_bytes(b"saved synthetic choices")
        original = self.prefs.read_bytes()
        with patch.object(subprocess, "run", side_effect=AssertionError("Unexpected desktop or shell command")):
            output = self.session("", "q")
        args = self.child_args()
        self.assertEqual(args[0], str(self.reports))
        self.assertEqual(args[args.index("--reference-sheet") + 1], str(self.reference))
        self.assertEqual(args[args.index("--crosswalk") + 1], str(self.review))
        self.assertIn("--check-crosswalk", args)
        self.assertIn("--terminal-review", args)
        self.assertNotIn("--create-crosswalk", args)
        self.assertNotIn("--open-review", args)
        self.assertNotIn("Finder", output)
        self.assertEqual(self.review.read_bytes(), b"saved synthetic choices")
        self.assertEqual(self.prefs.read_bytes(), original)

    def test_setup_accepts_paths_with_spaces_and_saves_a_new_review_for_later(self):
        def create_review(args):
            target = Path(args[args.index("--create-crosswalk") + 1])
            self.assertEqual(target, self.review)
            target.touch()
            return 3

        self.converter.side_effect = create_review
        self.session(self.reports, self.reference, "", self.review)
        args = self.child_args()
        self.assertIn("--terminal-review", args)
        self.assertNotIn("--open-review", args)
        self.assertEqual(json.loads(self.prefs.read_text(encoding="utf-8")), self.paths)
        if os.name == "posix":
            self.assertEqual(self.prefs.stat().st_mode & 0o077, 0)

    def test_quitting_or_end_of_input_during_setup_preserves_existing_locations(self):
        self.seed()
        original = self.prefs.read_bytes()
        for final_input in ((), ("q",)):
            with self.subTest(final_input=final_input):
                self.session("4", self.reports, *final_input)
                self.assertEqual(self.prefs.read_bytes(), original)
                self.converter.assert_not_called()

    def test_test_file_uses_saved_supplement_for_dates_and_blocks_without_a_second_prompt(self):
        dates = self.root / "PowerSchool dates.csv"
        dates.write_text("SSID,GraduationDate,FERPABlock\n0000000001,20260612,Y\n", encoding="utf-8")
        self.seed(dates=str(dates))
        screen = self.session("2", "", "q")
        self.assertNotIn("Filename >", screen)
        date_step = screen.split("How should the saved district supplement be used?", 1)[1].split("Where should", 1)[0]
        self.assertIn(dates.name, date_step)
        self.assertIn("CALPADS exit date", date_step)
        self.assertIn("SSID,GraduationDate,FERPABlock", date_step)
        self.assertIn("dates-only CSV", date_step)
        self.assertIn("Skip the supplement", date_step)
        self.assertIn("Return", date_step)
        self.assertNotIn("Use the FERPA column in the CALPADS reports", screen)
        args = self.child_args()
        for option, value in (("--supplement", dates), ("--test-output", self.outputs)):
            self.assertEqual(args[args.index(option) + 1], str(value))
        self.assertIn("--reference-sheet", args)
        self.assertIn("--crosswalk", args)
        self.assertNotIn("--output", args)
        self.assertEqual(set(json.loads(self.prefs.read_text())), {*self.paths, "dates"})

    def test_choosing_blocks_only_keeps_the_saved_csv_for_ferpa_blocks(self):
        dates = self.root / "Saved dates.csv"
        dates.write_text("SSID,GraduationDate\n0000000001,20260610\n", encoding="utf-8")
        self.seed(dates=str(dates))
        for response, use_saved in (("", True), ("B", False)):
            with self.subTest(response=response):
                self.converter.reset_mock()
                self.session("3", response, "q")
                args = self.child_args()
                self.assertEqual(args[args.index("--supplement") + 1], str(dates))
                self.assertEqual("--calpads-exit-dates" in args, not use_saved)
                self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(dates))

    def test_skip_omits_the_saved_supplement_for_this_output_and_preserves_preferences(self):
        supplement = self.root / "Supplement.csv"
        supplement.write_text("SSID,GraduationDate,FERPABlock\n0000000001,,Y\n", encoding="utf-8")
        self.seed(dates=str(supplement))
        self.session("2", "s", "q")
        args = self.child_args()
        self.assertNotIn("--supplement", args)
        self.assertNotIn("--calpads-exit-dates", args)
        self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(supplement))

    def test_change_replaces_available_saved_csv_after_generation(self):
        current = self.root / "First dates.csv"
        replacement = self.root / "Updated dates.csv"
        for path in (current, replacement):
            path.write_text("SSID,GraduationDate\n0000000001,20260610\n", encoding="utf-8")
        self.seed(dates=str(current))
        self.session("2", "C", replacement, "q")
        args = self.child_args()
        self.assertEqual(args[args.index("--supplement") + 1], str(replacement))
        self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(replacement))

    def test_invalid_date_choice_then_quit_preserves_files_and_preferences(self):
        dates = self.root / "Saved dates.csv"
        dates.write_text("SSID,GraduationDate\n0000000001,20260610\n", encoding="utf-8")
        for paths in ({}, {"dates": str(dates)}):
            with self.subTest(saved=bool(paths)):
                self.seed(**paths)
                original = self.prefs.read_bytes()
                output = self.session("3", "maybe", "Q")
                self.assertIn("Press Return or enter one of the letters shown.", output)
                self.assertNotIn("Where should the graduates file be saved?", output)
                self.converter.assert_not_called()
                self.assertEqual(self.prefs.read_bytes(), original)

    def test_full_file_saves_automatically_and_remembers_new_dates(self):
        dates = self.root / "New dates.csv"
        dates.write_text("SSID,GraduationDate\n0000000001,20260612\n", encoding="utf-8")
        self.seed()
        screen = self.session("3", "y", dates, "q")
        self.assertNotIn("Filename >", screen)
        args = self.child_args()
        self.assertEqual(args[args.index("--output") + 1], str(self.outputs))
        self.assertIn("--review-full", args)
        self.assertEqual(args[args.index("--supplement") + 1], str(dates))
        self.assertIn("--reference-sheet", args)
        self.assertIn("--crosswalk", args)
        self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(dates))

    def test_review_graduation_decisions_loads_the_supplement_without_writing_a_submission(self):
        supplement = self.root / "district.csv"
        supplement.write_text("SSID,GraduationDate,FERPABlock\n0000000001,20250101,Y\n", encoding="utf-8")
        self.seed(dates=str(supplement))
        screen = self.session("6", "", "q")
        self.assertIn("Review graduation decisions", screen)
        args = self.child_args()
        self.assertEqual(args[args.index("--supplement") + 1], str(supplement))
        self.assertEqual(args[args.index("--crosswalk") + 1], str(self.review))
        self.assertIn("--dry-run", args)
        self.assertIn("--review-decisions", args)
        self.assertNotIn("--output", args)
        self.assertNotIn("--test-output", args)

    def test_pausing_an_unfinished_review_remembers_the_selected_supplement(self):
        self.seed()
        selected = self.root / "selected supplement.csv"
        selected.write_text("SSID,GraduationDate,FERPABlock\n0000000001,20250101,Y\n", encoding="utf-8")
        self.converter.return_value = 3
        self.session("6", "y", selected)
        self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(selected))

    def test_failed_generation_preserves_previous_outputs(self):
        self.seed()
        self.outputs.mkdir()
        destination = self.outputs / "Graduates_previous.txt"
        destination.write_bytes(b"previous synthetic output")
        self.converter.return_value = 1
        screen = self.session("3", "", "q")
        self.assertIn("No graduates file was saved.", screen)
        self.assertIn("Retry with the same inputs", screen)
        self.assertNotIn("Open this file in ", screen)
        self.assertEqual(screen.count("What would you like to do?"), 1)
        self.assertEqual(list(self.outputs.iterdir()), [destination])
        self.assertEqual(destination.read_bytes(), b"previous synthetic output")

    def test_debug_reaches_converter_commands(self):
        with patch.object(sys, "argv", ["prepare-nsc-graduates", "--debug", "--guide"]):
            with patch.object(subprocess, "run") as run:
                self.real_converter(["input", "--check-headers"])
        self.assertIn("--debug", run.call_args.args[0])

    def test_change_column_j_for_one_workspace_and_recover_from_local_id_error(self):
        self.seed()
        self.converter.side_effect = [1, 0, 0]
        screen = self.session("2", "", "i", "3", "3", "i", "2", "", "q")
        policies = []
        for call in self.converter.call_args_list:
            flags = call.args[0]
            policies.append(flags[flags.index("--student-id") + 1])
        self.assertEqual(policies, ["local", "blank", "ssid"])
        self.assertIn("Leave column J blank", screen)
        self.assertEqual(terminal_guide.load_locations()["student_id"], "ssid")
        other = self.root / "Other district" / "input"
        other.mkdir(parents=True)
        self.assertNotIn("student_id", terminal_guide.from_folder(other, terminal_guide.load_locations()))

    def test_retry_keeps_test_and_supplement_choices(self):
        dates = self.root / "Supplement.csv"
        dates.write_text("SSID,GraduationDate,FERPABlock\n")
        self.seed()
        self.converter.side_effect = [1, terminal_guide.FINISHED_AFTER_SAVE]
        screen = self.session("2", "y", dates, "r")
        first, second = self.converter.call_args_list
        self.assertEqual(first, second)
        self.assertIn("--test-output", first.args[0])
        self.assertIn(str(dates), first.args[0])
        self.assertEqual(json.loads(self.prefs.read_text())["dates"], str(dates))
        self.assertEqual(screen.count("What would you like to do?"), 1)

    def test_file_review_shows_only_requested_graduate_total(self):
        saved = self.root / "Graduates_example.txt"
        saved.write_bytes(b"PH3\tpractice\r\nPD3\tfictional-value\r\nPT3\t3\r\n")
        for choices, expected in ((["m"], False), (["t", "m"], True)):
            with self.subTest(choices=choices):
                screen = io.StringIO()
                with redirect_stdout(screen), patch("builtins.input", side_effect=choices):
                    self.assertEqual(terminal_guide.review_saved_file(saved, sample=False), 0)
                self.assertEqual("Graduates included in this file · 1" in screen.getvalue(), expected)
                self.assertNotIn("fictional-value", screen.getvalue())

    def test_excel_and_upload_use_the_saved_test_and_wait_for_confirmation(self):
        saved = self.root / "TEST_Graduates_example_2.txt"
        write_saved_graduates(saved)
        (self.root / "TEST_Graduates_later.txt").touch()
        screen = io.StringIO()
        choices = iter(["", "c", "o", ""])

        def respond(prompt):
            choice = next(choices)
            if choice == "c":
                self.assertNotIn("Submit your Graduate test file", screen.getvalue())
            return choice

        with (
            redirect_stdout(screen),
            patch("builtins.input", side_effect=respond),
            patch.object(terminal_guide, "open_workbook") as excel,
            patch.object(terminal_guide.webbrowser, "open") as browser,
        ):
            self.assertEqual(terminal_guide.review_saved_file(saved, sample=True), terminal_guide.FINISHED_AFTER_SAVE)
        excel.assert_called_once_with(saved, excel_only=True)
        browser.assert_called_once_with("https://www.studentclearinghouse.org/")
        for item in ("User Login", "Submit a Test File", "Graduate", str(saved), "Upload", "Proposed Solution"):
            self.assertIn(item, screen.getvalue())

    def test_review_handles_excel_resave_or_finish_later_before_upload(self):
        saved = self.root / "TEST_Graduates_example.txt"
        for choice in ("c", "q"):
            with self.subTest(choice=choice):
                saved.write_text("synthetic original")
                screen = io.StringIO()
                with (
                    redirect_stdout(screen),
                    patch("builtins.input", side_effect=["", choice]),
                    patch.object(
                        terminal_guide, "open_workbook", side_effect=lambda *a, **k: saved.write_text("edited")
                    ),
                    patch.object(terminal_guide.webbrowser, "open") as browser,
                ):
                    status = terminal_guide.review_saved_file(saved, sample=True)
                self.assertNotIn("Submit your Graduate test file", screen.getvalue())
                browser.assert_not_called()
                if choice == "c":
                    self.assertEqual(status, 0)
                    self.assertIn("create a fresh test file", screen.getvalue())
                else:
                    self.assertEqual(status, terminal_guide.FINISHED_AFTER_SAVE)

    def test_saved_file_menu_checks_test_file_and_returns_to_menu(self):
        self.seed()
        saved = self.root / "TEST_Graduates_sample.txt"
        write_saved_graduates(saved)
        screen = self.session("7", saved, "q")
        self.assertIn("Saved file layout checked. No field warnings.", screen)
        self.assertEqual(screen.count("What would you like to do?"), 2)
        self.converter.assert_not_called()

    def test_resume_full_file_uses_the_selected_original_and_secure_ftp(self):
        self.seed()
        self.outputs.mkdir()
        full = self.outputs / "Graduates_20260630_121500.txt"
        write_saved_graduates(full)
        (self.outputs / "Graduates_newer.txt").touch()
        with (
            patch.object(terminal_guide, "open_workbook") as excel,
            patch.object(terminal_guide.webbrowser, "open") as browser,
        ):
            screen = self.session("5", self.outputs / "TEST_Graduates_example.txt", full, "", "c", "o", "")
        self.assertIn("Test files begin with TEST_", screen)
        self.assertIn(str(full.resolve()), screen)
        self.assertIn("FTP confirmation email", screen)
        self.assertNotIn("Submit a Test File", screen)
        self.assertEqual(screen.count("What would you like to do?"), 1)
        self.assertIn("Guide closed.", screen)
        excel.assert_called_once_with(full.resolve(), excel_only=True)
        browser.assert_called_once_with("https://ftps.nslc.org/")
        self.converter.assert_not_called()

    def test_submission_instructions_offer_an_explicit_way_back_to_the_menu(self):
        self.seed()
        full = self.root / "Graduates_example.txt"
        write_saved_graduates(full)
        screen = self.session("5", full, "c", "m", "q")
        self.assertIn("Finish and exit", screen)
        self.assertEqual(screen.count("What would you like to do?"), 2)
        self.converter.assert_not_called()

    def test_exit_words_at_main_menu_supplement_and_filename(self):
        self.seed()
        for word in ("Q", "Quit", "Exit", "CLOSE", " done "):
            with self.subTest(word=word):
                self.assertIn("Guide closed.", self.session(word))
                screen = self.session("2", word)
                self.assertEqual(screen.count("What would you like to do?"), 1)
                self.assertIn("Guide closed.", screen)
                with (
                    patch("builtins.input", return_value=word),
                    redirect_stdout(io.StringIO()),
                    self.assertRaises(terminal_guide.Finished),
                ):
                    terminal_guide.save_as("Save workbook", self.review, (".xlsx",))
        self.converter.assert_not_called()

    def test_full_review_cancellation_and_modification_do_not_open_ftp(self):
        self.seed()
        full = self.root / "Graduates_example.txt"
        write_saved_graduates(full)
        with (
            patch.object(terminal_guide.webbrowser, "open") as browser,
            patch.object(terminal_guide, "open_workbook", side_effect=lambda *a, **k: full.write_text("changed")),
        ):
            screen = self.session("5", "q", "5", full, "", "c", "q")
        self.assertIn("The text file changed during review. Choose 3", screen)
        self.assertNotIn("Submit your full graduates file", screen)
        browser.assert_not_called()
        self.converter.assert_not_called()

    def test_guided_defaults_create_review_and_generate_using_actual_converter(self):
        self.write_reference(self.reference)
        self.converter.side_effect = self.real_converter
        real_run = subprocess.run

        def run_conversion(command, **kwargs):
            self.assertEqual(command[1:3], ["-m", "calpads_to_nsc"])
            self.assertIn("--terminal-review", command)
            self.assertNotIn("--open-review", command)
            result = real_run(
                [*command, "--date", "20260630", "--confirm-no-ferpa-blocks"],
                stdin=subprocess.DEVNULL,
                capture_output=True,
                timeout=10,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            return result

        with patch.object(subprocess, "run", side_effect=run_conversion):
            self.session(self.reports, self.reference, "", self.review, "3", "", "q")
            (destination,) = self.outputs.glob("Graduates_*.txt")
            dates = self.root / "PowerSchool supplement.csv"
            dates.write_text("SSID,GraduationDate,FERPABlock\n0000000001,20260610,Y\n", encoding="utf-8")
            self.session("3", "y", dates, "q")
            (combined,) = set(self.outputs.glob("Graduates_*.txt")) - {destination}
            self.session("2", "b", "q")
            (flags_only,) = self.outputs.glob("TEST_Graduates_*.txt")
            dates.write_text("SSID,GraduationDate\n0000000001,20260610\n", encoding="utf-8")
            self.session("2", "", "q")
            (dates_only,) = set(self.outputs.glob("TEST_Graduates_*.txt")) - {flags_only}
        self.assertEqual(self.converter.call_count, 5)
        self.assertTrue(self.review.is_file())
        header, graduate, trailer = destination.read_text().splitlines()
        self.assertTrue(header.startswith("PH3\t"))
        self.assertEqual(graduate.split("\t")[11:15], ["20260612", "N", "SAMPLE HIGH SCHOOL", "012345"])
        self.assertEqual(trailer, "PT3\t3")
        self.assertEqual(combined.read_text().splitlines()[1].split("\t")[11:13], ["20260610", "Y"])
        self.assertEqual(flags_only.read_text().splitlines()[1].split("\t")[11:13], ["20260612", "Y"])
        self.assertEqual(dates_only.read_text().splitlines()[1].split("\t")[11:13], ["20260610", "N"])

    def test_single_folder_discovery_through_real_converter(self):
        project = self.root / "Private batch"
        folder = project / "input"
        folder.mkdir(parents=True)
        self.add_report_columns("Gender,Ethnicity/Race,ELAS Status,Special Education Status", "F,Asian,EL,Yes")
        self.report.rename(folder / "Campus.csv")
        supplement = folder / "PowerSchool export.csv"
        supplement.write_text("SSID,GraduationDate,FERPABlock,ProgramCode\n0000000001,20260609,Y,STEM\n")
        self.write_reference(folder / "Current sheet.xlsx")
        self.converter.side_effect = lambda args: self.real_converter([*args, "--date", "20260630"])
        self.picker.side_effect = AssertionError("Unexpected file dialog for complete inputs")
        with patch.object(Path, "cwd", return_value=project):
            self.session("", "", "2", "", "", "", "3", "", "", "", "q")
        self.assertTrue((project / "outputs" / "School review.xlsx").exists())
        (test_file,) = (project / "outputs").glob("TEST_Graduates_*.txt")
        (full_file,) = (project / "outputs").glob("Graduates_*.txt")
        for path in (test_file, full_file):
            record = path.read_text().splitlines()[1].split("\t")
            self.assertEqual(record[11:13], ["20260609", "Y"])
            self.assertEqual(record[15:18], ["F", "AS", ""])
            self.assertEqual((record[22], record[25], record[26]), ("Y", "Y", "STEM"))
        with patch.object(Path, "cwd", return_value=folder):
            self.assertEqual(terminal_guide.launch_input_folder(), folder)
            self.session("2", "b", "1", "2", "q")
        (blocks_only,) = set((project / "outputs").glob("TEST_Graduates_*.txt")) - {test_file}
        record = blocks_only.read_text().splitlines()[1].split("\t")
        self.assertEqual(record[11:13], ["20260612", "Y"])
        self.assertEqual((record[22], record[25], record[26]), ("Y", "", ""))
        with patch.object(Path, "cwd", return_value=folder):
            self.session("2", "s", "2", "q")
        (minimal,) = set((project / "outputs").glob("TEST_Graduates_*.txt")) - {test_file, blocks_only}
        record = minimal.read_text().splitlines()[1].split("\t")
        self.assertEqual(record[11:13], ["20260612", "N"])
        self.assertEqual(record[15:27], [""] * 12)

    def test_separate_reports_folder_saves_to_sibling_outputs_and_reopens_same_inputs(self):
        self.seed()
        self.session("2", "q")
        self.assertFalse(self.outputs.exists())
        self.session("2", "", "q")
        args = self.child_args()
        self.assertEqual(args[args.index("--test-output") + 1], str(self.outputs))
        self.assertIsNone(terminal_guide.launch_input_folder())
        self.assertIn(self.review.name, self.session("q"))

    def test_multiple_references_require_selection_and_supplement_selection_can_wait(self):
        folder = self.root / "input"
        folder.mkdir()
        self.report.rename(folder / "School.csv")
        ref = folder / "Sheet.xls"
        ref.write_bytes(b"synthetic reference")
        (folder / "Old.pdf").write_bytes(b"synthetic reference")
        for name in ("New.csv", "Old.csv"):
            (folder / name).write_text("SSID,GraduationDate\n")
        self.picker.side_effect = [ref]
        choices = terminal_guide.from_folder(folder, {})
        self.assertEqual(choices["reference"], str(ref))
        self.assertNotIn("dates", choices)
        self.assertEqual(terminal_guide.from_folder(folder, choices), choices)
        self.picker.assert_called_once()
        self.assertEqual([p.name for p in terminal_guide.folder_inputs(folder)["reports"]], ["School.csv"])

    def test_skip_or_replace_multiple_supplements_when_launched_outside_workspace(self):
        folder = self.root / "School records" / "input"
        folder.mkdir(parents=True)
        self.report.rename(folder / "Graduates.csv")
        self.write_reference(folder / "Schools.xlsx")
        for name in ("Fall supplement.csv", "Updated supplement.csv"):
            (folder / name).write_text("SSID,GraduationDate,FERPABlock\n0000000001,20260609,Y\n")
        other = self.root / "Reviewed supplement.csv"
        other.write_text("SSID,GraduationDate,FERPABlock\n0000000001,20260608,Y\n")
        self.converter.side_effect = lambda args: self.real_converter([*args, "--date", "20260630"])
        self.picker.side_effect = [folder, other]
        self.session("", "", "2", "", "3", "y", "q")
        outputs = folder.parent / "outputs"
        (test_file,) = outputs.glob("TEST_Graduates_*.txt")
        (full_file,) = outputs.glob("Graduates_*.txt")
        self.assertEqual(test_file.read_text().splitlines()[1].split("\t")[11:13], ["20260612", "N"])
        self.assertEqual(full_file.read_text().splitlines()[1].split("\t")[11:13], ["20260608", "Y"])
        self.assertEqual(terminal_guide.load_locations()["dates"], str(other))
        self.assertIn(other.name, self.session("q"))
        self.assertEqual(terminal_guide.load_locations()["dates"], str(other))
        self.assertEqual(self.picker.call_count, 2)
        self.assertFalse((self.launch / "outputs").exists())

    def test_cancelling_supplement_dialog_returns_to_supplement_choice(self):
        self.seed()
        self.picker.return_value = None
        self.picker.side_effect = None
        original = self.prefs.read_bytes()
        output = self.session("2", "y", "q")
        self.assertIn("Selection cancelled", output)
        self.assertEqual(self.prefs.read_bytes(), original)
        self.converter.assert_not_called()
        self.picker.assert_called_once()

    def test_supplement_picker_starts_in_selected_reports_folder(self):
        self.picker.side_effect = None
        selected = self.root / "Selected.csv"
        self.picker.return_value = selected
        self.assertEqual(terminal_guide.browse("dates", {"reports": str(self.reports)}), str(selected))
        self.assertEqual(self.picker.call_args.args[1:3], (self.reports, (".csv",)))

    def test_reporting_year_comes_from_reports_even_if_an_older_preference_exists(self):
        self.seed(year=2025)
        screen = self.session("2", "", "q")
        self.assertIn("CALPADS completion reporting year · 2025–26 (Aug 16, 2025–Aug 15, 2026)", screen)
        self.assertNotIn("--year", self.child_args())
        self.assertNotIn("year", terminal_guide.load_locations())
        self.report.write_text(
            self.report.read_text(encoding="utf-8").replace("2026-06-12", "2027-07-12"), encoding="utf-8"
        )
        self.assertIn("CALPADS completion reporting year · 2026–27", self.session("q"))

    def test_mixed_reporting_cycles_show_where_the_inputs_differ(self):
        self.seed()
        other = self.reports / "summer.csv"
        other.write_text(self.report.read_text(encoding="utf-8").replace("2026-06-12", "2026-08-16"), encoding="utf-8")
        screen = self.session("q")
        for fragment in ("report.csv", "summer.csv", "Row 2", "different reporting cycle"):
            self.assertIn(fragment, screen)
        self.assertNotIn("2026-08-16", screen)
        self.assertNotIn("0000000001", screen)

    def test_returning_to_menu_reuses_year_until_a_report_changes(self):
        self.seed()
        with patch.object(terminal_guide, "reporting_year", wraps=terminal_guide.reporting_year) as infer:
            self.session("1", "1", "q")
            self.assertEqual(infer.call_count, 1)
        earlier = self.report.read_bytes()

        def replace_report(*args):
            self.report.write_bytes(earlier.replace(b"2026-06-12", b"2027-06-12") + b"\n")
            return 0

        self.converter.side_effect = replace_report
        with patch.object(terminal_guide, "reporting_year", wraps=terminal_guide.reporting_year) as infer:
            screen = self.session("1", "q")
            self.assertIn("CALPADS completion reporting year · 2026–27", screen)
            self.assertEqual(infer.call_count, 2)

    def test_first_run_mixed_files_can_be_corrected_before_creating_workbook(self):
        self.seed()
        self.review.unlink()
        other = self.reports / "second.csv"
        other.write_bytes(self.report.read_bytes().replace(b"2026-06-12", b"2027-06-12"))
        original_answer = terminal_guide.answer

        def correct_on_retry(prompt):
            if "Retry after correcting" in prompt:
                other.unlink()
                return "r"
            return original_answer(prompt)

        with patch.object(terminal_guide, "answer", side_effect=correct_on_retry) as prompts:
            screen = self.session("q")
        self.assertTrue(any("Choose another input folder" in call.args[0] for call in prompts.call_args_list))
        self.assertIn("CALPADS completion reporting year · 2025–26", screen)
        self.assertIn("Create a new workbook", screen)
        self.picker.assert_not_called()
        self.converter.assert_not_called()

    def test_preferences_use_the_operating_system_user_config_folder(self):
        for platform, variables, expected in (
            ("darwin", {}, self.root / "Library" / "Application Support"),
            ("win32", {"APPDATA": str(self.root / "Roaming")}, self.root / "Roaming"),
            ("linux", {"XDG_CONFIG_HOME": str(self.root / "Config")}, self.root / "Config"),
            ("linux", {"XDG_CONFIG_HOME": "relative"}, self.root / ".config"),
        ):
            with (
                self.subTest(platform=platform, settings=variables),
                patch.object(sys, "platform", platform),
                patch.object(Path, "home", return_value=self.root),
                patch.dict(os.environ, variables),
            ):
                self.assertEqual(terminal_guide.preferences_path(), expected / "calpads-to-nsc" / "locations.json")

    def test_dialog_error_does_not_default_to_a_missing_supplement(self):
        self.seed()
        self.picker.side_effect = DialogUnavailable()
        with (
            patch.object(sys, "stdin", io.StringIO("2\ny\n")),
            redirect_stdout(io.StringIO()),
            redirect_stderr(io.StringIO()) as errors,
        ):
            self.assertEqual(terminal_guide.main(), 1)
        self.assertIn("file dialog", errors.getvalue())
        self.converter.assert_not_called()

    def test_mac_open_dialog_filters_csv_and_passes_paths_as_arguments(self):
        folder = self.root / "District's input files"
        folder.mkdir()
        source = folder / "Selected supplement.csv"
        with patch.object(sys, "platform", "darwin"), patch.object(subprocess, "run") as run:
            run.return_value = subprocess.CompletedProcess([], 0, str(source) + "\n", "")
            self.assertEqual(native_picker.pick_path("Choose a supplement", folder, (".csv",)), source)
            argv = run.call_args.args[0]
            self.assertEqual(argv[:2], ["/usr/bin/osascript", "-e"])
            self.assertEqual(argv[3:], ["Choose a supplement", str(folder), "file", "csv"])
            self.assertNotIn(str(folder), argv[2])
            run.return_value = subprocess.CompletedProcess([], 0, "\n", "")
            self.assertIsNone(native_picker.pick_path("Choose folder", folder, choose_folder=True))
            self.assertEqual(run.call_args.args[0][-1], "folder")
            native_picker.pick_path("Save the test file", folder, (".txt",), save_name="TEST_Graduates.txt")
            self.assertEqual(run.call_args.args[0][-2:], ["save", "TEST_Graduates.txt"])
