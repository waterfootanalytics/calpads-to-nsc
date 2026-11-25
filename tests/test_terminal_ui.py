import csv
import io
import os
import re
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from calpads_to_nsc import calpads, workbook_review


SGR = re.compile(r"\033\[[0-9;]*m")


class Terminal(io.StringIO):
    def isatty(self):
        return True


class TerminalStyleTest(unittest.TestCase):
    def test_color_requires_tty_and_respects_terminal_settings(self):
        for stream, term, no_color, expected in (
            (Terminal(), "xterm-256color", None, True),
            (io.StringIO(), "xterm-256color", None, False),
            (Terminal(), "dumb", None, False),
            (Terminal(), "xterm-256color", "", False),
            (Terminal(), "xterm-256color", "1", False),
        ):
            with self.subTest(tty=stream.isatty(), term=term, no_color=no_color):
                with patch.dict(os.environ, {"TERM": term}):
                    if no_color is None:
                        os.environ.pop("NO_COLOR", None)
                    else:
                        os.environ["NO_COLOR"] = no_color
                    label = workbook_review.terminal_style("Match schools", "36", stream)
                self.assertEqual(SGR.sub("", label), "Match schools")
                self.assertEqual(bool(SGR.search(label)), expected)
                if expected:
                    self.assertTrue(label.endswith("\033[0m"))


class ReviewDisplayTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        self.review = root / "School review.xlsx"
        account = root / "account.xlsx"
        reports = root / "reports.csv"
        self.choices = dict(
            zip(
                ("Acorn Academy", "Birch Academy", "Cedar Academy", "Delta Academy"),
                ("NORTH HS", "SOUTH HS", "EAST HS", "WEST HS"),
            )
        )
        self.long_school = "MEADOW REGIONAL HIGH SCHOOL FOR APPLIED SCIENCES AND ENGINEERING"
        with reports.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow(
                [
                    "SchoolName",
                    "SSID",
                    "FirstName",
                    "LastName",
                    "BirthDate",
                    "StudentEnrollmentExitDate",
                    "StudentCompletionStatus",
                ]
            )
            for index, campus in enumerate([*self.choices, self.long_school, self.long_school], 1):
                writer.writerow([campus, f"{index:010d}", "Alex", "Sample", "2008-06-01", "2026-06-12", "100"])
        book = Workbook()
        for row in (
            ("Organization Name:", "Example District"),
            ("Service Agreement/Account Number:", "001"),
            ("Total High School Count:", 5),
            ("ACT Code", "HS Operating Status", "High School Name"),
        ):
            book.active.append(row)
        for index, campus in enumerate([*self.choices.values(), self.long_school], 1):
            book.active.append((f"{index:06d}", "OPEN", campus))
        book.save(account)
        book.close()
        self.cli = [
            "calpads_to_nsc",
            str(reports),
            "--reference-sheet",
            str(account),
            "--year",
            "2026",
            "--date",
            "20260630",
            "--confirm-no-ferpa-blocks",
        ]
        with patch.object(sys, "argv", [*self.cli, "--create-crosswalk", str(self.review), "--no-prompt"]):
            with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                calpads.entrypoint()

    def check(self, stream_type):
        out = stream_type()
        err = stream_type()
        with patch.dict(os.environ, {"TERM": "xterm-256color"}), redirect_stdout(out), redirect_stderr(err):
            os.environ.pop("NO_COLOR", None)
            with patch.object(sys, "argv", [*self.cli, "--crosswalk", str(self.review), "--check-crosswalk"]):
                with patch.object(workbook_review, "offer_workbook", return_value=False):
                    try:
                        calpads.entrypoint()
                    except SystemExit as exc:
                        return exc.code, out.getvalue() + err.getvalue()
        return 0, out.getvalue() + err.getvalue()

    def save_school_matches(self):
        book = load_workbook(self.review)
        try:
            for index, campus in enumerate(self.choices.values(), 1):
                book["Match schools"][f"D{12 + index}"] = f"{campus} ({index:06d})"
            book.save(self.review)
        finally:
            book.close()

    def test_school_review_finishes_once_every_campus_is_matched(self):
        for stream_type in (Terminal, io.StringIO):
            with self.subTest(stream=stream_type.__name__):
                code, message = self.check(stream_type)
                self.assertNotEqual(code, 0)
                self.assertEqual(bool(SGR.search(message)), stream_type is Terminal)
                plain = SGR.sub("", message)
                self.assertIn("Match schools", plain)
                for row, campus in enumerate(self.choices, 13):
                    self.assertRegex(plain, rf"(?m)\bD{row}\b[^\n]*{re.escape(campus)}")
                self.assertLessEqual(plain.lower().count("choose the matching clearinghouse school"), 1)
        self.save_school_matches()
        code, message = self.check(io.StringIO)
        self.assertEqual(code, 0)
        self.assertNotRegex(message, SGR)
        self.assertIn("School matches checked", message)
        status, message = self.terminal_session([""], "--check-crosswalk")
        self.assertEqual(status, 0)

    def terminal_session(self, responses, *mode):
        out, err = Terminal(), Terminal()
        answers = iter(responses)

        def respond(prompt):
            out.write(prompt + "\n")
            return next(answers)

        with redirect_stdout(out), redirect_stderr(err), patch.object(sys.stdin, "isatty", return_value=True):
            with patch.object(sys, "argv", [*self.cli, "--crosswalk", str(self.review), "--terminal-review", *mode]):
                with (
                    patch("builtins.input", side_effect=respond),
                    patch.object(workbook_review, "open_workbook") as opener,
                ):
                    try:
                        calpads.entrypoint()
                        status = 0
                    except SystemExit as exc:
                        status = exc.code
                    opener.assert_not_called()
        self.assertEqual(list(answers), [])
        return status, SGR.sub("", out.getvalue() + err.getvalue())

    def test_terminal_review_saves_each_choice_then_resumes_and_writes_full_file(self):
        status, transcript = self.terminal_session(["000001", "f"], "--check-crosswalk")
        self.assertEqual(status, 3)
        self.assertIn("Flag saved", transcript)
        book = load_workbook(self.review)
        try:
            self.assertEqual(book["Match schools"]["D13"].value, "NORTH HS (000001)")
            self.assertEqual(book["Match schools"]["D14"].value, "Flag for review")
        finally:
            book.close()
        output = self.review.with_name("Graduates_2026.txt")
        status, transcript = self.terminal_session(["000002", "000003", "000004"], "--output", str(output))
        self.assertEqual(status, 0)
        self.assertIn("Graduates file saved", transcript)
        rows = [line.split("\t") for line in output.read_text().splitlines()]
        self.assertEqual(
            {row[14]: row[13] for row in rows[1:-1]},
            {**{f"{i:06d}": school for i, school in enumerate(self.choices.values(), 1)}, "000005": self.long_school},
        )

    def test_terminal_can_change_an_existing_match(self):
        status, _ = self.terminal_session(["000001", "m", "1", "000002", "q"], "--check-crosswalk")
        self.assertEqual(status, 3)
        book = load_workbook(self.review)
        try:
            self.assertEqual(book["Match schools"]["D13"].value, "SOUTH HS (000002)")
        finally:
            book.close()

    def test_exit_from_nested_school_list_preserves_saved_matches(self):
        self.save_school_matches()
        previous = self.review.read_bytes()
        for exit_key in ("q", "Close"):
            with self.subTest(exit_key=exit_key):
                status, screen = self.terminal_session(["m", exit_key], "--check-crosswalk")
                self.assertEqual(status, 3)
                self.assertEqual(screen.count("Every school has a match."), 1)
                self.assertEqual(self.review.read_bytes(), previous)

    def test_full_file_save_flows_into_review_then_secure_ftp_instructions(self):
        self.save_school_matches()
        status, screen = self.terminal_session(["c", ""], "--output", str(self.review.parent), "--review-full")
        self.assertEqual(status, 0)
        (saved,) = self.review.parent.glob("Graduates_*.txt")
        for step in (
            "Graduates file saved",
            "Review the full graduates file in ",
            "Submit your full graduates file",
            "FTP Login",
            str(saved.resolve()),
        ):
            self.assertIn(step, screen)
        self.assertLess(screen.index("Graduates file saved"), screen.index("I reviewed the file"))
        self.assertLess(screen.index("I reviewed the file"), screen.index("Submit your full graduates file"))
        self.assertNotIn("Submit a Test File", screen)

    def test_guide_requires_explicit_no_block_confirmation_for_full_file(self):
        self.save_school_matches()
        self.cli.remove("--confirm-no-ferpa-blocks")
        output = self.review.with_name("full.txt")
        for key, code in (("m", 5), ("q", 3), ("CONFIRM", 0)):
            entries = [key, ""] if key == "CONFIRM" else [key]
            status, screen = self.terminal_session(entries, "--guide-session", "--output", str(output))
            self.assertEqual(status, code)
            self.assertIn("Type CONFIRM", screen)
            self.assertEqual(output.is_file(), key == "CONFIRM")

    def test_pre_save_summary_can_return_to_menu_without_writing(self):
        self.save_school_matches()
        output = self.review.with_name("full.txt")
        status, screen = self.terminal_session(["m"], "--guide-session", "--output", str(output))
        self.assertEqual(status, 5)
        self.assertIn("Review before saving", screen)
        self.assertIn("Excluded campuses", screen)
        self.assertIn("Completion code 100", screen)
        self.assertFalse(output.exists())

    def test_guide_full_file_reports_ferpa_for_dates_only_supplement(self):
        self.save_school_matches()
        self.cli.remove("--confirm-no-ferpa-blocks")
        supplement = self.review.with_name("district.csv")
        supplement.write_text("SSID,GraduationDate\n0000000001,20260610\n")
        output = self.review.with_name("full.txt")
        status, screen = self.terminal_session(
            ["CONFIRM", ""], "--guide-session", "--supplement", str(supplement), "--output", str(output)
        )
        self.assertEqual(status, 0)
        self.assertIn("Type CONFIRM", screen)
        graduate = output.read_text().splitlines()[1].split("\t")
        self.assertEqual(graduate[11:13], ["20260610", "N"])

    def test_reference_spelling_with_a_length_warning_in_test_and_full_files(self):
        self.save_school_matches()
        original = self.review.read_bytes()
        for mode, filename, expected in (("--test-output", "Sample.txt", 1), ("--output", "Full.txt", 2)):
            with self.subTest(mode=mode):
                output = self.review.with_name(filename)
                status, screen = self.terminal_session([], mode, str(output))
                self.assertEqual(status, 0)
                warnings = [line for line in screen.splitlines() if "warning" in line.lower()]
                self.assertEqual(len(warnings), 1)
                self.assertIn("50", warnings[0])
                self.assertIn("column N", warnings[0])
                detail = f"ACT 000005  ·  {len(self.long_school)} characters  ·  {self.long_school}"
                self.assertEqual(screen.count(detail), 1)
                self.assertIn("exactly as supplied by NSC", screen)
                self.assertIn("may return a field-length validation error", screen)
                rows = [line.split("\t") for line in output.read_text().splitlines()[1:-1]]
                self.assertEqual([row[13] for row in rows if row[14] == "000005"], [self.long_school] * expected)
                self.assertEqual(self.review.read_bytes(), original)
