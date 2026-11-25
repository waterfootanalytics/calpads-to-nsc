import io
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import ExitStack, redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from calpads_to_nsc import calpads, workbook_review


class WorkbookPromptTest(unittest.TestCase):
    def setUp(self):
        mocks = ExitStack()
        self.addCleanup(mocks.close)
        mocks.enter_context(redirect_stdout(io.StringIO()))
        self.messages = mocks.enter_context(redirect_stdout(io.StringIO()))
        self.stdin = mocks.enter_context(patch.object(sys.stdin, "isatty", return_value=True))
        self.stdout = mocks.enter_context(patch.object(sys.stdout, "isatty", return_value=True))
        mocks.enter_context(patch.object(sys, "platform", "darwin"))
        self.answer = mocks.enter_context(patch("builtins.input"))
        self.open_app = mocks.enter_context(patch.object(subprocess, "run"))
        self.open_app.return_value = subprocess.CompletedProcess([], 0)
        self.startfile = mocks.enter_context(patch.object(os, "startfile", create=True))
        private = tempfile.TemporaryDirectory()
        self.addCleanup(private.cleanup)
        self.path = Path(private.name, "School review (2026).xlsx")
        self.path.touch()

    def test_unrecognized_answer_is_reprompted_then_enter_opens_the_exact_path(self):
        self.answer.side_effect = ["later?", ""]
        self.assertTrue(workbook_review.offer_workbook(self.path))
        self.assertEqual(self.answer.call_count, 2)
        if os.name == "nt":
            self.startfile.assert_called_once_with(str(self.path.resolve()))
        else:
            self.open_app.assert_called_once()
            argv = self.open_app.call_args.args[0]
            self.assertEqual(Path(argv[0]).name, "open")
            self.assertEqual(Path(argv[-1]).resolve(), self.path.resolve())
            self.assertFalse(self.open_app.call_args.kwargs.get("shell", False))

    def test_decline_or_end_of_input_preserves_the_saved_workbook(self):
        for response in (["n"], EOFError):
            with self.subTest(response=response):
                self.answer.side_effect = response
                self.assertFalse(workbook_review.offer_workbook(self.path))
                self.assertTrue(self.path.exists())
                self.open_app.assert_not_called()
                self.startfile.assert_not_called()

    def test_redirected_input_or_output_disables_both_prompt_and_opener(self):
        for stdin, stdout in ((True, False), (False, True), (False, False)):
            with self.subTest(stdin=stdin, stdout=stdout):
                self.stdin.return_value, self.stdout.return_value = stdin, stdout
                self.assertFalse(workbook_review.offer_workbook(self.path))
                self.answer.assert_not_called()
                self.open_app.assert_not_called()
                self.startfile.assert_not_called()

    def test_unavailable_opener_returns_control(self):
        self.answer.return_value = "yes"
        self.open_app.side_effect = FileNotFoundError()
        self.startfile.side_effect = FileNotFoundError()
        self.assertFalse(workbook_review.offer_workbook(self.path))
        self.assertTrue(self.path.exists())
        self.assertIn(str(self.path.resolve()), self.messages.getvalue())


class ReviewCommandTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        root = Path(directory.name)
        report = root / "calpads.csv"
        account = root / "account.xlsx"
        self.review = root / "School Review.xlsx"
        self.output = root / "graduates.txt"
        report.write_text(
            "SchoolName,SSID,FirstName,LastName,BirthDate,EnrollmentExitDate,StudentCompletionStatus\n"
            "Cedar Academy,0000000001,Alex,Sample,2008-06-01,2026-06-12,100\n",
            encoding="utf-8",
        )
        book = Workbook()
        for row in (
            ("Organization Name:", "Sample District"),
            ("Service Agreement/Account Number:", "001"),
            ("Total High School Count:", 1),
            ("ACT Code", "HS Operating Status", "High School Name"),
            ("012345", "OPEN", "MEADOW HS"),
        ):
            book.active.append(row)
        book.save(account)
        book.close()
        self.cli = ["calpads_to_nsc", str(report), "--reference-sheet", str(account), "--year", "2026"]

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, "-B", "-m", *self.cli, *args],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            timeout=10,
        )

    def save_choice(self):
        book = load_workbook(self.review)
        book["Match schools"]["D13"] = "MEADOW HS (012345)"
        book.save(self.review)
        book.close()

    def test_noninteractive_create_check_and_saved_edit(self):
        created = self.run_cli("--create-crosswalk", str(self.review))
        self.assertEqual(created.returncode, 0, created.stderr)
        self.assertTrue(self.review.is_file())
        self.assertIn(str(self.review.resolve()), created.stdout + created.stderr)
        self.assertIn("Match schools", created.stdout + created.stderr)
        checked = self.run_cli("--crosswalk", str(self.review), "--check-crosswalk")
        self.assertNotEqual(checked.returncode, 0)
        message = checked.stdout + checked.stderr
        for hint in ("Match schools", "D13", "Cedar Academy", str(self.review.resolve())):
            self.assertIn(hint, message)
        self.assertNotIn("0000000001", message)
        self.save_choice()
        passed = self.run_cli("--crosswalk", str(self.review), "--check-crosswalk")
        self.assertEqual(passed.returncode, 0, passed.stderr)

    def test_no_prompt_leaves_edit_for_later_without_opening_or_writing(self):
        self.assertEqual(self.run_cli("--create-crosswalk", str(self.review)).returncode, 0)
        argv = [*self.cli, "--no-prompt", "--crosswalk", str(self.review), "--output", str(self.output)]
        with patch.object(sys, "argv", argv), patch.object(workbook_review, "offer_workbook") as opener:
            with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as result:
                calpads.main()
        self.assertNotEqual(result.exception.code, 0)
        opener.assert_not_called()
        self.assertFalse(self.output.exists())

    def test_interactive_generation_rechecks_until_saved_and_continues_same_command(self):
        created = self.run_cli("--create-crosswalk", str(self.review))
        self.assertEqual(created.returncode, 0, created.stderr)
        argv = [
            *self.cli,
            "--crosswalk",
            str(self.review),
            "--output",
            str(self.output),
            "--date",
            "20260630",
            "--confirm-no-ferpa-blocks",
        ]
        with ExitStack() as mocks:
            mocks.enter_context(patch.object(sys, "argv", argv))
            mocks.enter_context(redirect_stdout(io.StringIO()))
            mocks.enter_context(redirect_stderr(io.StringIO()))
            opened = mocks.enter_context(patch.object(workbook_review, "offer_workbook", return_value=True))
            save_actions = iter([lambda: None, self.save_choice])

            def recheck():
                self.assertFalse(self.output.exists())
                next(save_actions)()
                return True

            rechecks = mocks.enter_context(patch.object(workbook_review, "ready_to_recheck", side_effect=recheck))
            calpads.main()
        opened.assert_called_once_with(self.review.resolve())
        self.assertEqual(rechecks.call_count, 2)
        rows = self.output.read_text().splitlines()
        self.assertEqual(rows[1].split("\t")[13:15], ["MEADOW HS", "012345"])
        self.assertEqual(rows[-1], "PT3\t3")
