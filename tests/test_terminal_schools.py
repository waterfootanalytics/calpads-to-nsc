import io
import os
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager, redirect_stdout
from pathlib import Path
from threading import Barrier
from unittest.mock import patch

from openpyxl import load_workbook

from calpads_to_nsc.crosswalk import EXCLUDE, SchoolEdit, create, reviewed_config
from calpads_to_nsc import terminal_schools
from calpads_to_nsc.terminal_input import Finished


class SchoolEditorTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.path = self.root / "review.xlsx"
        self.reference = dict(
            account_number="001", organization_name="Example", schools={"012345": "PARK HS", "012346": "LAKE HS"}
        )
        self.labels = {"Former Academy", "Park High School"}
        create(self.path, self.labels, self.reference, "fingerprint", year=2026)
        self.first = SchoolEdit("Match schools", "D13", "Former Academy")
        self.second = SchoolEdit("Match schools", "D14", "Park High School")
        self.lake = "LAKE HS (012346)"
        self.park = "PARK HS (012345)"

    def save(self, edit, version, value):
        return terminal_schools.save_choice(self.path, self.reference, edit, version, value)

    def test_changed_workbook_during_save_is_reloaded_and_temp_file_is_cleaned(self):
        replacement = self.root / "from-excel.xlsx"
        book = load_workbook(self.path)
        book[self.first.sheet][self.first.cell] = "Flag for review"
        book.save(replacement)
        book.close()
        latest = replacement.read_bytes()
        version = self.path.stat()
        lock = terminal_schools.commit_lock

        @contextmanager
        def simultaneous_save(path):
            with lock(path):
                os.replace(replacement, path)
                yield

        with patch.object(terminal_schools, "commit_lock", simultaneous_save), redirect_stdout(io.StringIO()):
            self.assertEqual(self.save(self.first, version, self.lake), "retry")
        self.assertEqual(self.path.read_bytes(), latest)
        self.assertEqual(list(self.root.iterdir()), [self.path])

    def test_two_operators_starting_from_same_version_cannot_overwrite_each_other(self):
        original = self.path.stat()
        start = Barrier(2)

        def change(edit):
            start.wait(timeout=4)
            return self.save(edit, original, self.lake)

        with redirect_stdout(io.StringIO()), ThreadPoolExecutor(2) as workers:
            results = list(workers.map(change, (self.first, self.second)))
        self.assertCountEqual(results, ["saved", "retry"])
        book = load_workbook(self.path, read_only=True)
        try:
            sheet = book[self.first.sheet]
            self.assertIn(
                (sheet[self.first.cell].value, sheet[self.second.cell].value),
                ((self.lake, self.park), (None, self.lake)),
            )
        finally:
            book.close()
        self.assertEqual(list(self.root.iterdir()), [self.path])

    def test_excel_open_marker_pauses_without_changing_workbook(self):
        self.path.with_name("~$" + self.path.name).touch()
        before = self.path.read_bytes()
        with redirect_stdout(io.StringIO()):
            status = self.save(self.first, self.path.stat(), self.lake)
        self.assertEqual(status, "pause")
        self.assertEqual(self.path.read_bytes(), before)

    def test_locked_destination_without_excel_marker_preserves_workbook_and_can_retry(self):
        before = self.path.read_bytes()
        version = self.path.stat()
        with patch.object(os, "replace", side_effect=PermissionError), redirect_stdout(io.StringIO()):
            self.assertEqual(self.save(self.first, version, self.lake), "pause")
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual(list(self.root.iterdir()), [self.path])
        with redirect_stdout(io.StringIO()):
            self.assertEqual(self.save(self.first, version, self.lake), "saved")

    def test_duplicate_code_is_rejected_then_an_unused_code_can_be_saved(self):
        screen = io.StringIO()
        with patch("builtins.input", side_effect=["012345", "012346"]), redirect_stdout(screen):
            status = terminal_schools.edit_school(self.path, self.reference, self.first, self.path.stat())
        self.assertEqual(status, "saved")
        self.assertIn("012345", screen.getvalue())
        self.assertIn("Park High School", screen.getvalue())
        schools = reviewed_config(self.path, self.labels, self.reference, "fingerprint", writing=True)["schools"]
        self.assertEqual(schools["Former Academy"]["act_code"], "012346")
        self.assertEqual(schools["Park High School"]["act_code"], "012345")

    def test_terminal_records_and_reverses_a_confirmed_school_exclusion(self):
        initial = self.path.read_bytes()
        private_reason = "=Separate agreement"
        screen = io.StringIO()

        def edit(replies):
            with patch("builtins.input", side_effect=replies), redirect_stdout(screen):
                return terminal_schools.edit_school(self.path, self.reference, self.first, self.path.stat(), year=2026)

        with self.assertRaises(Finished):
            edit(["e", "", "e", private_reason, "Y", "q"])
        self.assertEqual(self.path.read_bytes(), initial)
        self.assertEqual(edit(["e", private_reason, "EXCLUDE"]), "saved")
        book = load_workbook(self.path)
        self.assertEqual(book["Match schools"]["D13"].value, EXCLUDE)
        self.assertEqual(book["Match schools"]["F13"].data_type, "s")
        self.assertEqual(book["Match schools"]["F13"].value, private_reason)
        self.assertEqual(book["Match schools"]["G13"].value, "EXCLUDE")
        self.assertEqual(dict(book["Source"].values)["Reporting cycle"], 2026)
        self.assertEqual(book["NSC choices"]["A5"].value, EXCLUDE)
        book.close()
        with patch("builtins.input", side_effect=["1"]), redirect_stdout(screen):
            self.assertEqual(terminal_schools.pick_saved_match(self.path, self.reference), self.first)
        self.assertEqual(edit(["012346"]), "saved")
        result = reviewed_config(self.path, self.labels, self.reference, "fingerprint", writing=True, year=2026)
        self.assertEqual(result["excluded_schools"], {})
        self.assertEqual(result["schools"]["Former Academy"]["act_code"], "012346")
        self.assertNotIn(private_reason, screen.getvalue())
