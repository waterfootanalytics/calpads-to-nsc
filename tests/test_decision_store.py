import json
import os
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from unittest.mock import patch

from calpads_to_nsc.decision_store import DecisionStore, fingerprint
from calpads_to_nsc.decisions import ReviewDecisions
from calpads_to_nsc.generate import ValidationError


class DecisionStoreTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        self.book = self.root / "schools.xlsx"
        self.scope = fingerprint("graduation-decisions-v1", "account-a", 2026)

    def test_saved_completion_choices_are_scoped_to_district_and_cycle(self):
        choices = (("account-a", 2025, "exclude"), ("account-a", 2026, "Other"), ("account-b", 2026, "Regular Diploma"))
        for account, year, choice in choices:
            reviewer = ReviewDecisions(self.book, account, year, None)
            reviewer.store.save("completion:250", choice)
        for account, year, choice in choices:
            reviewer = ReviewDecisions(self.book, account, year, None)
            self.assertEqual(reviewer.store.load("completion:250"), choice)

    def test_concurrent_saves_merge_distinct_keys_and_detect_competing_changes(self):
        stores = [DecisionStore(self.book, self.scope) for _ in range(3)]
        with ThreadPoolExecutor(max_workers=2) as pool:
            a = pool.submit(stores[0].save, "completion:102", "Other")
            b = pool.submit(stores[1].save, "completion:250", "Regular Diploma")
            a.result()
            b.result()
        with self.assertRaisesRegex(ValidationError, "Another session changed"):
            stores[2].save("completion:250", "exclude")
        merged = DecisionStore(self.book, self.scope)
        self.assertEqual(merged.load("completion:102"), "Other")
        self.assertEqual(merged.load("completion:250"), "Regular Diploma")

    def test_date_decision_moves_with_folder_and_changed_evidence_requires_review(self):
        original = self.root / "original"
        original.mkdir()
        book = original / self.book.name
        reviewer = ReviewDecisions(book, "01234", 2026, original / "supplement.csv", interactive=True)
        ssid, birth, leave, evidence = "0000123456", date(2007, 1, 1), date(2026, 6, 1), [(8, "20250815")]
        with patch("builtins.input", return_value="1"):
            reviewer.date_source(
                (original / "roster.csv", 3), ssid, birth, leave, evidence, "Review", allow_diploma=True
            )
        private = book.with_suffix(".decisions.json")
        if os.name == "posix":
            self.assertEqual(private.stat().st_mode & 0o777, 0o600)
        stored = private.read_text()
        for value in (ssid, birth.isoformat(), leave.isoformat(), "20250815", "roster.csv", "supplement.csv"):
            self.assertNotIn(value, stored)
        moved = self.root / "moved"
        original.rename(moved)
        reader = ReviewDecisions(moved / book.name, "01234", 2026, moved / "supplement.csv")
        args = ((moved / "roster.csv", 3), ssid, birth, leave)
        self.assertEqual(reader.date_source(*args, evidence, "Review", allow_diploma=True), "calpads")
        with self.assertRaises(ValidationError):
            reader.date_source(*args, [(8, "20250814")], "Review", allow_diploma=True)

    def test_invalid_decision_file_is_preserved_and_its_contents_stay_private(self):
        private = self.book.with_suffix(".decisions.json")
        unsupported = [
            '{"private value"',
            json.dumps({"version": 2, "decisions": {}}),
            json.dumps({"version": True, "decisions": {}}),
            json.dumps({"version": 1, "decisions": {"private value": {}}}),
            json.dumps({"version": 1, "decisions": {self.scope: {"completion:250": "private value"}}}),
        ]
        for content in unsupported:
            with self.subTest(content=content):
                private.write_text(content)
                with self.assertRaises(ValidationError) as error:
                    DecisionStore(self.book, self.scope)
                self.assertNotIn("private value", str(error.exception))
                self.assertEqual(private.read_text(), content)
