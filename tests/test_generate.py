import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl.utils import get_column_letter

from calpads_to_nsc.generate import DetailColumn, ValidationError, name, save_submission


class GraduateFileTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def test_review_fields_match_clearinghouse_columns(self):
        self.assertEqual(get_column_letter(DetailColumn.GRADUATION_DATE + 1), "L")
        self.assertEqual(get_column_letter(DetailColumn.HIGH_SCHOOL_NAME + 1), "N")
        self.assertEqual(get_column_letter(DetailColumn.ACT_CODE + 1), "O")
        self.assertEqual(get_column_letter(DetailColumn.END_OF_DETAIL + 1), "AB")

    def test_explicit_output_preserves_contents_and_refuses_overwrite(self):
        output = self.root / "graduates.txt"
        rows = [["PH3", "00123", "Example District", "P", "20260630", "2026"], ["PT3", "2"]]
        self.assertEqual(save_submission(output, rows), output)
        expected = b"PH3\t00123\tExample District\tP\t20260630\t2026\r\nPT3\t2\r\n"
        self.assertEqual(output.read_bytes(), expected)
        with self.assertRaises(FileExistsError):
            save_submission(output, [["replacement"]])
        self.assertEqual(output.read_bytes(), expected)

    def test_simultaneous_automatic_names_preserve_every_submission(self):
        base = self.root / "TEST_Graduates_20261124_083315.txt"
        base.write_bytes(b"previous output")
        with patch("calpads_to_nsc.generate.datetime") as clock, ThreadPoolExecutor(max_workers=2) as pool:
            clock.now.return_value = datetime(2026, 11, 24, 8, 33, 15)
            writes = [
                pool.submit(save_submission, self.root, [[value]], name_prefix="TEST_Graduates")
                for value in ("one", "two")
            ]
            files = [write.result() for write in writes]
        self.assertEqual({p.name for p in files}, {base.stem + "_2.txt", base.stem + "_3.txt"})
        self.assertEqual([p.read_bytes() for p in files], [b"one\r\n", b"two\r\n"])
        self.assertEqual(base.read_bytes(), b"previous output")

    def test_student_name_cleanup_handles_unicode_and_is_repeatable(self):
        cases = (
            ("A’B", "A'B"),
            ("X—Y", "X-Y"),
            ("A\u0301", "A"),
            ("  A\tB  ", "A B"),
            ("Ｆ7_", "F"),
            ("D’A\u0301rcy", "D'Arcy"),
            ("  De\t\xa0la\nCruz  ", "De la Cruz"),
            ("Ó’Neill—Smith", "O'Neill-Smith"),
            ("Łéø", "Leo"),
            ("'t Hooft", "'t Hooft"),
            ("Ｍａｅ™_7!", "Mae"),
            ("Example-Smith", "Example-Smith"),
        )
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(name(raw, "last name"), expected)
                notices = set()
                self.assertEqual(name(expected, "last name", notices), expected)
                self.assertFalse(notices)
        self.assertEqual(name("...#9", "middle name"), "")
        self.assertEqual(name("N/A", "middle name"), "")

    def test_name_lengths_and_unrepresentable_fields(self):
        for label in ("first name", "middle name", "last name"):
            for size in (39, 40, 41):
                with self.subTest(label=label, size=size):
                    notices = set()
                    self.assertEqual(name("E" * size, label, notices), "E" * min(size, 40))
                    self.assertEqual(any("40-character" in notice for notice in notices), size > 40)
        for raw in ("", "123-*", "#N/A", "#REF!", "Private王", "王", "Priva\ufffdte"):
            with self.subTest(raw=raw), self.assertRaises(ValidationError) as caught:
                name(raw, "first name")
            if raw:
                self.assertNotIn(raw, str(caught.exception))
        with self.assertRaises(ValidationError):
            name("Private王", "middle name")


if __name__ == "__main__":
    unittest.main()
