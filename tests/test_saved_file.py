import io
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from calpads_to_nsc import terminal_guide
from calpads_to_nsc.generate import ValidationError
from calpads_to_nsc.saved_file import NAME_NOTICE, SCHOOL_NOTICE, check_saved_file


class SavedFileTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.path = Path(temp.name) / "Graduates_sample.txt"
        self.header = ["PH3", "01234", "Example District", "P", "20240630", "2024"]
        self.detail = [
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
        self.trailer = ["PT3", "3"]

    def write(self, *rows):
        self.path.write_bytes("".join("\t".join(row) + "\r\n" for row in rows).encode())

    def test_independent_cli_and_layout_warnings(self):
        self.detail[2] = "First-Name"
        self.detail[13] = "A" * 51
        self.write(self.header, self.detail, self.trailer)
        self.assertEqual(check_saved_file(self.path), {NAME_NOTICE, SCHOOL_NOTICE})
        result = subprocess.run(
            [sys.executable, "-m", "calpads_to_nsc", "--check-file", self.path], capture_output=True, text=True
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Saved file layout checked", result.stdout)
        self.assertNotIn("First-Name", result.stdout)

    def test_structural_and_field_errors_never_print_field_values(self):
        cases = [
            (self.header[:-1], self.detail, self.trailer),
            (self.header, self.detail[:-1], self.trailer),
            (self.header, self.detail),
            (self.header, self.detail, ["PT3", "2"]),
            (self.header, self.detail, self.trailer, self.detail),
            (self.header, ["PT3", "2"]),
        ]
        for col, value in [
            (1, "bad"),
            (2, "X" * 41),
            (8, "20000231"),
            (9, "Private-ID"),
            (11, "20240631"),
            (12, "Maybe"),
            (14, "12345"),
            (15, "X"),
            (18, "4"),
            (25, "Maybe"),
            (26, "=X"),
            (27, "BAD"),
        ]:
            row = self.detail.copy()
            row[col] = value
            cases.append((self.header, row, self.trailer))
        for rows in cases:
            with self.subTest(rows=rows):
                self.write(*rows)
                with self.assertRaises(ValidationError) as ctx:
                    check_saved_file(self.path)
                self.assertNotIn("Private-ID", str(ctx.exception))

    def test_blocks_upload_instructions_if_saved_file_is_invalid(self):
        self.write(self.header, self.detail, ["PT3", "900"])
        output = io.StringIO()
        with (
            redirect_stdout(output),
            patch("builtins.input", return_value="c"),
            patch.object(terminal_guide, "show_full_upload") as upload,
        ):
            terminal_guide.review_saved_file(self.path, sample=False)
        upload.assert_not_called()
        self.assertIn("PT3 must count every row", output.getvalue())
