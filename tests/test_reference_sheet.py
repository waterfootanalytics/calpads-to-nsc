import os
import shutil
import struct
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

from calpads_to_nsc.generate import ValidationError
from calpads_to_nsc.reference_sheet import match_school, parse_reference, read_reference


def write_biff2_sheet(path, rows):
    def record(kind, data=b""):
        return struct.pack("<HH", kind, len(data)) + data

    records = [record(0x0009, struct.pack("<HH", 2, 0x10)), record(0x0042, struct.pack("<H", 1252))]
    for row_index, row in enumerate(rows):
        for column_index, text in enumerate(row):
            label = text.encode("ascii")
            cell = struct.pack("<HH3xB", row_index, column_index, len(label)) + label
            records.append(record(0x0004, cell))
    records.append(record(0x000A))
    path.write_bytes(b"".join(records))


def write_text_pdf(path, lines):
    stream = ["BT", "/F1 9 Tf", "12 TL", "30 720 Td"]
    for line in lines:
        stream.extend([f"({line}) Tj", "T*"])
    stream.append("ET")
    content = "\n".join(stream).encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>",
        b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream",
    ]
    document = bytearray(b"%PDF-1.4\n")
    offsets = []
    for index, obj in enumerate(objects, 1):
        offsets.append(len(document))
        document.extend(f"{index} 0 obj\n".encode() + obj + b"\nendobj\n")
    startxref = len(document)
    document.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode())
    for offset in offsets:
        document.extend(f"{offset:010} 00000 n \n".encode())
    document.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{startxref}\n%%EOF\n".encode())
    path.write_bytes(document)


class ReferenceTest(unittest.TestCase):
    def setUp(self):
        header = "Account Name: TEST\nAccount Number: 001\nParticipating Organizations: 2\n"
        table = f"{'# Type Name':40}City          State ACT Code District Status\n"
        self.text = (
            header
            + table
            + f"{'1 HIGH SCHOOL SAMPLE':40}Any City      CA    012345   DISTRICT OPEN\n"
            + "              HIGH SCHOOL\nPage 1 of 2\n\f"
            + header
            + table
            + f"{'2 HIGH SCHOOL OLD SCHOOL':40}Any City      CA    012346   DISTRICT CLOSED\n"
        )

    def test_wrapped_name_and_status(self):
        for newline in ("\n", "\r\n", "\r"):
            with self.subTest(newline=newline):
                sheet = parse_reference(self.text.replace("\n", newline))
                self.assertEqual(sheet["schools"], {"012345": "SAMPLE HIGH SCHOOL"})
                self.assertEqual(match_school("Sample HS", sheet)["act_code"], "012345")
                self.assertFalse(match_school("Another school", sheet)["act_code"])

    def test_partial_or_conflicting_sheet(self):
        for text in (
            self.text.split("\f")[0],
            self.text.replace("012346", "012345"),
            self.text.replace("Account Number: 001", "Account Number: 002", 1),
        ):
            with self.assertRaises(ValidationError):
                parse_reference(text)

    def test_excel_reference_uses_open_schools_and_checks_total_and_codes(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "account.xlsx"
            book = Workbook()
            sheet = book.active
            for row in [
                (None, "Organization Name:", "TEST"),
                (None, "Service Agreement/Account Number:", "001"),
                (None, "Total High School Count:", 2),
                ("ACT Code", "HS Operating Status", "High School Name"),
                (12345, "Open", "TEST HS"),
                ("012346", "Closed", "FORMER HS"),
            ]:
                sheet.append(row)
            book.create_sheet("Metadata").sheet_state = "hidden"
            book.save(path)
            self.assertEqual(
                read_reference(path),
                dict(account_number="001", organization_name="TEST", schools={"012345": "TEST HS"}),
            )
            sheet["A6"] = "012345"
            book.save(path)
            with self.assertRaisesRegex(ValidationError, "duplicate") as problem:
                read_reference(path)
            message = str(problem.exception)
            for item in (path.name, "Worksheet Sheet", "Row 6", "row 5"):
                self.assertIn(item, message)
            for private in ("012345", "FORMER HS"):
                self.assertNotIn(private, message)
            sheet.delete_rows(6)
            book.save(path)
            with self.assertRaisesRegex(ValidationError, "count"):
                read_reference(path)

    def test_invalid_excel_cells_identify_location_and_field_without_echoing_values(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "Account.xlsx"
            book = Workbook()
            self.addCleanup(book.close)
            sheet = book.active
            sheet.title = "Campus table"
            for row in (
                ("Organization Name:", "Practice"),
                ("Service Agreement/Account Number:", "001"),
                ("Total High School Count:", 1),
                ("ACT Code", "HS Operating Status", "High School Name"),
                ("012345", "Open", "Practice School"),
            ):
                sheet.append(row)
            checks = (
                ("A5", "Private malformed code", "Row 5", "ACT Code"),
                ("B5", "Private invalid status", "Row 5", "HS Operating Status"),
                ("C5", 12345678, "Row 5", "registered school name"),
                ("B2", "Private invalid account", "Row 2", "Service Agreement"),
                ("B3", "²", "Row 3", "count"),
            )
            for cell, invalid, row, description in checks:
                with self.subTest(cell=cell):
                    original = sheet[cell].value
                    sheet[cell] = invalid
                    book.save(path)
                    with self.assertRaises(ValidationError) as problem:
                        read_reference(path)
                    message = str(problem.exception)
                    for item in (path.name, "Worksheet Campus table", row, description):
                        self.assertIn(item, message)
                    self.assertNotIn(str(invalid), message)
                    sheet[cell] = original

    def test_pdf_points_to_printed_school_and_original_duplicate_without_showing_values(self):
        path = Path("Campus table.pdf")
        for text, labels in (
            (self.text.replace("012346", "012345"), ("duplicate", "Page 1 · School row 1")),
            (self.text.replace("012346", "private"), ("ACT code",)),
            (self.text.replace("CLOSED", "private"), ("status",)),
        ):
            response = SimpleNamespace(returncode=0, stdout=text.encode("utf-8"))
            with self.subTest(labels=labels), patch("subprocess.run", return_value=response):
                with self.assertRaises(ValidationError) as problem:
                    read_reference(path)
                message = str(problem.exception)
                for detail in (path.name, "Page 2 · School row 2", *labels):
                    self.assertIn(detail, message)
                for private in ("private", "012345", "OLD SCHOOL"):
                    self.assertNotIn(private, message)

    def test_pdf_two_letter_words_in_city_still_parse(self):
        sheet = parse_reference(self.text.replace("Any City", "Casa DE Oro"))
        self.assertEqual(sheet["schools"], {"012345": "SAMPLE HIGH SCHOOL"})

    def test_pdf_layout_error_adds_poppler_path_hint_only_for_xpdf(self):
        path = Path("Account sheet.pdf")
        invalid = self.text.replace("012346", "private")
        extraction = SimpleNamespace(returncode=0, stdout=invalid.encode())
        for banner, has_hint in (
            (b"pdftotext version 4.06 [www.xpdfreader.com]\r\nCopyright Glyph & Cog, LLC\r\n", True),
            (b"Copyright The Poppler Developers\nCopyright Glyph & Cog, LLC\n", False),
        ):
            version = SimpleNamespace(returncode=99 if has_hint else 0, stdout=banner)
            with self.subTest(banner=banner), patch("subprocess.run", side_effect=[extraction, version]):
                with self.assertRaises(ValidationError) as failure:
                    read_reference(path)
                message = str(failure.exception)
                for detail in (path.name, "Page 2 · School row 2", "ACT code"):
                    self.assertIn(detail, message)
                self.assertEqual("ahead of Xpdf on PATH" in message, has_hint)
                self.assertNotIn("private", message)

    def test_xls_reads_a_binary_account_sheet(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "Practice district.xls"
            write_biff2_sheet(
                path,
                [
                    ("Organization Name:", "Practice District"),
                    ("Service Agreement/Account Number:", "0023"),
                    ("Total High School Count:", "2"),
                    ("ACT Code", "HS Operating Status", "High School Name"),
                    ("012345", "Open", "PRACTICE SCHOOL"),
                    ("012346", "Closed", "FORMER PRACTICE SCHOOL"),
                ],
            )
            self.assertEqual(
                read_reference(path),
                {
                    "account_number": "0023",
                    "organization_name": "Practice District",
                    "schools": {"012345": "PRACTICE SCHOOL"},
                },
            )

    @unittest.skipUnless(shutil.which("pdftotext"), "Poppler is optional for Excel-only installations")
    def test_pdf_reads_text_extracted_by_poppler(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "Practice district.pdf"
            write_text_pdf(
                path,
                [
                    "Account Name: Practice District",
                    "Account Number: 0023",
                    "Participating Organizations: 2",
                    f"{'# Type Name':40}City          State ACT Code District Status",
                    f"{'1 HIGH SCHOOL PRACTICE':40}Any City      CA    012345   DISTRICT OPEN",
                    "              SCHOOL",
                    f"{'2 HIGH SCHOOL FORMER':40}Any City      CA    012346   DISTRICT CLOSED",
                ],
            )
            self.assertEqual(
                read_reference(path),
                {
                    "account_number": "0023",
                    "organization_name": "Practice District",
                    "schools": {"012345": "PRACTICE SCHOOL"},
                },
            )

    def test_missing_reader_names_the_public_installation_dependency(self):
        with patch.dict(sys.modules, {"xlrd": None}), self.assertRaisesRegex(ValidationError, "Install xlrd"):
            read_reference(Path("Practice district.xls"))
        with patch.dict(os.environ, {"PATH": ""}), self.assertRaisesRegex(ValidationError, "Install Poppler"):
            read_reference(Path("Practice district.pdf"))
