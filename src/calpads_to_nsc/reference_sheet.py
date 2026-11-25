import re
import subprocess
import unicodedata
from contextlib import ExitStack, closing
from pathlib import Path
from typing import TypedDict
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from .generate import ValidationError, field, require


class AccountReference(TypedDict):
    account_number: str
    organization_name: str
    schools: dict[str, str]


def single_label(text: str, pattern: str) -> str:
    values = set(re.findall(pattern, text, re.MULTILINE))
    require(len(values) == 1, "Missing or inconsistent reference sheet account fields.")
    return values.pop().strip()


def read_reference(path: Path) -> AccountReference:
    try:
        return read_account_document(path)
    except ValidationError as exc:
        raise ValidationError(f"Account reference · {path.name}\n{exc}") from None
    except (BadZipFile, ParseError, ValueError, KeyError, IndexError):
        raise ValidationError(
            f"Account reference · {path.name}\nThe document could not be read. Export a fresh account sheet."
        ) from None


def read_account_document(path: Path) -> AccountReference:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        try:
            import xlrd
        except ImportError:
            raise ValidationError("Install xlrd in your Python environment to read .xls account sheets.") from None
        try:
            book = xlrd.open_workbook(str(path), on_demand=True)
            try:
                return excel_reference(
                    (s.name, (s.row_values(i) for i in range(s.nrows))) for s in book.sheets() if s.visibility == 0
                )
            finally:
                book.release_resources()
        except (xlrd.biffh.XLRDError, xlrd.compdoc.CompDocError):
            raise ValidationError("The .xls file could not be read. Export a fresh account sheet.") from None
    if suffix == ".xlsx":
        from openpyxl import load_workbook

        book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            with ExitStack() as streams:
                sheets = (
                    (s.title, streams.enter_context(closing(s.iter_rows(values_only=True))))
                    for s in book.worksheets
                    if s.sheet_state == "visible"
                )
                return excel_reference(sheets)
        finally:
            book.close()
    require(suffix == ".pdf", "Supply an NSC .xls, .xlsx or PDF account reference sheet.")
    try:
        process = subprocess.run(
            ["pdftotext", "-layout", "-enc", "UTF-8", str(path.resolve()), "-"], capture_output=True, timeout=30
        )
    except FileNotFoundError:
        raise ValidationError("Install Poppler to read an account reference PDF.") from None
    except subprocess.TimeoutExpired:
        raise ValidationError("Reference PDF extraction timed out.") from None
    require(process.returncode == 0, "Unable to extract the reference PDF.")
    try:
        return parse_reference(process.stdout.decode("utf-8"))
    except ValidationError as exc:
        if pdftotext_is_xpdf():
            raise ValidationError(
                f"{exc}\nThe pdftotext command is from Xpdf. Place Poppler’s executable directory ahead of Xpdf "
                "on PATH and try again."
            ) from None
        raise


def pdftotext_is_xpdf() -> bool:
    try:
        version = subprocess.run(["pdftotext", "-v"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=5)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return b"xpdfreader.com" in version.stdout.lower()


def cell_text(value: object) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return "" if value is None else str(value).strip()


def registered_name(value):
    require(
        isinstance(value, str)
        and value.strip()
        and all(c.isprintable() or unicodedata.category(c) == "Zs" for c in value),
        "The registered school name must fit in one text field.",
    )
    return value


ACCOUNT_LABELS = {
    "ORGANIZATION NAME": "name",
    "SERVICE AGREEMENT/ACCOUNT NUMBER": "account",
    "TOTAL HIGH SCHOOL COUNT": "count",
}
SCHOOL_HEADINGS = ("HIGH SCHOOL NAME", "ACT CODE", "HS OPERATING STATUS")


def collect_account_labels(values, keys, metadata, row, location):
    for index, label in enumerate(keys):
        if label not in ACCOUNT_LABELS:
            continue
        key = ACCOUNT_LABELS[label]
        item = next((value for value in values[index + 1 :] if value), "")
        require(item, location + f"{label.title()} is blank. Export the complete account sheet.")
        if key in metadata:
            raise ValidationError(location + f"{label.title()} is repeated. Check row {metadata[key][1]} as well.")
        metadata[key] = item, row


def collect_campus(cells, values, columns, schools, seen, row, location):
    name_col, code_col, status_col = columns
    require(len(values) > max(columns), location + "Incomplete school row. Export the full campus table.")
    name, code, status = values[name_col], values[code_col], values[status_col].upper()
    if not any((name, code, status)):
        return
    if isinstance(cells[code_col], (int, float)) and code.isascii() and code.isdigit():
        code = code.zfill(6)
    require(re.fullmatch("[0-9]{6}", code), location + "ACT Code must contain six digits. Check the account sheet.")
    if code in seen:
        raise ValidationError(
            location + f"The ACT code is a duplicate. Check row {seen[code]} on this worksheet as well."
        )
    require(name, location + "High School Name is blank. Check the account sheet.")
    require(
        status in ("OPEN", "CLOSED"), location + "HS Operating Status must be Open or Closed. Check the account sheet."
    )
    seen[code] = row
    if status == "OPEN":
        try:
            schools[code] = registered_name(cells[name_col])
        except ValidationError as exc:
            raise ValidationError(location + str(exc)) from None


def account_summary(sheet, metadata, schools, count):
    for label, key in ACCOUNT_LABELS.items():
        require(key in metadata, f"{sheet} · Missing {label.title()}. Export the original account sheet.")
    number, row = metadata["account"]
    require(re.fullmatch("[0-9]{1,10}", number), f"{sheet} · Row {row} · Check the Service Agreement/Account Number.")
    total, row = metadata["count"]
    require(
        re.fullmatch("[0-9]+", total) and (total.lstrip("0") or "0") == str(count),
        f"{sheet} · Row {row} · The total high school count must match the campus table.",
    )
    require(schools, f"{sheet} · No open campuses were found. Check that this is the current account sheet.")
    name, row = metadata["name"]
    try:
        organization = field(name, 80, "account name")
    except ValidationError as exc:
        raise ValidationError(f"{sheet} · Row {row} · {exc}") from None
    return dict(account_number=number, organization_name=organization, schools=schools)


def account_worksheet(sheet_name, rows):
    metadata, schools, seen = {}, {}, {}
    columns = None
    sheet = f"Worksheet {sheet_name}"
    for row, cells in enumerate(rows, 1):
        values = [cell_text(v) for v in cells]
        if not any(values):
            continue
        location = f"{sheet} · Row {row} · "
        if columns is not None:
            collect_campus(cells, values, columns, schools, seen, row, location)
            continue
        keys = [" ".join(v.upper().rstrip(":").split()) for v in values]
        if set(SCHOOL_HEADINGS).issubset(keys):
            require(
                all(keys.count(k) == 1 for k in SCHOOL_HEADINGS),
                location + "A school table heading is repeated. Export the original account sheet.",
            )
            columns = tuple(keys.index(k) for k in SCHOOL_HEADINGS)
        else:
            collect_account_labels(values, keys, metadata, row, location)
    if columns is not None:
        return account_summary(sheet, metadata, schools, len(seen))
    return None


def excel_reference(sheets):
    matches = [account for name, rows in sheets if (account := account_worksheet(name, rows)) is not None]
    require(len(matches) == 1, "Expected one NSC high school table on a visible worksheet.")
    return matches[0]


def parse_reference(text: str) -> AccountReference:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    number = single_label(text, r"^[ \t]*Account Number:[ \t]*([0-9]{1,10})[ \t]*$")
    org = single_label(text, r"^[ \t]*Account Name:[ \t]*([^\r\n]+)$")
    expected = single_label(text, r"^[ \t]*Participating Organizations:[ \t]*([0-9]+)[ \t]*$").lstrip("0")
    schools = {}
    for page_number, page in enumerate(text.split("\f"), 1):
        if page.strip():
            read_pdf_table(page, page_number, schools)
    require(
        schools and str(len(schools)) == expected,
        "Reference table does not match its participating organization count.",
    )
    opened = {}
    for code, school in schools.items():
        if school["status"] == "OPEN":
            try:
                opened[code] = registered_name(school["nsc_school_name"])
            except ValidationError as exc:
                raise ValidationError(f"{school['location']} · {exc}") from None
    return {"account_number": number, "organization_name": field(org, 80, "account name"), "schools": opened}


def read_pdf_table(page, page_number, schools):
    lines = page.splitlines()
    header = next((line for line in lines if re.match(r"^#\s+Type\s+Name\s+City\s+State\s+ACT Code", line)), None)
    require(header, f"Page {page_number} · Unsupported reference sheet table layout. Export the Excel account sheet.")
    city = header.index("City")
    active = None
    start = lines.index(header) + 1
    for line_number, line in enumerate(lines[start:], start + 1):
        if line.startswith("Page "):
            break
        if not line.strip():
            continue
        left = line[:city].strip()
        row = re.fullmatch(r"(\d+)\s+HIGH SCHOOL\s+(.+)", left)
        if row:
            count = len(schools) + 1
            location = f"Page {page_number} · School row {count}"
            require(
                row[1].lstrip("0") == str(count),
                f"{location} · The numbered rows are out of sequence. Check the original campus table.",
            )
            values = re.fullmatch(r".+?\s+[A-Z]{2}\s+([0-9]{6})\s+.+?\s+(\S+)\s*", line[city:])
            require(values, f"{location} · Check the six-digit ACT code and the alignment of the campus entry.")
            code, status = values.groups()
            require(status in ("OPEN", "CLOSED"), f"{location} · Check the campus operating status.")
            if code in schools:
                original = schools[code]["location"]
                raise ValidationError(f"{location} · The ACT code is a duplicate. Also check {original}.")
            active = code
            schools[active] = {"nsc_school_name": row[2], "status": status, "location": location}
        elif left:
            require(
                active and not re.match(r"^\d", left),
                f"Page {page_number} · Line {line_number} · Unrecognized school row. Check the original campus table.",
            )
            schools[active]["nsc_school_name"] += " " + left


def school_key(value: str) -> str:
    value = re.sub(r"\bhigh school\b", "hs", value.casefold())
    return re.sub(r"[^\w]", "", value)


def match_school(label, reference):
    matches = [(code, name) for code, name in reference["schools"].items() if school_key(label) == school_key(name)]
    if len(matches) == 1:
        code, name = matches[0]
        return {"act_code": code, "nsc_school_name": name}
    return {"act_code": "", "nsc_school_name": ""}
