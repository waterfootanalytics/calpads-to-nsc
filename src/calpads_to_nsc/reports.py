import csv
import re
from collections.abc import Iterable
from contextlib import closing
from datetime import date, datetime
from functools import lru_cache
from itertools import islice
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from .generate import ValidationError, field, require

DEMOGRAPHIC_ALIASES = {
    "gender": ("Gender", "GenderCode", "StudentGender", "StudentGenderCode"),
    "ethnicity": ("EthnicityRace", "RaceEthnicity", "Ethnicity", "ReportingEthnicity", "StudentEthnicityRace"),
    "economic_disadvantage": (
        "SocioEconomicallyDisadvantaged",
        "SocioEconomicallyDisadvantagedIndicator",
        "SocioEconomicallyDisadvantagedStatus",
        "EconomicallyDisadvantaged",
        "SED",
    ),
    "english_learner": (
        "ELAS",
        "ELASCode",
        "ELASStatus",
        "EnglishLanguageAcquisitionStatus",
        "EnglishLanguageAcquisitionStatusCode",
        "EnglishLearnerStatus",
        "EnglishLearnerIndicator",
    ),
    "disability": (
        "SpecialEducationStatus",
        "SpecialEducation",
        "SpecialEducationIndicator",
        "SpecialEd",
        "SpecialEdStatus",
        "StudentWithDisabilities",
        "StudentsWithDisabilities",
        "SWD",
    ),
}

ALIASES = {
    "school": ("SchoolName",),
    "ssid": ("SSID", "StateStudentID"),
    "local": ("LocalID", "LocalStudentID"),
    "combined": ("StudentName", "StudentLegalName"),
    "first": ("FirstName", "StudentFirstName", "LegalFirstName"),
    "middle": ("MiddleName", "StudentMiddleName", "LegalMiddleName"),
    "last": ("LastName", "StudentLastName", "LegalLastName"),
    "dob": ("BirthDate", "DateOfBirth", "DOB"),
    "exit": ("StudentEnrollmentExitDate", "EnrollmentExitDate", "ExitDate"),
    "reason": ("StudentExitReason", "StudentExitReasonCode", "ExitReason"),
    "status": (
        "StudentCompletionStatus",
        "SchoolCompletionStatus",
        "StudentSchoolCompletionStatus",
        "SchoolCompletionStatusCode",
        "StudentSchoolCompletionStatusCode",
        "CompletionStatus",
    ),
    "ferpa": ("FERPABlock",),
    **DEMOGRAPHIC_ALIASES,
}


HEADER_ROW_LIMIT = 25
DELIMITERS = ("\t", "^", "|", ",")
HEADER_FIELDS = {label.lower(): key for key, names in ALIASES.items() for label in names}


def norm(value):
    return re.sub("[^a-z0-9]", "", str(value).lower())


def columns(line):
    result = {}
    ambiguous = set()
    for position, value in enumerate(line):
        key = HEADER_FIELDS.get(norm(value))
        if key is None:
            continue
        if key in result:
            require(key in DEMOGRAPHIC_ALIASES, f"Ambiguous report header for {key}.")
            ambiguous.add(key)
        result[key] = position
    for key in ambiguous:
        del result[key]
    require(
        {"school", "ssid", "dob", "exit", "status"} <= result.keys(),
        "Missing school, SSID, birth date, enrollment exit date, or completion status header.",
    )
    require({"first", "last"} <= result.keys() or "combined" in result, "Missing legal name headers.")
    return result


def header_score(values):
    return sum(norm(value) in HEADER_FIELDS for value in values if value is not None)


def complete_header(values):
    try:
        return columns(values)
    except ValidationError:
        return None


def delimited_header(line):
    candidates = []
    for delimiter in DELIMITERS:
        if delimiter not in line:
            continue
        try:
            cells = next(csv.reader([line], delimiter=delimiter, strict=True))
        except csv.Error:
            continue
        score = header_score(cells)
        if score >= 2:
            candidates.append((bool(complete_header(cells)), score, delimiter, cells))
    if not candidates:
        return False, "", ()
    candidates.sort(reverse=True)
    complete, _, delimiter, cells = candidates[0]
    require(
        not complete or len(candidates) == 1 or not candidates[1][0],
        "The report header has ambiguous delimiters. Export one flat delimited table.",
    )
    return complete, delimiter, cells


def csv_rows(handle, delimiter, offset):
    reader = csv.reader(handle, delimiter=delimiter, strict=True)
    for values in reader:
        yield offset + reader.line_num, values


def delimited_rows(handle, *, calpads):
    if calpads:
        for number, line in enumerate(islice(handle, HEADER_ROW_LIMIT), 1):
            complete, delimiter, cells = delimited_header(line)
            yield number, cells
            if complete:
                yield from csv_rows(handle, delimiter, number)
                return
    else:
        first = handle.readline()
        separators = [c for c in DELIMITERS if c in first]
        require(len(separators) == 1, "Expected one flat, delimited report with a header.")
        handle.seek(0)
        yield from csv_rows(handle, separators[0], 0)


def input_rows(path, *, calpads=False):
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            require(len(workbook.worksheets) == 1, "Export one flat report worksheet.")
            yield from enumerate(workbook.worksheets[0].values, 1)
        finally:
            workbook.close()
    else:
        require(path.suffix.lower() in (".csv", ".tsv", ".txt"), "Use XLSX or delimited text.")
        with path.open("rb") as raw:
            prefix = raw.read(2)
        encoding = "utf-16" if prefix in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
        with path.open(encoding=encoding, newline="", errors="replace") as handle:
            yield from delimited_rows(handle, calpads=calpads)


def report_files(source):
    if not source.is_dir():
        return [source]
    paths = sorted(p for p in source.iterdir() if not p.name.startswith((".", "~$")))
    require(paths, "No reports found in the input folder.")
    require(
        all(p.is_file() and p.suffix.lower() in (".xlsx", ".txt", ".csv", ".tsv") for p in paths),
        "Use a folder containing only flat XLSX, TXT, CSV or TSV exports.",
    )
    return paths


def workbook_role(path):
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if {"Match schools", "Look up a school"} <= set(workbook.sheetnames):
            return "review"
        reference = False
        for sheet in workbook.worksheets:
            if sheet.sheet_state != "visible":
                continue
            for cells in islice(sheet.values, HEADER_ROW_LIMIT):
                if header_score(cells) >= 2 and complete_header(cells):
                    return "reports"
                labels = {norm(v) for v in cells if v is not None}
                if {"organizationname", "serviceagreementaccountnumber", "totalhighschoolcount"} & labels:
                    reference = True
        return "reference" if reference else None
    finally:
        workbook.close()


def has_supplement_header(path):
    from .supplement import is_supplement_header

    stream = input_rows(path)
    try:
        _, cells = next(stream, (1, ()))
        return is_supplement_header(cells)
    except (ValidationError, csv.Error, UnicodeError):
        return False
    finally:
        stream.close()


@lru_cache(maxsize=256)
def input_role(path, file_version):
    suffix = path.suffix.lower()
    if suffix == ".txt":
        with path.open("rb") as handle:
            if handle.read(4) == b"PH3\t":
                return "graduates"
    if suffix in (".pdf", ".xls"):
        return "reference"
    if suffix == ".xlsx":
        try:
            role = workbook_role(path)
        except (BadZipFile, ValueError, ParseError, OSError, KeyError):
            role = None
        if role:
            return role
        if "accountreference" in norm(path.stem):
            return "reference"
    if suffix == ".csv" and has_supplement_header(path):
        return "dates"
    return "reports"


def folder_inputs(folder, *, reference=None, review=None, supplement=None):
    inputs = {"reports": [], "reference": [], "dates": []}
    for path in sorted(folder.iterdir(), key=lambda item: item.name.casefold()):
        if path.name.startswith((".", "~$")) or path.suffix.lower() not in (
            ".csv",
            ".tsv",
            ".txt",
            ".pdf",
            ".xls",
            ".xlsx",
        ):
            continue
        if path == review or not path.is_file():
            continue
        if path == reference:
            role = "reference"
        elif path == supplement:
            role = "dates"
        else:
            stat = path.stat()
            role = input_role(path, (stat.st_ino, stat.st_mtime_ns, stat.st_size))
        if role in inputs:
            inputs[role].append(path)
    return inputs


def row_error(location, message, source="CALPADS file"):
    path, number = location
    return ValidationError(f"{source} · {path.name}\nRow {number} · {message}")


def located_rows(path, source, *, calpads=False):
    stream = input_rows(path, calpads=calpads)
    number = 1
    try:
        while True:
            try:
                end, values = next(stream)
            except StopIteration:
                return
            except ValidationError as exc:
                raise row_error((path, number), str(exc), source) from None
            except (csv.Error, UnicodeError):
                raise row_error(
                    (path, number), "Cannot read this row. Check the delimiter, quoting and text encoding.", source
                ) from None
            except (BadZipFile, KeyError, ParseError, ValueError):
                raise row_error(
                    (path, number), "Cannot read the workbook. Export a new flat copy and retry.", source
                ) from None
            if any(isinstance(value, str) and "\ufffd" in value for value in values):
                raise row_error(
                    (path, number), "Unreadable character. Check the source encoding and export again.", source
                )
            yield number, values
            number = end + 1
    finally:
        stream.close()


def report_header(path, rows):
    problem = None
    best_score = 0
    for number, values in islice(rows, HEADER_ROW_LIMIT):
        score = header_score(values)
        if score < 2:
            continue
        try:
            return columns(values)
        except ValidationError as exc:
            if score > best_score:
                problem = row_error((path, number), str(exc))
                best_score = score
    if problem is not None:
        raise problem
    raise row_error(
        (path, 1),
        f"No CALPADS column headings found in the first {HEADER_ROW_LIMIT} rows. "
        "Export the certified student list with its field names on one row.",
    )


def demographic_headers(paths: Iterable[Path]) -> set[str]:
    found = set()
    for path in paths:
        with closing(located_rows(path, "CALPADS file", calpads=True)) as rows:
            found.update(report_header(path, rows).keys() & DEMOGRAPHIC_ALIASES.keys())
    return found


def report_records(paths, headers_only=False):
    for path in paths:
        rows = located_rows(path, "CALPADS file", calpads=True)
        try:
            indexes = report_header(path, rows)
            if headers_only:
                continue
            last_required = max(i for key, i in indexes.items() if key not in DEMOGRAPHIC_ALIASES)
            for number, values in rows:
                if not any(v is not None and str(v).strip() for v in values):
                    continue
                location = (path, number)
                if len(values) <= last_required:
                    raise row_error(location, "Incomplete report row. Export all fields listed in the header.")
                record = {key: values[i] if i < len(values) else None for key, i in indexes.items()}
                block = record.get("ferpa")
                if block is not None and str(block).strip():
                    raise row_error(
                        location,
                        "FERPABlock belongs in the district supplement. Use the certified "
                        "CALPADS 1.23 export and move the flag to the matching supplement row.",
                    )
                yield location, record
        finally:
            rows.close()


def report_schools(paths):
    schools = set()
    for location, record in report_records(paths):
        school = record["school"]
        if not isinstance(school, str) or not school.strip():
            raise row_error(location, "Missing school name. Check the School Name field in the certified report.")
        try:
            schools.add(field(school.strip(), 255, "school name"))
        except ValidationError:
            raise row_error(
                location, "Invalid school name. Check the School Name field in the certified report."
            ) from None
    require(schools, "No school names found in the reports.")
    return schools


def parse_calendar_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = str(value).strip()
    patterns = (
        (r"[0-9]{8}", "%Y%m%d"),
        (r"[0-9]{4}-[0-9]{2}-[0-9]{2}", "%Y-%m-%d"),
        (r"[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}", "%m/%d/%Y"),
    )
    for pattern, fmt in patterns:
        if re.fullmatch(pattern, raw):
            try:
                return datetime.strptime(raw, fmt).date()
            except ValueError:
                break
    raise ValidationError("Invalid report date.")


def report_date(location, value, label):
    if value is None or not str(value).strip():
        raise row_error(location, f"Missing {label}. Check the source record and export it again.")
    try:
        return parse_calendar_date(value)
    except ValidationError:
        raise row_error(
            location, f"Invalid {label}. Use a calendar date with a four-digit year and export it again."
        ) from None


def completion_year_label(year: int) -> str:
    start = year - 1
    return f"CALPADS completion reporting year · {start}–{year % 100:02d} (Aug 16, {start}–Aug 15, {year})"


def cycle_for_exit(exit_day):
    return exit_day.year + ((exit_day.month, exit_day.day) > (8, 15))


def reporting_year(paths, expected=None):
    first_year = None
    first = None
    for location, record in report_records(paths):
        year = cycle_for_exit(report_date(location, record["exit"], "enrollment exit date"))
        if not 1901 <= year <= 9999:
            raise row_error(
                location, "Enrollment exit date is outside the supported reporting years. Check the export."
            )
        if first_year is not None and year != first_year:
            raise row_error(
                location,
                f"Enrollment exit date is in a different reporting cycle from "
                f"{first[0].name}, Row {first[1]}. Choose reports for one cycle and retry.",
            )
        if expected is not None and year != expected:
            raise row_error(
                location,
                f"Enrollment exit date is outside the selected CALPADS cycle. "
                f"Check the report or the requested reporting year ({expected}).",
            )
        if first_year is None:
            first_year, first = year, location
    require(first_year is not None, "No CALPADS student rows found. Choose a report with enrollment exit dates.")
    return first_year


def validate_ssid(value, location, source):
    if not value:
        raise row_error(location, "Missing SSID. Supply the ten-digit state student ID and export again.", source)
    if not re.fullmatch("[0-9]{10}", value):
        raise row_error(location, "Malformed SSID. Use ten digits and preserve any leading zeroes.", source)


def record_unique(value, seen, location, label):
    if value in seen:
        previous, number = seen[value]
        raise row_error(
            location,
            f"Duplicate {label}. First found in {previous.name}, Row {number}. Check both rows and correct the inputs.",
        )
    seen[value] = location
