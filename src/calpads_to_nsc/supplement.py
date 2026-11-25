import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date

from .decisions import PENDING
from .generate import DetailColumn, ValidationError, day, require
from .programs import PROGRAM_FIELDS, ProgramRow, merged_program_fields
from .reports import cycle_for_exit, located_rows, norm, parse_calendar_date, row_error, validate_ssid


def diploma_date_result(value):
    if isinstance(value, date):
        return value.strftime("%Y%m%d"), "Compared"
    raw = str(value).strip() if value is not None else ""
    if not raw:
        return None, "Blank diploma date in district supplement"
    if re.fullmatch(r"[0-9]{1,2}/[0-9]{1,2}/[0-9]{2}", raw):
        return None, "Two-digit year in diploma export"
    if re.fullmatch(r"[0-9]{8}|[0-9]{4}-[0-9]{2}-[0-9]{2}|[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}", raw):
        try:
            return parse_calendar_date(raw).strftime("%Y%m%d"), "Compared"
        except ValidationError:
            pass
    return None, "Invalid date in diploma export"


def ferpa_flag(value):
    flag = "" if value is None else str(value).strip().upper()
    require(flag in ("", "Y", "N"), "Check FERPABlock in the district supplement. Enter Y, N, or leave it blank.")
    return flag


@dataclass
class Supplement:
    dates: dict = field(default_factory=lambda: defaultdict(list))
    blocked: set = field(default_factory=set)
    has_block_column: bool = False
    has_block_value: bool = False
    has_date_column: bool = False
    program_columns: tuple[str, ...] = ()
    program_rows: dict[str, list[ProgramRow]] = field(default_factory=dict)

    def program_fields(self, ssid: str) -> dict[DetailColumn, str]:
        return merged_program_fields(self.program_rows.get(ssid, []))


SUPPLEMENT_HEADERS = {norm(name): name for name in ("SSID", "GraduationDate", "FERPABlock", *PROGRAM_FIELDS)}


def is_supplement_header(header) -> bool:
    labels = {norm(value) for value in header}
    return "ssid" in labels and "schoolname" not in labels and bool((labels - {"ssid"}) & SUPPLEMENT_HEADERS.keys())


def supplement_columns(path, header):
    labels = [norm(v) for v in header]
    unique = set(labels)
    if "ssid" not in unique or len(unique) < 2 or len(unique) != len(labels) or not unique <= SUPPLEMENT_HEADERS.keys():
        raise row_error(
            (path, 1),
            "Use SSID and at least one supported supplement column. Check the headers for misspellings or duplicates. "
            "See the district supplement field list.",
            "District supplement",
        )
    return [SUPPLEMENT_HEADERS[label] for label in labels]


def add_supplement_row(result, location, labels, values):
    source = "District supplement"
    if len(values) != len(labels):
        raise row_error(
            location, "The row has a different number of fields than the header. Check the CSV columns.", source
        )
    row = dict(zip(labels, values))
    try:
        flag = ferpa_flag(row.get("FERPABlock"))
    except ValidationError as exc:
        raise row_error(location, str(exc), source) from None
    ssid = str(row["SSID"] or "").strip()
    if not re.fullmatch("[0-9]{10}", ssid):
        if flag == "Y":
            validate_ssid(ssid, location, source)
        return False
    if flag:
        result.has_block_value = True
    if flag == "Y":
        result.blocked.add(ssid)
    if result.has_date_column:
        result.dates[ssid].append((location[1], *diploma_date_result(row["GraduationDate"])))
    fields = {name: str(row[name]).strip(" ") for name in result.program_columns if row[name] is not None}
    fields = {name: value for name, value in fields.items() if value}
    if fields:
        result.program_rows.setdefault(ssid, []).append((location, fields))
    return True


def read_supplement(path):
    rows = located_rows(path, "District supplement")
    skipped = False
    try:
        labels = supplement_columns(path, next(rows, (1, ()))[1])
        result = Supplement(
            has_block_column="FERPABlock" in labels,
            has_date_column="GraduationDate" in labels,
            program_columns=tuple(name for name in PROGRAM_FIELDS if name in labels),
        )
        for row_number, values in rows:
            if any(v is not None and str(v).strip() for v in values):
                if not add_supplement_row(result, (path, row_number), labels, values):
                    skipped = True
    finally:
        rows.close()
    if skipped:
        print("Skipped district supplement rows with a missing or malformed SSID and no indicated FERPA block.")
    return result


@dataclass(frozen=True)
class DateSelection:
    selected: date | None
    source: str
    status: str
    candidate: date | None = None


def review_date(location, ssid, exit_day, birth_day, dated_rows, candidates, decisions):
    if len(candidates) > 1:
        choice = decisions.date_source(
            location,
            ssid,
            birth_day,
            exit_day,
            dated_rows,
            "The supplement lists different diploma dates for this graduate.",
            allow_diploma=False,
        )
        if choice is PENDING:
            return DateSelection(None, "", "Needs decision: conflicting supplemental diploma dates")
        return DateSelection(exit_day, "CALPADS", "Confirmed CALPADS date for conflicting supplemental diploma dates")
    award_day = next(iter(candidates.values()))
    if cycle_for_exit(award_day) == cycle_for_exit(exit_day):
        return DateSelection(award_day, "District supplement", "Compared", award_day)
    choice = decisions.date_source(
        location,
        ssid,
        birth_day,
        exit_day,
        dated_rows,
        "The diploma date falls outside this report's completion cycle.",
        allow_diploma=True,
    )
    if choice is PENDING:
        return DateSelection(None, "", "Needs decision: diploma date outside reporting cycle", award_day)
    if choice == "supplement":
        return DateSelection(
            award_day, "District supplement", "Confirmed diploma date outside reporting cycle", award_day
        )
    return DateSelection(exit_day, "CALPADS", "Confirmed CALPADS date for out-of-cycle diploma date", award_day)


def resolve_graduation_date(dates, ssid, exit_day, birth_day, send, location, decisions):
    if dates is None:
        return DateSelection(exit_day, "CALPADS", "Using CALPADS exit dates")
    entries = dates.get(ssid, ())
    readable = {value for _, value, _ in entries if value}
    if not readable:
        statuses = {status for _, _, status in entries}
        status = "SSID missing from district supplement" if not entries else "No readable supplemental diploma date"
        return DateSelection(exit_day, "CALPADS", statuses.pop() if len(statuses) == 1 else status)
    candidates = {value: day(value, "graduation date") for value in readable}
    candidates = {value: award for value, award in candidates.items() if birth_day < award <= send}
    if not candidates:
        return DateSelection(exit_day, "CALPADS", "Supplemental date outside birth-to-transmission range")
    dated_rows = [(row, value) for row, value, _ in entries if value in candidates]
    return review_date(location, ssid, exit_day, birth_day, dated_rows, candidates, decisions)


def select_graduation_date(dates, ssid, exit_day, birth_day, send, location, decisions):
    result = resolve_graduation_date(dates, ssid, exit_day, birth_day, send, location, decisions)
    if result.selected is not None and not birth_day < result.selected <= send:
        raise row_error(
            location,
            "Check Birth Date and the selected graduation date. Graduation must be after birth "
            "and on or before the file transmission date.",
        )
    return result
