import json
import re
import sys
from collections.abc import Iterable, Sequence

from .demographics import demographic_values
from .generate import DIPLOMAS, DetailColumn, ValidationError, day, detail_row, require, submission
from .reports import cycle_for_exit, norm, record_unique, report_date, report_records, row_error, validate_ssid
from .saved_file import NAME_NOTICE, given_name_warning
from .supplement import select_graduation_date


def student_names(row):
    if (row.get("first") and row.get("last")) or ("first" in row and "last" in row and not row.get("combined")):
        return row["first"], row.get("middle", ""), row["last"]
    parts = row.get("combined", "").replace("，", ",").split(",")
    require(
        len(parts) == 2 and all(p.strip() for p in parts),
        "Use Last, First Middle or supply separate legal first and last name columns.",
    )
    given = parts[1].split(maxsplit=1)
    return given[0], given[1] if len(given) == 2 else "", parts[0].strip()


def graduate_test_sample(rows: list[list[str]], account: dict[str, object], send_date: str) -> list[list[str]]:
    by_school = {}
    for detail in rows[1:-1]:
        by_school.setdefault(detail[DetailColumn.ACT_CODE], detail)
    require(len(by_school) <= 1000, "The Graduate test uploader allows up to 1,000 student records per file.")
    sample = list(by_school.values())
    years = [row[DetailColumn.GRADUATION_DATE][:4] for row in sample]
    return submission(account, send_date, {min(years), max(years)}, sample)


DEFAULT_DIPLOMAS = {"100": "Regular Diploma", "120": None, "320": None, "330": None}


def completion_overrides(policy_path):
    if policy_path is None:
        return {}
    settings = json.loads(policy_path.read_text(encoding="utf-8-sig"))
    require(
        isinstance(settings, dict) and set(settings) == {"completion_types"},
        "Optional settings must contain only the completion_types object.",
    )
    overrides = settings["completion_types"]
    require(
        isinstance(overrides, dict)
        and all(
            re.fullmatch("[0-9]{3}", k) and (v is None or isinstance(v, str) and v in DIPLOMAS)
            for k, v in overrides.items()
        ),
        "Invalid completion type mapping.",
    )
    return overrides


def completion_code(row, location):
    if "reason" in row and not (re.match(r"^E230\b", row["reason"], re.I) or norm(row["reason"]) == "completerexit"):
        raise row_error(
            location, "Unexpected exit reason. Verify that the export is the certified 1.23 completer list."
        )
    status = re.match(r"^([0-9]{3})(?:\b|$)", row["status"])
    if not status:
        raise row_error(location, "Completion status needs a three-digit CALPADS code. Check the certified report.")
    return status[1]


def eligible_graduates(paths, year, diplomas, deferred, decisions=None):
    seen = {}
    for location, record in report_records(paths):
        row = {key: str(v).strip() if v is not None else "" for key, v in record.items()}
        if row["school"] in deferred:
            continue
        ssid = row["ssid"]
        validate_ssid(ssid, location, "CALPADS file")
        record_unique(ssid, seen, location, "SSID")
        exit_day = report_date(location, record["exit"], "enrollment exit date")
        if cycle_for_exit(exit_day) != year:
            raise row_error(
                location, "Enrollment exit date is outside the selected CALPADS cycle. Check the report and year."
            )
        code = completion_code(row, location)
        if decisions is None and code not in diplomas:
            raise row_error(
                location,
                f"Completion code {code} needs a district decision. Review completion codes "
                "in the guide or supply completion_types in --settings.",
            )
        diploma = decisions.completion(code, location, diplomas) if decisions else diplomas[code]
        if diploma is not None:
            yield location, record, row, exit_day, diploma


def student_identifier(row, location, policy, seen):
    if policy == "blank":
        return ""
    if policy == "ssid":
        return row["ssid"]
    local = row.get("local", "")
    if local:
        if not re.fullmatch("[A-Za-z0-9]{1,15}", local):
            raise row_error(
                location,
                "Malformed local student ID. Use 1–15 letters or digits, leave it blank, "
                "or choose a different student ID option.",
            )
        record_unique(local, seen, location, "local student ID")
    return local


def graduate_detail(row, location, birth, award, school, identifier, notices):
    try:
        first, middle, last = student_names(row)
        student = dict(first_name=first, middle_name=middle, last_name=last, birth_date=birth.strftime("%Y%m%d"))
        return detail_row(student, identifier, award, school, notices)
    except ValidationError as exc:
        raise row_error(location, f"{exc} Check the legal name in the source record and export again.") from None


def warn_school_lengths(details: Iterable[Sequence[str]]) -> None:
    long_names = sorted(
        {
            (row[DetailColumn.ACT_CODE], row[DetailColumn.HIGH_SCHOOL_NAME])
            for row in details
            if len(row[DetailColumn.HIGH_SCHOOL_NAME]) > 50
        }
    )
    if long_names:
        print(
            "Warning — NSC specifies a maximum of 50 characters for the high school name in column N.", file=sys.stderr
        )
        print("These names on the NSC Account Reference Sheet exceed that limit.", file=sys.stderr)
        for code, registered in long_names:
            print(f"  ACT {code}  ·  {len(registered)} characters  ·  {registered}", file=sys.stderr)
        print(
            "File preparation continues with each registered name exactly as supplied by NSC.\n"
            "NSC may return a field-length validation error when the file is submitted.",
            file=sys.stderr,
        )


def optional_detail_fields(args, row, supplement, notices):
    fields = {}
    if args.include_demographics:
        fields.update(
            demographic_values(row, include_disability=args.use_special_education_for_disability, notices=notices)
        )
    if not args.calpads_exit_dates:
        fields.update(supplement.program_fields(row["ssid"]))
    return fields


def convert(args, paths, reviewed, diplomas, deferred, supplement_dates, supplement, decisions):
    send = day(args.date, "transmission date")
    details, years, local_ids = [], set(), {}
    notices, date_sources, fallbacks = set(), set(), set()
    for location, record, row, exit_day, diploma in eligible_graduates(paths, args.year, diplomas, deferred, decisions):
        school = reviewed["schools"].get(row["school"])
        if school is None:
            raise row_error(
                location, "School Name is missing from the reviewed workbook. Review the school matches again."
            )
        birth_day = report_date(location, record["dob"], "birth date")
        selection = select_graduation_date(
            supplement_dates, row["ssid"], exit_day, birth_day, send, location, decisions
        )
        date_sources.add(selection.source)
        if selection.source == "CALPADS" and supplement_dates is not None:
            fallbacks.add(selection.status)
        grad = selection.selected.strftime("%Y%m%d")
        years.add(grad[:4])
        flag = "Y" if row["ssid"] in supplement.blocked else "N"
        identifier = student_identifier(row, location, args.student_id, local_ids)
        award = dict(graduation_date=grad, diploma_type=diploma, ferpa_block=flag)
        detail = graduate_detail(row, location, birth_day, award, school, identifier, notices)
        for column, value in optional_detail_fields(args, row, supplement, notices).items():
            detail[column] = value
        if given_name_warning(detail):
            notices.add(NAME_NOTICE)
        details.append(detail)
    require(
        details,
        "No graduates remain after applying school and completion decisions. "
        "Review school matches and graduation decisions, then retry.",
    )
    require(len(details) <= 250000, "The cohort exceeds the full file limit of 250,000 graduates.")
    show_cohort_review(args, details, deferred, date_sources, fallbacks, supplement, notices)
    return submission(reviewed, args.date, {min(years), max(years)}, details)


def show_cohort_review(args, details, deferred, date_sources, fallbacks, supplement, notices):
    print("\nReview before saving" if args.output or args.test_output else "\nGraduates file review")
    print("Excluded campuses · " + (", ".join(sorted(deferred)) or "None"))
    print("Graduation date sources used: " + ", ".join(sorted(date_sources)) + ".")
    for reason in sorted(fallbacks):
        print("Kept CALPADS exit dates · " + reason)
    if supplement.has_block_column and not supplement.has_block_value:
        print("FERPABlock has no Y or N entries on usable supplement rows. All graduates will have N (unknown).")
    for notice in sorted(notices):
        print(notice)
    warn_school_lengths(details)
