import re
import unicodedata
from pathlib import Path

from .generate import DetailColumn
from .reports import row_error


PROGRAM_FIELDS = {
    "Grade8MathAssessment": DetailColumn.GRADE8_MATH_ASSESSMENT,
    "Grade8ELAAssessment": DetailColumn.GRADE8_ELA_ASSESSMENT,
    "HighSchoolMathAssessment": DetailColumn.HIGH_SCHOOL_MATH_ASSESSMENT,
    "HighSchoolELAAssessment": DetailColumn.HIGH_SCHOOL_ELA_ASSESSMENT,
    "MathSemesters": DetailColumn.MATH_SEMESTERS,
    "DualEnrollment": DetailColumn.DUAL_ENROLLMENT,
    "ProgramCode": DetailColumn.PROGRAM_CODE,
}

ProgramRow = tuple[tuple[Path, int], dict[str, str]]


def valid_program_code(value: str) -> bool:
    canonical = unicodedata.normalize("NFKC", value)
    return (
        len(value) <= 50
        and value.isprintable()
        and not canonical.lstrip().startswith(("=", "+", "@"))
        and not any(c in "'\"−" or unicodedata.category(c) in {"Pd", "Pi", "Pf"} for c in canonical)
    )


def program_value(name: str, value: str, location: tuple[Path, int]) -> str:
    if name == "MathSemesters":
        if re.fullmatch(r"[0-9]{1,2}", value):
            return str(int(value))
        instruction = "Enter a whole number from 0 to 99."
    elif name == "DualEnrollment":
        flag = {"Y": "Y", "YES": "Y", "N": "N", "NO": "N"}.get(value.upper())
        if flag:
            return flag
        instruction = "Enter Y or N."
    elif name == "ProgramCode":
        if value.casefold() == "n/a":
            raise row_error(
                location, "Check ProgramCode. Leave this cell blank when no code applies.", "District supplement"
            )
        if valid_program_code(value):
            return value
        instruction = (
            "Use up to 50 printable characters without quotes or dashes. Separate codes with _ or . and avoid "
            "starting with =, + or @ when reviewing the file in a spreadsheet."
        )
    else:
        if value in {"1", "2", "3"}:
            return value
        instruction = "Enter the result on the NSC assessment scale of 1, 2 or 3."
    raise row_error(location, f"Check {name}. {instruction} Leave it blank if unknown.", "District supplement")


def merged_program_fields(rows: list[ProgramRow]) -> dict[DetailColumn, str]:
    result = {}
    first_rows = {}
    for location, fields in rows:
        for name, raw in fields.items():
            column = PROGRAM_FIELDS[name]
            value = program_value(name, raw, location)
            if column in result and result[column] != value:
                raise row_error(
                    location,
                    f"Conflicting {name} entries. Also check Row {first_rows[column]} and supply one value.",
                    "District supplement",
                )
            result[column] = value
            first_rows.setdefault(column, location[1])
    return result
