import re
from pathlib import Path

from .demographics import ETHNICITY
from .generate import DIPLOMAS, DetailColumn as Col, ValidationError, day, require
from .programs import valid_program_code
from .reference_sheet import registered_name


NAME_NOTICE = (
    "NSC specifies letters only in first and middle names (C and D). "
    "A name contains a space, hyphen or apostrophe. Check the Graduate test result for this spelling."
)
SCHOOL_NOTICE = (
    "NSC specifies a maximum of 50 characters in column N. A school name exceeds that limit. "
    "Check its spelling against the NSC Account Reference Sheet. NSC may flag it during submission."
)
BLANK_OR_YN = {"", "Y", "N"}
ALLOWED = {
    Col.DIPLOMA_TYPE: DIPLOMAS,
    Col.FERPA_BLOCK: {"Y", "N"},
    Col.GENDER: {"", "M", "F"},
    Col.ETHNICITY: {"", *ETHNICITY.values()},
    Col.ECONOMIC_DISADVANTAGE: BLANK_OR_YN,
    Col.ENGLISH_LEARNER: BLANK_OR_YN,
    Col.DUAL_ENROLLMENT: BLANK_OR_YN,
    Col.DISABILITY: BLANK_OR_YN,
    Col.END_OF_DETAIL: {"ED"},
}
ASSESSMENTS = range(Col.GRADE8_MATH_ASSESSMENT, Col.HIGH_SCHOOL_ELA_ASSESSMENT + 1)
PATTERNS = {
    Col.SOCIAL_SECURITY_NUMBER: r"NO SSN|[0-9]{9}",
    Col.FIRST_NAME: r"[A-Za-z '-]{1,40}",
    Col.MIDDLE_NAME: r"[A-Za-z '-]{0,40}",
    Col.LAST_NAME: r"[A-Za-z '-]{1,40}",
    5: r"[A-Za-z]{0,3}",
    6: r"[A-Za-z]{0,40}",
    7: r"[A-Za-z]{0,40}",
    Col.STUDENT_ID: r"[A-Za-z0-9]{0,15}",
    Col.ACT_CODE: r"[0-9]{6}",
    Col.MATH_SEMESTERS: r"[0-9]{0,2}",
}


def column_label(index: int) -> str:
    return chr(65 + index) if index < 26 else "A" + chr(65 + index - 26)


def header_values(row: list[str]):
    require(len(row) == 6 and row[0] == "PH3", "Row 1 needs a six-field PH3 graduates header.")
    require(re.fullmatch(r"[0-9]{1,10}", row[1]), "Row 1 · Check the account number in B.")
    require(0 < len(row[2].strip()) <= 80 and row[2].isprintable(), "Row 1 · Check the organization name in C.")
    require(row[3] == "P", "Row 1 · File type D must be P.")
    sent = day(row[4], "transmission date (row 1, E)")
    require(re.fullmatch(r"[0-9]{4}(?:-[0-9]{4})?", row[5]), "Row 1 · Check the diploma period in F.")
    period = row[5].split("-")
    require(period[0] <= period[-1], "Row 1 · Diploma period in F is reversed.")
    return sent, period


def check_detail_fields(row, line):
    require(len(row) == 28 and row[0] == "PD3", f"Row {line} needs a 28-field PD3 detail record.")
    for index, choices in ALLOWED.items():
        require(row[index] in choices, f"Row {line} · Check column {column_label(index)}.")
    for index, pattern in PATTERNS.items():
        require(re.fullmatch(pattern, row[index]), f"Row {line} · Check column {column_label(index)}.")
    for index in ASSESSMENTS:
        require(row[index] in {"", "1", "2", "3"}, f"Row {line} · Check column {column_label(index)}.")
    for index in (Col.FIRST_NAME, Col.MIDDLE_NAME, Col.LAST_NAME):
        if row[index]:
            require(re.search("[A-Za-z]", row[index]), f"Row {line} · Check column {column_label(index)}.")


def given_name_warning(row):
    return any(re.search("[ '-]", row[index]) for index in (Col.FIRST_NAME, Col.MIDDLE_NAME))


def check_detail(row, line, sent, period, warnings):
    check_detail_fields(row, line)
    if given_name_warning(row):
        warnings.add(NAME_NOTICE)
    school = row[Col.HIGH_SCHOOL_NAME]
    try:
        registered_name(school)
    except ValidationError:
        raise ValidationError(f"Row {line} · Check the school name in N.") from None
    if len(school) > 50:
        warnings.add(SCHOOL_NOTICE)
    code = row[Col.PROGRAM_CODE]
    require(not code or valid_program_code(code) and code.casefold() != "n/a", f"Row {line} · Check program code AA.")
    birth = day(row[Col.BIRTH_DATE], f"birth date (row {line}, I)")
    award = day(row[Col.GRADUATION_DATE], f"diploma date (row {line}, L)")
    require(birth < award <= sent, f"Row {line} · Check birth, diploma and transmission date order.")
    year = row[Col.GRADUATION_DATE][:4]
    require(period[0] <= year <= period[-1], f"Row {line} · Diploma date falls outside the header period.")


def check_saved_file(path: Path) -> set[str]:
    require(path.suffix.lower() == ".txt", "Choose a graduates .txt file.")
    warnings = set()
    with path.open(encoding="utf-8", newline=None) as handle:
        sent, period = header_values(handle.readline().rstrip("\n").split("\t"))
        trailer_seen = False
        for line, raw in enumerate(handle, 2):
            require(not trailer_seen, f"Row {line} follows the PT3 trailer. The trailer must be last.")
            row = raw.rstrip("\n").split("\t")
            if row[0] == "PT3":
                require(line >= 3, "The file needs at least one PD3 detail record.")
                require(row == ["PT3", str(line)], f"Row {line} · PT3 must count every row, including PH3 and PT3.")
                trailer_seen = True
            else:
                check_detail(row, line, sent, period, warnings)
                require(line <= 250001, "The file exceeds the limit of 250,000 graduates.")
        require(trailer_seen, "The file is missing its final PT3 trailer.")
    return warnings


def show_file_check(path: Path) -> None:
    try:
        warnings = check_saved_file(path)
    except UnicodeError:
        raise ValidationError("Check the encoding of the saved file. Choose the original UTF-8 text file.") from None
    for notice in sorted(warnings):
        print("Warning · " + notice)
    print("Saved file layout checked." if warnings else "Saved file layout checked. No field warnings.")
