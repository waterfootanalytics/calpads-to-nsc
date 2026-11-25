import re
import unicodedata
from collections.abc import Mapping
from datetime import datetime
from enum import IntEnum

from .storage import private_output_path, publish_new


DIPLOMAS = {
    "Endorsed/Advanced Diploma",
    "General Education Development Credential (GED)",
    "High School Equivalency Credential (other than GED)",
    "International Baccalaureate",
    "Modified Diploma",
    "Regular Diploma",
    "Vocational Certificate",
    "Other",
}
NAME_EQUIVALENTS = str.maketrans(
    {
        **dict.fromkeys("‘’ʼʻ＇`", "'"),
        **dict.fromkeys("‐‑‒–—−﹣－", "-"),
        "Æ": "AE",
        "æ": "ae",
        "Œ": "OE",
        "œ": "oe",
        "Ø": "O",
        "ø": "o",
        "Ł": "L",
        "ł": "l",
        "Đ": "D",
        "đ": "d",
        "Ð": "D",
        "ð": "d",
        "Þ": "Th",
        "þ": "th",
        "ẞ": "SS",
        "ß": "ss",
        "ı": "i",
    }
)


class ValidationError(ValueError):
    pass


def require(condition, message):
    if not condition:
        raise ValidationError(message)


def field(value, limit, label):
    require(
        isinstance(value, str) and 0 < len(value) <= limit and value.isprintable() and value == value.strip(),
        "Invalid " + label + ".",
    )
    return value


def day(value, label):
    require(re.fullmatch(r"[0-9]{8}", value), "Invalid " + label + ".")
    try:
        return datetime.strptime(value, "%Y%m%d").date()
    except ValueError:
        raise ValidationError("Invalid " + label + ".") from None


def name(value, label, notices=None):
    if notices is None:
        notices = set()
    require(
        "\ufffd" not in value, f"The {label} contains an unreadable character. Check the export encoding or spelling."
    )
    token = value.strip().upper()
    require(
        token not in {"#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"},
        f"The {label} contains a spreadsheet error. Supply the verified spelling.",
    )
    text = "" if token in {"N/A", "#N/A"} else value.translate(NAME_EQUIVALENTS)
    kept = "".join(c for c in text if c.isalpha() or c.isspace() or c in "'-")
    folded = unicodedata.normalize("NFKD", kept).translate(NAME_EQUIVALENTS)
    require(
        not any(c.isalpha() and not c.isascii() for c in folded),
        f"The {label} includes letters the formatter cannot convert. Supply the verified Latin spelling.",
    )
    cleaned = " ".join(re.sub(r"[^A-Za-z\s'-]", "", folded).split())
    if len(cleaned) > 40:
        cleaned = cleaned[:40].rstrip()
        notices.add("The 40-character limit shortened a student name field in the NSC output.")
    if not re.search("[A-Za-z]", cleaned):
        require(label == "middle name", f"The {label} has no letters after formatting. Supply the verified legal name.")
        cleaned = ""
    if cleaned != value:
        notices.add("Student names were formatted automatically for the NSC file.")
    return cleaned


class DetailColumn(IntEnum):
    RECORD_TYPE = 0
    SOCIAL_SECURITY_NUMBER = 1
    FIRST_NAME = 2
    MIDDLE_NAME = 3
    LAST_NAME = 4
    BIRTH_DATE = 8
    STUDENT_ID = 9
    DIPLOMA_TYPE = 10
    GRADUATION_DATE = 11
    FERPA_BLOCK = 12
    HIGH_SCHOOL_NAME = 13
    ACT_CODE = 14
    GENDER = 15
    ETHNICITY = 16
    ECONOMIC_DISADVANTAGE = 17
    GRADE8_MATH_ASSESSMENT = 18
    GRADE8_ELA_ASSESSMENT = 19
    HIGH_SCHOOL_MATH_ASSESSMENT = 20
    HIGH_SCHOOL_ELA_ASSESSMENT = 21
    ENGLISH_LEARNER = 22
    MATH_SEMESTERS = 23
    DUAL_ENROLLMENT = 24
    DISABILITY = 25
    PROGRAM_CODE = 26
    END_OF_DETAIL = 27


def detail_row(
    student: Mapping[str, str],
    key: str,
    award: Mapping[str, str],
    school: Mapping[str, str],
    name_notices: set[str] | None = None,
) -> list[str]:
    row = [""] * (DetailColumn.END_OF_DETAIL + 1)
    row[DetailColumn.RECORD_TYPE] = "PD3"
    row[DetailColumn.SOCIAL_SECURITY_NUMBER] = "NO SSN"
    row[DetailColumn.FIRST_NAME] = name(student["first_name"], "first name", name_notices)
    row[DetailColumn.MIDDLE_NAME] = name(student["middle_name"], "middle name", name_notices)
    row[DetailColumn.LAST_NAME] = name(student["last_name"], "last name", name_notices)
    row[DetailColumn.BIRTH_DATE] = student["birth_date"]
    row[DetailColumn.STUDENT_ID] = key
    row[DetailColumn.DIPLOMA_TYPE] = award["diploma_type"]
    row[DetailColumn.GRADUATION_DATE] = award["graduation_date"]
    row[DetailColumn.FERPA_BLOCK] = award["ferpa_block"]
    row[DetailColumn.HIGH_SCHOOL_NAME] = school["nsc_school_name"]
    row[DetailColumn.ACT_CODE] = school["act_code"]
    row[DetailColumn.END_OF_DETAIL] = "ED"
    return row


def submission(account, send_date, years, details):
    require(isinstance(account, dict), "Invalid account configuration.")
    number = field(account.get("account_number"), 10, "account number")
    require(re.fullmatch("[0-9]+", number), "Invalid account number.")
    organization = field(account.get("organization_name"), 80, "organization name")
    header = ["PH3", number, organization, "P", send_date, "-".join(sorted(years))]
    return [header, *details, ["PT3", str(len(details) + 2)]]


def save_submission(output, rows, *, name_prefix=None):
    auto_name = name_prefix is not None and output.is_dir()
    if auto_name:
        output = output / f"{name_prefix}_{datetime.now():%Y%m%d_%H%M%S}.txt"
    require(re.fullmatch(r"[A-Za-z0-9_-]+\.txt", output.name), "Invalid output filename.")
    with private_output_path(output.parent) as temp:
        with temp.open("w", encoding="utf-8", newline="") as handle:
            for row in rows:
                handle.write("\t".join(row) + "\r\n")
        candidate = output
        number = 1
        while True:
            try:
                publish_new(temp, candidate)
                return candidate
            except FileExistsError:
                if not auto_name:
                    raise
                number += 1
                candidate = output.with_name(f"{output.stem}_{number}{output.suffix}")
