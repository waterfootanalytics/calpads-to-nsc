import re
from collections.abc import Mapping

from .generate import DetailColumn


DEMOGRAPHIC_LABELS = {
    "gender": "Gender",
    "ethnicity": "Race/ethnicity",
    "economic_disadvantage": "Economic disadvantage",
    "english_learner": "English learner",
    "disability": "Disability",
}


def _labels(groups: dict[str, tuple[str, ...]]) -> dict[str, str]:
    return {label.casefold(): code for code, labels in groups.items() for label in (code, *labels)}


GENDER = _labels({"F": ("Female",), "M": ("Male",)})
ETHNICITY = _labels(
    {
        "AM": (
            "Am Ind/Alaska Nat",
            "Am Indian/Alskn Nat",
            "American Indian/Alaska Native",
            "American Indian/Alaskan Native",
            "American Indian or Alaska Native",
            "American Indian or Alaskan Native",
            "American Indian or Native Alaskan",
        ),
        "AS": ("Asian", "Filipino"),
        "BL": ("Black", "African American", "Black/African Am", "Black/African American", "Black or African American"),
        "HI": ("Hispanic", "Hispanic/Latino", "Hispanic or Latino", "Hispanic or Latino of any race"),
        "PI": (
            "Nat Hawaii/Pac Isl",
            "Nat Hwiin/Other Pac Islndr",
            "Pacific Islander",
            "Hawaiian/Pacific Islander",
            "Native Hawaiian/Pacific Islander",
            "Native Hawaiian/Other Pacific Islander",
            "Native Hawaiian or Pacific Islander",
            "Native Hawaiian or Other Pacific Islander",
        ),
        "WH": ("White",),
        "MU": ("Multi", "Multiple", "Multiple Races", "Two or More", "Two or More Races"),
    }
)
YES_NO = _labels({"Y": ("Yes", "True"), "N": ("No", "False")})
ECONOMIC_DISADVANTAGE = {**YES_NO, "1": "Y", "0": "N"}
ENGLISH_LEARNER = {
    **YES_NO,
    **_labels(
        {
            "Y": ("EL", "English Learner"),
            "N": (
                "EO",
                "English Only",
                "IFEP",
                "Initial Fluent English Proficient",
                "Initially Fluent English Proficient",
                "RFEP",
                "Reclassified Fluent English Proficient",
            ),
        }
    ),
}
DISABILITY = {**YES_NO, "eligible and participating": "Y", "1": "Y"}
FIELDS = {
    "gender": (DetailColumn.GENDER, GENDER),
    "ethnicity": (DetailColumn.ETHNICITY, ETHNICITY),
    "economic_disadvantage": (DetailColumn.ECONOMIC_DISADVANTAGE, ECONOMIC_DISADVANTAGE),
    "english_learner": (DetailColumn.ENGLISH_LEARNER, ENGLISH_LEARNER),
    "disability": (DetailColumn.DISABILITY, DISABILITY),
}


def _convert(text: str, accepted: dict[str, str]) -> str:
    key = re.sub(r"\s*/\s*", "/", " ".join(text.casefold().split()))
    if key in accepted:
        return accepted[key]
    pair = re.fullmatch(r"(.+?)\s*(?:[-–—:]\s*(.+)|\(([^()]+)\))", key)
    if pair:
        first = accepted.get(pair[1].strip())
        second = accepted.get((pair[2] or pair[3]).strip())
        if first and first == second:
            return first
    return ""


def demographic_values(
    row: Mapping[str, object], *, include_disability: bool, notices: set[str]
) -> dict[DetailColumn, str]:
    fields = {}
    for key, (column, accepted) in FIELDS.items():
        if key == "disability" and not include_disability:
            continue
        label = DEMOGRAPHIC_LABELS[key]
        if key not in row:
            notices.add(f"CALPADS has a missing or repeated {label} heading. That report's NSC field was left blank.")
            continue
        value = row[key]
        if value is None or not str(value).strip():
            continue
        code = _convert(str(value), accepted)
        if code:
            fields[column] = code
        else:
            notices.add(f"Unsupported CALPADS entries for {label} were left blank.")
    return fields
