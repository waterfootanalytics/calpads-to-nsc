import re
from contextlib import closing
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from .generate import ValidationError, require
from .storage import private_output_path, publish_new
from .reference_sheet import match_school


FLAG = "Flag for review"
EXCLUDE = "Exclude from this submission"
SELECTION_POLICY = "Explicit school selections"
FIRST_ROW = 13
MATCH_HEADERS = (
    "School in CALPADS",
    "Suggested match",
    "Clearinghouse school",
    "Status",
    "Reason for exclusion",
    "Type EXCLUDE to confirm",
)
UNSUPPORTED_WORKBOOK = "Unsupported school review workbook. Create a new school review workbook from your input files."


@dataclass(frozen=True)
class SchoolEdit:
    sheet: str
    cell: str
    label: str
    instruction: str = ""


class WorkbookEditsNeeded(ValidationError):
    def __init__(self, message, edits=(), excluded=()):
        super().__init__(message + ("\n\n" + exclusion_notice(excluded) if excluded else ""))
        self.edits = edits
        self.excluded = tuple(sorted(excluded))


def exclusion_notice(excluded):
    return "Excluded from this submission:\n" + "\n".join(f"  {school}" for school in sorted(excluded))


def choices(reference):
    return {f"{name} ({code})": code for code, name in reference["schools"].items()}


def tokens(label):
    return set(re.findall(r"[a-z]+", label.lower())) - {
        "high",
        "school",
        "hs",
        "the",
        "of",
        "and",
        "academy",
        "technology",
        "tech",
        "college",
        "ready",
        "complex",
    }


def exact_choice(label, reference):
    school = match_school(label, reference)
    if school["act_code"]:
        return f"{school['nsc_school_name']} ({school['act_code']})"
    return None


def comparable_names(label, reference):
    names = {choice: reference["schools"][code] for choice, code in choices(reference).items()}
    first_word = re.match(r"[a-z]+", reference.get("organization_name", "").strip().lower())
    if first_word:
        prefix = re.compile(rf"^{first_word[0]}\b\s*", re.IGNORECASE)
        if sum(bool(prefix.match(name)) for name in names.values()) >= 2:
            label = prefix.sub("", label)
            names = {choice: prefix.sub("", name) for choice, name in names.items()}
    grade_span = r"\b(?:k|[0-9]{1,2})\s*[-–—]\s*(?:8|12)\b"
    label = re.sub(grade_span, "", label, flags=re.IGNORECASE)
    names = {choice: re.sub(grade_span, "", name, flags=re.IGNORECASE) for choice, name in names.items()}
    numbers = re.findall(r"[0-9]+", label)
    names = {choice: name for choice, name in names.items() if re.findall(r"[0-9]+", name) == numbers}
    return label, names


def suggestion(label, reference):
    exact = exact_choice(label, reference)
    if exact:
        return exact
    label, names = comparable_names(label, reference)
    source = tokens(label)
    words = [
        w
        for w in re.findall(r"[a-z]+", label.lower())
        if w not in {"school", "high", "hs", "the", "of", "and", "complex"}
    ]
    initials = "".join(w[0] for w in words)
    acronym_matches = [choice for choice, name in names.items() if len(words) >= 3 and tokens(name) == {initials}]
    if len(acronym_matches) == 1:
        return acronym_matches[0]
    ranked = []
    for choice, name in names.items():
        target = tokens(name)
        overlap = source & target
        if overlap:
            ranked.append((len(overlap) / len(source | target), choice))
    ranked.sort(reverse=True)
    if ranked and ranked[0][0] >= 0.35 and (len(ranked) == 1 or ranked[0][0] - ranked[1][0] >= 0.15):
        return ranked[0][1]
    return "No clear match"


def dropdown(sheet, column, formula, first, last):
    rule = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True)
    rule.error = "Choose an entry from the list."
    sheet.add_data_validation(rule)
    rule.add(f"{column}{first}:{column}{last}")


def banner(sheet, row, text, *, title=False):
    sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    cell = sheet.cell(row, 2, text)
    cell.data_type = "s"
    cell.font = Font(name="Calibri", size=18 if title else 12, bold=title, color="FFFFFF" if title else "18384A")
    cell.fill = PatternFill("solid", fgColor="214A5F" if title else "EDF4F6")
    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    sheet.row_dimensions[row].height = 38 if title else 32


def create_review_sheets(book, reference, rehearsal):
    sheet = book.active
    sheet.title = "Match schools"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "C13"
    for column, width in {"A": 3, "B": 43, "C": 44, "D": 52, "E": 24, "F": 44, "G": 26}.items():
        sheet.column_dimensions[column].width = width
    banner(sheet, 2, "Match your high schools", title=True)
    banner(sheet, 3, reference["organization_name"] + ("  ·  Rehearsal only" if rehearsal else ""))
    banner(sheet, 5, "1. Compare the CALPADS campus in B with the account school and ACT code suggested in C.")
    banner(sheet, 6, "2. Keep correct exact matches in D. Use blank yellow cells to choose a school or flag it.")
    banner(sheet, 7, "3. For an uncertain match, follow Look up a school. Save this workbook and return to the guide.")
    banner(
        sheet,
        8,
        "To leave out a campus, choose Exclude from this submission. Enter a school-level reason in F "
        "and type EXCLUDE in G. Leave F and G blank for included schools.",
    )
    sheet.row_dimensions[8].height = 40
    for col, label in enumerate(MATCH_HEADERS, 2):
        cell = sheet.cell(12, col, label)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="214A5F")
        cell.alignment = Alignment(vertical="center", indent=1)
    sheet.row_dimensions[12].height = 26
    create_lookup_sheet(book, reference)
    return sheet


def create_lookup_sheet(book, reference):
    guide = book.create_sheet("Look up a school")
    guide.sheet_view.showGridLines = False
    for column, width in {"A": 3, "B": 58, "C": 16, "D": 10, "E": 25}.items():
        guide.column_dimensions[column].width = width
    banner(guide, 2, "Look up a school", title=True)
    banner(guide, 4, "1. Find the campus on the account sheet. Use its city and state to tell schools apart.")
    banner(guide, 5, "2. Check the six-digit ACT code in your district directory or ask the campus about former names.")
    banner(guide, 6, "3. Choose the school on Match schools. If missing, flag it and get a current account sheet.")
    guide.cell(9, 2, "Registered school name")
    guide.cell(9, 3, "ACT code")
    for col in (2, 3):
        guide.cell(9, col).font = Font(bold=True, color="FFFFFF")
        guide.cell(9, col).fill = PatternFill("solid", fgColor="214A5F")
    for row, (code, name) in enumerate(sorted(reference["schools"].items(), key=lambda pair: (pair[1], pair[0])), 10):
        guide.cell(row, 2, name).data_type = "s"
        guide.cell(row, 3, code).data_type = "s"
        guide.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="center")
        guide.row_dimensions[row].height = 32
    guide.freeze_panes = "B10"
    guide.auto_filter.ref = f"B9:C{guide.max_row}"


def valid_reason(value):
    return isinstance(value, str) and 0 < len(value.strip()) <= 240 and value.isprintable()


def workbook_metadata(book):
    sheets = {"Match schools", "Look up a school", "NSC choices", "Source"}
    require(sheets <= set(book.sheetnames) and "SchoolChoices" in book.defined_names, UNSUPPORTED_WORKBOOK)
    sheet = book["Match schools"]
    headers = next(sheet.iter_rows(min_row=12, max_row=12, min_col=2, max_col=7, values_only=True), ())
    require(headers == MATCH_HEADERS, UNSUPPORTED_WORKBOOK)
    meta = dict(book["Source"].iter_rows(max_col=2, values_only=True))
    require(meta.get(SELECTION_POLICY) == 1, UNSUPPORTED_WORKBOOK)
    require(meta.get("Use") in ("Rehearsal", "Submission"), UNSUPPORTED_WORKBOOK)
    return meta


def previous_choices(previous, labels, reference, digest, year):
    saved = {}
    with closing(load_workbook(previous, read_only=True, data_only=False, keep_links=False)) as old:
        meta = workbook_metadata(old)
        require(
            meta.get("Reference sheet fingerprint") == digest, "Account reference sheet changed. Start a fresh review."
        )
        directory = choices(reference)
        for n, label, pick, reason, confirmation in decisions(old):
            require(label in labels and label not in saved, f"Unexpected or duplicate school on row {n}.")
            require(pick in (None, "", FLAG, EXCLUDE) or pick in directory, f"Check the saved selection on row {n}.")
            if pick == EXCLUDE and year is not None:
                require(
                    meta.get("Reporting cycle") == year,
                    "A saved exclusion belongs to another cycle. Start a fresh school review for this year.",
                )
            saved[label] = pick, reason, confirmation
        require(saved.keys() == labels, "Report schools changed. Start a fresh review.")
    return saved, meta["Use"] == "Rehearsal"


def add_workbook_sources(book, reference, digest, rehearsal, directory, year):
    source = book.create_sheet("Source")
    source.append(["Reference sheet fingerprint", digest])
    source.append(["Use", "Rehearsal" if rehearsal else "Submission"])
    source.append(["NSC organization", reference["organization_name"]])
    source.append(["Agreement", reference["account_number"]])
    source["B3"].data_type = "s"
    source.append([SELECTION_POLICY, 1])
    if year is not None:
        source.append(["Reporting cycle", year])
    source.sheet_state = "hidden"
    listing = book.create_sheet("NSC choices")
    listing.append(["Registered school and ACT code"])
    listing.append([FLAG])
    for item in sorted(directory):
        listing.append([item])
        listing.cell(listing.max_row, 1).data_type = "s"
    listing.append([EXCLUDE])
    listing.column_dimensions["A"].width = 80
    listing.sheet_state = "hidden"
    book.defined_names.add(DefinedName("SchoolChoices", attr_text=f"'NSC choices'!$A$2:$A${listing.max_row}"))


def school_status_formula(row, last):
    pick = f"D{row}"
    return (
        f'=IF({pick}="{EXCLUDE}",IF(AND(LEN(TRIM(F{row}))>0,LEN(TRIM(F{row}))<=240,EXACT(G{row},"EXCLUDE")),'
        f'"Excluded","Complete F and G"),IF(OR(F{row}<>"",G{row}<>""),"Clear F and G",'
        f'IF({pick}="","Choose a school",IF({pick}="{FLAG}","Flagged",'
        f'IF(SUMPRODUCT(--(SchoolChoices={pick}))<>1,"Check choice",'
        f'IF(SUMPRODUCT(--($D${FIRST_ROW}:$D${last}={pick}))>1,"Used twice","Matched"))))))'
    )


def fill_school_rows(sheet, labels, reference, saved):
    last = FIRST_ROW + len(labels) - 1
    for row, label in enumerate(sorted(labels), FIRST_ROW):
        for col in range(2, 6):
            cell = sheet.cell(row, col)
            cell.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
            cell.font = Font(name="Calibri", size=12, color="18384A")
            if col == 4:
                cell.fill = PatternFill("solid", fgColor="FFF0BE")
        sheet.row_dimensions[row].height = 44
        sheet.cell(row, 2, label).data_type = "s"
        hint = suggestion(label, reference)
        sheet.cell(row, 3, hint).data_type = "s"
        pick_value, reason, confirmation = saved.get(label, (exact_choice(label, reference), None, None))
        if pick_value:
            sheet.cell(row, 4, pick_value).data_type = "s"
        sheet.cell(row, 5, school_status_formula(row, last))
        for col, value in ((6, reason), (7, confirmation)):
            cell = sheet.cell(row, col, value)
            cell.fill = PatternFill("solid", fgColor="FFF7E0")
            cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
            cell.number_format = "@"
            if value is not None:
                cell.data_type = "s"
    dropdown(sheet, "D", "SchoolChoices", FIRST_ROW, last)
    sheet["B10"] = (
        f'=COUNTIF(E{FIRST_ROW}:E{last},"Matched")&" matched  ·  "&COUNTIF(E{FIRST_ROW}:E{last},"Excluded")&" excluded"'
    )
    sheet.auto_filter.ref = f"B12:G{last}"


def create(path, labels, reference, digest, rehearsal=False, previous=None, year=None):
    require(path.suffix.lower() == ".xlsx", "Use .xlsx for the review workbook.")
    require(labels, "No school labels to review.")
    directory = choices(reference)
    require(directory, "No open high schools on the reference sheet.")
    saved, was_rehearsal = previous_choices(previous, labels, reference, digest, year) if previous else ({}, False)
    with closing(Workbook()) as book:
        book.properties.creator = "CALPADS to NSC"
        sheet = create_review_sheets(book, reference, rehearsal or was_rehearsal)
        add_workbook_sources(book, reference, digest, rehearsal or was_rehearsal, directory, year)
        fill_school_rows(sheet, labels, reference, saved)
        with private_output_path(path.parent, suffix=".xlsx") as temp:
            book.save(temp)
            publish_new(temp, path)


def decisions(book):
    rows = book["Match schools"].iter_rows(min_row=FIRST_ROW, min_col=2, max_col=7)
    for n, (label, _, pick, _, reason, confirmation) in enumerate(rows, FIRST_ROW):
        if any(cell.value is not None for cell in (label, pick, reason, confirmation)):
            require(
                reason.data_type != "f" and confirmation.data_type != "f",
                f"Match schools row {n}: type the reason and confirmation as plain text.",
            )
            yield n, label.value, pick.value, reason.value, confirmation.value


def review_metadata(book, labels, digest, writing, deferred, test_file):
    meta = workbook_metadata(book)
    require(
        meta.get("Reference sheet fingerprint") == digest,
        "Account reference sheet changed. Create a new workbook and review it.",
    )
    require(
        not writing or (meta["Use"] == "Submission" and not test_file),
        "A rehearsal workbook or test file cannot be used for production output.",
    )
    if deferred:
        require(meta["Use"] == "Rehearsal" or test_file, "Use --defer-school with a rehearsal workbook or a test file.")
        require(deferred <= labels, "Copy the full name for --defer-school from the School in CALPADS column.")
        require(labels - deferred, "Keep at least one school in the rehearsal.")
    return meta


@dataclass
class SchoolReview:
    names: dict
    directory: dict
    excluded: dict = field(default_factory=dict)
    schools: dict = field(default_factory=dict)
    used: dict = field(default_factory=dict)
    issues: list = field(default_factory=list)
    edits: list = field(default_factory=list)

    def issue(self, n, label, detail="", instruction=""):
        self.issues.append(f"  D{n}  {label}" + (" — " + detail if detail else ""))
        self.edits.append(SchoolEdit("Match schools", f"D{n}", label, instruction))

    def record(self, n, label, selected, reason, confirmation, deferred, correct_cycle):
        if selected == EXCLUDE:
            require(
                correct_cycle,
                "A saved exclusion is outside this workbook’s reporting cycle. "
                "Start a fresh school review for this year.",
            )
            if valid_reason(reason) and confirmation == "EXCLUDE":
                self.excluded[label] = reason.strip()
            else:
                self.issue(
                    n,
                    label,
                    f"enter a reason (1–240 characters) in F{n} and type EXCLUDE in G{n}.",
                    "Choose E to enter the reason and confirm the exclusion, or choose a school to include it.",
                )
        elif reason or confirmation:
            self.issue(
                n,
                label,
                f"clear F{n} and G{n} when including or flagging this school.",
                "Select the school again to clear its old exclusion details.",
            )
        elif label in deferred:
            require(
                selected in (None, "", FLAG), f"Review row {n} already has a choice. Check the school being deferred."
            )
        elif selected not in self.directory:
            self.issue(n, label, "Flagged (see Look up a school)" if selected == FLAG else "")
        else:
            code = self.directory[selected]
            if code in self.used:
                self.issue(n, label, f"same school as {self.used[code]}. Check both choices.")
            self.used.setdefault(code, f"D{n} ({label})")
            self.schools[label] = {"act_code": code, "nsc_school_name": self.names[code]}

    def check(self, rows, labels, deferred, correct_cycle):
        seen = {}
        for n, label, selected, reason, confirmation in rows:
            require(
                label in labels and label not in seen, f"Review row {n} has an unexpected or duplicate CALPADS school."
            )
            seen[label] = n
            self.record(n, label, selected, reason, confirmation, deferred, correct_cycle)
        require(seen.keys() == labels, "Report schools changed or a row was removed. Rebuild the crosswalk.")
        if not self.issues and not self.schools:
            for label in self.excluded:
                self.issue(
                    seen[label],
                    label,
                    "keep at least one school in this submission.",
                    "At least one school must be included. Choose its registered school.",
                )
        if self.issues:
            title = "Choose a school match" if len(self.issues) == 1 else f"Review {len(self.issues)} school matches"
            instructions = "On the Match schools tab, use the dropdown in each yellow cell."
            message = f"{title}\n{instructions}\n\n" + "\n".join(self.issues)
            raise WorkbookEditsNeeded(message, self.edits, self.excluded)


def reviewed_config(path, labels, reference, digest, writing, deferred=frozenset(), test_file=False, year=None):
    with closing(load_workbook(path, read_only=True, data_only=False, keep_links=False)) as book:
        meta = review_metadata(book, labels, digest, writing, deferred, test_file)
        review = SchoolReview(reference["schools"], choices(reference))
        review.check(decisions(book), labels, deferred, year is None or meta.get("Reporting cycle") == year)
    return {
        "account_number": reference["account_number"],
        "organization_name": reference["organization_name"],
        "schools": review.schools,
        "excluded_schools": review.excluded,
    }
