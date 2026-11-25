import os
import sys
from contextlib import closing

from openpyxl import load_workbook

from .workbook_review import terminal_style
from .crosswalk import EXCLUDE, FLAG, SchoolEdit, choices, decisions, suggestion, valid_reason, workbook_metadata
from .generate import ValidationError
from .storage import private_output_path, private_update_lock as commit_lock
from .terminal_input import answer


def heading(text):
    print("\n" + terminal_style(text, "1;36", sys.stdout))


def saved_matches(path, reference):
    book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        workbook_metadata(book)
        directory = choices(reference)
        return [
            (SchoolEdit("Match schools", f"D{row}", label), EXCLUDE if pick == EXCLUDE else directory.get(pick))
            for row, label, pick, _, _ in decisions(book)
        ]
    finally:
        book.close()


def pick_saved_match(path, reference):
    matches = [(edit, code) for edit, code in saved_matches(path, reference) if code]
    if not matches:
        print("There are no saved matches yet.")
        return None
    heading("Saved school matches")
    for number, (edit, code) in enumerate(matches, 1):
        summary = EXCLUDE if code == EXCLUDE else f"{reference['schools'][code]} (ACT {code})"
        print(f"  {number:>2}  {edit.label}\n      {summary}")
    while True:
        value = answer(terminal_style("Match number, Return to go back, or Q to exit > ", "1", sys.stdout))
        if not value:
            return None
        for number, (edit, _) in enumerate(matches, 1):
            if value == str(number):
                return edit
        print(f"Choose a number from 1 to {len(matches)}.")


def signature(info):
    return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns


def check_version(path, version):
    if path.with_name("~$" + path.name).exists():
        print("Save and close the open workbook, then start the school review again.")
        return "pause"
    if signature(path.stat()) != signature(version):
        print("The workbook changed during this review. Reading its latest saved choices.")
        return "retry"
    return None


def set_selection(book, edit, value, year, reason, confirmation):
    metadata = workbook_metadata(book)
    if value == EXCLUDE:
        if not valid_reason(reason) or confirmation != "EXCLUDE":
            raise ValidationError("Enter a school-level reason and type EXCLUDE before saving an exclusion.")
        if year is not None and metadata.get("Reporting cycle") != year:
            raise ValidationError("Start a fresh school review for this reporting year before recording an exclusion.")
    sheet = book[edit.sheet]
    cell = sheet[edit.cell]
    for col, note in ((6, reason), (7, confirmation)):
        note_cell = sheet.cell(cell.row, col)
        note_cell.value = note if value == EXCLUDE else None
        if note_cell.value is not None:
            note_cell.data_type = "s"
    cell.value = value
    cell.data_type = "s"


def commit_choice(path, edit, version, value, year, reason, confirmation):
    try:
        with commit_lock(path), private_output_path(path.parent, suffix=".xlsx") as temporary:
            status = check_version(path, version)
            if status:
                return status
            with path.open("rb") as source, closing(load_workbook(source, keep_links=False)) as book:
                set_selection(book, edit, value, year, reason, confirmation)
                book.save(temporary)
            status = check_version(path, version)
            if status:
                return status
            os.replace(temporary, path)
    except (BlockingIOError, FileExistsError):
        print("Another session is saving this workbook. Reading its latest choices.")
        return "retry"
    except PermissionError:
        print("Close the workbook and check that you can save files in its folder.")
        return "pause"
    return None


def save_choice(path, reference, edit, version, value, *, reason=None, confirmation=None, year=None):
    status = commit_choice(path, edit, version, value, year, reason, confirmation)
    if status:
        return status
    if value == FLAG:
        print("Flag saved. Check the city and ACT code against the account reference sheet.")
        return "pause"
    if value == EXCLUDE:
        print(terminal_style(f"Exclusion saved for {edit.label}.", "1;33", sys.stdout))
        return "saved"
    saved = reference["schools"][choices(reference)[value]]
    print(terminal_style(f"Saved  {edit.label}  →  {saved}", "1;32", sys.stdout))
    return "saved"


def school_menu(path, reference, edit):
    print("CALPADS school    " + terminal_style(edit.label, "1;33", sys.stdout))
    if edit.instruction:
        print(edit.instruction)
    directory = choices(reference)
    matches = saved_matches(path, reference)
    current = next((code for other, code in matches if other.cell == edit.cell), None)
    used = {code: other.label for other, code in matches if code and other.cell != edit.cell}
    if current == EXCLUDE:
        print("Current choice: exclude this campus. Choose a school below to include it again.")
    if not current:
        hint = suggestion(edit.label, reference)
        if hint in directory:
            print(f"Suggested school  {hint}. Check the campus and ACT code before selecting it.")
    ordered = sorted(directory, key=lambda name: (reference["schools"][directory[name]].casefold(), directory[name]))
    candidates = [name for name in ordered if directory[name] not in used or directory[name] == current]
    if candidates:
        print("These Clearinghouse schools are available.")
    else:
        print("All Clearinghouse schools have a saved match. Use M to check which campuses have those codes.")
    available = {code: name for name, code in directory.items()}
    for index, name in enumerate(candidates, 1):
        code = directory[name]
        status = "  - current match" if code == current else ""
        print(f"  {index:>2}  {reference['schools'][code]}  (ACT {code}){status}")
        available[str(index)] = name
    print("\nF  Flag for review and finish later\nE  Exclude from this submission")
    print("M  View or change a saved match     X  Open workbook     Q  Exit guide")
    return available, used


def exclusion_choice(path, reference, edit, version, year):
    print(f"All graduates from {edit.label} will be left out of the test and full files.")
    while True:
        reason = answer("School-level reason (no student details), or Return to go back > ")
        if not reason:
            return None
        if valid_reason(reason):
            break
        print("Use a single line up to 240 characters.")
    confirmed = answer("Type EXCLUDE to save this choice, or Return to go back > ")
    if confirmed != "EXCLUDE":
        print("No exclusion saved.")
        return None
    return save_choice(path, reference, edit, version, EXCLUDE, reason=reason, confirmation=confirmed, year=year)


def read_school_choice(available, used, reference):
    while True:
        value = answer(terminal_style("School number or six-digit ACT code > ", "1", sys.stdout))
        key = value.casefold()
        if key in ("x", "m", "e", "f"):
            return key
        name = available.get(value)
        if not name:
            print("Choose a number from the list or enter the six-digit ACT code shown beside the school.")
            continue
        code = choices(reference)[name]
        if code in used:
            print(f"ACT {code} is already saved for {used[code]}. Use M to check that match, or choose another school.")
            continue
        return name


def edit_school(path, reference, edit, version, *, year=None):
    while True:
        status = check_version(path, version)
        if status:
            return status
        available, used = school_menu(path, reference, edit)
        key = read_school_choice(available, used, reference)
        if key == "x":
            return "excel"
        if key == "m":
            edit = pick_saved_match(path, reference) or edit
        elif key == "e":
            status = exclusion_choice(path, reference, edit, version, year)
            if status:
                return status
        else:
            return save_choice(path, reference, edit, version, FLAG if key == "f" else key, year=year)
