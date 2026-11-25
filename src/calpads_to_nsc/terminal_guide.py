import hashlib
import json
import os
import shlex
import shutil
import subprocess
import sys
import tempfile
import textwrap
import webbrowser
from pathlib import Path

from . import NON_AFFILIATION
from .demographics import DEMOGRAPHIC_LABELS
from .reports import completion_year_label, demographic_headers, folder_inputs, report_files, reporting_year
from .workbook_review import open_workbook, terminal_style
from .generate import ValidationError
from .saved_file import show_file_check
from .native_picker import DialogUnavailable, pick_path
from .terminal_input import FINISHED_AFTER_SAVE, PAUSED, RETURN_TO_MENU, Finished, answer


def preferences_path():
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    elif sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
        if not base.is_absolute():
            base = Path.home() / ".config"
    return base / "calpads-to-nsc" / "locations.json"


STUDENT_IDS = {
    "local": "District local student ID",
    "ssid": "California state student ID (SSID)",
    "blank": "Leave column J blank",
}
PREFERENCES = preferences_path()
CLEARINGHOUSE = "https://www.studentclearinghouse.org/"
SECURE_FTP = "https://ftps.nslc.org/"
INPUT_LABELS = {
    "reports": "Certified CALPADS 1.23 reports",
    "reference": "Clearinghouse Account Reference Sheet",
    "review": "School review workbook",
    "dates": "District supplement CSV (dates, FERPA blocks and optional program fields)",
}
EXTENSIONS = {"reference": (".xls", ".xlsx", ".pdf"), "review": (".xlsx",), "dates": (".csv",)}
SOURCE_STEPS = {
    "reports": "Ask for each high school’s certified 1.23 Graduates and Completers student-list export. "
    "Put the exports together with the reference sheet and optional supplement in one input folder.",
    "reference": "Ask your StudentTracker contact for the current Account Reference Sheet. "
    "Use the original Excel file or text-based PDF.",
}


class SelectionCancelled(Exception):
    pass


def load_locations():
    try:
        data = json.loads(PREFERENCES.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}
    if not isinstance(data, dict):
        return {}
    selected = {
        key: value
        for key, value in data.items()
        if key in (*INPUT_LABELS, "input_folder") and isinstance(value, str) and Path(value).is_absolute()
    }
    if "reports" in selected and isinstance(data.get("student_id"), str) and data["student_id"] in STUDENT_IDS:
        selected["student_id"] = data["student_id"]
    return selected


def save_locations(locations):
    PREFERENCES.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile("w", dir=PREFERENCES.parent, encoding="utf-8", delete=False)
    temporary = Path(handle.name)
    try:
        with handle:
            json.dump(locations, handle, indent=2)
        os.replace(temporary, PREFERENCES)
    finally:
        temporary.unlink(missing_ok=True)


def heading(title):
    print("\n" + terminal_style(title, "1;36", sys.stdout))


def choose(title, options, detail="", *, default=1):
    heading(title)
    if detail:
        print(detail)
    for number, option in enumerate(options, 1):
        key = terminal_style(str(number), "1;33", sys.stdout)
        print(f"  {key}  {option}")
    print("  " + terminal_style("Q", "1;33", sys.stdout) + "  Exit guide")
    while True:
        value = answer(f"Press Return for {default}, or enter your choice > ") or str(default)
        if value in {str(i) for i in range(1, len(options) + 1)}:
            return int(value)
        print(f"Enter a number from 1 to {len(options)}, or Q to finish.")


def unused_name(path):
    candidate = path
    number = 1
    while os.path.lexists(candidate):
        number += 1
        candidate = path.with_name(f"{path.stem}_{number}{path.suffix}")
    return candidate


def unquote_path(value):
    if value.startswith(('"', "'")) or (os.name != "nt" and "\\" in value):
        try:
            parts = shlex.split(value, posix=os.name != "nt")
            if len(parts) == 1:
                value = parts[0].strip('"') if os.name == "nt" else parts[0]
        except ValueError:
            pass
    return value


def save_as(title, default, extensions):
    heading(title)
    print(f"Press Return to save {default.name} in {default.parent}.\nB  Choose a name and location in the Save dialog")
    while True:
        value = answer("Filename (Q to exit) > ")
        if value.casefold() == "b":
            picked = pick_path(title, default.parent, extensions, save_name=default.name)
            if picked is None:
                continue
            value = str(picked)
        value = unquote_path(value)
        try:
            path = Path(value or default).expanduser()
        except RuntimeError:
            print("That home folder could not be found. Enter the full path.")
            continue
        if not path.is_absolute():
            path = Path(default).parent / path
        path = path.absolute()
        if not path.suffix:
            path = path.with_suffix(extensions[0])
        if path.suffix.lower() not in extensions:
            print("Choose a file ending in " + " or ".join(extensions) + ".")
        elif os.path.lexists(path):
            print("That name is in use. Enter a new filename.")
        elif not path.parent.is_dir():
            print("The destination folder could not be found. Press B to choose a location.")
        else:
            return str(path)


def selected_reports(path, locations):
    if locations.get("input_folder") != str(path):
        return report_files(path)
    return folder_inputs(
        path,
        reference=Path(locations["reference"]) if locations.get("reference") else None,
        review=Path(locations["review"]) if locations.get("review") else None,
        supplement=Path(locations["dates"]) if locations.get("dates") else None,
    )["reports"]


def files_to_check(key, path, locations):
    if key != "reports":
        if not path.is_file():
            raise ValidationError("File not found.")
        if path.suffix.lower() not in EXTENSIONS[key]:
            raise ValidationError("Choose a file ending in " + " or ".join(EXTENSIONS[key]) + ".")
        return [path]
    if not path.is_dir():
        raise ValidationError("Folder not found.")
    files = selected_reports(path, locations)
    if not files:
        raise ValidationError("No school reports found. Put the CALPADS exports directly in the input folder.")
    report_paths = {file.resolve() for file in files}
    for extra in EXTENSIONS:
        saved = locations.get(extra)
        if saved and Path(saved).resolve() in report_paths:
            raise ValidationError(f"Move {Path(saved).name} out of the CALPADS folder, then check again.")
    return files


def input_problem(key, location, locations=None):
    if not location:
        return "No location saved."
    try:
        files = files_to_check(key, Path(location), locations or {})
        for file in files:
            if not os.access(file, os.R_OK):
                return f"Cannot read {file.name}. Check file access."
            if file.stat().st_size == 0:
                return f"{file.name} is empty. Save a complete copy."
    except ValidationError as exc:
        return str(exc)
    except OSError:
        return "Cannot access this location. Check that it is available on this computer."
    return ""


def show_inputs(locations, problems):
    heading("Input check")
    if locations.get("input_folder"):
        print("Input folder · " + locations["input_folder"])
    width = max(40, min(96, shutil.get_terminal_size().columns))
    for key, label in INPUT_LABELS.items():
        issue = problems[key]
        location = locations.get(key)
        statuses = {"dates": ("OPTIONAL", "36"), "review": ("SET UP", "1;33")}
        status, color = statuses.get(key, ("NEEDED", "1;33")) if issue else ("FOUND", "1;32")
        print("  " + terminal_style(f"{status:<9}", color, sys.stdout) + label)
        if not issue and key != "reports":
            print("           " + Path(location).name)
        if issue:
            if location:
                print("           " + issue)
                print("           " + location)
            if key in SOURCE_STEPS:
                step = SOURCE_STEPS[key]
            elif key == "dates":
                step = (
                    "Use a CSV from district records to add diploma dates, FERPA blocks or program fields. "
                    "You can select it when making a file."
                )
            else:
                step = "After the two required inputs are found, create this workbook or locate your saved copy."
            print(textwrap.fill(step, width, initial_indent="           ", subsequent_indent="           "))


def browse(key, locations, previous=None):
    folder = Path(locations.get("reports", Path.cwd()))
    if previous:
        prior = Path(previous)
        candidate = prior if prior.is_dir() else prior.parent
        if candidate.is_dir() and "input_folder" not in locations:
            folder = candidate
    print(
        "Choose the input folder in the folder dialog."
        if key == "reports"
        else f"Choose the {INPUT_LABELS[key]} in the Open dialog.",
        flush=True,
    )
    location = pick_path(
        "Choose the input folder" if key == "reports" else INPUT_LABELS[key],
        folder,
        EXTENSIONS.get(key, ()),
        choose_folder=key == "reports",
    )
    if location is None:
        raise SelectionCancelled()
    return str(location)


def locate_input(key, previous=None, locations=None):
    while True:
        location = browse(key, locations or {}, previous)
        problem = input_problem(key, location, locations)
        if not problem:
            return location
        print(problem)
        previous = location


def launch_input_folder():
    cwd = Path.cwd()
    if cwd.name.casefold() == "input":
        return cwd
    if (cwd / "input").is_dir():
        return cwd / "input"
    return None


def output_folder(locations):
    reports = Path(locations["reports"])
    workspace = reports.parent if reports.name.casefold() == "input" or "input_folder" not in locations else reports
    return workspace / "outputs"


def set_detected_file(selected, key, candidates, saved, reselect):
    if key == "dates" and saved and not reselect and Path(saved).is_file():
        selected[key] = saved
    elif len(candidates) == 1:
        selected[key] = str(candidates[0])
    elif len(candidates) > 1:
        if key == "dates":
            print(
                "Several district supplements are available. When making a file, "
                "choose one in the file dialog or continue without a supplement."
            )
        elif saved and Path(saved) in candidates and not reselect:
            selected[key] = saved
        else:
            print("Several files match. Select the one for this request.")
            selected[key] = locate_input(key, locations=selected)
    elif saved and Path(saved).is_file():
        selected[key] = saved


def from_folder(folder, previous, *, reselect=False):
    found = folder_inputs(folder)
    same = str(folder) == previous.get("reports")
    selected = {"reports": str(folder)}
    if same and previous.get("student_id") in STUDENT_IDS:
        selected["student_id"] = previous["student_id"]
    is_input = folder.name.casefold() == "input" or (same and "input_folder" in previous)
    if found["reference"] or found["dates"] or is_input:
        selected["input_folder"] = str(folder)
    for key in ("reference", "dates"):
        set_detected_file(selected, key, found[key], previous.get(key) if same else None, reselect)
    review = previous.get("review")
    if review and same:
        selected["review"] = review
    else:
        default = output_folder(selected) / "School review.xlsx"
        if default.is_file():
            selected["review"] = str(default)
    return selected


def choose_workbook(locations):
    options = ["Create a new workbook", "Use an existing workbook"]
    saved = locations.get("review")
    if saved and not input_problem("review", saved):
        options.insert(0, "Keep " + Path(saved).name)
    mode = choose("School review workbook", options)
    if len(options) == 3:
        if mode == 1:
            return locations, ["--crosswalk", saved, "--check-crosswalk"]
        mode -= 1
    if mode == 1:
        reports = Path(locations["reports"])
        output_dir = output_folder(locations)
        output_dir.mkdir(mode=0o700, exist_ok=True)
        default = Path(locations.get("review", output_dir / "School review.xlsx"))
        if default.suffix.lower() != ".xlsx" or not default.parent.is_dir() or default.parent == reports:
            default = output_dir / "School review.xlsx"
        review = save_as("Where should the new workbook be saved?", unused_name(default), (".xlsx",))
        review_args = ["--create-crosswalk", review]
    else:
        review = locate_input("review", locations.get("review"), locations)
        review_args = ["--crosswalk", review, "--check-crosswalk"]
    return {**locations, "review": review}, review_args


def run_converter(arguments):
    debug = ["--debug"] if "--debug" in sys.argv[1:] else []
    return subprocess.run([sys.executable, "-m", "calpads_to_nsc", "--guide-session", *debug, *arguments]).returncode


def change_student_id(locations):
    value = choose(
        "What should appear in student ID column J?",
        list(STUDENT_IDS.values()),
        "This optional identifier helps your district match returned results. "
        "The guide always uses SSID to look up supplement records.",
    )
    locations["student_id"] = list(STUDENT_IDS)[value - 1]
    save_locations(locations)
    print("Column J · " + STUDENT_IDS[locations["student_id"]])


def choose_demographics(available):
    if not available:
        print("No demographic columns were found in the CALPADS reports. Those NSC fields will be blank.")
        return []
    labels = ", ".join(label for key, label in DEMOGRAPHIC_LABELS.items() if key in available)
    choice = choose(
        "Include demographics from CALPADS?",
        ["Include available demographics", "Exclude demographics for this file"],
        f"Available in the report headings · {labels}\nMissing or unsupported values stay blank.",
    )
    if choice == 2:
        return []
    flags = ["--include-demographics"]
    if "disability" in available and choose_disability():
        flags.append("--use-special-education-for-disability")
    return flags


def choose_disability():
    choice = choose(
        "Use CALPADS Special Education Status for NSC disability?",
        ["Include disability using Special Education Status", "Proceed without disability for this file"],
        "Column Z uses the report’s special education participation indicator. "
        "Choose 2 if your district uses a different measure or wants to leave this field blank.",
    )
    return choice == 1


def show_supplement_choices(locations, saved_path, saved, review_only):
    heading("How should the saved district supplement be used?" if saved else "Add a district supplement?")
    if saved:
        print("Saved file  " + Path(saved_path).name)
    print(
        "Common CSV headers   SSID,GraduationDate,FERPABlock\n"
        "A dates-only CSV also works. Optional columns can supply assessments, math semesters,\n"
        "dual enrollment and a program code. Empty program fields stay blank.\n\n"
        "Use the diploma date matched by SSID when available, or keep the CALPADS exit date.\n"
        "The guide will show you the source rows if dates conflict or a diploma date is outside this cycle.\n"
        "For a FERPA block, enter Y. Use N for no block or leave it blank if unknown.\n"
        "NSC accepts N for no block or unknown status."
    )
    default = "Use the available fields in this CSV" if saved else "Continue without a supplement"
    print("\n" + terminal_style("Return", "1;36", sys.stdout) + "   " + default)
    if saved:
        print(
            "B        Use only FERPA blocks from this CSV and keep CALPADS exit dates\n"
            "C        Choose a different supplement CSV\n"
            "S        Skip the supplement, including its dates, blocks and program fields, for this file"
        )
    else:
        print("Y        Choose a supplement CSV")
    if not review_only:
        print("I        Change the student ID in column J · " + STUDENT_IDS[locations.get("student_id", "local")])
    print("Q        Exit guide")


def choose_supplement(locations, *, review_only=False):
    saved_path = locations.get("dates")
    saved = not input_problem("dates", saved_path)
    show_supplement_choices(locations, saved_path, saved, review_only)
    while True:
        choice = answer("> ").casefold()
        if choice == "i" and not review_only:
            change_student_id(locations)
            show_supplement_choices(locations, saved_path, saved, review_only)
            continue
        if not choice:
            return (saved_path if saved else None), False
        if saved and choice == "b":
            return saved_path, True
        if (not saved and choice == "y") or (saved and choice == "c"):
            try:
                return locate_input("dates", saved_path, locations), False
            except SelectionCancelled:
                print("Selection cancelled. Press Return for the choice shown above, or choose another option.")
                continue
        if saved and choice == "s":
            return None, False
        print("Press Return or enter one of the letters shown.")


def source_arguments(locations):
    args = [locations["reports"], "--reference-sheet", locations["reference"], "--terminal-review"]
    if locations.get("input_folder"):
        args.append("--input-folder")
    return args


def show_test_upload(path):
    heading("Submit your Graduate test file")
    print(
        f"1. Go to {CLEARINGHOUSE} and click User Login.\n"
        "   Sign in with your StudentTracker Web ID and password.\n"
        "2. On the StudentTracker home page, click Submit a Test File. Choose Graduate.\n"
        "3. Click Browse and select this text file:\n"
        f"   {path}\n"
        "4. Click Upload. Read the results on the Error Resolution screen.\n"
        "   Review any warnings. For errors, follow the Proposed Solution, correct the inputs,\n"
        "   and generate a fresh test file. When the test passes, use menu option 3 to prepare the full file.\n\n"
        "NSC instructions · https://help.studentclearinghouse.org/sths/knowledge-base/submitting-a-test-file/"
    )
    offer_website(CLEARINGHOUSE, "Clearinghouse website")


def show_full_upload(path):
    heading("Submit your full graduates file")
    print(
        f"1. Open {SECURE_FTP} and sign in with your district’s secure FTP account.\n"
        "   You can also reach it through FTP Login on the Clearinghouse website.\n"
        "2. On the FTP home page choose Upload, then Browse. If you see Choose File, click it.\n"
        "3. Select this full graduates file and follow the Upload prompt in your browser.\n"
        f"   {path}\n"
        "4. Keep the FTP confirmation email as your transfer receipt. NSC validates the file next.\n"
        "   Watch the district technical contact’s email for any corrections.\n\n"
        "NSC instructions · https://help.studentclearinghouse.org/dataexchange/knowledge-base/"
        "uploading-files-via-the-web-browser-interface/"
    )
    offer_website(SECURE_FTP, "secure FTP login")


def offer_website(url, label):
    while True:
        choice = answer(f"\nReturn  Finish and exit (or Q)\nO       Open the {label}\nM       Main menu\n> ").casefold()
        if not choice:
            raise Finished()
        if choice == "m":
            return
        if choice == "o":
            try:
                if webbrowser.open(url):
                    continue
            except (OSError, webbrowser.Error):
                pass
            print(f"Open your browser and go to {url}")
        else:
            print("Press Return to finish, O to open the website, or M for the menu.")


def file_digest(path):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
    return digest.digest()


def review_file_intro(path, sample, title, app):
    heading(f"Review the {title} in {app}")
    print(path)
    if sample:
        print("The file contains one graduate per included campus.")
    else:
        print(
            "This file includes eligible graduates from the selected schools. "
            "Review NSC test results before submitting."
        )
    print(
        "Check the student ID (J), diploma date (L), FERPA block (M), school name (N) and ACT code (O).\n"
        "If you included demographics, check gender (P), race/ethnicity (Q), economic disadvantage (R),\n"
        "English learner (W) and disability (Z). Supplemental academic and program fields use S–V, X, Y and AA.\n"
        "If asked how to import the file, choose Tab as the delimiter and Text for the columns\n"
        "to preserve leading zeros. Close the file without saving changes."
    )


def count_graduates(path):
    with path.open("rb") as handle:
        total = sum(line.startswith(b"PD3\t") for line in handle)
    print(f"Graduates included in this file · {total}")


def ready_for_upload(path, original, retry):
    if file_digest(path) != original:
        print(f"The text file changed during review. {retry}")
        return False
    try:
        show_file_check(path)
    except ValidationError as exc:
        print(f"{exc} {retry}")
        return False
    return True


def review_saved_file(path, *, sample):
    title = "Graduate test file" if sample else "full graduates file"
    retry = "Choose 2 to create a fresh test file." if sample else "Choose 3 to create a fresh full file."
    app = "LibreOffice Calc" if sys.platform.startswith("linux") else "Excel"
    try:
        review_file_intro(path, sample, title, app)
        original = file_digest(path)
        while True:
            choice = answer(
                f"\nReturn  Open this file in {app}\n"
                "C       I reviewed the file. Show me how to submit it.\n"
                "T       Show the number of graduates in this file\n"
                "M       Main menu\nQ       Exit guide\n> "
            ).casefold()
            if not choice or choice == "x":
                if open_workbook(path, excel_only=True):
                    print("After reviewing the file, return to Terminal and enter C.")
            elif choice == "t":
                count_graduates(path)
            elif choice == "c":
                if not ready_for_upload(path, original, retry):
                    break
                show_upload = show_test_upload if sample else show_full_upload
                show_upload(path)
                return 0
            elif choice == "m":
                return 0
            else:
                print("Press Return or enter C, T, M or Q.")
    except OSError:
        print(f"The saved file could not be read at {path}. {retry}")
    except (Finished, EOFError, KeyboardInterrupt):
        return FINISHED_AFTER_SAVE
    return 0


def choose_saved_full_file(locations):
    folder = output_folder(locations)
    while True:
        path = pick_path("Select the Graduates_ text file saved by option 3", folder, (".txt",))
        if path is None:
            return 0
        path = path.resolve()
        folder = path.parent
        if path.suffix.lower() != ".txt" or not path.name.lower().startswith("graduates_"):
            print("Choose the full Graduates_ file created by option 3. Test files begin with TEST_.")
            continue
        try:
            with path.open("rb") as handle:
                if handle.read(4) != b"PH3\t":
                    print("This file has no NSC graduates header. Choose the original text file from option 3.")
                    continue
        except OSError:
            print("The file could not be opened. Select an accessible copy.")
            continue
        return review_saved_file(path, sample=False)


def check_selected_file(locations):
    path = pick_path("Select a saved NSC graduates text file", output_folder(locations), (".txt",))
    if path is not None:
        try:
            show_file_check(path)
        except (OSError, ValidationError) as exc:
            print(str(exc) if isinstance(exc, ValidationError) else "The selected file could not be opened.")


def generation_command(locations, sample, review_only, available):
    csv_path, use_calpads_dates = choose_supplement(locations, review_only=review_only)
    if csv_path and csv_path != locations.get("dates"):
        locations["dates"] = csv_path
        save_locations(locations)
    arguments = ["--supplement", csv_path] if csv_path else []
    if use_calpads_dates:
        arguments.append("--calpads-exit-dates")
    output_dir = output_folder(locations)
    if review_only:
        arguments += ["--dry-run", "--review-decisions"]
    else:
        arguments.extend(choose_demographics(available))
        arguments.append("--review-test" if sample else "--review-full")
        arguments += ["--test-output" if sample else "--output", output_dir]
    return [*source_arguments(locations), "--crosswalk", locations["review"], *arguments]


def retry_generation(locations, review_only):
    print("No graduates file was saved.", file=sys.stderr)
    heading("Fix the issue above, then try again")
    print("R  Retry with the same inputs")
    if not review_only:
        print("I  Change column J to a local ID, SSID or blank and retry")
    print("M  Main menu\nQ  Exit guide")
    while True:
        choice = answer("> ").casefold()
        if choice == "i" and not review_only:
            change_student_id(locations)
            return True
        if choice == "r":
            return True
        if choice == "m":
            return False
        print("Choose one of the letters shown.")


def generate(locations, sample, available, *, review_only=False):
    command = generation_command(locations, sample, review_only, available)
    output_dir = output_folder(locations)
    while True:
        try:
            output_dir.mkdir(mode=0o700, exist_ok=True)
        except OSError:
            print(f"Could not create the output folder at {output_dir}.", file=sys.stderr)
        else:
            id_args = [] if review_only else ["--student-id", locations.get("student_id", "local")]
            code = run_converter([*command, *id_args])
            if code == FINISHED_AFTER_SAVE:
                raise Finished()
            if code in (0, RETURN_TO_MENU):
                return 0
            if code == PAUSED:
                return code
        if not retry_generation(locations, review_only):
            return 0


def initial_locations():
    locations = load_locations()
    folder = launch_input_folder()
    if folder is None and locations.get("reports") and Path(locations["reports"]) == Path.cwd():
        folder = Path.cwd()
    if folder is None and locations.get("input_folder") and Path(locations["input_folder"]).is_dir():
        folder = Path(locations["input_folder"])
    if folder:
        selected = from_folder(folder, locations)
        if selected != locations:
            save_locations(selected)
        return selected
    return locations


def complete_inputs(locations, missing):
    print("\nSelect the missing input in the dialog that opens.")
    updated = locations.copy()
    for key in missing:
        if key == "reports":
            chosen = browse(key, updated, updated.get(key))
            proposed = from_folder(Path(chosen), updated)
            if proposed.get("input_folder") or updated.get("reports") != chosen:
                updated = proposed
            if input_problem(key, chosen, updated):
                break
        elif input_problem(key, updated.get(key), updated):
            updated[key] = locate_input(key, updated.get(key), updated)
            if key == "reference" and Path(updated[key]).parent == Path(updated["reports"]):
                updated["input_folder"] = updated["reports"]
        save_locations(updated)
    return updated


def file_version(path):
    info = path.stat()
    return path, info.st_ino, info.st_size, info.st_mtime_ns


def report_summary(files):
    try:
        year = reporting_year(files)
        cycle = completion_year_label(year)
        return cycle, demographic_headers(files), ""
    except ValidationError as exc:
        return "", set(), str(exc)


def other_folder(locations):
    folder = Path(browse("reports", locations, locations["reports"]))
    return from_folder(folder, locations, reselect=True)


def cycle_recovery(error, locations):
    heading("Check the CALPADS reports")
    print(error)
    choice = answer("R  Retry after correcting the files   C  Choose another input folder   Q  Finish > ").casefold()
    if choice == "c":
        try:
            locations = other_folder(locations)
            save_locations(locations)
        except SelectionCancelled:
            pass
    return locations


def change_inputs(locations):
    try:
        selected = other_folder(locations)
        if not selected.get("reference"):
            selected["reference"] = locate_input("reference", locations.get("reference"), selected)
            if Path(selected["reference"]).parent == Path(selected["reports"]):
                selected["input_folder"] = selected["reports"]
    except SelectionCancelled:
        return None
    problem = input_problem("reports", selected["reports"], selected)
    if problem:
        print(problem)
        return None
    return selected


def run_menu(locations, needs_workbook, cycle, available):
    selected = locations
    if needs_workbook:
        print(cycle)
        action = 4
    else:
        action = choose(
            "What would you like to do?",
            [
                "Review school matches",
                "Make a Graduate test file",
                "Make the full graduates file",
                "Change input files",
                "Review and submit a saved full file",
                "Review graduation decisions",
                "Check a saved NSC graduates file",
            ],
            f"{cycle}\nWorkbook · {Path(locations['review']).name}\nOutputs · {output_folder(locations)}",
        )
        if action == 4:
            selected = change_inputs(locations)
            if selected is None:
                return locations, 0
    if action == 4:
        selected, arguments = choose_workbook(selected)
        code = run_converter([*source_arguments(selected), *arguments])
        if code in (0, PAUSED):
            save_locations(selected)
            return selected, code
    elif action == 1:
        code = run_converter([*source_arguments(locations), "--crosswalk", locations["review"], "--check-crosswalk"])
    elif action == 7:
        check_selected_file(locations)
        code = 0
    elif action == 5:
        code = choose_saved_full_file(locations)
        if code == FINISHED_AFTER_SAVE:
            raise Finished()
    else:
        code = generate(locations, sample=action == 2, available=available, review_only=action == 6)
    return locations, code


def guide_session():
    locations = initial_locations()
    last_check = last_reports = None
    while True:
        problems = {key: input_problem(key, locations.get(key), locations) for key in INPUT_LABELS}
        check = {key: locations.get(key) for key in (*INPUT_LABELS, "input_folder")}, problems
        if check != last_check:
            show_inputs(locations, problems)
            last_check = check
        missing = [key for key in SOURCE_STEPS if problems[key]]
        if missing:
            locations = complete_inputs(locations, missing)
            continue
        files = selected_reports(Path(locations["reports"]), locations)
        version = [file_version(path) for path in files]
        if version != last_reports:
            cycle, available, error = report_summary(files)
            last_reports = version
        if error:
            locations = cycle_recovery(error, locations)
            last_reports = None
            continue
        locations, code = run_menu(locations, problems["review"], cycle, available)
        if code == PAUSED:
            return 0
        if code:
            print("\nCheck the message above. Choose your files again if a location has changed.")


def main():
    heading("Prepare an NSC graduates file")
    print(NON_AFFILIATION)
    print("At any Terminal prompt, type Q and press Return to exit. Use Cancel to leave a file dialog.")
    try:
        return guide_session()
    except (Finished, SelectionCancelled, EOFError, KeyboardInterrupt):
        print("\nGuide closed.")
        return 0
    except DialogUnavailable:
        print("The file dialog could not open. Run the guide from a desktop Terminal session.", file=sys.stderr)
        return 1
    except OSError:
        print("Check that your files and private working folder are accessible, then try again.", file=sys.stderr)
        return 1
