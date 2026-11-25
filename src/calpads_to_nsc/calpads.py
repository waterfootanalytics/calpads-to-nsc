import csv
import json
import sys
from pathlib import Path
from zipfile import BadZipFile

from .arguments import parse_arguments
from .comparison import compare_dates
from .decisions import ReviewDecisions
from .generate import ValidationError, day, require, save_submission
from .graduates import DEFAULT_DIPLOMAS, completion_overrides, convert, graduate_test_sample
from .reference_sheet import read_reference
from .reports import completion_year_label, folder_inputs, report_files, report_records, reporting_year
from .saved_file import show_file_check
from .supplement import Supplement, read_supplement
from .terminal_input import PAUSED, RETURN_TO_MENU, Finished, ReturnToMenu, answer
from .workbook_review import review_schools, terminal_style


def selected_reports(args):
    if args.input_folder:
        require(args.report.is_dir(), "Choose the input folder containing the CALPADS exports.")
        paths = folder_inputs(
            args.report, reference=args.reference_sheet, review=args.crosswalk, supplement=args.supplement
        )["reports"]
        require(paths, "No CALPADS exports found. Put the school reports directly in the input folder.")
        for target in (args.create_crosswalk, args.compare_dates):
            require(
                not target or target.resolve().parent != args.report.resolve(),
                "Save the generated file in the output folder or another folder outside the input files.",
            )
        return paths
    if args.report.is_dir() and not args.check_headers:
        destinations = (
            args.output,
            args.test_output,
            args.settings,
            args.create_crosswalk,
            args.crosswalk,
            args.reference_sheet,
            args.supplement,
            args.compare_dates,
        )
        require(
            all(p is None or not p.resolve().is_relative_to(args.report.resolve()) for p in destinations),
            "Use --input-folder for reports stored beside a reference sheet or supplement. Save outputs separately.",
        )
    return report_files(args.report)


def reviewed_schools(args, paths, reference, deferred):
    from .crosswalk import exclusion_notice

    reviewed = review_schools(args, paths, reference, deferred)
    if reviewed and reviewed["excluded_schools"]:
        print(exclusion_notice(reviewed["excluded_schools"]))
    if args.create_crosswalk:
        if reviewed is not None:
            print("School matches checked. Use this saved workbook when generating the graduates file.")
        return reviewed
    if deferred:
        print("Partial rehearsal. Leaving out " + ", ".join(sorted(deferred)) + ".")
    deferred.update(reviewed["excluded_schools"])
    if args.check_crosswalk:
        message = "School matches checked for the included campuses." if deferred else "School matches checked."
        print(terminal_style(message, "1;32", sys.stdout))
    return reviewed


def confirm_block_source(args, supplement, interactive):
    if not args.output or supplement.has_block_column or args.confirm_no_ferpa_blocks:
        return
    message = (
        "The full file will use N (no known block or unknown) for every graduate. "
        "The selected inputs have no district FERPABlock column."
    )
    if not args.guide_session or not interactive:
        raise ValidationError(
            message + " Supply a supplement with FERPABlock or confirm the district has no "
            "known blocks with --confirm-no-ferpa-blocks."
        )
    print("\n" + message)
    print("To include known blocks, return to the main menu and choose a district supplement containing FERPABlock.")
    entry = answer("Type CONFIRM if your district has no known blocks. M  Main menu   Q  Exit > ")
    if entry.casefold() == "m":
        raise ReturnToMenu()
    require(entry == "CONFIRM", "Confirmation was not recorded. Review the FERPA block source and retry.")


def confirm_save():
    while True:
        choice = answer("Return  Save this graduates file   M  Main menu   Q  Exit > ").casefold()
        if not choice:
            return
        if choice == "m":
            raise ReturnToMenu()
        print("Press Return to save or choose M or Q.")


def save_result(args, rows, reviewed):
    sample = bool(args.test_output)
    if sample:
        rows = graduate_test_sample(rows, reviewed, args.date)
    saved = save_submission(
        args.test_output or args.output, rows, name_prefix="TEST_Graduates" if sample else "Graduates"
    )
    print(terminal_style("Graduate test file saved" if sample else "Graduates file saved", "1;32", sys.stdout))
    if (args.review_test or args.review_full) and not args.no_prompt and sys.stdin.isatty() and sys.stdout.isatty():
        from .terminal_guide import review_saved_file

        status = review_saved_file(saved.resolve(), sample=sample)
        if args.guide_session:
            raise SystemExit(status)
        return
    print(saved.resolve())
    if sample:
        print("In StudentTracker choose Submit a Test File > Graduate.")


def process_graduates(args, paths, reference, reviewed, deferred):
    supplement = read_supplement(args.supplement) if args.supplement else Supplement()
    overrides = completion_overrides(args.settings)
    diplomas = DEFAULT_DIPLOMAS | overrides
    interactive = sys.stdin.isatty() and sys.stdout.isatty() and not args.no_prompt
    require(not args.review_decisions or interactive, "Open an interactive Terminal to use --review-decisions.")
    decisions = ReviewDecisions(
        args.crosswalk,
        reference["account_number"] if reference else None,
        args.year,
        args.supplement,
        interactive=interactive and (args.guide_session or args.review_decisions),
        revisit=args.review_decisions,
        guided=args.guide_session,
        preview=bool(args.compare_dates),
    )
    dates = supplement.dates if supplement.has_date_column and not args.calpads_exit_dates else None
    try:
        if args.compare_dates:
            require(supplement.has_date_column, "To compare dates, include GraduationDate in the district supplement.")
            compare_dates(
                args.compare_dates,
                paths,
                args.year,
                diplomas,
                deferred,
                dates,
                day(args.date, "transmission date"),
                decisions,
            )
            print("Comparison CSV saved. Check Eligibility and Status for pending decisions.")
            return
        rows = convert(args, paths, reviewed, diplomas, deferred, dates, supplement, decisions)
    finally:
        decisions.summary()
    if args.output or args.test_output:
        confirm_block_source(args, supplement, interactive)
        if args.guide_session and interactive:
            confirm_save()
        save_result(args, rows, reviewed)
    elif deferred:
        print("Dry run passed for the included campuses.")
    else:
        print(terminal_style("Validation passed.", "1;32", sys.stdout))


def main() -> None:
    arguments = sys.argv[1:]
    if [arg for arg in arguments if arg != "--debug"] in ([], ["--guide"]):
        from .terminal_guide import main as guide

        raise SystemExit(guide())
    args = parse_arguments(arguments)
    if args.check_file:
        show_file_check(args.check_file)
        return
    if args.compare_dates:
        args.no_prompt = True
    paths = selected_reports(args)
    if args.check_headers:
        for _ in report_records(paths, headers_only=True):
            pass
        print("Report headers passed.")
        return
    args.year = reporting_year(paths, args.year)
    print(completion_year_label(args.year))
    deferred = set(args.defer_school)
    reference = read_reference(args.reference_sheet) if args.reference_sheet else None
    reviewed = None
    if args.create_crosswalk or args.crosswalk:
        reviewed = reviewed_schools(args, paths, reference, deferred)
        if args.create_crosswalk or args.check_crosswalk:
            return
    process_graduates(args, paths, reference, reviewed, deferred)


ERRORS = {
    FileExistsError: "A file already exists there. Choose a new filename or an output directory.",
    PermissionError: "Access denied. Check read access to the inputs and write access to the output folder.",
    csv.Error: "Check the text export. Use a flat delimited file saved with UTF-8 or UTF-16 encoding.",
    UnicodeError: "Check the text export encoding. Save the export with UTF-8 or UTF-16 encoding.",
    json.JSONDecodeError: "The settings file contains invalid JSON. Check its quotes, commas and braces.",
    BadZipFile: "The Excel workbook could not be read. Save a valid .xlsx copy or export the report again.",
    ImportError: "A required Python package is missing. Install the project and its dependencies in this environment.",
    OSError: "A file operation failed. Check file access and free space, then retry.",
}


def error_message(exc: Exception) -> str:
    if isinstance(exc, ValidationError):
        return str(exc)
    if isinstance(exc, FileNotFoundError):
        filename = Path(exc.filename).name if exc.filename else "selected input"
        return f"File or folder not found: {filename}. Check the selected location."
    return next(
        (message for error, message in ERRORS.items() if isinstance(exc, error)),
        "Unable to read the selected input. Check its format and try again.",
    )


def debug_location(exc: Exception) -> str:
    frame = exc.__traceback__
    location = "entrypoint"
    while frame:
        module = frame.tb_frame.f_globals.get("__name__", "")
        if module.startswith("calpads_to_nsc."):
            location = f"{module}.{frame.tb_frame.f_code.co_name}"
        frame = frame.tb_next
    return f"Diagnostic · {type(exc).__name__} · {location}"


def entrypoint() -> None:
    try:
        main()
    except ReturnToMenu:
        sys.exit(RETURN_TO_MENU)
    except (Finished, EOFError, KeyboardInterrupt):
        print("\nReview closed. Saved decisions will be available next time.")
        sys.exit(PAUSED)
    except Exception as exc:
        if "--debug" in sys.argv[1:]:
            print(debug_location(exc), file=sys.stderr)
        sys.exit(error_message(exc))


if __name__ == "__main__":
    entrypoint()
