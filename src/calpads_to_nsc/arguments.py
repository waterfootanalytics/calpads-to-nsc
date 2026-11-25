import argparse
from collections.abc import Sequence
from datetime import date
from pathlib import Path

from . import NON_AFFILIATION, __version__
from .generate import day, require


def parse_arguments(arguments: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="prepare-nsc-graduates",
        description="Prepare a StudentTracker graduates file from certified CALPADS 1.23 reports.",
        epilog="Run with --guide by itself for the interactive Terminal menu. " + NON_AFFILIATION,
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--debug", action="store_true", help="Show the error type and program location on failure.")
    add_policy_options(parser)
    add_inputs(parser)
    add_review_options(parser)
    add_date_options(parser)
    add_operations(parser)
    args = parser.parse_args(arguments)
    if args.output or args.test_output:
        require(
            day(args.date, "transmission date") <= date.today(),
            "The file transmission date is in the future. Use today or an earlier date and try again.",
        )
    if args.check_file:
        flags = {arg.split("=")[0] for arg in arguments if arg.startswith("--")}
        if args.report is not None or flags - {"--check-file", "--debug"}:
            parser.error("Use --check-file FILE on its own.")
    elif args.report is None:
        parser.error("Choose a CALPADS report or folder.")
    validate_options(args)
    return args


def add_inputs(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("report", nargs="?", type=Path, help="One report or a folder of per-school reports.")
    parser.add_argument("--guide-session", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--settings", type=Path, help="Optional JSON file containing district completion_types decisions."
    )
    parser.add_argument(
        "--review-decisions",
        action="store_true",
        help="Review saved graduation-date and completion decisions interactively for this workbook.",
    )
    parser.add_argument(
        "--input-folder",
        action="store_true",
        help="Read CALPADS exports from a folder shared with the reference sheet and supplement.",
    )
    parser.add_argument("--crosswalk", type=Path, help="Excel workbook with reviewed campus selections.")


def add_review_options(parser: argparse.ArgumentParser) -> None:
    prompts = parser.add_mutually_exclusive_group()
    prompts.add_argument(
        "--no-prompt", action="store_true", help="Show review instructions without opening a workbook or waiting."
    )
    prompts.add_argument("--terminal-review", action="store_true", help="Review and save school matches in Terminal.")
    prompts.add_argument(
        "--open-review", action="store_true", help="Open the workbook directly during an interactive school review."
    )
    parser.add_argument(
        "--review-test",
        action="store_true",
        help="Offer Excel review and upload instructions after saving a Graduate test file.",
    )
    parser.add_argument(
        "--review-full",
        action="store_true",
        help="Offer Excel review and secure FTP instructions after saving a full graduates file.",
    )
    parser.add_argument(
        "--rehearsal", action="store_true", help="Create a workbook for local checks and Graduate test files."
    )
    parser.add_argument(
        "--defer-school",
        action="append",
        default=[],
        metavar="CALPADS_NAME",
        help="Leave a named unresolved school out of a rehearsal or Graduate test file. Repeat for another school.",
    )


def add_date_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--reference-sheet", type=Path, help="NSC account reference sheet (.xls, .xlsx or text PDF).")
    parser.add_argument(
        "--year",
        type=int,
        help="Verify that every CALPADS exit date belongs to this reporting year. "
        "By default the year is determined from the reports (August 16 through August 15).",
    )
    parser.add_argument(
        "--date",
        metavar="YYYYMMDD",
        default=date.today().strftime("%Y%m%d"),
        help="File transmission date, defaulting to today. A test or full file can use today or an earlier date.",
    )
    parser.add_argument(
        "--supplement",
        type=Path,
        metavar="CSV",
        help="District CSV keyed by SSID with optional diploma dates, FERPA blocks, academic fields or program codes.",
    )
    parser.add_argument(
        "--calpads-exit-dates",
        action="store_true",
        help="Use only the FERPA block column from the supplement and keep CALPADS exit dates.",
    )


def add_policy_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--include-demographics",
        action="store_true",
        help="Fill supported demographic fields from the certified CALPADS reports.",
    )
    parser.add_argument(
        "--use-special-education-for-disability",
        action="store_true",
        help="Use CALPADS special education status for the district's NSC disability indicator (column Z).",
    )
    parser.add_argument(
        "--student-id",
        choices=("local", "ssid", "blank"),
        default="local",
        help="Value to put in NSC column J. The default is the local student ID.",
    )
    parser.add_argument(
        "--confirm-no-ferpa-blocks",
        action="store_true",
        help="Confirm no known FERPA blocks for a full file without a supplement block column.",
    )


def add_operations(parser: argparse.ArgumentParser) -> None:
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check-file", type=Path, metavar="TXT", help="Check a saved graduates text file on its own.")
    mode.add_argument("--check-headers", action="store_true", help="Check the field names in every CALPADS report.")
    mode.add_argument(
        "--create-crosswalk",
        type=Path,
        metavar="NEW_XLSX",
        help="Generate the Excel school review from CALPADS and the account reference sheet.",
    )
    mode.add_argument(
        "--check-crosswalk", action="store_true", help="Validate campus selections against the account reference sheet."
    )
    mode.add_argument(
        "--compare-dates",
        type=Path,
        metavar="NEW_CSV",
        help="Compare CALPADS exit dates with --supplement and save private grouped counts.",
    )
    mode.add_argument("--dry-run", action="store_true", help="Validate the cohort and report any unresolved decisions.")
    mode.add_argument(
        "--test-output",
        type=Path,
        metavar="FILE_OR_DIR",
        help="Save one eligible graduate per campus. An existing folder gets a timestamped test file.",
    )
    mode.add_argument(
        "--output",
        type=Path,
        metavar="FILE_OR_DIR",
        help="Save all eligible graduates. An existing folder gets a timestamped graduates file.",
    )


def validate_options(args: argparse.Namespace) -> None:
    require(args.year is None or 1901 <= args.year <= 9999, "Use a four-digit CALPADS reporting year from 1901 onward.")
    require(not args.review_test or args.test_output, "Use --review-test with --test-output.")
    require(not args.review_full or args.output, "Use --review-full with --output.")
    require(
        not args.open_review or args.create_crosswalk or args.check_crosswalk,
        "Use --open-review when creating or checking the school workbook.",
    )
    validate_school_options(args)
    generating = bool(args.dry_run or args.test_output or args.output)
    require(not args.include_demographics or generating, "Choose a dry run, test file or full file for demographics.")
    require(
        not args.use_special_education_for_disability or args.include_demographics,
        "Use --use-special-education-for-disability with --include-demographics.",
    )
    validate_graduation_options(args, generating)


def validate_school_options(args: argparse.Namespace) -> None:
    deferred = args.defer_school
    require(
        not deferred
        or (args.crosswalk and (args.check_crosswalk or args.compare_dates or args.dry_run or args.test_output)),
        "Use --defer-school with --crosswalk and --check-crosswalk, --compare-dates, --dry-run or --test-output.",
    )
    require(not args.rehearsal or args.create_crosswalk, "Use --rehearsal with --create-crosswalk.")
    require(
        not args.crosswalk
        or args.create_crosswalk
        or args.check_crosswalk
        or args.compare_dates
        or args.dry_run
        or args.test_output
        or args.output,
        "Choose a crosswalk review, date comparison or graduate-file operation.",
    )


def validate_graduation_options(args: argparse.Namespace, generating: bool) -> None:
    require(
        not (args.check_crosswalk or generating) or args.crosswalk,
        "Pass the reviewed workbook with --crosswalk. Create it first using --create-crosswalk.",
    )
    require(
        not (args.crosswalk or args.create_crosswalk) or args.reference_sheet,
        "Supply the current NSC Account Reference Sheet with --reference-sheet.",
    )
    require(
        not args.review_decisions or (args.crosswalk and generating and not args.no_prompt),
        "Use --review-decisions with a reviewed school workbook and a graduation-file operation in Terminal.",
    )
    require(
        not args.settings or args.compare_dates or generating,
        "Use --settings with --compare-dates, --dry-run, --test-output or --output.",
    )
    require(not args.compare_dates or args.supplement, "Supply --supplement with --compare-dates.")
    require(
        not args.supplement or args.compare_dates or generating,
        "Use --supplement with --compare-dates, --dry-run, --test-output or --output.",
    )
    require(
        not args.calpads_exit_dates or generating or args.compare_dates,
        "Choose a comparison, dry run, test file or full file when using --calpads-exit-dates.",
    )
