import hashlib
import os
import re
import shlex
import shutil
import subprocess
import sys

from .generate import ValidationError, require
from .reports import report_schools
from .terminal_input import PAUSED, answer


def terminal_style(text, code, stream):
    if stream.isatty() and os.environ.get("TERM") != "dumb" and "NO_COLOR" not in os.environ:
        return f"\033[{code}m{text}\033[0m"
    return text


def show_review_step(message, stream):
    heading, *lines = message.splitlines()
    print("\n" + terminal_style(heading, "1;36", stream), file=stream)
    for line in lines:
        if line.startswith("  "):
            line = re.sub(r"\b[A-Z]+[0-9]+\b", lambda match: terminal_style(match[0], "1;33", stream), line, count=1)
        print(line, file=stream)
    stream.flush()


def open_workbook(path, *, excel_only=False):
    filename = str(path.resolve())
    try:
        if os.name == "nt":
            if excel_only:
                os.startfile("excel.exe", arguments=subprocess.list2cmdline([filename]))
            else:
                os.startfile(filename)
            return True
        if sys.platform == "darwin":
            result = subprocess.run(
                ["open", "-a", "Microsoft Excel", filename], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            if result.returncode == 0:
                return True
        elif excel_only and (os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")):
            calc = shutil.which("libreoffice") or shutil.which("soffice")
            if calc:
                subprocess.Popen(
                    [calc, "--calc", "--view", filename], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
                return True
        if not excel_only:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.run([opener, filename], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
    except (OSError, subprocess.CalledProcessError):
        pass
    print("In your spreadsheet application, choose File > Open and select this file:\n" + filename)
    return False


def offer_workbook(path):
    if not (sys.stdin.isatty() and sys.stdout.isatty()):
        return False
    enter = terminal_style("Enter", "1;36", sys.stdout)
    later = terminal_style("N", "1", sys.stdout)
    while True:
        try:
            choice = answer(f"\nOpen the Excel workbook?   {enter} = Open    {later} = Later    Q = Exit   ").casefold()
        except EOFError:
            print()
            return False
        if choice in ("n", "no"):
            return False
        if choice in ("", "y", "yes"):
            return open_workbook(path)
        print("Press Enter to open the workbook, or N for later.")


def ready_to_recheck():
    shortcut = "⌘S" if sys.platform == "darwin" else "Ctrl+S"
    enter = terminal_style("Enter", "1;36", sys.stdout)
    later = terminal_style("Q", "1", sys.stdout)
    while True:
        try:
            choice = answer(
                f"\nSave in Excel ({shortcut}), then return here.\n"
                f"{enter} = Check saved changes    {later} = Exit guide   "
            )
        except EOFError:
            print()
            return False
        if not choice:
            return True
        print("Press Enter to read your saved choices, or Q to exit.")


def check_workbook_command(args, path):
    command = [
        sys.executable,
        "-m",
        "calpads_to_nsc",
        str(args.report.resolve()),
        "--reference-sheet",
        str(args.reference_sheet.resolve()),
        "--crosswalk",
        str(path),
        "--check-crosswalk",
    ]
    if getattr(args, "input_folder", False):
        command.append("--input-folder")
    return subprocess.list2cmdline(command) if os.name == "nt" else shlex.join(command)


def completed_review_action(path, reference, result, version, year):
    from .terminal_schools import edit_school, pick_saved_match

    status = (
        "School review complete, including saved exclusions."
        if result["excluded_schools"]
        else "Every school has a match."
    )
    print("\n" + terminal_style(status, "1;32", sys.stdout))
    key = answer(
        "Return  Keep saved matches\nM       View or change a match or exclusion\n"
        "X       Open Excel\nQ       Exit guide\n> "
    ).casefold()
    if not key:
        return "done"
    if key == "m":
        selected = pick_saved_match(path, reference)
        return edit_school(path, reference, selected, version, year=year) if selected else "retry"
    if key == "x":
        return "excel"
    print("Press Return or enter one of the letters shown.")
    return "retry"


def edit_remaining(path, reference, problem, version, year):
    from .terminal_schools import edit_school

    if not problem.edits:
        show_review_step(str(problem), sys.stdout)
        key = answer("X  Open Excel to fix the workbook\nQ  Exit guide\n> ").casefold()
        return "excel" if key == "x" else "retry"
    count = len(problem.edits)
    title = "Match the remaining school" if count == 1 else f"{count} schools to match"
    print("\n" + terminal_style(title, "1;36", sys.stdout))
    return edit_school(path, reference, problem.edits[0], version, year=year)


def terminal_school_review(path, labels, reference, digest, args, deferred):
    from .crosswalk import WorkbookEditsNeeded, exclusion_notice, reviewed_config

    review_only = bool(args.check_crosswalk or args.create_crosswalk)
    last_notice = ()
    try:
        while True:
            version = path.stat()
            try:
                result = reviewed_config(
                    path,
                    labels,
                    reference,
                    digest,
                    bool(args.output),
                    deferred=deferred,
                    test_file=bool(args.test_output),
                    year=args.year,
                )
            except WorkbookEditsNeeded as exc:
                if exc.excluded and exc.excluded != last_notice:
                    print(exclusion_notice(exc.excluded))
                last_notice = exc.excluded
                action = edit_remaining(path, reference, exc, version, args.year)
            else:
                if not review_only:
                    return result
                action = completed_review_action(path, reference, result, version, args.year)
                if action == "done":
                    return result
            if action == "excel":
                open_workbook(path)
                action = "retry" if ready_to_recheck() else "pause"
            if action == "pause":
                break
    except EOFError:
        print()
    print("Review paused. Your saved choices will be here next time.")
    raise SystemExit(PAUSED)


def create_review(args, path, labels, reference, digest, interactive):
    from .crosswalk import create

    try:
        create(path, labels, reference, digest, args.rehearsal, previous=args.crosswalk, year=args.year)
    except FileExistsError:
        raise ValidationError(
            f"A workbook already exists at {path}. To start a new review, choose a new filename. "
            "To open your saved review and check the choices, run\n" + check_workbook_command(args, path)
        ) from None
    print(terminal_style("School review workbook created", "1;32", sys.stdout))
    if not interactive:
        print(path)


def read_excel_choices(path, labels, reference, digest, args, deferred, interactive):
    from .crosswalk import WorkbookEditsNeeded, reviewed_config

    try:
        return reviewed_config(
            path,
            labels,
            reference,
            digest,
            bool(args.output),
            deferred=deferred,
            test_file=bool(args.test_output),
            year=args.year,
        )
    except WorkbookEditsNeeded as exc:
        stream = sys.stdout if args.create_crosswalk else sys.stderr
        if not interactive and not args.create_crosswalk:
            print(path, file=stream)
        show_review_step(str(exc), stream)
        return None


def excel_school_review(path, labels, reference, digest, args, deferred, interactive):
    creating = bool(args.create_crosswalk)
    direct_open = args.open_review and interactive
    opened = False
    while True:
        reviewed = read_excel_choices(path, labels, reference, digest, args, deferred, interactive)
        if reviewed is not None:
            if opened or not (creating or direct_open):
                return reviewed
            show_review_step(
                "Review your matches\nOn Match schools, keep correct choices and change any that "
                "need a different campus.",
                sys.stdout,
            )
        if not opened:
            if direct_open:
                open_workbook(path)
                opened = True
            else:
                opened = not args.no_prompt and offer_workbook(path)
        if not opened or not ready_to_recheck():
            break
    if direct_open:
        print("Review paused. You can resume from Terminal.")
        raise SystemExit(PAUSED)
    if creating:
        message = (
            "Review paused. You can resume from Terminal."
            if interactive
            else "Check your saved choices with this command.\n" + check_workbook_command(args, path)
        )
        print(message)
        return None
    if reviewed is not None:
        return reviewed
    print("Save the workbook and run the same command to continue.", file=sys.stderr)
    raise SystemExit(1)


def review_schools(args, paths, reference, deferred):
    require(reference, "Supply the account reference sheet with --reference-sheet.")
    path = (args.create_crosswalk or args.crosswalk).resolve()
    digest = hashlib.sha256(args.reference_sheet.read_bytes()).hexdigest()
    labels = report_schools(paths)
    interactive = sys.stdin.isatty() and sys.stdout.isatty() and not args.no_prompt
    if args.create_crosswalk:
        create_review(args, path, labels, reference, digest, interactive)
    else:
        require(
            path.is_file(), f"Excel review workbook not found at {path}. Choose your saved workbook with --crosswalk."
        )
    if interactive and not args.open_review:
        return terminal_school_review(path, labels, reference, digest, args, deferred)
    return excel_school_review(path, labels, reference, digest, args, deferred, interactive)
