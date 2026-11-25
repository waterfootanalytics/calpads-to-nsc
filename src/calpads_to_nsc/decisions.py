from .decision_store import DecisionStore, fingerprint
from .generate import DIPLOMAS, ValidationError
from .terminal_input import ask_graduation_decision


PENDING = object()


def source_location(location):
    path, row = location
    return f"CALPADS file · {path.name}\nRow {row}"


class ReviewDecisions:
    def __init__(
        self,
        workbook,
        account,
        year,
        supplement,
        *,
        interactive=False,
        revisit=False,
        guided=False,
        preview=False,
    ):
        self.store = DecisionStore(workbook, fingerprint("graduation-decisions-v1", account, year))
        self.supplement = supplement
        self.guided = guided
        self.interactive = interactive
        self.revisit = revisit
        self.preview = preview
        self.completions = {}
        self.used = False

    def select(self, key, message, options, *, completion=None):
        choices = dict(options)
        saved = self.store.load(key)
        if saved is not None and saved not in choices:
            raise ValidationError(
                "The saved selection is invalid. Restore the private decisions file and review it again."
            )
        if saved is not None and (self.preview or not self.revisit):
            self.used = True
            return saved
        if self.preview:
            return PENDING
        if not self.interactive:
            raise ValidationError(
                message + "\nOpen the terminal guide and choose Review graduation decisions. "
                "For a scripted interactive run, include the school workbook and --review-decisions."
            )
        selection = ask_graduation_decision(message, options, saved, completion, self.guided)
        if selection is not None:
            self.store.save(key, selection)
            print("Saved · " + choices[selection])
        self.used = True
        return saved if selection is None else selection

    def completion(self, code, location, diplomas):
        if code in diplomas:
            self.completions[code] = diplomas[code]
            return diplomas[code]
        if code not in self.completions:
            description = " (Adult Education High School Diploma)" if code == "250" else ""
            message = (
                f"{source_location(location)} · Completion code {code}{description} — district decision.\n"
                "Check the credential awarded. Choose its NSC diploma type, or exclude this code for this cycle."
            )
            options = [(d, d) for d in sorted(DIPLOMAS)] + [("exclude", "Exclude graduates with this completion code")]
            choice = self.select("completion:" + code, message, options, completion=code)
            self.completions[code] = None if choice == "exclude" else choice
        return self.completions[code]

    def date_source(self, location, ssid, birth_day, exit_day, entries, reason, *, allow_diploma):
        path, row_number = location
        file = self.supplement
        rows = ", ".join(str(num) for num, _ in entries)
        message = (
            f"{source_location(location)} · {reason}\n"
            f"District supplement · {file.name} · {'Row' if len(entries) == 1 else 'Rows'} {rows}\n"
            "Check these rows against the district's diploma record before choosing the date to submit."
        )
        evidence = (ssid, birth_day.isoformat(), exit_day.isoformat(), entries)
        key = "date:" + fingerprint("v2", path.name, row_number, file.name, *evidence)
        options = [("calpads", "Use the CALPADS exit date for this graduate")]
        if allow_diploma:
            options.append(("supplement", "Use the confirmed diploma date from the district supplement"))
        return self.select(key, message, options)

    def summary(self):
        for code, choice in sorted(self.completions.items()):
            if choice is not PENDING:
                print(f"Completion code {code} · {choice or 'Exclude from submission'}")
        if self.used and self.store.path and not self.preview:
            print("To change saved choices, select Review graduation decisions.")
