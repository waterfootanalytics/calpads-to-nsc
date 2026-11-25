from .generate import ValidationError


PAUSED = 3
FINISHED_AFTER_SAVE = 4
RETURN_TO_MENU = 5


class Finished(Exception):
    pass


def answer(prompt):
    value = input(prompt).strip()
    if value.casefold() in {"q", "quit", "exit", "close", "done"}:
        raise Finished()
    return value


class ReturnToMenu(Exception):
    pass


def confirm_completion(code, choice, guided):
    action = "EXCLUDE" if choice == "exclude" else "APPLY"
    label = "Exclude from submission" if choice == "exclude" else choice
    print(f"Every graduate with completion code {code} in this reporting cycle · {label}")
    entry = answer(f"Type {action} to confirm, or press Return to choose again > ")
    if guided and entry.casefold() == "m":
        raise ReturnToMenu()
    return entry == action


def ask_graduation_decision(message, options, saved, completion, guided):
    print("\n" + message)
    if saved is not None:
        print(f"Return  Keep the saved choice · {dict(options)[saved]}")
    numbered = {}
    for index, (choice, label) in enumerate(options, 1):
        numbered[str(index)] = choice
        print(f"  {index}  {label}")
    print("  R  Correct the source file and retry")
    if guided:
        print("  M  Main menu")
    print("  Q  Save progress and exit")
    while True:
        entry = answer("Choose an option > ")
        if not entry and saved is not None:
            return None
        if entry.casefold() == "r":
            raise ValidationError("After correcting and saving your source file, retry with the same inputs.")
        if guided and entry.casefold() == "m":
            raise ReturnToMenu()
        if entry not in numbered:
            print("Choose one of the options shown.")
            continue
        choice = numbered[entry]
        if completion and not confirm_completion(completion, choice, guided):
            print("Selection unchanged. Choose an option.")
            continue
        return choice
