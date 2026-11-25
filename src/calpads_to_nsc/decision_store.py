import hashlib
import json
import os
import re

from .generate import DIPLOMAS, ValidationError
from .storage import private_output_path, private_update_lock


def fingerprint(*fields):
    return hashlib.sha256(json.dumps(fields, ensure_ascii=True, separators=(",", ":")).encode()).hexdigest()


def valid_choice(key, choice):
    if not isinstance(choice, str):
        return False
    if re.fullmatch("completion:[0-9]{3}", key):
        return choice in DIPLOMAS or choice == "exclude"
    return bool(re.fullmatch("date:[a-f0-9]{64}", key)) and choice in {"calpads", "supplement"}


def decode_json(path):
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or type(data.get("version")) is not int or data["version"] != 1:
        raise ValueError()
    if not isinstance(data.get("decisions"), dict):
        raise ValueError()
    choices = data["decisions"]
    for scope, rows in choices.items():
        if not re.fullmatch("[a-f0-9]{64}", scope) or not isinstance(rows, dict):
            raise ValueError()
        if not all(valid_choice(key, choice) for key, choice in rows.items()):
            raise ValueError()
    return choices


class DecisionStore:
    def __init__(self, workbook, scope):
        self.scope = scope
        self.path = workbook.with_suffix(".decisions.json") if workbook else None
        self.choices = self.read()

    def read(self):
        if self.path is None:
            return {}
        try:
            return decode_json(self.path)
        except FileNotFoundError:
            return {}
        except (ValueError, OSError):
            raise ValidationError(
                "The saved graduation decisions could not be read. Check access or restore the private decision file."
            ) from None

    def load(self, key):
        return self.choices.get(self.scope, {}).get(key)

    def save(self, key, choice):
        if self.path is None:
            raise ValidationError("Select the school review workbook to save district decisions.")
        try:
            with private_update_lock(self.path):
                current = self.read()
                group = current.setdefault(self.scope, {})
                if group.get(key) not in (self.load(key), choice):
                    raise ValidationError("Another session changed this decision. Open the review again.")
                group[key] = choice
                with private_output_path(self.path.parent) as temp:
                    temp.write_text(
                        json.dumps({"version": 1, "decisions": current}, sort_keys=True) + "\n", encoding="utf-8"
                    )
                    os.replace(temp, self.path)
                self.choices = current
        except OSError:
            raise ValidationError(
                "Could not save graduation decisions. Close other editing sessions, check write access and retry."
            ) from None
