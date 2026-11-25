import csv
from collections import Counter

from .decisions import PENDING
from .generate import require
from .graduates import eligible_graduates
from .reports import report_date
from .storage import private_output_path, publish_new
from .supplement import select_graduation_date


def compare_dates(destination, paths, year, diplomas, deferred, dates, send, decisions):
    require(destination.suffix.lower() == ".csv", "Use a new .csv filename for the date comparison.")
    counts = Counter()
    for location, record, row, exit_day, diploma in eligible_graduates(paths, year, diplomas, deferred, decisions):
        dob = report_date(location, record["dob"], "birth date")
        result = select_graduation_date(dates, row["ssid"], exit_day, dob, send, location, decisions)
        eligibility = f"Needs decision: completion code {row['status'][:3]}" if diploma is PENDING else "Eligible"
        known = diploma is not PENDING and result.selected is not None
        evidence = result.candidate
        counts[
            (
                result.status,
                exit_day.isoformat(),
                evidence.isoformat() if evidence else "",
                (exit_day - evidence).days if evidence else "",
                result.source if known else "",
                result.selected.isoformat() if known else "",
                eligibility,
            )
        ] += 1
    require(counts, "No eligible or pending CALPADS graduates to compare.")
    with private_output_path(destination.parent) as temp:
        with temp.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(
                [
                    "Status",
                    "CALPADSExitDate",
                    "SupplementGraduationDate",
                    "DaysDifference",
                    "DateSource",
                    "NSCGraduationDate",
                    "Students",
                    "Eligibility",
                ]
            )
            for group, count in sorted(counts.items()):
                writer.writerow([*group[:6], count, group[-1]])
        publish_new(temp, destination)
