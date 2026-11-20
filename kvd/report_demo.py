"""Build and issue synthetic MC10 and MC02 reports for evaluation."""
import argparse
from pathlib import Path

from openpyxl import load_workbook

from . import reports
from .demo import REPORTS, seed
from .excel_commit import commit
from .intake import parse
from .proposals import approve, preview


ANCHOR = "2026-01-15"


def snapshot(directory):
    source, journal = seed(directory)
    north = "KVDT2-WF-000001"
    east = "KVDT2-WF-000002"
    book = load_workbook(source)
    try:
        fixtures = {
            2: {"A5": "SYN-ACTION-1", "B5": "2026-01-12", "D5": "Confirm mock access route",
                "E5": "High", "F5": "Mock Coordinator", "G5": "2026-01-16", "H5": "Open",
                "J5": "Review mock route sketch",
                "A6": "SYN-ACTION-CLOSED", "B6": "2026-01-12", "D6": "Archived mock action", "H6": "Closed"},
            5: {"A5": "SYN-PREV-1", "B5": "2026-01-13", "C5": "Day", "D5": "SYN-CUT",
                "E5": "SYN-BQ-1", "F5": "Mock North Zone", "G5": "Trial excavation",
                "H5": "Mock Crew A", "I5": "m3", "J5": 15, "K5": 14, "L5": 10, "R5": north,
                "A6": "SYN-PREV-2", "B6": ANCHOR, "C6": "Night", "D6": "SYN-DRAIN",
                "F6": "Mock East Zone", "G6": "Drain bedding", "H6": "Mock Crew B",
                "I6": "m", "J6": 14, "K6": 0, "L6": 0, "R6": east},
            7: {"A5": "SYN-RFI-1", "C5": "2026-01-14", "E5": "Confirm mock trench detail",
                "I5": "Mock Design Lead", "J5": "2026-01-17", "K5": "Open",
                "A6": "SYN-RFI-CLOSED", "C6": "2026-01-12", "E6": "Archived mock query", "K6": "Closed"},
            8: {"A5": "SYN-INSPECT-1", "B5": ANCHOR, "C5": "SYN-DRAIN", "H5": "Bedding inspection",
                "I5": "Pending", "K5": "Mock Inspector"},
            9: {"A5": "SYN-NCR-1", "B5": ANCHOR, "D5": "Mock compaction follow-up", "H5": "Open",
                "I5": "Review mock results"},
            14: {"A5": "SYN-SUB-1", "B5": "Mock excavation method", "E5": "2026-01-13", "H5": "Pending"},
            20: {"B5": "2026-01-14", "E5": "Record mock bedding measurement",
                 "F5": "Mock Junior Engineer", "G5": "Medium", "H5": "2026-01-16", "L5": "Pending"},
        }
        for index, cells in fixtures.items():
            for reference, value in cells.items():
                book.worksheets[index][reference] = value
        book.save(source)
    finally:
        book.close()
    drafts = parse(REPORTS.read_text(encoding="utf-8"))
    choices = [{"verified_qty": "7", "verification_evidence": "SYN-MEASURE-1"}] + [{}] * 4
    token, _ = preview(journal, source, drafts, choices, "Synthetic Operator")
    approve(journal, token, "Synthetic Quantity Reviewer", True)
    output = source.with_name("synthetic-reporting-snapshot.xlsx")
    commit(journal, token, output, "Synthetic Operator")
    return output, journal


def samples(directory):
    output, journal = snapshot(directory)
    exported = []
    for command in ("MC10", "MC02"):
        identifier = reports.create(journal, output, command, ANCHOR, "Synthetic Report Author")
        reports.approve(journal, identifier, "Synthetic Report Reviewer")
        reports.issue(journal, identifier, "Synthetic Issuer")
        exported.append(reports.export(journal, identifier, directory))
    return output, journal, exported


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory")
    args = parser.parse_args()
    output, journal, folders = samples(Path(args.directory))
    print(f'python -m kvd.report_desktop --journal "{journal.path}" --workbook "{output}"')
    for folder in folders:
        print(folder)


if __name__ == "__main__":
    main()
