"""Build disposable files and a reconciliation manifest for the Excel test."""
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

from kvd.demo import REPORTS, seed
from kvd.excel_commit import commit
from kvd.intake import parse
from kvd.proposals import approve, preview
from kvd.journal import fingerprint


def prepare(directory):
    path, journal = seed(directory)
    book = load_workbook(path)
    try:
        for index, entries in ((10, {"A5": "SYN-CUT", "E5": "m3", "F5": 100,
                                      "A6": "SYN-DRAIN", "E6": "m", "F6": 40}),
                               (12, {"A5": "SYN-BQ-1", "C5": "m3", "D5": 100})):
            for cell, value in entries.items():
                book.worksheets[index][cell] = value
        book.save(path)
    finally:
        book.close()
    drafts = parse((REPORTS.parent / "synthetic-report-ms.txt").read_text(encoding="utf-8"))
    drafts += parse(REPORTS.read_text(encoding="utf-8"))[1:]
    decisions = [{"verified_qty": "7", "verification_evidence": "SYN-QA-01"}] + [{}] * 4
    token, _ = preview(journal, path, drafts, decisions, "Synthetic Operator")
    approve(journal, token, "Synthetic Reviewer", True)
    output = path.with_name("synthetic-committed.xlsx")
    result = commit(journal, token, output, "Synthetic Operator")
    assert fingerprint(path) == fingerprint(result["backup"])
    return result


if __name__ == "__main__":
    print(json.dumps(prepare(Path(sys.argv[1])), ensure_ascii=True))
