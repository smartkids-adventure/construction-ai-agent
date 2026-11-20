"""Populate a local journal from synthetic P4 data and issue one example."""
import argparse
from pathlib import Path
import json

from openpyxl import load_workbook

from .demo import REPORTS, seed
from .excel_commit import commit
from .intake import parse
from .journal import fingerprint
from .materials import apply, read
from .material_requisition import export
from .proposals import approve, preview


ANCHOR = "2026-01-15"
LOCATION = "Mock Central Store"
NORTH = "KVDT2-WF-000001"
EAST = "KVDT2-WF-000002"
MANIFEST = REPORTS.parent / "p4-materials.json"


def workbook_snapshot(directory):
    source, journal = seed(directory)
    book = load_workbook(source)
    try:
        row = book.worksheets[11]
        for cell, value in {"A5": "SYN-CUT", "C5": "Synthetic excavation programme",
                            "D5": ANCHOR, "E5": "2026-01-28", "F5": "m", "G5": 100}.items():
            row[cell] = value
        book.save(source)
    finally:
        book.close()
    draft = parse(REPORTS.read_text(encoding="utf-8"))
    token, _ = preview(journal, source, draft, [{}] * len(draft), "Synthetic Intake Operator")
    approve(journal, token, "Synthetic Intake Reviewer")
    output = source.with_name("synthetic-material-snapshot.xlsx")
    commit(journal, token, output, "Synthetic Intake Operator")
    return output, journal


def fixtures(directory):
    workbook, journal = workbook_snapshot(directory)
    actor = "Synthetic Planner"
    apply(journal, actor, json.loads(MANIFEST.read_text(encoding="utf-8")))
    for short, material, unit, workfront, activity, design, opening, daily, rate, end in (
        ("REBAR", "SYN-REBAR", "kg", NORTH, "SYN-CUT", 300, 50, None, 2, "2026-01-28"),
        ("DUCT", "SYN-DUCT", "m", EAST, "SYN-DRAIN", 60, 120, 5, 1, "2026-01-26"),
        ("GROUT", "SYN-GROUT", "kg", EAST, "SYN-DRAIN", 30, 35, 5, 1, "2026-01-21"),
        ("PANEL", "SYN-PANEL", "ea", EAST, "SYN-DRAIN", 10, 8, None, None, "2026-01-21"),
    ):
        drawing = apply(journal, actor, [{"operation": "drawing", "number": f"SYN-DR-{short}", "revision": "A",
                         "status": "PENDING" if short == "PANEL" else "APPROVED",
                         "reference": f"Synthetic drawing schedule {short}/A", "snapshot": fingerprint(workbook)}])[0]
        apply(journal, actor, [
            {"operation": "takeoff", "id": f"TO-{short}", "drawing": drawing, "workfront": workfront,
             "material_id": material, "element": "Synthetic schedule total", "dimensions": {
             "measured": {"value": design, "unit": unit}}, "formula": "measured", "unit": unit,
             "source": f"Synthetic drawing schedule {short}/A line 1", "allowance_pct": "5"},
            {"operation": "movement", "id": f"OPEN-{short}", "material_id": material, "unit": unit,
             "location": LOCATION, "date": "2026-01-10", "kind": "opening", "quantity": opening,
             "source": "Synthetic counted opening stock"},
        ])
        for action, approver in (("verify", "Synthetic Quantity Verifier"), ("approve", "Synthetic Engineer")):
            if short != "PANEL":
                apply(journal, approver, [{"operation": "takeoff-review", "id": f"TO-{short}",
                                          "action": action, "reference": f"SYN-{action}-{short}"}])
        plan = {"id": f"PLAN-{short}", "takeoff": f"TO-{short}", "location": LOCATION,
                "workdays": [0, 1, 2, 3, 4], "rate": rate, "unit": unit}
        if short == "REBAR":
            plan.update(operation="programme-import", workbook=str(workbook), row=5, production_unit="m")
        else:
            plan.update(operation="programme", activity=activity, start=ANCHOR, finish=end, daily=daily,
                        source="Synthetic workfront programme", production_unit="m", snapshot=fingerprint(workbook))
        apply(journal, actor, [plan])
        if short != "PANEL":
            apply(journal, "Synthetic Programme Reviewer", [{"operation": "programme-verify",
                  "id": f"PLAN-{short}", "reference": f"SYN-PROGRAMME-{short}"}])

    moves = {"material_id": "SYN-REBAR", "unit": "kg", "location": LOCATION,
             "workfront": NORTH, "source": "Synthetic store note"}
    order = {"material_id": "SYN-REBAR", "unit": "kg", "location": LOCATION,
             "placed": "2026-01-10", "reference": "Synthetic order file"}
    apply(journal, "Synthetic Storekeeper", [
        {"operation": "order", "id": "ORDER-01", "po": "PO-SYN-01", "quantity": 30,
         "due": "2026-01-14", "eta": "2026-01-14", "confirmed_on": "2026-01-11", **order},
        {"operation": "order", "id": "ORDER-02", "po": "PO-SYN-02", "quantity": 25,
         "due": "2026-01-20", "eta": "2026-01-20", "confirmed_on": "2026-01-12", **order},
        {"operation": "arrival", "id": "DN-SYN-01", "order": "ORDER-01", "unit": "kg",
         "date": "2026-01-14", "reference": "Synthetic partial delivery note",
         "accepted": 8, "rejected": 2, "quarantined": 5},
        {"operation": "movement", "id": "RESERVE-01", "date": "2026-01-15",
         "kind": "reserve", "quantity": 20, **moves},
        {"operation": "movement", "id": "ISSUE-01", "date": "2026-01-15",
         "kind": "issue", "quantity": 12, **moves},
        {"operation": "movement", "id": "CONSUME-01", "date": "2026-01-15",
         "kind": "consume", "quantity": 9, **moves},
    ])
    return workbook, journal


def sample(directory):
    workbook, journal = fixtures(directory)
    ids = apply(journal, "Synthetic Buyer", [{"operation": "prepare-requisitions", "start": ANCHOR}])[0]
    identifier = ids[0]
    for action, actor in (("review", "Synthetic Materials Reviewer"),
                          ("approve", "Synthetic Requisition Approver"), ("issue", "Synthetic Buyer")):
        apply(journal, actor, [{"operation": "requisition-action", "id": identifier,
                                "action": action, "reference": f"SYN-MR-{action}"}])
    path = read(journal, export, id=identifier, directory=directory)
    return workbook, journal, path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory")
    args = parser.parse_args()
    workbook, journal, path = sample(Path(args.directory))
    print(f'python -m kvd.report_desktop --journal "{journal.path}" --workbook "{workbook}"')
    print(f"Material forecast start: {ANCHOR}")
    print(f"Issued requisition files: {path}")


if __name__ == "__main__":
    main()
