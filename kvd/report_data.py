"""Read reporting data from fingerprint-verified synthetic workbook commits."""
from collections import defaultdict
from datetime import date
from io import BytesIO
from pathlib import Path
import hashlib
import math
import sqlite3

from openpyxl import load_workbook

from .proposals import wire
from .workbook import FIRST_ROW, NEW_END, validate


SECTIONS = {
    "progress": "Workfront progress",
    "resources": "Manpower and machinery",
    "technical": "Open technical issues",
    "quality": "Open QA/QC issues",
    "toolbox": "Safety toolbox and attendance",
    "utility": "Utility and relocation status",
    "readiness": "Workfront readiness",
    "tasks": "Junior engineer tasks",
    "actions": "Outstanding actions",
}
FILTERS = ("from_date", "to_date", "workfront", "activity", "subcontractor", "status")


def approved_snapshot(journal, path):
    path = Path(path).resolve()
    raw = path.read_bytes()
    digest = hashlib.sha256(raw).hexdigest()
    with journal.transaction() as db:
        try:
            rows = db.execute(
                "SELECT token, output FROM proposal_commits WHERE output_sha=? AND state='committed'",
                (digest,),
            ).fetchall()
        except sqlite3.OperationalError as error:
            if "no such table" not in str(error):
                raise
            rows = []
        commit = next((row for row in rows if Path(row["output"]).resolve() == path), None)
        if commit is None:
            raise ValueError("Select a committed synthetic revision recorded at this path in the journal.")
    book = load_workbook(BytesIO(raw))
    try:
        validate(book, expanded=True)
        return {"fingerprint": digest, "commit": commit["token"], "file": path.name,
                "path": str(path),
                "sections": extract(book)}
    finally:
        book.close()


def text(value):
    value = wire(value)
    return "" if value is None else str(value)


def quantity(value):
    if value is None:
        return None
    if type(value) not in (int, float) or not math.isfinite(value):
        raise ValueError("Quantity cells must contain numbers or be blank.")
    return value


def iso(value):
    if not value:
        return ""
    try:
        result = date.fromisoformat(text(value))
    except ValueError as error:
        raise ValueError(f"Invalid report date: {value}") from error
    if result.isoformat() != text(value):
        raise ValueError("Dates must use YYYY-MM-DD.")
    return result.isoformat()


def norm(value):
    return text(value).strip().casefold()


def claim_status(row):
    values = {key: norm(row[index]) for key, index in
              (("utility", 5), ("drawing", 7), ("access", 8), ("toolbox", 9))}
    if "blocked" in values.values():
        return "Blocked"
    if (all(status == "verified" for status in values.values()) and row[6] and
            norm(row[10]) == "none" and norm(row[11]) == "released" and row[12] and row[13]):
        return "Release recorded"
    return "Pending review"


def extract(book):
    result = {key: [] for key in SECTIONS}

    def rows(index):
        sheet = book.worksheets[index]
        for number, raw in enumerate(sheet.iter_rows(min_row=FIRST_ROW, max_row=NEW_END,
                                                      values_only=True), FIRST_ROW):
            if raw[0] is not None:
                yield number, raw

    def record(section, index, number, raw, *, title, status, day=None, workfront=None,
               activity=None, subcontractor=None, details=(), **data):
        item = {"id": text(raw[0]), "sheet": index + 1, "row": number, "day": iso(day),
                "title": text(title), "status": text(status) or "Not recorded",
                "workfront": text(workfront), "activity": text(activity),
                "subcontractor": text(subcontractor),
                "details": [(label, text(value)) for label, value in details if text(value)]}
        item.update(data)
        result[section].append(item)
        return item

    site = defaultdict(list)
    for n, r in rows(5):
        planned, reported, verified = [quantity(value) for value in r[9:12]]
        item = record("progress", 5, n, r, title=r[6], status="Unverified" if verified is None else "Verified",
                      day=r[1], workfront=r[17], activity=r[3], subcontractor=r[7],
                      details=(("Shift", r[2]), ("Location", r[5]), ("Evidence", r[15]), ("Constraints", r[12])),
                      unit=text(r[8]), planned=planned, reported=reported, verified=verified)
        site[norm(r[0])].append(item)

    def parent(index, identifier):
        matches = index.get(norm(identifier), [])
        return matches[0] if len(matches) == 1 else {}

    for n, r in rows(6):
        linked = parent(site, r[3])
        record("resources", 6, n, r, title=r[7], status=r[6], day=r[1],
               workfront=linked.get("workfront"), activity=linked.get("activity"),
               subcontractor=r[5] or linked.get("subcontractor"),
               details=(("Site record", r[3]), ("Count", r[8]), ("Hours per unit", r[9])))

    for n, r in rows(7):
        if norm(r[10]) not in {"closed", "accepted", "rejected"}:
            record("technical", 7, n, r, title=r[4], status=r[10], day=r[2],
                   details=(("Owner", r[8]), ("Due", r[9]), ("Action", r[12])))
    for n, r in rows(14):
        if norm(r[7]) not in {"closed", "accepted", "rejected"}:
            record("technical", 14, n, r, title=r[1], status=r[7], day=r[4],
                   details=(("Owner", r[5]), ("Due", r[6]), ("Action", r[9])))
    for n, r in rows(8):
        if norm(r[8]) not in {"closed", "accepted"}:
            record("quality", 8, n, r, title=r[7], status=r[8], day=r[1], activity=r[2],
                   details=(("Inspector", r[10]), ("Result", r[9]), ("Action", r[12])))
    for n, r in rows(9):
        if norm(r[7]) != "closed":
            record("quality", 9, n, r, title=r[3], status=r[7], day=r[1],
                   details=(("Owner", r[5]), ("Due", r[6]), ("Action", r[8])))

    toolbox = defaultdict(list)
    for n, r in rows(15):
        item = record("toolbox", 15, n, r, title=r[7] or "Toolbox record", status=r[11], day=r[1],
                      workfront=r[3], activity=r[5], details=(("Lead", r[6]),
                      ("Attendance evidence", r[10]), ("Evidence", r[9]), ("Follow-up", r[12])))
        toolbox[norm(r[0])].append(item)
    for n, r in rows(16):
        linked = parent(toolbox, r[1])
        record("toolbox", 16, n, r, title="Attendance record", status=r[9], day=r[2],
               workfront=linked.get("workfront"), activity=linked.get("activity"),
               details=(("Toolbox record", r[1]), ("Attendance", r[6]), ("Evidence", r[7])))

    for n, r in rows(17):
        record("utility", 17, n, r, title=r[4], status=r[11], day=r[2], workfront=r[1],
               details=(("Conflict", r[7]), ("Clearance reference", r[12]),
                        ("Owner", r[5]), ("Required action", r[8])))
    for n, r in rows(18):
        record("utility", 18, n, r, title="Relocation record", status=r[7], day=r[6], workfront=r[2],
               details=(("Utility", r[1]), ("Target", r[5]), ("Instruction", r[4]),
                        ("Clearance reference", r[9]), ("Outstanding action", r[10])))
    for n, r in rows(19):
        record("readiness", 19, n, r, title=r[4], status=claim_status(r), day=r[1], workfront=r[0],
               activity=r[4], details=(("Reported constraint", r[10]),
               ("Recorded release status", r[11]), ("Release evidence", r[12]), ("Follow-up", r[15])))
    for n, r in rows(20):
        record("tasks", 20, n, r, title=r[4], status=r[8], day=r[1], workfront=r[2],
               details=(("Assignee", r[5]), ("Due", r[7]), ("Priority", r[6]),
                        ("Reported outcome", r[10]), ("Senior review", r[11]), ("Follow-up", r[12])))
    for n, r in rows(2):
        if norm(r[7]) not in {"closed", "accepted", "rejected"}:
            record("actions", 2, n, r, title=r[3], status=r[7], day=r[1],
                   details=(("Owner", r[5]), ("Due", r[6]), ("Priority", r[4]), ("Next action", r[9])))
    return result


def select(snapshot, **filters):
    if filters.keys() - set(FILTERS):
        raise ValueError("Unknown dashboard filter")
    start, end = iso(filters.get("from_date")), iso(filters.get("to_date"))
    if start and end and start > end:
        raise ValueError("Start date must be on or before end date.")
    dimensions = {field: norm(filters.get(field)) for field in FILTERS[2:] if filters.get(field)}

    def matches(row):
        day = row["day"]
        if start and (not day or day < start):
            return False
        if end and (not day or day > end):
            return False
        return all(norm(row[field]) == selected for field, selected in dimensions.items())

    return {section: [row for row in records if matches(row)]
            for section, records in snapshot["sections"].items()}


def comparison(records):
    groups = defaultdict(list)
    for row in records:
        # A missing unit or activity cannot support a meaningful cross-record total.
        dimensions = tuple(row[field] for field in ("workfront", "activity", "subcontractor", "unit"))
        key = dimensions + ((row["sheet"], row["row"]) if not row["unit"] or not row["activity"] else ())
        groups[key].append(row)
    result = []
    for key, rows in sorted(groups.items()):
        counts, sums = {}, {}
        for field in ("planned", "reported", "verified"):
            known = [row[field] for row in rows if row[field] is not None]
            counts[field], sums[field] = len(known), sum(known) if known else None
        result.append({"dimensions": list(key[:4]), "records": len(rows),
                       "counts": counts, "sums": sums,
                       "sources": [f'{row["id"]} (row {row["row"]})' for row in rows]})
    return result


def display_quantity(value, count, records):
    if value is None:
        return "Not recorded"
    label = f"{value:,.10g}"
    return label if count == records else f"{label} ({count}/{records})"
