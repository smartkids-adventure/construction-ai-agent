"""Commit an approved batch to a new synthetic workbook revision."""
import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import hashlib
import json
import math
import os
import sqlite3

from openpyxl import load_workbook

from .journal import fingerprint
from .proposals import checksum, wire
from .workbook import FIRST_ROW, NEW_END, validate, values


WRITABLE = {
    5: set("ABCDEFGHIKLMPR"), 6: set("ABCDEFGHIJ"),
    15: set("ABCDEFGHIJKL"), 16: set("ABCDGHJ"), 17: set("ABCDEFG"),
    18: set("ABCE"), 19: set("ABCDEKP"), 20: set("JK"),
}
GUARDED = {
    (15, "K"): {"Pending"}, (15, "L"): {"Pending"},
    (16, "G"): {"Present"}, (16, "J"): {"Pending"},
}
DATES = {(5, "B"), (6, "B"), (15, "B"), (16, "C"), (17, "C"), (19, "B")}
NUMBERS = {(5, "K"), (5, "L"), (6, "I"), (6, "J")}


def setup(db):
    db.execute("""
        CREATE TABLE IF NOT EXISTS proposal_commits (
            token TEXT PRIMARY KEY, source_sha TEXT NOT NULL, source TEXT NOT NULL,
            output_sha TEXT NOT NULL, output TEXT NOT NULL, backup TEXT NOT NULL,
            actor TEXT NOT NULL, state TEXT NOT NULL
                CHECK(state IN ('pending','committed','aborted'))
        )
    """)
    db.execute("""CREATE UNIQUE INDEX IF NOT EXISTS proposal_successor
                  ON proposal_commits(source_sha) WHERE state <> 'aborted'""")
    db.execute("""CREATE UNIQUE INDEX IF NOT EXISTS proposal_destination
                  ON proposal_commits(output) WHERE state <> 'aborted'""")


def approved_payload(db, token):
    payload = approved = None
    for row in db.execute("""SELECT event, detail FROM events
                         WHERE event IN ('p2_proposal_previewed','p2_proposal_approved') ORDER BY id"""):
        item = json.loads(row["detail"])
        if item["token"] != token:
            continue
        if row["event"] == "p2_proposal_previewed":
            payload = item["payload"]
        else:
            approved = item
    if payload is None or approved is None or checksum(payload) != token:
        raise ValueError("Select an approved, unchanged proposal")
    for field in ("changes", "records", "workbook_sha"):
        if payload[field] != approved[field]:
            raise ValueError("Approval and preview differ")
    return payload, approved


def apply_cells(book, payload, approved):
    touched = set()
    verified_rows = set()
    for change in payload["changes"]:
        index = change["sheet_index"]
        match = re.fullmatch(r"([A-Z])([1-9][0-9]*)", change["cell"])
        if not match or index not in WRITABLE:
            raise ValueError("Proposed cell is outside the update scope")
        column, row = match[1], int(match[2])
        sheet = book.worksheets[index]
        target = sheet[change["cell"]]
        position = (index, column)
        if (sheet.title != change["sheet"] or not FIRST_ROW <= row <= NEW_END
                or column not in WRITABLE[index] or target.data_type == "f"
                or (index, target.coordinate) in touched):
            raise ValueError(f"Proposed cell is protected or repeated: {target.coordinate}")
        touched.add((index, target.coordinate))
        if wire(target.value) != change["before"]:
            raise ValueError(f"Before value changed at {target.coordinate}")
        value = change["after"]
        if position in GUARDED and value not in GUARDED[position]:
            raise ValueError("Proposed status is outside the permitted values")
        if position in DATES:
            if change["type"] != "date":
                raise ValueError("Expected an ISO date")
            value = date.fromisoformat(value)
        elif position in NUMBERS:
            if change["type"] != "number" or type(value) not in (int, float):
                raise ValueError("Expected a numeric quantity")
            if not math.isfinite(value) or value < 0:
                raise ValueError("Quantities must be finite and nonnegative")
            if position == (5, "L"):
                verified_rows.add((row, value))
        elif change["type"] != "text" or (value is not None and not isinstance(value, str)):
            raise ValueError("Expected a literal text value")
        target.value = value
        if isinstance(value, str):
            target.data_type = "s"
    if verified_rows:
        allowed = set()
        site = book.worksheets[5]
        for record in payload["records"]:
            verification = record.get("verification")
            if not verification or not verification["evidence"].strip():
                continue
            fields = record["fields"]
            for row, qty in verified_rows:
                if (qty == verification["quantity"] and site[f"R{row}"].value == record["workfront"]
                        and wire(site[f"B{row}"].value) == fields["date"]
                        and site[f"C{row}"].value == fields["shift"]
                        and site[f"D{row}"].value == fields["activity id"]
                        and site[f"I{row}"].value == fields["unit"]
                        and site[f"A{row}"].value == record.get("site_record_id")):
                    allowed.add((row, qty))
        if not approved["verification_attested"] or allowed != verified_rows:
            raise ValueError("Verified quantity requires matching separately approved evidence")
    return touched


def criterion(value):
    if value is None or value == "":
        return ""
    if not isinstance(value, str) or re.search(r"[*?~]|^[<>=]", value):
        raise ValueError("Activity, BQ and unit keys must be literal Excel criteria")
    return value.casefold()


def reconcile(book):
    """Independently calculate expected verified totals for the existing roll-ups."""
    source = book.worksheets[5]
    progress, bq = {}, {}
    for row in source.iter_rows(min_row=FIRST_ROW, max_row=NEW_END, values_only=True):
        qty = row[11]
        if qty is None:
            continue
        if type(qty) not in (int, float) or not math.isfinite(qty) or qty < 0:
            raise ValueError("Invalid verified quantity in synthetic workbook")
        for totals, key in ((progress, row[3]), (bq, row[4])):
            if key:
                unit = criterion(row[8])
                if not unit:
                    raise ValueError("Verified rows require a unit")
                pair = (criterion(key), unit)
                totals[pair] = totals.get(pair, 0) + qty
    summaries = {}
    for index, unit_index, column, totals in ((10, 4, "H", progress), (12, 2, "E", bq)):
        expected = {}
        matched = set()
        for row, cells in enumerate(book.worksheets[index].iter_rows(
                min_row=FIRST_ROW, max_row=NEW_END, values_only=True), FIRST_ROW):
            if cells[0] is None:
                continue
            pair = (criterion(cells[0]), criterion(cells[unit_index]))
            expected[f"{column}{row}"] = totals.get(pair)
            if pair in totals:
                matched.add(pair)
        summaries[str(index)] = {
            "expected": expected,
            "unmatched": [list(pair) + [totals[pair]] for pair in sorted(totals.keys() - matched)],
        }
    return summaries


def verify_result(staged, original, payload, touched, expected_totals):
    book = load_workbook(staged)
    try:
        validate(book, expanded=True)
        before_values = {key: dict(value) for key, value in original.items()}
        after_values = values(book)
        for change in payload["changes"]:
            name, address = change["sheet"], change["cell"]
            cell = book[name][address]
            if (wire(cell.value) != change["after"] or
                    (change["type"] == "text" and cell.value is not None and cell.data_type != "s")):
                raise ValueError(f"Round-trip value mismatch at {address}")
        for index, address in touched:
            name = book.worksheets[index].title
            before_values[name].pop(address, None)
            after_values[name].pop(address, None)
        if after_values != before_values or reconcile(book) != expected_totals:
            raise ValueError("Staged revision has unapproved data changes or inconsistent totals")
    finally:
        book.close()


def ensure_available(db, token, source_sha):
    rows = db.execute("""SELECT token FROM proposal_commits WHERE state <> 'aborted'
                       AND (token=? OR source_sha=?)""", (token, source_sha)).fetchall()
    if rows:
        raise ValueError("This batch or source revision already has a commit; recover or use its successor")


def commit(journal, token, output, actor):
    if not actor.strip():
        raise ValueError("A commit operator is required")
    output = Path(output).resolve()
    with journal.transaction() as db:
        setup(db)
        payload, approval = approved_payload(db, token)
        ensure_available(db, token, payload["workbook_sha"])
    source = Path(payload["workbook"]).resolve()
    if output == source or output.exists() or output.suffix.lower() != ".xlsx":
        raise ValueError("Choose a separate new .xlsx output file")
    if not payload["changes"]:
        raise ValueError("This approval contains no workbook changes")
    raw = source.read_bytes()
    if hashlib.sha256(raw).hexdigest() != payload["workbook_sha"]:
        raise ValueError("Approved source fingerprint changed")
    backups = journal.path.parent / "backups"
    backups.mkdir(exist_ok=True)
    backup = backups / f'{payload["workbook_sha"]}.xlsx'
    with TemporaryDirectory(dir=output.parent) as directory:
        staged = Path(directory) / "staged.xlsx"
        book = load_workbook(BytesIO(raw))
        try:
            validate(book, expanded=True)
            original = values(book)
            before = reconcile(book)
            touched = apply_cells(book, payload, approval)
            after = reconcile(book)
            book.save(staged)
        finally:
            book.close()
        verify_result(staged, original, payload, touched, after)
        output_sha = fingerprint(staged)
        with journal.transaction() as db:
            ensure_available(db, token, payload["workbook_sha"])
            if not backup.exists():
                with TemporaryDirectory(dir=backups) as backup_dir:
                    candidate = Path(backup_dir) / "snapshot.xlsx"
                    candidate.write_bytes(raw)
                    os.link(candidate, backup)
            if fingerprint(backup) != payload["workbook_sha"]:
                raise ValueError("Backup fingerprint mismatch")
            if output.exists() or fingerprint(source) != payload["workbook_sha"]:
                raise ValueError("Source or output changed during staging")
            db.execute("DELETE FROM proposal_commits WHERE token=? AND state='aborted'", (token,))
            try:
                db.execute("INSERT INTO proposal_commits VALUES (?,?,?,?,?,?,?,'pending')",
                           (token, payload["workbook_sha"], str(source), output_sha,
                            str(output), str(backup), actor))
            except sqlite3.IntegrityError as error:
                raise ValueError("Destination reserved by another batch") from error
            journal.log(db, "p21_commit_prepared", actor, {
                "token": token, "source_sha": payload["workbook_sha"], "output_sha": output_sha,
                "backup": str(backup), "output": str(output), "changes": payload["changes"],
                "reconciliation_before": before, "reconciliation_after": after,
            })
        with journal.transaction() as db:
            prepared = db.execute(
                "SELECT state, output, output_sha FROM proposal_commits WHERE token=?", (token,)
            ).fetchone()
            if tuple(prepared) != ("pending", str(output), output_sha):
                raise ValueError("Commit was reconciled while publishing")
            if fingerprint(source) != payload["workbook_sha"] or fingerprint(backup) != payload["workbook_sha"]:
                raise ValueError("Source or backup fingerprint changed before publication")
            os.link(staged, output)
            db.execute("UPDATE proposal_commits SET state='committed' WHERE token=?", (token,))
            journal.log(db, "p21_commit_completed", actor, {"token": token, "output_sha": output_sha})
    return {"output": str(output), "output_sha": output_sha, "backup": str(backup), "reconciliation": after}


def recover(journal, actor):
    if not actor.strip():
        raise ValueError("A recovery operator is required")
    outcomes = []
    with journal.transaction() as db:
        setup(db)
        for row in db.execute("SELECT * FROM proposal_commits WHERE state='pending'").fetchall():
            output, backup = Path(row["output"]), Path(row["backup"])
            if (not backup.is_file() or fingerprint(backup) != row["source_sha"]
                    or (output.exists() and fingerprint(output) != row["output_sha"])):
                outcomes.append((row["token"], "conflict — retained for manual review"))
                continue
            state = "committed" if output.exists() else "aborted"
            db.execute("UPDATE proposal_commits SET state=? WHERE token=?", (state, row["token"]))
            journal.log(db, "p21_commit_recovered", actor, {"token": row["token"], "state": state})
            outcomes.append((row["token"], state))
    return outcomes

def main():
    import argparse
    from .journal import Journal

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--journal", default="local/audit.sqlite3")
    sub = parser.add_subparsers(dest="command", required=True)
    write = sub.add_parser("commit")
    write.add_argument("token")
    write.add_argument("output")
    write.add_argument("--actor", required=True)
    sub.add_parser("recover").add_argument("--actor", required=True)
    args = parser.parse_args()
    try:
        journal = Journal(args.journal)
        if args.command == "commit":
            result = commit(journal, args.token, args.output, args.actor)
        else:
            result = recover(journal, args.actor)
        print(json.dumps(result, ensure_ascii=True, indent=2))
    except (ValueError, OSError, sqlite3.Error) as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
