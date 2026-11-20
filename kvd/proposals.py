"""Build and approve exact cell proposals for an expanded synthetic workbook."""
import hashlib
import json
import math
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter

from .intake import check
from .journal import fingerprint
from .workbook import FIRST_ROW, NEW_END, validate
from .workfront import normalize, resolve

SITE, UTILITY, RELOCATION, READINESS, TASK, TOOLBOX, ATTENDANCE = 5, 17, 18, 19, 20, 15, 16
RESOURCES = 6
ADDABLE = {"MC01", "MC14"}
ACTIONS = {"add", "keep distinct", "suppress", "correct existing"}


def checksum(obj):
    encoded = json.dumps(obj, ensure_ascii=False, sort_keys=True, allow_nan=False).encode()
    return hashlib.sha256(encoded).hexdigest()


def wire(value):
    if isinstance(value, (date, datetime)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    return value


def number(raw):
    value = float(raw)
    if not math.isfinite(value) or value < 0:
        raise ValueError("Quantity must be finite and nonnegative")
    return value


def key(kind, *parts):
    return [kind, *(normalize(str(wire(part) or "")) for part in parts)]


class Planner:
    def __init__(self, book):
        self.book = book
        self.changes = []
        self.reserved = set()
        self.approved_cells = set()

    def rows(self, index):
        return self.book.worksheets[index].iter_rows(min_row=FIRST_ROW, max_row=NEW_END,
                                                     values_only=True)

    def lookup(self, index, identifier, required=True):
        matches = [row for row, values in enumerate(self.rows(index), FIRST_ROW)
                   if normalize(str(values[0] or "")) == normalize(identifier)]
        if len(matches) > 1 or (required and not matches):
            raise ValueError(f"Expected one workbook row for {identifier}")
        return matches[0] if matches else None

    def new(self, index, data):
        sheet = self.book.worksheets[index]
        for row in range(FIRST_ROW, NEW_END + 1):
            if (index, row) not in self.reserved and all(
                cell.value is None or cell.data_type == "f" for cell in sheet[row]
            ):
                self.reserved.add((index, row))
                self.write(index, row, data)
                return
        raise ValueError(f"{sheet.title} has no unused prepared row")

    def write(self, index, row, data):
        sheet = self.book.worksheets[index]
        for column, value in data.items():
            cell = sheet[f"{column}{row}"]
            if cell.data_type == "f":
                raise ValueError(f"Formula cell is outside the write scope: {cell.coordinate}")
            before = wire(cell.value)
            if before == value:
                continue
            if (index, cell.coordinate) in self.approved_cells or any(
                change["sheet_index"] == index and change["cell"] == cell.coordinate
                for change in self.changes
            ):
                raise ValueError("Another draft or approved proposal targets the same cell")
            self.changes.append({"sheet_index": index, "sheet": sheet.title,
                                 "cell": cell.coordinate, "before": before, "after": value,
                                 "type": "date" if column in DATE_COLUMNS.get(index, ())
                                 else "number" if isinstance(value, (int, float)) else "text"})


DATE_COLUMNS = {
    SITE: {"B"}, RESOURCES: {"B"}, TOOLBOX: {"B"}, ATTENDANCE: {"C"},
    UTILITY: {"C"}, READINESS: {"B"},
}


def identity(command, fields, workfront):
    common = (workfront, fields["date"], fields["shift"])
    if command in ADDABLE:
        return key("site", *common, fields["activity id"], fields.get("bq item"),
                   fields.get("crew"), fields["unit"])
    if command == "MC11":
        return key("toolbox", *common, fields["activity"], fields["lead"])
    return key("readiness", *common, fields["activity"])


def workbook_matches(planner, match_key, fields=None):
    kind = match_key[0]
    if kind == "site":
        index, columns = SITE, (18, 2, 3, 4, 5, 8, 9)
    elif kind == "toolbox":
        index, columns = TOOLBOX, (4, 2, 3, 6, 7)
    else:
        index, columns = READINESS, (1, 2, 3, 5)
    matches = []
    for row, values in enumerate(planner.rows(index), FIRST_ROW):
        if not values[0]:
            continue
        observed = key(kind, *(values[column - 1] for column in columns))
        if observed == match_key:
            matches.append({"row": row, "record": str(values[0]),
                            "reported": wire(values[10]) if kind == "site" else None,
                            "verified": wire(values[11]) if kind == "site" else None})
    if kind == "readiness" and fields:
        row = planner.lookup(UTILITY, fields["utility id"], required=False)
        if row:
            matches.append({"row": None, "record": fields["utility id"], "reason": "existing utility ID"})
    return matches


def subject(draft, front, day, digest, verification):
    fields = draft.fields
    basic = (front["id"], front["location"], day, fields["shift"], fields["activity"], fields["evidence"])
    reference = "SYN-" + digest[:16].upper()
    if draft.command in ADDABLE:
        if front["activities"] and fields["activity id"] not in front["activities"]:
            raise ValueError("Activity ID is outside the approved synthetic workfront index")
        data = {"A": reference, "B": day, "C": fields["shift"], "D": fields["activity id"],
                "E": fields.get("bq item"), "F": front["location"], "G": fields["activity"],
                "H": fields.get("crew"), "I": fields["unit"], "K": number(fields["reported qty"]),
                "M": fields.get("notes"), "P": fields["evidence"], "R": front["id"]}
        if verification:
            data["L"] = number(verification["quantity"])
        return data
    return basic, reference


def stage_rows(planner, draft, front, day, digest, verification, target):
    fields = draft.fields
    prepared = subject(draft, front, day, digest, verification)
    if draft.command in ADDABLE:
        if target is None:
            planner.new(SITE, prepared)
        else:
            planner.write(SITE, target, {column: prepared[column]
                                         for column in ("G", "K", "M", "P", "L") if column in prepared})
        if fields.get("resources"):
            if target is not None:
                raise ValueError("Resource corrections require separate review")
            for index, resource in enumerate(fields["resources"].split(";"), 1):
                parts = [part.strip() for part in resource.split("|")]
                if len(parts) != 4 or not all(parts):
                    raise ValueError("Use resources: type | description | count | hours per unit")
                kind, description, count, hours = parts
                planner.new(RESOURCES, {"A": f"{prepared['A']}-R{index}", "B": day,
                    "C": fields["shift"], "D": prepared["A"], "E": front["location"],
                    "F": fields.get("crew"), "G": kind, "H": description, "I": number(count),
                    "J": number(hours)})
        if draft.command == "MC14":
            task = planner.lookup(TASK, fields["task id"])
            cell = planner.book.worksheets[TASK][f"C{task}"]
            if cell.value != front["id"]:
                raise ValueError("Task belongs to a different workfront")
            planner.write(TASK, task, {"J": fields["evidence"], "K": fields["outcome"]})
        return
    (workfront, location, day, shift, activity, evidence), reference = prepared
    if draft.command == "MC11":
        workers = [worker.strip() for worker in fields["attendees"].split(";")]
        if any(not worker for worker in workers) or len({normalize(w) for w in workers}) != len(workers):
            raise ValueError("Provide distinct attendee IDs separated by semicolons")
        planner.new(TOOLBOX, {"A": reference, "B": day, "C": shift, "D": workfront,
                             "E": location, "F": activity, "G": fields["lead"], "H": fields["topics"],
                             "I": fields.get("jsa ref"), "J": evidence, "K": "Pending", "L": "Pending"})
        for i, worker in enumerate(workers, 1):
            planner.new(ATTENDANCE, {"A": f"{reference}-{i}", "B": reference, "C": day,
                                     "D": worker, "G": "Present", "H": evidence, "J": "Pending"})
        return
    if fields["utility type"] not in {"Electrical", "Data / telecom", "Other", "Unknown"}:
        raise ValueError("Choose a valid utility type")
    prior_utility = planner.lookup(UTILITY, fields["utility id"], required=False)
    if prior_utility:
        sheet = planner.book.worksheets[UTILITY]
        if sheet[f"B{prior_utility}"].value != workfront:
            raise ValueError("Utility ID belongs to a different workfront")
    else:
        planner.new(UTILITY, {"A": fields["utility id"], "B": workfront, "C": day,
                              "D": location, "E": fields["utility type"], "F": fields.get("owner"),
                              "G": evidence, "H": "Conflict" if fields.get("blocker") else "Unknown",
                              "L": "Blocked" if fields.get("blocker") else "Pending"})
    if fields.get("relocation instruction"):
        planner.new(RELOCATION, {"A": reference, "B": fields["utility id"], "C": workfront,
                                 "E": fields["relocation instruction"], "H": "Not started"})
    planner.new(READINESS, {"A": workfront, "B": day, "C": shift, "D": location, "E": activity,
                           "F": "Blocked" if fields.get("blocker") else "Pending",
                           "H": "Pending", "I": "Pending", "J": "Pending",
                           "K": fields.get("blocker"), "L": "Not released", "P": evidence})


def approved_state(db):
    entries = db.execute(
        "SELECT id, detail FROM events WHERE event='p2_proposal_approved' ORDER BY id"
    ).fetchall()
    return (entries[-1]["id"] if entries else 0,
            [json.loads(entry["detail"]) for entry in entries])


def preview(journal, workbook, drafts, decisions, actor):
    if not drafts or len(drafts) != len(decisions):
        raise ValueError("Provide a review decision for each draft")
    path = Path(workbook).resolve()
    raw = path.read_bytes()
    sha = hashlib.sha256(raw).hexdigest()
    book = load_workbook(BytesIO(raw))
    try:
        validate(book, expanded=True)
        return build_preview(journal, Planner(book), path, sha, drafts, decisions, actor)
    finally:
        book.close()


def build_preview(journal, planner, path, sha, drafts, decisions, actor):
    with journal.transaction() as db:
        epoch, history = approved_state(db)
    for proposal in history:
        if proposal["workbook_sha"] != sha:
            continue
        for change in proposal["changes"]:
            index, address = change["sheet_index"], change["cell"]
            planner.approved_cells.add((index, address))
            planner.reserved.add((index, planner.book.worksheets[index][address].row))
    seen = set()
    seen_pastes = set()
    used_utility = {
        normalize(record["fields"]["utility id"])
        for proposal in history if proposal["workbook_sha"] == sha
        for record in proposal["records"]
        if record["command"] == "MC12" and record["decision"] != "suppress"
    }
    results = []
    replay_ids = {record["source_hash"] for result in history for record in result["records"]
                  if record["decision"] != "suppress"}
    prior_records = [record for result in history for record in result["records"]
                     if record["decision"] != "suppress"]
    for draft, decision in zip(drafts, decisions):
        action = decision.get("action", "add")
        reason = decision.get("reason", "").strip()
        if action not in ACTIONS or (action != "add" and not reason):
            raise ValueError("Choose a duplicate decision and give a reason for an override")
        source_hash = checksum(normalize(draft.source))
        item = {"command": draft.command, "source": draft.source, "source_hash": source_hash,
                "fields": dict(draft.fields), "decision": action, "reason": reason}
        if action == "suppress":
            results.append(item)
            continue
        day = check(draft)
        fields = draft.fields
        front = resolve(journal, fields["workfront"])
        match_key = identity(draft.command, fields, front["id"])
        candidates = workbook_matches(planner, match_key, fields)
        for record in prior_records:
            if record.get("match_key") == match_key:
                candidates.append({"row": None, "record": record["source_hash"][:12],
                                   "reason": "previously approved proposal",
                                   "reported": record["fields"].get("reported qty")})
        if tuple(match_key) in seen:
            candidates.append({"row": None, "record": "another draft in this paste"})
        if source_hash in replay_ids or source_hash in seen_pastes:
            raise ValueError("Exact source replay: suppress this draft or paste the corrected source")
        if candidates and action == "add":
            labels = ", ".join(f'{c["record"]} (reported {c.get("reported")}, verified {c.get("verified")})'
                               for c in candidates)
            raise ValueError(f"Possible duplicate for {front['id']}: {labels}. Choose a review decision.")
        target = None
        if action == "correct existing":
            record_id = decision.get("record_id", "").strip()
            target = next((c["row"] for c in candidates if c["record"] == record_id), None)
            if draft.command not in ADDABLE or target is None:
                raise ValueError("Corrections require a matched site record ID in this workbook")
        verification = None
        if decision.get("verified_qty", "").strip() or decision.get("verification_evidence", "").strip():
            evidence = decision.get("verification_evidence", "").strip()
            if draft.command not in ADDABLE or not evidence or not decision.get("verified_qty", "").strip():
                raise ValueError("A site verification proposal requires numeric quantity and separate evidence")
            verification = {"quantity": number(decision["verified_qty"]), "evidence": evidence}
        if draft.command == "MC12":
            identifier = normalize(fields["utility id"])
            if identifier in used_utility:
                raise ValueError("Only one draft per utility ID is supported in a review batch")
            used_utility.add(identifier)
        digest = checksum([source_hash, fields, match_key, len(results)])
        stage_rows(planner, draft, front, day, digest, verification, target)
        item.update({"match_key": match_key, "workfront": front["id"], "duplicates": candidates,
                     "verification": verification, "target_record": decision.get("record_id") if target else None})
        results.append(item)
        seen.add(tuple(match_key))
        seen_pastes.add(source_hash)
    payload = {"workbook": str(path), "workbook_sha": sha, "records": results,
               "changes": planner.changes, "epoch": epoch, "prepared_by": actor}
    token = checksum(payload)
    with journal.transaction() as db:
        journal.log(db, "p2_proposal_previewed", actor, {"token": token, "payload": payload})
    return token, payload


def approve(journal, token, actor, verification_attested=False):
    with journal.transaction() as db:
        previews = db.execute(
            "SELECT detail FROM events WHERE event='p2_proposal_previewed' ORDER BY id DESC"
        )
        payload = None
        for row in previews:
            candidate = json.loads(row["detail"])
            if candidate["token"] == token:
                payload = candidate["payload"]
                break
        if payload is None or checksum(payload) != token:
            raise ValueError("Preview token is missing or changed")
        epoch, history = approved_state(db)
        if any(item["token"] == token for item in history):
            raise ValueError("This proposal was already approved")
        if epoch != payload["epoch"] or fingerprint(payload["workbook"]) != payload["workbook_sha"]:
            raise ValueError("Workbook or approval history changed; refresh the preview")
        if any(item.get("verification") for item in payload["records"]):
            if not verification_attested or normalize(actor) == normalize(payload["prepared_by"]):
                raise ValueError("A distinct quantity reviewer must attest to the evidence")
        detail = {"token": token, "records": payload["records"], "changes": payload["changes"],
                  "workbook_sha": payload["workbook_sha"], "verification_attested": verification_attested}
        journal.log(db, "p2_proposal_approved", actor, detail)
        return detail
