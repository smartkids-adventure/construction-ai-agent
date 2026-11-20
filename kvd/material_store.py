"""P4 material register, drawing evidence, programme and inventory ledger."""
import ast
from datetime import date, timedelta
from decimal import Decimal, DecimalException, ROUND_CEILING
from io import BytesIO
import json
import re

from openpyxl import load_workbook

from .journal import fingerprint
from .workbook import FIRST_ROW, NEW_END, validate
from .proposals import wire
from .workfront import normalize


SCALE = 1_000_000
LIMIT = 10**14
SCHEMA = """
CREATE TABLE IF NOT EXISTS p4_schema (version INTEGER PRIMARY KEY);
INSERT OR IGNORE INTO p4_schema VALUES (1);
CREATE TABLE IF NOT EXISTS p4_material (
    id TEXT PRIMARY KEY, description TEXT NOT NULL, specification TEXT NOT NULL, unit TEXT NOT NULL,
    lead INTEGER, safety INTEGER, variance TEXT
);
CREATE TABLE IF NOT EXISTS p4_drawing (
    id INTEGER PRIMARY KEY, number TEXT NOT NULL, revision TEXT NOT NULL,
    status TEXT NOT NULL, reference TEXT NOT NULL, snapshot TEXT,
    UNIQUE(number,revision)
);
CREATE TABLE IF NOT EXISTS p4_takeoff (
    id TEXT PRIMARY KEY, drawing INTEGER NOT NULL REFERENCES p4_drawing(id),
    workfront TEXT NOT NULL, material TEXT NOT NULL REFERENCES p4_material(id), element TEXT NOT NULL,
    dimensions TEXT NOT NULL, formula TEXT NOT NULL, source TEXT NOT NULL, unit TEXT NOT NULL,
    design INTEGER NOT NULL, allowance INTEGER NOT NULL, procurement INTEGER NOT NULL,
    state TEXT NOT NULL, preparer TEXT NOT NULL, verifier TEXT,
    UNIQUE(drawing,workfront,material,element)
);
CREATE TABLE IF NOT EXISTS p4_plan (
    id TEXT PRIMARY KEY, takeoff TEXT NOT NULL REFERENCES p4_takeoff(id),
    material TEXT NOT NULL REFERENCES p4_material(id), workfront TEXT NOT NULL,
    location TEXT NOT NULL, activity TEXT NOT NULL, start TEXT NOT NULL, finish TEXT NOT NULL,
    weekdays TEXT NOT NULL, exclusions TEXT NOT NULL, daily INTEGER,
    rate INTEGER, production_unit TEXT NOT NULL, source TEXT NOT NULL, snapshot TEXT,
    preparer TEXT NOT NULL, reviewer TEXT, active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS p4_movement (
    id TEXT PRIMARY KEY, material TEXT NOT NULL REFERENCES p4_material(id),
    location TEXT NOT NULL, day TEXT NOT NULL, kind TEXT NOT NULL, amount INTEGER NOT NULL,
    workfront TEXT, usable INTEGER NOT NULL, quarantine INTEGER NOT NULL,
    reserved INTEGER NOT NULL, source TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS p4_moves_by_store ON p4_movement(material,location,day);
CREATE TABLE IF NOT EXISTS p4_order (
    id TEXT PRIMARY KEY, material TEXT NOT NULL REFERENCES p4_material(id),
    location TEXT NOT NULL, po TEXT NOT NULL, requisition TEXT, workfront TEXT,
    ordered INTEGER NOT NULL, placed TEXT NOT NULL, due TEXT NOT NULL, eta TEXT,
    confirmed_on TEXT, reference TEXT NOT NULL, UNIQUE(po,material,location)
);
CREATE TABLE IF NOT EXISTS p4_arrival (
    id TEXT PRIMARY KEY, order_id TEXT NOT NULL REFERENCES p4_order(id),
    day TEXT NOT NULL, accepted INTEGER NOT NULL, rejected INTEGER NOT NULL,
    quarantined INTEGER NOT NULL, reference TEXT NOT NULL, UNIQUE(order_id,reference)
);
CREATE TABLE IF NOT EXISTS p4_requisition (
    id INTEGER PRIMARY KEY, scope TEXT NOT NULL, revision INTEGER NOT NULL,
    state TEXT NOT NULL, body TEXT NOT NULL, digest TEXT NOT NULL,
    author TEXT NOT NULL, reviewer TEXT, approver TEXT, closed TEXT, word BLOB, pdf BLOB,
    word_sha TEXT, pdf_sha TEXT, UNIQUE(scope,revision)
);
"""


def setup(db):
    for statement in SCHEMA.split(";"):
        if statement.strip():
            db.execute(statement)
    if db.execute("SELECT MAX(version) FROM p4_schema").fetchone()[0] != 1:
        raise ValueError("This application cannot open the newer P4 schema.")


def required(value, label):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{label} is required.")
    return value.strip()


def day(value):
    raw = wire(value)
    result = date.fromisoformat(required(raw, "Date"))
    if result.isoformat() != raw:
        raise ValueError("Use YYYY-MM-DD.")
    return result.isoformat()


def decimal(value):
    try:
        parsed = Decimal(str(value))
    except DecimalException as error:
        raise ValueError("Enter a finite decimal number.") from error
    if not parsed.is_finite() or abs(parsed) > LIMIT:
        raise ValueError("Number is outside the supported range.")
    return parsed


def units(value, *, signed=False):
    qty = decimal(value)
    if not signed and qty < 0:
        raise ValueError("Quantity must be nonnegative.")
    result = qty * SCALE
    if result != result.to_integral_value() or abs(result) > LIMIT:
        raise ValueError("Use at most six decimal places and a supported quantity.")
    return int(result)


def show(amount):
    return None if amount is None else format((Decimal(amount) / SCALE).normalize(), "f")


def real_amount(amount):
    qty = units(amount)
    if not qty:
        raise ValueError("Quantity must be greater than zero.")
    return qty


def found(db, table, identifier):
    row = db.execute(f"SELECT * FROM {table} WHERE id=?", (identifier,)).fetchone()
    if row is None:
        raise ValueError(f"Unknown record: {identifier}")
    return row


def material(db, identifier, unit):
    row = found(db, "p4_material", identifier)
    if unit != row["unit"]:
        raise ValueError(f'Unit mismatch. Expected {row["unit"]}.')
    return row


def front(db, identifier, activity=None, *, active=True):
    identifier = required(identifier, "Workfront ID")
    match = re.fullmatch(r"KVDT2-WF-(\d{6})", identifier)
    row = db.execute("SELECT * FROM workfronts WHERE sequence=?", (int(match[1]),)).fetchone() if match else None
    if row is None or (active and row["retired"]):
        raise ValueError("Use an existing active canonical Workfront ID.")
    approved = json.loads(row["activities"])
    if activity and approved and activity not in approved:
        raise ValueError("Activity does not belong to the workfront index.")
    return identifier


def log(db, journal, event, actor, detail):
    journal.log(db, f"p4_{event}", actor, detail)


def material_add(db, journal, actor, id, description, specification, unit, lead=None, safety=None, variance=None):
    id, description, specification, unit = [
        required(value, label) for value, label in
        ((id, "Material ID"), (description, "Description"), (specification, "Specification"), (unit, "Unit"))]
    lead, safety, variance = policy(lead, safety, variance)
    db.execute("INSERT INTO p4_material VALUES (?,?,?,?,?,?,?)",
               (id, description, specification, unit, lead, safety, str(variance) if variance is not None else None))
    log(db, journal, "material_added", actor, {"id": id, "unit": unit, "lead": lead, "safety": show(safety)})
    return id


def policy(lead, safety, variance):
    if lead is not None and (type(lead) is not int or not 0 <= lead <= 3650):
        raise ValueError("Lead time must be a whole number of calendar days, 0 to 3650.")
    safety = units(safety) if safety is not None else None
    variance = decimal(variance) if variance is not None else None
    if variance is not None and variance < 0:
        raise ValueError("Variance threshold must be nonnegative.")
    return lead, safety, variance


def material_policy(db, journal, actor, id, lead, safety, variance, reference):
    previous = found(db, "p4_material", id)
    lead, safety, variance = policy(lead, safety, variance)
    threshold = str(variance) if variance is not None else None
    db.execute("UPDATE p4_material SET lead=?,safety=?,variance=? WHERE id=?",
               (lead, safety, threshold, id))
    log(db, journal, "material_policy_updated", actor, {
        "id": id, "before": [previous["lead"], show(previous["safety"]), previous["variance"]],
        "after": [lead, show(safety), threshold], "reference": required(reference, "Policy evidence"),
    })
    return id


def snapshot_ref(db, digest):
    if not digest:
        return None
    exists = db.execute("SELECT name FROM sqlite_master WHERE name='proposal_commits'").fetchone()
    if not exists or not db.execute(
        "SELECT 1 FROM proposal_commits WHERE state='committed' AND output_sha=?", (digest,)
    ).fetchone():
        raise ValueError("Snapshot must refer to a committed workbook in this journal.")
    return digest


def drawing_add(db, journal, actor, number, revision, status, reference, snapshot=None):
    number = required(number, "Drawing number")
    revision = required(revision, "Revision")
    reference = required(reference, "Drawing reference")
    if status not in {"APPROVED", "PENDING", "REJECTED", "SUPERSEDED"}:
        raise ValueError("Use an explicit drawing approval status.")
    key = db.execute("INSERT INTO p4_drawing(number,revision,status,reference,snapshot) VALUES (?,?,?,?,?)",
                     (number, revision, status, reference, snapshot_ref(db, snapshot))).lastrowid
    log(db, journal, "drawing_recorded", actor, {"id": key, "number": number, "revision": revision, "status": status,
                                                 "reference": reference, "snapshot": snapshot})
    return key


def drawing_status(db, journal, actor, id, status, reference):
    previous = found(db, "p4_drawing", id)
    if status not in {"APPROVED", "PENDING", "REJECTED", "SUPERSEDED"}:
        raise ValueError("Use an explicit drawing approval status.")
    db.execute("UPDATE p4_drawing SET status=?,reference=? WHERE id=?",
               (status, required(reference, "Drawing approval evidence"), id))
    log(db, journal, "drawing_status", actor, {"id": id, "before": previous["status"],
                                              "after": status, "previous_reference": previous["reference"], "reference": reference})
    return id


def quantity_current(db, takeoff):
    current = db.execute(
        """SELECT d.* FROM p4_drawing d WHERE d.number =
           (SELECT number FROM p4_drawing WHERE id=?) ORDER BY id DESC LIMIT 1""",
        (takeoff["drawing"],),
    ).fetchone()
    return (takeoff["state"] == "APPROVED" and current["id"] == takeoff["drawing"]
            and current["status"] == "APPROVED")


def signature(unit):
    unit = required(unit, "Dimension unit").replace("²", "2").replace("³", "3")
    if unit == "1":
        return {}
    match = re.fullmatch(r"([A-Za-z]+)([1-9])?", unit)
    if not match:
        raise ValueError("Use a plain unit such as ea, m, m2, m3 or kg.")
    return {match[1]: int(match[2] or "1")}


def design_quantity(formula, dimensions, output_unit):
    if not isinstance(dimensions, dict) or not dimensions:
        raise ValueError("Supply drawing dimensions and their source units.")
    if not isinstance(formula, str) or len(formula) > 500:
        raise ValueError("Enter a dimension formula of up to 500 characters.")
    variables = {}
    for name, measurement in dimensions.items():
        if not re.fullmatch(r"[a-zA-Z][a-zA-Z0-9_]*", name):
            raise ValueError("Use dimension variable names such as length or depth.")
        if not isinstance(measurement, dict) or set(measurement) != {"value", "unit"}:
            raise ValueError(f"Supply a value and unit for {name}.")
        variables[name] = (decimal(measurement["value"]), signature(measurement["unit"]))
        if variables[name][0] < 0:
            raise ValueError("Dimensions must be nonnegative.")

    def calculate(node):
        if isinstance(node, ast.Name) and node.id in variables:
            return variables[node.id]
        if isinstance(node, ast.Constant) and type(node.value) in (int, float):
            return decimal(ast.get_source_segment(formula, node)), {}
        if isinstance(node, ast.BinOp):
            (left, left_unit), (right, right_unit) = calculate(node.left), calculate(node.right)
            if isinstance(node.op, (ast.Add, ast.Sub)):
                if left_unit != right_unit:
                    raise ValueError("Cannot add or subtract mismatched dimension units.")
                return (left + right if isinstance(node.op, ast.Add) else left - right), left_unit
            if isinstance(node.op, (ast.Mult, ast.Div)):
                multiply = isinstance(node.op, ast.Mult)
                result_unit = dict(left_unit)
                for base, power in right_unit.items():
                    result_unit[base] = result_unit.get(base, 0) + (power if multiply else -power)
                return (left * right if multiply else left / right), {k: v for k, v in result_unit.items() if v}
        raise ValueError("Use supplied dimension names, numbers, parentheses and + - * / only.")

    try:
        tree = ast.parse(formula, mode="eval")
        if not any(isinstance(node, ast.Name) for node in ast.walk(tree)):
            raise ValueError("Reference at least one recorded drawing dimension.")
        value, result_unit = calculate(tree.body)
        if result_unit != signature(output_unit):
            raise ValueError("Formula dimensions do not match the material unit. Convert and record dimensions explicitly.")
        value = decimal(value)
    except (SyntaxError, DecimalException) as error:
        raise ValueError("Drawing formula could not be calculated.") from error
    return units(value)


def takeoff_add(db, journal, actor, id, drawing, workfront, material_id, element, dimensions,
                formula, unit, source, allowance_pct):
    workfront = front(db, workfront)
    material(db, material_id, unit)
    found(db, "p4_drawing", drawing)
    design = design_quantity(formula, dimensions, unit)
    percentage = decimal(allowance_pct)
    if percentage < 0:
        raise ValueError("Allowance must be explicit and nonnegative.")
    allowance = int((Decimal(design) * percentage / 100).to_integral_value(rounding=ROUND_CEILING))
    if design + allowance > LIMIT:
        raise ValueError("Procurement quantity is outside the supported range.")
    db.execute("""INSERT INTO p4_takeoff VALUES (?,?,?,?,?,?,?,?,?,?,?,?, 'DRAFT', ?, NULL)""",
               (required(id, "Take-off ID"), drawing, workfront, material_id,
                required(element, "Drawing element"), json.dumps(dimensions), formula,
                required(source, "Dimension source reference"), unit, design, allowance, design + allowance, actor))
    log(db, journal, "takeoff_recorded", actor, {"id": id, "design": show(design),
                                                 "allowance": show(allowance), "procurement": show(design + allowance)})
    return id


def takeoff_review(db, journal, actor, id, action, reference):
    takeoff = found(db, "p4_takeoff", id)
    reference = required(reference, "Review evidence")
    if normalize(actor) == normalize(takeoff["preparer"]):
        raise ValueError("A different named person must review the calculated quantity.")
    expected, updated = {"verify": ("DRAFT", "VERIFIED"), "approve": ("VERIFIED", "APPROVED")}.get(action, (None, None))
    if takeoff["state"] != expected:
        raise ValueError("Verify the draft before approving the quantity.")
    if action == "approve" and normalize(actor) == normalize(takeoff["verifier"]):
        raise ValueError("Quantity approval requires a person other than its verifier.")
    drawing = found(db, "p4_drawing", takeoff["drawing"])
    latest = db.execute("SELECT MAX(id) FROM p4_drawing WHERE number=?", (drawing["number"],)).fetchone()[0]
    if drawing["id"] != latest or drawing["status"] != "APPROVED":
        raise ValueError("The drawing revision is superseded or unapproved.")
    db.execute("UPDATE p4_takeoff SET state=?,verifier=? WHERE id=?",
               (updated, actor if action == "verify" else takeoff["verifier"], id))
    log(db, journal, "takeoff_" + updated.lower(), actor, {"id": id, "reference": reference})
    return id


def dates(start, finish):
    current, end = date.fromisoformat(start), date.fromisoformat(finish)
    while current <= end:
        yield current
        current += timedelta(days=1)


def programme_add(db, journal, actor, id, takeoff, activity, location, start, finish,
                  workdays, rate, unit, production_unit, source, daily=None, target=None, exclusions=(), snapshot=None):
    take = found(db, "p4_takeoff", takeoff)
    activity = required(activity, "Programme activity")
    front(db, take["workfront"], activity)
    material(db, take["material"], unit)
    start, finish = day(start), day(finish)
    if not 0 <= (date.fromisoformat(finish) - date.fromisoformat(start)).days <= 3660:
        raise ValueError("Programme finish must follow start within ten years.")
    if not workdays or any(type(value) is not int or value not in range(7) for value in workdays):
        raise ValueError("Supply working weekdays, Monday = 0 through Sunday = 6.")
    exclusions = sorted({day(item) for item in exclusions})
    eligible = sum(d.weekday() in workdays and d.isoformat() not in exclusions for d in dates(start, finish))
    if not eligible:
        raise ValueError("Programme interval has no configured working days.")
    if daily is not None and target is not None:
        raise ValueError("Provide daily production or the programme target.")
    if target is not None:
        daily = show((real_amount(target) + eligible - 1) // eligible)
    if daily is not None:
        daily = real_amount(daily)
    rate = real_amount(rate) if rate is not None else None
    if rate is not None and daily is not None and rate * daily > LIMIT * SCALE:
        raise ValueError("Daily demand is outside the supported range.")
    overlap = db.execute("""SELECT 1 FROM p4_plan WHERE active=1 AND workfront=? AND material=?
                          AND activity=? AND takeoff=? AND start<=? AND finish>=?""",
                         (take["workfront"], take["material"], activity, takeoff, finish, start)).fetchone()
    if overlap:
        raise ValueError("Programme entry overlaps demand already recorded for this take-off.")
    db.execute("""INSERT INTO p4_plan VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,NULL,1)""",
               (required(id, "Programme ID"), takeoff, take["material"], take["workfront"],
                required(location, "Stock location"), activity, start, finish,
                json.dumps(sorted(set(workdays))), json.dumps(exclusions), daily, rate,
                required(production_unit, "Programme production unit"),
                required(source, "Programme source"), snapshot_ref(db, snapshot), actor))
    log(db, journal, "programme_recorded", actor, {"id": id, "snapshot": snapshot, "source": source})
    return id


def programme_import(db, journal, actor, workbook, row, **inputs):
    if type(row) is not int or row not in range(FIRST_ROW, NEW_END + 1):
        raise ValueError("Select a prepared programme row.")
    from pathlib import Path
    path = Path(workbook).resolve()
    digest = fingerprint(path)
    snapshot_ref(db, digest)
    known = db.execute("SELECT output FROM proposal_commits WHERE output_sha=? AND state='committed'",
                       (digest,)).fetchall()
    if not any(Path(item["output"]).resolve() == path for item in known):
        raise ValueError("Use the committed workbook path recorded in the journal.")
    raw = path.read_bytes()
    from .reports import digest as hash_bytes
    if hash_bytes(raw) != digest:
        raise ValueError("Programme snapshot changed during import.")
    book = load_workbook(BytesIO(raw))
    try:
        validate(book, expanded=True)
        record = book.worksheets[11][row]
        activity, start, finish, source_unit, target = [record[index].value for index in (0, 3, 4, 5, 6)]
        expected_unit = inputs.pop("production_unit")
        if not source_unit or source_unit != expected_unit:
            raise ValueError("Programme production unit does not match the supplied rate basis.")
        return programme_add(db, journal, actor, activity=activity, start=wire(start), finish=wire(finish),
                             target=target, source=f"Approved programme, sheet 12 row {row}",
                             production_unit=expected_unit,
                             snapshot=digest, **inputs)
    finally:
        book.close()


def programme_verify(db, journal, actor, id, reference):
    row = found(db, "p4_plan", id)
    if not row["active"] or row["reviewer"] or normalize(row["preparer"]) == normalize(actor):
        raise ValueError("Use a different named reviewer for an unverified programme.")
    take = found(db, "p4_takeoff", row["takeoff"])
    if not quantity_current(db, take) or row["daily"] is None or row["rate"] is None:
        raise ValueError("An approved take-off and complete production and consumption rates are required.")
    db.execute("UPDATE p4_plan SET reviewer=? WHERE id=?", (actor, id))
    log(db, journal, "programme_verified", actor, {"id": id, "reference": required(reference, "Review evidence")})
    return id


def programme_retire(db, journal, actor, id, reference):
    row = found(db, "p4_plan", id)
    if not row["active"]:
        raise ValueError("Programme entry is already retired.")
    db.execute("UPDATE p4_plan SET active=0 WHERE id=?", (id,))
    log(db, journal, "programme_retired", actor, {"id": id, "reference": required(reference, "Revision evidence")})
    return id


def stock(db, material_id, location, by=None):
    moves = db.execute(
        """SELECT * FROM p4_movement WHERE material=? AND location=? AND (? IS NULL OR day<=?)
           ORDER BY day,rowid""", (material_id, location, by, by),
    ).fetchall()
    reservations = {}
    for move in moves:
        if move["reserved"]:
            key = move["workfront"]
            reservations[key] = reservations.get(key, 0) + move["reserved"]
    usable = sum(move["usable"] for move in moves)
    quarantined = sum(move["quarantine"] for move in moves)
    reserved = sum(reservations.values())
    return {"opening_known": any(move["kind"] == "opening" for move in moves), "physical": usable + quarantined,
            "usable": usable, "quarantine": quarantined, "reserved": reserved,
            "available": usable - reserved, "allocations": reservations, "moves": moves}


def movement(db, journal, actor, id, material_id, unit, location, date, kind, quantity, source, workfront=None):
    material(db, material_id, unit)
    location = required(location, "Stock location")
    when = day(date)
    current = stock(db, material_id, location)
    if current["moves"] and when < current["moves"][-1]["day"]:
        raise ValueError("Stock movements must be posted in effective date order for this location.")
    amount = units(quantity, signed=kind == "adjustment")
    if not amount and kind != "opening":
        raise ValueError("Only an opening stock count may be zero.")
    if workfront:
        workfront = front(db, workfront)
    if kind == "opening":
        if current["opening_known"]:
            raise ValueError("Opening stock already exists for this material and location.")
        if workfront:
            raise ValueError("Opening stock must be counted for the location.")
    elif not current["opening_known"]:
        raise ValueError("Record an opening stock count, including explicit zero if applicable.")
    usable = quarantined = reserved = 0
    own = current["allocations"].get(workfront, 0)
    if kind in {"opening", "receipt", "adjustment"}:
        usable = amount
    elif kind == "quarantine":
        quarantined = amount
    elif kind == "quarantine-release":
        usable, quarantined = amount, -amount
    elif kind == "quarantine-reject":
        quarantined = -amount
    elif kind in {"reserve", "release", "issue", "consume"}:
        if not workfront:
            raise ValueError("Allocation, issue and consumption require a Workfront ID.")
        if kind == "reserve":
            reserved = amount
        elif kind == "release":
            reserved = -amount
        elif kind == "issue":
            usable, reserved = -amount, -min(own, amount)
        else:
            issued = sum(move["amount"] for move in current["moves"]
                         if move["kind"] == "issue" and move["workfront"] == workfront)
            consumed = sum(move["amount"] for move in current["moves"]
                           if move["kind"] == "consume" and move["workfront"] == workfront)
            if consumed + amount > issued:
                raise ValueError("Actual consumption exceeds material already issued to this workfront.")
    else:
        raise ValueError("Unknown stock movement.")
    if (own + reserved < 0 or current["quarantine"] + quarantined < 0 or
            current["available"] + usable - reserved < 0):
        raise ValueError("Insufficient unallocated, reserved or quarantined stock for this movement.")
    db.execute("INSERT INTO p4_movement VALUES (?,?,?,?,?,?,?,?,?,?,?)",
               (required(id, "Movement ID"), material_id, location, when, kind, amount, workfront,
                usable, quarantined, reserved, required(source, "Movement source")))
    log(db, journal, "stock_movement", actor, {"id": id, "kind": kind, "quantity": show(amount),
                                             "available_before": show(current["available"]),
                                             "available_after": show(current["available"] + usable - reserved)})
    return id
