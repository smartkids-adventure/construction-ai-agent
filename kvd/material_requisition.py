"""Create shortage-based requisitions and issue approved Word/PDF files."""
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import shutil
from tempfile import mkdtemp
from xml.sax.saxutils import escape

from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate

from .material_forecast import forecast
from .material_store import found, log, required
from .proposals import checksum
from .workfront import normalize


def body(result):
    return {"material": result["material"], "description": result["description"],
            "specification": result["specification"], "location": result["location"],
            "workfront": result["workfront"], "unit": result["unit"], "start": result["start"],
            "quantity": result["shortage"], "required": result["stockout"],
            "requisition_by": result["requisition_by"], "references": result["references"]}


def save_draft(db, journal, actor, result, scope, revision):
    data = body(result)
    key = db.execute(
        """INSERT INTO p4_requisition(scope,revision,state,body,digest,author)
           VALUES (?,?,'DRAFT',?,?,?)""",
        (scope, revision, json.dumps(data, ensure_ascii=False), checksum(data), actor),
    ).lastrowid
    log(db, journal, "requisition_drafted", actor, {"id": key, "revision": revision, "scope": scope,
                                                     "required": data["required"], "quantity": data["quantity"]})
    return key


def prepare(db, journal, actor, start, workfront=None, material_id=None):
    ids = []
    for result in forecast(db, start, workfront=workfront, material_id=material_id, status="RED"):
        scope = checksum([result[key] for key in ("material", "location", "workfront")])
        previous = db.execute("SELECT * FROM p4_requisition WHERE scope=? ORDER BY revision DESC LIMIT 1",
                              (scope,)).fetchone()
        if previous and previous["state"] != "SUPERSEDED" and not previous["closed"]:
            continue
        if db.execute("SELECT 1 FROM p4_requisition WHERE scope=? AND digest=? AND state='ISSUED'",
                      (scope, checksum(body(result)))).fetchone():
            continue
        ids.append(save_draft(db, journal, actor, result, scope, previous["revision"] + 1 if previous else 1))
    return ids


def inspect(db, identifier):
    row = found(db, "p4_requisition", identifier)
    data = json.loads(row["body"])
    if checksum(data) != row["digest"]:
        raise ValueError("Review data changed after the revision was created.")
    return row, data


def current(db, data, start):
    for result in forecast(db, start, workfront=data["workfront"],
                           material_id=data["material"], status="RED"):
        if result["location"] == data["location"]:
            return result
    raise ValueError("There is no verified shortage for this scope and forecast window.")


def revise(db, journal, actor, id, start, reference):
    row, data = inspect(db, id)
    if row["state"] not in {"DRAFT", "REVIEWED", "APPROVED"}:
        raise ValueError("Only an unissued current revision can be superseded.")
    latest = db.execute("SELECT MAX(revision) FROM p4_requisition WHERE scope=?", (row["scope"],)).fetchone()[0]
    if latest != row["revision"]:
        raise ValueError("Select the latest revision.")
    result = current(db, data, start)
    db.execute("UPDATE p4_requisition SET state='SUPERSEDED' WHERE id=?", (id,))
    successor = save_draft(db, journal, actor, result, row["scope"], latest + 1)
    log(db, journal, "requisition_revised", actor, {"before": id, "after": successor,
                                                    "reference": required(reference, "Revision evidence")})
    return successor


def requisition_action(db, journal, actor, id, action, reference):
    row, data = inspect(db, id)
    reference = required(reference, "Decision reference")
    if action == "close":
        if row["state"] != "ISSUED" or row["closed"]:
            raise ValueError("Only an open issued requisition can be closed.")
        db.execute("UPDATE p4_requisition SET closed=? WHERE id=?", (reference, id))
        log(db, journal, "requisition_closed", actor, {"id": id, "reference": reference})
        return id
    transitions = {"review": ("DRAFT", "REVIEWED"), "approve": ("REVIEWED", "APPROVED"),
                   "issue": ("APPROVED", "ISSUED")}
    if action not in transitions or row["state"] != transitions[action][0]:
        raise ValueError("Follow draft, reviewed, approved and issued order.")
    target = transitions[action][1]
    if action in {"review", "approve"} and (
        normalize(actor) == normalize(row["author"]) or
        (action == "approve" and normalize(actor) == normalize(row["reviewer"]))
    ):
        raise ValueError("Use separate named people to prepare, review and approve the requisition.")
    if body(current(db, data, data["start"])) != data:
        raise ValueError("Forecast or source references changed. Create a new draft revision.")
    if action == "review":
        db.execute("UPDATE p4_requisition SET state=?,reviewer=? WHERE id=?", (target, actor, id))
    elif action == "approve":
        db.execute("UPDATE p4_requisition SET state=?,approver=? WHERE id=?", (target, actor, id))
    else:
        word, pdf = render(row, data, actor)
        db.execute("""UPDATE p4_requisition SET state=?,word=?,pdf=?,word_sha=?,pdf_sha=? WHERE id=?""",
                   (target, word, pdf, sha256(word).hexdigest(), sha256(pdf).hexdigest(), id))
    log(db, journal, "requisition_" + target.lower(), actor, {"id": id, "reference": reference,
                                                               "digest": row["digest"]})
    return id


def render(row, data, issuer):
    title = "Material Requisition"
    lines = [
        f'MR-{row["id"]:04d}  |  Revision {row["revision"]:02d}  |  Issued  |  Synthetic data',
        f'Workfront: {data["workfront"]}  |  Stock location: {data["location"]}',
        f'Material: {data["material"]} - {data["description"]}',
        f'Specification: {data["specification"]}',
        f'Quantity requested: {data["quantity"]} {data["unit"]}',
        f'Required delivery: {data["required"]}  |  Requisition recommended by: {data["requisition_by"]}',
        f'Prepared by {row["author"]} | Reviewed by {row["reviewer"]}',
        f'Approved by {row["approver"]} | Issued by {issuer}',
        "Supporting references",
    ]
    for item in data["references"]:
        lines.append(f'{item["drawing"]} ({item["drawing_ref"]}); take-off {item["takeoff"]} ({item["quantity_ref"]})')
        if item["drawing_snapshot"]:
            lines.append(f'Drawing snapshot SHA-256: {item["drawing_snapshot"]}')
        lines.append(f'Programme {item["plan"]}: {item["programme"]}')
        if item["snapshot"]:
            lines.append(f'Snapshot SHA-256: {item["snapshot"]}')
    doc, word_stream, pdf_stream = Document(), BytesIO(), BytesIO()
    doc.add_heading(title, 0)
    for line in lines:
        if line == "Supporting references":
            doc.add_heading(line, 1)
        else:
            doc.add_paragraph(line)
    doc.core_properties.title = title
    doc.save(word_stream)
    styles = getSampleStyleSheet()
    blocks = [Paragraph(title, styles["Title"])]
    blocks.extend(Paragraph(escape(line), styles["Heading2" if line == "Supporting references" else "BodyText"])
                  for line in lines)
    SimpleDocTemplate(pdf_stream, pagesize=A4, title=title).build(blocks)
    return word_stream.getvalue(), pdf_stream.getvalue()


def export(db, id, directory):
    row, _ = inspect(db, id)
    if row["state"] != "ISSUED":
        raise ValueError("Issue the approved requisition before exporting.")
    files = {"docx": row["word"], "pdf": row["pdf"]}
    if sha256(files["docx"]).hexdigest() != row["word_sha"] or sha256(files["pdf"]).hexdigest() != row["pdf_sha"]:
        raise ValueError("Stored requisition output failed the checksum check.")
    prefix = f'MR-{id:04d}_r{row["revision"]:02d}'
    parent = Path(directory).resolve()
    destination = parent / prefix
    if destination.exists():
        if all((destination / f"{prefix}.{ext}").is_file()
               and (destination / f"{prefix}.{ext}").read_bytes() == data for ext, data in files.items()):
            return str(destination)
        raise ValueError("Export destination contains different files.")
    staging = Path(mkdtemp(dir=parent, prefix=".requisition-"))
    try:
        for ext, data in files.items():
            (staging / f"{prefix}.{ext}").write_bytes(data)
        staging.rename(destination)
    finally:
        if staging.exists():
            shutil.rmtree(staging)
    return str(destination)
