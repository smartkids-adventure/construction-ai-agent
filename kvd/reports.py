"""Create, approve and issue revisions from an immutable reporting snapshot."""
from datetime import date, timedelta
import hashlib
import json
from pathlib import Path
import shutil
from tempfile import mkdtemp

from .proposals import checksum
from .report_data import approved_snapshot, comparison, iso, norm, select
from .report_render import TITLES, pdf, word


def setup(db):
    db.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY, command TEXT NOT NULL, start TEXT NOT NULL,
            end TEXT NOT NULL, revision INTEGER NOT NULL, payload TEXT NOT NULL,
            digest TEXT NOT NULL, state TEXT NOT NULL CHECK(state IN ('draft','approved','issued')),
            approver TEXT, docx BLOB, pdf BLOB, docx_sha TEXT, pdf_sha TEXT,
            UNIQUE(command,start,end,revision)
        )
    """)


def period(command, anchor):
    if command not in TITLES or not iso(anchor):
        raise ValueError("Choose MC10 or MC02 and an anchor date in YYYY-MM-DD.")
    day = date.fromisoformat(anchor)
    if command == "MC02":
        day -= timedelta(days=day.weekday())
        return day.isoformat(), (day + timedelta(days=6)).isoformat()
    return day.isoformat(), day.isoformat()


def make_payload(snapshot, command, anchor, author):
    start, end = period(command, anchor)
    dated = select(snapshot, from_date=start, to_date=end)
    for section in dated.keys() - {"progress", "resources", "toolbox"}:
        dated[section] = [item for item in snapshot["sections"][section]
                          if not item["day"] or item["day"] <= end]
    return {"command": command, "start": start, "end": end, "author": author.strip(),
            "fingerprint": snapshot["fingerprint"], "source_commit": snapshot["commit"],
            "source": snapshot["file"], "sections": dated,
            "comparison": comparison(dated["progress"])}


def create(journal, path, command, anchor, author, expected_sha=None):
    if not author.strip():
        raise ValueError("Enter the report author's name.")
    snapshot = approved_snapshot(journal, path)
    if expected_sha and snapshot["fingerprint"] != expected_sha:
        raise ValueError("The loaded workbook changed. Reload its committed snapshot.")
    payload = make_payload(snapshot, command, anchor, author)
    with journal.transaction() as db:
        setup(db)
        key = (command, payload["start"], payload["end"])
        payload["revision"] = db.execute(
            "SELECT COALESCE(MAX(revision), 0)+1 FROM reports WHERE command=? AND start=? AND end=?", key
        ).fetchone()[0]
        digest = checksum(payload)
        identifier = db.execute(
            """INSERT INTO reports(command,start,end,revision,payload,digest,state)
               VALUES (?,?,?,?,?,?,'draft')""",
            (*key, payload["revision"], json.dumps(payload, ensure_ascii=False), digest),
        ).lastrowid
        journal.log(db, "p3_report_drafted", author, {"report": identifier, "digest": digest,
                                                     "snapshot": payload["fingerprint"]})
    return identifier


def lookup(db, identifier):
    row = db.execute("SELECT * FROM reports WHERE id=?", (identifier,)).fetchone()
    if row is None:
        raise ValueError("Select a report revision.")
    payload = json.loads(row["payload"])
    if checksum(payload) != row["digest"]:
        raise ValueError("Report revision content changed.")
    return row, payload


def get(journal, identifier):
    with journal.transaction() as db:
        setup(db)
        row, payload = lookup(db, identifier)
        return row["state"], row["approver"], payload


def history(journal):
    with journal.transaction() as db:
        setup(db)
        return [dict(row) for row in db.execute(
            "SELECT id, command, start, end, revision, state, approver FROM reports ORDER BY id DESC"
        )]


def approve(journal, identifier, reviewer):
    if not reviewer.strip():
        raise ValueError("Enter the report reviewer's name.")
    with journal.transaction() as db:
        setup(db)
        row, payload = lookup(db, identifier)
        if row["state"] != "draft":
            raise ValueError("Only a draft can be approved.")
        if norm(reviewer) == norm(payload["author"]):
            raise ValueError("The report needs a different named reviewer.")
        db.execute("UPDATE reports SET state='approved', approver=? WHERE id=?", (reviewer, identifier))
        journal.log(db, "p3_report_approved", reviewer, {"report": identifier, "digest": row["digest"]})


def digest(data):
    return hashlib.sha256(data).hexdigest()


def issue(journal, identifier, actor):
    if not actor.strip():
        raise ValueError("Enter the report issuer's name.")
    with journal.transaction() as db:
        setup(db)
        row, payload = lookup(db, identifier)
        if row["state"] == "draft":
            raise ValueError("Approve the revision before issuing it.")
        if row["state"] == "issued":
            return
        docx_data = word(payload, row["approver"])
        pdf_data = pdf(payload, row["approver"])
        docx_sha, pdf_sha = digest(docx_data), digest(pdf_data)
        db.execute("""UPDATE reports SET state='issued', docx=?, pdf=?, docx_sha=?, pdf_sha=?
                      WHERE id=?""", (docx_data, pdf_data, docx_sha, pdf_sha, identifier))
        journal.log(db, "p3_report_issued", actor, {"report": identifier, "digest": row["digest"],
                                                   "docx_sha": docx_sha, "pdf_sha": pdf_sha})


def export(journal, identifier, directory):
    with journal.transaction() as db:
        setup(db)
        row, payload = lookup(db, identifier)
        if row["state"] != "issued":
            raise ValueError("Issue the approved revision before exporting.")
        files = {extension: row[extension] for extension in ("docx", "pdf")}
        if any(digest(data) != row[f"{ext}_sha"] for ext, data in files.items()):
            raise ValueError("Stored report file fingerprint differs.")
    parent = Path(directory).resolve()
    stem = f'{payload["command"]}_{payload["start"]}_r{payload["revision"]:02d}'
    path = parent / stem
    if path.exists():
        if path.is_dir() and all((path / f"{stem}.{ext}").is_file() and
                                 digest((path / f"{stem}.{ext}").read_bytes()) == digest(data)
                                 for ext, data in files.items()):
            return path
        raise ValueError("Export destination already exists with different contents.")
    staging = Path(mkdtemp(dir=parent, prefix=".report-"))
    try:
        for ext, data in files.items():
            (staging / f"{stem}.{ext}").write_bytes(data)
        staging.rename(path)
    finally:
        if staging.exists():
            shutil.rmtree(staging)
    return path
