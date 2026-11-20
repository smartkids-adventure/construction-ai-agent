"""Local SQLite audit journal and immutable-by-workflow workbook revisions."""
import hashlib
import json
import os
import shutil
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory

from .workbook import stage_migration


def fingerprint(path):
    with Path(path).open("rb") as source:
        return hashlib.file_digest(source, "sha256").hexdigest()


class Journal:
    def __init__(self, path):
        self.path = Path(path).resolve()
        with self.transaction() as db:
            db.executescript("""
                CREATE TABLE IF NOT EXISTS events (
                    id INTEGER PRIMARY KEY, event TEXT NOT NULL,
                    actor TEXT NOT NULL, detail TEXT NOT NULL,
                    created TEXT NOT NULL DEFAULT
                        (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
                );
                CREATE TABLE IF NOT EXISTS revisions (
                    id INTEGER PRIMARY KEY, source_sha TEXT NOT NULL,
                    output_sha TEXT NOT NULL, backup TEXT NOT NULL,
                    output TEXT NOT NULL, actor TEXT NOT NULL,
                    state TEXT NOT NULL CHECK (state IN ('pending','committed','aborted'))
                );
                CREATE UNIQUE INDEX IF NOT EXISTS active_output ON revisions(output)
                    WHERE state <> 'aborted';
                CREATE TABLE IF NOT EXISTS workfronts (
                    sequence INTEGER PRIMARY KEY AUTOINCREMENT CHECK (sequence <= 999999),
                    section TEXT NOT NULL, location TEXT NOT NULL,
                    activities TEXT NOT NULL, retired INTEGER NOT NULL DEFAULT 0
                );
                CREATE TABLE IF NOT EXISTS aliases (
                    alias TEXT PRIMARY KEY, sequence INTEGER NOT NULL REFERENCES workfronts(sequence)
                );
            """)

    @contextmanager
    def transaction(self):
        db = sqlite3.connect(self.path, timeout=15)
        db.row_factory = sqlite3.Row
        try:
            db.execute("PRAGMA foreign_keys = ON")
            db.execute("BEGIN IMMEDIATE")
            yield db
            db.commit()
        finally:
            db.close()

    @staticmethod
    def log(db, event, actor, detail):
        if not actor.strip():
            raise ValueError("An operator name is required")
        db.execute("INSERT INTO events(event, actor, detail) VALUES (?,?,?)",
                   (event, actor, json.dumps(detail, ensure_ascii=False)))

    def recover(self, actor):
        outcomes = []
        with self.transaction() as db:
            for row in db.execute("SELECT * FROM revisions WHERE state='pending'").fetchall():
                output = Path(row["output"])
                if output.exists() and fingerprint(output) != row["output_sha"]:
                    outcomes.append((row["id"], "conflict: output differs; retained for review"))
                    continue
                state = "committed" if output.exists() else "aborted"
                db.execute("UPDATE revisions SET state=? WHERE id=?", (state, row["id"]))
                self.log(db, "migration_recovered", actor, {"revision": row["id"], "state": state})
                outcomes.append((row["id"], state))
        return outcomes


def migrate(source, output, journal, actor, approved_sha):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.exists():
        raise ValueError("Specify a new output file, separate from the source")
    if source.suffix.lower() != ".xlsx" or output.suffix.lower() != ".xlsx":
        raise ValueError("Workbook paths must end in .xlsx")
    if not actor.strip() or fingerprint(source) != approved_sha:
        raise ValueError("Operator and exact approved source fingerprint are required")
    backups = journal.path.parent / "backups"
    backups.mkdir(exist_ok=True)
    backup = backups / f"{approved_sha}.xlsx"
    with TemporaryDirectory(dir=output.parent) as temp:
        snapshot = Path(temp) / "source.xlsx"
        staged = Path(temp) / "staged.xlsx"
        shutil.copyfile(source, snapshot)
        if fingerprint(snapshot) != approved_sha:
            raise ValueError("Source changed while staging")
        stage_migration(snapshot, staged)
        output_sha = fingerprint(staged)
        if fingerprint(source) != approved_sha:
            raise ValueError("Source changed after approval")
        with journal.transaction() as db:
            if not backup.exists():
                # Stage on the backup filesystem. Publish the complete file exclusively.
                with TemporaryDirectory(dir=backups) as backup_temp:
                    candidate = Path(backup_temp) / "backup.xlsx"
                    shutil.copyfile(snapshot, candidate)
                    os.link(candidate, backup)
            if fingerprint(backup) != approved_sha:
                raise ValueError("Existing backup checksum mismatch")
            revision = db.execute(
                """INSERT INTO revisions(source_sha,output_sha,backup,output,actor,state)
                   VALUES (?,?,?,?,?,'pending')""",
                (approved_sha, output_sha, str(backup), str(output), actor),
            ).lastrowid
            journal.log(db, "migration_approved", actor, {
                "revision": revision, "input": str(source), "input_sha": approved_sha,
                "output_sha": output_sha, "capacity": 1000,
            })
        with journal.transaction() as db:
            if db.execute("SELECT state FROM revisions WHERE id=?", (revision,)).fetchone()[0] != "pending":
                raise ValueError("Revision was reconciled while publishing")
            if fingerprint(source) != approved_sha:
                raise ValueError("Source changed before publication")
            # Atomic and exclusive on NTFS. An existing workbook is never overwritten.
            os.link(staged, output)
            db.execute("UPDATE revisions SET state='committed' WHERE id=?", (revision,))
            journal.log(db, "migration_committed", actor, {"revision": revision})
    return revision
