"""Project-wide Workfront IDs and exact, normalized alias resolution."""
import json
import re


def normalize(alias):
    return " ".join(alias.casefold().split())


def workfront_id(sequence):
    return f"KVDT2-WF-{sequence:06d}"


def create(journal, actor, section, location, activities, aliases):
    normalized = list(dict.fromkeys(normalize(alias) for alias in aliases))
    if not section.strip() or not location.strip() or not normalized or not all(normalized):
        raise ValueError("Section, location and at least one nonblank alias are required")
    with journal.transaction() as db:
        for alias in normalized:
            if db.execute("SELECT 1 FROM aliases WHERE alias=?", (alias,)).fetchone():
                raise ValueError(f"Alias is already allocated: {alias}")
        sequence = db.execute(
            "INSERT INTO workfronts(section,location,activities) VALUES (?,?,?)",
            (section, location, json.dumps(activities)),
        ).lastrowid
        db.executemany("INSERT INTO aliases(alias,sequence) VALUES (?,?)",
                       [(alias, sequence) for alias in normalized])
        identifier = workfront_id(sequence)
        journal.log(db, "workfront_created", actor, {
            "id": identifier, "aliases": normalized, "section": section,
            "location": location, "activities": activities,
        })
    return identifier


def resolve(journal, lookup):
    with journal.transaction() as db:
        match = re.fullmatch(r"KVDT2-WF-(\d{6})", lookup.strip(), re.IGNORECASE)
        if match:
            row = db.execute("SELECT * FROM workfronts WHERE sequence=? AND retired=0",
                             (int(match[1]),)).fetchone()
        else:
            row = db.execute("""SELECT w.* FROM workfronts w JOIN aliases a USING(sequence)
                                WHERE a.alias=? AND w.retired=0""", (normalize(lookup),)).fetchone()
        if row is None:
            raise ValueError("Unknown or retired workfront")
        return {"id": workfront_id(row["sequence"]), "section": row["section"],
                "location": row["location"], "activities": json.loads(row["activities"])}


def retire(journal, actor, identifier):
    match = re.fullmatch(r"KVDT2-WF-(\d{6})", identifier)
    if not match:
        raise ValueError("Supply a canonical Workfront ID")
    with journal.transaction() as db:
        result = db.execute("UPDATE workfronts SET retired=1 WHERE sequence=? AND retired=0",
                            (int(match[1]),))
        if not result.rowcount:
            raise ValueError("Unknown or already retired workfront")
        journal.log(db, "workfront_retired", actor, {"id": identifier})
