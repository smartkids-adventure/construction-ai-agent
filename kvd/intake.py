"""Deterministic parser for manually pasted synthetic site reports."""
from dataclasses import dataclass
from datetime import date
import re


COMMANDS = {
    "MC01": "Daily Site Update",
    "MC14": "Junior Engineer Site Update",
    "MC11": "Safety Toolbox and Attendance",
    "MC12": "Utility Clearance and Workfront Readiness",
}
COMMON = ("workfront", "date", "shift", "activity", "evidence")
SITE = ("activity id", "bq item", "crew", "unit", "reported qty", "notes", "resources")
FIELDS = {
    "MC01": COMMON + SITE,
    "MC14": COMMON + SITE + ("task id", "outcome"),
    "MC11": COMMON + ("lead", "topics", "jsa ref", "attendees"),
    "MC12": COMMON + ("utility id", "utility type", "owner", "blocker", "relocation instruction"),
}
REQUIRED = {
    "MC01": ("activity id", "unit", "reported qty"),
    "MC14": ("activity id", "unit", "reported qty", "task id", "outcome"),
    "MC11": ("lead", "topics", "attendees"),
    "MC12": ("utility id", "utility type"),
}

LABELS_MS = {
    "had kerja": "workfront", "zon kerja": "workfront", "kawasan kerja": "workfront",
    "lokasi kerja": "workfront", "arahan": "command", "perintah": "command",
    "tarikh": "date", "syif": "shift", "aktiviti": "activity", "bukti": "evidence",
    "id aktiviti": "activity id", "item bq": "bq item", "kru": "crew",
    "kuantiti dilaporkan": "reported qty", "kuantiti lapor": "reported qty",
    "catatan": "notes", "sumber": "resources", "id tugasan": "task id",
    "hasil": "outcome", "ketua": "lead", "topik": "topics", "rujukan jsa": "jsa ref",
    "kehadiran": "attendees", "id utiliti": "utility id", "jenis utiliti": "utility type",
    "pemilik": "owner", "halangan": "blocker", "arahan pemindahan": "relocation instruction",
}


@dataclass
class Draft:
    command: str
    fields: dict
    source: str


def parse(text, default_command="MC01"):
    if default_command not in COMMANDS:
        raise ValueError("Select a supported Master Command")
    if not text.strip() or len(text) > 100_000:
        raise ValueError("Paste a synthetic report of 1–100,000 characters")
    groups = []
    for line in text.splitlines():
        if not line.strip():
            continue
        match = re.fullmatch(r"\s*([a-zA-Z][a-zA-Z _-]*):\s*(.*?)\s*", line)
        if not match:
            raise ValueError(f"Expected 'Field: value': {line[:70]}")
        key = " ".join(match[1].lower().replace("_", " ").replace("-", " ").split())
        key = LABELS_MS.get(key, key)
        if key == "workfront":
            groups.append(({}, []))
        if not groups:
            raise ValueError("Start each report with Workfront: alias or Zon kerja: alias")
        fields, lines = groups[-1]
        if key in fields:
            raise ValueError(f"Repeated field: {key}")
        fields[key] = match[2]
        lines.append(line)
    if len(groups) > 100:
        raise ValueError("Review at most 100 workfront blocks per paste")
    drafts = []
    for fields, lines in groups:
        command = fields.pop("command", default_command).upper()
        if command not in COMMANDS:
            raise ValueError(f"Unsupported command: {command}")
        unknown = fields.keys() - set(FIELDS[command])
        if unknown:
            raise ValueError(f"{command}: unsupported fields: {', '.join(sorted(unknown))}")
        drafts.append(Draft(command, fields, "\n".join(lines)))
    return drafts


def check(draft):
    fields = draft.fields
    if draft.command not in COMMANDS or fields.keys() - set(FIELDS[draft.command]):
        raise ValueError("Draft contains fields outside its command allowlist")
    for field in ("workfront", "date", "shift", "activity", "evidence") + REQUIRED[draft.command]:
        if not fields.get(field, "").strip():
            raise ValueError(f"{draft.command}: {field} is required")
    day = date.fromisoformat(fields["date"])
    if day.isoformat() != fields["date"]:
        raise ValueError("Use YYYY-MM-DD for the work date")
    for text in fields.values():
        if len(text) > 2000:
            raise ValueError("Each draft field must be at most 2,000 characters")
    return day.isoformat()
