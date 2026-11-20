"""Command-line entry point and transaction boundary for local P4 workflows."""
import argparse
import json
import sqlite3
from pathlib import Path

from . import material_store as store
from . import material_delivery as delivery
from . import material_requisition as requests
from .journal import Journal
from .material_forecast import forecast


ACTIONS = {
    "material": store.material_add,
    "material-policy": store.material_policy,
    "drawing": store.drawing_add,
    "drawing-status": store.drawing_status,
    "takeoff": store.takeoff_add,
    "takeoff-review": store.takeoff_review,
    "programme": store.programme_add,
    "programme-import": store.programme_import,
    "programme-verify": store.programme_verify,
    "programme-retire": store.programme_retire,
    "movement": store.movement,
    "order": delivery.order_add,
    "order-confirm": delivery.order_confirm,
    "arrival": delivery.arrival_add,
    "prepare-requisitions": requests.prepare,
    "revise-requisition": requests.revise,
    "requisition-action": requests.requisition_action,
}


def apply(journal, actor, operations):
    store.required(actor, "Operator")
    results = []
    with journal.transaction() as db:
        store.setup(db)
        for operation in operations:
            name = operation.get("operation")
            if name not in ACTIONS:
                raise ValueError(f"Unknown material operation: {name}")
            fields = {key: value for key, value in operation.items() if key != "operation"}
            results.append(ACTIONS[name](db, journal, actor, **fields))
    return results


def read(journal, reader, **arguments):
    with journal.transaction() as db:
        store.setup(db)
        return reader(db, **arguments)


def list_requisitions(db):
    return [{**dict(row), **json.loads(row["body"])} for row in db.execute(
        "SELECT id,revision,state,body,author,reviewer,approver,closed FROM p4_requisition ORDER BY id DESC"
    )]


def list_takeoffs(db):
    result = []
    for row in db.execute("""SELECT t.*,d.number,d.revision,d.reference FROM p4_takeoff t
                           JOIN p4_drawing d ON d.id=t.drawing ORDER BY t.id"""):
        result.append({**dict(row), "dimensions": json.loads(row["dimensions"]),
                       "current_approved": store.quantity_current(db, row),
                       **{key: store.show(row[key]) for key in ("design", "allowance", "procurement")}})
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--journal", default="local/audit.sqlite3")
    sub = parser.add_subparsers(dest="command", required=True)
    batch = sub.add_parser("apply")
    batch.add_argument("operations", help="JSON object or array")
    batch.add_argument("--actor", required=True)
    view = sub.add_parser("forecast")
    view.add_argument("start")
    view.add_argument("--workfront")
    view.add_argument("--material", dest="material_id")
    view.add_argument("--status")
    sub.add_parser("requisitions")
    sub.add_parser("takeoffs")
    export = sub.add_parser("export")
    export.add_argument("id", type=int)
    export.add_argument("directory")
    args = parser.parse_args()
    try:
        journal = Journal(args.journal)
        if args.command == "apply":
            value = json.loads(Path(args.operations).read_text(encoding="utf-8-sig"))
            result = apply(journal, args.actor, value if isinstance(value, list) else [value])
        elif args.command == "forecast":
            result = read(journal, forecast, start=args.start, workfront=args.workfront,
                          material_id=args.material_id, status=args.status)
        elif args.command == "requisitions":
            result = read(journal, list_requisitions)
        elif args.command == "takeoffs":
            result = read(journal, list_takeoffs)
        else:
            result = read(journal, requests.export, id=args.id, directory=args.directory)
        print(json.dumps(result, ensure_ascii=True, indent=2))
    except (ValueError, TypeError, AttributeError, OSError, sqlite3.Error) as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
