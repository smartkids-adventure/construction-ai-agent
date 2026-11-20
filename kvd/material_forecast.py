"""Daily material forecast and consumption variance for a shared stock location."""
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal
import json

from .material_delivery import delivery_register
from .material_store import SCALE, day, dates, found, quantity_current, show, stock


def due(plan, current):
    return (plan["start"] <= current.isoformat() <= plan["finish"]
            and current.weekday() in json.loads(plan["weekdays"])
            and current.isoformat() not in json.loads(plan["exclusions"]))


def consumption(plan):
    return (plan["daily"] * plan["rate"] + SCALE - 1) // SCALE


def sum_moves(moves, front, kind, start=None, end=None):
    return sum(move["amount"] for move in moves if move["workfront"] == front and move["kind"] == kind
               and (not start or start <= move["day"]) and (not end or move["day"] <= end))


def variance_info(db, policy, moves, workfront, location, start, end, known, daily):
    all_plans = db.execute(
        "SELECT DISTINCT takeoff FROM p4_plan WHERE active=1 AND material=? AND workfront=? AND location=?",
        (policy["id"], workfront, location),
    )
    takeoffs = [found(db, "p4_takeoff", row["takeoff"]) for row in all_plans]
    design = sum(t["design"] for t in takeoffs) if takeoffs and all(quantity_current(db, t) for t in takeoffs) else None
    issued = sum_moves(moves, workfront, "issue")
    actual_rows = [move for move in moves if move["workfront"] == workfront and move["kind"] == "consume"
                   and start <= move["day"] <= end]
    actual = sum(move["amount"] for move in actual_rows) if actual_rows else None
    planned = sum(values["demand"] for values in daily) if known else None
    threshold = Decimal(policy["variance"]) if policy["variance"] is not None else None

    def over(value, target):
        return (value is not None and target is not None and threshold is not None
                and Decimal(value - target) * 100 > Decimal(target) * threshold)

    return {"design": show(design), "issued": show(issued), "planned": show(planned), "actual": show(actual),
            "variance": show(actual - planned if actual is not None and planned is not None else None),
            "potential_wastage": over(actual, planned), "design_over_issue": over(issued, design),
            "variance_threshold": str(threshold) if threshold is not None else None}


def forecast(db, start, workfront=None, material_id=None, status=None):
    start = day(start)
    end = (date.fromisoformat(start) + timedelta(days=13)).isoformat()
    horizon = list(dates(start, end))
    if status and status not in {"GREEN", "AMBER", "RED", "UNKNOWN"}:
        raise ValueError("Select GREEN, AMBER, RED or UNKNOWN.")
    deliveries = delivery_register(db, start)
    plans = [dict(row) for row in db.execute(
        "SELECT * FROM p4_plan WHERE active=1 AND start<=? AND finish>=? ORDER BY workfront,id", (end, start))]
    pools = defaultdict(list)
    for row in plans:
        pools[(row["material"], row["location"])].append(row)
    for row in db.execute("SELECT DISTINCT material,location FROM p4_movement WHERE day<=?", (start,)):
        pools[(row["material"], row["location"])]
    for row in deliveries:
        pools[(row["material"], row["location"])]
    result = []
    for (identifier, location), rows in sorted(pools.items()):
        if material_id and identifier != material_id:
            continue
        policy = found(db, "p4_material", identifier)
        opening = stock(db, identifier, location, start)
        movements = stock(db, identifier, location, end)["moves"]
        here = [item for item in deliveries if item["material"] == identifier and item["location"] == location]
        fronts = sorted({row["workfront"] for row in rows} | {key for key, amount in opening["allocations"].items() if amount}
                        | {row["workfront"] for row in here if row["workfront"]})
        owned = dict(opening["allocations"])
        shared = opening["available"]
        backlog = defaultdict(int)
        reasons = []
        if not opening["opening_known"]:
            reasons.append("Opening stock count missing")
        if policy["lead"] is None or policy["safety"] is None:
            reasons.append("Lead time or safety stock missing")
        if not rows:
            reasons.append("No programme in this period")
        takeoffs = {}
        for row in rows:
            take = found(db, "p4_takeoff", row["takeoff"])
            takeoffs[take["id"]] = take
            if (not row["reviewer"] or row["rate"] is None or row["daily"] is None
                    or not quantity_current(db, take)):
                reasons.append(f'Unverified programme or drawing: {row["id"]}')
            if not db.execute("SELECT 1 FROM workfronts WHERE sequence=? AND retired=0",
                              (int(row["workfront"][-6:]),)).fetchone():
                reasons.append(f'Retired workfront: {row["workfront"]}')
        data = defaultdict(list)
        first_safety = None
        for current in horizon:
            when = current.isoformat()
            incoming = [item for item in here if item["confirmed"] and item["eta"] == when and item["remaining"]]
            for item in incoming:
                if item["workfront"]:
                    key = item["workfront"]
                    owned[key] = owned.get(key, 0) + item["remaining"]
                else:
                    shared += item["remaining"]
            for front in fronts or [""]:
                selected = [row for row in rows if row["workfront"] == front and due(row, current)]
                demand = sum(consumption(row) for row in selected if row["rate"] is not None and row["daily"] is not None)
                available = shared + owned.get(front, 0)
                if not reasons:
                    need = backlog[front] + demand
                    from_reserved = min(need, owned.get(front, 0))
                    owned[front] = owned.get(front, 0) - from_reserved
                    from_shared = min(need - from_reserved, shared)
                    shared -= from_shared
                    backlog[front] = need - from_reserved - from_shared
                data[front].append({"date": when, "demand": None if reasons else demand,
                                    "incoming": sum(item["remaining"] for item in incoming
                                                    if not item["workfront"] or item["workfront"] == front),
                                    "available": None if reasons else available,
                                    "deficit": None if reasons else backlog[front]})
            balance = shared + sum(owned.values()) - sum(backlog.values())
            free_balance = shared - sum(backlog.values())
            if not reasons and free_balance < policy["safety"] and first_safety is None:
                first_safety = when
            for front in fronts or [""]:
                data[front][-1]["projected"] = None if reasons else balance
                data[front][-1]["unallocated"] = None if reasons else free_balance
        for front in fronts or [""]:
            if workfront and front != workfront:
                continue
            own_plans = [row for row in rows if row["workfront"] == front]
            front_reasons = reasons + ([] if own_plans else ["No workfront demand plan in this period"])
            daily = data[front]
            first = next((row["date"] for row in daily if row["deficit"]), None)
            overdue = any(item["overdue"] and item["workfront"] in (None, front) for item in here)
            level = "UNKNOWN" if front_reasons else "RED" if first else "AMBER" if first_safety or overdue else "GREEN"
            if status and status != level:
                continue
            required_day = first or first_safety
            requisition_day = ((date.fromisoformat(required_day) - timedelta(days=policy["lead"])).isoformat()
                               if required_day and policy["lead"] is not None and not front_reasons else None)
            references = []
            for row in own_plans:
                take = takeoffs[row["takeoff"]]
                drawing = found(db, "p4_drawing", take["drawing"])
                references.append({"plan": row["id"], "programme": row["source"], "snapshot": row["snapshot"],
                                   "drawing": f'{drawing["number"]} rev {drawing["revision"]}',
                                   "drawing_ref": drawing["reference"], "drawing_snapshot": drawing["snapshot"],
                                   "takeoff": take["id"], "quantity_ref": take["source"]})
            result.append({
                "material": identifier, "description": policy["description"], "specification": policy["specification"],
                "unit": policy["unit"], "location": location, "workfront": front, "status": level,
                "reasons": front_reasons, "stockout": first, "requisition_by": requisition_day,
                "safety_stock": show(policy["safety"]), "safety_breach": first_safety,
                "shortage": None if front_reasons else show(max(row["deficit"] for row in daily)),
                "demand": None if front_reasons else show(sum(row["demand"] for row in daily)),
                "confirmed": show(sum(row["incoming"] for row in daily)), "start": start, "end": end,
                "stock": {key: show(opening[key]) if opening["opening_known"] else None
                          for key in ("physical", "usable", "quarantine", "reserved", "available")},
                "deliveries": [{**item, **{key: show(item[key]) for key in
                                          ("ordered", "remaining", "accepted", "rejected", "quarantined")}}
                               for item in here if item["workfront"] in (None, front)],
                "references": references,
                "variance": variance_info(db, policy, movements, front, location, start, end,
                                          not front_reasons, daily),
                "daily": [{key: value if key == "date" else show(value) for key, value in record.items()}
                          for record in daily],
            })
    return result
