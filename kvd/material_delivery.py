"""PO promises, partial deliveries and separate quality dispositions."""
import json

from .material_store import (day, found, front, log, material, movement,
                             real_amount, required, show, units)


def order_add(db, journal, actor, id, material_id, unit, location, po, quantity, placed, due,
              reference, workfront=None, requisition=None, eta=None, confirmed_on=None):
    material(db, material_id, unit)
    location = required(location, "Stock location")
    if workfront:
        workfront = front(db, workfront)
    total = real_amount(quantity)
    placed, due = day(placed), day(due)
    if due < placed:
        raise ValueError("Planned delivery cannot precede the order.")
    eta, confirmed_on = confirmation(eta, confirmed_on, placed)
    if requisition is not None:
        request = found(db, "p4_requisition", requisition)
        body = json.loads(request["body"])
        if (request["state"] != "ISSUED" or
                (body["material"], body["location"], body["workfront"]) != (material_id, location, workfront)):
            raise ValueError("PO must match an issued requisition's material, location and workfront.")
        earlier = db.execute("SELECT COALESCE(SUM(ordered),0) FROM p4_order WHERE requisition=?",
                             (str(requisition),)).fetchone()[0]
        if earlier + total > units(body["quantity"]):
            raise ValueError("PO quantities exceed this requisition.")
    db.execute("INSERT INTO p4_order VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
               (required(id, "Order ID"), material_id, location, required(po, "PO reference"),
                str(requisition) if requisition is not None else None, workfront, total, placed, due,
                eta, confirmed_on, required(reference, "Order evidence")))
    log(db, journal, "order_recorded", actor, {"id": id, "po": po, "requisition": requisition, "eta": eta})
    return id


def confirmation(eta, confirmed_on, placed):
    if (eta is None) != (confirmed_on is None):
        raise ValueError("Provide both confirmation date and confirmed ETA.")
    if eta is None:
        return None, None
    eta, confirmed_on = day(eta), day(confirmed_on)
    if not placed <= confirmed_on <= eta:
        raise ValueError("Confirmation date must be between order placement and the ETA.")
    return eta, confirmed_on


def order_confirm(db, journal, actor, id, eta, confirmed_on, reference):
    order = found(db, "p4_order", id)
    eta, confirmed_on = confirmation(eta, confirmed_on, order["placed"])
    if order["confirmed_on"] and (not confirmed_on or confirmed_on < order["confirmed_on"]):
        raise ValueError("New confirmation must be dated on or after the last confirmation.")
    db.execute("UPDATE p4_order SET eta=?,confirmed_on=?,reference=? WHERE id=?",
               (eta, confirmed_on, required(reference, "Delivery confirmation reference"), id))
    log(db, journal, "order_confirmed", actor, {"id": id, "previous_eta": order["eta"], "eta": eta,
                                                "confirmed_on": confirmed_on, "reference": reference})
    return id


def arrivals(db, order, by=None):
    return db.execute("""SELECT * FROM p4_arrival WHERE order_id=? AND (? IS NULL OR day<=?)
                         ORDER BY day,id""", (order, by, by)).fetchall()


def delivered(rows):
    return sum(row["accepted"] + row["rejected"] + row["quarantined"] for row in rows)


def arrival_add(db, journal, actor, id, order, unit, date, reference, accepted, rejected, quarantined):
    item = found(db, "p4_order", order)
    material(db, item["material"], unit)
    when = day(date)
    reference, id = required(reference, "Delivery note reference"), required(id, "Arrival ID")
    prior = arrivals(db, order)
    if when < item["placed"] or (prior and when < prior[-1]["day"]):
        raise ValueError("An arrival must be posted in chronological order after the PO date.")
    parts = {name: units(value) for name, value in
             (("accepted", accepted), ("rejected", rejected), ("quarantined", quarantined))}
    total = sum(parts.values())
    if not total or delivered(prior) + total > item["ordered"]:
        raise ValueError("Delivery exceeds remaining PO quantity or has no units.")
    db.execute("INSERT INTO p4_arrival VALUES (?,?,?,?,?,?,?)",
               (id, order, when, parts["accepted"], parts["rejected"], parts["quarantined"], reference))
    for key, kind in (("accepted", "receipt"), ("quarantined", "quarantine")):
        if parts[key]:
            movement(db, journal, actor, f"delivery:{id}:{kind}", item["material"], unit,
                     item["location"], when, kind, show(parts[key]), reference)
    if parts["accepted"] and item["workfront"]:
        movement(db, journal, actor, f"delivery:{id}:reserve", item["material"], unit, item["location"],
                 when, "reserve", show(parts["accepted"]), reference, item["workfront"])
    log(db, journal, "arrival_recorded", actor, {"id": id, "order": order, "date": when,
                                                 **{key: show(value) for key, value in parts.items()}})
    return id


def delivery_register(db, as_of):
    result = []
    for order in db.execute("SELECT * FROM p4_order WHERE placed<=? ORDER BY due,id", (as_of,)):
        actual = arrivals(db, order["id"], as_of)
        sums = {key: sum(row[key] for row in actual) for key in ("accepted", "rejected", "quarantined")}
        remaining = order["ordered"] - sum(sums.values())
        confirmed = bool(order["eta"] and order["confirmed_on"] <= as_of and order["eta"] >= as_of)
        result.append({**dict(order), **sums, "remaining": remaining, "confirmed": confirmed,
                       "actual_dates": sorted({row["day"] for row in actual}),
                       "overdue": bool(remaining and order["due"] < as_of),
                       "last_arrival_late": bool(actual and actual[-1]["day"] > order["due"])})
    return result
