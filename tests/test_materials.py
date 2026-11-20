from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import unittest

from docx import Document

from kvd import material_requisition, materials, reports
from kvd.journal import Journal, fingerprint
from kvd.material_demo import ANCHOR, EAST, LOCATION, NORTH, fixtures
from kvd.material_forecast import forecast
from kvd.material_store import design_quantity, show, stock


class MaterialControl(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        directory = TemporaryDirectory()
        cls.addClassCleanup(directory.cleanup)
        cls.workbook, cls.seed = fixtures(Path(directory.name))
        cls.sha = fingerprint(cls.workbook)

    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.directory = Path(directory.name)
        journal_path = self.directory / "audit.sqlite3"
        journal_path.write_bytes(self.seed.path.read_bytes())
        self.journal = Journal(journal_path)

    def action(self, operation, actor="Synthetic Operator", **values):
        return materials.apply(self.journal, actor, [{"operation": operation, **values}])[0]

    def project(self, **filters):
        return materials.read(self.journal, forecast, start=filters.pop("start", ANCHOR), **filters)

    def move(self, id, kind, qty, **overrides):
        values = {"material_id": "SYN-REBAR", "unit": "kg", "location": LOCATION,
                  "date": ANCHOR, "source": "Synthetic source", "workfront": NORTH}
        values.update(overrides)
        return self.action("movement", id=id, kind=kind, quantity=qty, **values)

    def test_four_statuses_date_and_forecast_values(self):
        rows = self.project()
        self.assertEqual({row["status"] for row in rows}, {"GREEN", "AMBER", "RED", "UNKNOWN"})
        red, = self.project(workfront=NORTH, material_id="SYN-REBAR", status="RED")
        self.assertEqual(len(red["daily"]), 14)
        self.assertEqual(red["daily"][2]["demand"], "0")
        self.assertEqual(red["demand"], "200")
        self.assertEqual(red["stockout"], "2026-01-19")
        self.assertEqual(red["requisition_by"], "2026-01-14")
        self.assertEqual(red["shortage"], "129")
        self.assertEqual(red["confirmed"], "25")
        self.assertEqual(red["references"][0]["snapshot"], self.sha)
        unknown, = self.project(status="UNKNOWN")
        self.assertIsNone(unknown["demand"])
        self.assertIsNone(unknown["shortage"])
        self.assertEqual(self.project(start="2026-02-01", status="RED"), [])

    def test_geometry_and_units(self):
        dims = {"l": {"value": 3, "unit": "m"}, "w": {"value": 2, "unit": "m"},
                "d": {"value": "0.5", "unit": "m"}}
        self.assertEqual(show(design_quantity("l*w*d", dims, "m3")), "3")
        with self.assertRaisesRegex(ValueError, "dimension names"):
            design_quantity("l*w*h", dims, "m3")
        dims["w"]["unit"] = "mm"
        with self.assertRaisesRegex(ValueError, "dimensions do not match"):
            design_quantity("l*w*d", dims, "m3")
        with self.assertRaisesRegex(ValueError, "Unit mismatch"):
            self.move("wrong", "issue", 2, unit="m")
        takeoff = next(row for row in materials.read(self.journal, materials.list_takeoffs) if row["id"] == "TO-REBAR")
        self.assertEqual((takeoff["design"], takeoff["allowance"], takeoff["procurement"]), ("300", "15", "315"))

    def test_drawing_revision_invalidates_previous_demand(self):
        self.action("drawing", number="SYN-DR-REBAR", revision="B", status="PENDING", reference="SYN-RB")
        row, = self.project(material_id="SYN-REBAR")
        self.assertEqual(row["status"], "UNKNOWN")
        self.assertIn("drawing", row["reasons"][0])
        self.assertEqual(self.action("prepare-requisitions", start=ANCHOR), [])

    def test_allowance_and_verification_are_separate(self):
        with self.assertRaisesRegex(ValueError, "Verify the draft"):
            self.action("takeoff-review", id="TO-PANEL", action="approve", reference="SYN-GATE")
        with self.assertRaisesRegex(ValueError, "superseded or unapproved"):
            self.action("takeoff-review", id="TO-PANEL", action="verify", reference="SYN-GATE")

    def test_reservations_and_concurrent_allocation(self):
        def attempt(item):
            try:
                self.move(f"reserve-{item}", "reserve", 30, workfront=EAST)
                return True
            except ValueError:
                return False

        with ThreadPoolExecutor(max_workers=2) as worker:
            self.assertEqual(sum(worker.map(attempt, (1, 2))), 1)
        with self.assertRaisesRegex(ValueError, "Insufficient"):
            self.move("issue-restricted", "issue", 20, workfront=NORTH)
        state = materials.read(self.journal, stock, material_id="SYN-REBAR", location=LOCATION)
        self.assertEqual([show(state[k]) for k in ("physical", "usable", "quarantine", "reserved", "available")],
                         ["51", "46", "5", "38", "8"])
        with self.assertRaisesRegex(sqlite3.IntegrityError, "UNIQUE"):
            self.move("ISSUE-01", "issue", 1)
        self.move("q-accepted", "quarantine-release", 5, workfront=None)
        with self.assertRaisesRegex(ValueError, "quarantined"):
            self.move("q-double", "quarantine-release", 5, workfront=None)

    def test_partial_late_and_reconfirmed_deliveries(self):
        first, = [row for row in self.project(material_id="SYN-REBAR")[0]["deliveries"] if row["id"] == "ORDER-01"]
        self.assertTrue(first["overdue"])
        self.assertFalse(first["confirmed"])
        self.assertEqual((first["accepted"], first["rejected"], first["quarantined"], first["remaining"]),
                         ("8", "2", "5", "15"))
        self.action("order-confirm", id="ORDER-01", eta="2026-01-17", confirmed_on=ANCHOR, reference="SYN-ETA")
        self.assertEqual(self.project(material_id="SYN-REBAR")[0]["confirmed"], "40")
        self.action("arrival", id="DN-SYN-02", order="ORDER-01", unit="kg", date=ANCHOR, reference="SYN-DN",
                    accepted=12, rejected=1, quarantined=2)
        first, = [row for row in self.project(material_id="SYN-REBAR")[0]["deliveries"] if row["id"] == "ORDER-01"]
        self.assertTrue(first["last_arrival_late"])
        self.assertFalse(first["overdue"])
        self.assertEqual(first["remaining"], "0")
        with self.assertRaisesRegex(ValueError, "exceeds"):
            self.action("arrival", id="DN-SYN-03", order="ORDER-01", unit="kg", date=ANCHOR,
                        reference="SYN-DN", accepted=1, rejected=0, quarantined=0)

    def test_actual_consumption_is_not_issued_quantity(self):
        self.move("excess-issue", "issue", 31, workfront=EAST, material_id="SYN-GROUT")
        self.move("excess-use", "consume", 29, workfront=EAST, material_id="SYN-GROUT")
        values = self.project(material_id="SYN-GROUT")[0]["variance"]
        self.assertEqual((values["planned"], values["actual"], values["design"], values["issued"]),
                         ("25", "29", "30", "31"))
        self.assertTrue(values["potential_wastage"])
        self.assertFalse(values["design_over_issue"])
        with self.assertRaisesRegex(ValueError, "already issued"):
            self.move("not-issued", "consume", 3, workfront=EAST, material_id="SYN-GROUT")

    def test_requisition_duplicate_revision_approval_and_export(self):
        first, = self.action("prepare-requisitions", start=ANCHOR)
        self.assertEqual(self.action("prepare-requisitions", start=ANCHOR), [])
        with self.assertRaisesRegex(ValueError, "reviewed"):
            self.action("requisition-action", id=first, action="issue", reference="SYN-ISSUE")
        second = self.action("revise-requisition", id=first, start=ANCHOR, reference="SYN-REVISION")
        with self.assertRaisesRegex(ValueError, "separate named"):
            self.action("requisition-action", id=second, action="review", reference="SYN-REVIEW")
        for action, person in (("review", "Mock Reviewer"), ("approve", "Mock Approver"), ("issue", "Mock Issuer")):
            self.action("requisition-action", actor=person, id=second, action=action, reference=f"SYN-{action}")
        rows = materials.read(self.journal, materials.list_requisitions)
        self.assertEqual([(r["revision"], r["state"]) for r in rows], [(2, "ISSUED"), (1, "SUPERSEDED")])
        result = Path(materials.read(self.journal, material_requisition.export, id=second, directory=self.directory))
        self.assertIn("Quantity requested: 129 kg",
                      "\n".join(p.text for p in Document(next(result.glob("*.docx"))).paragraphs))
        self.assertTrue(next(result.glob("*.pdf")).read_bytes().startswith(b"%PDF-"))
        self.action("order", id="PO-REQ", material_id="SYN-REBAR", unit="kg", location=LOCATION, po="SYN-PO-RQ",
                    quantity=129, placed=ANCHOR, due="2026-01-21", reference="SYN-PO-EVIDENCE",
                    workfront=NORTH, requisition=second)
        with self.assertRaisesRegex(ValueError, "exceed this requisition"):
            self.action("order", id="PO-REQ-2", material_id="SYN-REBAR", unit="kg", location=LOCATION,
                        po="SYN-PO-RQ2", quantity=1, placed=ANCHOR, due="2026-01-21",
                        reference="SYN-PO-EVIDENCE", workfront=NORTH, requisition=second)
        self.action("requisition-action", id=second, action="close", reference="SYN-CLOSED")
        self.assertEqual(self.action("prepare-requisitions", start=ANCHOR), [])

    def test_issue_rechecks_supply_and_source(self):
        identifier, = self.action("prepare-requisitions", start=ANCHOR)
        self.action("requisition-action", actor="Mock Reviewer", id=identifier,
                    action="review", reference="SYN-REVIEW")
        self.action("requisition-action", actor="Mock Approver", id=identifier,
                    action="approve", reference="SYN-APPROVE")
        self.move("additional-stock", "receipt", 500, workfront=None)
        with self.assertRaisesRegex(ValueError, "no verified shortage"):
            self.action("requisition-action", actor="Mock Issuer", id=identifier,
                        action="issue", reference="SYN-ISSUE")

    def test_missing_policy_and_missing_opening_are_unknown(self):
        self.action("material-policy", id="SYN-GROUT", lead=None, safety=15, variance=10,
                    reference="SYN-POLICY")
        grout, = self.project(material_id="SYN-GROUT")
        self.assertEqual(grout["status"], "UNKNOWN")
        self.assertIn("Lead time or safety stock missing", grout["reasons"])
        with self.assertRaisesRegex(ValueError, "opening stock"):
            self.move("unknown-opening", "reserve", 1, location="Other mock store")

    def test_programme_import_unit_and_calendar_exclusion(self):
        with self.assertRaisesRegex(ValueError, "production unit"):
            self.action("programme-import", workbook=str(self.workbook), row=5,
                        production_unit="m3", id="UNITS", takeoff="TO-REBAR",
                        location=LOCATION, workdays=[0, 1, 2, 3, 4], rate=2, unit="kg")
        self.action("programme-retire", id="PLAN-GROUT", reference="SYN-CALENDAR")
        self.action("programme", id="REVISED", takeoff="TO-GROUT", activity="SYN-DRAIN",
                    location=LOCATION, start=ANCHOR, finish="2026-01-21",
                    workdays=[0, 1, 2, 3, 4], exclusions=["2026-01-19"], rate=1, daily=5,
                    unit="kg", production_unit="m", source="Synthetic holiday revision")
        self.action("programme-verify", actor="Mock Reviewer", id="REVISED", reference="SYN-WORKDAYS")
        grout, = self.project(material_id="SYN-GROUT")
        self.assertEqual(grout["demand"], "20")
        self.assertEqual(grout["daily"][4]["demand"], "0")

    def test_p3_reports_on_same_workbook_and_journal(self):
        identifier = reports.create(self.journal, self.workbook, "MC10", ANCHOR, "Mock Author")
        reports.approve(self.journal, identifier, "Mock Reviewer")
        reports.issue(self.journal, identifier, "Mock Issuer")
        self.assertEqual(reports.get(self.journal, identifier)[0], "issued")
        self.assertEqual(fingerprint(self.workbook), self.sha)
        self.assertEqual(len(self.project()), 4)


if __name__ == "__main__":
    unittest.main()
