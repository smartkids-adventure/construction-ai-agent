from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from docx import Document

from kvd import reports
from kvd.journal import fingerprint
from kvd.report_data import approved_snapshot, comparison, select
from kvd.report_demo import ANCHOR, snapshot


class Reporting(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        folder = TemporaryDirectory()
        cls.addClassCleanup(folder.cleanup)
        cls.directory = Path(folder.name)
        cls.workbook, cls.journal = snapshot(cls.directory / "demo")
        cls.snapshot = approved_snapshot(cls.journal, cls.workbook)

    def setUp(self):
        with self.journal.transaction() as db:
            reports.setup(db)
            db.execute("DELETE FROM reports")
            db.execute("DELETE FROM events WHERE event LIKE 'p3_%'")

    def draft(self, command="MC10", anchor=ANCHOR):
        return reports.create(self.journal, self.workbook, command, anchor, "Synthetic Author")

    def test_approved_source_and_fingerprint_gate(self):
        with self.assertRaisesRegex(ValueError, "committed synthetic revision"):
            approved_snapshot(self.journal, self.workbook.with_name("synthetic-five.xlsx"))
        with TemporaryDirectory() as temp:
            copy = Path(temp) / "copy.xlsx"
            copy.write_bytes(self.workbook.read_bytes())
            with self.assertRaisesRegex(ValueError, "committed synthetic revision"):
                approved_snapshot(self.journal, copy)
        with patch("kvd.report_data.Path.read_bytes", return_value=self.workbook.read_bytes() + b"tamper"):
            with self.assertRaisesRegex(ValueError, "committed synthetic revision"):
                approved_snapshot(self.journal, self.workbook)

    def test_dashboard_filters_and_distinct_units(self):
        rows = select(self.snapshot, from_date=ANCHOR, to_date=ANCHOR, workfront="kvdt2-wf-000002",
                      activity="SYN-DRAIN", subcontractor="Mock Crew B", status="Verified")
        self.assertEqual(len(rows["progress"]), 1)
        self.assertEqual(rows["progress"][0]["verified"], 0)
        self.assertEqual(rows["technical"], [])
        totals = comparison(select(self.snapshot, from_date=ANCHOR, to_date=ANCHOR)["progress"])
        self.assertEqual({item["dimensions"][-1] for item in totals}, {"m3", "m"})
        east = next(item for item in totals if item["dimensions"][-1] == "m")
        self.assertEqual(east["sums"], {"planned": 14, "reported": 12.5, "verified": 0})
        self.assertEqual(east["counts"], {"planned": 1, "reported": 2, "verified": 1})
        with self.assertRaisesRegex(ValueError, "Start date"):
            select(self.snapshot, from_date="2026-01-16", to_date=ANCHOR)

    def test_daily_and_weekly_windows_and_open_registers(self):
        daily = reports.get(self.journal, self.draft())[2]
        weekly = reports.get(self.journal, self.draft("MC02"))[2]
        self.assertEqual((daily["start"], daily["end"]), (ANCHOR, ANCHOR))
        self.assertEqual((weekly["start"], weekly["end"]), ("2026-01-12", "2026-01-18"))
        self.assertEqual(len(weekly["sections"]["progress"]), len(daily["sections"]["progress"]) + 1)
        self.assertEqual(len(weekly["sections"]["technical"]), 2)
        self.assertEqual(len(daily["sections"]["quality"]), 2)
        self.assertEqual(len(daily["sections"]["actions"]), 1)
        self.assertEqual(len(daily["sections"]["tasks"]), 1)
        self.assertEqual(len(daily["sections"]["toolbox"]), 3)
        self.assertEqual(len(daily["sections"]["utility"]), 3)
        self.assertEqual({row["status"] for row in daily["sections"]["readiness"]}, {"Pending review"})
        self.assertTrue(all(row["status"] == "Not recorded" for row in daily["sections"]["utility"]))

    def test_approval_issue_and_both_artifacts(self):
        original = fingerprint(self.workbook)
        for command in ("MC10", "MC02"):
            identifier = self.draft(command)
            with self.assertRaisesRegex(ValueError, "Approve the revision"):
                reports.issue(self.journal, identifier, "Synthetic Issuer")
            with self.assertRaisesRegex(ValueError, "different named reviewer"):
                reports.approve(self.journal, identifier, "Synthetic Author")
            reports.approve(self.journal, identifier, "Synthetic Reviewer")
            reports.issue(self.journal, identifier, "Synthetic Issuer")
            path = reports.export(self.journal, identifier, self.directory)
            self.assertEqual(reports.export(self.journal, identifier, self.directory), path)
            doc = Document(BytesIO(next(path.glob("*.docx")).read_bytes()))
            self.assertIn("Approved by Synthetic Reviewer", "\n".join(p.text for p in doc.paragraphs))
            self.assertEqual(doc.tables[0].rows[0].cells[-3].text, "Planned")
            self.assertTrue(next(path.glob("*.pdf")).read_bytes().startswith(b"%PDF-"))
        self.assertEqual(fingerprint(self.workbook), original)

    def test_revision_preserves_issued_data(self):
        first = self.draft()
        reports.approve(self.journal, first, "Synthetic Reviewer")
        reports.issue(self.journal, first, "Synthetic Issuer")
        second = self.draft()
        revisions = reports.history(self.journal)
        self.assertEqual([(r["id"], r["revision"], r["state"]) for r in revisions],
                         [(second, 2, "draft"), (first, 1, "issued")])

    def test_render_failure_keeps_approved_revision_retryable(self):
        identifier = self.draft()
        reports.approve(self.journal, identifier, "Synthetic Reviewer")
        with patch("kvd.reports.pdf", side_effect=ValueError("Synthetic render failure")):
            with self.assertRaisesRegex(ValueError, "Synthetic render failure"):
                reports.issue(self.journal, identifier, "Synthetic Issuer")
        self.assertEqual(reports.get(self.journal, identifier)[0], "approved")
        reports.issue(self.journal, identifier, "Synthetic Issuer")
        self.assertEqual(reports.get(self.journal, identifier)[0], "issued")


if __name__ == "__main__":
    unittest.main()
