import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from kvd.demo import REPORTS, seed
from kvd.intake import parse
from kvd.journal import fingerprint
from kvd.proposals import approve, preview


class SiteIntake(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = TemporaryDirectory()
        cls.addClassCleanup(cls.temp.cleanup)
        cls.workbook, cls.journal = seed(Path(cls.temp.name) / "demo")
        cls.text = REPORTS.read_text(encoding="utf-8")

    def setUp(self):
        self.drafts = parse(self.text)
        with self.journal.transaction() as db:
            db.execute("DELETE FROM events WHERE event LIKE 'p2_%'")

    def preview(self, drafts=None, decisions=None, workbook=None, actor="Synthetic Operator"):
        drafts = self.drafts if drafts is None else drafts
        decisions = [{} for _ in drafts] if decisions is None else decisions
        return preview(self.journal, workbook or self.workbook, drafts, decisions, actor)

    def test_five_reports_four_commands_and_audited_proposal(self):
        original = fingerprint(self.workbook)
        token, draft = self.preview()
        self.assertEqual(len(draft["records"]), 5)
        self.assertEqual({row["command"] for row in draft["records"]}, {"MC01", "MC14", "MC11", "MC12"})
        cells = {(row["sheet_index"], row["cell"]): row["after"] for row in draft["changes"]}
        self.assertEqual(cells[(5, "K5")], 18)
        self.assertEqual(cells[(5, "K6")], 12.5)
        self.assertEqual(cells[(6, "D5")], cells[(5, "A5")])
        self.assertNotIn((6, "K5"), cells)
        self.assertNotIn((5, "L5"), cells)
        self.assertNotIn((5, "L6"), cells)
        self.assertEqual(cells[(20, "K5")], self.drafts[1].fields["outcome"])
        self.assertNotIn((20, "I5"), cells)
        self.assertEqual(cells[(16, "B6")], cells[(15, "A5")])
        self.assertEqual(cells[(16, "J6")], "Pending")
        self.assertEqual(cells[(19, "K5")], self.drafts[3].fields["blocker"])
        for index, address in ((19, "F5"), (19, "L5"), (19, "L6"), (17, "L5")):
            self.assertNotIn((index, address), cells)
        self.assertNotIn((19, "O5"), cells)
        self.assertNotIn((19, "M5"), cells)
        self.assertNotIn((17, "M5"), cells)
        approved = approve(self.journal, token, "Synthetic Reviewer")
        self.assertEqual(approved["changes"], draft["changes"])
        self.assertEqual(fingerprint(self.workbook), original)
        with self.journal.transaction() as db:
            item = db.execute("SELECT actor, detail FROM events WHERE event='p2_proposal_approved'").fetchone()
        self.assertEqual(item["actor"], "Synthetic Reviewer")
        self.assertEqual(json.loads(item["detail"])["records"][0]["source"], self.drafts[0].source)

    def test_replay_is_rejected_and_suppression_is_journaled(self):
        token, _ = self.preview()
        approve(self.journal, token, "Synthetic Reviewer")
        with self.assertRaisesRegex(ValueError, "Exact source replay"):
            self.preview()
        token, item = self.preview(decisions=[{"action": "suppress", "reason": "Replay"}] * 5)
        self.assertEqual(item["changes"], [])
        approve(self.journal, token, "Synthetic Reviewer")

    def test_within_batch_match_and_distinct_override(self):
        other = parse(self.drafts[0].source.replace("18", "20"))[0]
        drafts = [self.drafts[0], other]
        with self.assertRaisesRegex(ValueError, "Possible duplicate"):
            self.preview(drafts)
        _, item = self.preview(drafts, [{}, {"action": "keep distinct", "reason": "Separate synthetic load"}])
        self.assertEqual(len(item["records"][1]["duplicates"]), 1)

    def test_existing_record_correction_preserves_verified_quantity(self):
        with TemporaryDirectory() as temp:
            workbook = Path(temp) / "copy.xlsx"
            book = load_workbook(self.workbook)
            try:
                fields = {1: "SYN-EXIST", 2: "2026-01-15", 3: "Day", 4: "SYN-CUT", 5: "SYN-BQ-1",
                          8: "Mock Crew A", 9: "m3", 11: 16, 12: 4, 18: "KVDT2-WF-000001"}
                for column, value in fields.items():
                    book.worksheets[5].cell(5, column, value)
                book.save(workbook)
            finally:
                book.close()
            self.drafts[0].fields.pop("resources")
            _, item = self.preview([self.drafts[0]], [{"action": "correct existing", "reason": "New measurement",
                                                        "record_id": "SYN-EXIST"}], workbook)
            change = next(change for change in item["changes"] if change["cell"] == "K5")
            self.assertEqual((change["before"], change["after"]), (16, 18))
            self.assertFalse(any(change["cell"] == "L5" for change in item["changes"]))

    def test_verified_zero_requires_separate_evidence_and_reviewer(self):
        token, item = self.preview([self.drafts[0]], [{"verified_qty": "0", "verification_evidence": "SYN-QA-1"}])
        self.assertEqual(next(change["after"] for change in item["changes"] if change["cell"] == "L5"), 0)
        with self.assertRaisesRegex(ValueError, "distinct quantity reviewer"):
            approve(self.journal, token, "Synthetic Operator", True)
        with self.assertRaisesRegex(ValueError, "distinct quantity reviewer"):
            approve(self.journal, token, "Synthetic Reviewer")
        approve(self.journal, token, "Synthetic Reviewer", True)

    def test_stale_approval_and_row_conflicts(self):
        first, _ = self.preview([self.drafts[0]])
        second, _ = self.preview([self.drafts[1]])
        approve(self.journal, first, "Synthetic Reviewer")
        with self.assertRaisesRegex(ValueError, "refresh the preview"):
            approve(self.journal, second, "Synthetic Reviewer")
        new, item = self.preview([self.drafts[1]])
        self.assertTrue(any(change["cell"] == "A6" and change["sheet_index"] == 5
                            for change in item["changes"]))
        approve(self.journal, new, "Synthetic Reviewer")

    def test_workbook_revision_change_blocks_approval(self):
        with TemporaryDirectory() as temp:
            workbook = Path(temp) / "copy.xlsx"
            workbook.write_bytes(self.workbook.read_bytes())
            token, _ = self.preview([self.drafts[0]], workbook=workbook)
            with workbook.open("ab") as output:
                output.write(b"synthetic change")
            with self.assertRaisesRegex(ValueError, "refresh the preview"):
                approve(self.journal, token, "Synthetic Reviewer")

    def test_parser_rejects_clearance_and_verified_fields(self):
        for extra in ("\nRelease: Released", "\nVerified qty: 99"):
            with self.assertRaisesRegex(ValueError, "unsupported fields"):
                parse(self.drafts[0].source + extra)


if __name__ == "__main__":
    unittest.main()
