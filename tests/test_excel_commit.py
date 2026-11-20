import json
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from kvd.demo import REPORTS, seed
from kvd.excel_commit import commit, recover
from kvd.intake import parse
from kvd.journal import Journal, fingerprint
from kvd.proposals import approve, checksum, preview
from kvd.workbook import schema


class ExcelCommit(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        folder = TemporaryDirectory()
        cls.addClassCleanup(folder.cleanup)
        cls.template, cls.demo_journal = seed(Path(folder.name))
        cls.malay = (REPORTS.parent / "synthetic-report-ms.txt").read_text(encoding="utf-8")

    def setUp(self):
        tmp = TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.folder = Path(tmp.name)
        self.source = self.folder / "source.xlsx"
        self.source.write_bytes(self.template.read_bytes())
        journal = self.folder / "journal.sqlite3"
        journal.write_bytes(self.demo_journal.path.read_bytes())
        self.journal = Journal(journal)
        self.output = self.folder / "successor.xlsx"
        self.drafts = parse(self.malay) + parse(REPORTS.read_text(encoding="utf-8"))[1:]

    def approval(self, drafts=None, verification=False):
        records = self.drafts if drafts is None else drafts
        decisions = [{} for _ in records]
        if verification:
            decisions[0] = {"verified_qty": "7", "verification_evidence": "SYN-MEASURE"}
        token, payload = preview(self.journal, self.source, records, decisions, "Synthetic Operator")
        approve(self.journal, token, "Synthetic Reviewer", verification)
        return token, payload

    def run_commit(self, token, output=None):
        return commit(self.journal, token, output or self.output, "Synthetic Operator")

    def test_mixed_batch_reconciles_without_using_reported_quantity(self):
        book = load_workbook(self.source)
        try:
            for index, values in ((10, {"A5": "SYN-CUT", "E5": "m3", "A6": "SYN-DRAIN", "E6": "m"}),
                                  (12, {"A5": "SYN-BQ-1", "C5": "m3"})):
                for address, value in values.items():
                    book.worksheets[index][address] = value
            expected_schema = schema(book)
            book.save(self.source)
        finally:
            book.close()
        original = fingerprint(self.source)
        self.drafts[0].fields["notes"] = "=2+2"
        token, payload = self.approval(verification=True)
        result = self.run_commit(token)
        self.assertEqual(fingerprint(self.source), original)
        self.assertEqual(fingerprint(result["backup"]), original)
        self.assertEqual(fingerprint(self.output), result["output_sha"])
        self.assertEqual(result["reconciliation"]["10"]["expected"], {"H5": 7, "H6": None})
        self.assertEqual(result["reconciliation"]["12"]["expected"], {"E5": 7})
        book = load_workbook(self.output)
        try:
            self.assertEqual(schema(book), expected_schema)
            site = book.worksheets[5]
            self.assertEqual(site["K5"].value, 18)
            self.assertEqual(site["L5"].value, 7)
            self.assertEqual(site["K6"].value, 12.5)
            self.assertIsNone(site["L6"].value)
            self.assertIsInstance(site["B5"].value, datetime)
            self.assertEqual((site["M5"].value, site["M5"].data_type), ("=2+2", "s"))
            for index, cells in ((17, ("L5", "M5")), (18, ("J5",)),
                                 (19, ("F5", "G5", "L5", "M5", "N5"))):
                for cell in cells:
                    self.assertIsNone(book.worksheets[index][cell].value)
        finally:
            book.close()
        with self.journal.transaction() as db:
            event = db.execute("SELECT detail FROM events WHERE event='p21_commit_prepared'").fetchone()
            self.assertEqual(json.loads(event["detail"])["changes"], payload["changes"])

    def test_duplicate_execution_and_second_successor_are_rejected(self):
        first, _ = self.approval(self.drafts[:1])
        second, _ = self.approval(self.drafts[1:2])
        self.run_commit(first)
        for token in (first, second):
            with self.assertRaisesRegex(ValueError, "already has a commit"):
                self.run_commit(token, self.folder / "other.xlsx")

    def test_stale_source_is_rejected(self):
        token, _ = self.approval()
        with self.source.open("ab") as source:
            source.write(b"test mutation")
        with self.assertRaisesRegex(ValueError, "fingerprint changed"):
            self.run_commit(token)
        self.assertFalse(self.output.exists())

    def test_no_overwrite_or_backup_mismatch(self):
        token, payload = self.approval()
        with self.assertRaisesRegex(ValueError, "separate new"):
            self.run_commit(token, self.source)
        backups = self.folder / "backups"
        backups.mkdir()
        (backups / f'{payload["workbook_sha"]}.xlsx').write_bytes(b"incorrect backup")
        with self.assertRaisesRegex(ValueError, "Backup fingerprint mismatch"):
            self.run_commit(token)

    def test_recover_published_revision_after_audit_interruption(self):
        token, _ = self.approval()
        log = Journal.log

        def interrupt(db, event, actor, detail):
            if event == "p21_commit_completed":
                raise RuntimeError("synthetic interruption")
            return log(db, event, actor, detail)

        with patch.object(Journal, "log", side_effect=interrupt):
            with self.assertRaisesRegex(RuntimeError, "synthetic interruption"):
                self.run_commit(token)
        self.assertTrue(self.output.is_file())
        self.assertEqual(recover(self.journal, "Recovery"), [(token, "committed")])
        self.assertEqual(recover(self.journal, "Recovery"), [])

    def test_recover_unpublished_revision_and_retry(self):
        from os import link
        token, _ = self.approval()

        def interrupt(source, output):
            if output == self.output:
                raise RuntimeError("synthetic interruption")
            return link(source, output)

        with patch("kvd.excel_commit.os.link", side_effect=interrupt):
            with self.assertRaisesRegex(RuntimeError, "synthetic interruption"):
                self.run_commit(token)
        self.assertEqual(recover(self.journal, "Recovery"), [(token, "aborted")])
        self.run_commit(token)
        self.assertTrue(self.output.is_file())

    def test_protected_cell_is_rejected_even_with_matching_approval(self):
        _, payload = preview(self.journal, self.source, self.drafts, [{}] * 5,
                             "Synthetic Operator")
        change = dict(payload["changes"][0])
        book = load_workbook(self.source, read_only=True)
        try:
            change.update(sheet_index=19, sheet=book.sheetnames[19],
                          cell="L5", before=None, after="Released", type="text")
        finally:
            book.close()
        payload["changes"] = [change]
        token = checksum(payload)
        with self.journal.transaction() as db:
            self.journal.log(db, "p2_proposal_previewed", "Synthetic Operator",
                             {"token": token, "payload": payload})
        approve(self.journal, token, "Synthetic Reviewer")
        with self.assertRaisesRegex(ValueError, "protected"):
            self.run_commit(token)

    def test_bilingual_values_are_preserved_and_required_fields_are_not_invented(self):
        bilingual = self.malay.replace("Tarikh:", "Date:")
        drafts = parse(bilingual)
        self.assertEqual(drafts[0].fields["activity"], "Kerja percubaan")
        self.assertEqual(drafts[0].fields["date"], "2026-01-15")
        with self.assertRaisesRegex(ValueError, "Repeated field"):
            parse(self.malay + "Reported qty: 7\n")
        missing = parse("\n".join(line for line in self.malay.splitlines()
                                    if not line.startswith("Kuantiti")))
        with self.assertRaisesRegex(ValueError, "reported qty is required"):
            self.approval(missing)


if __name__ == "__main__":
    unittest.main()
