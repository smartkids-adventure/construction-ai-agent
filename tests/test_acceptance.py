import sqlite3
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import load_workbook

from kvd.journal import Journal, fingerprint, migrate
from kvd.workbook import baseline, make_fixture, schema, validate
from kvd.workfront import create, resolve, retire


class Acceptance(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.source = self.folder / "synthetic.xlsx"
        self.output = self.folder / "synthetic-1000.xlsx"
        self.journal = Journal(self.folder / "audit.sqlite3")
        self.actor = "Synthetic Reviewer"
        make_fixture(self.source)

    def edit(self, operation):
        book = load_workbook(self.source)
        try:
            operation(book)
            book.save(self.source)
        finally:
            book.close()

    def migrate(self):
        return migrate(self.source, self.output, self.journal, self.actor, fingerprint(self.source))

    def test_migrate_preserves_source_and_values_and_expands_boundaries(self):
        self.edit(lambda book: book.worksheets[5].cell(5, 1, "SYN-001"))
        original = fingerprint(self.source)
        self.assertEqual(self.migrate(), 1)
        self.assertEqual(fingerprint(self.source), original)
        self.assertEqual(fingerprint(self.folder / "backups" / f"{original}.xlsx"), original)
        book = load_workbook(self.output)
        try:
            validate(book, expanded=True)
            observed = schema(book)
            self.assertEqual(len(observed), 21)
            self.assertEqual(sum(len(entry["tables"]) for entry in observed.values()), 19)
            self.assertEqual(sum(len(entry["validation"]) for entry in observed.values()), 21)
            self.assertEqual(sum(len(entry["formulas"]) for entry in observed.values()), 11016)
            for name, entry in baseline().items():
                self.assertTrue(entry["formulas"].keys() <= observed[name]["formulas"].keys())
            self.assertEqual(book.worksheets[5]["A5"].value, "SYN-001")
            self.assertEqual(book.worksheets[5]["R4"].value, "Workfront ID")
            self.assertIn("A5:A1004", book.worksheets[0]["B5"].value)
            for row in (105, 1004):
                self.assertIn(f"I{row}:J{row}", book.worksheets[6][f"K{row}"].value)
                self.assertIn(f"F{row}", book.worksheets[19][f"O{row}"].value)
                self.assertIn("$L$5:$L$1004", book.worksheets[10][f"H{row}"].value)
                self.assertIn("$L$5:$L$1004", book.worksheets[12][f"E{row}"].value)
            self.assertEqual(str(book.worksheets[2].data_validations.dataValidation[0].sqref),
                             "H5:H1004")
        finally:
            book.close()

    def test_reject_formula_drift(self):
        self.edit(lambda book: setattr(book.worksheets[19]["O5"], "value", "=1"))
        with self.assertRaisesRegex(ValueError, "formulas"):
            self.migrate()
        self.assertFalse(self.output.exists())

    def test_reject_invalid_list_entry(self):
        self.edit(lambda book: setattr(book.worksheets[17]["H5"], "value", "Synthetic invalid"))
        with self.assertRaisesRegex(ValueError, "invalid list value"):
            self.migrate()

    def test_reject_table_header_drift(self):
        self.edit(lambda book: setattr(book.worksheets[5]["A4"], "value", "Wrong"))
        with self.assertRaisesRegex(ValueError, "headers"):
            self.migrate()

    def test_stale_approval_and_existing_output_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "fingerprint"):
            migrate(self.source, self.output, self.journal, self.actor, "bad")
        self.output.write_bytes(b"existing output")
        with self.assertRaisesRegex(ValueError, "new output"):
            self.migrate()
        self.assertEqual(self.output.read_bytes(), b"existing output")

    def test_recover_after_publication_before_journal_commit(self):
        log = Journal.log

        def interrupt(db, event, actor, detail):
            if event == "migration_committed":
                raise RuntimeError("Synthetic interruption")
            log(db, event, actor, detail)

        with patch.object(Journal, "log", side_effect=interrupt):
            with self.assertRaisesRegex(RuntimeError, "Synthetic interruption"):
                self.migrate()
        self.assertTrue(self.output.exists())
        self.assertEqual(self.journal.recover(self.actor), [(1, "committed")])
        self.assertEqual(self.journal.recover(self.actor), [])

    def test_workfront_identifiers_are_stable_and_aliases_unique(self):
        args = (self.journal, self.actor, "Synthetic Section", "Synthetic km 1", ["SYN-A"])
        identifier = create(*args, ["Mock  Workfront"])
        self.assertEqual(identifier, "KVDT2-WF-000001")
        self.assertEqual(resolve(self.journal, " mock workfront ")["id"], identifier)
        with self.assertRaisesRegex(ValueError, "already allocated"):
            create(*args, ["MOCK WORKFRONT"])
        retire(self.journal, self.actor, identifier)
        with self.assertRaisesRegex(ValueError, "retired"):
            resolve(self.journal, identifier)
        self.assertEqual(create(*args, ["Other Mock Front"]), "KVDT2-WF-000002")
        with sqlite3.connect(self.journal.path) as db:
            self.assertEqual(db.execute("SELECT count(*) FROM events").fetchone()[0], 3)

    def test_reject_unmarked_workbook(self):
        self.edit(lambda book: setattr(book.worksheets[1]["A1"], "value", None))
        with self.assertRaisesRegex(ValueError, "synthetic"):
            self.migrate()


if __name__ == "__main__":
    unittest.main()
