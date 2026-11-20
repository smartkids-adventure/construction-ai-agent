# Construction AI Management System

P0/P1 is a local Windows Python CLI for synthetic workbook acceptance. It uses no paid API or cloud runtime. The supplied Master Register V2 was read only to extract structural metadata. Generate the schema manifest locally from your reference workbook. The manifest, workbook copies, and journal are excluded from Git.

## Windows installation

Install Python 3.11 or later. Use a local NTFS folder for the workbook copies, audit database, and backups. Use an existing desktop Excel installation for the recalculation check.

Clone the review branch and run these commands in PowerShell from the repository root.

```powershell
git clone --branch codex/mvp-p2-1 https://github.com/smartkids-adventure/construction-ai-agent.git
cd construction-ai-agent
py -3 -m venv .venv
$python = ".\.venv\Scripts\python.exe"
& $python -m pip install -r requirements.txt
& $python -m kvd baseline "C:\path\to\your\Master Register V2.xlsx"
& $python -m unittest discover -s tests -v
```

## Run a synthetic migration

```powershell
New-Item -ItemType Directory -Force .\local
$inputBook = ".\local\synthetic-100.xlsx"
$outputBook = ".\local\synthetic-1000.xlsx"
$audit = ".\local\audit.sqlite3"
& $python -m kvd fixture $inputBook
$approved = & $python -m kvd validate $inputBook
& $python -m kvd --journal $audit migrate $inputBook $outputBook --actor "Synthetic Reviewer" --approved-sha $approved
& $python -m kvd validate $outputBook --expanded
.\tests\excel-boundary.ps1 -WorkbookPath $outputBook
```

Inspect the fixture and validate its fingerprint before passing that fingerprint as approval. Use a fresh output filename for each revision. The migration expands all 19 tables and 21 validation ranges through row 1004, widens the dependent formula ranges, and fills the existing formula patterns. Daily Site Data gains Workfront ID in column R. The original fixture stays in place, with a verified copy under `local/backups`. The Excel acceptance script opens the result read-only, tests rows 5, 105, and 1004 in memory, and discards its edits.

If publication was interrupted, close Excel and run:

```powershell
& $python -m kvd --journal $audit recover --actor "Synthetic Reviewer"
```

A matching output is marked committed. A missing output is marked aborted and the migration can be retried. A mismatched output remains pending for manual comparison. Restore a verified backup by copying it to a fresh filename, then run validation. Keep the database with its backups. The named actor is a P0/P1 audit label. Use the baseline command once per checkout, from the repository root. The generated workbook-baseline.json contains reference metadata and stays local.

## Workfront index

```powershell
& $python -m kvd --journal $audit workfront-add --actor "Synthetic Reviewer" --section "Mock Section" --location "Mock km 1" --activity SYN-A --alias "Mock Front"
& $python -m kvd --journal $audit workfront-find "Mock Front"
& $python -m kvd --journal $audit workfront-retire KVDT2-WF-000001 --actor "Synthetic Reviewer"
```

IDs and aliases are unique within this project's journal. Keep one authoritative local database. Retirement preserves the allocation. Daily Site Data's new column is available for later approved intake.

## Review scope

The Python tests check structural preservation, existing value preservation, range boundaries, rejected drift, stale approval, backup and recovery, and Workfront ID behavior. They do not evaluate Excel formulas. Run the PowerShell check on the target laptop before signing off on P1, and measure Excel save and recalculation time there.

The migration operates on generated synthetic fixtures through openpyxl. The actual project workbook's Excel save and recalculation path, authorized staff roster, reporting cutoff, and operational workbook rollout require owner confirmation. WhatsApp intake, record-level duplicate review, operational clearance approval, dashboard, and report generation are later priorities.

## P2 desktop review

Use the `codex/mvp-p2-1` branch for desktop intake and staged Excel commits. Complete the local
baseline setup above once from the repository root. Then use a fresh
demo directory inside `local`:

```powershell
& $python -m kvd.demo .\local\p2-demo
& $python -m kvd.desktop --workbook .\local\p2-demo\synthetic-five.xlsx --journal .\local\p2-demo\audit.sqlite3
```

Paste `examples/five-synthetic-reports.txt` and enter a synthetic operator
name. Each block starts with `Workfront:`. Enter one `Field: value` per line
and use `Command:` for mixed command batches. Use semicolons between attendee
IDs or resource entries. Resources take `type | description | count | hours per unit`.
The demo seeds five aliases and an open junior task.

Select each draft to correct fields and review its source. Build an Excel
preview to see cell-level changes. Likely matches need a reasoned decision:
retain a distinct entry, suppress a replay, or correct a matched site record.
For quantity verification, enter the separate value and evidence in draft
review. A distinct named reviewer must attest at approval. Reviewer names
are local audit labels for synthetic acceptance.

Approval records the preview digest, workbook fingerprint, source excerpts,
edited data, duplicate decisions, optional quantity verification evidence,
and exact typed cell proposals in SQLite. Clearance and release fields stay
outside the write scope. Reported blockers can be recorded as observations.
The original and synthetic source workbooks remain read-only.

Free-form WhatsApp prose, attachments, authenticated reviewer roles,
and production workbook editing are outside this P2.1 workflow. No live data should be entered. Run the full acceptance suite with
`python -m unittest discover -s tests -v`.

### Commit and recover a synthetic workbook revision

1. Paste English or Bahasa Melayu labels. For example, use `Workfront` or
   `Zon kerja`, `Date` or `Tarikh`, and `Reported qty` or
   `Kuantiti dilaporkan`. You can mix labels in the same paste. See
   `examples/synthetic-report-ms.txt`. Values are preserved as entered and
   missing required fields must be supplied during review.
2. Review the proposal and record approval. Choose **Save approved revision**
   and enter a new `.xlsx` path on a local NTFS volume. The application
   verifies the reviewed source, records the staged revision and its
   before-and-after values in SQLite, and publishes it to the new path.
   Reopen the successor for the next batch. Generate a fresh preview for
   any older proposal that targets a protected field or a superseded source.
3. After an interruption, use **Recover** with the operator field filled in.
   A matching output and backup complete the journal entry. A missing
   output with a verified backup allows a retry. A mismatch stays pending
   for manual comparison.

Approved tokens can also be used after restarting the desktop. The status
shows the leading characters of the token. The complete token is recorded
in the SQLite preview and approval events.

```powershell
$token = Read-Host "Full approval token"
& $python -m kvd.excel_commit --journal .\local\p2-demo\audit.sqlite3 commit $token .\local\p2-demo\revision-2.xlsx --actor "Synthetic Operator"
& $python -m kvd.excel_commit --journal .\local\p2-demo\audit.sqlite3 recover --actor "Synthetic Operator"
& $python -m unittest discover -s tests -v
.\tests\excel-commit-acceptance.ps1 -Python $python -DemoDirectory .\local\windows-commit-test
```

Use a fresh Windows acceptance-test directory. The Python tests verify formula
definitions and independently compute expected progress and BQ totals from
verified quantities. The PowerShell acceptance test uses desktop Excel to
recalculate the committed fixture in memory and compare its results. Because
openpyxl does not evaluate formulas, desktop Excel recalculation is required
to populate displayed totals. Approval and audit use local synthetic labels,
and the commit workflow supports local synthetic workbooks only.
