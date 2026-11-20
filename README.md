# Construction AI Management System

P0/P1 is a local Windows Python CLI for synthetic workbook acceptance. It uses no paid API or cloud runtime. The supplied Master Register V2 was read only to extract structural metadata. Generate the schema manifest locally from your reference workbook. The manifest, workbook copies, and journal are excluded from Git.

## Windows installation

Install Python 3.11 or later. Use a local NTFS folder for the workbook copies, audit database, and backups. Use an existing desktop Excel installation for the recalculation check.

Clone the review branch and run these commands in PowerShell from the repository root.

```powershell
git clone --branch codex/mvp-p0-p1 https://github.com/smartkids-adventure/construction-ai-agent.git
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
