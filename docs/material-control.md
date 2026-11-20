# Material control instructions

Run these commands from the repository root after the README installation.
Use a dedicated local synthetic journal and an empty demo folder. The demo
registers four synthetic materials, imports one lookahead row from a committed
workbook, enters three additional programme lines, and issues one example
requisition through separate synthetic review and approval decisions.

```powershell
$journal = ".\local\p4-demo\audit.sqlite3"
& $python -m kvd.materials --journal $journal forecast 2026-01-15
& $python -m kvd.materials --journal $journal forecast 2026-01-15 --material SYN-REBAR --status RED
& $python -m kvd.materials --journal $journal takeoffs
& $python -m kvd.materials --journal $journal requisitions
& $python -m kvd.materials --journal $journal export 1 .\local
```

## Enter and review records

Write one operation or an array of operations to a UTF-8 JSON file. Use
`apply events.json --actor "Synthetic Operator"` in the same CLI. A batch is
one SQLite transaction. If any operation fails, its changes and audit events
roll back. Use separate batches for the different people who record or approve
decisions. Identifiers are explicit, and quantities accept up to six decimal
places. Dates are `YYYY-MM-DD`, and units must agree with the material register.

| Operation | Required JSON fields in addition to `operation` |
|---|---|
| `material` | `id`, `description`, `specification`, `unit`. Add `lead` in calendar days, `safety` in material units, and `variance` as a percentage when known. |
| `material-policy` | `id`, `lead`, `safety`, `variance`, `reference`. Use JSON `null` for a policy value pending confirmation. |
| `drawing` | `number`, `revision`, `status` (`APPROVED`, `PENDING`, `REJECTED`, or `SUPERSEDED`), and `reference`. An optional `snapshot` is a committed workbook SHA-256 in the same journal. The operation returns the drawing ID. |
| `drawing-status` | Drawing `id`, new `status`, and approval evidence `reference`. Record drawing revisions in publication order. A new recorded revision suspends reliance on the prior revision. |
| `takeoff` | `id`, drawing ID `drawing`, `workfront`, `material_id`, `element`, `dimensions`, `formula`, `unit`, `source`, and explicit `allowance_pct`. Dimensions use objects such as `"length": {"value": "3.2", "unit": "m"}`. The output retains separate design, allowance and procurement quantities. |
| `takeoff-review` | Take-off `id`, `action` of `verify` or `approve`, and evidence `reference`. Use separate named people for preparation, verification, and approval. |
| `programme` | `id`, `takeoff`, `activity`, `location`, `start`, `finish`, `workdays`, `rate`, `unit`, `production_unit`, and `source`. Give `daily` production or interval `target`. Optional `exclusions` hold non-working dates. |
| `programme-import` | Committed `workbook` path, physical lookahead `row`, `id`, `takeoff`, `location`, `workdays`, `rate`, `unit`, and `production_unit`. Reads activity, dates, target and production unit from the row. The take-off supplies the Workfront ID, and the activity must match the workfront index. |
| `programme-verify` / `programme-retire` | Programme `id` and evidence `reference`. Verify through a different named reviewer. Retire the old programme before recording a replacement. |
| `movement` | `id`, `material_id`, `unit`, `location`, `date`, `kind`, `quantity` and `source`. Add `workfront` for allocation, issue and consumption. |
| `order` | `id`, `material_id`, `unit`, `location`, `po`, `quantity`, `placed`, `due`, and `reference`. Optional `workfront`, issued `requisition` ID and both `eta` and `confirmed_on` for a confirmed promise. |
| `order-confirm` | Order `id`, `eta`, `confirmed_on`, and `reference`. Reconfirm a missed ETA before its outstanding balance can re-enter the forecast. |
| `arrival` | `id`, `order`, `unit`, `date`, delivery-note `reference`, and explicit `accepted`, `rejected`, and `quarantined` quantities. Repeat for partial deliveries, using a new delivery note. |
| `prepare-requisitions` | Forecast `start`, optionally `workfront` and `material_id`. Creates drafts for verified RED shortages and skips scopes with an open requisition. |
| `revise-requisition` | Current unissued requisition `id`, forecast `start`, and `reference`. Recalculates the scope and supersedes the old revision. |
| `requisition-action` | Requisition `id`, `action` (`review`, `approve`, `issue`, or `close`), and evidence `reference`. Preparation, review, and approval require different names. Issue creates Word and PDF locally. Close an issued scope explicitly when settled. |

Example batch to calculate the forecast and prepare shortage requisitions:

```json
{"operation": "prepare-requisitions", "start": "2026-01-15", "material_id": "SYN-REBAR"}
```

Use `opening`, `receipt`, `issue`, `adjustment`, `reserve`, `release`,
`quarantine`, `quarantine-release`, `quarantine-reject`, and `consume`
movement kinds. Start with an explicit counted opening balance at each
material and location. Post effective dates in order. A positive adjustment
adds available stock and a negative adjustment removes stock that is neither
reserved nor quarantined. Issue consumes that workfront’s reservation first,
then unallocated stock. Actual consumption is recorded separately against
units already issued to that workfront and location.

## Forecast interpretation

- The 14-day window includes the start date. A programme can define weekdays
  (Monday = 0 through Sunday = 6) and exclusions. A target is spread evenly
  over eligible dates, rounded up at six-decimal precision, then multiplied by
  the documented consumption rate in material units per production unit.
- Available stock excludes quarantine and existing reservations. The forecast
  uses a workfront’s own reservation first, then shares unallocated stock in
  Workfront ID order. A confirmed delivery earmarked for one workfront remains
  reserved for that workfront. All workfronts compete in the calculation
  before dashboard filters are applied. Shared location balances shown on
  multiple rows must be counted once.
- Confirmed future receipts are applied before demand on their ETA. Actual
  accepted receipts appear in current stock. A delivery counts as arrived
  when accepted, rejected, or quarantined. Only accepted units enter usable
  stock. Rejected units need a separate replacement promise if still required.
- RED is the first workfront deficit. The shortage shown is maximum
  unfilled demand during the window. The recommended requisition date
  subtracts the supplier’s lead time from the first shortage date.
  AMBER covers a safety-stock breach or an overdue delivery without a
  projected deficit. GREEN requires a complete verified forecast with no
  warning. UNKNOWN identifies missing or unverified inputs. Only a verified
  RED shortage generates a draft. Buffer purchases require a separate
  procurement decision.
- Variance compares explicitly recorded consumption with planned consumption.
  Design vs issued compares current approved design quantity with historical
  issues at the selected location. Excesses above the configured percentage
  flag potential waste for review.

## Known limitations

P4 uses local named audit actors and synthetic workbooks. It does not
authenticate corporate identities or place orders. Registers are entered
through the CLI, and the dashboard provides read-only review. Drawings
require manually transcribed dimensions and references. No CAD/PDF extraction
or implicit unit conversion is performed. Drawing and programme revisions
reflect current records rather than historical reconstruction. Forecasting
uses a deterministic Workfront ID order for shared stock and does not
optimize alternative suppliers, packaging or stock transfers. Quarantine
movements are tracked by material and location, with source references
providing the link to a recorded delivery note.

## Windows acceptance checklist

- Install Python 3.11+ with Tcl/Tk on the target laptop, install the branch
  requirements, generate the local baseline, and run all unit tests.
- Use a fresh demo folder on a local NTFS volume. Launch the P3 dashboard
  against its committed workbook and confirm the Material control tab opens.
- Set the start date to `2026-01-15`. Confirm four warning levels. Filter to
  `SYN-REBAR` and verify 200 kg demand, 25 kg confirmed incoming, 129 kg
  shortage, 19 Jan stockout and 14 Jan recommended requisition.
- Inspect Deliveries. Verify the first order is overdue with 8 kg accepted,
  2 kg rejected, 5 kg quarantined, and 15 kg outstanding. Inspect the
  consumption variance and urgent procurement views.
- Open the issued sample in Word and a PDF viewer. Compare quantities,
  workfront, dates, approval names, and supporting references.
- Verify the source workbook hash is unchanged and P3 report generation works
  against the same journal. Run the separate P2.1 desktop Excel acceptance
  before signing off on operational workbook handling.
