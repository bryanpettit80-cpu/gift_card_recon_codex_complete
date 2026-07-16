# Gift Card Reconciliation

This program automatically reconciles weekly gift card activity to verified Micros POS and tender evidence for Richmond (`9354`) and Virginia Beach (`9355`), then carries the completed activity into the monthly-close workflow.

## Source, Development, and Operations

The project uses three deliberately separate locations:

- [GitHub](https://github.com/bryanpettit80-cpu/gift_card_recon_codex_complete) is the authoritative source and history.
- A local clone is the sole Git working copy used for development, testing, commits, and pushes. A typical location is created from `$env:USERPROFILE`, such as `Documents\Repos\gift_card_recon_codex_complete`.
- Dropbox holds live inputs, evidence, reports, archives, and a deployed program snapshot. The Dropbox snapshot must not contain `.git`, virtual environments, caches, tests, or other development state.

This separation keeps Git's frequently changing internal files out of Dropbox while preserving the double-click operator workflow.

## Run It

1. Save one Monday-Sunday Gift Card Activity workbook in the correct store's `activity` inbox.
2. Close the workbook in Excel.
3. Click `Run Weekly Gift Card Reconciliation.cmd`.

The program retrieves that store's matching POS data, validates all seven business dates and tender evidence, creates the workbook, stages the Activity report for monthly close, and retains a hash-verified weekly evidence package. No POS control CSV entry is required.

Stores are processed independently. An empty inbox is a normal no-op. A malformed or incomplete report remains in its inbox, does not publish a workbook, and causes the launcher to identify the review item while still allowing the other store to run.

## Folders

```text
00 START HERE - Gift Card Reconciliation.txt
Run Weekly Gift Card Reconciliation.cmd
Run Monthly Gift Card Close.cmd
Check Gift Card Reconciliation Health.cmd
01 Weekly Gift Card Activity Reports/
  9354 Richmond/activity/
  9355 Virginia Beach/activity/
02 Monthly Close Inputs/
  9354 Richmond/
  9355 Virginia Beach/
  Darden Reports - Drop Here/
03 Finished Reports/
  Weekly/
  Monthly Close/
  Monthly Close - Review Required/
04 Archive/
  Weekly Reconciliation/  # weekly source, POS evidence, workbook, and manifest
  Monthly Close/          # canonical close evidence and manifests
  Generated Reports/      # preserved historical workbooks
  Legacy Reconciliation/  # pre-current-process Darden material
  Cleanup Manifests/      # hash-verified organization records
_automation_runs/
Gift Card Reconciliation Automation/  # deployed program snapshot; no .git
  deployment-manifest.json
```

Operators normally use only the two reconciliation launchers, weekly Activity inboxes, Darden inbox, and finished reports. `Gift Card Reconciliation Automation` contains the deployed runtime files; `_automation_runs` contains logs, QA output, and review quarantine.

`Check Gift Card Reconciliation Health.cmd` validates the deployed revision and hashes, runtime readiness, Excel, both Micros sources, Dropbox file accessibility, output-folder writability, and pending inputs without performing a reconciliation.

The Python environment, package cache, compiled Python cache, and temporary extraction files are kept outside Dropbox under `%LOCALAPPDATA%\GiftCardRecon`. They are not part of the repository or monthly-close evidence.

## Automatic Weekly POS Controls

The normal weekly runner reads `DLYSYSTT.TXT` and `TENDER_DETAIL.TXT` from each store's configured external Micros export. It requires one correct-store Monday-Sunday Activity report, exact weekly POS coverage, valid money fields, and matching tender evidence. A scheduled Monday may be absent only when Activity, POS, and tender evidence for that day are all zero. Missing values are never replaced with Activity totals.

A completed store/week is retained under `04 Archive\Weekly Reconciliation` with the original Activity report, a compact seven-day POS/tender CSV, an identical archived copy of the finished workbook, and `weekly_manifest.json` containing sizes and SHA-256 hashes. Exact reruns are idempotent; conflicting duplicate weeks are sent to `_automation_runs\review\duplicate-inputs`.

Completed weekly workbooks remain under `03 Finished Reports\Weekly`, including workbooks with a `REVIEW` status. The separate `Monthly Close - Review Required` folder is only for blocked monthly-close diagnostics.

## Workbook

The first tab is `Reconciliation`. It now includes:

- The POS tie-out rows
- A `Gift Card Activity File Totals` section with the activity file name, report period, row count, activations, redemptions, and net activity
- POS control totals

Other tabs keep the weekly, daily, raw detail, source file, and exception details.

## Setup

If this is a fresh deployment, double-click either operator runner. The first run creates `%LOCALAPPDATA%\GiftCardRecon\operator\venv` and installs the required packages. Later runs reuse that environment; it refreshes when the dependency specification or deployed application payload changes. `_program\.venv` is no longer used.

Operator temporary files and caches live under `%LOCALAPPDATA%\GiftCardRecon\operator`. Development tests use an independent `%LOCALAPPDATA%\GiftCardRecon\development` runtime, cache, and temporary workspace so test dependencies cannot alter the operator environment.

## Test

From the local development checkout, run:

```powershell
.\_program\run_tests.ps1
```

The test runner creates or refreshes the separate development runtime and keeps its cache outside Dropbox.

GitHub CI runs the unit suite on Linux with Python 3.10, 3.11, and 3.12, plus a Windows smoke job on Python 3.14. The Windows job parses every PowerShell script, validates the operator-launcher contracts, runs the isolated migration and deployment/health fixtures, and runs the unit suite with Excel PDF automation mocked. The real Excel COM/PDF integration test remains a local Windows check on a workstation with Microsoft Excel installed.

## Monthly Close From Micros

The Darden credit memo is the final checkbox for the month. Put every new Darden PDF in the shared inbox:

```text
02 Monthly Close Inputs\Darden Reports - Drop Here\
```

Then double-click `Run Monthly Gift Card Close.cmd`. A no-argument run scans every PDF in the inbox, reads the store and fiscal service period from the report, and processes Richmond and Virginia Beach independently. There is no background watcher.

The source folders remain store- and period-specific:

```text
02 Monthly Close Inputs\9355 Virginia Beach\FY27 M01 - Fiscal June\
  summary\
    07.05.2026 9355 Gift Card Summary.xlsx
  activity\
    five Monday-Sunday Gift Card Activity reports
```

Weekly activity files are staged automatically by the weekly runner. Completed evidence is retained under `04 Archive\Monthly Close` with a SHA-256 close manifest.

### Close dispositions

- `CLOSED`: every required control passes.
- `CLOSED WITH REVIEW`: evidence is complete, Darden matches, and every nonzero weekly and period POS/tender variance is no more than `$5.00`.
- `REVIEW REQUIRED`: identity, completeness, coverage, Darden, archive/publication, or a larger variance fails.

The Darden result is shown separately as `MATCHED` or `MISMATCHED`; it is not the overall close status. Summary-to-activity and Darden-to-Summary controls must match to the cent.

The runner requires exactly one correct-store activity report for every expected week, exact fiscal-date coverage, and both Micros evidence files. A scheduled Monday may be absent from Micros only when activity and tender evidence are also zero. An existing Monday POS row is included normally. Missing POS values are never replaced with activity totals.

### Deliverables

Successful close reports are written as matching workbook and PDF files:

```text
03 Finished Reports\Monthly Close\<fiscal period>\
  Richmond_9354_<period>_Monthly_Close.xlsx
  Richmond_9354_<period>_Monthly_Close.pdf
  Virginia_Beach_9355_<period>_Monthly_Close.xlsx
  Virginia_Beach_9355_<period>_Monthly_Close.pdf
```

The first worksheet is an intentional two-page, letter-landscape executive accounting report:

- Page 1 shows the location and fiscal period, an overall status band, `Settlement Tie-Out` cards, a `Close Controls` table, and `Open Items Summary`.
- Page 2 repeats the location and period, then shows `Weekly Variance Detail`, `Variance Summary`, deduplicated `Review Items`, and `Evidence and Audit Trail`.

The report uses Arial, a navy/light-blue accounting palette, and green/amber/red only for assessed status. Settlement amounts use neutral accounting formatting, including negative values. Both formats carry report metadata plus generated-time, location, and page details in the footer. The fixed 85% print scale, merged-cell borders, and compact follow-up text keep the report readable at exactly two pages.

A successful close is published and archived only after both the workbook and Excel-exported PDF are verified as a matching pair. If PDF export fails, no canonical close report is published and no source evidence is archived or removed.

A blocked run always attempts a red diagnostic workbook and an Excel-exported diagnostic PDF under `03 Finished Reports\Monthly Close - Review Required`. If Excel cannot create the PDF, the new workbook may be published by itself; the command reports the original close blocker, the exact PDF export error, the authoritative workbook path, and that no diagnostic PDF was published for that run. Any older same-named diagnostic PDF is retired transactionally first, so an old PDF cannot appear to match the new workbook. If a locked old diagnostic prevents that retirement, the old pair is preserved and the new diagnostic is not published; the command reports both the original blocker and the diagnostic-publication failure. A locked canonical file likewise fails clearly and never causes an alternate filename.

### Archive-backed reissues

Use the dedicated archive mode to reproduce a completed store-period from its retained evidence instead of manually assembling source paths:

```powershell
& ".\Run Monthly Gift Card Close.cmd" `
  -Store 9355 `
  -Period FY27-M01 `
  -ReissueFromArchive
```

`-ReissueFromArchive` requires both `-Store` and `-Period`. It cannot be combined with `-InputDir`, `-DardenPath`, or `-MicrosPath`; the program derives the canonical archived Summary, activity, Darden, and Micros paths from the existing close manifest. It verifies every retained source against that manifest, rejects missing, changed, or out-of-archive evidence, and permits the archived Micros snapshot only after containment validation.

Archive mode forces the equivalents of `-NoStageWeekly` and `-NoCleanup`. It does not stage weekly files, delete sources, or touch live input folders. Reissued canonical artifacts still follow the normal transactional workbook/PDF publication rules.

### Explicit reruns

Store, period, Darden, input, output, and Micros options remain available for controlled reruns:

```powershell
& ".\Run Monthly Gift Card Close.cmd" `
  -Store 9355 `
  -Period FY27-M01 `
  -DardenPath "C:\path\to\Darden credit memo.pdf" `
  -InputDir ".\04 Archive\Monthly Close\9355\FY27 M01 - Fiscal June" `
  -NoStageWeekly `
  -NoCleanup
```

`-Period 2026-06` maps to Darden Fiscal June 2026 (`FY27-M01`), covering `2026-06-01` through `2026-07-05`.

All new evidence is written beneath `04 Archive\Monthly Close`. A historical lowercase `monthly-close` path remains readable when supplied explicitly with `-InputDir`, but it is never used as a silent fallback.

Default Micros sources are location-specific:

- Richmond / `9354`: the external Dropbox folder `micros_data\RC-Richmond-current`
- Virginia Beach / `9355`: the external Dropbox folder `GETLinkedData-VB`

`-MicrosPath` may point to the configured location source or to a store-identified archived snapshot (an extracted folder, `.zip`, or `.7z`). Arbitrary folders and the other location's source are rejected.

## Deploying the Program and Operator Assets

Both PowerShell entrypoints accept `-OperationsRoot`. The Dropbox-facing launchers pass their own folder explicitly, so inputs, reports, archives, logs, and review files remain outside the deployed program directory. Relative override paths are resolved from the operations root, not from the code checkout. Store Micros exports remain external siblings of the operations folder.

Develop and commit only from the local clone. After updating and testing that clone, deploy a SHA-256-verified program snapshot to Dropbox with:

```powershell
$repoRoot = Join-Path $env:USERPROFILE "Documents\Repos\gift_card_recon_codex_complete"
$operationsRoot = Join-Path $env:USERPROFILE "Dropbox\Gift Card Reconciliation"

& "$repoRoot\_program\maintenance\deploy_operator_program.ps1" `
  -OperationsRoot $operationsRoot
```

The deployment excludes Git metadata and development-only files, verifies copied files, writes `Gift Card Reconciliation Automation\deployment-manifest.json`, and refreshes the operator-facing launchers and notes. For an assets-only refresh, use `_program\install_operator_assets.ps1 -OperationsRoot $operationsRoot` from the local clone.

After deployment, double-click `Check Gift Card Reconciliation Health.cmd` or run it from PowerShell:

```powershell
& "$operationsRoot\Check Gift Card Reconciliation Health.cmd"
```
