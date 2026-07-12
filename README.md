# Gift Card Reconciliation

This program reconciles weekly gift card activity files to POS control totals for stores `9354` and `9355`.

## Run It

1. Download or copy one weekly Gift Card Activity file into the store's `activity` folder.
2. Enter POS totals in the store's `pos_controls.csv`.
3. Click `Run-Gift-Card-Reconciliation.cmd`.

The finished workbook is created in `Output`.
After a workbook is created, the two POS total cells are cleared so the same file is ready for the next week.
The weekly activity file is moved into that store's Darden fiscal monthly close folder so month-end is easier to prepare.

If a weekly folder has more than one Gift Card Activity file, that store is skipped and the program lists the files to fix. Remove or move the extra file, then run again.

## Folders

```text
9354 - Weekly/
  activity/
  pos_controls.csv
9355 - Weekly/
  activity/
  pos_controls.csv
Monthly Close/
  9354/
  9355/
  Darden Reports - Drop Here/
Output/
Archive - Old Files/
  Monthly Close/          # canonical close evidence and manifests
  Generated Reports/      # preserved historical workbooks
  Legacy Reconciliation/  # pre-current-process Darden material
  Cleanup Manifests/      # hash-verified organization records
_program/
```

`_program` contains the code and tests. Operators normally only use the weekly folders, `Monthly Close`, and `Output`.

The Python environment, package cache, compiled Python cache, and temporary extraction files are kept outside Dropbox under `%LOCALAPPDATA%\GiftCardRecon`. They are not part of the repository or monthly-close evidence.

## POS Controls

Each `pos_controls.csv` has one line to fill in:

```csv
store,period,pos_gift_card_issue,pos_gift_card_payment
9354,auto,,
```

Leave `period` as `auto`. The program reads the week-ending date from the Gift Card Activity file and names the workbook with the correct week.
After a successful run, the program clears only `pos_gift_card_issue` and `pos_gift_card_payment`. If the workbook is not created, the entered totals stay in place so they can be corrected and reused.

## Workbook

The first tab is `Reconciliation`. It now includes:

- The POS tie-out rows
- A `Gift Card Activity File Totals` section with the activity file name, report period, row count, activations, redemptions, and net activity
- POS control totals

Other tabs keep the weekly, daily, raw detail, source file, and exception details.

## Setup

If this is a fresh download, double-click either runner. The first run creates `%LOCALAPPDATA%\GiftCardRecon\venv` and installs the required packages. Later runs reuse that environment without reinstalling; setup runs again only when `requirements.txt` or `pyproject.toml` changes. `_program\.venv` is no longer used.

Temporary Micros extraction uses `%LOCALAPPDATA%\GiftCardRecon\temp\micros-extract`. Python bytecode, pytest state, and package downloads use `%LOCALAPPDATA%\GiftCardRecon\cache`.

## Test

For verification, run:

```powershell
.\_program\run_tests.ps1
```

The test runner uses the same local runtime and keeps its cache outside Dropbox.

## Monthly Close From Micros

The Darden credit memo is the final checkbox for the month. Put every new Darden PDF in the shared inbox:

```text
Monthly Close\Darden Reports - Drop Here\
```

Then double-click `Run-Monthly-Close.cmd`. A no-argument run scans every PDF in the inbox, reads the store and fiscal service period from the report, and processes Richmond and Virginia Beach independently. There is no background watcher.

The source folders remain store- and period-specific:

```text
Monthly Close\9355\FY27 M01 - Fiscal June\
  summary\
    07.05.2026 9355 Gift Card Summary.xlsx
  activity\
    five Monday-Sunday Gift Card Activity reports
```

Weekly activity files are staged from the weekly archive when available. Completed evidence is retained under `Archive - Old Files\Monthly Close` with a SHA-256 close manifest.

### Close dispositions

- `CLOSED`: every required control passes.
- `CLOSED WITH REVIEW`: evidence is complete, Darden matches, and every nonzero weekly and period POS/tender variance is no more than `$5.00`.
- `REVIEW REQUIRED`: identity, completeness, coverage, Darden, archive/publication, or a larger variance fails.

The Darden result is shown separately as `MATCHED` or `MISMATCHED`; it is not the overall close status. Summary-to-activity and Darden-to-Summary controls must match to the cent.

The runner requires exactly one correct-store activity report for every expected week, exact fiscal-date coverage, and both Micros evidence files. A scheduled Monday may be absent from Micros only when activity and tender evidence are also zero. An existing Monday POS row is included normally. Missing POS values are never replaced with activity totals.

### Deliverables

Successful close reports are written as matching workbook and PDF files:

```text
Output\Monthly Close\<fiscal period>\
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

A blocked run always attempts a red diagnostic workbook and an Excel-exported diagnostic PDF under `Output\Review Required`. If Excel cannot create the PDF, the new workbook may be published by itself; the command reports the original close blocker, the exact PDF export error, the authoritative workbook path, and that no diagnostic PDF was published for that run. Any older same-named diagnostic PDF is retired transactionally first, so an old PDF cannot appear to match the new workbook. If a locked old diagnostic prevents that retirement, the old pair is preserved and the new diagnostic is not published; the command reports both the original blocker and the diagnostic-publication failure. A locked canonical file likewise fails clearly and never causes an alternate filename.

### Archive-backed reissues

Use the dedicated archive mode to reproduce a completed store-period from its retained evidence instead of manually assembling source paths:

```powershell
.\Run-Monthly-Close.cmd `
  -Store 9355 `
  -Period FY27-M01 `
  -ReissueFromArchive
```

`-ReissueFromArchive` requires both `-Store` and `-Period`. It cannot be combined with `-InputDir`, `-DardenPath`, or `-MicrosPath`; the program derives the canonical archived Summary, activity, Darden, and Micros paths from the existing close manifest. It verifies every retained source against that manifest, rejects missing, changed, or out-of-archive evidence, and permits the archived Micros snapshot only after containment validation.

Archive mode forces the equivalents of `-NoStageWeekly` and `-NoCleanup`. It does not stage weekly files, delete sources, or touch live input folders. Reissued canonical artifacts still follow the normal transactional workbook/PDF publication rules.

### Explicit reruns

Store, period, Darden, input, output, and Micros options remain available for controlled reruns:

```powershell
.\Run-Monthly-Close.cmd `
  -Store 9355 `
  -Period FY27-M01 `
  -DardenPath "C:\path\to\Darden credit memo.pdf" `
  -InputDir ".\Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June" `
  -NoStageWeekly `
  -NoCleanup
```

`-Period 2026-06` maps to Darden Fiscal June 2026 (`FY27-M01`), covering `2026-06-01` through `2026-07-05`.

All new evidence is written beneath `Archive - Old Files\Monthly Close`. A historical lowercase `monthly-close` path remains readable when supplied explicitly with `-InputDir`, but it is never used as a silent fallback.

Default Micros sources are location-specific:

- Richmond / `9354`: `..\micros_data\RC-Richmond-current`
- Virginia Beach / `9355`: `..\GETLinkedData-VB`

`-MicrosPath` may point to the configured location source or to a store-identified archived snapshot (an extracted folder, `.zip`, or `.7z`). Arbitrary folders and the other location's source are rejected.
