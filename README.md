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
Output/
Archive - Old Files/
_program/
```

`_program` contains the code and tests. Operators normally only use the weekly folders, `Monthly Close`, and `Output`.

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

If this is a fresh download, double-click `Run-Gift-Card-Reconciliation.cmd`. It will install what it needs the first time, then run the weekly reconciliation.

## Test

For verification, run:

```powershell
.\_program\run_tests.ps1
```

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

The first worksheet is an intentional two-page report. Page 1 shows the location, period, overall disposition, Darden result, executive controls, and open actions. Page 2 repeats the location and period, then shows weekly POS/tender variances, coverage, status, unified exceptions, evidence notes, and page numbering.

Blocked runs create red diagnostic files under `Output\Review Required`; they never publish a canonical close report or remove live evidence. A locked canonical file fails clearly and does not create an alternate filename.

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

Default Micros sources are location-specific:

- Richmond / `9354`: `..\micros_data\RC-Richmond-current`
- Virginia Beach / `9355`: `..\GETLinkedData-VB`

`-MicrosPath` may point to the configured location source or to a store-identified archived snapshot (an extracted folder, `.zip`, or `.7z`). Arbitrary folders and the other location's source are rejected.
