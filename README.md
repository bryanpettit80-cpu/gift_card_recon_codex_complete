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

Monthly close uses the checked-in Darden fiscal calendar derived from `Darden Fiscal Calendar as of 11_2025.pdf` in Google Drive. For Fiscal June FY27, put the monthly Gift Card Summary in the store's summary folder, for example:

```text
Monthly Close\9355\FY27 M01 - Fiscal June\summary\
```

The weekly activity files should already be in:

```text
Monthly Close\9355\FY27 M01 - Fiscal June\activity\
```

Then run:

```powershell
.\Run-Monthly-Close.cmd `
  -Store 9355 `
  -Period FY27-M01
```

You can also double-click `Run-Monthly-Close.cmd` to run the same monthly close defaults. The command accepts options when you need to pass a different store, fiscal period, archive folder, or Micros path.

`-Period 2026-06` is also accepted and maps to Darden Fiscal June 2026 (`FY27-M01`). That fiscal period runs from `2026-06-01` through `2026-07-05`.

By default, monthly close uses the store-specific Micros export folder:

- `9355`: `..\GETLinkedData-VB`
- `9354`: `..\micros_data\RC-Richmond-current`

`-MicrosPath` can point to another extracted Micros export folder or to a `Micros3700.7z` archive when 7-Zip is available. Passing the extracted folder is the most reliable option. This monthly-close runner derives POS Gift Card Issue and POS Gift Card Payment from the Micros export, creates the standard reconciliation workbook, appends `Weekly POS Variance Detail` on the existing `Reconciliation` tab, then moves the monthly source files to `Archive - Old Files\monthly-close`.

To run Richmond / store `9354`:

```powershell
.\Run-Monthly-Close.cmd `
  -Store 9354 `
  -Period FY27-M01
```

Richmond's current Micros export is published by the server-side helper scripts in `_program\support\richmond_micros_export`. The installed server task on `RESSERVER` is `Gift Card Export Copy to Dropbox`, scheduled for `06:35` daily, and writes the current files into `C:\Users\customer\Dropbox\micros_data\RC-Richmond-current`.

To check month-end readiness without creating the workbook, run:

```powershell
.\Run-Monthly-Close.cmd `
  -Store 9355 `
  -Period FY27-M01 `
  -PrepareOnly
```

The prepare step creates `Monthly Close\<store>\<fiscal period>\summary` and `Monthly Close\<store>\<fiscal period>\activity`, checks for the expected week-ending activity files, verifies the Gift Card Summary, and confirms the Micros export reaches the Darden fiscal period end.

To rerun a completed monthly close from archived source files, point `-InputDir` at the archived fiscal period folder and turn off weekly staging and cleanup:

```powershell
.\Run-Monthly-Close.cmd `
  -Store 9355 `
  -Period FY27-M01 `
  -InputDir ".\Archive - Old Files\monthly-close\9355\FY27 M01 - Fiscal June" `
  -NoStageWeekly `
  -NoCleanup
```
