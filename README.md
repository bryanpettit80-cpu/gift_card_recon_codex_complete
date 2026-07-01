# Gift Card Reconciliation

This program reconciles weekly gift card activity files to POS control totals for stores `9354` and `9355`.

## Run It

1. Put the weekly Gift Card Activity file in the store's `activity` folder.
2. Enter POS totals in the store's `pos_controls.csv`.
3. Click `Run-Gift-Card-Reconciliation.cmd`.

The finished workbook is created in `output`.
After a workbook is created, the two POS total cells are cleared so the same file is ready for the next week.

## Folders

```text
input/
  9354/
    weekly/
      activity/
      summary/
      pos_controls.csv
  9355/
    weekly/
      activity/
      summary/
      pos_controls.csv
output/
```

Use `summary` only if you have an optional weekly Gift Card Summary file.

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

If this is a fresh download, run `install.ps1` once. After that, use `Run-Gift-Card-Reconciliation.cmd`.

## Test

For verification, run:

```powershell
.\run_tests.ps1
```

## Monthly Close From Micros

For June store `9355`, put the monthly Gift Card Summary and weekly Gift Card Activity files in `input\9355\2026-06`, then run:

```powershell
.\run_monthly_close.ps1 `
  -Store 9355 `
  -Period 2026-06 `
  -MicrosPath .\_inspect_micros3700
```

`-MicrosPath` can point to an extracted Micros export folder or a `Micros3700.7z` archive. This monthly-close runner derives POS Gift Card Issue and POS Gift Card Payment from the Micros export, creates the standard reconciliation workbook, and appends `Weekly POS Variance Detail` on the existing `Reconciliation` tab.
