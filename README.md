# Gift Card Reconciliation Automation

This repo reconciles restaurant gift card activity against POS control totals. It supports monthly reconciliations with a required gift card summary and weekly reconciliations where the summary is optional.

## What it does

Monthly input:

- One monthly `Gift Card Summary.xlsx`
- One or more weekly `Gift Card Activity.xls` or `.xlsx` files
- POS controls either through `pos_controls.csv` or command-line arguments

Weekly input:

- One or more `Gift Card Activity.xls` or `.xlsx` files
- POS controls either through `pos_controls.csv` or command-line arguments
- Optional weekly `Gift Card Summary.xlsx`

Output:

- A formatted Excel reconciliation workbook in `output/`

Workbook tabs:

1. `Reconciliation`
2. `Weekly Activity Detail`
3. `Daily Activity Detail`
4. `Raw Detail`
5. `Source Files`
6. `Exception Log`

## Folder layout

Simple weekly workflow:

```text
input/
  9354/
    weekly/
      activity/
        drop the current 9354 Gift Card Activity .xls/.xlsx file here
      summary/
        optional weekly Gift Card Summary.xlsx
      pos_controls.csv
  9355/
    weekly/
      activity/
        drop the current 9355 Gift Card Activity .xls/.xlsx file here
      summary/
        optional weekly Gift Card Summary.xlsx
      pos_controls.csv
output/
```

For weekly runs, the program reads the activity report date range and names the workbook for the actual ISO week. You do not need to type `2026-W23` or the week-ending date.

Monthly example:

```text
input/
  9354/
    2026-05/
      summary/
        05.31.2026 9354 Gift Card Summary.xlsx
      activity/
        05.03.2026 9354 Gift Card Activity.xls
        05.10.2026 9354 Gift Card Activity.xls
        05.17.2026 9354 Gift Card Activity.xls
        05.24.2026 9354 Gift Card Activity.xls
        05.31.2026 9354 Gift Card Activity.xls
      pos_controls.csv
output/
```

Older weekly period-folder example:

```text
input/
  9354/
    2026-W22/
      summary/
        optional weekly Gift Card Summary.xlsx
      activity/
        05.31.2026 9354 Gift Card Activity.xls
      pos_controls.csv
output/
```

## One-time setup

Open PowerShell 7 in this repo folder:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\install.ps1
```

## Run weekly

1. Put each store's current activity file in `input\<store>\weekly\activity\`.
2. Enter POS totals in `input\<store>\weekly\pos_controls.csv`.
3. Click `Run-Gift-Card-Reconciliation.cmd`.

The output workbook is created in `output\`. The first tab is `Reconciliation`; for weekly runs its title says `Week Ending` using the date found in the activity file.

## Run May 2026 store 9354 monthly

Place the summary and activity files in the folders shown above, then run:

```powershell
.\run_recon.ps1 `
  -Mode monthly `
  -Store 9354 `
  -Period 2026-05 `
  -PeriodEnd 2026-05-31 `
  -PosGiftCardIssue 11642.00 `
  -PosGiftCardPayment 49869.75
```

Alternative using `pos_controls.csv`:

```powershell
.\run_recon.ps1 `
  -Mode monthly `
  -Store 9354 `
  -Period 2026-05 `
  -PeriodEnd 2026-05-31 `
  -PosControls .\input\9354\2026-05\pos_controls.csv
```

Expected May 2026 results:

- Weekly activity activations: `$11,507.00`
- Weekly activity redemptions: `$49,867.48`
- POS Gift Card Issue: `$11,642.00`
- POS Gift Card Payment: `$49,869.75`
- POS issue variance: `$135.00`
- POS payment variance: `$2.27`
- POS net variance: `$132.73`

## Advanced weekly command

Weekly mode can run activity-to-POS without a summary file. If a weekly summary file is present, it is included; if not, summary-only values show as `N/A`.

Using `pos_controls.csv`:

```powershell
.\run_recon.ps1 `
  -Mode weekly `
  -Store 9354 `
  -Period 2026-W22 `
  -PeriodEnd 2026-05-31 `
  -InputDir .\input\9354\2026-W22 `
  -PosControls .\input\9354\2026-W22\pos_controls.csv
```

Using direct POS totals:

```powershell
.\run_recon.ps1 `
  -Mode weekly `
  -Store 9354 `
  -Period 2026-W22 `
  -PeriodEnd 2026-05-31 `
  -InputDir .\input\9354\2026-W22 `
  -PosGiftCardIssue 2730.00 `
  -PosGiftCardPayment 7446.47
```

## Create a new month

```powershell
.\scripts\new_period_folders.ps1 -Store 9354 -Period 2026-06
```

Then drop the new summary/activity files into the generated folders and update the POS controls file.

## Run tests

```powershell
.\run_tests.ps1
```

The synthetic test proves the math and workbook generation work. The actual-file test automatically runs if you copy real May 2026 files into `input/9354/2026-05`.

## Direct Python command

```powershell
.\.venv\Scripts\python.exe -m gift_card_recon `
  --mode monthly `
  --store 9354 `
  --period 2026-05 `
  --period-end 2026-05-31 `
  --input-dir .\input\9354\2026-05 `
  --output-dir .\output `
  --pos-gift-card-issue 11642.00 `
  --pos-gift-card-payment 49869.75
```

Weekly direct Python command:

```powershell
.\.venv\Scripts\python.exe -m gift_card_recon `
  --mode weekly `
  --store 9354 `
  --period 2026-W22 `
  --period-end 2026-05-31 `
  --input-dir .\input\9354\2026-W22 `
  --output-dir .\output `
  --pos-controls .\input\9354\2026-W22\pos_controls.csv
```

## Codex usage

Optional repo review:

```powershell
.\scripts\run_codex_review.ps1
```

Recommended Codex instruction:

```text
Review this gift card reconciliation repo. Run the pytest suite. Fix only genuine bugs. Do not change business rules unless a test proves they are wrong. Preserve POS controls on the Reconciliation tab.
```

## Business rules

- Activations are positive.
- Redemptions are negative in gift card activity detail.
- POS Gift Card Payment is treated as a positive POS control and compared against the absolute value of activity redemptions.
- Net Gift Card Impact equals issue less payment.
- Conversion promo redemptions are identified from promo codes listed in the summary when a summary is supplied.
- Weekly mode does not require a summary; summary-only rows are marked `N/A` when no weekly summary is supplied.
- The Reconciliation tab includes POS controls directly. No separate POS tie-out tab is required.
