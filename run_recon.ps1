param(
    [ValidateSet("monthly", "weekly")]
    [string]$Mode = "monthly",
    [string]$Store = "9354",
    [string]$Period = "2026-05",
    [string]$PeriodEnd = "2026-05-31",
    [string]$InputDir = ".\input\9354\2026-05",
    [string]$OutputDir = ".\output",
    [string]$PosControls = "",
    [decimal]$PosGiftCardIssue = 11642.00,
    [decimal]$PosGiftCardPayment = 49869.75,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
}

if (-not $SkipInstall) {
    .\.venv\Scripts\python.exe -m pip install --upgrade pip
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe -m pip install -e .
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

if (-not (Test-Path $InputDir)) {
    throw "Input directory not found: $InputDir"
}

$ActivityFiles = @(
    Get-ChildItem -Path (Join-Path $InputDir "activity") -Filter "*Gift Card Activity*.xls*" -File -ErrorAction SilentlyContinue
    Get-ChildItem -Path $InputDir -Filter "*Gift Card Activity*.xls*" -File -ErrorAction SilentlyContinue
) | Select-Object -Unique

if ($ActivityFiles.Count -eq 0) {
    throw "No Gift Card Activity .xls/.xlsx files found for $Mode mode in $InputDir or activity\."
}

$SummaryFiles = @(
    Get-ChildItem -Path (Join-Path $InputDir "summary") -Filter "*Gift Card Summary*.xlsx" -File -ErrorAction SilentlyContinue
    Get-ChildItem -Path $InputDir -Filter "*Gift Card Summary*.xlsx" -File -ErrorAction SilentlyContinue
) | Select-Object -Unique

if ($Mode -eq "monthly" -and $SummaryFiles.Count -ne 1) {
    throw "Monthly mode requires exactly one Gift Card Summary .xlsx file in $InputDir or summary\. Found $($SummaryFiles.Count)."
}

if ($Mode -eq "weekly" -and $SummaryFiles.Count -gt 1) {
    throw "Weekly mode allows at most one optional Gift Card Summary .xlsx file in $InputDir or summary\. Found $($SummaryFiles.Count)."
}

if ($PosControls -eq "" -and -not (Test-Path (Join-Path $InputDir "pos_controls.csv")) -and -not (Test-Path (Join-Path $InputDir "pos_controls.xlsx")) -and ($null -eq $PosGiftCardIssue -or $null -eq $PosGiftCardPayment)) {
    throw "POS controls missing. Provide -PosControls, place pos_controls.csv/xlsx in $InputDir, or pass both -PosGiftCardIssue and -PosGiftCardPayment."
}

$ArgsList = @(
    "-m", "gift_card_recon",
    "--mode", $Mode,
    "--store", $Store,
    "--period", $Period,
    "--period-end", $PeriodEnd,
    "--input-dir", $InputDir,
    "--output-dir", $OutputDir
)

if ($PosControls -ne "") {
    $ArgsList += @("--pos-controls", $PosControls)
} else {
    $ArgsList += @("--pos-gift-card-issue", $PosGiftCardIssue, "--pos-gift-card-payment", $PosGiftCardPayment)
}

.\.venv\Scripts\python.exe @ArgsList
