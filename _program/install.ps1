param(
    [switch]$ForceInstall
)

# One-time setup. Run from the repository root in Windows PowerShell or PowerShell 7+.
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -ForceInstall:$ForceInstall

foreach ($store in @("9354", "9355")) {
    $weeklyDir = Join-Path $RepoRoot "$store - Weekly"
    $activityDir = Join-Path $weeklyDir "activity"
    $posPath = Join-Path $weeklyDir "pos_controls.csv"
    New-Item -ItemType Directory -Force -Path $activityDir | Out-Null
    if (-not (Test-Path -LiteralPath $posPath)) {
        @(
            "store,period,pos_gift_card_issue,pos_gift_card_payment"
            "$store,auto,,"
        ) | Set-Content -LiteralPath $posPath -Encoding UTF8
    }
}

foreach ($folder in @("Monthly Close", "Output", "Archive - Old Files")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $RepoRoot $folder) | Out-Null
}

Write-Host "Setup complete. Local runtime: $($Runtime.RuntimeRoot)" -ForegroundColor Green
Write-Host "Use Run-Gift-Card-Reconciliation.cmd for weekly work or Run-Monthly-Close.cmd for month-end close."
