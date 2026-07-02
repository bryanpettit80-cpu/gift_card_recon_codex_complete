param(
    [string]$Store = "9355",
    [string]$Period = "2026-06",
    [string]$PeriodEnd = "",
    [string]$InputRoot = ".\input",
    [string]$InputDir = "",
    [string]$OutputDir = ".\output",
    [string]$OutputFile = "",
    [string]$MicrosPath = ".\_inspect_micros3700",
    [string]$MicrosWorkDir = ".\tmp\monthly_close_micros",
    [switch]$PrepareOnly,
    [switch]$NoStageWeekly,
    [switch]$NoBoundaryAdjustment,
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

if ($InputDir -eq "") {
    $InputDir = Join-Path (Join-Path $InputRoot $Store) $Period
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $MicrosWorkDir | Out-Null

$ArgsList = @(
    "-m", "gift_card_recon.monthly_close",
    "--store", $Store,
    "--period", $Period,
    "--input-root", $InputRoot,
    "--input-dir", $InputDir,
    "--output-dir", $OutputDir,
    "--micros-path", $MicrosPath,
    "--micros-work-dir", $MicrosWorkDir
)

if ($PeriodEnd -ne "") {
    $ArgsList += @("--period-end", $PeriodEnd)
}

if ($OutputFile -ne "") {
    $ArgsList += @("--output-file", $OutputFile)
}

if ($PrepareOnly) {
    $ArgsList += @("--prepare-only")
}

if ($NoStageWeekly) {
    $ArgsList += @("--no-stage-weekly")
}

if ($NoBoundaryAdjustment) {
    $ArgsList += @("--no-boundary-adjustment")
}

.\.venv\Scripts\python.exe @ArgsList
