param(
    [string]$Store = "9355",
    [string]$Period = "2026-06",
    [string]$PeriodEnd = "",
    [string]$InputDir = "",
    [string]$OutputDir = ".\output",
    [string]$OutputFile = "",
    [string]$MicrosPath = ".\_inspect_micros3700",
    [string]$MicrosWorkDir = ".\tmp\monthly_close_micros",
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
    $InputDir = ".\input\$Store\$Period"
}

if (-not (Test-Path $InputDir)) {
    throw "Input directory not found: $InputDir"
}

if (-not (Test-Path $MicrosPath)) {
    throw "Micros export path not found: $MicrosPath"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $MicrosWorkDir | Out-Null

$ArgsList = @(
    "-m", "gift_card_recon.monthly_close",
    "--store", $Store,
    "--period", $Period,
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

if ($NoBoundaryAdjustment) {
    $ArgsList += @("--no-boundary-adjustment")
}

.\.venv\Scripts\python.exe @ArgsList
