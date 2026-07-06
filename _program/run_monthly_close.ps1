param(
    [string]$Store = "9355",
    [string]$Period = "FY27-M01",
    [string]$PeriodEnd = "",
    [string]$InputRoot = ".\Monthly Close",
    [string]$InputDir = "",
    [string]$OutputDir = ".\Output",
    [string]$OutputFile = "",
    [string]$MicrosPath = ".\_inspect_micros3700",
    [string]$MicrosWorkDir = ".\_program\tmp\monthly_close_micros",
    [string]$ArchiveRoot = ".\Archive - Old Files",
    [switch]$PrepareOnly,
    [switch]$NoStageWeekly,
    [switch]$NoCleanup,
    [switch]$NoBoundaryAdjustment,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $RepoRoot
$VenvPython = Join-Path $ProgramRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    python -m venv (Join-Path $ProgramRoot ".venv")
}

if (-not $SkipInstall) {
    & $VenvPython -m pip install --upgrade pip
    & $VenvPython -m pip install -r (Join-Path $ProgramRoot "requirements.txt")
    & $VenvPython -m pip install -e $ProgramRoot
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $MicrosWorkDir | Out-Null
New-Item -ItemType Directory -Force -Path $ArchiveRoot | Out-Null

$ArgsList = @(
    "-m", "gift_card_recon.monthly_close",
    "--store", $Store,
    "--period", $Period,
    "--input-root", $InputRoot,
    "--output-dir", $OutputDir,
    "--micros-path", $MicrosPath,
    "--micros-work-dir", $MicrosWorkDir,
    "--archive-root", $ArchiveRoot
)

if ($InputDir -ne "") {
    $ArgsList += @("--input-dir", $InputDir)
}

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

if ($NoCleanup) {
    $ArgsList += @("--no-cleanup")
}

if ($NoBoundaryAdjustment) {
    $ArgsList += @("--no-boundary-adjustment")
}

& $VenvPython @ArgsList
