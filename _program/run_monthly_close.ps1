param(
    [string]$Store = "",
    [string]$Period = "",
    [string]$InputRoot = ".\Monthly Close",
    [string]$InputDir = "",
    [string]$OutputDir = ".\Output",
    [string]$OutputFile = "",
    [string]$DardenPath = "",
    [string]$MicrosPath = "",
    [string]$MicrosWorkDir = ".\_program\tmp\monthly_close_micros",
    [string]$ArchiveRoot = ".\Archive - Old Files",
    [switch]$PrepareOnly,
    [switch]$NoStageWeekly,
    [switch]$NoCleanup,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $RepoRoot
$VenvPython = Join-Path $ProgramRoot ".venv\Scripts\python.exe"

if ($MicrosPath -eq "" -and $Store -ne "") {
    if ($Store -eq "9354") {
        $MicrosPath = "..\micros_data\RC-Richmond-current"
    } else {
        $MicrosPath = "..\GETLinkedData-VB"
    }
}

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
    "--input-root", $InputRoot,
    "--output-dir", $OutputDir,
    "--micros-work-dir", $MicrosWorkDir,
    "--archive-root", $ArchiveRoot
)

if ($Store -ne "") {
    $ArgsList += @("--store", $Store)
}

if ($Period -ne "") {
    $ArgsList += @("--period", $Period)
}

if ($MicrosPath -ne "") {
    $ArgsList += @("--micros-path", $MicrosPath)
}

if ($InputDir -ne "") {
    $ArgsList += @("--input-dir", $InputDir)
}

if ($OutputFile -ne "") {
    $ArgsList += @("--output-file", $OutputFile)
}

if ($DardenPath -ne "") {
    $ArgsList += @("--darden-path", $DardenPath)
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

& $VenvPython @ArgsList
exit $LASTEXITCODE
