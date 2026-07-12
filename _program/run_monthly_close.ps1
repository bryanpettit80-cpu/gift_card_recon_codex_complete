param(
    [string]$Store = "",
    [string]$Period = "",
    [string]$InputRoot = ".\Monthly Close",
    [string]$InputDir = "",
    [string]$OutputDir = ".\Output",
    [string]$OutputFile = "",
    [string]$DardenPath = "",
    [string]$MicrosPath = "",
    [string]$MicrosWorkDir = "",
    [string]$ArchiveRoot = ".\Archive - Old Files",
    [switch]$PrepareOnly,
    [switch]$ReissueFromArchive,
    [switch]$NoStageWeekly,
    [switch]$NoCleanup,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $RepoRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -SkipInstall:$SkipInstall
$VenvPython = $Runtime.PythonPath

if ($MicrosWorkDir -eq "") {
    $MicrosWorkDir = $Runtime.MicrosExtractDir
}

if ($ReissueFromArchive) {
    if ($Store -eq "" -or $Period -eq "") {
        throw "-ReissueFromArchive requires both -Store and -Period."
    }
    if ($InputDir -ne "" -or $DardenPath -ne "" -or $MicrosPath -ne "") {
        throw "-ReissueFromArchive cannot be combined with -InputDir, -DardenPath, or -MicrosPath. Archived inputs are derived from the verified close manifest."
    }
}

if (-not $ReissueFromArchive -and $MicrosPath -eq "" -and $Store -ne "") {
    switch ($Store) {
        "9354" { $MicrosPath = "..\micros_data\RC-Richmond-current" }
        "9355" { $MicrosPath = "..\GETLinkedData-VB" }
        default { throw "Unsupported store '$Store'. Use 9354 or 9355." }
    }
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

if ($ReissueFromArchive) {
    $ArgsList += @("--reissue-from-archive")
}

if ($NoStageWeekly -or $ReissueFromArchive) {
    $ArgsList += @("--no-stage-weekly")
}

if ($NoCleanup -or $ReissueFromArchive) {
    $ArgsList += @("--no-cleanup")
}

& $VenvPython @ArgsList
$exitCode = $LASTEXITCODE
exit $exitCode
