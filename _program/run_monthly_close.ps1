param(
    [string]$OperationsRoot = "",
    [string]$Store = "",
    [string]$Period = "",
    [string]$InputRoot = "",
    [string]$InputDir = "",
    [string]$OutputDir = "",
    [string]$OutputFile = "",
    [string]$DardenPath = "",
    [string]$MicrosPath = "",
    [string]$MicrosWorkDir = "",
    [string]$ArchiveRoot = "",
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
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -Profile Operator -SkipInstall:$SkipInstall
$VenvPython = $Runtime.PythonPath

function ConvertTo-OperationsPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

$UseOrganizedLayout = -not [string]::IsNullOrWhiteSpace($OperationsRoot)
if ($UseOrganizedLayout) {
    $OperationsRoot = ConvertTo-OperationsPath -Path $OperationsRoot -BasePath $RepoRoot
}
else {
    # Backward-compatible local-development mode for a checkout that has not
    # yet been installed into the deployed operator layout.
    $OperationsRoot = $RepoRoot
}

if ([string]::IsNullOrWhiteSpace($InputRoot)) {
    $InputRoot = if ($UseOrganizedLayout) { "02 Monthly Close Inputs" } else { "Monthly Close" }
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = if ($UseOrganizedLayout) { "03 Finished Reports" } else { "Output" }
}
if ([string]::IsNullOrWhiteSpace($ArchiveRoot)) {
    $ArchiveRoot = if ($UseOrganizedLayout) { "04 Archive" } else { "Archive - Old Files" }
}

$InputRoot = ConvertTo-OperationsPath -Path $InputRoot -BasePath $OperationsRoot
$OutputDir = ConvertTo-OperationsPath -Path $OutputDir -BasePath $OperationsRoot
$ArchiveRoot = ConvertTo-OperationsPath -Path $ArchiveRoot -BasePath $OperationsRoot
if (-not [string]::IsNullOrWhiteSpace($InputDir)) {
    $InputDir = ConvertTo-OperationsPath -Path $InputDir -BasePath $OperationsRoot
}
if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = ConvertTo-OperationsPath -Path $OutputFile -BasePath $OperationsRoot
}
if (-not [string]::IsNullOrWhiteSpace($DardenPath)) {
    $DardenPath = ConvertTo-OperationsPath -Path $DardenPath -BasePath $OperationsRoot
}
if (-not [string]::IsNullOrWhiteSpace($MicrosPath)) {
    $MicrosPath = ConvertTo-OperationsPath -Path $MicrosPath -BasePath $OperationsRoot
}

if ($MicrosWorkDir -eq "") {
    $MicrosWorkDir = $Runtime.MicrosExtractDir
}
else {
    $MicrosWorkDir = ConvertTo-OperationsPath -Path $MicrosWorkDir -BasePath $OperationsRoot
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
    $DropboxRoot = Split-Path -Parent $OperationsRoot
    switch ($Store) {
        "9354" { $MicrosPath = Join-Path $DropboxRoot "micros_data\RC-Richmond-current" }
        "9355" { $MicrosPath = Join-Path $DropboxRoot "GETLinkedData-VB" }
        default { throw "Unsupported store '$Store'. Use 9354 or 9355." }
    }
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $MicrosWorkDir | Out-Null
New-Item -ItemType Directory -Force -Path $ArchiveRoot | Out-Null

$ArgsList = @(
    "-m", "gift_card_recon.monthly_close",
    "--operations-root", $OperationsRoot,
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

Set-Location $OperationsRoot
& $VenvPython @ArgsList
$exitCode = $LASTEXITCODE
exit $exitCode
