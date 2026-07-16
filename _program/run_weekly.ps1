param(
    [string]$OperationsRoot = "",
    [string]$InputRoot = "",
    [string]$OutputDir = "",
    [string]$MonthlyCloseRoot = "",
    [string]$ArchiveRoot = "",
    [string]$ReviewRoot = "",
    [string[]]$Store = @(),
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $RepoRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -Profile Operator -SkipInstall:$SkipInstall

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
    $InputRoot = if ($UseOrganizedLayout) { "01 Weekly Gift Card Activity Reports" } else { "." }
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = if ($UseOrganizedLayout) { "03 Finished Reports\Weekly" } else { "Output" }
}
if ([string]::IsNullOrWhiteSpace($MonthlyCloseRoot)) {
    $MonthlyCloseRoot = if ($UseOrganizedLayout) { "02 Monthly Close Inputs" } else { "Monthly Close" }
}
if ([string]::IsNullOrWhiteSpace($ArchiveRoot)) {
    $ArchiveRoot = if ($UseOrganizedLayout) { "04 Archive\Weekly Reconciliation" } else { "Archive - Old Files\Weekly Reconciliation" }
}
if ([string]::IsNullOrWhiteSpace($ReviewRoot)) {
    $ReviewRoot = "_automation_runs\review"
}

$InputRoot = ConvertTo-OperationsPath -Path $InputRoot -BasePath $OperationsRoot
$OutputDir = ConvertTo-OperationsPath -Path $OutputDir -BasePath $OperationsRoot
$MonthlyCloseRoot = ConvertTo-OperationsPath -Path $MonthlyCloseRoot -BasePath $OperationsRoot
$ArchiveRoot = ConvertTo-OperationsPath -Path $ArchiveRoot -BasePath $OperationsRoot
$ReviewRoot = ConvertTo-OperationsPath -Path $ReviewRoot -BasePath $OperationsRoot

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $ArchiveRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ReviewRoot | Out-Null
$ArgsList = @(
    "-m", "gift_card_recon.auto_run",
    "--operations-root", $OperationsRoot,
    "--input-root", $InputRoot,
    "--output-dir", $OutputDir,
    "--monthly-close-root", $MonthlyCloseRoot,
    "--archive-root", $ArchiveRoot,
    "--review-root", $ReviewRoot
)
foreach ($storeNumber in $Store) {
    if (-not [string]::IsNullOrWhiteSpace($storeNumber)) {
        $ArgsList += @("--store", $storeNumber)
    }
}

Set-Location $OperationsRoot
& $Runtime.PythonPath @ArgsList
$exitCode = $LASTEXITCODE
exit $exitCode
