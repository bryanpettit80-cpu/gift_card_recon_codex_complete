param(
    [string]$OperationsRoot = "",
    [string]$DropboxRoot = "",
    [string]$ConfigPath = "",
    [string]$InputRoot = "",
    [string]$OutputDir = "",
    [string]$MonthlyCloseRoot = "",
    [string]$ArchiveRoot = "",
    [string]$ReviewRoot = "",
    [string[]]$Store = @(),
    [string[]]$ApproveCorrection = @(),
    [string[]]$RestoreRevision = @(),
    [switch]$Readiness,
    [switch]$Json,
    [switch]$ShowConfiguration,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $RepoRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Get-GiftCardReconRuntime
$PathConfiguration = Get-GiftCardReconPathConfiguration -ConfigPath $ConfigPath
if ($null -ne $PathConfiguration) {
    if ([string]::IsNullOrWhiteSpace($OperationsRoot)) { $OperationsRoot = $PathConfiguration.OperationsRoot }
    if ([string]::IsNullOrWhiteSpace($DropboxRoot) -and ([IO.Path]::GetFullPath($OperationsRoot)).TrimEnd('\', '/') -eq ([string]$PathConfiguration.OperationsRoot).TrimEnd('\', '/')) {
        $DropboxRoot = $PathConfiguration.DropboxRoot
    }
}

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
    # yet been installed into the nested program-only layout.
    $OperationsRoot = $RepoRoot
}
$DropboxRoot = Resolve-GiftCardReconDropboxRoot `
    -OperationsRoot $OperationsRoot `
    -DropboxRoot $DropboxRoot

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

if ($ShowConfiguration) {
    [ordered]@{ OperationsRoot = $OperationsRoot; DropboxRoot = $DropboxRoot; InputRoot = $InputRoot; OutputDir = $OutputDir; ArchiveRoot = $ArchiveRoot; RuntimeRoot = $Runtime.RuntimeRoot } | ConvertTo-Json
    exit 0
}
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -SkipInstall
if ($Readiness) {
    if ($ApproveCorrection.Count -gt 0 -or $RestoreRevision.Count -gt 0) { throw "-Readiness cannot approve a correction or restore a revision." }
    $ReadinessArguments = @("-m", "gift_card_recon.coordinator", "weekly", "--operations-root", $OperationsRoot, "--dropbox-root", $DropboxRoot, "--input-root", $InputRoot, "--archive-root", $ArchiveRoot)
    foreach ($storeNumber in $Store) { $ReadinessArguments += @("--store", $storeNumber) }
    & $Runtime.PythonPath @ReadinessArguments
    exit $LASTEXITCODE
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $ArchiveRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ReviewRoot | Out-Null
$ArgsList = @(
    "-m", "gift_card_recon.auto_run",
    "--operations-root", $OperationsRoot,
    "--dropbox-root", $DropboxRoot,
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
foreach ($Approval in $ApproveCorrection) { $ArgsList += @("--approve-correction", $Approval) }
foreach ($Revision in $RestoreRevision) { $ArgsList += @("--restore-revision", $Revision) }
if ($Json) { $ArgsList += "--json" }

Set-Location $OperationsRoot
& $Runtime.PythonPath @ArgsList
$exitCode = $LASTEXITCODE
exit $exitCode
