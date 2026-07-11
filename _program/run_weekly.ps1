param(
    [string]$InputRoot = ".",
    [string]$OutputDir = ".\Output",
    [string]$MonthlyCloseRoot = ".\Monthly Close",
    [string[]]$Store = @(),
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $RepoRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -SkipInstall:$SkipInstall

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$ArgsList = @(
    "-m", "gift_card_recon.auto_run",
    "--input-root", $InputRoot,
    "--output-dir", $OutputDir,
    "--monthly-close-root", $MonthlyCloseRoot
)
foreach ($storeNumber in $Store) {
    if (-not [string]::IsNullOrWhiteSpace($storeNumber)) {
        $ArgsList += @("--store", $storeNumber)
    }
}

& $Runtime.PythonPath @ArgsList
$exitCode = $LASTEXITCODE
exit $exitCode
