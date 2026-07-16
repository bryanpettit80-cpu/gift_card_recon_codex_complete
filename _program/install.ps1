param(
    [switch]$ForceInstall,
    [string]$OperationsRoot = ""
)

# One-time setup. Run from the repository root in Windows PowerShell or PowerShell 7+.
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -Profile Operator -ForceInstall:$ForceInstall
$OperatorInstaller = Join-Path $ProgramRoot "install_operator_assets.ps1"
if ([string]::IsNullOrWhiteSpace($OperationsRoot)) {
    & $OperatorInstaller
}
else {
    & $OperatorInstaller -OperationsRoot $OperationsRoot
}

Write-Host "Setup complete. Operator runtime: $($Runtime.RuntimeRoot)" -ForegroundColor Green
Write-Host "Use 'Run Weekly Gift Card Reconciliation.cmd' for weekly work or 'Run Monthly Gift Card Close.cmd' for month-end close."
