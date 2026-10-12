param(
    [switch]$ForceInstall,
    [string]$OperationsRoot = "",
    [string]$DropboxRoot = "",
    [string]$ConfigPath = ""
)

# One-time setup. Run from the repository root in Windows PowerShell or PowerShell 7+.
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -ForceInstall:$ForceInstall
$OperatorInstaller = Join-Path $ProgramRoot "install_operator_assets.ps1"
if ([string]::IsNullOrWhiteSpace($OperationsRoot)) {
    $ExistingConfiguration = Get-GiftCardReconPathConfiguration -ConfigPath $ConfigPath
    if ($null -ne $ExistingConfiguration) {
        $OperationsRoot = $ExistingConfiguration.OperationsRoot
        if ([string]::IsNullOrWhiteSpace($DropboxRoot)) { $DropboxRoot = $ExistingConfiguration.DropboxRoot }
    } elseif ((Split-Path -Leaf $RepoRoot) -eq "Gift Card Reconciliation Automation") {
        $OperationsRoot = Split-Path -Parent $RepoRoot
    } else { $OperationsRoot = $RepoRoot }
}
if (-not [IO.Path]::IsPathRooted($OperationsRoot)) { $OperationsRoot = Join-Path $RepoRoot $OperationsRoot }
& $OperatorInstaller -OperationsRoot $OperationsRoot
Set-GiftCardReconPathConfiguration -ProgramRoot $ProgramRoot -OperationsRoot $OperationsRoot -DropboxRoot $DropboxRoot -ConfigPath $ConfigPath

Write-Host "Setup complete. Local runtime: $($Runtime.RuntimeRoot)" -ForegroundColor Green
Write-Host "Use 'Run Weekly Gift Card Reconciliation.cmd' for weekly work or 'Run Monthly Gift Card Close.cmd' for month-end close."
