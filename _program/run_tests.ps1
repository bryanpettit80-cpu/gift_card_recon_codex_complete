param(
    [switch]$SkipInstall,
    [switch]$SkipExcelIntegration
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -SkipInstall:$SkipInstall

$pytestArguments = @(
    "-m", "pytest", "-q", "-o", "cache_dir=$($Runtime.PytestCacheDir)"
)
if ($SkipExcelIntegration) {
    $pytestArguments += "--ignore=tests/test_windows_pdf_integration.py"
}

& $Runtime.PythonPath @pytestArguments
$exitCode = $LASTEXITCODE
exit $exitCode
