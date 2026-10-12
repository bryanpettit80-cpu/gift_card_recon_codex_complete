param(
    [switch]$SkipInstall,
    [switch]$SkipExcelIntegration,
    [string]$RuntimeRoot = "",
    [string]$TestTempRoot = ""
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
if ([string]::IsNullOrWhiteSpace($RuntimeRoot)) {
    $RuntimeRoot = Get-GiftCardReconTestRuntimeRoot -ProgramRoot $ProgramRoot
}
if ([IO.Path]::GetFullPath($RuntimeRoot).TrimEnd('\', '/') -eq (Get-GiftCardReconRuntime).RuntimeRoot.TrimEnd('\', '/')) {
    throw "Tests require a separate runtime; the reporting runtime cannot be used."
}
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -SkipInstall:$SkipInstall -RuntimeRoot $RuntimeRoot
if ([string]::IsNullOrWhiteSpace($TestTempRoot)) {
    $TestTempRoot = Join-Path (Get-GiftCardReconRuntime).RuntimeRoot "t"
}
# Pytest owns only a unique child. Short paths avoid legacy Windows CopyFile2
# limits when test fixtures contain the full operator/archive folder hierarchy.
$CheckoutKey = Split-Path -Leaf (Get-GiftCardReconTestRuntimeRoot -ProgramRoot $ProgramRoot)
$TestBaseTemp = Join-Path ([IO.Path]::GetFullPath($TestTempRoot)) ("{0}-{1}" -f $CheckoutKey.Substring(0, 6), [Guid]::NewGuid().ToString("N").Substring(0, 8))
if (Test-Path -LiteralPath $TestBaseTemp) { throw "Test temporary path is unexpectedly occupied: $TestBaseTemp" }
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $TestBaseTemp))

$pytestArguments = @(
    "-m", "pytest", "-q", "-o", "cache_dir=$($Runtime.PytestCacheDir)", "--basetemp", $TestBaseTemp
)
if ($SkipExcelIntegration) {
    $pytestArguments += "--ignore=tests/test_windows_pdf_integration.py"
}

& $Runtime.PythonPath @pytestArguments
$exitCode = $LASTEXITCODE
exit $exitCode
