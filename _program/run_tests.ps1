param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -SkipInstall:$SkipInstall

& $Runtime.PythonPath -m pytest -q -o "cache_dir=$($Runtime.PytestCacheDir)"
$exitCode = $LASTEXITCODE
exit $exitCode
