param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $ProgramRoot
. (Join-Path $ProgramRoot "runtime.ps1")
$Runtime = Initialize-GiftCardReconRuntime -ProgramRoot $ProgramRoot -Profile Development -SkipInstall:$SkipInstall

$systemRoot = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath("$env:SystemDrive\"))
$testTemp = Join-Path $systemRoot ("gcrt-{0}-{1}" -f $PID, ([guid]::NewGuid().ToString("N").Substring(0, 8)))
$resolvedTestTemp = [System.IO.Path]::GetFullPath($testTemp)
if (
    [System.IO.Path]::GetDirectoryName($resolvedTestTemp) -ne $systemRoot -or
    -not [System.IO.Path]::GetFileName($resolvedTestTemp).StartsWith("gcrt-", [System.StringComparison]::OrdinalIgnoreCase)
) {
    throw "Refusing to use an unexpected pytest temporary path: $resolvedTestTemp"
}

try {
    New-Item -ItemType Directory -Path $resolvedTestTemp -Force | Out-Null
    & $Runtime.PythonPath -m pytest -q `
        -o "cache_dir=$($Runtime.PytestCacheDir)" `
        "--basetemp=$resolvedTestTemp"
    $exitCode = $LASTEXITCODE
}
finally {
    if (Test-Path -LiteralPath $resolvedTestTemp) {
        Remove-Item -LiteralPath $resolvedTestTemp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

exit $exitCode
