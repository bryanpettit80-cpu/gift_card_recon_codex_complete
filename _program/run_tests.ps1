$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
Set-Location $ProgramRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    & (Join-Path $ProgramRoot "install.ps1")
}

.\.venv\Scripts\python.exe -m pytest -q
