$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    .\install.ps1
}

.\.venv\Scripts\python.exe -m pytest -q
