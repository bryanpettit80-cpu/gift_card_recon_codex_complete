# One-click weekly runner. Drop each store's current activity file into:
# input\<store>\weekly\activity
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    .\install.ps1
}

New-Item -ItemType Directory -Force -Path ".\output" | Out-Null
.\.venv\Scripts\python.exe -m gift_card_recon.auto_run --input-root ".\input" --output-dir ".\output"

Write-Host ""
Write-Host "Open the output folder to find the finished workbook(s):" -ForegroundColor Green
Write-Host (Join-Path $RepoRoot "output")
