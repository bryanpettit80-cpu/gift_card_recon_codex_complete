# One-time setup. Run from the repository root in PowerShell 7+.
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ProgramRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ProgramRoot
$venvPython = Join-Path $ProgramRoot ".venv\Scripts\python.exe"

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath failed with exit code $LASTEXITCODE."
    }
}

function Test-VenvPython {
    if (-not (Test-Path $venvPython)) {
        return $false
    }

    & $venvPython --version *> $null
    if ($LASTEXITCODE -ne 0) {
        return $false
    }

    & $venvPython -m pip --version *> $null
    return $LASTEXITCODE -eq 0
}

if (-not (Test-VenvPython)) {
    Invoke-Checked python @("-m", "venv", "--clear", (Join-Path $ProgramRoot ".venv"))
}

Invoke-Checked $venvPython @("-m", "pip", "install", "-r", (Join-Path $ProgramRoot "requirements.txt"))
Invoke-Checked $venvPython @("-m", "pip", "install", "-e", $ProgramRoot)

foreach ($store in @("9354", "9355")) {
    $weeklyDir = Join-Path $RepoRoot "$store - Weekly"
    $activityDir = Join-Path $weeklyDir "activity"
    $posPath = Join-Path $weeklyDir "pos_controls.csv"
    New-Item -ItemType Directory -Force -Path $activityDir | Out-Null
    if (-not (Test-Path -LiteralPath $posPath)) {
        @(
            "store,period,pos_gift_card_issue,pos_gift_card_payment"
            "$store,auto,,"
        ) | Set-Content -LiteralPath $posPath -Encoding UTF8
    }
}

foreach ($folder in @("Monthly Close", "Output", "Archive - Old Files")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $RepoRoot $folder) | Out-Null
}

Write-Host "Setup complete. Click Run-Gift-Card-Reconciliation.cmd next." -ForegroundColor Green
