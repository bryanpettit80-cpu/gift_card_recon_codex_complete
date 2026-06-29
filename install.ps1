# One-time setup. Run from the repository root in PowerShell 7+.
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

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

$venvPython = ".\.venv\Scripts\python.exe"

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
    Invoke-Checked python @("-m", "venv", "--clear", ".venv")
}

Invoke-Checked $venvPython @("-m", "pip", "install", "-r", "requirements.txt")
Invoke-Checked $venvPython @("-m", "pip", "install", "-e", ".")

Write-Host "Setup complete. Click Run-Gift-Card-Reconciliation.cmd next." -ForegroundColor Green
