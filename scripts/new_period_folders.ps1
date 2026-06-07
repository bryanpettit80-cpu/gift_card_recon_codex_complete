param(
    [string]$Store = "9354",
    [string]$Period = "2026-06"
)

$ErrorActionPreference = "Stop"
$Base = ".\input\$Store\$Period"
New-Item -ItemType Directory -Force -Path "$Base\summary" | Out-Null
New-Item -ItemType Directory -Force -Path "$Base\activity" | Out-Null

$Csv = "$Base\pos_controls.csv"
if (-not (Test-Path $Csv)) {
    "store,period,pos_gift_card_issue,pos_gift_card_payment" | Set-Content -Path $Csv -Encoding UTF8
    "$Store,$Period,," | Add-Content -Path $Csv -Encoding UTF8
}

Write-Host "Created input folders: $Base" -ForegroundColor Green
