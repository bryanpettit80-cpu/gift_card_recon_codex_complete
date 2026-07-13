[CmdletBinding()]
param(
    [string]$OperationsRoot = ""
)

$ErrorActionPreference = "Stop"

$ProgramRoot = Split-Path -Parent $PSCommandPath
$ProjectRoot = Split-Path -Parent $ProgramRoot

if ([string]::IsNullOrWhiteSpace($OperationsRoot)) {
    if ((Split-Path -Leaf $ProjectRoot) -eq "Gift Card Reconciliation Automation") {
        $OperationsRoot = Split-Path -Parent $ProjectRoot
    }
    else {
        # Safe local-development default before the program-only repo is nested.
        $OperationsRoot = $ProjectRoot
    }
}

if ([System.IO.Path]::IsPathRooted($OperationsRoot)) {
    $OperationsRoot = [System.IO.Path]::GetFullPath($OperationsRoot)
}
else {
    $OperationsRoot = [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot $OperationsRoot))
}

$RequiredFolders = @(
    "01 Weekly Gift Card Activity Reports\9354 Richmond\activity",
    "01 Weekly Gift Card Activity Reports\9355 Virginia Beach\activity",
    "02 Monthly Close Inputs\Darden Reports - Drop Here",
    "02 Monthly Close Inputs\9354 Richmond",
    "02 Monthly Close Inputs\9355 Virginia Beach",
    "03 Finished Reports\Weekly\9354 Richmond",
    "03 Finished Reports\Weekly\9355 Virginia Beach",
    "03 Finished Reports\Monthly Close",
    "03 Finished Reports\Review Required",
    "04 Archive\Monthly Close",
    "04 Archive\Weekly Reconciliation",
    "04 Archive\Generated Reports",
    "04 Archive\Legacy Reconciliation",
    "04 Archive\Cleanup Manifests",
    "_automation_runs\logs",
    "_automation_runs\qa",
    "_automation_runs\review",
    "_automation_runs\test-output"
)

foreach ($RelativeFolder in $RequiredFolders) {
    New-Item -ItemType Directory -Force -Path (Join-Path $OperationsRoot $RelativeFolder) | Out-Null
}

$Assets = @(
    [pscustomobject]@{
        Template = Join-Path $ProjectRoot "templates\00 START HERE - Gift Card Reconciliation.txt"
        Target = Join-Path $OperationsRoot "00 START HERE - Gift Card Reconciliation.txt"
    },
    [pscustomobject]@{
        Template = Join-Path $ProjectRoot "templates\Run Weekly Gift Card Reconciliation.cmd"
        Target = Join-Path $OperationsRoot "Run Weekly Gift Card Reconciliation.cmd"
    },
    [pscustomobject]@{
        Template = Join-Path $ProjectRoot "templates\Run Monthly Gift Card Close.cmd"
        Target = Join-Path $OperationsRoot "Run Monthly Gift Card Close.cmd"
    },
    [pscustomobject]@{
        Template = Join-Path $ProjectRoot "templates\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"
        Target = Join-Path $OperationsRoot "01 Weekly Gift Card Activity Reports\9354 Richmond\activity\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"
    },
    [pscustomobject]@{
        Template = Join-Path $ProjectRoot "templates\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"
        Target = Join-Path $OperationsRoot "01 Weekly Gift Card Activity Reports\9355 Virginia Beach\activity\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"
    },
    [pscustomobject]@{
        Template = Join-Path $ProjectRoot "templates\00 DROP DARDEN CREDIT MEMOS HERE.txt"
        Target = Join-Path $OperationsRoot "02 Monthly Close Inputs\Darden Reports - Drop Here\00 DROP DARDEN CREDIT MEMOS HERE.txt"
    }
)

foreach ($Asset in $Assets) {
    if (-not (Test-Path -LiteralPath $Asset.Template -PathType Leaf)) {
        throw "Operator file template not found: $($Asset.Template)"
    }

    Copy-Item -LiteralPath $Asset.Template -Destination $Asset.Target -Force

    $TemplateHash = (Get-FileHash -LiteralPath $Asset.Template -Algorithm SHA256).Hash
    $TargetHash = (Get-FileHash -LiteralPath $Asset.Target -Algorithm SHA256).Hash
    if ($TemplateHash -ne $TargetHash) {
        throw "Operator file refresh failed; target does not match template: $($Asset.Target)"
    }

    Write-Host "Operator file verified: $($Asset.Target)"
}

Write-Host "Gift Card Reconciliation operator layout is ready: $OperationsRoot"
