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
    "03 Finished Reports\Monthly Close - Review Required",
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
        Template = Join-Path $ProjectRoot "templates\Check Gift Card Reconciliation Health.cmd"
        Target = Join-Path $OperationsRoot "Check Gift Card Reconciliation Health.cmd"
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

$TransactionId = [guid]::NewGuid().ToString("N").Substring(0, 12)
$PreparedAssets = @()

try {
    # Validate and stage every asset before changing any live operator file.
    foreach ($Asset in $Assets) {
        if (-not (Test-Path -LiteralPath $Asset.Template -PathType Leaf)) {
            throw "Operator file template not found: $($Asset.Template)"
        }
        if ((Test-Path -LiteralPath $Asset.Target) -and -not (Test-Path -LiteralPath $Asset.Target -PathType Leaf)) {
            throw "Operator file target exists but is not a file: $($Asset.Target)"
        }

        $TargetFolder = Split-Path -Parent $Asset.Target
        New-Item -ItemType Directory -Force -Path $TargetFolder | Out-Null
        $AssetIndex = $PreparedAssets.Count
        $StagePath = Join-Path $TargetFolder (".gcs-{0}-{1}.tmp" -f $TransactionId, $AssetIndex)
        $Prepared = [pscustomobject]@{
            Template = $Asset.Template
            Target = $Asset.Target
            Stage = $StagePath
            Index = $AssetIndex
            Hash = $null
            Backup = $null
            Published = $false
        }
        $PreparedAssets += $Prepared

        Copy-Item -LiteralPath $Asset.Template -Destination $StagePath -Force
        $TemplateHash = (Get-FileHash -LiteralPath $Asset.Template -Algorithm SHA256).Hash
        $StageHash = (Get-FileHash -LiteralPath $StagePath -Algorithm SHA256).Hash
        if ($TemplateHash -ne $StageHash) {
            throw "Operator file staging failed; staged copy does not match template: $($Asset.Target)"
        }
        $Prepared.Hash = $TemplateHash
    }

    try {
        # Back up the complete live set before publishing any staged asset.
        foreach ($Prepared in $PreparedAssets) {
            if (Test-Path -LiteralPath $Prepared.Target -PathType Leaf) {
                $Prepared.Backup = Join-Path (Split-Path -Parent $Prepared.Target) `
                    (".gcb-{0}-{1}.tmp" -f $TransactionId, $Prepared.Index)
                Move-Item -LiteralPath $Prepared.Target -Destination $Prepared.Backup -Force
            }
        }

        foreach ($Prepared in $PreparedAssets) {
            Move-Item -LiteralPath $Prepared.Stage -Destination $Prepared.Target -Force
            $Prepared.Published = $true
            $TargetHash = (Get-FileHash -LiteralPath $Prepared.Target -Algorithm SHA256).Hash
            if ($Prepared.Hash -ne $TargetHash) {
                throw "Operator file refresh failed; target does not match template: $($Prepared.Target)"
            }
        }
    }
    catch {
        $PublishFailure = $_
        $RollbackErrors = @()
        for ($Index = $PreparedAssets.Count - 1; $Index -ge 0; $Index--) {
            $Prepared = $PreparedAssets[$Index]
            if ($Prepared.Published -and (Test-Path -LiteralPath $Prepared.Target)) {
                try {
                    Remove-Item -LiteralPath $Prepared.Target -Force
                }
                catch {
                    $RollbackErrors += "could not remove $($Prepared.Target): $($_.Exception.Message)"
                }
            }
            if ($Prepared.Backup -and (Test-Path -LiteralPath $Prepared.Backup -PathType Leaf)) {
                try {
                    Move-Item -LiteralPath $Prepared.Backup -Destination $Prepared.Target -Force
                }
                catch {
                    $RollbackErrors += "could not restore $($Prepared.Target): $($_.Exception.Message)"
                }
            }
        }
        if ($RollbackErrors.Count -gt 0) {
            throw "Operator file refresh failed and rollback was incomplete: $($RollbackErrors -join '; '). Original error: $($PublishFailure.Exception.Message)"
        }
        throw "Operator file refresh failed; every prior operator file was restored. $($PublishFailure.Exception.Message)"
    }

    foreach ($Prepared in $PreparedAssets) {
        if ($Prepared.Backup -and (Test-Path -LiteralPath $Prepared.Backup -PathType Leaf)) {
            try {
                Remove-Item -LiteralPath $Prepared.Backup -Force
            }
            catch {
                Write-Warning "Operator file is current, but its prior backup could not be removed: $($Prepared.Backup)"
            }
        }
        Write-Host "Operator file verified: $($Prepared.Target)"
    }
}
finally {
    foreach ($Prepared in $PreparedAssets) {
        if ($Prepared.Stage -and (Test-Path -LiteralPath $Prepared.Stage -PathType Leaf)) {
            Remove-Item -LiteralPath $Prepared.Stage -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "Gift Card Reconciliation operator layout is ready: $OperationsRoot"
