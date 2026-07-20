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

$RetiredAssets = @(
    [pscustomobject]@{
        Target = Join-Path $OperationsRoot "Check Gift Card Reconciliation Health.cmd"
        # The only released versions of the retired launcher used LF or CRLF
        # line endings. An unrecognized same-name file belongs to the operator
        # and must be preserved rather than guessed to be ours.
        KnownSha256 = @(
            "797A0B78B611DBE93317457422FA68043A3AAB35B535F9DB4D54A98988199955",
            "3614BEA1DCA286BC14EAA9116853FE4CFE40B50CF6E696FC284AD232802DB1EA"
        )
    }
)

$TransactionId = [guid]::NewGuid().ToString("N").Substring(0, 12)
$PreparedAssets = @()
$PreparedRetirements = @()

try {
    # Validate and stage the complete managed set before changing a live file.
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
        if (Test-Path -LiteralPath $StagePath) {
            throw "Operator file staging path is unexpectedly occupied: $StagePath"
        }
        $Prepared = [pscustomobject]@{
            Target = $Asset.Target
            Stage = $StagePath
            Index = $AssetIndex
            Hash = $null
            OriginalHash = $null
            Backup = $null
            BackedUp = $false
            Published = $false
            Retired = $false
        }
        $PreparedAssets += $Prepared

        Copy-Item -LiteralPath $Asset.Template -Destination $StagePath
        $TemplateHash = (Get-FileHash -LiteralPath $Asset.Template -Algorithm SHA256).Hash
        $StageHash = (Get-FileHash -LiteralPath $StagePath -Algorithm SHA256).Hash
        if ($TemplateHash -ne $StageHash) {
            throw "Operator file staging failed; staged copy does not match template: $($Asset.Target)"
        }
        $Prepared.Hash = $TemplateHash
    }

    # Retire only the exact launcher released by this project. A modified or
    # unrelated same-name file is deliberately left untouched with a warning.
    foreach ($RetiredAsset in $RetiredAssets) {
        if (-not (Test-Path -LiteralPath $RetiredAsset.Target)) {
            continue
        }
        if (-not (Test-Path -LiteralPath $RetiredAsset.Target -PathType Leaf)) {
            Write-Warning "Preserving unrecognized retired-asset path because it is not a file: $($RetiredAsset.Target)"
            continue
        }

        try {
            $RetiredHash = (Get-FileHash -LiteralPath $RetiredAsset.Target -Algorithm SHA256).Hash
        }
        catch {
            Write-Warning "Preserving retired-asset candidate because its hash could not be verified: $($RetiredAsset.Target). $($_.Exception.Message)"
            continue
        }
        if ($RetiredHash -notin $RetiredAsset.KnownSha256) {
            Write-Warning "Preserving unrecognized same-name operator file: $($RetiredAsset.Target)"
            continue
        }

        $PreparedRetirements += [pscustomobject]@{
            Target = $RetiredAsset.Target
            Stage = $null
            Index = $PreparedAssets.Count + $PreparedRetirements.Count
            Hash = $RetiredHash
            OriginalHash = $RetiredHash
            Backup = $null
            BackedUp = $false
            Published = $false
            Retired = $true
        }
    }

    $TransactionItems = @($PreparedAssets) + @($PreparedRetirements)
    try {
        # Back up the complete live set before publishing any staged asset.
        # A recognized retired launcher participates in the same transaction,
        # so a late failure restores it along with every managed file.
        foreach ($Prepared in $TransactionItems) {
            if (-not (Test-Path -LiteralPath $Prepared.Target)) {
                continue
            }
            if (-not (Test-Path -LiteralPath $Prepared.Target -PathType Leaf)) {
                throw "Operator file target changed after preflight and is not a file: $($Prepared.Target)"
            }

            $CurrentHash = (Get-FileHash -LiteralPath $Prepared.Target -Algorithm SHA256).Hash
            if ($Prepared.Retired -and $CurrentHash -ne $Prepared.Hash) {
                throw "Retired operator file changed after preflight and was preserved: $($Prepared.Target)"
            }
            $Prepared.OriginalHash = $CurrentHash
            $Prepared.Backup = Join-Path (Split-Path -Parent $Prepared.Target) `
                (".gcb-{0}-{1}.tmp" -f $TransactionId, $Prepared.Index)
            if (Test-Path -LiteralPath $Prepared.Backup) {
                throw "Operator file backup path is unexpectedly occupied: $($Prepared.Backup)"
            }
            Move-Item -LiteralPath $Prepared.Target -Destination $Prepared.Backup
            $Prepared.BackedUp = $true
        }

        foreach ($Prepared in $PreparedAssets) {
            if (Test-Path -LiteralPath $Prepared.Target) {
                throw "Operator file target became occupied during publication: $($Prepared.Target)"
            }
            Move-Item -LiteralPath $Prepared.Stage -Destination $Prepared.Target
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
        for ($Index = $TransactionItems.Count - 1; $Index -ge 0; $Index--) {
            $Prepared = $TransactionItems[$Index]
            if ($Prepared.Published -and (Test-Path -LiteralPath $Prepared.Target)) {
                try {
                    Remove-Item -LiteralPath $Prepared.Target -Force
                }
                catch {
                    $RollbackErrors += "could not remove $($Prepared.Target): $($_.Exception.Message)"
                }
            }
            if ($Prepared.BackedUp) {
                try {
                    if (-not (Test-Path -LiteralPath $Prepared.Backup -PathType Leaf)) {
                        throw "verified backup is missing"
                    }
                    if (Test-Path -LiteralPath $Prepared.Target) {
                        throw "restore target is occupied"
                    }
                    Move-Item -LiteralPath $Prepared.Backup -Destination $Prepared.Target
                    $RestoredHash = (Get-FileHash -LiteralPath $Prepared.Target -Algorithm SHA256).Hash
                    if ($Prepared.OriginalHash -ne $RestoredHash) {
                        throw "restored file hash does not match the pre-refresh file"
                    }
                    $Prepared.BackedUp = $false
                }
                catch {
                    $RollbackErrors += "could not restore $($Prepared.Target): $($_.Exception.Message)"
                }
            }
        }
        if ($RollbackErrors.Count -gt 0) {
            throw "Operator file refresh failed and rollback was incomplete: $($RollbackErrors -join '; '). Original error: $($PublishFailure.Exception.Message)"
        }
        throw "Operator file refresh failed; the prior operator asset set was restored. Original error: $($PublishFailure.Exception.Message)"
    }

    foreach ($Prepared in $TransactionItems) {
        if ($Prepared.BackedUp) {
            try {
                if (-not (Test-Path -LiteralPath $Prepared.Backup -PathType Leaf)) {
                    throw "verified backup is missing"
                }
                Remove-Item -LiteralPath $Prepared.Backup -Force
                $Prepared.BackedUp = $false
            }
            catch {
                Write-Warning "Operator assets are current, but a prior backup could not be removed: $($Prepared.Backup)"
            }
        }
    }
    foreach ($Prepared in $PreparedAssets) {
        Write-Host "Operator file verified: $($Prepared.Target)"
    }
    foreach ($Prepared in $PreparedRetirements) {
        Write-Host "Retired obsolete operator file: $($Prepared.Target)"
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
