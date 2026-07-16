[CmdletBinding()]
param(
    [string]$OperationsRoot = "",
    [string]$DropboxRoot = "",
    [ValidateRange(1, 720)]
    [int]$MaxMicrosAgeHours = 72,
    [ValidateRange(1, 100000)]
    [int]$MaxPlaceholderFiles = 10000,
    [string]$ReportPath = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$ProgramRoot = Split-Path -Parent $PSCommandPath
$ProjectRoot = Split-Path -Parent $ProgramRoot
if ([string]::IsNullOrWhiteSpace($OperationsRoot)) {
    if ((Split-Path -Leaf $ProjectRoot) -eq "Gift Card Reconciliation Automation") {
        $OperationsRoot = Split-Path -Parent $ProjectRoot
    }
    else {
        $OperationsRoot = $ProjectRoot
    }
}
$OperationsRoot = [IO.Path]::GetFullPath($OperationsRoot).TrimEnd('\', '/')
if ([string]::IsNullOrWhiteSpace($DropboxRoot)) {
    $DropboxRoot = Split-Path -Parent $OperationsRoot
}
$DropboxRoot = [IO.Path]::GetFullPath($DropboxRoot).TrimEnd('\', '/')

$results = [Collections.Generic.List[object]]::new()

function Add-HealthResult {
    param(
        [Parameter(Mandatory = $true)][ValidateSet("PASS", "INFO", "WARNING", "BLOCKER")][string]$Status,
        [Parameter(Mandatory = $true)][string]$Control,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $script:results.Add([pscustomobject]@{
        status = $Status
        control = $Control
        message = $Message
    })
}

function Test-ChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$Child
    )

    $parentFull = [IO.Path]::GetFullPath($Parent).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $childFull = [IO.Path]::GetFullPath($Child)
    return $childFull.StartsWith($parentFull, [StringComparison]::OrdinalIgnoreCase)
}

function Get-ManifestTreeHash {
    param([Parameter(Mandatory = $true)][object[]]$Files)

    $rows = foreach ($file in $Files | Sort-Object path) {
        "$($file.path)`t$($file.sha256)`t$($file.bytes)"
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($rows -join "`n"))
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") })
    }
    finally {
        $sha.Dispose()
    }
}

function Test-DeploymentManifest {
    $manifestPath = Join-Path $ProjectRoot "deployment-manifest.json"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        Add-HealthResult -Status BLOCKER -Control "Deployment manifest" -Message "Missing deployment-manifest.json in $ProjectRoot. Deploy from the clean local Git checkout."
        return
    }

    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if ($manifest.project -ne "gift_card_recon_codex_complete" -or [int]$manifest.schema_version -ne 1) {
            throw "unsupported project or schema"
        }
        $manifestFiles = @($manifest.files)
        if ($manifestFiles.Count -ne [int]$manifest.file_count -or $manifestFiles.Count -eq 0) {
            throw "file count does not match the manifest payload"
        }

        $failures = [Collections.Generic.List[string]]::new()
        foreach ($file in $manifestFiles) {
            $relative = [string]$file.path
            if ([IO.Path]::IsPathRooted($relative) -or $relative -match "(^|[\\/])\.\.([\\/]|$)") {
                $failures.Add("unsafe path $relative")
                continue
            }
            $target = [IO.Path]::GetFullPath((Join-Path $ProjectRoot $relative.Replace('/', [IO.Path]::DirectorySeparatorChar)))
            if (-not (Test-ChildPath -Parent $ProjectRoot -Child $target)) {
                $failures.Add("path outside program root $relative")
                continue
            }
            if (-not (Test-Path -LiteralPath $target -PathType Leaf)) {
                $failures.Add("missing $relative")
                continue
            }
            $item = Get-Item -LiteralPath $target
            $actualHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actualHash -ne ([string]$file.sha256).ToLowerInvariant() -or $item.Length -ne [long]$file.bytes) {
                $failures.Add("changed $relative")
            }
            if ($failures.Count -ge 10) {
                break
            }
        }

        $treeHash = Get-ManifestTreeHash -Files $manifestFiles
        if ($treeHash -ne ([string]$manifest.source_tree_sha256).ToLowerInvariant()) {
            $failures.Add("manifest tree hash mismatch")
        }
        $expectedPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($file in $manifestFiles) {
            [void]$expectedPaths.Add(([string]$file.path).Replace("\", "/"))
        }
        [void]$expectedPaths.Add("deployment-manifest.json")
        $extraFiles = foreach ($item in Get-ChildItem -LiteralPath $ProjectRoot -File -Recurse -Force) {
            $relative = $item.FullName.Substring($ProjectRoot.Length).TrimStart([char[]]"\/").Replace("\", "/")
            if (-not $expectedPaths.Contains($relative)) {
                $relative
            }
        }
        if (@($extraFiles).Count -gt 0) {
            $failures.Add("unexpected deployed files: $(@($extraFiles | Select-Object -First 10) -join ', ')")
        }
        if ($failures.Count -gt 0) {
            Add-HealthResult -Status BLOCKER -Control "Deployment manifest" -Message ("Deployment verification failed: " + ($failures -join "; "))
        }
        else {
            Add-HealthResult -Status PASS -Control "Deployment manifest" -Message "Commit $($manifest.commit) verified across $($manifest.file_count) deployed files."
        }
    }
    catch {
        Add-HealthResult -Status BLOCKER -Control "Deployment manifest" -Message "Manifest could not be validated: $($_.Exception.Message)"
    }

    if (Test-Path -LiteralPath (Join-Path $ProjectRoot ".git")) {
        Add-HealthResult -Status BLOCKER -Control "Dropbox Git metadata" -Message "The Dropbox program folder still contains .git metadata. Deploy the source-only operator snapshot."
    }
    else {
        Add-HealthResult -Status PASS -Control "Dropbox Git metadata" -Message "No .git folder is present in the operator deployment."
    }
    if (Test-Path -LiteralPath (Join-Path $OperationsRoot ".git")) {
        Add-HealthResult -Status BLOCKER -Control "Operations-root Git metadata" -Message "The Dropbox operations root contains orphaned .git metadata. Remove it after verifying the local Git clone and GitHub remote."
    }
    else {
        Add-HealthResult -Status PASS -Control "Operations-root Git metadata" -Message "No orphaned .git folder is present at the Dropbox operations root."
    }
}

function Test-OperatorAssets {
    $assets = @(
        [pscustomobject]@{ Template = "templates\00 START HERE - Gift Card Reconciliation.txt"; Target = "00 START HERE - Gift Card Reconciliation.txt" },
        [pscustomobject]@{ Template = "templates\Run Weekly Gift Card Reconciliation.cmd"; Target = "Run Weekly Gift Card Reconciliation.cmd" },
        [pscustomobject]@{ Template = "templates\Run Monthly Gift Card Close.cmd"; Target = "Run Monthly Gift Card Close.cmd" },
        [pscustomobject]@{ Template = "templates\Check Gift Card Reconciliation Health.cmd"; Target = "Check Gift Card Reconciliation Health.cmd" },
        [pscustomobject]@{ Template = "templates\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"; Target = "01 Weekly Gift Card Activity Reports\9354 Richmond\activity\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt" },
        [pscustomobject]@{ Template = "templates\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"; Target = "01 Weekly Gift Card Activity Reports\9355 Virginia Beach\activity\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt" },
        [pscustomobject]@{ Template = "templates\00 DROP DARDEN CREDIT MEMOS HERE.txt"; Target = "02 Monthly Close Inputs\Darden Reports - Drop Here\00 DROP DARDEN CREDIT MEMOS HERE.txt" }
    )
    $failures = [Collections.Generic.List[string]]::new()
    foreach ($asset in $assets) {
        $templatePath = Join-Path $ProjectRoot $asset.Template
        $targetPath = Join-Path $OperationsRoot $asset.Target
        if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) {
            $failures.Add("missing deployed template $($asset.Template)")
            continue
        }
        if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
            $failures.Add("missing operator asset $($asset.Target)")
            continue
        }
        try {
            if ((Get-FileHash -LiteralPath $templatePath -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash) {
                $failures.Add("outdated operator asset $($asset.Target)")
            }
        }
        catch {
            $failures.Add("unreadable operator asset $($asset.Target)")
        }
    }
    if ($failures.Count -gt 0) {
        Add-HealthResult -Status BLOCKER -Control "Operator launchers and guides" -Message ($failures -join "; ")
    }
    else {
        Add-HealthResult -Status PASS -Control "Operator launchers and guides" -Message "$($assets.Count) installed operator assets match their deployed templates."
    }
}

function Test-OperatorRuntime {
    try {
        . (Join-Path $ProgramRoot "runtime.ps1")
        $runtime = Get-GiftCardReconRuntime -Profile Operator
        $expected = Get-GiftCardReconDependencyFingerprint -ProgramRoot $ProgramRoot -Profile Operator
        $installed = ""
        if (Test-Path -LiteralPath $runtime.DependencyFingerprintPath -PathType Leaf) {
            $installed = (Get-Content -LiteralPath $runtime.DependencyFingerprintPath -Raw).Trim()
        }
        if (-not (Test-GiftCardReconRuntime -Runtime $runtime)) {
            Add-HealthResult -Status BLOCKER -Control "Operator runtime" -Message "Runtime validation failed at $($runtime.RuntimeRoot). Run the deployed _program\install.ps1."
        }
        elseif ($installed -ne $expected) {
            Add-HealthResult -Status BLOCKER -Control "Operator runtime" -Message "Runtime dependencies or deployed source changed. Run the deployed _program\install.ps1."
        }
        else {
            Add-HealthResult -Status PASS -Control "Operator runtime" -Message "Stable operator runtime is ready at $($runtime.RuntimeRoot)."
        }
    }
    catch {
        Add-HealthResult -Status BLOCKER -Control "Operator runtime" -Message $_.Exception.Message
    }
}

function Test-ExcelAvailability {
    try {
        $excelType = [type]::GetTypeFromProgID("Excel.Application")
        if ($null -eq $excelType) {
            throw "Excel.Application is not registered."
        }
        Add-HealthResult -Status PASS -Control "Microsoft Excel" -Message "Excel desktop automation is registered."
    }
    catch {
        Add-HealthResult -Status BLOCKER -Control "Microsoft Excel" -Message "Excel desktop automation is unavailable: $($_.Exception.Message)"
    }
}

function Test-MicrosEvidence {
    $sources = @(
        [pscustomobject]@{ Store = "9354 Richmond"; Folder = Join-Path $DropboxRoot "micros_data\RC-Richmond-current" },
        [pscustomobject]@{ Store = "9355 Virginia Beach"; Folder = Join-Path $DropboxRoot "GETLinkedData-VB" }
    )
    $now = Get-Date
    foreach ($source in $sources) {
        foreach ($name in @("DLYSYSTT.TXT", "TENDER_DETAIL.TXT")) {
            $path = Join-Path $source.Folder $name
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                Add-HealthResult -Status BLOCKER -Control "Micros $($source.Store)" -Message "Missing $path"
                continue
            }
            try {
                $item = Get-Item -LiteralPath $path
                if ($item.Length -le 0) {
                    throw "$name is empty"
                }
                $ageHours = ($now - $item.LastWriteTime).TotalHours
                if ($ageHours -lt -1) {
                    Add-HealthResult -Status BLOCKER -Control "Micros $($source.Store)" -Message "$name has a future timestamp: $($item.LastWriteTime.ToString('s'))."
                }
                elseif ($ageHours -gt $MaxMicrosAgeHours) {
                    Add-HealthResult -Status BLOCKER -Control "Micros $($source.Store)" -Message "$name is $([math]::Round($ageHours, 1)) hours old; maximum is $MaxMicrosAgeHours hours."
                }
                else {
                    Add-HealthResult -Status PASS -Control "Micros $($source.Store)" -Message "$name is present, nonempty, and $([math]::Round($ageHours, 1)) hours old."
                }
            }
            catch {
                Add-HealthResult -Status BLOCKER -Control "Micros $($source.Store)" -Message "$path could not be inspected: $($_.Exception.Message)"
            }
        }
    }
}

function Test-OutputFolder {
    $outputRoot = Join-Path $OperationsRoot "03 Finished Reports"
    if (-not (Test-Path -LiteralPath $outputRoot -PathType Container)) {
        Add-HealthResult -Status BLOCKER -Control "Output folder" -Message "Missing output folder: $outputRoot"
        return
    }
    $probe = Join-Path $outputRoot ".gift-card-health-$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($probe, "health-check", [Text.Encoding]::ASCII)
        if ((Get-Content -LiteralPath $probe -Raw) -ne "health-check") {
            throw "probe read-back did not match"
        }
        Add-HealthResult -Status PASS -Control "Output folder" -Message "Output folder is writable: $outputRoot"
    }
    catch {
        Add-HealthResult -Status BLOCKER -Control "Output folder" -Message "Output folder write test failed: $($_.Exception.Message)"
    }
    finally {
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    }
}

function Test-DropboxPlaceholders {
    if (-not (Test-Path -LiteralPath $OperationsRoot -PathType Container)) {
        Add-HealthResult -Status BLOCKER -Control "Dropbox placeholders" -Message "Operations root is missing: $OperationsRoot"
        return
    }

    if (-not ("GiftCardRecon.NativeFileAttributes" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace GiftCardRecon {
    public static class NativeFileAttributes {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetFileAttributes(string path);
    }
}
"@
    }

    $scanned = 0
    $conflicts = [Collections.Generic.List[string]]::new()
    try {
        foreach ($item in Get-ChildItem -LiteralPath $OperationsRoot -File -Recurse -Force -ErrorAction Stop) {
            $scanned += 1
            if ($scanned -gt $MaxPlaceholderFiles) {
                Add-HealthResult -Status BLOCKER -Control "Dropbox placeholders" -Message "Placeholder scan exceeded the safety limit of $MaxPlaceholderFiles files. Narrow or clean the operations folder."
                return
            }
            $attributes = [GiftCardRecon.NativeFileAttributes]::GetFileAttributes($item.FullName)
            if ($attributes -eq [uint32]::MaxValue) {
                $conflicts.Add("unreadable: $($item.FullName)")
            }
            elseif (($attributes -band 0x80000) -ne 0 -and ($attributes -band 0x100000) -ne 0) {
                $conflicts.Add("pinned and unpinned: $($item.FullName)")
            }
            if ($conflicts.Count -ge 20) {
                break
            }
        }
    }
    catch {
        Add-HealthResult -Status BLOCKER -Control "Dropbox placeholders" -Message "Placeholder scan failed: $($_.Exception.Message)"
        return
    }

    if ($conflicts.Count -gt 0) {
        Add-HealthResult -Status BLOCKER -Control "Dropbox placeholders" -Message ("Invalid or unreadable Dropbox files found: " + ($conflicts -join "; "))
    }
    else {
        Add-HealthResult -Status PASS -Control "Dropbox placeholders" -Message "$scanned files scanned; no conflicting pinned/unpinned states found."
    }
}

function Add-PendingInputSummary {
    $dardenRoot = Join-Path $OperationsRoot "02 Monthly Close Inputs\Darden Reports - Drop Here"
    $dardenCount = if (Test-Path -LiteralPath $dardenRoot) {
        @(Get-ChildItem -LiteralPath $dardenRoot -File -Filter "*.pdf" -ErrorAction SilentlyContinue).Count
    }
    else { 0 }
    $weeklyRoot = Join-Path $OperationsRoot "01 Weekly Gift Card Activity Reports"
    $weeklyCount = if (Test-Path -LiteralPath $weeklyRoot) {
        @(Get-ChildItem -LiteralPath $weeklyRoot -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike "00 *" -and $_.Extension -in @(".xls", ".xlsx", ".xlsm") }).Count
    }
    else { 0 }
    Add-HealthResult -Status INFO -Control "Pending inputs" -Message "$weeklyCount weekly Activity workbook(s) and $dardenCount Darden PDF(s) are waiting."
}

Write-Host ""
Write-Host "GIFT CARD RECONCILIATION - OPERATOR HEALTH CHECK" -ForegroundColor Cyan
Write-Host "Operations root: $OperationsRoot"
Write-Host ""

Test-DeploymentManifest
Test-OperatorAssets
Test-OperatorRuntime
Test-ExcelAvailability
Test-MicrosEvidence
Test-OutputFolder
Test-DropboxPlaceholders
Add-PendingInputSummary

foreach ($result in $results) {
    $color = switch ($result.status) {
        "PASS" { "Green" }
        "INFO" { "Cyan" }
        "WARNING" { "Yellow" }
        default { "Red" }
    }
    Write-Host ("[{0}] {1}: {2}" -f $result.status, $result.control, $result.message) -ForegroundColor $color
}

$blockers = @($results | Where-Object { $_.status -eq "BLOCKER" })
$report = [ordered]@{
    generated_at_utc = [DateTime]::UtcNow.ToString("o")
    operations_root = $OperationsRoot
    program_root = $ProjectRoot
    overall_status = if ($blockers.Count -eq 0) { "READY" } else { "BLOCKED" }
    blocker_count = $blockers.Count
    controls = @($results)
}
if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $resolvedReport = if ([IO.Path]::IsPathRooted($ReportPath)) {
        [IO.Path]::GetFullPath($ReportPath)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $OperationsRoot $ReportPath))
    }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resolvedReport) | Out-Null
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resolvedReport -Encoding UTF8
    Write-Host "Health report: $resolvedReport"
}

Write-Host ""
if ($blockers.Count -eq 0) {
    Write-Host "READY: all required operator controls passed." -ForegroundColor Green
    exit 0
}
Write-Host "BLOCKED: $($blockers.Count) required control(s) need attention." -ForegroundColor Red
exit 2
