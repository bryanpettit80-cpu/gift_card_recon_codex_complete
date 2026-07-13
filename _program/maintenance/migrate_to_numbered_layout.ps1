#Requires -Version 5.1

<#
.SYNOPSIS
Migrates Gift Card Reconciliation business files into the numbered Dropbox layout.

.DESCRIPTION
The default mode is a read-only dry run. The migration inventories only known
business-data roots, calculates SHA-256 for every file, and maps each file to
its numbered-layout destination. Program files, Git metadata, launchers,
caches, and temporary files are never migration inputs.

Apply mode performs a complete preflight before changing anything. Each file
is copied to a same-directory partial file, hash-verified, atomically published,
and only then is the source quarantined, reverified, and removed. A preflight
manifest is written before the first change and a checkpointed post manifest
records progress. Re-running Apply is safe: matching destinations are verified
and remaining duplicate sources are removed; conflicting destinations block
the whole preflight.

Verify validates a prior manifest without changing files. Rollback restores
the old source paths from a post manifest, preserving destinations that existed
before migration and removing only destinations created by that migration.

.EXAMPLE
.\migrate_to_numbered_layout.ps1 -OperationsRoot "C:\Users\bryan\Dropbox\Gift Card Reconciliation"

.EXAMPLE
.\migrate_to_numbered_layout.ps1 -OperationsRoot "C:\Users\bryan\Dropbox\Gift Card Reconciliation" -Apply -ExpectedPlanSha256 <hash>

.EXAMPLE
.\migrate_to_numbered_layout.ps1 -OperationsRoot <path> -Verify -ManifestPath <post-manifest.json>

.EXAMPLE
.\migrate_to_numbered_layout.ps1 -OperationsRoot <path> -Rollback -ManifestPath <post-manifest.json>
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OperationsRoot,

    [Parameter()]
    [switch]$Apply,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [switch]$Verify,

    [Parameter()]
    [switch]$Rollback,

    [Parameter()]
    [string]$ManifestPath,

    [Parameter()]
    [string]$ManifestDirectory,

    [Parameter()]
    [string]$ExpectedPlanSha256
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:QuarantineSuffix = ".gc-layout-source-quarantine"
$script:PartialMarker = ".gc-layout-"

if ($null -eq ("GiftCardRecon.LayoutReparsePoint" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace GiftCardRecon
{
    public static class LayoutReparsePoint
    {
        private const uint FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000;
        private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
        private const uint OPEN_EXISTING = 3;
        private const uint FSCTL_GET_REPARSE_POINT = 0x000900A8;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string fileName, uint desiredAccess, FileShare shareMode,
            IntPtr securityAttributes, uint creationDisposition,
            uint flagsAndAttributes, IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(
            SafeFileHandle device, uint controlCode, IntPtr inputBuffer,
            int inputBufferSize, byte[] outputBuffer, int outputBufferSize,
            out int bytesReturned, IntPtr overlapped);

        public static uint GetTag(string path)
        {
            using (SafeFileHandle handle = CreateFile(
                path, 0, FileShare.Read | FileShare.Write | FileShare.Delete,
                IntPtr.Zero, OPEN_EXISTING,
                FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                IntPtr.Zero))
            {
                if (handle.IsInvalid)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not open reparse point: " + path);
                byte[] buffer = new byte[16 * 1024];
                int bytesReturned;
                if (!DeviceIoControl(handle, FSCTL_GET_REPARSE_POINT, IntPtr.Zero,
                    0, buffer, buffer.Length, out bytesReturned, IntPtr.Zero))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not read reparse tag: " + path);
                if (bytesReturned < 8)
                    throw new InvalidDataException("Reparse buffer was too short for: " + path);
                return BitConverter.ToUInt32(buffer, 0);
            }
        }

        public static bool IsCloudProjection(uint tag)
        {
            return (tag & 0xFFFF0FFFu) == 0x9000001Au;
        }
    }
}
"@
}

$selectedModes = @(@($Apply, $DryRun, $Verify, $Rollback) | Where-Object { $_ })
if ($selectedModes.Count -gt 1) {
    throw "Choose only one mode: -DryRun, -Apply, -Verify, or -Rollback."
}
if (-not $Apply -and -not $Verify -and -not $Rollback) {
    $DryRun = $true
}
if (($Verify -or $Rollback) -and [string]::IsNullOrWhiteSpace($ManifestPath)) {
    throw "-ManifestPath is required with -Verify or -Rollback."
}
if ($Apply -and [string]::IsNullOrWhiteSpace($ExpectedPlanSha256)) {
    throw "-Apply requires -ExpectedPlanSha256 from a reviewed dry run."
}
if ($DryRun -and -not [string]::IsNullOrWhiteSpace($ManifestPath)) {
    throw "-ManifestPath is only used by -Verify or -Rollback. Dry run prints its complete manifest without writing files."
}

function Get-FullPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [System.IO.Path]::GetFullPath($Path)
}

function Assert-WithinRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )
    $full = Get-FullPath -Path $Path
    $rootFull = Get-FullPath -Path $Root
    $prefix = $rootFull.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $full.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase) -and
        -not $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escapes the approved operations root. Path='$full'; Root='$rootFull'."
    }
    $operationsRootVariable = Get-Variable -Name OperationsRootResolved -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $operationsRootVariable) {
        Assert-NoUnsafeReparsePoint -Path $full -Root $script:OperationsRootResolved
    }
    return $full
}

function Assert-NoUnsafeReparsePoint {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )
    $candidate = Get-FullPath -Path $Path
    $rootFull = Get-FullPath -Path $Root
    $rootPrefix = $rootFull.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $candidate.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase) -and
        -not $candidate.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escapes the approved operations root. Path='$candidate'; Root='$rootFull'."
    }

    $probe = $candidate
    while (-not (Test-Path -LiteralPath $probe)) {
        if ($probe.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase)) { break }
        $next = [System.IO.Path]::GetDirectoryName($probe)
        if ([string]::IsNullOrWhiteSpace($next) -or $next -eq $probe) {
            throw "Could not resolve an existing ancestor inside the operations root: $candidate"
        }
        $probe = $next
    }
    while ($true) {
        if (Test-Path -LiteralPath $probe) {
            $item = Get-Item -LiteralPath $probe -Force
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $tag = [GiftCardRecon.LayoutReparsePoint]::GetTag($item.FullName)
                if (-not [GiftCardRecon.LayoutReparsePoint]::IsCloudProjection($tag)) {
                    throw ("Unsafe filesystem link blocked at '{0}' (reparse tag 0x{1})." -f $item.FullName, $tag.ToString("x8"))
                }
            }
        }
        if ($probe.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase)) { break }
        $next = [System.IO.Path]::GetDirectoryName($probe)
        if ([string]::IsNullOrWhiteSpace($next) -or $next -eq $probe) {
            throw "Existing path chain escaped the approved operations root: $candidate"
        }
        $probe = $next
    }
}

function Get-RelativePathLiteral {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $rootFull = Get-FullPath -Path $Root
    $pathFull = Assert-WithinRoot -Path $Path -Root $rootFull
    if ($pathFull.Equals($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }
    return $pathFull.Substring($rootFull.TrimEnd('\', '/').Length).TrimStart('\', '/')
}

function Join-RootPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )
    return Assert-WithinRoot -Path (Join-Path $Root $RelativePath) -Root $Root
}

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file is missing: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-BytesSha256 {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        return -join ($hasher.ComputeHash($Bytes) | ForEach-Object { $_.ToString("x2") })
    }
    finally {
        $hasher.Dispose()
    }
}

function Get-StringSha256 {
    param([Parameter(Mandatory = $true)][string]$Value)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $script:Utf8NoBom.GetBytes($Value)
        $hash = $sha.ComputeHash($bytes)
        return ([BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Assert-Fingerprint {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][long]$SizeBytes,
        [Parameter(Mandatory = $true)][string]$Sha256
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file is missing: $Path"
    }
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.Length -ne $SizeBytes) {
        throw "Size mismatch for '$Path'. Expected $SizeBytes; found $($item.Length)."
    }
    $actual = Get-Sha256 -Path $Path
    if ($actual -ne $Sha256.ToLowerInvariant()) {
        throw "SHA-256 mismatch for '$Path'. Expected $Sha256; found $actual."
    }
}

function Assert-StableReadable {
    param([Parameter(Mandatory = $true)][string]$Path)
    $stream = $null
    try {
        # Existing writers and delete-capable handles block this open. Other
        # readers remain allowed, which is compatible with Dropbox indexing.
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read
        )
    }
    catch {
        throw "File is locked or not stably readable: '$Path'. $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    [void](Assert-WithinRoot -Path $Path -Root $script:OperationsRootResolved)
    if (Test-Path -LiteralPath $Path) {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            throw "A file occupies a required directory path: $Path"
        }
        return
    }
    [void][System.IO.Directory]::CreateDirectory($Path)
}

function Write-JsonAtomic {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )
    $full = Assert-WithinRoot -Path $Path -Root $script:OperationsRootResolved
    $parent = [System.IO.Path]::GetDirectoryName($full)
    Ensure-Directory -Path $parent
    $temp = Join-Path $parent (".{0}.{1}.tmp" -f [System.IO.Path]::GetFileName($full), [Guid]::NewGuid().ToString("N"))
    $backup = Join-Path $parent (".{0}.{1}.bak" -f [System.IO.Path]::GetFileName($full), [Guid]::NewGuid().ToString("N"))
    $json = $Value | ConvertTo-Json -Depth 30
    try {
        [System.IO.File]::WriteAllText($temp, $json + [Environment]::NewLine, $script:Utf8NoBom)
        if (Test-Path -LiteralPath $full -PathType Leaf) {
            [System.IO.File]::Replace($temp, $full, $backup, $true)
            if (Test-Path -LiteralPath $backup -PathType Leaf) {
                Remove-Item -LiteralPath $backup -Force
            }
        }
        else {
            [System.IO.File]::Move($temp, $full)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temp -PathType Leaf) { Remove-Item -LiteralPath $temp -Force }
        if (Test-Path -LiteralPath $backup -PathType Leaf) {
            if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
                [System.IO.File]::Move($backup, $full)
            }
            else {
                Remove-Item -LiteralPath $backup -Force
            }
        }
    }
}

function Test-ProtectedProgramPath {
    param([Parameter(Mandatory = $true)][string]$RelativePath)
    $parts = @($RelativePath -split '[\\/]')
    $excludedDirectories = @(
        ".git", ".venv", "venv", "node_modules", "__pycache__", ".pytest_cache",
        "build", "dist", "htmlcov", "Gift Card Reconciliation Automation", "_program"
    )
    foreach ($part in $parts) {
        if ($part -in $excludedDirectories -or $part.EndsWith(".egg-info", [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Test-ExcludedBusinessFile {
    param([Parameter(Mandatory = $true)][string]$RelativePath)
    $parts = @($RelativePath -split '[\\/]')
    if (Test-ProtectedProgramPath -RelativePath $RelativePath) { return $true }
    $name = $parts[-1]
    if ($name -in @(".gitkeep", ".DS_Store", "Thumbs.db", "desktop.ini")) { return $true }
    if ($name.EndsWith(".pyc", [StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($name.Contains($script:PartialMarker) -and $name.EndsWith(".partial", [StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $false
}

function Resolve-DestinationRelative {
    param([Parameter(Mandatory = $true)][string]$SourceRelative)
    $path = $SourceRelative.Replace('/', '\')

    if ($path.StartsWith("Archive - Old Files\", [StringComparison]::OrdinalIgnoreCase)) {
        return "04 Archive\" + $path.Substring("Archive - Old Files\".Length)
    }

    foreach ($store in @(
        [pscustomobject]@{ Old = "9354 - Weekly"; New = "9354 Richmond"; Id = "9354" },
        [pscustomobject]@{ Old = "9355 - Weekly"; New = "9355 Virginia Beach"; Id = "9355" }
    )) {
        $activityPrefix = "$($store.Old)\activity\"
        if ($path.StartsWith($activityPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            return "01 Weekly Gift Card Activity Reports\$($store.New)\activity\" + $path.Substring($activityPrefix.Length)
        }
        if ($path.Equals("$($store.Old)\pos_controls.csv", [StringComparison]::OrdinalIgnoreCase)) {
            return "04 Archive\Legacy Reconciliation\Manual POS Controls\$($store.Id)\pos_controls.csv"
        }
        $weeklyPrefix = "$($store.Old)\"
        if ($path.StartsWith($weeklyPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            return "04 Archive\Legacy Reconciliation\Weekly Input\$($store.Id)\" + $path.Substring($weeklyPrefix.Length)
        }
    }

    if ($path.StartsWith("Monthly Close\Darden Reports - Drop Here\", [StringComparison]::OrdinalIgnoreCase)) {
        return "02 Monthly Close Inputs\Darden Reports - Drop Here\" + $path.Substring("Monthly Close\Darden Reports - Drop Here\".Length)
    }
    if ($path.StartsWith("Monthly Close\9354\", [StringComparison]::OrdinalIgnoreCase)) {
        return "02 Monthly Close Inputs\9354 Richmond\" + $path.Substring("Monthly Close\9354\".Length)
    }
    if ($path.StartsWith("Monthly Close\9355\", [StringComparison]::OrdinalIgnoreCase)) {
        return "02 Monthly Close Inputs\9355 Virginia Beach\" + $path.Substring("Monthly Close\9355\".Length)
    }
    if ($path.StartsWith("Monthly Close\", [StringComparison]::OrdinalIgnoreCase)) {
        return "02 Monthly Close Inputs\" + $path.Substring("Monthly Close\".Length)
    }

    if ($path.StartsWith("Output\Weekly\9354\", [StringComparison]::OrdinalIgnoreCase)) {
        return "03 Finished Reports\Weekly\9354 Richmond\" + $path.Substring("Output\Weekly\9354\".Length)
    }
    if ($path.StartsWith("Output\Weekly\9355\", [StringComparison]::OrdinalIgnoreCase)) {
        return "03 Finished Reports\Weekly\9355 Virginia Beach\" + $path.Substring("Output\Weekly\9355\".Length)
    }
    if ($path.StartsWith("Output\", [StringComparison]::OrdinalIgnoreCase)) {
        return "03 Finished Reports\" + $path.Substring("Output\".Length)
    }

    if ($path.StartsWith("input\", [StringComparison]::OrdinalIgnoreCase)) {
        return "04 Archive\Legacy Reconciliation\input\" + $path.Substring("input\".Length)
    }
    if ($path.StartsWith("reports\", [StringComparison]::OrdinalIgnoreCase)) {
        return "04 Archive\Generated Reports\Legacy Reports\" + $path.Substring("reports\".Length)
    }

    throw "No approved business-data mapping exists for: $SourceRelative"
}

function Get-RequiredDirectoryRelatives {
    return @(
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
        "_automation_runs\test-output",
        "_automation_runs\migration"
    )
}

function Get-LegacySourceRootRelatives {
    return @("9354 - Weekly", "9355 - Weekly", "Monthly Close", "Output", "Archive - Old Files", "input", "reports")
}

function Get-TargetBusinessRootRelatives {
    return @("01 Weekly Gift Card Activity Reports", "02 Monthly Close Inputs", "03 Finished Reports", "04 Archive")
}

function Get-OrphanPartials {
    $orphans = New-Object System.Collections.Generic.List[string]
    foreach ($targetRelative in Get-TargetBusinessRootRelatives) {
        $target = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $targetRelative
        if (-not (Test-Path -LiteralPath $target -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $target -File -Recurse -Force) {
            if ($file.Name.Contains($script:PartialMarker) -and $file.Name.EndsWith(".partial", [StringComparison]::OrdinalIgnoreCase)) {
                [void]$orphans.Add($file.FullName)
            }
        }
    }
    return @($orphans | ForEach-Object { $_ })
}

function New-Inventory {
    $raw = New-Object System.Collections.Generic.List[object]
    $destinationOwners = @{}
    $sourceRoots = Get-LegacySourceRootRelatives

    foreach ($sourceRootRelative in $sourceRoots) {
        $sourceRoot = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $sourceRootRelative
        if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) { continue }

        foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -Force) {
            $full = $file.FullName
            $logicalFull = $full
            $fromQuarantine = $false
            if ($file.Name.EndsWith($script:QuarantineSuffix, [StringComparison]::OrdinalIgnoreCase)) {
                $logicalFull = $full.Substring(0, $full.Length - $script:QuarantineSuffix.Length)
                $fromQuarantine = $true
                if (Test-Path -LiteralPath $logicalFull) {
                    throw "Both a source and its migration quarantine exist: $logicalFull"
                }
            }
            $sourceRelative = Get-RelativePathLiteral -Root $script:OperationsRootResolved -Path $logicalFull
            if (Test-ExcludedBusinessFile -RelativePath $sourceRelative) { continue }
            $destinationRelative = Resolve-DestinationRelative -SourceRelative $sourceRelative
            $destinationFull = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $destinationRelative
            $key = $destinationFull.ToLowerInvariant()
            if ($destinationOwners.ContainsKey($key)) {
                throw "Multiple legacy files map to the same destination: '$($destinationOwners[$key])' and '$sourceRelative'."
            }
            $destinationOwners[$key] = $sourceRelative

            Assert-StableReadable -Path $full
            $sha = Get-Sha256 -Path $full
            $destinationExists = Test-Path -LiteralPath $destinationFull
            $destinationPreexisting = Test-Path -LiteralPath $destinationFull -PathType Leaf
            $initialStatus = "planned_move"
            if ($destinationExists -and -not $destinationPreexisting) {
                $destinationPreexisting = $true
                $initialStatus = "conflict"
            }
            elseif ($destinationPreexisting) {
                Assert-StableReadable -Path $destinationFull
                $destItem = Get-Item -LiteralPath $destinationFull -Force
                $destHash = Get-Sha256 -Path $destinationFull
                if ($destItem.Length -ne $file.Length -or $destHash -ne $sha) {
                    $initialStatus = "conflict"
                }
                else {
                    $initialStatus = "planned_deduplicate"
                }
            }

            [void]$raw.Add([pscustomobject][ordered]@{
                id = $null
                action = "move_file"
                source_relative = $sourceRelative
                source_absolute = $logicalFull
                source_quarantine_absolute = if ($fromQuarantine) { $full } else { $logicalFull + $script:QuarantineSuffix }
                destination_relative = $destinationRelative
                destination_absolute = $destinationFull
                sha256 = $sha
                size_bytes = [long]$file.Length
                last_write_time_utc = $file.LastWriteTimeUtc.ToString("o")
                destination_preexisting = [bool]$destinationPreexisting
                status = $initialStatus
                note = if ($fromQuarantine) { "Resumable source quarantine was inventoried." } else { $null }
            })
        }
    }

    $plannedDestinations = @{}
    foreach ($entry in $raw) { $plannedDestinations[$entry.destination_absolute.ToLowerInvariant()] = $true }
    foreach ($targetRelative in Get-TargetBusinessRootRelatives) {
        $target = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $targetRelative
        if (-not (Test-Path -LiteralPath $target -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $target -File -Recurse -Force) {
            $destinationRelative = Get-RelativePathLiteral -Root $script:OperationsRootResolved -Path $file.FullName
            if (Test-ExcludedBusinessFile -RelativePath $destinationRelative) { continue }
            if ($plannedDestinations.ContainsKey($file.FullName.ToLowerInvariant())) { continue }
            Assert-StableReadable -Path $file.FullName
            [void]$raw.Add([pscustomobject][ordered]@{
                id = $null
                action = "verify_existing"
                source_relative = $null
                source_absolute = $null
                source_quarantine_absolute = $null
                destination_relative = $destinationRelative
                destination_absolute = $file.FullName
                sha256 = Get-Sha256 -Path $file.FullName
                size_bytes = [long]$file.Length
                last_write_time_utc = $file.LastWriteTimeUtc.ToString("o")
                destination_preexisting = $true
                status = "already_migrated"
                note = "File already resides under a numbered business-data root."
            })
        }
    }

    $sorted = @($raw | Sort-Object destination_relative, source_relative)
    for ($index = 0; $index -lt $sorted.Count; $index++) {
        $sorted[$index].id = "file-{0:D5}" -f ($index + 1)
    }
    return $sorted
}

function Get-PlanSha256 {
    param([Parameter(Mandatory = $true)][object[]]$Files)
    $lines = @($Files | ForEach-Object {
        @(
            $_.action,
            $(if ($null -eq $_.source_relative) { "" } else { $_.source_relative }),
            $_.destination_relative,
            $_.sha256,
            [string]$_.size_bytes,
            [string]$_.destination_preexisting
        ) -join "`t"
    })
    return Get-StringSha256 -Value (($lines -join "`n") + "`n")
}

function Get-ManifestSummary {
    param([Parameter(Mandatory = $true)][object[]]$Files)
    return [pscustomobject][ordered]@{
        total_files = $Files.Count
        planned_moves = @($Files | Where-Object { $_.status -eq "planned_move" }).Count
        planned_deduplications = @($Files | Where-Object { $_.status -eq "planned_deduplicate" }).Count
        already_migrated = @($Files | Where-Object { $_.status -eq "already_migrated" }).Count
        conflicts = @($Files | Where-Object { $_.status -eq "conflict" }).Count
        total_bytes = [long](($Files | Measure-Object -Property size_bytes -Sum).Sum)
    }
}

function New-Manifest {
    param(
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][object[]]$Files,
        [Parameter(Mandatory = $true)][string]$RunId
    )
    return [pscustomobject][ordered]@{
        schema_version = 1
        migration = "gift-card-numbered-dropbox-layout"
        mode = $Mode
        status = $Status
        run_id = $RunId
        generated_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
        completed_at_utc = $null
        operations_root = $script:OperationsRootResolved
        plan_sha256 = Get-PlanSha256 -Files $Files
        archive_policy = "Archive - Old Files is renamed to 04 Archive with every internal relative path and file byte preserved."
        required_directories = @(Get-RequiredDirectoryRelatives)
        summary = Get-ManifestSummary -Files $Files
        files = $Files
        removed_placeholders = @()
        retained_legacy_paths = @()
        error = $null
    }
}

function Copy-Verify-PublishAndRemoveSource {
    param([Parameter(Mandatory = $true)][pscustomobject]$File)
    $source = $File.source_absolute
    $quarantine = $File.source_quarantine_absolute
    $destination = $File.destination_absolute
    $effectiveSource = if (Test-Path -LiteralPath $quarantine -PathType Leaf) { $quarantine } else { $source }
    $destinationExisted = Test-Path -LiteralPath $destination -PathType Leaf

    Assert-Fingerprint -Path $effectiveSource -SizeBytes ([long]$File.size_bytes) -Sha256 $File.sha256
    Assert-StableReadable -Path $effectiveSource
    if ($destinationExisted) {
        Assert-Fingerprint -Path $destination -SizeBytes ([long]$File.size_bytes) -Sha256 $File.sha256
    }
    else {
        $destinationParent = [System.IO.Path]::GetDirectoryName($destination)
        Ensure-Directory -Path $destinationParent
        $partial = Join-Path $destinationParent (".{0}{1}{2}.partial" -f [System.IO.Path]::GetFileName($destination), $script:PartialMarker, [Guid]::NewGuid().ToString("N"))
        $sourceLock = $null
        try {
            $sourceLock = [System.IO.File]::Open(
                $effectiveSource,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::Read
            )
            $target = [System.IO.File]::Open($partial, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try { $sourceLock.CopyTo($target) }
            finally { $target.Dispose() }
            Assert-Fingerprint -Path $partial -SizeBytes ([long]$File.size_bytes) -Sha256 $File.sha256
            if (Test-Path -LiteralPath $destination) {
                throw "Destination appeared during publication: $destination"
            }
            [System.IO.File]::Move($partial, $destination)
            Assert-Fingerprint -Path $destination -SizeBytes ([long]$File.size_bytes) -Sha256 $File.sha256
        }
        finally {
            if ($null -ne $sourceLock) { $sourceLock.Dispose() }
            if (Test-Path -LiteralPath $partial -PathType Leaf) { Remove-Item -LiteralPath $partial -Force }
        }
    }

    if ($effectiveSource.Equals($source, [StringComparison]::OrdinalIgnoreCase)) {
        if (Test-Path -LiteralPath $quarantine) {
            throw "Source quarantine path is occupied: $quarantine"
        }
        try {
            [System.IO.File]::Move($source, $quarantine)
        }
        catch {
            throw "Destination was safely published, but the source could not be quarantined (possibly locked): '$source'. Re-run after resolving the lock. $($_.Exception.Message)"
        }
    }

    try {
        Assert-Fingerprint -Path $quarantine -SizeBytes ([long]$File.size_bytes) -Sha256 $File.sha256
        Assert-Fingerprint -Path $destination -SizeBytes ([long]$File.size_bytes) -Sha256 $File.sha256
        Remove-Item -LiteralPath $quarantine -Force
        if (Test-Path -LiteralPath $quarantine) {
            throw "Source quarantine could not be removed: $quarantine"
        }
    }
    catch {
        if ((Test-Path -LiteralPath $quarantine -PathType Leaf) -and -not (Test-Path -LiteralPath $source)) {
            [System.IO.File]::Move($quarantine, $source)
        }
        throw
    }

    $File.status = if ($destinationExisted) { "deduplicated" } else { "moved" }
    $File.note = "Destination and source quarantine were hash-verified before source removal."
}

function Remove-LegacyPlaceholdersAndEmptyDirectories {
    param([Parameter(Mandatory = $true)][pscustomobject]$Manifest)
    $removed = New-Object System.Collections.Generic.List[object]
    $retained = New-Object System.Collections.Generic.List[string]
    foreach ($relative in Get-LegacySourceRootRelatives) {
        $root = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $relative
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        foreach ($placeholder in Get-ChildItem -LiteralPath $root -File -Recurse -Force | Where-Object { $_.Name -eq ".gitkeep" }) {
            $placeholderRelative = Get-RelativePathLiteral -Root $script:OperationsRootResolved -Path $placeholder.FullName
            if (Test-ProtectedProgramPath -RelativePath $placeholderRelative) { continue }
            $hash = Get-Sha256 -Path $placeholder.FullName
            $bytes = [System.IO.File]::ReadAllBytes($placeholder.FullName)
            [void]$removed.Add([pscustomobject][ordered]@{
                path = $placeholderRelative
                sha256 = $hash
                size_bytes = [long]$placeholder.Length
                content_base64 = [Convert]::ToBase64String($bytes)
            })
            Remove-Item -LiteralPath $placeholder.FullName -Force
        }
        $directories = @(Get-ChildItem -LiteralPath $root -Directory -Recurse -Force | Sort-Object { $_.FullName.Length } -Descending)
        foreach ($directory in $directories) {
            $directoryRelative = Get-RelativePathLiteral -Root $script:OperationsRootResolved -Path $directory.FullName
            if (Test-ProtectedProgramPath -RelativePath $directoryRelative) { continue }
            if (@(Get-ChildItem -LiteralPath $directory.FullName -Force).Count -eq 0) {
                Remove-Item -LiteralPath $directory.FullName -Force
            }
        }
        if ((Test-Path -LiteralPath $root -PathType Container) -and @(Get-ChildItem -LiteralPath $root -Force).Count -eq 0) {
            Remove-Item -LiteralPath $root -Force
        }
        elseif (Test-Path -LiteralPath $root -PathType Container) {
            [void]$retained.Add($relative)
        }
    }
    $Manifest.removed_placeholders = @($removed | ForEach-Object { $_ })
    $Manifest.retained_legacy_paths = @($retained | ForEach-Object { $_ })
}

function Assert-ManifestFilesMigrated {
    param([Parameter(Mandatory = $true)][pscustomobject]$Manifest)
    foreach ($file in @($Manifest.files)) {
        Assert-Fingerprint -Path $file.destination_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
        if ($file.action -eq "move_file") {
            if (Test-Path -LiteralPath $file.source_absolute -PathType Leaf) {
                throw "Legacy source remains after migration: $($file.source_absolute)"
            }
            if (Test-Path -LiteralPath $file.source_quarantine_absolute -PathType Leaf) {
                throw "Source quarantine remains after migration: $($file.source_quarantine_absolute)"
            }
            if ($file.source_relative.StartsWith("Archive - Old Files\", [StringComparison]::OrdinalIgnoreCase)) {
                $expected = "04 Archive\" + $file.source_relative.Substring("Archive - Old Files\".Length)
                if (-not $file.destination_relative.Equals($expected, [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Archive internal path was not preserved for: $($file.source_relative)"
                }
            }
        }
    }
}

function Assert-NoUnplannedLegacyBusinessFiles {
    foreach ($sourceRootRelative in Get-LegacySourceRootRelatives) {
        $sourceRoot = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $sourceRootRelative
        if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -Force) {
            $relative = Get-RelativePathLiteral -Root $script:OperationsRootResolved -Path $file.FullName
            if (Test-ExcludedBusinessFile -RelativePath $relative) { continue }
            throw (
                "A new or unplanned legacy business file appeared during migration: $relative. " +
                "The migration is blocked; run a new dry run and review its fingerprint before resuming."
            )
        }
    }
}

function Assert-ManifestIntegrity {
    param([Parameter(Mandatory = $true)][pscustomobject]$Manifest)
    if ($Manifest.schema_version -ne 1 -or $Manifest.migration -ne "gift-card-numbered-dropbox-layout") {
        throw "Unsupported numbered-layout migration manifest."
    }
    if (-not ([string]$Manifest.operations_root).Equals($script:OperationsRootResolved, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Manifest operations root does not match -OperationsRoot."
    }
    $files = @($Manifest.files)
    $calculatedPlan = Get-PlanSha256 -Files $files
    if ($calculatedPlan -ne ([string]$Manifest.plan_sha256).ToLowerInvariant()) {
        throw "Migration manifest plan fingerprint is invalid or the manifest was modified."
    }
    $destinationSet = @{}
    foreach ($file in $files) {
        if ([string]$file.sha256 -notmatch '^[0-9a-fA-F]{64}$' -or [long]$file.size_bytes -lt 0) {
            throw "Invalid fingerprint fields in manifest file entry: $($file.id)"
        }
        $expectedDestination = Join-RootPath -Root $script:OperationsRootResolved -RelativePath ([string]$file.destination_relative)
        $actualDestination = Assert-WithinRoot -Path ([string]$file.destination_absolute) -Root $script:OperationsRootResolved
        if (-not $expectedDestination.Equals($actualDestination, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Manifest destination absolute/relative paths disagree for entry: $($file.id)"
        }
        $key = $actualDestination.ToLowerInvariant()
        if ($destinationSet.ContainsKey($key)) {
            throw "Manifest contains duplicate destination entries: $($file.destination_relative)"
        }
        $destinationSet[$key] = $true

        if ($file.action -eq "move_file") {
            $expectedSource = Join-RootPath -Root $script:OperationsRootResolved -RelativePath ([string]$file.source_relative)
            $actualSource = Assert-WithinRoot -Path ([string]$file.source_absolute) -Root $script:OperationsRootResolved
            if (-not $expectedSource.Equals($actualSource, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Manifest source absolute/relative paths disagree for entry: $($file.id)"
            }
            $expectedQuarantine = $actualSource + $script:QuarantineSuffix
            $actualQuarantine = Assert-WithinRoot -Path ([string]$file.source_quarantine_absolute) -Root $script:OperationsRootResolved
            if (-not $expectedQuarantine.Equals($actualQuarantine, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Manifest source quarantine path is invalid for entry: $($file.id)"
            }
            $mappedDestination = Resolve-DestinationRelative -SourceRelative ([string]$file.source_relative)
            if (-not $mappedDestination.Equals([string]$file.destination_relative, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Manifest source-to-destination mapping is invalid for entry: $($file.id)"
            }
        }
        elseif ($file.action -ne "verify_existing") {
            throw "Unsupported manifest action '$($file.action)' for entry: $($file.id)"
        }
    }
    foreach ($relative in @($Manifest.required_directories)) {
        [void](Join-RootPath -Root $script:OperationsRootResolved -RelativePath ([string]$relative))
    }
    foreach ($placeholder in @($Manifest.removed_placeholders)) {
        $placeholderPath = Join-RootPath -Root $script:OperationsRootResolved -RelativePath ([string]$placeholder.path)
        if ([System.IO.Path]::GetFileName($placeholderPath) -ne ".gitkeep") {
            throw "Manifest placeholder entry is not a .gitkeep file: $($placeholder.path)"
        }
        try {
            $bytes = [Convert]::FromBase64String([string]$placeholder.content_base64)
        }
        catch {
            throw "Manifest placeholder content is not valid base64: $($placeholder.path)"
        }
        if ($bytes.LongLength -ne [long]$placeholder.size_bytes -or
            (Get-BytesSha256 -Bytes $bytes) -ne ([string]$placeholder.sha256).ToLowerInvariant()) {
            throw "Manifest placeholder content does not match its recorded fingerprint: $($placeholder.path)"
        }
    }
}

function Invoke-VerifyManifest {
    param([Parameter(Mandatory = $true)][string]$Path)
    $full = Assert-WithinRoot -Path $Path -Root $script:OperationsRootResolved
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "Manifest not found: $full" }
    $manifest = Get-Content -LiteralPath $full -Raw | ConvertFrom-Json
    Assert-ManifestIntegrity -Manifest $manifest
    if ($manifest.status -ne "completed") {
        throw "Only a completed post manifest can be verified. Manifest status: $($manifest.status)"
    }
    Assert-ManifestFilesMigrated -Manifest $manifest
    foreach ($relative in @($manifest.required_directories)) {
        $directory = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $relative
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
            throw "Required numbered-layout directory is missing: $directory"
        }
    }
    Write-Host "[VERIFY] Verified $(@($manifest.files).Count) files against manifest '$full'."
}

function Invoke-RollbackManifest {
    param([Parameter(Mandatory = $true)][string]$Path)
    $full = Assert-WithinRoot -Path $Path -Root $script:OperationsRootResolved
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "Manifest not found: $full" }
    $manifest = Get-Content -LiteralPath $full -Raw | ConvertFrom-Json
    Assert-ManifestIntegrity -Manifest $manifest

    # Full rollback preflight: no destination or source is changed unless every
    # applicable operation is safe to reverse.
    foreach ($file in @($manifest.files | Where-Object { $_.action -eq "move_file" })) {
        $destinationExists = Test-Path -LiteralPath $file.destination_absolute -PathType Leaf
        $sourceExists = Test-Path -LiteralPath $file.source_absolute -PathType Leaf
        $rollbackQuarantine = $file.destination_absolute + ".gc-layout-rollback-quarantine"
        if (-not $destinationExists -and (Test-Path -LiteralPath $file.destination_absolute)) {
            throw "A non-file entry occupies rollback destination path: $($file.destination_absolute)"
        }
        if (-not $destinationExists -and -not [bool]$file.destination_preexisting -and $sourceExists) {
            Assert-Fingerprint -Path $file.source_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
            if (Test-Path -LiteralPath $rollbackQuarantine -PathType Leaf) {
                Assert-Fingerprint -Path $rollbackQuarantine -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
            }
            elseif (Test-Path -LiteralPath $rollbackQuarantine) {
                throw "A non-file entry occupies rollback quarantine path: $rollbackQuarantine"
            }
            continue
        }
        if (Test-Path -LiteralPath $rollbackQuarantine) {
            throw "Rollback quarantine is occupied while the destination still exists: $rollbackQuarantine"
        }
        Assert-Fingerprint -Path $file.destination_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
        Assert-StableReadable -Path $file.destination_absolute
        if ($sourceExists) {
            Assert-Fingerprint -Path $file.source_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
        }
        elseif (Test-Path -LiteralPath $file.source_absolute) {
            throw "A non-file entry occupies rollback source path: $($file.source_absolute)"
        }
    }
    foreach ($placeholder in @($manifest.removed_placeholders)) {
        $placeholderPath = Join-RootPath -Root $script:OperationsRootResolved -RelativePath ([string]$placeholder.path)
        if (Test-Path -LiteralPath $placeholderPath -PathType Leaf) {
            Assert-Fingerprint -Path $placeholderPath -SizeBytes ([long]$placeholder.size_bytes) -Sha256 $placeholder.sha256
        }
        elseif (Test-Path -LiteralPath $placeholderPath) {
            throw "A non-file entry occupies rollback placeholder path: $placeholderPath"
        }
    }

    $rollbackFiles = New-Object System.Collections.Generic.List[object]
    foreach ($file in @($manifest.files | Where-Object { $_.action -eq "move_file" } | Sort-Object source_relative)) {
        if ((Test-Path -LiteralPath $file.source_absolute -PathType Leaf) -and
            -not (Test-Path -LiteralPath $file.destination_absolute -PathType Leaf) -and
            -not [bool]$file.destination_preexisting) {
            $rollbackQuarantine = $file.destination_absolute + ".gc-layout-rollback-quarantine"
            if (Test-Path -LiteralPath $rollbackQuarantine -PathType Leaf) {
                Assert-Fingerprint -Path $rollbackQuarantine -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
                Assert-Fingerprint -Path $file.source_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
                Remove-Item -LiteralPath $rollbackQuarantine -Force
                if (Test-Path -LiteralPath $rollbackQuarantine) {
                    throw "Rollback quarantine could not be removed: $rollbackQuarantine"
                }
            }
            [void]$rollbackFiles.Add([pscustomobject][ordered]@{
                id = $file.id
                source_relative = $file.source_relative
                destination_relative = $file.destination_relative
                sha256 = $file.sha256
                size_bytes = [long]$file.size_bytes
                destination_preserved = $false
                status = "already_restored"
            })
            continue
        }
        if (-not (Test-Path -LiteralPath $file.source_absolute -PathType Leaf)) {
            $sourceParent = [System.IO.Path]::GetDirectoryName($file.source_absolute)
            Ensure-Directory -Path $sourceParent
            $partial = Join-Path $sourceParent (".{0}{1}{2}.partial" -f [System.IO.Path]::GetFileName($file.source_absolute), $script:PartialMarker, [Guid]::NewGuid().ToString("N"))
            try {
                Copy-Item -LiteralPath $file.destination_absolute -Destination $partial
                Assert-Fingerprint -Path $partial -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
                [System.IO.File]::Move($partial, $file.source_absolute)
            }
            finally {
                if (Test-Path -LiteralPath $partial -PathType Leaf) { Remove-Item -LiteralPath $partial -Force }
            }
        }
        Assert-Fingerprint -Path $file.source_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
        if (-not [bool]$file.destination_preexisting) {
            $quarantine = $file.destination_absolute + ".gc-layout-rollback-quarantine"
            if (Test-Path -LiteralPath $quarantine) { throw "Rollback quarantine is occupied: $quarantine" }
            [System.IO.File]::Move($file.destination_absolute, $quarantine)
            try {
                Assert-Fingerprint -Path $quarantine -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
                Assert-Fingerprint -Path $file.source_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
                Remove-Item -LiteralPath $quarantine -Force
            }
            catch {
                if ((Test-Path -LiteralPath $quarantine -PathType Leaf) -and -not (Test-Path -LiteralPath $file.destination_absolute)) {
                    [System.IO.File]::Move($quarantine, $file.destination_absolute)
                }
                throw
            }
        }
        [void]$rollbackFiles.Add([pscustomobject][ordered]@{
            id = $file.id
            source_relative = $file.source_relative
            destination_relative = $file.destination_relative
            sha256 = $file.sha256
            size_bytes = [long]$file.size_bytes
            destination_preserved = [bool]$file.destination_preexisting
            status = "restored"
        })
    }
    foreach ($placeholder in @($manifest.removed_placeholders)) {
        $placeholderPath = Join-RootPath -Root $script:OperationsRootResolved -RelativePath $placeholder.path
        if (-not (Test-Path -LiteralPath $placeholderPath)) {
            Ensure-Directory -Path ([System.IO.Path]::GetDirectoryName($placeholderPath))
            [System.IO.File]::WriteAllBytes(
                $placeholderPath,
                [Convert]::FromBase64String([string]$placeholder.content_base64)
            )
        }
        Assert-Fingerprint -Path $placeholderPath -SizeBytes ([long]$placeholder.size_bytes) -Sha256 $placeholder.sha256
    }
    $rollbackReport = [pscustomobject][ordered]@{
        schema_version = 1
        migration = "gift-card-numbered-dropbox-layout-rollback"
        source_manifest = $full
        operations_root = $script:OperationsRootResolved
        completed_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
        status = "completed"
        files = @($rollbackFiles | ForEach-Object { $_ })
    }
    $rollbackPath = Join-Path ([System.IO.Path]::GetDirectoryName($full)) ("{0}.rollback.json" -f [System.IO.Path]::GetFileNameWithoutExtension($full))
    Write-JsonAtomic -Path $rollbackPath -Value $rollbackReport
    Write-Host "[ROLLBACK] Restored $($rollbackFiles.Count) legacy files. Report: $rollbackPath"
}

$script:OperationsRootResolved = Get-FullPath -Path $OperationsRoot
if (-not (Test-Path -LiteralPath $script:OperationsRootResolved -PathType Container)) {
    throw "Operations root does not exist: $script:OperationsRootResolved"
}

if ($Verify) {
    Invoke-VerifyManifest -Path $ManifestPath
    return
}
if ($Rollback) {
    Invoke-RollbackManifest -Path $ManifestPath
    return
}

$orphans = @(Get-OrphanPartials)
if ($orphans.Count -gt 0) {
    throw "Orphan migration partial files require review before continuing: $($orphans -join '; ')"
}

$files = @(New-Inventory)
$runId = [DateTimeOffset]::UtcNow.ToString("yyyyMMddTHHmmssfffZ")
$modeName = if ($Apply) { "apply" } else { "dry_run" }
$manifest = New-Manifest -Mode $modeName -Status "preflight_verified" -Files $files -RunId $runId
if ($manifest.summary.conflicts -gt 0) {
    $conflicts = @($manifest.files | Where-Object { $_.status -eq "conflict" } | ForEach-Object { $_.destination_relative })
    throw "Migration preflight found conflicting destination content: $($conflicts -join '; ')"
}
if (-not [string]::IsNullOrWhiteSpace($ExpectedPlanSha256) -and
    $manifest.plan_sha256 -ne $ExpectedPlanSha256.ToLowerInvariant()) {
    throw "Plan fingerprint changed. Expected $ExpectedPlanSha256; found $($manifest.plan_sha256). Run a new dry run and review the new plan."
}

if ($DryRun) {
    Write-Host ("[DRY-RUN] {0} files: {1} moves, {2} duplicate removals, {3} already migrated. Plan SHA-256: {4}" -f `
        $manifest.summary.total_files,
        $manifest.summary.planned_moves,
        $manifest.summary.planned_deduplications,
        $manifest.summary.already_migrated,
        $manifest.plan_sha256)
    Write-Output ($manifest | ConvertTo-Json -Depth 30)
    return
}

if ([string]::IsNullOrWhiteSpace($ManifestDirectory)) {
    $ManifestDirectory = Join-RootPath -Root $script:OperationsRootResolved -RelativePath "04 Archive\Cleanup Manifests"
}
else {
    $automationRunsRoot = Join-RootPath -Root $script:OperationsRootResolved -RelativePath "_automation_runs"
    $ManifestDirectory = Assert-WithinRoot -Path $ManifestDirectory -Root $automationRunsRoot
}
Ensure-Directory -Path $ManifestDirectory
$preManifestPath = Join-Path $ManifestDirectory ("GiftCard_Layout_Migration_{0}.pre.json" -f $runId)
$postManifestPath = Join-Path $ManifestDirectory ("GiftCard_Layout_Migration_{0}.post.json" -f $runId)
Write-JsonAtomic -Path $preManifestPath -Value $manifest

$manifest.status = "in_progress"
$manifest.mode = "apply"
Write-JsonAtomic -Path $postManifestPath -Value $manifest
try {
    foreach ($relative in Get-RequiredDirectoryRelatives) {
        Ensure-Directory -Path (Join-RootPath -Root $script:OperationsRootResolved -RelativePath $relative)
    }
    foreach ($file in @($manifest.files)) {
        if ($file.action -eq "move_file") {
            Copy-Verify-PublishAndRemoveSource -File $file
        }
        else {
            Assert-Fingerprint -Path $file.destination_absolute -SizeBytes ([long]$file.size_bytes) -Sha256 $file.sha256
            $file.status = "verified_existing"
        }
        $manifest.summary = Get-ManifestSummary -Files @($manifest.files)
        Write-JsonAtomic -Path $postManifestPath -Value $manifest
    }
    Remove-LegacyPlaceholdersAndEmptyDirectories -Manifest $manifest
    Assert-NoUnplannedLegacyBusinessFiles
    Assert-ManifestFilesMigrated -Manifest $manifest
    $manifest.status = "completed"
    $manifest.completed_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
    $manifest.summary = Get-ManifestSummary -Files @($manifest.files)
    Write-JsonAtomic -Path $postManifestPath -Value $manifest
    Write-Host "[APPLY] Migration completed. Preflight: $preManifestPath"
    Write-Host "[APPLY] Verified post manifest: $postManifestPath"
}
catch {
    $manifest.status = "blocked"
    $manifest.completed_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
    $manifest.error = $_.Exception.Message
    Write-JsonAtomic -Path $postManifestPath -Value $manifest
    throw
}
