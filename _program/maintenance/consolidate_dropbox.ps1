#Requires -Version 5.1

<#
.SYNOPSIS
Safely consolidates the verified Gift Card Reconciliation Dropbox layout.

.DESCRIPTION
The default mode is a read-only dry run. Pass -Apply to make changes. The
script uses the exact SHA-256 inventory verified on 2026-07-11. Every move is
implemented as copy-to-temporary, hash verification, atomic rename, and only
then exact source-file removal. Duplicate deletion requires both the source
and retained reference to match the expected hash.

No directory junctions are created. No data-bearing directory is recursively
deleted. Directory pruning uses non-recursive removal after a fresh empty check.

.EXAMPLE
.\consolidate_dropbox.ps1

Performs a read-only preflight and prints the planned work.

.EXAMPLE
.\consolidate_dropbox.ps1 -Apply

Writes the cleanup manifest atomically and applies the verified consolidation.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$DropboxRoot = (Join-Path $env:USERPROFILE "Dropbox"),

    [Parameter()]
    [switch]$Apply,

    [Parameter()]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($Apply -and $DryRun) {
    throw "Choose either -Apply or -DryRun, not both. Dry run is the default."
}

$script:IsDryRun = -not $Apply
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:RunStartedAt = [DateTimeOffset]::UtcNow
$script:SafetyRoot = $null
$script:CheckpointCallback = $null

if ($null -eq ("GiftCardRecon.NativeReparsePoint" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace GiftCardRecon
{
    public static class NativeReparsePoint
    {
        private const uint FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000;
        private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
        private const uint OPEN_EXISTING = 3;
        private const uint FSCTL_GET_REPARSE_POINT = 0x000900A8;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string fileName,
            uint desiredAccess,
            FileShare shareMode,
            IntPtr securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(
            SafeFileHandle device,
            uint controlCode,
            IntPtr inputBuffer,
            int inputBufferSize,
            byte[] outputBuffer,
            int outputBufferSize,
            out int bytesReturned,
            IntPtr overlapped);

        public static uint GetTag(string path)
        {
            using (SafeFileHandle handle = CreateFile(
                path,
                0,
                FileShare.Read | FileShare.Write | FileShare.Delete,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                IntPtr.Zero))
            {
                if (handle.IsInvalid)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not open reparse point: " + path);
                }

                byte[] buffer = new byte[16 * 1024];
                int bytesReturned;
                if (!DeviceIoControl(
                    handle,
                    FSCTL_GET_REPARSE_POINT,
                    IntPtr.Zero,
                    0,
                    buffer,
                    buffer.Length,
                    out bytesReturned,
                    IntPtr.Zero))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not read reparse tag: " + path);
                }
                if (bytesReturned < 8)
                {
                    throw new InvalidDataException("Reparse buffer was too short for: " + path);
                }
                return BitConverter.ToUInt32(buffer, 0);
            }
        }

        public static bool IsCloudProjection(uint tag)
        {
            // IO_REPARSE_TAG_CLOUD and CLOUD_1 through CLOUD_F vary only in
            // the provider/variant nibble masked out below.
            return (tag & 0xFFFF0FFFu) == 0x9000001Au;
        }
    }
}
"@
}

function Write-OperationLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $prefix = if ($script:IsDryRun) { "[DRY-RUN]" } else { "[APPLY]" }
    Write-Host "$prefix $Message"
}

function Get-NormalizedFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path)
}

function Assert-PathWithinRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $candidate = Get-NormalizedFullPath -Path $Path
    $root = Get-NormalizedFullPath -Path $AllowedRoot
    $rootPrefix = $root.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar

    if (-not $candidate.Equals($root, [System.StringComparison]::OrdinalIgnoreCase) -and
        -not $candidate.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Resolved path escapes its approved root. Path='$candidate'; Root='$root'."
    }

    return $candidate
}

function Assert-NoFilesystemLinkInExistingChain {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $candidate = Assert-PathWithinRoot -Path $Path -AllowedRoot $AllowedRoot
    $root = Get-NormalizedFullPath -Path $AllowedRoot
    $probe = $candidate

    while (-not (Test-Path -LiteralPath $probe)) {
        if ($probe.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $nextProbe = [System.IO.Path]::GetDirectoryName($probe)
        if ([string]::IsNullOrWhiteSpace($nextProbe) -or $nextProbe -eq $probe) {
            throw "Could not resolve an existing ancestor within the approved root: $candidate"
        }
        $probe = $nextProbe
    }

    while ($true) {
        if (Test-Path -LiteralPath $probe) {
            $item = Get-Item -LiteralPath $probe -Force
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $reparseTag = [GiftCardRecon.NativeReparsePoint]::GetTag($item.FullName)
                if (-not [GiftCardRecon.NativeReparsePoint]::IsCloudProjection($reparseTag)) {
                    $formattedTag = $reparseTag.ToString("x8")
                    throw "Only Dropbox cloud-projection reparse points are allowed; blocked tag 0x$formattedTag at: $($item.FullName)"
                }
            }
        }

        if ($probe.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $nextProbe = [System.IO.Path]::GetDirectoryName($probe)
        if ([string]::IsNullOrWhiteSpace($nextProbe) -or $nextProbe -eq $probe) {
            throw "Existing path chain escaped its approved root: $candidate"
        }
        $probe = $nextProbe
    }
}

function Resolve-ExistingDirectoryLiteral {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $fullPath = Assert-PathWithinRoot -Path $Path -AllowedRoot $AllowedRoot
    if (-not (Test-Path -LiteralPath $fullPath -PathType Container)) {
        throw "Required directory does not exist: $fullPath"
    }

    Assert-NoFilesystemLinkInExistingChain -Path $fullPath -AllowedRoot $AllowedRoot
    $item = Get-Item -LiteralPath $fullPath -Force

    return $item.FullName
}

function Join-ApprovedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    return Assert-PathWithinRoot -Path (Join-Path $Root $RelativePath) -AllowedRoot $Root
}

function Get-ParentFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $parent = [System.IO.Path]::GetDirectoryName((Get-NormalizedFullPath -Path $Path))
    if ([string]::IsNullOrWhiteSpace($parent)) {
        throw "Path does not have a parent directory: $Path"
    }
    return $parent
}

function Get-Sha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($null -ne $script:SafetyRoot) {
        Assert-NoFilesystemLinkInExistingChain -Path $Path -AllowedRoot $script:SafetyRoot
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file does not exist: $Path"
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-StringSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $bytes = $script:Utf8NoBom.GetBytes($Value)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    }
    finally {
        $sha.Dispose()
    }

    return ([System.BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant()
}

function Assert-FileFingerprint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedSha256,

        [Parameter(Mandatory = $true)]
        [long]$ExpectedSizeBytes
    )

    if ($null -ne $script:SafetyRoot) {
        Assert-NoFilesystemLinkInExistingChain -Path $Path -AllowedRoot $script:SafetyRoot
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file is missing: $Path"
    }

    $item = Get-Item -LiteralPath $Path -Force
    if ($item.Length -ne $ExpectedSizeBytes) {
        throw "Size mismatch for '$Path'. Expected $ExpectedSizeBytes bytes; found $($item.Length)."
    }

    $actualHash = Get-Sha256 -Path $Path
    if ($actualHash -ne $ExpectedSha256.ToLowerInvariant()) {
        throw "SHA-256 mismatch for '$Path'. Expected $ExpectedSha256; found $actualHash."
    }
}

function Ensure-DirectoryLiteral {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $fullPath = Assert-PathWithinRoot -Path $Path -AllowedRoot $AllowedRoot
    Assert-NoFilesystemLinkInExistingChain -Path $fullPath -AllowedRoot $AllowedRoot
    if (Test-Path -LiteralPath $fullPath) {
        if (-not (Test-Path -LiteralPath $fullPath -PathType Container)) {
            throw "A file occupies the required directory path: $fullPath"
        }
        return
    }

    if ($script:IsDryRun) {
        Write-OperationLog "Would create directory: $fullPath"
        return
    }

    [System.IO.Directory]::CreateDirectory($fullPath) | Out-Null
    if (-not (Test-Path -LiteralPath $fullPath -PathType Container)) {
        throw "Failed to create directory: $fullPath"
    }
}

function Write-JsonAtomic {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    if ($script:IsDryRun) {
        return
    }

    $fullPath = Assert-PathWithinRoot -Path $Path -AllowedRoot $AllowedRoot
    Assert-NoFilesystemLinkInExistingChain -Path $fullPath -AllowedRoot $AllowedRoot
    $parent = Get-ParentFullPath -Path $fullPath
    Ensure-DirectoryLiteral -Path $parent -AllowedRoot $AllowedRoot

    $nonce = [Guid]::NewGuid().ToString("N")
    $temporaryPath = Join-Path $parent (".{0}.{1}.tmp" -f ([System.IO.Path]::GetFileName($fullPath)), $nonce)
    $backupPath = Join-Path $parent (".{0}.{1}.bak" -f ([System.IO.Path]::GetFileName($fullPath)), $nonce)
    $json = $Value | ConvertTo-Json -Depth 20

    try {
        [System.IO.File]::WriteAllText($temporaryPath, $json + [Environment]::NewLine, $script:Utf8NoBom)
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            # Windows PowerShell 5.1 cannot reliably bind a null backup path to
            # File.Replace. A same-directory backup keeps the replacement
            # atomic and gives the finally block a recovery copy if needed.
            [System.IO.File]::Replace($temporaryPath, $fullPath, $backupPath, $true)
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                throw "Atomic JSON replacement did not publish the destination: $fullPath"
            }
            if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
                Remove-Item -LiteralPath $backupPath -Force
            }
        }
        else {
            [System.IO.File]::Move($temporaryPath, $fullPath)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
        if (Test-Path -LiteralPath $backupPath -PathType Leaf) {
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                [System.IO.File]::Move($backupPath, $fullPath)
            }
            else {
                Remove-Item -LiteralPath $backupPath -Force
            }
        }
    }
}

function Set-OperationStatus {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter()]
        [string]$Note
    )

    $Operation.status = $Status
    $Operation.note = $Note
    if ($Status -in @("completed", "already_complete", "retained_nonempty")) {
        $Operation.completed_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
    }
    if (-not $script:IsDryRun -and $null -ne $script:CheckpointCallback) {
        & $script:CheckpointCallback
    }
}

function Get-DeletionQuarantinePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return "$Path.gc-recon-delete-quarantine"
}

function Remove-VerifiedSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedSha256,

        [Parameter(Mandatory = $true)]
        [long]$ExpectedSizeBytes,

        [Parameter(Mandatory = $true)]
        [string]$RetainedPath
    )

    $quarantinePath = Get-DeletionQuarantinePath -Path $Path
    $sourceExists = Test-Path -LiteralPath $Path -PathType Leaf
    $quarantineExists = Test-Path -LiteralPath $quarantinePath -PathType Leaf
    if ($sourceExists -and $quarantineExists) {
        throw "Both source and deletion quarantine exist; refusing ambiguous cleanup: $Path"
    }
    if (-not $sourceExists -and -not $quarantineExists) {
        throw "Verified source and its deletion quarantine are both absent: $Path"
    }

    $verifiedObjectPath = if ($quarantineExists) { $quarantinePath } else { $Path }
    Assert-FileFingerprint -Path $verifiedObjectPath -ExpectedSha256 $ExpectedSha256 -ExpectedSizeBytes $ExpectedSizeBytes
    Assert-FileFingerprint -Path $RetainedPath -ExpectedSha256 $ExpectedSha256 -ExpectedSizeBytes $ExpectedSizeBytes
    if ($script:IsDryRun) {
        Write-OperationLog "Would quarantine, reverify, and remove source file: $Path"
        return
    }

    if (-not $quarantineExists) {
        [System.IO.File]::Move($Path, $quarantinePath)
    }

    $retainedLock = $null
    try {
        Assert-FileFingerprint -Path $quarantinePath -ExpectedSha256 $ExpectedSha256 -ExpectedSizeBytes $ExpectedSizeBytes
        # Deny write/delete sharing on the retained copy while the quarantine is
        # irreversibly removed. Other readers remain allowed.
        $retainedLock = [System.IO.File]::Open(
            $RetainedPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read
        )
        Assert-FileFingerprint -Path $RetainedPath -ExpectedSha256 $ExpectedSha256 -ExpectedSizeBytes $ExpectedSizeBytes
        Remove-Item -LiteralPath $quarantinePath -Force
        if (Test-Path -LiteralPath $quarantinePath) {
            throw "Deletion quarantine could not be removed: $quarantinePath"
        }
        Assert-FileFingerprint -Path $RetainedPath -ExpectedSha256 $ExpectedSha256 -ExpectedSizeBytes $ExpectedSizeBytes
    }
    catch {
        if ((Test-Path -LiteralPath $quarantinePath -PathType Leaf) -and -not (Test-Path -LiteralPath $Path)) {
            [System.IO.File]::Move($quarantinePath, $Path)
        }
        throw
    }
    finally {
        if ($null -ne $retainedLock) {
            $retainedLock.Dispose()
        }
    }
}

function Copy-Verify-AtomicallyPublishFile {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $source = $Operation.source_absolute
    $destination = $Operation.destination_absolute
    $expectedHash = $Operation.sha256
    $expectedSize = [long]$Operation.size_bytes
    $sourceQuarantine = Get-DeletionQuarantinePath -Path $source
    foreach ($candidate in @($source, $sourceQuarantine, $destination)) {
        Assert-NoFilesystemLinkInExistingChain -Path $candidate -AllowedRoot $AllowedRoot
        if ((Test-Path -LiteralPath $candidate) -and -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            throw "A non-file entry occupies a required file path for operation '$($Operation.id)': $candidate"
        }
    }
    $sourceExists = Test-Path -LiteralPath $source -PathType Leaf
    $sourceQuarantineExists = Test-Path -LiteralPath $sourceQuarantine -PathType Leaf
    $destinationExists = Test-Path -LiteralPath $destination -PathType Leaf
    if ($sourceExists -and $sourceQuarantineExists) {
        throw "Both source and deletion quarantine exist for operation '$($Operation.id)': $source"
    }

    if ($destinationExists) {
        Assert-FileFingerprint -Path $destination -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize
        if ($sourceExists -or $sourceQuarantineExists) {
            Remove-VerifiedSourceFile -Path $source -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize -RetainedPath $destination
            Set-OperationStatus -Operation $Operation -Status $(if ($script:IsDryRun) { "planned" } else { "completed" }) -Note "Destination already matched; verified source removal completed or planned."
        }
        else {
            Set-OperationStatus -Operation $Operation -Status "already_complete" -Note "Destination already contains the verified artifact and source is absent."
        }
        return
    }

    if (-not $sourceExists) {
        if ($sourceQuarantineExists) {
            throw "A source deletion quarantine exists without its required published destination: $source"
        }
        throw "Neither verified source nor destination exists for operation '$($Operation.id)'."
    }

    Assert-FileFingerprint -Path $source -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize
    if ($script:IsDryRun) {
        Write-OperationLog "Would copy, verify, publish, and remove source: '$source' -> '$destination'"
        Set-OperationStatus -Operation $Operation -Status "planned" -Note "Dry run verified source fingerprint."
        return
    }

    $destinationParent = Get-ParentFullPath -Path $destination
    Ensure-DirectoryLiteral -Path $destinationParent -AllowedRoot $AllowedRoot
    $temporaryPath = Join-Path $destinationParent (".{0}.{1}.partial" -f ([System.IO.Path]::GetFileName($destination)), [Guid]::NewGuid().ToString("N"))

    try {
        Copy-Item -LiteralPath $source -Destination $temporaryPath
        Assert-FileFingerprint -Path $temporaryPath -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize
        [System.IO.File]::Move($temporaryPath, $destination)
        Assert-FileFingerprint -Path $destination -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize
        Remove-VerifiedSourceFile -Path $source -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize -RetainedPath $destination
        Set-OperationStatus -Operation $Operation -Status "completed" -Note "Copied to a verified temporary file, atomically published, then source removed."
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

function Remove-VerifiedDuplicateFile {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Operation
    )

    $source = $Operation.source_absolute
    $reference = $Operation.verified_against_absolute
    $expectedHash = $Operation.sha256
    $expectedSize = [long]$Operation.size_bytes
    $sourceQuarantine = Get-DeletionQuarantinePath -Path $source

    foreach ($candidate in @($source, $sourceQuarantine)) {
        Assert-NoFilesystemLinkInExistingChain -Path $candidate -AllowedRoot $script:SafetyRoot
        if ((Test-Path -LiteralPath $candidate) -and -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            throw "A non-file entry occupies a duplicate source/quarantine path: $candidate"
        }
    }

    if ((Get-NormalizedFullPath -Path $source).Equals((Get-NormalizedFullPath -Path $reference), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Duplicate source and retained reference resolve to the same lexical path: $source"
    }

    Assert-FileFingerprint -Path $reference -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize

    $sourceExists = Test-Path -LiteralPath $source -PathType Leaf
    $sourceQuarantineExists = Test-Path -LiteralPath $sourceQuarantine -PathType Leaf
    if (-not $sourceExists -and -not $sourceQuarantineExists) {
        Set-OperationStatus -Operation $Operation -Status "already_complete" -Note "Duplicate source is absent and retained reference remains verified."
        return
    }

    Remove-VerifiedSourceFile -Path $source -ExpectedSha256 $expectedHash -ExpectedSizeBytes $expectedSize -RetainedPath $reference
    if ($script:IsDryRun) {
        Set-OperationStatus -Operation $Operation -Status "planned" -Note "Dry run verified source/quarantine and retained reference."
    }
    else {
        Set-OperationStatus -Operation $Operation -Status "completed" -Note "Source was atomically quarantined, reverified with the retained reference, and removed."
    }
}

function Remove-DirectoryIfEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $path = Assert-PathWithinRoot -Path $Operation.source_absolute -AllowedRoot $AllowedRoot
    Assert-NoFilesystemLinkInExistingChain -Path $path -AllowedRoot $AllowedRoot
    if (-not (Test-Path -LiteralPath $path)) {
        Set-OperationStatus -Operation $Operation -Status "already_complete" -Note "Directory is already absent."
        return
    }
    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        throw "Expected a directory while pruning: $path"
    }

    $children = @(Get-ChildItem -LiteralPath $path -Force)
    if ($children.Count -ne 0) {
        Set-OperationStatus -Operation $Operation -Status "retained_nonempty" -Note "Directory was retained because it is not empty."
        Write-OperationLog "Retaining nonempty directory: $path"
        return
    }

    if ($script:IsDryRun) {
        Write-OperationLog "Would remove verified-empty directory: $path"
        Set-OperationStatus -Operation $Operation -Status "planned" -Note "Dry run verified that directory is empty."
        return
    }

    [System.IO.Directory]::Delete($path, $false)
    if (Test-Path -LiteralPath $path) {
        throw "Verified-empty directory could not be removed: $path"
    }
    Set-OperationStatus -Operation $Operation -Status "completed" -Note "Directory was freshly verified empty and removed non-recursively."
}

function Assert-DirectoryFileSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [object[]]$ExpectedFiles
    )

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        throw "Required directory does not exist: $Directory"
    }

    $subdirectories = @(Get-ChildItem -LiteralPath $Directory -Directory -Force)
    if ($subdirectories.Count -ne 0) {
        throw "Unexpected subdirectories found in verified flat archive: $Directory"
    }

    $actualFiles = @(Get-ChildItem -LiteralPath $Directory -File -Force)
    $expectedNames = @($ExpectedFiles | ForEach-Object { $_.name })
    $unexpected = @($actualFiles | Where-Object { $_.Name -notin $expectedNames })
    $missing = @($expectedNames | Where-Object { -not (Test-Path -LiteralPath (Join-Path $Directory $_) -PathType Leaf) })
    if ($unexpected.Count -ne 0 -or $missing.Count -ne 0 -or $actualFiles.Count -ne $ExpectedFiles.Count) {
        throw "Directory file set does not match the approved inventory: $Directory"
    }

    foreach ($expected in $ExpectedFiles) {
        Assert-FileFingerprint -Path (Join-Path $Directory $expected.name) -ExpectedSha256 $expected.sha256 -ExpectedSizeBytes ([long]$expected.size_bytes)
    }
}

function Remove-TemporaryFlatDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $fullPath = Assert-PathWithinRoot -Path $Path -AllowedRoot $AllowedRoot
    Assert-NoFilesystemLinkInExistingChain -Path $fullPath -AllowedRoot $AllowedRoot
    if (-not (Test-Path -LiteralPath $fullPath -PathType Container)) {
        return
    }

    $subdirectories = @(Get-ChildItem -LiteralPath $fullPath -Directory -Force)
    if ($subdirectories.Count -ne 0) {
        throw "Refusing to clean a temporary directory containing subdirectories: $fullPath"
    }

    foreach ($file in @(Get-ChildItem -LiteralPath $fullPath -File -Force)) {
        Remove-Item -LiteralPath $file.FullName -Force
    }
    if (@(Get-ChildItem -LiteralPath $fullPath -Force).Count -ne 0) {
        throw "Temporary directory is not empty after exact-file cleanup: $fullPath"
    }
    [System.IO.Directory]::Delete($fullPath, $false)
}

function Publish-LegacyDardenDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$DestinationDirectory,

        [Parameter(Mandatory = $true)]
        [object[]]$Files,

        [Parameter(Mandatory = $true)]
        [pscustomobject[]]$Operations,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    if ((Test-Path -LiteralPath $DestinationDirectory) -and -not (Test-Path -LiteralPath $DestinationDirectory -PathType Container)) {
        throw "A file occupies the legacy Darden destination-directory path."
    }
    $destinationExists = Test-Path -LiteralPath $DestinationDirectory -PathType Container
    if ($destinationExists) {
        Assert-DirectoryFileSet -Directory $DestinationDirectory -ExpectedFiles $Files
    }

    if (Test-Path -LiteralPath $SourceDirectory) {
        if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
            throw "A file occupies the legacy Darden source-directory path."
        }
        $allowedSourceNames = @($Files | ForEach-Object { $_.name }) +
            @($Files | ForEach-Object { "{0}.gc-recon-delete-quarantine" -f $_.name }) +
            @("README - MOVED.txt")
        $unexpectedFiles = @(Get-ChildItem -LiteralPath $SourceDirectory -File -Force | Where-Object { $_.Name -notin $allowedSourceNames })
        $unexpectedDirectories = @(Get-ChildItem -LiteralPath $SourceDirectory -Directory -Force)
        if ($unexpectedFiles.Count -ne 0 -or $unexpectedDirectories.Count -ne 0) {
            throw "Unexpected content exists in the legacy Darden source folder; no move was attempted."
        }
    }

    $sourceDataFilesPresent = @($Files | Where-Object { Test-Path -LiteralPath (Join-Path $SourceDirectory $_.name) -PathType Leaf })
    $sourceQuarantinesPresent = @($Files | Where-Object { Test-Path -LiteralPath (Get-DeletionQuarantinePath -Path (Join-Path $SourceDirectory $_.name)) -PathType Leaf })
    foreach ($file in $Files) {
        $sourceFile = Join-Path $SourceDirectory $file.name
        $sourceQuarantine = Get-DeletionQuarantinePath -Path $sourceFile
        if ((Test-Path -LiteralPath $sourceFile -PathType Leaf) -and (Test-Path -LiteralPath $sourceQuarantine -PathType Leaf)) {
            throw "Both a Darden source member and its deletion quarantine exist: $sourceFile"
        }
        if (Test-Path -LiteralPath $sourceQuarantine -PathType Leaf) {
            Assert-FileFingerprint -Path $sourceQuarantine -ExpectedSha256 $file.sha256 -ExpectedSizeBytes ([long]$file.size_bytes)
        }
    }
    if ($sourceDataFilesPresent.Count -eq 0 -and $sourceQuarantinesPresent.Count -eq 0) {
        if (-not $destinationExists) {
            throw "Legacy Darden source data and verified destination are both absent."
        }
        foreach ($operation in $Operations) {
            Set-OperationStatus -Operation $operation -Status "already_complete" -Note "Verified destination set exists and source data is absent."
        }
        return
    }

    if (-not $destinationExists -and ($sourceDataFilesPresent.Count -ne $Files.Count -or $sourceQuarantinesPresent.Count -ne 0)) {
        throw "Legacy Darden source contains only part of the approved 11-file set."
    }

    foreach ($file in $sourceDataFilesPresent) {
        Assert-FileFingerprint -Path (Join-Path $SourceDirectory $file.name) -ExpectedSha256 $file.sha256 -ExpectedSizeBytes ([long]$file.size_bytes)
    }

    if ($script:IsDryRun) {
        Write-OperationLog "Would publish the verified 11-file Darden archive and remove only its verified source files."
        foreach ($operation in $Operations) {
            Set-OperationStatus -Operation $operation -Status "planned" -Note "Dry run verified each present source or recoverable quarantine against the destination state."
        }
        return
    }

    if (-not $destinationExists) {
        $destinationParent = Get-ParentFullPath -Path $DestinationDirectory
        Ensure-DirectoryLiteral -Path $destinationParent -AllowedRoot $AllowedRoot
        $temporaryDirectory = Join-Path $destinationParent (".Darden-GC-Reconciliations.{0}.partial" -f [Guid]::NewGuid().ToString("N"))
        try {
            [System.IO.Directory]::CreateDirectory($temporaryDirectory) | Out-Null
            foreach ($file in $Files) {
                Copy-Item -LiteralPath (Join-Path $SourceDirectory $file.name) -Destination (Join-Path $temporaryDirectory $file.name)
            }
            Assert-DirectoryFileSet -Directory $temporaryDirectory -ExpectedFiles $Files
            [System.IO.Directory]::Move($temporaryDirectory, $DestinationDirectory)
            Assert-DirectoryFileSet -Directory $DestinationDirectory -ExpectedFiles $Files
        }
        finally {
            if (Test-Path -LiteralPath $temporaryDirectory -PathType Container) {
                Remove-TemporaryFlatDirectory -Path $temporaryDirectory -AllowedRoot $AllowedRoot
            }
        }
    }

    foreach ($file in $Files) {
        $sourceFile = Join-Path $SourceDirectory $file.name
        $sourceQuarantine = Get-DeletionQuarantinePath -Path $sourceFile
        $destinationFile = Join-Path $DestinationDirectory $file.name
        $operation = $Operations | Where-Object { $_.source_absolute -eq $sourceFile } | Select-Object -First 1
        if ((Test-Path -LiteralPath $sourceFile -PathType Leaf) -or (Test-Path -LiteralPath $sourceQuarantine -PathType Leaf)) {
            Remove-VerifiedSourceFile -Path $sourceFile -ExpectedSha256 $file.sha256 -ExpectedSizeBytes ([long]$file.size_bytes) -RetainedPath $destinationFile
            Set-OperationStatus -Operation $operation -Status "completed" -Note "Destination set was verified before source quarantine, re-verification, and removal."
        }
        else {
            Set-OperationStatus -Operation $operation -Status "already_complete" -Note "Destination set is verified and this source member was already absent."
        }
    }
}

function Write-RedirectAtomic {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $true)]
        [string]$AllowedRoot
    )

    $path = $Operation.destination_absolute
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        $actualContent = [System.IO.File]::ReadAllText($path, $script:Utf8NoBom)
        if ($actualContent -ne $Content) {
            throw "Existing redirect content differs from the approved text: $path"
        }
        if ((Get-Sha256 -Path $path) -ne $Operation.sha256) {
            throw "Existing redirect hash differs from the approved content: $path"
        }
        Set-OperationStatus -Operation $Operation -Status "already_complete" -Note "Redirect already exists with approved content."
        return
    }

    if ($script:IsDryRun) {
        Write-OperationLog "Would create text redirect: $path"
        Set-OperationStatus -Operation $Operation -Status "planned" -Note "Dry run; redirect content hash is recorded."
        return
    }

    $parent = Get-ParentFullPath -Path $path
    Ensure-DirectoryLiteral -Path $parent -AllowedRoot $AllowedRoot
    $temporaryPath = Join-Path $parent (".README-MOVED.{0}.tmp" -f [Guid]::NewGuid().ToString("N"))
    try {
        [System.IO.File]::WriteAllText($temporaryPath, $Content, $script:Utf8NoBom)
        if ((Get-Sha256 -Path $temporaryPath) -ne $Operation.sha256) {
            throw "Temporary redirect content failed hash verification."
        }
        [System.IO.File]::Move($temporaryPath, $path)
        if ((Get-Sha256 -Path $path) -ne $Operation.sha256) {
            throw "Published redirect content failed hash verification."
        }
        Set-OperationStatus -Operation $Operation -Status "completed" -Note "Text redirect was atomically published."
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

function Assert-CloseManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedManifestSha256,

        [Parameter(Mandatory = $true)]
        [long]$ExpectedManifestSizeBytes,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedStore,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedStatus,

        [Parameter(Mandatory = $true)]
        [string]$ArchiveRoot,

        [Parameter(Mandatory = $true)]
        [string]$GcReconRoot
    )

    Assert-FileFingerprint -Path $ManifestPath -ExpectedSha256 $ExpectedManifestSha256 -ExpectedSizeBytes $ExpectedManifestSizeBytes
    $document = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    if ([string]$document.store -ne $ExpectedStore) {
        throw "Close manifest store mismatch: $ManifestPath"
    }
    if ([string]$document.status -ne $ExpectedStatus) {
        throw "Close manifest status mismatch: $ManifestPath"
    }
    if (@($document.sources).Count -ne 9 -or @($document.artifacts).Count -ne 2) {
        throw "Close manifest evidence/artifact counts changed: $ManifestPath"
    }

    foreach ($source in @($document.sources)) {
        $relativeArchivePath = ([string]$source.archive_path).Replace("/", [System.IO.Path]::DirectorySeparatorChar)
        $sourcePath = Join-ApprovedPath -Root $ArchiveRoot -RelativePath $relativeArchivePath
        Assert-FileFingerprint -Path $sourcePath -ExpectedSha256 ([string]$source.sha256) -ExpectedSizeBytes ([long]$source.size_bytes)
    }
    foreach ($artifact in @($document.artifacts)) {
        $artifactPath = Join-ApprovedPath -Root $GcReconRoot -RelativePath ([string]$artifact.path)
        Assert-FileFingerprint -Path $artifactPath -ExpectedSha256 ([string]$artifact.sha256) -ExpectedSizeBytes ([long]$artifact.size_bytes)
    }
}

function Get-PlanFingerprint {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Operations
    )

    $lines = foreach ($operation in ($Operations | Sort-Object id)) {
        @(
            $operation.id,
            $operation.action,
            $operation.source,
            $operation.destination,
            $operation.sha256,
            $operation.size_bytes,
            $operation.verified_against
        ) -join "|"
    }
    return Get-StringSha256 -Value (($lines -join "`n") + "`n")
}

# Resolve the two approved roots before constructing any operation paths.
$dropboxRootResolved = Resolve-ExistingDirectoryLiteral -Path $DropboxRoot -AllowedRoot (Get-NormalizedFullPath -Path $DropboxRoot)
$script:SafetyRoot = $dropboxRootResolved
$gcReconRoot = Resolve-ExistingDirectoryLiteral -Path (Join-Path $dropboxRootResolved "Gift Card Reconciliation") -AllowedRoot $dropboxRootResolved
$archiveRoot = Resolve-ExistingDirectoryLiteral -Path (Join-Path $gcReconRoot "Archive - Old Files") -AllowedRoot $gcReconRoot

$canonicalCloseChecks = @(
    [pscustomobject]@{
        store = "9354"
        status = "CLOSED WITH REVIEW"
        manifest_relative = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\close_manifest.json"
        manifest_sha256 = "9bff19a0247e69f38018c177b3505422857d3350f47a5dc63c9320e0a1d70907"
        manifest_size_bytes = 5379
    },
    [pscustomobject]@{
        store = "9355"
        status = "CLOSED"
        manifest_relative = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\close_manifest.json"
        manifest_sha256 = "4fbad3f9d2b09f90b066a26b94d91efe72c0defa4a344cc66ec64d0a4ea15746"
        manifest_size_bytes = 5333
    }
)

foreach ($check in $canonicalCloseChecks) {
    Assert-CloseManifest `
        -ManifestPath (Join-ApprovedPath -Root $gcReconRoot -RelativePath $check.manifest_relative) `
        -ExpectedManifestSha256 $check.manifest_sha256 `
        -ExpectedManifestSizeBytes ([long]$check.manifest_size_bytes) `
        -ExpectedStore $check.store `
        -ExpectedStatus $check.status `
        -ArchiveRoot $archiveRoot `
        -GcReconRoot $gcReconRoot
}

$operations = New-Object System.Collections.ArrayList

function Add-FileMoveOperation {
    param(
        [string]$Id,
        [string]$Group,
        [string]$SourceAbsolute,
        [string]$DestinationAbsolute,
        [string]$SourceDisplay,
        [string]$DestinationDisplay,
        [string]$Sha256,
        [long]$SizeBytes
    )

    $operation = [pscustomobject][ordered]@{
        id = $Id
        group = $Group
        action = "copy_verify_atomic_publish_then_remove_source"
        source = $SourceDisplay
        destination = $DestinationDisplay
        sha256 = $Sha256
        size_bytes = $SizeBytes
        verified_against = $null
        status = "planned"
        completed_at_utc = $null
        note = $null
        source_absolute = $SourceAbsolute
        destination_absolute = $DestinationAbsolute
        verified_against_absolute = $null
    }
    [void]$operations.Add($operation)
    return $operation
}

function Add-DuplicateRemovalOperation {
    param(
        [string]$Id,
        [string]$Group,
        [string]$SourceAbsolute,
        [string]$ReferenceAbsolute,
        [string]$SourceDisplay,
        [string]$ReferenceDisplay,
        [string]$Sha256,
        [long]$SizeBytes
    )

    $operation = [pscustomobject][ordered]@{
        id = $Id
        group = $Group
        action = "remove_hash_proven_duplicate"
        source = $SourceDisplay
        destination = $null
        sha256 = $Sha256
        size_bytes = $SizeBytes
        verified_against = $ReferenceDisplay
        status = "planned"
        completed_at_utc = $null
        note = $null
        source_absolute = $SourceAbsolute
        destination_absolute = $null
        verified_against_absolute = $ReferenceAbsolute
    }
    [void]$operations.Add($operation)
    return $operation
}

function Add-PruneOperation {
    param(
        [string]$Id,
        [string]$SourceAbsolute,
        [string]$SourceDisplay
    )

    $operation = [pscustomobject][ordered]@{
        id = $Id
        group = "empty_directory_prune"
        action = "remove_only_if_verified_empty_nonrecursive"
        source = $SourceDisplay
        destination = $null
        sha256 = $null
        size_bytes = $null
        verified_against = $null
        status = "planned"
        completed_at_utc = $null
        note = $null
        source_absolute = $SourceAbsolute
        destination_absolute = $null
        verified_against_absolute = $null
    }
    [void]$operations.Add($operation)
    return $operation
}

$legacyDardenSourceDirectory = Join-ApprovedPath -Root $dropboxRootResolved -RelativePath "BP\Darden GC Reconciliations"
$legacyDardenDestinationDirectory = Join-ApprovedPath -Root $gcReconRoot -RelativePath "Archive - Old Files\Legacy Reconciliation\Darden GC Reconciliations"
$legacyDardenFiles = @(
    [pscustomobject]@{ name = "03.29.2026 9354 Gift Card Summary.xlsx"; sha256 = "5de4324b4e2a8bca907cad2d37987ab5371005d8e95a92767ebba1e747ce6b34"; size_bytes = 14822 },
    [pscustomobject]@{ name = "03.29.2026 9355 Gift Card Summary.xlsx"; sha256 = "03295c2871de34638f7da0eb690771c6c2aec13c9f744c39c38187541bcbd993"; size_bytes = 14827 },
    [pscustomobject]@{ name = "03.29.2026 Sorensen Gift Card Summary.xlsx"; sha256 = "2e903c8c2ba9a3cf2aa40eda1f1e2668d8c3164b9567be8baa0e16cebfb046d8"; size_bytes = 15219 },
    [pscustomobject]@{ name = "Feb_FY26_-_Sorensen_2.pdf"; sha256 = "da00addd5149dbc0d2bf5baf31169aa6c71e31bee54e455625df62725adec578"; size_bytes = 136593 },
    [pscustomobject]@{ name = "Feb_FY26_-_Sorensen-Prime_Steak.pdf"; sha256 = "c9821758e5e0947fa02eea1977de453c0bbe16111939d38a3826f1970d2a173b"; size_bytes = 137989 },
    [pscustomobject]@{ name = "Gift_Card_Reconciliation_9355_2026-05_with_weekly_variance.xlsx"; sha256 = "a0d949e883a9353e4973a2700d294edd552968fbb4979081f7ca03b32bf90e2f"; size_bytes = 36426 },
    [pscustomobject]@{ name = "Mar_FY26_-_Sorensen_2.pdf"; sha256 = "4d88ee0cddda367bc8a61e65b9583f3c70f666b6ace28ac060aa8d44dba28b58"; size_bytes = 136863 },
    [pscustomobject]@{ name = "Mar_FY26_-_Sorensen-Prime_Steak.pdf"; sha256 = "ae0e43573167f1c0d0f7b459c1e4701664cc2f8fec59ad03d7f30e69747d8b6e"; size_bytes = 137410 },
    [pscustomobject]@{ name = "March_2026_Gift_Card_Reconciliation_Report.docx"; sha256 = "2c790958d7f855a029045c5949428d5ddd1047fb14fed38b058a23ea37da58ef"; size_bytes = 37825 },
    [pscustomobject]@{ name = "RC-Prime_Steak.pdf"; sha256 = "80d0ac4da11dc06b504b151956c417c68b83a7ae0b2aab2cc9028c9d7b0be952"; size_bytes = 137018 },
    [pscustomobject]@{ name = "RC-Prime_Steak_2.pdf"; sha256 = "9d0573c510de9cd41f15a240054dc7471af25f16b22b64e4b6ef1e82c6bc0c5f"; size_bytes = 136443 }
)
if ($legacyDardenFiles.Count -ne 11) {
    throw "Internal safety error: expected exactly 11 legacy Darden files."
}

$legacyDardenOperations = @()
for ($index = 0; $index -lt $legacyDardenFiles.Count; $index++) {
    $file = $legacyDardenFiles[$index]
    $legacyDardenOperations += Add-FileMoveOperation `
        -Id ("legacy-darden-{0:D2}" -f ($index + 1)) `
        -Group "legacy_darden_directory" `
        -SourceAbsolute (Join-ApprovedPath -Root $legacyDardenSourceDirectory -RelativePath $file.name) `
        -DestinationAbsolute (Join-ApprovedPath -Root $legacyDardenDestinationDirectory -RelativePath $file.name) `
        -SourceDisplay ("BP\Darden GC Reconciliations\{0}" -f $file.name) `
        -DestinationDisplay ("Gift Card Reconciliation\Archive - Old Files\Legacy Reconciliation\Darden GC Reconciliations\{0}" -f $file.name) `
        -Sha256 $file.sha256 `
        -SizeBytes ([long]$file.size_bytes)
}

$redirectContent = @"
MOVED TO THE GC RECON ARCHIVE

This historical folder was consolidated on 2026-07-11.

Historical archive:
..\..\Gift Card Reconciliation\Archive - Old Files\Legacy Reconciliation\Darden GC Reconciliations

New monthly Darden reports:
..\..\Gift Card Reconciliation\Monthly Close\Darden Reports - Drop Here

Do not place new files in this folder.
"@
$redirectContent = $redirectContent.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
$redirectPath = Join-ApprovedPath -Root $legacyDardenSourceDirectory -RelativePath "README - MOVED.txt"
$redirectOperation = [pscustomobject][ordered]@{
    id = "legacy-darden-redirect"
    group = "legacy_darden_directory"
    action = "write_text_redirect_atomically"
    source = $null
    destination = "BP\Darden GC Reconciliations\README - MOVED.txt"
    sha256 = Get-StringSha256 -Value $redirectContent
    size_bytes = $script:Utf8NoBom.GetByteCount($redirectContent)
    verified_against = $null
    status = "planned"
    completed_at_utc = $null
    note = $null
    source_absolute = $null
    destination_absolute = $redirectPath
    verified_against_absolute = $null
}
[void]$operations.Add($redirectOperation)

$preservedFiles = @(
    [pscustomobject]@{ id = "weekly-9354-w25"; group = "legacy_weekly_reports"; source = "Archive - Old Files\pre-restructure-20260706\legacy-output\Gift_Card_Reconciliation_9354_2026-W25.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly\9354\2026\Gift_Card_Reconciliation_9354_2026-W25.xlsx"; sha256 = "dd96ee7e307f801bd33259b350e78e640d6c5483a3669c90e815b7c312568751"; size_bytes = 17297 },
    [pscustomobject]@{ id = "weekly-9354-w26"; group = "legacy_weekly_reports"; source = "Archive - Old Files\pre-restructure-20260706\legacy-output\Gift_Card_Reconciliation_9354_2026-W26.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly\9354\2026\Gift_Card_Reconciliation_9354_2026-W26.xlsx"; sha256 = "6c9fb6918f7d0460ee9e914ba29d9371dd13634cd195fffbb3bc2a0980af8430"; size_bytes = 18463 },
    [pscustomobject]@{ id = "weekly-9354-w27"; group = "legacy_weekly_reports"; source = "Archive - Old Files\pre-restructure-20260706\legacy-output\Gift_Card_Reconciliation_9354_2026-W27.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly\9354\2026\Gift_Card_Reconciliation_9354_2026-W27.xlsx"; sha256 = "4c96da0620c59fa0fb5c967ab9aa6c75b2766ac6b96e2b8e569c012b0d6b9fb1"; size_bytes = 16064 },
    [pscustomobject]@{ id = "weekly-9355-w25"; group = "legacy_weekly_reports"; source = "Archive - Old Files\pre-restructure-20260706\legacy-output\Gift_Card_Reconciliation_9355_2026-W25.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly\9355\2026\Gift_Card_Reconciliation_9355_2026-W25.xlsx"; sha256 = "10df8358de2c4bea2fbe805376e04796723435643cdd231f598fabc907937463"; size_bytes = 14658 },
    [pscustomobject]@{ id = "weekly-9355-w26"; group = "legacy_weekly_reports"; source = "Archive - Old Files\pre-restructure-20260706\legacy-output\Gift_Card_Reconciliation_9355_2026-W26.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly\9355\2026\Gift_Card_Reconciliation_9355_2026-W26.xlsx"; sha256 = "ab5c3b7266c49b1d9327bb91ca7821c58465687226c9630c276f8f359c7dd73a"; size_bytes = 15626 },
    [pscustomobject]@{ id = "weekly-9355-w27"; group = "legacy_weekly_reports"; source = "Archive - Old Files\pre-restructure-20260706\legacy-output\Gift_Card_Reconciliation_9355_2026-W27.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly\9355\2026\Gift_Card_Reconciliation_9355_2026-W27.xlsx"; sha256 = "b20bbe9d3df5c48b5d31a43d42b81e662e19e539e20c63e94b31d24c63ffd94f"; size_bytes = 14949 },
    [pscustomobject]@{ id = "weekly-output-snapshot-9354-w27"; group = "weekly_output_snapshots"; source = "Output\Gift_Card_Reconciliation_9354_2026-W27.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly Output Snapshot\9354\2026\Gift_Card_Reconciliation_9354_2026-W27.xlsx"; sha256 = "8d554533dd55c9dc696a43db79aad4d5f196b2c226be4c2fe2d13e9dab07158f"; size_bytes = 16061 },
    [pscustomobject]@{ id = "weekly-output-snapshot-9355-w27"; group = "weekly_output_snapshots"; source = "Output\Gift_Card_Reconciliation_9355_2026-W27.xlsx"; destination = "Archive - Old Files\Generated Reports\Weekly Output Snapshot\9355\2026\Gift_Card_Reconciliation_9355_2026-W27.xlsx"; sha256 = "a08723437c32bc7bb23f6c5d9a2d20b60f20d05876af509c0929d7770d3d9b13"; size_bytes = 14944 },
    [pscustomobject]@{ id = "legacy-monthly-9354"; group = "legacy_generated_reports"; source = "Output\Gift_Card_Reconciliation_9354_FY27-M01.xlsx"; destination = "Archive - Old Files\Generated Reports\Monthly Legacy\FY27-M01\Gift_Card_Reconciliation_9354_FY27-M01.xlsx"; sha256 = "ce3a0802a75030c601b9f5246c788ebf0cac5d5daa91ad4bd469768c9d97c991"; size_bytes = 39924 },
    [pscustomobject]@{ id = "legacy-monthly-9355"; group = "legacy_generated_reports"; source = "Output\Gift_Card_Reconciliation_9355_FY27-M01.xlsx"; destination = "Archive - Old Files\Generated Reports\Monthly Legacy\FY27-M01\Gift_Card_Reconciliation_9355_FY27-M01.xlsx"; sha256 = "56331b388bff105e9303bfe64c43434ce29bd82a93fad89245b5757be7a32648"; size_bytes = 32967 },
    [pscustomobject]@{ id = "diagnostic-9354-incomplete"; group = "incomplete_diagnostics"; source = "Output\Review Required\Richmond_9354_FY27-M01_Review_Required.xlsx"; destination = "Archive - Old Files\Generated Reports\Diagnostics\9354\FY27-M01\Richmond_9354_FY27-M01_Review_Required.xlsx"; sha256 = "fac959b0458abf6b4ab552510e9b1c73fcd60cb7e5f2f035d617e586e6074949"; size_bytes = 41326 }
)
if (@($preservedFiles | Where-Object { $_.group -eq "legacy_weekly_reports" }).Count -ne 6) {
    throw "Internal safety error: expected exactly 6 legacy weekly reports."
}
if (@($preservedFiles | Where-Object { $_.group -eq "weekly_output_snapshots" }).Count -ne 2) {
    throw "Internal safety error: expected exactly 2 weekly output snapshots."
}
if (@($preservedFiles | Where-Object { $_.group -eq "legacy_generated_reports" }).Count -ne 2) {
    throw "Internal safety error: expected exactly 2 legacy generated monthly reports."
}
if (@($preservedFiles | Where-Object { $_.group -eq "incomplete_diagnostics" }).Count -ne 1) {
    throw "Internal safety error: expected exactly 1 incomplete diagnostic."
}

$preserveOperations = @()
foreach ($file in $preservedFiles) {
    $preserveOperations += Add-FileMoveOperation `
        -Id $file.id `
        -Group $file.group `
        -SourceAbsolute (Join-ApprovedPath -Root $gcReconRoot -RelativePath $file.source) `
        -DestinationAbsolute (Join-ApprovedPath -Root $gcReconRoot -RelativePath $file.destination) `
        -SourceDisplay ("Gift Card Reconciliation\{0}" -f $file.source) `
        -DestinationDisplay ("Gift Card Reconciliation\{0}" -f $file.destination) `
        -Sha256 $file.sha256 `
        -SizeBytes ([long]$file.size_bytes)
}

$duplicateDefinitions = @(
    # The exact 18 legacy-input files.
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\2026-06\activity\06.07.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.07.2026 9354 Gift Card Activity.xls"; sha256 = "bb2276297acecb1b41b4e4db26494e5fdc07bdfe4c5d998eb0cceca2c63e68cb"; size_bytes = 40448 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\2026-06\activity\06.14.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.14.2026 9354 Gift Card Activity.xls"; sha256 = "f887dd5c87f941f1ab3e927827e2f3333209de46f629f532a255715ec7bd85b3"; size_bytes = 33792 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\2026-06\activity\06.21.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.21.2026 9354 Gift Card Activity.xls"; sha256 = "e0bda3c7e3b9835b127bd27378796d5aeb7039eaef4e4d678f589f5e2c215f5e"; size_bytes = 45056 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\2026-06\activity\06.28.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.28.2026 9354 Gift Card Activity.xls"; sha256 = "d5cce9f56b6bd1e009d20be96a26e32752e67dc968394e1a0981072aee0086e3"; size_bytes = 49664 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\2026-07\activity\07.05.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\07.05.2026 9354 Gift Card Activity.xls"; sha256 = "6cde183da0cacf0c011df06546b253197990c521948c1b87dd45cd221981c0dd"; size_bytes = 40448 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\weekly\activity\07.05.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\07.05.2026 9354 Gift Card Activity.xls"; sha256 = "6cde183da0cacf0c011df06546b253197990c521948c1b87dd45cd221981c0dd"; size_bytes = 40448 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\weekly\archive\2026-W25\06.21.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.21.2026 9354 Gift Card Activity.xls"; sha256 = "e0bda3c7e3b9835b127bd27378796d5aeb7039eaef4e4d678f589f5e2c215f5e"; size_bytes = 45056 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\weekly\archive\2026-W26\06.28.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.28.2026 9354 Gift Card Activity.xls"; sha256 = "d5cce9f56b6bd1e009d20be96a26e32752e67dc968394e1a0981072aee0086e3"; size_bytes = 49664 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\weekly\pos_controls.csv"; reference = "9354 - Weekly\pos_controls.csv"; sha256 = "158822ecde5c90afa606f4314a1cddbff4a5c4368234657e3ef6dadba7ce74a8"; size_bytes = 69 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\2026-06\activity\06.07.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.07.2026 9355 Gift Card Activity.xls"; sha256 = "bfac9e0dd13101914366eb2f006de01d022dccee874ea63c5fe6c8516b3bc04a"; size_bytes = 36864 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\2026-06\activity\06.14.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.14.2026 9355 Gift Card Activity.xls"; sha256 = "f0cf0e90a9a4fc3f628d563612d2edc5232f0c4b994dd5ec0fe3a4f489127b9e"; size_bytes = 36864 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\2026-06\activity\06.21.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.21.2026 9355 Gift Card Activity.xls"; sha256 = "1ca3f60e1d0e86f42092b9d80cc7c52ba2b3153b40e804b1319162f38311ae4c"; size_bytes = 34816 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\2026-06\activity\06.28.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.28.2026 9355 Gift Card Activity.xls"; sha256 = "628d1785bd62bc4a4e817b7493a870af1504e2fdce79cc2a4a01afccaf4a528d"; size_bytes = 38400 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\2026-07\activity\07.05.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\07.05.2026 9355 Gift Card Activity.xls"; sha256 = "fca157bd782d0cb622b8ec7a61bbf885ebc55c92acbf35c5159def1a1272c46a"; size_bytes = 35840 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\weekly\activity\07.05.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\07.05.2026 9355 Gift Card Activity.xls"; sha256 = "fca157bd782d0cb622b8ec7a61bbf885ebc55c92acbf35c5159def1a1272c46a"; size_bytes = 35840 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\weekly\archive\2026-W25\06.21.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.21.2026 9355 Gift Card Activity.xls"; sha256 = "1ca3f60e1d0e86f42092b9d80cc7c52ba2b3153b40e804b1319162f38311ae4c"; size_bytes = 34816 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\weekly\archive\2026-W26\06.28.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.28.2026 9355 Gift Card Activity.xls"; sha256 = "628d1785bd62bc4a4e817b7493a870af1504e2fdce79cc2a4a01afccaf4a528d"; size_bytes = 38400 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\weekly\pos_controls.csv"; reference = "9355 - Weekly\pos_controls.csv"; sha256 = "b3607c531615dd14fa50ef81cc08b70d0b6cc9159617fab790ca21b83d42dbcd"; size_bytes = 69 },

    # The exact 12 legacy-tmp files.
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\downloads\06.07.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.07.2026 9354 Gift Card Activity.xls"; sha256 = "bb2276297acecb1b41b4e4db26494e5fdc07bdfe4c5d998eb0cceca2c63e68cb"; size_bytes = 40448 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\downloads\06.07.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.07.2026 9355 Gift Card Activity.xls"; sha256 = "bfac9e0dd13101914366eb2f006de01d022dccee874ea63c5fe6c8516b3bc04a"; size_bytes = 36864 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\downloads\06.14.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.14.2026 9354 Gift Card Activity.xls"; sha256 = "f887dd5c87f941f1ab3e927827e2f3333209de46f629f532a255715ec7bd85b3"; size_bytes = 33792 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\downloads\06.14.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.14.2026 9355 Gift Card Activity.xls"; sha256 = "f0cf0e90a9a4fc3f628d563612d2edc5232f0c4b994dd5ec0fe3a4f489127b9e"; size_bytes = 36864 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\downloads\06.28.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.28.2026 9354 Gift Card Activity.xls"; sha256 = "d5cce9f56b6bd1e009d20be96a26e32752e67dc968394e1a0981072aee0086e3"; size_bytes = 49664 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9354\2026-06\activity\06.07.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.07.2026 9354 Gift Card Activity.xls"; sha256 = "bb2276297acecb1b41b4e4db26494e5fdc07bdfe4c5d998eb0cceca2c63e68cb"; size_bytes = 40448 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9354\2026-06\activity\06.14.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.14.2026 9354 Gift Card Activity.xls"; sha256 = "f887dd5c87f941f1ab3e927827e2f3333209de46f629f532a255715ec7bd85b3"; size_bytes = 33792 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9354\2026-06\activity\06.28.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.28.2026 9354 Gift Card Activity.xls"; sha256 = "d5cce9f56b6bd1e009d20be96a26e32752e67dc968394e1a0981072aee0086e3"; size_bytes = 49664 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9354\weekly\activity\06.28.2026 9354 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9354\FY27 M01 - Fiscal June\activity\06.28.2026 9354 Gift Card Activity.xls"; sha256 = "d5cce9f56b6bd1e009d20be96a26e32752e67dc968394e1a0981072aee0086e3"; size_bytes = 49664 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9355\2026-06\activity\06.07.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.07.2026 9355 Gift Card Activity.xls"; sha256 = "bfac9e0dd13101914366eb2f006de01d022dccee874ea63c5fe6c8516b3bc04a"; size_bytes = 36864 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9355\2026-06\activity\06.14.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.14.2026 9355 Gift Card Activity.xls"; sha256 = "f0cf0e90a9a4fc3f628d563612d2edc5232f0c4b994dd5ec0fe3a4f489127b9e"; size_bytes = 36864 },
    [pscustomobject]@{ source = "Archive - Old Files\pre-restructure-20260706\legacy-tmp\gmail_import_validation_20260702-232759\input\9355\weekly\activity\06.28.2026 9355 Gift Card Activity.xls"; reference = "Archive - Old Files\Monthly Close\9355\FY27 M01 - Fiscal June\activity\06.28.2026 9355 Gift Card Activity.xls"; sha256 = "628d1785bd62bc4a4e817b7493a870af1504e2fdce79cc2a4a01afccaf4a528d"; size_bytes = 38400 }
)

if ($duplicateDefinitions.Count -ne 30) {
    throw "Internal safety error: expected exactly 30 pre-restructure duplicate definitions."
}

$duplicateOperations = @()
for ($index = 0; $index -lt $duplicateDefinitions.Count; $index++) {
    $file = $duplicateDefinitions[$index]
    $duplicateOperations += Add-DuplicateRemovalOperation `
        -Id ("duplicate-{0:D2}" -f ($index + 1)) `
        -Group "verified_pre_restructure_duplicates" `
        -SourceAbsolute (Join-ApprovedPath -Root $gcReconRoot -RelativePath $file.source) `
        -ReferenceAbsolute (Join-ApprovedPath -Root $gcReconRoot -RelativePath $file.reference) `
        -SourceDisplay ("Gift Card Reconciliation\{0}" -f $file.source) `
        -ReferenceDisplay ("Gift Card Reconciliation\{0}" -f $file.reference) `
        -Sha256 $file.sha256 `
        -SizeBytes ([long]$file.size_bytes)
}

$lowercaseArchiveDefinitions = @(
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\activity\06.07.2026 9354 Gift Card Activity.xls"; sha256 = "bb2276297acecb1b41b4e4db26494e5fdc07bdfe4c5d998eb0cceca2c63e68cb"; size_bytes = 40448 },
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\activity\06.14.2026 9354 Gift Card Activity.xls"; sha256 = "f887dd5c87f941f1ab3e927827e2f3333209de46f629f532a255715ec7bd85b3"; size_bytes = 33792 },
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\activity\06.21.2026 9354 Gift Card Activity.xls"; sha256 = "e0bda3c7e3b9835b127bd27378796d5aeb7039eaef4e4d678f589f5e2c215f5e"; size_bytes = 45056 },
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\activity\06.28.2026 9354 Gift Card Activity.xls"; sha256 = "d5cce9f56b6bd1e009d20be96a26e32752e67dc968394e1a0981072aee0086e3"; size_bytes = 49664 },
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\activity\07.05.2026 9354 Gift Card Activity.xls"; sha256 = "6cde183da0cacf0c011df06546b253197990c521948c1b87dd45cd221981c0dd"; size_bytes = 40448 },
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\darden\Jun_FY27_-_Sorensen-Prime_Steak.pdf"; sha256 = "c73570fd9f583bc92b881eb17b41c3ceb1556aaf6049dd399612bd55a85df497"; size_bytes = 123731 },
    [pscustomobject]@{ relative = "9354\FY27 M01 - Fiscal June\summary\07.05.2026 9354 Gift Card Summary.xlsx"; sha256 = "885d7e30e0d82aeb755ebbcb345bd26a160b13df42c1e613eeb66921beca8508"; size_bytes = 14879 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\activity\06.07.2026 9355 Gift Card Activity.xls"; sha256 = "bfac9e0dd13101914366eb2f006de01d022dccee874ea63c5fe6c8516b3bc04a"; size_bytes = 36864 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\activity\06.14.2026 9355 Gift Card Activity.xls"; sha256 = "f0cf0e90a9a4fc3f628d563612d2edc5232f0c4b994dd5ec0fe3a4f489127b9e"; size_bytes = 36864 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\activity\06.21.2026 9355 Gift Card Activity.xls"; sha256 = "1ca3f60e1d0e86f42092b9d80cc7c52ba2b3153b40e804b1319162f38311ae4c"; size_bytes = 34816 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\activity\06.28.2026 9355 Gift Card Activity.xls"; sha256 = "628d1785bd62bc4a4e817b7493a870af1504e2fdce79cc2a4a01afccaf4a528d"; size_bytes = 38400 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\activity\07.05.2026 9355 Gift Card Activity.xls"; sha256 = "fca157bd782d0cb622b8ec7a61bbf885ebc55c92acbf35c5159def1a1272c46a"; size_bytes = 35840 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\darden\Jun_FY27_-_Sorensen_2.pdf"; sha256 = "44f53a14b189547641b7c8b793f3cedd069d894397619420ecf9db8b85767f85"; size_bytes = 122805 },
    [pscustomobject]@{ relative = "9355\FY27 M01 - Fiscal June\summary\07.05.2026 9355 Gift Card Summary.xlsx"; sha256 = "f31124a560c53d7c320bc3dbaf03d364c74be1470e69607c4a540c39534120d8"; size_bytes = 14861 }
)

if ($lowercaseArchiveDefinitions.Count -ne 14) {
    throw "Internal safety error: expected exactly 14 lowercase archive duplicate definitions."
}

$lowercaseArchiveOperations = @()
for ($index = 0; $index -lt $lowercaseArchiveDefinitions.Count; $index++) {
    $file = $lowercaseArchiveDefinitions[$index]
    $sourceRelative = "Archive - Old Files\monthly-close\$($file.relative)"
    $referenceRelative = "Archive - Old Files\Monthly Close\$($file.relative)"
    $lowercaseArchiveOperations += Add-DuplicateRemovalOperation `
        -Id ("lowercase-archive-{0:D2}" -f ($index + 1)) `
        -Group "lowercase_archive_migration" `
        -SourceAbsolute (Join-ApprovedPath -Root $gcReconRoot -RelativePath $sourceRelative) `
        -ReferenceAbsolute (Join-ApprovedPath -Root $gcReconRoot -RelativePath $referenceRelative) `
        -SourceDisplay ("Gift Card Reconciliation\{0}" -f $sourceRelative) `
        -ReferenceDisplay ("Gift Card Reconciliation\{0}" -f $referenceRelative) `
        -Sha256 $file.sha256 `
        -SizeBytes ([long]$file.size_bytes)
}

$sharedInboxPath = Join-ApprovedPath -Root $gcReconRoot -RelativePath "Monthly Close\Darden Reports - Drop Here"
$sharedInboxOperation = [pscustomobject][ordered]@{
    id = "shared-darden-inbox"
    group = "operator_layout"
    action = "ensure_directory"
    source = $null
    destination = "Gift Card Reconciliation\Monthly Close\Darden Reports - Drop Here"
    sha256 = $null
    size_bytes = $null
    verified_against = $null
    status = "planned"
    completed_at_utc = $null
    note = $null
    source_absolute = $null
    destination_absolute = $sharedInboxPath
    verified_against_absolute = $null
}
[void]$operations.Add($sharedInboxOperation)

# Derive prune candidates only from approved removal sources and fixed empty-shell paths.
$pruneStops = @(
    (Join-ApprovedPath -Root $gcReconRoot -RelativePath "Archive - Old Files\pre-restructure-20260706\legacy-input"),
    (Join-ApprovedPath -Root $gcReconRoot -RelativePath "Archive - Old Files\pre-restructure-20260706\legacy-tmp"),
    (Join-ApprovedPath -Root $gcReconRoot -RelativePath "Archive - Old Files\monthly-close")
)
$prunePathMap = @{}
foreach ($operation in @($duplicateOperations) + @($lowercaseArchiveOperations)) {
    $parent = Get-ParentFullPath -Path $operation.source_absolute
    $matchingStop = $pruneStops | Where-Object {
        $prefix = $_.TrimEnd('\') + '\'
        $parent.Equals($_, [System.StringComparison]::OrdinalIgnoreCase) -or $parent.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1
    if ($null -eq $matchingStop) {
        throw "Internal safety error: prune source is outside approved cleanup roots: $parent"
    }

    while ($true) {
        $prunePathMap[$parent.ToLowerInvariant()] = $parent
        if ($parent.Equals($matchingStop, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $next = Get-ParentFullPath -Path $parent
        if ($next -eq $parent) {
            throw "Internal safety error while deriving prune candidates."
        }
        $parent = $next
    }
}

$emptyJuneShells = @(
    "Monthly Close\9354\FY27 M01 - Fiscal June\activity",
    "Monthly Close\9354\FY27 M01 - Fiscal June\summary",
    "Monthly Close\9354\FY27 M01 - Fiscal June",
    "Monthly Close\9355\FY27 M01 - Fiscal June\activity",
    "Monthly Close\9355\FY27 M01 - Fiscal June\summary",
    "Monthly Close\9355\FY27 M01 - Fiscal June"
)
$knownEmptyLegacyDirectories = @(
    "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\2026-06\summary",
    "Archive - Old Files\pre-restructure-20260706\legacy-input\9354\weekly\summary",
    "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\2026-06\summary",
    "Archive - Old Files\pre-restructure-20260706\legacy-input\9355\weekly\summary"
)
foreach ($relativePath in @($emptyJuneShells) + @($knownEmptyLegacyDirectories)) {
    $absolutePath = Join-ApprovedPath -Root $gcReconRoot -RelativePath $relativePath
    $prunePathMap[$absolutePath.ToLowerInvariant()] = $absolutePath
}

$sortedPrunePaths = @(
    $prunePathMap.Values |
        Sort-Object @{ Expression = { $_.Length }; Descending = $true }, @{ Expression = { $_ }; Descending = $false }
)
$pruneOperations = @()
for ($index = 0; $index -lt $sortedPrunePaths.Count; $index++) {
    $path = $sortedPrunePaths[$index]
    $display = $path.Substring($dropboxRootResolved.Length).TrimStart('\')
    $pruneOperations += Add-PruneOperation -Id ("prune-empty-{0:D2}" -f ($index + 1)) -SourceAbsolute $path -SourceDisplay $display
}

$manifestPath = Join-ApprovedPath -Root $gcReconRoot -RelativePath "Archive - Old Files\Cleanup Manifests\Dropbox_Consolidation_2026-07-11.json"

# In apply mode, complete a full read-only preflight of every data-bearing
# operation before the planned manifest is written and before any mutation.
if (-not $script:IsDryRun) {
    $script:IsDryRun = $true
    try {
        Publish-LegacyDardenDirectory `
            -SourceDirectory $legacyDardenSourceDirectory `
            -DestinationDirectory $legacyDardenDestinationDirectory `
            -Files $legacyDardenFiles `
            -Operations $legacyDardenOperations `
            -AllowedRoot $dropboxRootResolved
        Write-RedirectAtomic -Operation $redirectOperation -Content $redirectContent -AllowedRoot $dropboxRootResolved
        foreach ($operation in $preserveOperations) {
            Copy-Verify-AtomicallyPublishFile -Operation $operation -AllowedRoot $gcReconRoot
        }
        foreach ($operation in @($duplicateOperations) + @($lowercaseArchiveOperations)) {
            Remove-VerifiedDuplicateFile -Operation $operation
        }
        Assert-NoFilesystemLinkInExistingChain -Path $sharedInboxPath -AllowedRoot $gcReconRoot
        if ((Test-Path -LiteralPath $sharedInboxPath) -and -not (Test-Path -LiteralPath $sharedInboxPath -PathType Container)) {
            throw "A file occupies the shared Darden inbox path: $sharedInboxPath"
        }
    }
    finally {
        $script:IsDryRun = $false
        foreach ($operation in $operations) {
            $operation.status = "planned"
            $operation.completed_at_utc = $null
            $operation.note = $null
        }
    }
}

$manifestOperations = @($operations | ForEach-Object {
    [pscustomobject][ordered]@{
        id = $_.id
        group = $_.group
        action = $_.action
        source = $_.source
        destination = $_.destination
        sha256 = $_.sha256
        size_bytes = $_.size_bytes
        verified_against = $_.verified_against
        status = $_.status
        completed_at_utc = $_.completed_at_utc
        note = $_.note
    }
})
$manifest = [pscustomobject][ordered]@{
    schema_version = 1
    inventory_verified_date = "2026-07-11"
    run_started_at_utc = $script:RunStartedAt.ToString("o")
    completed_at_utc = $null
    mode = if ($script:IsDryRun) { "dry_run" } else { "apply" }
    status = "planned"
    plan_sha256 = Get-PlanFingerprint -Operations $manifestOperations
    roots = [pscustomobject][ordered]@{
        dropbox = $dropboxRootResolved
        gc_recon = $gcReconRoot
    }
    expected_counts = [pscustomobject][ordered]@{
        legacy_darden_files = 11
        legacy_weekly_files = 6
        weekly_output_snapshots = 2
        legacy_generated_monthly_files = 2
        incomplete_diagnostics = 1
        pre_restructure_duplicates = 30
        lowercase_archive_duplicates = 14
        canonical_close_manifests = 2
        canonical_close_artifacts = 4
    }
    operations = $manifestOperations
    error = $null
}

function Sync-ManifestOperationState {
    for ($index = 0; $index -lt $manifest.operations.Count; $index++) {
        $manifestOperation = $manifest.operations[$index]
        $liveOperation = $operations | Where-Object { $_.id -eq $manifestOperation.id } | Select-Object -First 1
        if ($null -ne $liveOperation) {
            $manifestOperation.status = $liveOperation.status
            $manifestOperation.completed_at_utc = $liveOperation.completed_at_utc
            $manifestOperation.note = $liveOperation.note
        }
    }
}

Write-OperationLog ("Prepared {0} fixed or derived operations. Plan fingerprint: {1}" -f $operations.Count, $manifest.plan_sha256)
if ($script:IsDryRun) {
    Write-Host ($manifest | ConvertTo-Json -Depth 20)
}
else {
    Write-JsonAtomic -Path $manifestPath -Value $manifest -AllowedRoot $gcReconRoot
    Write-OperationLog "Wrote planned cleanup manifest atomically: $manifestPath"
    $script:CheckpointCallback = {
        Sync-ManifestOperationState
        $manifest.status = "in_progress"
        Write-JsonAtomic -Path $manifestPath -Value $manifest -AllowedRoot $gcReconRoot
    }
}

try {
    if (Test-Path -LiteralPath $redirectPath -PathType Leaf) {
        $existingRedirectHash = Get-Sha256 -Path $redirectPath
        if ($existingRedirectHash -ne $redirectOperation.sha256) {
            throw "Existing legacy-folder redirect does not match the approved content; no source data was moved."
        }
    }

    Publish-LegacyDardenDirectory `
        -SourceDirectory $legacyDardenSourceDirectory `
        -DestinationDirectory $legacyDardenDestinationDirectory `
        -Files $legacyDardenFiles `
        -Operations $legacyDardenOperations `
        -AllowedRoot $dropboxRootResolved

    Write-RedirectAtomic -Operation $redirectOperation -Content $redirectContent -AllowedRoot $dropboxRootResolved

    foreach ($operation in $preserveOperations) {
        Copy-Verify-AtomicallyPublishFile -Operation $operation -AllowedRoot $gcReconRoot
    }

    foreach ($operation in $duplicateOperations) {
        Remove-VerifiedDuplicateFile -Operation $operation
    }

    foreach ($operation in $lowercaseArchiveOperations) {
        Remove-VerifiedDuplicateFile -Operation $operation
    }

    Assert-NoFilesystemLinkInExistingChain -Path $sharedInboxPath -AllowedRoot $gcReconRoot
    if (Test-Path -LiteralPath $sharedInboxPath) {
        if (-not (Test-Path -LiteralPath $sharedInboxPath -PathType Container)) {
            throw "A file occupies the shared Darden inbox path: $sharedInboxPath"
        }
        Set-OperationStatus -Operation $sharedInboxOperation -Status "already_complete" -Note "Shared Darden inbox already exists."
    }
    elseif ($script:IsDryRun) {
        Write-OperationLog "Would create shared Darden inbox: $sharedInboxPath"
        Set-OperationStatus -Operation $sharedInboxOperation -Status "planned" -Note "Dry run."
    }
    else {
        Ensure-DirectoryLiteral -Path $sharedInboxPath -AllowedRoot $gcReconRoot
        Set-OperationStatus -Operation $sharedInboxOperation -Status "completed" -Note "Shared Darden inbox created."
    }

    foreach ($operation in $pruneOperations) {
        Remove-DirectoryIfEmpty -Operation $operation -AllowedRoot $gcReconRoot
    }

    # Postflight: canonical close certificates and evidence must remain byte-identical.
    foreach ($check in $canonicalCloseChecks) {
        Assert-CloseManifest `
            -ManifestPath (Join-ApprovedPath -Root $gcReconRoot -RelativePath $check.manifest_relative) `
            -ExpectedManifestSha256 $check.manifest_sha256 `
            -ExpectedManifestSizeBytes ([long]$check.manifest_size_bytes) `
            -ExpectedStore $check.store `
            -ExpectedStatus $check.status `
            -ArchiveRoot $archiveRoot `
            -GcReconRoot $gcReconRoot
    }

    if (-not $script:IsDryRun) {
        Assert-DirectoryFileSet -Directory $legacyDardenDestinationDirectory -ExpectedFiles $legacyDardenFiles
        if (-not (Test-Path -LiteralPath $sharedInboxPath -PathType Container)) {
            throw "Postflight failed: shared Darden inbox is missing."
        }
        foreach ($operation in $preserveOperations) {
            Assert-FileFingerprint -Path $operation.destination_absolute -ExpectedSha256 $operation.sha256 -ExpectedSizeBytes ([long]$operation.size_bytes)
            if (Test-Path -LiteralPath $operation.source_absolute -PathType Leaf) {
                throw "Postflight failed: preserved source still exists after verified publication: $($operation.source_absolute)"
            }
        }
        foreach ($operation in @($duplicateOperations) + @($lowercaseArchiveOperations)) {
            if (Test-Path -LiteralPath $operation.source_absolute -PathType Leaf) {
                throw "Postflight failed: verified duplicate source still exists: $($operation.source_absolute)"
            }
            Assert-FileFingerprint -Path $operation.verified_against_absolute -ExpectedSha256 $operation.sha256 -ExpectedSizeBytes ([long]$operation.size_bytes)
        }
        if ((Get-Sha256 -Path $redirectPath) -ne $redirectOperation.sha256) {
            throw "Postflight failed: redirect hash mismatch."
        }
    }

    Sync-ManifestOperationState
    $manifest.status = if ($script:IsDryRun) { "dry_run_verified" } else { "completed" }
    $manifest.completed_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
    if (-not $script:IsDryRun) {
        Write-JsonAtomic -Path $manifestPath -Value $manifest -AllowedRoot $gcReconRoot
        Write-OperationLog "Consolidation completed and manifest finalized: $manifestPath"
    }
    else {
        Write-OperationLog "Dry run completed. No Dropbox content or manifest was changed."
    }
}
catch {
    Sync-ManifestOperationState
    $manifest.status = "blocked"
    $manifest.completed_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
    $manifest.error = $_.Exception.Message
    if (-not $script:IsDryRun) {
        Write-JsonAtomic -Path $manifestPath -Value $manifest -AllowedRoot $gcReconRoot
    }
    throw
}
