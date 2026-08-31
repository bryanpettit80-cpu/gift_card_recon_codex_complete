[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceScript
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$TaskName = "Gift Card Export Copy to Dropbox"
$PayloadFileName = "Copy-GiftCardExportToDropbox.cmd"
$TrustedInstallerFileName = "Trusted-InstallDailyGiftCardCopyTask.ps1"
$PayloadDirectoryName = "payload"
$ReleaseId = "richmond-gift-card-export-1"
$ReleaseVersion = [int64]1
$TrustedPayloadSha256 = "83CB42082D0E3308B2B34AE436955F04ADA6113CEDDD51C427E9E29DC775EDCF"
$TrustedPayloadSizeBytes = [int64]1296
$TrustedRoot = Split-Path -Parent $PSCommandPath
$LocalAppDataRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
$GiftCardReconRoot = Join-Path $LocalAppDataRoot "GiftCardRecon"
$ExpectedTrustedRoot = Join-Path $GiftCardReconRoot "RichmondMicrosExport"
$PayloadDirectory = Join-Path $TrustedRoot $PayloadDirectoryName
$PayloadPath = Join-Path $PayloadDirectory $PayloadFileName
$SystemDirectory = [Environment]::SystemDirectory
$SchtasksPath = Join-Path $SystemDirectory "schtasks.exe"
$CmdPath = Join-Path $SystemDirectory "cmd.exe"
$SafeAction = ('"{0}" /d /c exit 0' -f $CmdPath)
$CurrentUserSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$AllowedWriterSids = @(
    $CurrentUserSid,
    "S-1-5-18", # LOCAL SYSTEM
    "S-1-5-32-544", # BUILTIN\\Administrators
    "S-1-3-4" # OWNER RIGHTS; the owner is separately restricted above.
)
$WriteRights = (
    [System.Security.AccessControl.FileSystemRights]::WriteData -bor
    [System.Security.AccessControl.FileSystemRights]::AppendData -bor
    [System.Security.AccessControl.FileSystemRights]::Delete -bor
    [System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor
    [System.Security.AccessControl.FileSystemRights]::WriteAttributes -bor
    [System.Security.AccessControl.FileSystemRights]::WriteExtendedAttributes -bor
    [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
    [System.Security.AccessControl.FileSystemRights]::TakeOwnership
)

function Get-IdentitySid {
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Principal.IdentityReference]$Identity
    )

    try {
        return $Identity.Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        throw "Cannot resolve ACL identity '$Identity' to a SID."
    }
}

function Assert-RestrictedWriteAcl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.PSIsContainer) {
        $acl = [System.IO.Directory]::GetAccessControl($item.FullName)
    }
    else {
        $acl = [System.IO.File]::GetAccessControl($item.FullName)
    }
    $ownerSid = Get-IdentitySid -Identity ([System.Security.Principal.NTAccount]$acl.Owner)
    if ($ownerSid -notin $AllowedWriterSids) {
        throw "Trusted path owner is not an approved local authority: $Path"
    }

    foreach ($rule in $acl.Access) {
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) {
            continue
        }
        if (($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) {
            continue
        }
        if (([int64]$rule.FileSystemRights -band [int64]$WriteRights) -eq 0) {
            continue
        }

        $sid = Get-IdentitySid -Identity $rule.IdentityReference
        if ($sid -notin $AllowedWriterSids) {
            throw "Trusted path grants write access to an unapproved principal ($sid): $Path"
        }
    }
}

function Assert-TrustedDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (-not $item.PSIsContainer) {
        throw "Trusted path is not a directory: $Path"
    }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Trusted directory cannot be a reparse point: $Path"
    }
    Assert-RestrictedWriteAcl -Path $Path
}

function Assert-TrustedFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.PSIsContainer) {
        throw "Trusted path is not a file: $Path"
    }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Trusted file cannot be a reparse point: $Path"
    }
    Assert-RestrictedWriteAcl -Path $Path
}

function Get-FileFingerprint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [pscustomobject]@{
            Sha256 = [System.BitConverter]::ToString($sha256.ComputeHash($bytes)).Replace("-", "")
            SizeBytes = [int64]$bytes.LongLength
        }
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-TrustedReleaseMetadata {
    if ($TrustedPayloadSha256 -notmatch "^[0-9A-F]{64}$") {
        throw "Embedded trusted payload SHA-256 is invalid."
    }
    if ($TrustedPayloadSizeBytes -lt 0) {
        throw "Embedded trusted payload size cannot be negative."
    }
    return [pscustomobject]@{
        ReleaseId = $ReleaseId
        ReleaseVersion = $ReleaseVersion
        PayloadSha256 = $TrustedPayloadSha256
        PayloadSizeBytes = $TrustedPayloadSizeBytes
    }
}

function Assert-FingerprintMatchesRelease {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [object]$Release,
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $fingerprint = Get-FileFingerprint -Path $Path
    if ($fingerprint.SizeBytes -ne $Release.PayloadSizeBytes -or
        $fingerprint.Sha256 -cne $Release.PayloadSha256) {
        throw "$Description does not match the embedded trusted release metadata."
    }
}

function Test-EquivalentWindowsPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Left,
        [Parameter(Mandatory = $true)]
        [string]$Right
    )

    $trimCharacters = [char[]]@(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $normalizedLeft = [System.IO.Path]::GetFullPath($Left).TrimEnd($trimCharacters)
    $normalizedRight = [System.IO.Path]::GetFullPath($Right).TrimEnd($trimCharacters)
    return [string]::Equals(
        $normalizedLeft,
        $normalizedRight,
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

try {
    if ([string]::IsNullOrWhiteSpace($LocalAppDataRoot) -or
        -not (Test-Path -LiteralPath $SchtasksPath -PathType Leaf) -or
        -not (Test-Path -LiteralPath $CmdPath -PathType Leaf)) {
        throw "Required trusted Windows paths are unavailable."
    }
    & $SchtasksPath /Create /TN $TaskName /TR $SafeAction /SC DAILY /ST 06:35 /F
    if ($LASTEXITCODE -ne 0) {
        throw "Cannot neutralize the existing scheduled task. Exit code: $LASTEXITCODE"
    }
    if (-not (Test-EquivalentWindowsPath -Left $TrustedRoot -Right $ExpectedTrustedRoot)) {
        throw "Trusted verifier must run from the protected RichmondMicrosExport local path."
    }
    if (-not [string]::Equals(
            [System.IO.Path]::GetFileName($PSCommandPath),
            $TrustedInstallerFileName,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
        throw "Trusted verifier filename is unexpected: $PSCommandPath"
    }
    Assert-TrustedDirectory -Path $LocalAppDataRoot
    Assert-TrustedDirectory -Path $GiftCardReconRoot
    Assert-TrustedDirectory -Path $TrustedRoot
    Assert-TrustedFile -Path $PSCommandPath

    $sourcePath = [System.IO.Path]::GetFullPath($SourceScript)
    $sourceItem = Get-Item -LiteralPath $sourcePath -Force -ErrorAction Stop
    if ($sourceItem.PSIsContainer) {
        throw "Dropbox payload path is not a file: $sourcePath"
    }
    if (-not [string]::Equals(
            [System.IO.Path]::GetFileName($sourcePath),
            $PayloadFileName,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
        throw "Dropbox payload name is unexpected: $sourcePath"
    }

    $release = Get-TrustedReleaseMetadata
    Assert-FingerprintMatchesRelease -Path $sourcePath -Release $release -Description "Dropbox payload"

    if (-not (Test-Path -LiteralPath $PayloadDirectory)) {
        New-Item -ItemType Directory -Path $PayloadDirectory -ErrorAction Stop | Out-Null
    }
    Assert-TrustedDirectory -Path $PayloadDirectory

    $stagePath = Join-Path $PayloadDirectory (".gcs-{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
    Copy-Item -LiteralPath $sourcePath -Destination $stagePath -ErrorAction Stop
    Assert-TrustedFile -Path $stagePath
    Assert-FingerprintMatchesRelease -Path $stagePath -Release $release -Description "Staged task payload"

    Copy-Item -LiteralPath $stagePath -Destination $PayloadPath -Force -ErrorAction Stop
    Assert-TrustedFile -Path $PayloadPath
    Assert-FingerprintMatchesRelease -Path $PayloadPath -Release $release -Description "Installed task payload"

    Remove-Item -LiteralPath $stagePath -Force -ErrorAction Stop

    & $SchtasksPath /Create /TN $TaskName /TR ('"{0}"' -f $PayloadPath) /SC DAILY /ST 06:35 /F
    if ($LASTEXITCODE -ne 0) {
        throw "Secure scheduled task activation failed. Exit code: $LASTEXITCODE"
    }

    Write-Output "Authenticated release $($release.ReleaseId) (version $($release.ReleaseVersion))."
    Write-Output "Installed scheduled task '$TaskName' for 06:35 daily."
    Write-Output "Verified private task script: $PayloadPath"
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
