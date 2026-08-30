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
$ManifestFileName = "GiftCardExportReleaseManifest.json"
$TrustedInstallerFileName = "Trusted-InstallDailyGiftCardCopyTask.ps1"
$PayloadDirectoryName = "payload"
$TrustedRoot = Split-Path -Parent $PSCommandPath
$LocalAppDataRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
$GiftCardReconRoot = Join-Path $LocalAppDataRoot "GiftCardRecon"
$ExpectedTrustedRoot = Join-Path $GiftCardReconRoot "RichmondMicrosExport"
$ManifestPath = Join-Path $TrustedRoot $ManifestFileName
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

function Get-RequiredManifestProperty {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Manifest,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $property = $Manifest.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        throw "Trusted release manifest is missing '$Name'."
    }
    return $property.Value
}

function Read-TrustedReleaseManifest {
    Assert-TrustedFile -Path $ManifestPath
    try {
        $manifest = [System.IO.File]::ReadAllText($ManifestPath, [System.Text.Encoding]::UTF8) |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Trusted release manifest is not valid JSON: $($_.Exception.Message)"
    }

    $schemaVersion = [int](Get-RequiredManifestProperty -Manifest $manifest -Name "schema_version")
    if ($schemaVersion -ne 1) {
        throw "Unsupported trusted release manifest schema version: $schemaVersion"
    }

    $releaseId = [string](Get-RequiredManifestProperty -Manifest $manifest -Name "release_id")
    if ([string]::IsNullOrWhiteSpace($releaseId)) {
        throw "Trusted release manifest release_id cannot be empty."
    }

    $releaseVersion = [int64](Get-RequiredManifestProperty -Manifest $manifest -Name "release_version")
    if ($releaseVersion -lt 1) {
        throw "Trusted release manifest release_version must be positive."
    }

    $payloadFilename = [string](Get-RequiredManifestProperty -Manifest $manifest -Name "payload_filename")
    if ($payloadFilename -cne $PayloadFileName) {
        throw "Trusted release manifest authorizes an unexpected payload name: $payloadFilename"
    }

    $payloadSha256 = [string](Get-RequiredManifestProperty -Manifest $manifest -Name "payload_sha256")
    if ($payloadSha256 -notmatch "^[0-9A-Fa-f]{64}$") {
        throw "Trusted release manifest payload_sha256 is invalid."
    }

    try {
        $payloadSizeBytes = [int64](Get-RequiredManifestProperty -Manifest $manifest -Name "payload_size_bytes")
    }
    catch {
        throw "Trusted release manifest payload_size_bytes is invalid."
    }
    if ($payloadSizeBytes -lt 0) {
        throw "Trusted release manifest payload_size_bytes cannot be negative."
    }

    return [pscustomobject]@{
        ReleaseId = $releaseId
        ReleaseVersion = $releaseVersion
        PayloadSha256 = $payloadSha256.ToUpperInvariant()
        PayloadSizeBytes = $payloadSizeBytes
    }
}

function Assert-FingerprintMatchesManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [object]$Manifest,
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $fingerprint = Get-FileFingerprint -Path $Path
    if ($fingerprint.SizeBytes -ne $Manifest.PayloadSizeBytes -or
        $fingerprint.Sha256 -cne $Manifest.PayloadSha256) {
        throw "$Description does not match the trusted release manifest."
    }
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
    if ([System.IO.Path]::GetFullPath($TrustedRoot).TrimEnd("\\") -cne $ExpectedTrustedRoot.TrimEnd("\\")) {
        throw "Trusted verifier must run from the protected RichmondMicrosExport local path."
    }
    if ([System.IO.Path]::GetFileName($PSCommandPath) -cne $TrustedInstallerFileName) {
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
    if ([System.IO.Path]::GetFileName($sourcePath) -cne $PayloadFileName) {
        throw "Dropbox payload name is unexpected: $sourcePath"
    }

    $manifest = Read-TrustedReleaseManifest
    Assert-FingerprintMatchesManifest -Path $sourcePath -Manifest $manifest -Description "Dropbox payload"

    if (-not (Test-Path -LiteralPath $PayloadDirectory)) {
        New-Item -ItemType Directory -Path $PayloadDirectory -ErrorAction Stop | Out-Null
    }
    Assert-TrustedDirectory -Path $PayloadDirectory

    $stagePath = Join-Path $PayloadDirectory (".gcs-{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
    Copy-Item -LiteralPath $sourcePath -Destination $stagePath -ErrorAction Stop
    Assert-TrustedFile -Path $stagePath
    Assert-FingerprintMatchesManifest -Path $stagePath -Manifest $manifest -Description "Staged task payload"

    Copy-Item -LiteralPath $stagePath -Destination $PayloadPath -Force -ErrorAction Stop
    Assert-TrustedFile -Path $PayloadPath
    Assert-FingerprintMatchesManifest -Path $PayloadPath -Manifest $manifest -Description "Installed task payload"

    Remove-Item -LiteralPath $stagePath -Force -ErrorAction Stop

    & $SchtasksPath /Create /TN $TaskName /TR ('"{0}"' -f $PayloadPath) /SC DAILY /ST 06:35 /F
    if ($LASTEXITCODE -ne 0) {
        throw "Secure scheduled task activation failed. Exit code: $LASTEXITCODE"
    }

    Write-Output "Authenticated release $($manifest.ReleaseId) (version $($manifest.ReleaseVersion))."
    Write-Output "Installed scheduled task '$TaskName' for 06:35 daily."
    Write-Output "Verified private task script: $PayloadPath"
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
