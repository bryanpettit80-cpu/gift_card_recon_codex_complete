[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourceRoot = "",
    [Parameter(Mandatory = $true)]
    [string]$OperationsRoot,
    [string]$ProgramFolderName = "Gift Card Reconciliation Automation",
    [switch]$SkipUpstreamCheck,
    [switch]$AllowNonMainBranch,
    [switch]$SkipOperatorAssetRefresh
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$BasePath = ""
    )

    if ([IO.Path]::IsPathRooted($Path)) {
        return [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    }
    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        $BasePath = (Get-Location).Path
    }
    return [IO.Path]::GetFullPath((Join-Path $BasePath $Path)).TrimEnd('\', '/')
}

function Assert-ChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$Child,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $parentWithSeparator = $Parent.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $Child.StartsWith($parentWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Description must remain beneath $Parent. Resolved path: $Child"
    }
}

function Invoke-GitText {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $text = & git @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Arguments -join ' ') failed: $($text -join [Environment]::NewLine)"
    }
    return @($text)
}

function Get-TreeHash {
    param([Parameter(Mandatory = $true)][object[]]$Files)

    $filesByPath = [Collections.Generic.SortedDictionary[string, object]]::new([StringComparer]::Ordinal)
    foreach ($file in $Files) {
        $path = [string]$file.path
        if ([string]::IsNullOrWhiteSpace($path) -or $filesByPath.ContainsKey($path)) {
            throw "Deployment manifest contains a blank or duplicate path: $path"
        }
        $filesByPath.Add($path, $file)
    }
    $rows = foreach ($entry in $filesByPath.GetEnumerator()) {
        $file = $entry.Value
        "$($entry.Key)`t$([string]$file.sha256)`t$([long]$file.bytes)"
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

function Remove-VerifiedDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$TargetRoot,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $resolved = Resolve-AbsolutePath -Path $Path
    Assert-ChildPath -Parent $TargetRoot -Child $resolved -Description "Cleanup path"
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        if (-not (Test-Path -LiteralPath $resolved)) {
            return
        }
        try {
            Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction Stop
        }
        catch {
            if ($attempt -eq 5) {
                throw
            }
        }
        if (Test-Path -LiteralPath $resolved) {
            if ($attempt -eq 5) {
                throw "Directory cleanup remained incomplete after 5 attempts: $resolved"
            }
            Start-Sleep -Milliseconds (250 * $attempt)
        }
    }
}

function Invoke-MoveDirectoryWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination,
        [int]$Attempts = 5
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            Move-Item -LiteralPath $Source -Destination $Destination -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -eq $Attempts) {
                throw
            }
            Start-Sleep -Milliseconds (250 * $attempt)
        }
    }
}

function Assert-ExactPayload {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][object[]]$Files,
        [Parameter(Mandatory = $true)][string]$ExpectedManifestPath
    )

    $expected = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $Files) {
        [void]$expected.Add(([string]$file.path).Replace("\", "/"))
        $path = Resolve-AbsolutePath -Path ([string]$file.path).Replace('/', [IO.Path]::DirectorySeparatorChar) -BasePath $Root
        Assert-ChildPath -Parent $Root -Child $path -Description "Payload verification path"
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Deployed payload is missing $($file.path)."
        }
        $actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualHash -ne ([string]$file.sha256).ToLowerInvariant() -or (Get-Item -LiteralPath $path).Length -ne [long]$file.bytes) {
            throw "Deployed payload hash or size mismatch for $($file.path)."
        }
    }
    [void]$expected.Add("deployment-manifest.json")

    $actual = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in Get-ChildItem -LiteralPath $Root -File -Recurse -Force) {
        $relative = $item.FullName.Substring($Root.Length).TrimStart([char[]]"\/").Replace("\", "/")
        [void]$actual.Add($relative)
    }
    $extras = @($actual | Where-Object { -not $expected.Contains($_) })
    $missing = @($expected | Where-Object { -not $actual.Contains($_) })
    if ($extras.Count -gt 0 -or $missing.Count -gt 0) {
        throw "Deployed payload is not exact. Extra: $($extras -join ', '); missing: $($missing -join ', ')."
    }

    $manifestPath = Join-Path $Root "deployment-manifest.json"
    $expectedManifestHash = (Get-FileHash -LiteralPath $ExpectedManifestPath -Algorithm SHA256).Hash
    $actualManifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    if ($expectedManifestHash -ne $actualManifestHash) {
        throw "Deployed manifest hash verification failed."
    }
}

$ScriptRoot = Split-Path -Parent $PSCommandPath
if ([string]::IsNullOrWhiteSpace($SourceRoot)) {
    $SourceRoot = Split-Path -Parent (Split-Path -Parent $ScriptRoot)
}
$SourceRoot = Resolve-AbsolutePath -Path $SourceRoot
$OperationsRoot = Resolve-AbsolutePath -Path $OperationsRoot -BasePath $SourceRoot
$TargetRoot = Resolve-AbsolutePath -Path (Join-Path $OperationsRoot $ProgramFolderName)

if (-not (Test-Path -LiteralPath (Join-Path $SourceRoot ".git") -PathType Container)) {
    throw "SourceRoot must be a Git working copy: $SourceRoot"
}
if (-not (Test-Path -LiteralPath $OperationsRoot -PathType Container)) {
    throw "OperationsRoot does not exist: $OperationsRoot"
}
if ($TargetRoot.Equals($SourceRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "The deployment target cannot be the source Git working copy."
}
Assert-ChildPath -Parent $OperationsRoot -Child $TargetRoot -Description "Program deployment target"

$status = (Invoke-GitText -Arguments @("-C", $SourceRoot, "status", "--porcelain=v1")) -join "`n"
if (-not [string]::IsNullOrWhiteSpace($status)) {
    throw "The source working tree is not clean. Commit or discard changes before deployment."
}

$commit = ((Invoke-GitText -Arguments @("-C", $SourceRoot, "rev-parse", "HEAD")) -join "").Trim()
$branch = ((Invoke-GitText -Arguments @("-C", $SourceRoot, "rev-parse", "--abbrev-ref", "HEAD")) -join "").Trim()
$remoteUrl = ((Invoke-GitText -Arguments @("-C", $SourceRoot, "remote", "get-url", "origin")) -join "").Trim()
$isMainBranch = $branch -eq "main"
if (-not $isMainBranch -and -not $AllowNonMainBranch) {
    throw "Deployment is blocked from branch '$branch'. Switch to main or explicitly use -AllowNonMainBranch for controlled recovery."
}
$upstreamCommit = ""
if (-not $SkipUpstreamCheck) {
    $remoteBranch = if ($isMainBranch) { "main" } else { $branch }
    $remoteRefSpec = "+refs/heads/$($remoteBranch):refs/remotes/origin/$($remoteBranch)"
    [void](Invoke-GitText -Arguments @("-C", $SourceRoot, "fetch", "--quiet", "origin", $remoteRefSpec))
    $upstreamCommit = ((Invoke-GitText -Arguments @("-C", $SourceRoot, "rev-parse", "refs/remotes/origin/$remoteBranch")) -join "").Trim()
    if ($commit -ne $upstreamCommit) {
        throw "Deployment is blocked because local HEAD $commit does not match live origin/$remoteBranch $upstreamCommit. Push the commit first."
    }
}

$requiredProgramFiles = @(
    "_program/install.ps1",
    "_program/install_operator_assets.ps1",
    "_program/run_weekly.ps1",
    "_program/run_monthly_close.ps1",
    "_program/runtime.ps1",
    "_program/requirements.txt",
    "_program/pyproject.toml",
    "_program/check_operator_health.ps1"
)
$allTrackedFiles = @(
    @(Invoke-GitText -Arguments @("-C", $SourceRoot, "ls-files")) |
        ForEach-Object { $_.Trim().Replace("\", "/") } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
)
$trackedFiles = @(
    $allTrackedFiles | Where-Object {
        $_ -like "_program/src/*" -or
        $_ -like "templates/*" -or
        $_ -in $requiredProgramFiles
    }
)
if ($trackedFiles.Count -eq 0) {
    throw "The source repository contains no tracked operator files."
}
$missingRequired = @($requiredProgramFiles | Where-Object { $_ -notin $trackedFiles })
if ($missingRequired.Count -gt 0) {
    throw "Required operator files are not tracked: $($missingRequired -join ', ')"
}

$localAppData = $env:LOCALAPPDATA
if ([string]::IsNullOrWhiteSpace($localAppData)) {
    $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
}
$stagingBase = Resolve-AbsolutePath -Path (Join-Path $localAppData "GiftCardRecon\deploy-staging")
$stagingRoot = Resolve-AbsolutePath -Path (Join-Path $stagingBase ([guid]::NewGuid().ToString("N")))
Assert-ChildPath -Parent $stagingBase -Child $stagingRoot -Description "Local deployment staging path"

$manifestName = "deployment-manifest.json"
$manifestPath = Join-Path $TargetRoot $manifestName
$deploymentId = [guid]::NewGuid().ToString("N")
$publishStagingRoot = Resolve-AbsolutePath -Path (Join-Path $OperationsRoot ".$ProgramFolderName.deploy-$deploymentId")
$backupRoot = Resolve-AbsolutePath -Path (Join-Path $OperationsRoot ".$ProgramFolderName.backup-$deploymentId")
$failedRoot = Resolve-AbsolutePath -Path (Join-Path $OperationsRoot ".$ProgramFolderName.failed-$deploymentId")
foreach ($siblingPath in @($publishStagingRoot, $backupRoot, $failedRoot)) {
    Assert-ChildPath -Parent $OperationsRoot -Child $siblingPath -Description "Transactional deployment path"
}
$files = @()

try {
    New-Item -ItemType Directory -Force -Path $stagingRoot | Out-Null

    foreach ($relativePath in $trackedFiles) {
        if ([IO.Path]::IsPathRooted($relativePath) -or $relativePath -match "(^|[\\/])\.\.([\\/]|$)") {
            throw "Unsafe tracked path returned by Git: $relativePath"
        }
        $sourcePath = Resolve-AbsolutePath -Path $relativePath -BasePath $SourceRoot
        Assert-ChildPath -Parent $SourceRoot -Child $sourcePath -Description "Tracked source path"
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw "Tracked source file is missing: $sourcePath"
        }

        $stagePath = Resolve-AbsolutePath -Path $relativePath -BasePath $stagingRoot
        Assert-ChildPath -Parent $stagingRoot -Child $stagePath -Description "Staged source path"
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $stagePath) | Out-Null
        Copy-Item -LiteralPath $sourcePath -Destination $stagePath -Force

        $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $stageHash = (Get-FileHash -LiteralPath $stagePath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($sourceHash -ne $stageHash) {
            throw "Local staging verification failed for $relativePath."
        }
        $files += [ordered]@{
            path = $relativePath.Replace("\", "/")
            sha256 = $sourceHash
            bytes = (Get-Item -LiteralPath $stagePath).Length
        }
    }

    $manifest = [ordered]@{
        schema_version = 1
        project = "gift_card_recon_codex_complete"
        deployed_at_utc = [DateTime]::UtcNow.ToString("o")
        source_repository = $remoteUrl
        branch = $branch
        commit = $commit
        upstream_commit = $upstreamCommit
        file_count = $files.Count
        source_tree_sha256 = (Get-TreeHash -Files $files)
        files = $files
    }
    $stagedManifest = Join-Path $stagingRoot $manifestName
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $stagedManifest -Encoding UTF8

    if (-not $PSCmdlet.ShouldProcess($TargetRoot, "Deploy verified operator program commit $commit")) {
        return
    }

    New-Item -ItemType Directory -Force -Path $publishStagingRoot | Out-Null
    foreach ($file in $files) {
        $relativeNative = $file.path.Replace('/', [IO.Path]::DirectorySeparatorChar)
        $stagePath = Resolve-AbsolutePath -Path $relativeNative -BasePath $stagingRoot
        $publishPath = Resolve-AbsolutePath -Path $relativeNative -BasePath $publishStagingRoot
        Assert-ChildPath -Parent $publishStagingRoot -Child $publishPath -Description "Dropbox staging file path"
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $publishPath) | Out-Null
        Copy-Item -LiteralPath $stagePath -Destination $publishPath -Force
    }
    Copy-Item -LiteralPath $stagedManifest -Destination (Join-Path $publishStagingRoot $manifestName) -Force
    Assert-ExactPayload -Root $publishStagingRoot -Files $files -ExpectedManifestPath $stagedManifest

    $targetBackedUp = $false
    $newTargetPublished = $false
    $deploymentCommitted = $false
    try {
        if (Test-Path -LiteralPath $TargetRoot) {
            Invoke-MoveDirectoryWithRetry -Source $TargetRoot -Destination $backupRoot
            $targetBackedUp = $true
        }
        Invoke-MoveDirectoryWithRetry -Source $publishStagingRoot -Destination $TargetRoot
        $newTargetPublished = $true
        Assert-ExactPayload -Root $TargetRoot -Files $files -ExpectedManifestPath $stagedManifest

        if (-not $SkipOperatorAssetRefresh) {
            $installer = Join-Path $TargetRoot "_program\install_operator_assets.ps1"
            & $installer -OperationsRoot $OperationsRoot
        }
        $deploymentCommitted = $true
    }
    catch {
        $deploymentFailure = $_
        try {
            if ($newTargetPublished -and (Test-Path -LiteralPath $TargetRoot)) {
                Invoke-MoveDirectoryWithRetry -Source $TargetRoot -Destination $failedRoot
                $newTargetPublished = $false
            }
            if ($targetBackedUp -and (Test-Path -LiteralPath $backupRoot)) {
                Invoke-MoveDirectoryWithRetry -Source $backupRoot -Destination $TargetRoot
                $targetBackedUp = $false
            }
            if (Test-Path -LiteralPath $failedRoot) {
                Remove-VerifiedDirectory -TargetRoot $OperationsRoot -Path $failedRoot
            }
        }
        catch {
            throw "Deployment failed and automatic rollback also failed. New error: $($_.Exception.Message). Original error: $($deploymentFailure.Exception.Message)"
        }
        throw "Deployment failed; the prior operator program was restored. $($deploymentFailure.Exception.Message)"
    }

    # The verified target is now committed. Backup cleanup is deliberately
    # outside the rollback block: a partial delete must never replace a valid
    # new deployment with a partially deleted prior one.
    if ($deploymentCommitted -and $targetBackedUp -and (Test-Path -LiteralPath $backupRoot)) {
        try {
            Remove-VerifiedDirectory -TargetRoot $OperationsRoot -Path $backupRoot
            $targetBackedUp = $false
        }
        catch {
            throw "Deployment commit $commit is verified, but prior-backup cleanup is incomplete at $backupRoot. Do not roll back the valid deployment; retry cleanup after Dropbox releases the folder. $($_.Exception.Message)"
        }
    }

    [pscustomobject]@{
        OperationsRoot = $OperationsRoot
        ProgramRoot = $TargetRoot
        Commit = $commit
        FileCount = $files.Count
        SourceTreeSha256 = $manifest.source_tree_sha256
        ManifestPath = $manifestPath
    }
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Assert-ChildPath -Parent $stagingBase -Child $stagingRoot -Description "Local deployment staging cleanup"
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $publishStagingRoot) {
        Remove-VerifiedDirectory -TargetRoot $OperationsRoot -Path $publishStagingRoot
    }
}
