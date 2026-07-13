#Requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Write-TestFile {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Content
    )
    $path = Join-Path $Root $RelativePath
    $parent = [System.IO.Path]::GetDirectoryName($path)
    [void][System.IO.Directory]::CreateDirectory($parent)
    [System.IO.File]::WriteAllText($path, $Content, (New-Object System.Text.UTF8Encoding($false)))
    return $path
}

function Get-OnlyPostManifest {
    param([Parameter(Mandatory = $true)][string]$ManifestDirectory)
    $posts = @(Get-ChildItem -LiteralPath $ManifestDirectory -Filter "*.post.json" -File | Sort-Object LastWriteTimeUtc)
    if ($posts.Count -eq 0) { throw "No post manifest was produced in $ManifestDirectory" }
    return $posts[-1].FullName
}

function Write-TestJson {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )
    [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($Path))
    [System.IO.File]::WriteAllText(
        $Path,
        ($Value | ConvertTo-Json -Depth 30),
        (New-Object System.Text.UTF8Encoding($false))
    )
}

function Invoke-ApplyResumeFixture {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][bool]$MoveOneFileBeforeResume,
        [Parameter(Mandatory = $true)][bool]$WritePostBeforeResume,
        [Parameter(Mandatory = $true)][string]$MigrationScript
    )
    if ($MoveOneFileBeforeResume -and -not $WritePostBeforeResume) {
        throw "A moved-file resume fixture requires a post checkpoint."
    }
    [void][System.IO.Directory]::CreateDirectory($Root)
    [void](Write-TestFile -Root $Root -RelativePath "9354 - Weekly\activity\resume-a.xls" -Content "resume-a")
    [void](Write-TestFile -Root $Root -RelativePath "9355 - Weekly\activity\resume-b.xls" -Content "resume-b")

    $dry = (& $MigrationScript -OperationsRoot $Root) | ConvertFrom-Json
    $manifestDirectory = Join-Path $Root "_automation_runs\migration"
    $runId = [string]$dry.run_id
    $prePath = Join-Path $manifestDirectory ("GiftCard_Layout_Migration_{0}.pre.json" -f $runId)
    $postPath = Join-Path $manifestDirectory ("GiftCard_Layout_Migration_{0}.post.json" -f $runId)

    $pre = $dry | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $pre.mode = "apply"
    $pre.status = "preflight_verified"
    Write-TestJson -Path $prePath -Value $pre

    if ($WritePostBeforeResume) {
        $post = $pre | ConvertTo-Json -Depth 30 | ConvertFrom-Json
        $post.status = "in_progress"
        if ($MoveOneFileBeforeResume) {
            $entry = @($post.files | Where-Object { $_.action -eq "move_file" })[0]
            [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName([string]$entry.destination_absolute))
            Copy-Item -LiteralPath $entry.source_absolute -Destination $entry.destination_absolute
            Remove-Item -LiteralPath $entry.source_absolute -Force
            $entry.status = "moved"
        }
        Write-TestJson -Path $postPath -Value $post
    }

    & $MigrationScript `
        -OperationsRoot $Root `
        -Apply `
        -ManifestDirectory $manifestDirectory `
        -ExpectedPlanSha256 $dry.plan_sha256

    $completed = Get-Content -LiteralPath $postPath -Raw | ConvertFrom-Json
    Assert-True ($completed.status -eq "completed") "resumed apply completes"
    Assert-True ($completed.plan_sha256 -eq $dry.plan_sha256) "resumed apply retains original reviewed fingerprint"
    Assert-True (@($completed.files | Where-Object { $_.status -eq "moved" }).Count -eq 2) "resumed apply completes every planned move"
    Assert-True (@(Get-ChildItem -LiteralPath $manifestDirectory -Filter "*.post.json" -File).Count -eq 1) "resume continues the original checkpoint instead of creating a new plan"
}

$migrationScript = Join-Path $PSScriptRoot "migrate_to_numbered_layout.ps1"
$tempRoot = [System.IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')
$fixture = Join-Path $tempRoot ("GiftCardLayoutMigrationTest-{0}" -f [Guid]::NewGuid().ToString("N"))

try {
    [void][System.IO.Directory]::CreateDirectory($fixture)
    $files = [ordered]@{}
    $files.weekly9354 = Write-TestFile -Root $fixture -RelativePath "9354 - Weekly\activity\07.12.2026 9354 Gift Card Activity.xls" -Content "activity-9354"
    $files.controls9354 = Write-TestFile -Root $fixture -RelativePath "9354 - Weekly\pos_controls.csv" -Content "date,issues,payments`n2026-07-12,1.00,2.00"
    $files.weekly9355 = Write-TestFile -Root $fixture -RelativePath "9355 - Weekly\activity\07.12.2026 9355 Gift Card Activity.xls" -Content "activity-9355"
    $files.controls9355 = Write-TestFile -Root $fixture -RelativePath "9355 - Weekly\pos_controls.csv" -Content "date,issues,payments`n2026-07-12,3.00,4.00"
    $files.darden = Write-TestFile -Root $fixture -RelativePath "Monthly Close\Darden Reports - Drop Here\Jul_FY27.pdf" -Content "darden"
    $files.monthlyInput = Write-TestFile -Root $fixture -RelativePath "Monthly Close\9354\FY27 M02\activity\source.xls" -Content "monthly-input"
    $files.monthlyReport = Write-TestFile -Root $fixture -RelativePath "Output\Monthly Close\FY27 M02\Richmond.xlsx" -Content "monthly-report"
    $files.weeklyReport = Write-TestFile -Root $fixture -RelativePath "Output\Weekly\9354\2026\week-28.xlsx" -Content "weekly-report"
    $files.archive = Write-TestFile -Root $fixture -RelativePath "Archive - Old Files\Monthly Close\9354\FY27 M01\close_manifest.json" -Content '{"archive":"unchanged"}'
    $files.legacyInput = Write-TestFile -Root $fixture -RelativePath "input\9355\old-source.xls" -Content "legacy-input"
    $files.legacyReport = Write-TestFile -Root $fixture -RelativePath "reports\old-report.xlsx" -Content "legacy-report"
    $files.program = Write-TestFile -Root $fixture -RelativePath "_program\cache\program.txt" -Content "must-not-move"
    $files.nestedProgramPlaceholder = Write-TestFile -Root $fixture -RelativePath "Archive - Old Files\_program\empty\.gitkeep" -Content "program-placeholder`r`n"
    $files.weeklyPlaceholder = Write-TestFile -Root $fixture -RelativePath "9354 - Weekly\activity\.gitkeep" -Content "`r`n"
    $files.archivePlaceholder = Write-TestFile -Root $fixture -RelativePath "Archive - Old Files\.gitkeep" -Content "`r`n"
    $beforeHashes = @{}
    foreach ($key in $files.Keys) {
        $beforeHashes[$key] = (Get-FileHash -LiteralPath $files[$key] -Algorithm SHA256).Hash
    }

    # Dry run returns a complete manifest but creates no numbered directories.
    $dryJson = & $migrationScript -OperationsRoot $fixture
    $dryManifest = $dryJson | ConvertFrom-Json
    Assert-True ($dryManifest.status -eq "preflight_verified") "dry-run status"
    Assert-True ($dryManifest.summary.total_files -eq 11) "dry-run inventories all and only business files"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $fixture "01 Weekly Gift Card Activity Reports"))) "dry run must not create directories"
    Assert-True (@($dryManifest.files | Where-Object { $_.source_relative -like "_program*" }).Count -eq 0) "program files must be excluded"
    Assert-True (@($dryManifest.files | Where-Object { $_.source_relative -like "*\_program\*" }).Count -eq 0) "nested program files must be excluded"

    [void](Write-TestFile -Root $fixture -RelativePath "04 Archive\Cleanup Manifests\GiftCard_Layout_Migration_prior.pre.json" -Content '{"generated":"pre"}')
    [void](Write-TestFile -Root $fixture -RelativePath "04 Archive\Cleanup Manifests\GiftCard_Layout_Migration_prior.post.json" -Content '{"generated":"post"}')
    [void](Write-TestFile -Root $fixture -RelativePath "04 Archive\Cleanup Manifests\GiftCard_Layout_Migration_prior.post.rollback.json" -Content '{"generated":"rollback"}')
    [void](Write-TestFile -Root $fixture -RelativePath "04 Archive\Cleanup Manifests\.GiftCard_Layout_Migration_prior.post.json.0123456789abcdef0123456789abcdef.tmp" -Content '{"generated":"atomic-temp"}')
    [void](Write-TestFile -Root $fixture -RelativePath "04 Archive\Cleanup Manifests\.GiftCard_Layout_Migration_prior.post.json.fedcba9876543210fedcba9876543210.bak" -Content '{"generated":"atomic-backup"}')
    $withGeneratedManifests = (& $migrationScript -OperationsRoot $fixture) | ConvertFrom-Json
    Assert-True ($withGeneratedManifests.plan_sha256 -eq $dryManifest.plan_sha256) "migration-generated manifests must not change the reviewed plan fingerprint"
    Assert-True ($withGeneratedManifests.summary.total_files -eq 11) "migration-generated manifests must be excluded from the business plan"

    $missingFingerprintBlocked = $false
    try {
        & $migrationScript -OperationsRoot $fixture -Apply -ManifestDirectory (Join-Path $fixture "_automation_runs\migration")
    }
    catch {
        $missingFingerprintBlocked = $_.Exception.Message -like "*-ExpectedPlanSha256*"
    }
    Assert-True $missingFingerprintBlocked "Apply must require a reviewed dry-run fingerprint"

    $manifestDirectory = Join-Path $fixture "_automation_runs\migration"
    & $migrationScript -OperationsRoot $fixture -Apply -ManifestDirectory $manifestDirectory -ExpectedPlanSha256 $dryManifest.plan_sha256
    $firstPostPath = Get-OnlyPostManifest -ManifestDirectory $manifestDirectory
    $firstPost = Get-Content -LiteralPath $firstPostPath -Raw | ConvertFrom-Json
    Assert-True ($firstPost.status -eq "completed") "apply post manifest status"
    Assert-True (@($firstPost.files | Where-Object { $_.status -eq "moved" }).Count -eq 11) "all fixture business files moved"
    Assert-True (Test-Path -LiteralPath $files.nestedProgramPlaceholder -PathType Leaf) "excluded nested program placeholder is preserved"

    $expectedDestinations = @(
        "01 Weekly Gift Card Activity Reports\9354 Richmond\activity\07.12.2026 9354 Gift Card Activity.xls",
        "01 Weekly Gift Card Activity Reports\9355 Virginia Beach\activity\07.12.2026 9355 Gift Card Activity.xls",
        "02 Monthly Close Inputs\Darden Reports - Drop Here\Jul_FY27.pdf",
        "02 Monthly Close Inputs\9354 Richmond\FY27 M02\activity\source.xls",
        "03 Finished Reports\Monthly Close\FY27 M02\Richmond.xlsx",
        "03 Finished Reports\Weekly\9354 Richmond\2026\week-28.xlsx",
        "04 Archive\Monthly Close\9354\FY27 M01\close_manifest.json",
        "04 Archive\Legacy Reconciliation\Manual POS Controls\9354\pos_controls.csv",
        "04 Archive\Legacy Reconciliation\Manual POS Controls\9355\pos_controls.csv",
        "04 Archive\Legacy Reconciliation\input\9355\old-source.xls",
        "04 Archive\Generated Reports\Legacy Reports\old-report.xlsx"
    )
    foreach ($relative in $expectedDestinations) {
        Assert-True (Test-Path -LiteralPath (Join-Path $fixture $relative) -PathType Leaf) "expected destination exists: $relative"
    }
    Assert-True ((Get-Content -LiteralPath (Join-Path $fixture "04 Archive\Monthly Close\9354\FY27 M01\close_manifest.json") -Raw) -eq '{"archive":"unchanged"}') "archive bytes and internal path preserved"
    Assert-True ((Get-FileHash -LiteralPath $files.program -Algorithm SHA256).Hash -eq $beforeHashes.program) "program file remains untouched"
    $retainedLegacyFiles = @(Get-ChildItem -LiteralPath (Join-Path $fixture "Archive - Old Files") -File -Recurse -Force)
    Assert-True ($retainedLegacyFiles.Count -eq 1 -and $retainedLegacyFiles[0].FullName -eq $files.nestedProgramPlaceholder) "legacy archive retains only the excluded nested program placeholder"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $fixture "9354 - Weekly"))) "nonprotected empty legacy weekly root pruned"

    # Verification is read-only and driven by the post manifest.
    & $migrationScript -OperationsRoot $fixture -Verify -ManifestPath $firstPostPath

    # A modified manifest cannot redirect verification or rollback outside the
    # reviewed plan. The plan fingerprint detects the edit before file access.
    $tampered = Get-Content -LiteralPath $firstPostPath -Raw | ConvertFrom-Json
    $tampered.files[0].destination_absolute = Join-Path $env:SystemRoot "not-a-migration-target.bin"
    $tamperedPath = Join-Path $manifestDirectory "tampered.post.json"
    [System.IO.File]::WriteAllText($tamperedPath, ($tampered | ConvertTo-Json -Depth 30), (New-Object System.Text.UTF8Encoding($false)))
    $tamperBlocked = $false
    try {
        & $migrationScript -OperationsRoot $fixture -Verify -ManifestPath $tamperedPath
    }
    catch {
        $tamperBlocked = $_.Exception.Message -like "*fingerprint*" -or $_.Exception.Message -like "*escapes*"
    }
    Assert-True $tamperBlocked "tampered manifest must be rejected"

    # Re-running Apply on the finished layout is idempotent.
    Start-Sleep -Milliseconds 20
    $secondDryJson = & $migrationScript -OperationsRoot $fixture
    $secondDry = $secondDryJson | ConvertFrom-Json
    Assert-True ($secondDry.summary.total_files -eq 11) "migration manifests do not enter later inventories"
    & $migrationScript -OperationsRoot $fixture -Apply -ManifestDirectory $manifestDirectory -ExpectedPlanSha256 $secondDry.plan_sha256
    $secondPostPath = Get-OnlyPostManifest -ManifestDirectory $manifestDirectory
    $secondPost = Get-Content -LiteralPath $secondPostPath -Raw | ConvertFrom-Json
    Assert-True ($secondPost.status -eq "completed") "second apply completes"
    Assert-True ($secondPost.summary.planned_moves -eq 0) "second apply plans no moves"
    Assert-True (@($secondPost.files | Where-Object { $_.status -eq "verified_existing" }).Count -eq 11) "second apply verifies existing files"

    # Rollback restores original locations and removes only destinations that
    # did not exist before the first migration. Simulate a hard crash after
    # one destination was quarantined but before that verified duplicate was
    # deleted; the resumed rollback must clean it up idempotently.
    $crashEntry = @($firstPost.files | Where-Object { $_.action -eq "move_file" -and -not [bool]$_.destination_preexisting })[0]
    [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName([string]$crashEntry.source_absolute)) | Out-Null
    Copy-Item -LiteralPath $crashEntry.destination_absolute -Destination $crashEntry.source_absolute
    $crashQuarantine = [string]$crashEntry.destination_absolute + ".gc-layout-rollback-quarantine"
    [System.IO.File]::Move([string]$crashEntry.destination_absolute, $crashQuarantine)
    & $migrationScript -OperationsRoot $fixture -Rollback -ManifestPath $firstPostPath
    Assert-True (-not (Test-Path -LiteralPath $crashQuarantine)) "resumed rollback removes verified quarantine duplicate"
    foreach ($key in @($files.Keys | Where-Object { $_ -ne "program" })) {
        Assert-True (Test-Path -LiteralPath $files[$key] -PathType Leaf) "rollback source restored: $key"
        Assert-True ((Get-FileHash -LiteralPath $files[$key] -Algorithm SHA256).Hash -eq $beforeHashes[$key]) "rollback hash restored: $key"
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $fixture "04 Archive\Monthly Close\9354\FY27 M01\close_manifest.json"))) "migration-created destination removed on rollback"
    & $migrationScript -OperationsRoot $fixture -Rollback -ManifestPath $firstPostPath

    # A conflicting destination blocks the full Apply preflight; unrelated
    # sources must remain in place.
    $conflictPath = Write-TestFile -Root $fixture -RelativePath "03 Finished Reports\Monthly Close\FY27 M02\Richmond.xlsx" -Content "different-content"
    $unrelatedHash = (Get-FileHash -LiteralPath $files.weekly9355 -Algorithm SHA256).Hash
    $conflictBlocked = $false
    try {
        & $migrationScript -OperationsRoot $fixture -Apply -ManifestDirectory $manifestDirectory -ExpectedPlanSha256 ("0" * 64)
    }
    catch {
        $conflictBlocked = $_.Exception.Message -like "*conflicting destination content*"
    }
    Assert-True $conflictBlocked "conflicting destination must block preflight"
    Assert-True ((Get-FileHash -LiteralPath $files.weekly9355 -Algorithm SHA256).Hash -eq $unrelatedHash) "conflict preflight makes no unrelated changes"
    Remove-Item -LiteralPath $conflictPath -Force

    # A directory occupying a mapped file destination is also a complete
    # preflight conflict, not a late publication error.
    [void][System.IO.Directory]::CreateDirectory($conflictPath)
    $directoryConflictBlocked = $false
    try {
        & $migrationScript -OperationsRoot $fixture -Apply -ManifestDirectory $manifestDirectory -ExpectedPlanSha256 ("0" * 64)
    }
    catch {
        $directoryConflictBlocked = $_.Exception.Message -like "*conflicting destination content*"
    }
    Assert-True $directoryConflictBlocked "destination directory must block preflight"
    Assert-True ((Get-FileHash -LiteralPath $files.weekly9355 -Algorithm SHA256).Hash -eq $unrelatedHash) "directory conflict preflight makes no unrelated changes"
    Remove-Item -LiteralPath $conflictPath -Force

    # An exclusively locked source blocks before any migration changes.
    $lockDryJson = & $migrationScript -OperationsRoot $fixture
    $lockDry = $lockDryJson | ConvertFrom-Json
    $lock = [System.IO.File]::Open($files.controls9354, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $lockBlocked = $false
    try {
        try {
            & $migrationScript -OperationsRoot $fixture -Apply -ManifestDirectory $manifestDirectory -ExpectedPlanSha256 $lockDry.plan_sha256
        }
        catch {
            $lockBlocked = $_.Exception.Message -like "*locked or not stably readable*"
        }
    }
    finally {
        $lock.Dispose()
    }
    Assert-True $lockBlocked "locked source must block preflight"
    Assert-True (Test-Path -LiteralPath $files.controls9354 -PathType Leaf) "locked source remains"

    # A hard interruption after only the immutable preflight write or after one
    # completed move resumes that plan with the original reviewed fingerprint.
    Invoke-ApplyResumeFixture `
        -Root (Join-Path $fixture "resume-after-preflight") `
        -MoveOneFileBeforeResume $false `
        -WritePostBeforeResume $false `
        -MigrationScript $migrationScript
    Invoke-ApplyResumeFixture `
        -Root (Join-Path $fixture "resume-after-one-move") `
        -MoveOneFileBeforeResume $true `
        -WritePostBeforeResume $true `
        -MigrationScript $migrationScript

    Write-Host "PASS: numbered-layout migration dry-run, apply, resume, verify, idempotency, rollback, conflict, and lock tests."
}
finally {
    $fixtureFull = [System.IO.Path]::GetFullPath($fixture)
    $requiredPrefix = $tempRoot + [System.IO.Path]::DirectorySeparatorChar + "GiftCardLayoutMigrationTest-"
    if ((Test-Path -LiteralPath $fixtureFull) -and $fixtureFull.StartsWith($requiredPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $fixtureFull -Recurse -Force
    }
}
