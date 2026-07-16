$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw "ASSERTION FAILED: $Message"
    }
}

function Invoke-NativeChecked {
    param([string]$FilePath, [string[]]$Arguments)
    & $FilePath @Arguments | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath $($Arguments -join ' ') failed with exit code $LASTEXITCODE."
    }
}

$MaintenanceRoot = Split-Path -Parent $PSCommandPath
$ProgramRoot = Split-Path -Parent $MaintenanceRoot
$ProjectRoot = Split-Path -Parent $ProgramRoot
$TestBase = Join-Path ([IO.Path]::GetTempPath()) "gift-card-operator-tooling-tests"
$TestRoot = Join-Path $TestBase ([guid]::NewGuid().ToString("N"))

try {
    New-Item -ItemType Directory -Force -Path $TestRoot | Out-Null

    . (Join-Path $ProgramRoot "runtime.ps1")
    $operator = Get-GiftCardReconRuntime -Profile Operator
    $development = Get-GiftCardReconRuntime -Profile Development
    Assert-True ($operator.RuntimeRoot -ne $development.RuntimeRoot) "operator and development runtimes must use different roots"
    Assert-True ($operator.RuntimeRoot -like "*\GiftCardRecon\operator") "operator runtime path should be stable"
    Assert-True ($development.RuntimeRoot -like "*\GiftCardRecon\development") "development runtime path should be isolated"

    $programCopyA = Join-Path $TestRoot "program-a"
    $programCopyB = Join-Path $TestRoot "program-b"
    foreach ($copy in @($programCopyA, $programCopyB)) {
        New-Item -ItemType Directory -Force -Path $copy | Out-Null
        Copy-Item -LiteralPath (Join-Path $ProgramRoot "requirements.txt") -Destination $copy
        Copy-Item -LiteralPath (Join-Path $ProgramRoot "pyproject.toml") -Destination $copy
        Copy-Item -LiteralPath (Join-Path $ProgramRoot "src") -Destination $copy -Recurse
        @(Get-ChildItem -LiteralPath (Join-Path $copy "src") -Directory -Recurse -Force |
            Where-Object { $_.Name -eq "__pycache__" -or $_.Name -like "*.egg-info" } |
            Sort-Object { $_.FullName.Length } -Descending) |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }
    }
    $ignoredEggInfo = Join-Path $programCopyB "src\gift_card_recon.egg-info"
    $ignoredPycache = Join-Path $programCopyB "src\gift_card_recon\__pycache__"
    New-Item -ItemType Directory -Force -Path $ignoredEggInfo, $ignoredPycache | Out-Null
    Set-Content -LiteralPath (Join-Path $ignoredEggInfo "PKG-INFO") -Value "ignored build metadata"
    Set-Content -LiteralPath (Join-Path $ignoredPycache "ignored.pyc") -Value "ignored bytecode"

    $operatorA = Get-GiftCardReconDependencyFingerprint -ProgramRoot $programCopyA -Profile Operator
    $operatorB = Get-GiftCardReconDependencyFingerprint -ProgramRoot $programCopyB -Profile Operator
    $developmentA = Get-GiftCardReconDependencyFingerprint -ProgramRoot $programCopyA -Profile Development
    $developmentB = Get-GiftCardReconDependencyFingerprint -ProgramRoot $programCopyB -Profile Development
    Assert-True ($operatorA -eq $operatorB) "identical operator payloads must have path-independent fingerprints"
    Assert-True ($developmentA -ne $developmentB) "editable development fingerprints must remain checkout-specific"

    $fakePython = Join-Path $TestRoot "fake-python.cmd"
    $stagedFileCapture = Join-Path $TestRoot "staged-files.txt"
    Set-Content -LiteralPath $fakePython -Value @(
        "@echo off",
        "set `"last=`"",
        "for %%A in (%*) do set `"last=%%~A`"",
        "dir /b /s `"%last%`" > `"$stagedFileCapture`"",
        "exit /b 0"
    ) -Encoding ASCII
    $fakeRuntime = [pscustomobject]@{
        PythonPath = $fakePython
        TempRoot = Join-Path $TestRoot "operator-temp"
    }
    $beforeInstall = @(Get-ChildItem -LiteralPath $programCopyB -File -Recurse | Sort-Object FullName | ForEach-Object {
        "$($_.FullName.Substring($programCopyB.Length))=$((Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash)"
    }) -join "`n"
    Install-GiftCardReconOperatorPackage -Runtime $fakeRuntime -ProgramRoot $programCopyB
    $afterInstall = @(Get-ChildItem -LiteralPath $programCopyB -File -Recurse | Sort-Object FullName | ForEach-Object {
        "$($_.FullName.Substring($programCopyB.Length))=$((Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash)"
    }) -join "`n"
    Assert-True ($beforeInstall -eq $afterInstall) "operator package installation must not mutate its deployed source"
    $capturedStage = Get-Content -LiteralPath $stagedFileCapture -Raw
    Assert-True ($capturedStage -notmatch "\.egg-info|__pycache__|\.pyc") "operator package staging must exclude build metadata and bytecode"
    Assert-True (@(Get-ChildItem -LiteralPath $fakeRuntime.TempRoot -Force -ErrorAction SilentlyContinue).Count -eq 0) "operator package staging should be removed"

    $assetOperations = Join-Path $TestRoot "asset-transaction"
    New-Item -ItemType Directory -Force -Path $assetOperations | Out-Null
    $guideTarget = Join-Path $assetOperations "00 START HERE - Gift Card Reconciliation.txt"
    Set-Content -LiteralPath $guideTarget -Value "prior operator guide" -NoNewline
    $invalidLateTarget = Join-Path $assetOperations "02 Monthly Close Inputs\Darden Reports - Drop Here\00 DROP DARDEN CREDIT MEMOS HERE.txt"
    New-Item -ItemType Directory -Force -Path $invalidLateTarget | Out-Null
    $assetPreflightBlocked = $false
    $assetPreflightMessage = "no error was raised"
    try {
        & (Join-Path $ProgramRoot "install_operator_assets.ps1") -OperationsRoot $assetOperations | Out-Null
    }
    catch {
        $assetPreflightMessage = $_.Exception.Message
        $assetPreflightBlocked = $assetPreflightMessage -like "*target exists but is not a file*"
    }
    Assert-True $assetPreflightBlocked "operator asset refresh must reject an invalid late target; received: $assetPreflightMessage"
    Assert-True ((Get-Content -LiteralPath $guideTarget -Raw) -eq "prior operator guide") "asset preflight failure must preserve every live operator file"
    $assetResidue = @(Get-ChildItem -LiteralPath $assetOperations -File -Recurse -Force | Where-Object { $_.Name -like ".gcs-*.tmp" -or $_.Name -like ".gcb-*.tmp" })
    Assert-True ($assetResidue.Count -eq 0) "failed operator asset refresh must not leave transaction residue"
    Remove-Item -LiteralPath $invalidLateTarget -Recurse -Force
    & (Join-Path $ProgramRoot "install_operator_assets.ps1") -OperationsRoot $assetOperations | Out-Null
    $expectedGuideHash = (Get-FileHash -LiteralPath (Join-Path $ProjectRoot "templates\00 START HERE - Gift Card Reconciliation.txt") -Algorithm SHA256).Hash
    Assert-True ((Get-FileHash -LiteralPath $guideTarget -Algorithm SHA256).Hash -eq $expectedGuideHash) "successful asset refresh must publish the verified guide"

    $source = Join-Path $TestRoot "source"
    $operations = Join-Path $TestRoot "operations"
    $target = Join-Path $operations "Gift Card Reconciliation Automation"
    New-Item -ItemType Directory -Force -Path $source, $operations, (Join-Path $target ".git"), (Join-Path $target "__pycache__") | Out-Null
    $fixtureFiles = @(
        "_program/install.ps1",
        "_program/install_operator_assets.ps1",
        "_program/run_weekly.ps1",
        "_program/run_monthly_close.ps1",
        "_program/runtime.ps1",
        "_program/requirements.txt",
        "_program/pyproject.toml",
        "_program/check_operator_health.ps1",
        "_program/src/gift_card_recon/__init__.py",
        "templates/Run Test.cmd"
    )
    foreach ($relative in $fixtureFiles) {
        $path = Join-Path $source $relative
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
        if ($relative -eq "_program/check_operator_health.ps1") {
            Copy-Item -LiteralPath (Join-Path $ProgramRoot "check_operator_health.ps1") -Destination $path
        }
        else {
            Set-Content -LiteralPath $path -Value "verified deployment: $relative" -NoNewline
        }
    }
    Set-Content -LiteralPath (Join-Path $source "README.md") -Value "development-only file" -NoNewline
    Set-Content -LiteralPath (Join-Path $target "__pycache__\stale.pyc") -Value "stale" -NoNewline
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "init", "-q")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "config", "user.name", "GC Recon Test")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "config", "user.email", "gc-recon-test@example.invalid")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "remote", "add", "origin", "https://example.invalid/gift-card-recon.git")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "add", ".")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "commit", "-q", "-m", "test")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "branch", "-M", "main")

    $deployment = & (Join-Path $MaintenanceRoot "deploy_operator_program.ps1") `
        -SourceRoot $source `
        -OperationsRoot $operations `
        -SkipUpstreamCheck `
        -SkipOperatorAssetRefresh
    Assert-True (Test-Path -LiteralPath (Join-Path $target "_program\runtime.ps1") -PathType Leaf) "tracked operator source file should be deployed"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $target "README.md"))) "development-only tracked files should be excluded"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $target ".git"))) "deployed Git metadata should be removed"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $target "__pycache__"))) "deployed cache should be removed"
    $manifestPath = Join-Path $target "deployment-manifest.json"
    Assert-True (Test-Path -LiteralPath $manifestPath -PathType Leaf) "deployment manifest should be published last"
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    Assert-True ($manifest.commit -eq $deployment.Commit) "manifest should record the deployed commit"
    Assert-True ([int]$manifest.file_count -eq $fixtureFiles.Count) "manifest should record each deployed operator file"
    $runtimeEntry = @($manifest.files | Where-Object { $_.path -eq "_program/runtime.ps1" })[0]
    Assert-True ((Get-FileHash -LiteralPath (Join-Path $target "_program\runtime.ps1") -Algorithm SHA256).Hash.ToLowerInvariant() -eq $runtimeEntry.sha256) "deployed file hash should match manifest"

    $priorCommit = [string]$manifest.commit
    $priorRuntimeHash = (Get-FileHash -LiteralPath (Join-Path $target "_program\runtime.ps1") -Algorithm SHA256).Hash
    Add-Content -LiteralPath (Join-Path $source "_program\runtime.ps1") -Value "`nnew version"
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "add", "_program/runtime.ps1")
    Invoke-NativeChecked -FilePath "git" -Arguments @("-C", $source, "commit", "-q", "-m", "deployment that must roll back")
    $rollbackTriggered = $false
    try {
        & (Join-Path $MaintenanceRoot "deploy_operator_program.ps1") `
            -SourceRoot $source `
            -OperationsRoot $operations `
            -SkipUpstreamCheck | Out-Null
    }
    catch {
        $rollbackTriggered = $_.Exception.Message -like "*prior operator program was restored*"
    }
    Assert-True $rollbackTriggered "a post-publication failure should trigger rollback"
    $restoredManifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    Assert-True ($restoredManifest.commit -eq $priorCommit) "rollback should restore the prior manifest"
    Assert-True ((Get-FileHash -LiteralPath (Join-Path $target "_program\runtime.ps1") -Algorithm SHA256).Hash -eq $priorRuntimeHash) "rollback should restore the prior program payload"
    $transactionResidue = @(Get-ChildItem -LiteralPath $operations -Force | Where-Object { $_.Name -like ".Gift Card Reconciliation Automation.*-*" })
    Assert-True ($transactionResidue.Count -eq 0) "transactional deployment should not leave staging or backup folders"

    Set-Content -LiteralPath (Join-Path $source "dirty.txt") -Value "not committed"
    $dirtyBlocked = $false
    try {
        & (Join-Path $MaintenanceRoot "deploy_operator_program.ps1") `
            -SourceRoot $source `
            -OperationsRoot $operations `
            -SkipUpstreamCheck `
            -SkipOperatorAssetRefresh | Out-Null
    }
    catch {
        $dirtyBlocked = $_.Exception.Message -like "*working tree is not clean*"
    }
    Assert-True $dirtyBlocked "deployment must reject a dirty source working tree"

    $report = Join-Path $TestRoot "health.json"
    $shell = (Get-Command powershell.exe -ErrorAction Stop).Source
    & $shell -NoLogo -NoProfile -ExecutionPolicy Bypass -File (Join-Path $target "_program\check_operator_health.ps1") `
        -OperationsRoot $operations `
        -DropboxRoot $TestRoot `
        -ReportPath $report *> $null
    Assert-True ($LASTEXITCODE -eq 2) "health check should return 2 when required evidence is missing"
    $health = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json
    Assert-True ($health.overall_status -eq "BLOCKED") "blocked health result should be recorded"
    Assert-True ([int]$health.blocker_count -gt 0) "health report should include blockers"
    $manifestHealth = @($health.controls | Where-Object { $_.control -eq "Deployment manifest" })
    Assert-True ($manifestHealth.Count -eq 1 -and $manifestHealth[0].status -eq "PASS") "deployed health check must independently verify the manifest tree hash"

    $launcher = Get-Content -LiteralPath (Join-Path $ProjectRoot "templates\Check Gift Card Reconciliation Health.cmd") -Raw
    Assert-True ($launcher.Contains('set "OPERATIONS_ROOT=%~dp0."')) "health launcher should resolve the operator root"
    Assert-True ($launcher.Contains("_program\check_operator_health.ps1")) "health launcher should call the checker"
    Assert-True ($launcher.Contains("exit /b %EXITCODE%")) "health launcher should propagate failures"

    Write-Host "Operator tooling tests passed." -ForegroundColor Green
    $global:LASTEXITCODE = 0
}
finally {
    if (Test-Path -LiteralPath $TestRoot) {
        $resolvedBase = [IO.Path]::GetFullPath($TestBase).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        $resolvedTest = [IO.Path]::GetFullPath($TestRoot)
        if (-not $resolvedTest.StartsWith($resolvedBase, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing unsafe test cleanup: $resolvedTest"
        }
        Remove-Item -LiteralPath $resolvedTest -Recurse -Force -ErrorAction SilentlyContinue
    }
}
