[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ProgramRoot = Split-Path -Parent $PSScriptRoot
$ProjectRoot = Split-Path -Parent $ProgramRoot
$Installer = Join-Path $ProgramRoot "install_operator_assets.ps1"
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "ASSERTION FAILED: $Message"
    }
}

function Get-ManagedAssetTemplates {
    return [ordered]@{
        "00 START HERE - Gift Card Reconciliation.txt" = Join-Path $ProjectRoot "templates\00 START HERE - Gift Card Reconciliation.txt"
        "Run Weekly Gift Card Reconciliation.cmd" = Join-Path $ProjectRoot "templates\Run Weekly Gift Card Reconciliation.cmd"
        "Run Monthly Gift Card Close.cmd" = Join-Path $ProjectRoot "templates\Run Monthly Gift Card Close.cmd"
        "01 Weekly Gift Card Activity Reports\9354 Richmond\activity\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt" = Join-Path $ProjectRoot "templates\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"
        "01 Weekly Gift Card Activity Reports\9355 Virginia Beach\activity\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt" = Join-Path $ProjectRoot "templates\00 DROP WEEKLY GIFT CARD ACTIVITY HERE.txt"
        "02 Monthly Close Inputs\Darden Reports - Drop Here\00 DROP DARDEN CREDIT MEMOS HERE.txt" = Join-Path $ProjectRoot "templates\00 DROP DARDEN CREDIT MEMOS HERE.txt"
    }
}

function Write-TextFile {
    param(
        [string]$Path,
        [string]$Content
    )

    $Parent = Split-Path -Parent $Path
    New-Item -ItemType Directory -Force -Path $Parent | Out-Null
    [System.IO.File]::WriteAllText($Path, $Content, $Utf8NoBom)
}

function Write-LegacyHealthLauncher {
    param([string]$Path)

    $Lines = @(
        "@echo off",
        "setlocal",
        "title Gift Card Reconciliation Health Check",
        'set "OPERATIONS_ROOT=%~dp0."',
        'set "PROGRAM_ROOT=%~dp0Gift Card Reconciliation Automation"',
        'set "CHECKER=%PROGRAM_ROOT%\_program\check_operator_health.ps1"',
        "",
        "echo.",
        'if not exist "%CHECKER%" (',
        "  echo ATTENTION NEEDED: The operator health checker was not found:",
        "  echo %CHECKER%",
        "  echo Deploy the current program from the clean local Git checkout.",
        "  echo.",
        "  pause",
        "  exit /b 2",
        ")",
        "",
        "where pwsh >nul 2>&1",
        "if errorlevel 1 (",
        '  powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CHECKER%" -OperationsRoot "%OPERATIONS_ROOT%" %*',
        ") else (",
        '  pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CHECKER%" -OperationsRoot "%OPERATIONS_ROOT%" %*',
        ")",
        'set "EXITCODE=%ERRORLEVEL%"',
        "",
        "echo.",
        'if "%EXITCODE%"=="0" (',
        "  echo The Gift Card Reconciliation operator environment is ready.",
        ") else (",
        "  echo ATTENTION NEEDED: One or more health controls are blocked.",
        ")",
        "echo.",
        "pause",
        "exit /b %EXITCODE%"
    )
    Write-TextFile -Path $Path -Content (($Lines -join "`r`n") + "`r`n")
    $Hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    Assert-True ($Hash -eq "3614BEA1DCA286BC14EAA9116853FE4CFE40B50CF6E696FC284AD232802DB1EA") `
        "legacy launcher fixture must match the released CRLF fingerprint"
}

function Initialize-PriorAssetSet {
    param([string]$OperationsRoot)

    $Before = @{}
    $Index = 0
    foreach ($RelativePath in (Get-ManagedAssetTemplates).Keys) {
        $Target = Join-Path $OperationsRoot $RelativePath
        Write-TextFile -Path $Target -Content "prior operator asset $Index for $RelativePath`r`n"
        $Before[$RelativePath] = (Get-FileHash -LiteralPath $Target -Algorithm SHA256).Hash
        $Index++
    }
    return $Before
}

function Assert-ManagedAssetsMatchTemplates {
    param([string]$OperationsRoot)

    $Templates = Get-ManagedAssetTemplates
    foreach ($RelativePath in $Templates.Keys) {
        $Target = Join-Path $OperationsRoot $RelativePath
        Assert-True (Test-Path -LiteralPath $Target -PathType Leaf) "managed asset exists: $RelativePath"
        $ExpectedHash = (Get-FileHash -LiteralPath $Templates[$RelativePath] -Algorithm SHA256).Hash
        $ActualHash = (Get-FileHash -LiteralPath $Target -Algorithm SHA256).Hash
        Assert-True ($ActualHash -eq $ExpectedHash) "managed asset matches template: $RelativePath"
    }
}

function Assert-ManagedAssetsMatchSnapshot {
    param(
        [string]$OperationsRoot,
        [hashtable]$Snapshot
    )

    foreach ($RelativePath in $Snapshot.Keys) {
        $Target = Join-Path $OperationsRoot $RelativePath
        Assert-True (Test-Path -LiteralPath $Target -PathType Leaf) "prior asset restored: $RelativePath"
        $ActualHash = (Get-FileHash -LiteralPath $Target -Algorithm SHA256).Hash
        Assert-True ($ActualHash -eq $Snapshot[$RelativePath]) "prior asset hash restored: $RelativePath"
    }
}

function Assert-NoTransactionResidue {
    param([string]$OperationsRoot)

    $Residue = @(
        Get-ChildItem -LiteralPath $OperationsRoot -Recurse -Force -File |
            Where-Object { $_.Name -like ".gcs-*.tmp" -or $_.Name -like ".gcb-*.tmp" }
    )
    Assert-True ($Residue.Count -eq 0) "operator asset transaction leaves no stage or backup files"
}

$TestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("gift-card-operator-assets-{0}" -f [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $TestRoot | Out-Null

try {
    $SuccessRoot = Join-Path $TestRoot "success"
    New-Item -ItemType Directory -Path $SuccessRoot | Out-Null
    [void](Initialize-PriorAssetSet -OperationsRoot $SuccessRoot)
    $UnrelatedFile = Join-Path $SuccessRoot "Operator Notes.txt"
    Write-TextFile -Path $UnrelatedFile -Content "operator-owned notes must remain unchanged`r`n"
    $UnrelatedHash = (Get-FileHash -LiteralPath $UnrelatedFile -Algorithm SHA256).Hash
    $HealthLauncher = Join-Path $SuccessRoot "Check Gift Card Reconciliation Health.cmd"
    Write-LegacyHealthLauncher -Path $HealthLauncher

    & $Installer -OperationsRoot $SuccessRoot | Out-Null

    Assert-ManagedAssetsMatchTemplates -OperationsRoot $SuccessRoot
    Assert-True (-not (Test-Path -LiteralPath $HealthLauncher)) "released stale health launcher is retired"
    Assert-True ((Get-FileHash -LiteralPath $UnrelatedFile -Algorithm SHA256).Hash -eq $UnrelatedHash) `
        "unrelated operator file is preserved"
    Assert-NoTransactionResidue -OperationsRoot $SuccessRoot

    Write-TextFile -Path $HealthLauncher -Content "operator-owned same-name file`r`n"
    $CustomHealthHash = (Get-FileHash -LiteralPath $HealthLauncher -Algorithm SHA256).Hash
    & $Installer -OperationsRoot $SuccessRoot | Out-Null
    Assert-True ((Get-FileHash -LiteralPath $HealthLauncher -Algorithm SHA256).Hash -eq $CustomHealthHash) `
        "unrecognized same-name health file is preserved"
    Assert-NoTransactionResidue -OperationsRoot $SuccessRoot

    $FailureRoot = Join-Path $TestRoot "late-failure"
    New-Item -ItemType Directory -Path $FailureRoot | Out-Null
    $PriorSnapshot = Initialize-PriorAssetSet -OperationsRoot $FailureRoot
    $FailureUnrelated = Join-Path $FailureRoot "Operator Notes.txt"
    Write-TextFile -Path $FailureUnrelated -Content "failure fixture unrelated file`r`n"
    $FailureUnrelatedHash = (Get-FileHash -LiteralPath $FailureUnrelated -Algorithm SHA256).Hash
    $LockedHealthLauncher = Join-Path $FailureRoot "Check Gift Card Reconciliation Health.cmd"
    Write-LegacyHealthLauncher -Path $LockedHealthLauncher
    $LockedHealthHash = (Get-FileHash -LiteralPath $LockedHealthLauncher -Algorithm SHA256).Hash

    $Lock = [System.IO.File]::Open(
        $LockedHealthLauncher,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )
    $FailureCaught = $false
    $FailureMessage = ""
    try {
        & $Installer -OperationsRoot $FailureRoot | Out-Null
    }
    catch {
        $FailureCaught = $true
        $FailureMessage = $_.Exception.Message
    }
    finally {
        $Lock.Dispose()
    }

    Assert-True $FailureCaught "locked late retirement must fail the refresh"
    Assert-True ($FailureMessage -like "*prior operator asset set was restored*") `
        "late failure must report successful rollback; received: $FailureMessage"
    Assert-ManagedAssetsMatchSnapshot -OperationsRoot $FailureRoot -Snapshot $PriorSnapshot
    Assert-True ((Get-FileHash -LiteralPath $LockedHealthLauncher -Algorithm SHA256).Hash -eq $LockedHealthHash) `
        "late failure restores the retired launcher"
    Assert-True ((Get-FileHash -LiteralPath $FailureUnrelated -Algorithm SHA256).Hash -eq $FailureUnrelatedHash) `
        "late failure preserves unrelated files"
    Assert-NoTransactionResidue -OperationsRoot $FailureRoot

    Write-Host "PASS: transactional operator asset refresh, safe launcher retirement, and late-failure rollback."
}
finally {
    $ExpectedPrefix = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    $ResolvedTestRoot = [System.IO.Path]::GetFullPath($TestRoot)
    if ($ResolvedTestRoot.StartsWith($ExpectedPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $ResolvedTestRoot) -like "gift-card-operator-assets-*") {
        Remove-Item -LiteralPath $ResolvedTestRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
