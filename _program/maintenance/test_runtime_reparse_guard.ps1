#Requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\runtime.ps1")

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Invoke-VenvGuard {
    param([Parameter(Mandatory = $true)][pscustomobject]$Runtime)

    try {
        Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
        return [pscustomobject]@{ Blocked = $false; Message = "" }
    } catch {
        return [pscustomobject]@{ Blocked = $true; Message = $_.Exception.Message }
    }
}

function Get-ChildPathSnapshot {
    param([Parameter(Mandatory = $true)][string]$Root)

    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    return @(
        Get-ChildItem -LiteralPath $fullRoot -Force -Recurse |
            ForEach-Object { $_.FullName.Substring($fullRoot.Length).TrimStart('\', '/') } |
            Sort-Object
    )
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw "The runtime reparse guard fixture requires Windows junction support."
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$fixture = Join-Path $tempRoot ("GiftCardRuntimeReparseTest-{0}" -f [Guid]::NewGuid().ToString("N"))
$junctions = [Collections.Generic.List[string]]::new()
$originalLocalAppData = $env:LOCALAPPDATA

try {
    [void][IO.Directory]::CreateDirectory($fixture)

    # A normal LocalAppData runtime is accepted both before and after the
    # ordinary venv directory exists.
    $normalLocalAppData = Join-Path $fixture "NormalLocalAppData"
    [void][IO.Directory]::CreateDirectory($normalLocalAppData)
    $env:LOCALAPPDATA = $normalLocalAppData
    $normalRuntime = Get-GiftCardReconRuntime
    $normalMissing = Invoke-VenvGuard -Runtime $normalRuntime
    Assert-True (-not $normalMissing.Blocked) "normal missing venv path should be accepted"
    [void][IO.Directory]::CreateDirectory($normalRuntime.VenvRoot)
    $normalExisting = Invoke-VenvGuard -Runtime $normalRuntime
    Assert-True (-not $normalExisting.Blocked) "normal existing venv directory should be accepted"
    Assert-True (
        $normalRuntime.RuntimeRoot -eq (Join-Path $normalLocalAppData "GiftCardRecon")
    ) "normal runtime remains rooted under LocalAppData"

    # Reproduce the review bypass: RuntimeRoot is a junction while venv is an
    # ordinary child in the junction target.
    $parentLocalAppData = Join-Path $fixture "ParentReparseLocalAppData"
    $externalRuntime = Join-Path $fixture "ExternalRuntime"
    $parentJunction = Join-Path $parentLocalAppData "GiftCardRecon"
    [void][IO.Directory]::CreateDirectory($parentLocalAppData)
    [void][IO.Directory]::CreateDirectory((Join-Path $externalRuntime "venv"))
    $sentinel = Join-Path $externalRuntime "venv\sentinel.txt"
    [IO.File]::WriteAllText($sentinel, "must remain")
    New-Item -ItemType Junction -Path $parentJunction -Target $externalRuntime | Out-Null
    $junctions.Add($parentJunction)
    $parentRuntime = [pscustomobject]@{
        RuntimeRoot = $parentJunction
        VenvRoot = Join-Path $parentJunction "venv"
    }
    $parentResult = Invoke-VenvGuard -Runtime $parentRuntime
    Assert-True $parentResult.Blocked "a reparse parent must be rejected"
    Assert-True ($parentResult.Message -like "*$parentJunction*") "the blocked parent must be identified"
    Assert-True (Test-Path -LiteralPath $sentinel -PathType Leaf) "guard must leave target contents untouched"

    # An ancestor above RuntimeRoot is equally unsafe because venv --clear
    # follows every redirected component in the path.
    $ancestorContainer = Join-Path $fixture "AncestorReparseContainer"
    $externalLocalAppData = Join-Path $fixture "ExternalLocalAppData"
    $localAppDataJunction = Join-Path $ancestorContainer "LocalAppData"
    [void][IO.Directory]::CreateDirectory($ancestorContainer)
    [void][IO.Directory]::CreateDirectory((Join-Path $externalLocalAppData "GiftCardRecon\venv"))
    New-Item -ItemType Junction -Path $localAppDataJunction -Target $externalLocalAppData | Out-Null
    $junctions.Add($localAppDataJunction)
    $ancestorRuntime = [pscustomobject]@{
        RuntimeRoot = Join-Path $localAppDataJunction "GiftCardRecon"
        VenvRoot = Join-Path $localAppDataJunction "GiftCardRecon\venv"
    }
    $ancestorResult = Invoke-VenvGuard -Runtime $ancestorRuntime
    Assert-True $ancestorResult.Blocked "a reparse ancestor above RuntimeRoot must be rejected"
    Assert-True (
        $ancestorResult.Message -like "*$localAppDataJunction*"
    ) "the blocked ancestor must be identified"

    # The leaf entry remains protected as an alternate malicious path class.
    $leafRuntimeRoot = Join-Path $fixture "LeafReparseRuntime"
    $leafTarget = Join-Path $fixture "LeafTarget"
    $leafJunction = Join-Path $leafRuntimeRoot "venv"
    [void][IO.Directory]::CreateDirectory($leafRuntimeRoot)
    [void][IO.Directory]::CreateDirectory($leafTarget)
    New-Item -ItemType Junction -Path $leafJunction -Target $leafTarget | Out-Null
    $junctions.Add($leafJunction)
    $leafRuntime = [pscustomobject]@{ RuntimeRoot = $leafRuntimeRoot; VenvRoot = $leafJunction }
    $leafResult = Invoke-VenvGuard -Runtime $leafRuntime
    Assert-True $leafResult.Blocked "a reparse venv leaf must be rejected"
    Assert-True ($leafResult.Message -like "*$leafJunction*") "the blocked leaf must be identified"

    # Exercise the actual initialization decision tree. These cases use a real
    # venv junction but replace process execution with a mutation canary, so a
    # missed guard would visibly alter the external sentinel before failing.
    $script:fixtureRuntime = $null
    $script:fixtureRuntimeValid = $false
    $script:fixturePythonUsable = $false
    $script:fixtureInstallCommandInvoked = $false
    $script:fixtureSentinel = ""

    function Get-GiftCardReconRuntime { return $script:fixtureRuntime }
    function Get-GiftCardReconDependencyFingerprint {
        param([Parameter(Mandatory = $true)][string]$ProgramRoot)
        return "fixture-fingerprint"
    }
    function Test-GiftCardReconRuntime {
        param([Parameter(Mandatory = $true)][pscustomobject]$Runtime)
        return [bool]$script:fixtureRuntimeValid
    }
    function Test-GiftCardReconPython {
        param([Parameter(Mandatory = $true)][pscustomobject]$Runtime)
        return [bool]$script:fixturePythonUsable
    }
    function Invoke-GiftCardReconChecked {
        param(
            [Parameter(Mandatory = $true)][string]$FilePath,
            [Parameter(Mandatory = $true)][string[]]$Arguments
        )
        $script:fixtureInstallCommandInvoked = $true
        [IO.File]::WriteAllText($script:fixtureSentinel, "modified through junction")
        throw "INSTALL MUTATION CANARY WAS REACHED"
    }

    $programRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
    foreach ($installCase in @(
        [pscustomobject]@{ Name = "create"; RuntimeValid = $false; PythonUsable = $false; Force = $false },
        [pscustomobject]@{ Name = "repair"; RuntimeValid = $false; PythonUsable = $true; Force = $false },
        [pscustomobject]@{ Name = "refresh"; RuntimeValid = $true; PythonUsable = $true; Force = $true }
    )) {
        $caseRoot = Join-Path $fixture ("InstallCase-{0}" -f $installCase.Name)
        $caseLocalAppData = Join-Path $caseRoot "LocalAppData"
        $caseRuntimeRoot = Join-Path $caseLocalAppData "GiftCardRecon"
        $caseExternalVenv = Join-Path $caseRoot "ExternalVenv"
        $caseVenvJunction = Join-Path $caseRuntimeRoot "venv"
        [void][IO.Directory]::CreateDirectory($caseRuntimeRoot)
        [void][IO.Directory]::CreateDirectory((Join-Path $caseExternalVenv "Scripts"))
        [IO.File]::WriteAllText((Join-Path $caseExternalVenv "pyvenv.cfg"), "fixture")
        [IO.File]::WriteAllText((Join-Path $caseExternalVenv "Scripts\python.exe"), "fixture")
        $caseSentinel = Join-Path $caseExternalVenv "sentinel.txt"
        [IO.File]::WriteAllText($caseSentinel, "must remain")
        New-Item -ItemType Junction -Path $caseVenvJunction -Target $caseExternalVenv | Out-Null
        $junctions.Add($caseVenvJunction)

        $env:LOCALAPPDATA = $caseLocalAppData
        $script:fixtureRuntime = [pscustomobject]@{
            RuntimeRoot = $caseRuntimeRoot
            VenvRoot = $caseVenvJunction
            PythonPath = Join-Path $caseVenvJunction "Scripts\python.exe"
            CacheRoot = Join-Path $caseRuntimeRoot "cache"
            PipCacheDir = Join-Path $caseRuntimeRoot "cache\pip"
            PycacheDir = Join-Path $caseRuntimeRoot "cache\pycache"
            PytestCacheDir = Join-Path $caseRuntimeRoot "cache\pytest"
            TempRoot = Join-Path $caseRuntimeRoot "temp"
            MicrosExtractDir = Join-Path $caseRuntimeRoot "temp\micros-extract"
            DependencyFingerprintPath = Join-Path $caseRuntimeRoot "dependency-fingerprint.sha256"
        }
        $script:fixtureRuntimeValid = $installCase.RuntimeValid
        $script:fixturePythonUsable = $installCase.PythonUsable
        $script:fixtureInstallCommandInvoked = $false
        $script:fixtureSentinel = $caseSentinel
        $runtimeBefore = Get-ChildPathSnapshot -Root $caseRuntimeRoot
        $externalBefore = Get-ChildPathSnapshot -Root $caseExternalVenv

        $blockedMessage = ""
        try {
            Invoke-GiftCardReconRuntimeInitialization `
                -ProgramRoot $programRoot `
                -ForceInstall:$installCase.Force | Out-Null
            throw "the $($installCase.Name) path unexpectedly accepted a reparse-point venv"
        } catch {
            $blockedMessage = $_.Exception.Message
        }

        Assert-True (
            $blockedMessage -like "*link, junction, or other reparse point*"
        ) "the $($installCase.Name) path must fail at the reparse guard"
        Assert-True (
            $blockedMessage -like "*$caseVenvJunction*"
        ) "the $($installCase.Name) error must identify the venv junction"
        Assert-True (
            -not $script:fixtureInstallCommandInvoked
        ) "the $($installCase.Name) path must block before an install command"
        Assert-True (
            [IO.File]::ReadAllText($caseSentinel) -eq "must remain"
        ) "the $($installCase.Name) path must preserve the external sentinel"
        Assert-True (
            (($runtimeBefore -join "`n") -eq ((Get-ChildPathSnapshot -Root $caseRuntimeRoot) -join "`n"))
        ) "the $($installCase.Name) path must not add runtime directories before blocking"
        Assert-True (
            (($externalBefore -join "`n") -eq ((Get-ChildPathSnapshot -Root $caseExternalVenv) -join "`n"))
        ) "the $($installCase.Name) path must not modify the external venv tree"
    }

    # -SkipInstall remains a no-write preflight: it reports the existing
    # out-of-date-runtime error without creating runtime directories or
    # reaching any install command.
    $skipCaseRoot = Join-Path $fixture "SkipInstallCase"
    $skipLocalAppData = Join-Path $skipCaseRoot "LocalAppData"
    $skipRuntimeRoot = Join-Path $skipLocalAppData "GiftCardRecon"
    $skipExternalVenv = Join-Path $skipCaseRoot "ExternalVenv"
    $skipVenvJunction = Join-Path $skipRuntimeRoot "venv"
    [void][IO.Directory]::CreateDirectory($skipRuntimeRoot)
    [void][IO.Directory]::CreateDirectory($skipExternalVenv)
    $skipSentinel = Join-Path $skipExternalVenv "sentinel.txt"
    [IO.File]::WriteAllText($skipSentinel, "must remain")
    New-Item -ItemType Junction -Path $skipVenvJunction -Target $skipExternalVenv | Out-Null
    $junctions.Add($skipVenvJunction)
    $script:fixtureRuntime = [pscustomobject]@{
        RuntimeRoot = $skipRuntimeRoot
        VenvRoot = $skipVenvJunction
        PythonPath = Join-Path $skipVenvJunction "Scripts\python.exe"
        CacheRoot = Join-Path $skipRuntimeRoot "cache"
        PipCacheDir = Join-Path $skipRuntimeRoot "cache\pip"
        PycacheDir = Join-Path $skipRuntimeRoot "cache\pycache"
        PytestCacheDir = Join-Path $skipRuntimeRoot "cache\pytest"
        TempRoot = Join-Path $skipRuntimeRoot "temp"
        MicrosExtractDir = Join-Path $skipRuntimeRoot "temp\micros-extract"
        DependencyFingerprintPath = Join-Path $skipRuntimeRoot "dependency-fingerprint.sha256"
    }
    $script:fixtureRuntimeValid = $true
    $script:fixturePythonUsable = $true
    $script:fixtureInstallCommandInvoked = $false
    $script:fixtureSentinel = $skipSentinel
    $skipBefore = Get-ChildPathSnapshot -Root $skipRuntimeRoot
    $skipMessage = ""
    try {
        Invoke-GiftCardReconRuntimeInitialization `
            -ProgramRoot $programRoot `
            -SkipInstall `
            -ForceInstall | Out-Null
        throw "the -SkipInstall preflight unexpectedly succeeded"
    } catch {
        $skipMessage = $_.Exception.Message
    }
    Assert-True (
        $skipMessage -like "*runtime is missing or out of date*"
    ) "-SkipInstall must keep its established out-of-date-runtime error"
    Assert-True (-not $script:fixtureInstallCommandInvoked) "-SkipInstall must not run an install command"
    Assert-True (
        [IO.File]::ReadAllText($skipSentinel) -eq "must remain"
    ) "-SkipInstall must preserve the external sentinel"
    Assert-True (
        (($skipBefore -join "`n") -eq ((Get-ChildPathSnapshot -Root $skipRuntimeRoot) -join "`n"))
    ) "-SkipInstall must not add runtime directories"

    Write-Host "Runtime reparse guard tests passed."
} finally {
    if ($null -eq $originalLocalAppData) {
        Remove-Item Env:LOCALAPPDATA -ErrorAction SilentlyContinue
    } else {
        $env:LOCALAPPDATA = $originalLocalAppData
    }

    for ($index = $junctions.Count - 1; $index -ge 0; $index--) {
        $junction = $junctions[$index]
        if (Test-Path -LiteralPath $junction) {
            [IO.Directory]::Delete($junction)
        }
    }

    $resolvedFixture = [IO.Path]::GetFullPath($fixture)
    if (-not $resolvedFixture.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Test fixture escaped the temporary root: $resolvedFixture"
    }
    if ([IO.Directory]::Exists($resolvedFixture)) {
        [IO.Directory]::Delete($resolvedFixture, $true)
    }
}
