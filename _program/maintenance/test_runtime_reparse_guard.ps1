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
    $paths = [Collections.Generic.List[string]]::new()
    $pendingDirectories = [Collections.Generic.Stack[string]]::new()
    $pendingDirectories.Push($fullRoot)
    while ($pendingDirectories.Count -gt 0) {
        $directory = $pendingDirectories.Pop()
        foreach ($child in @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)) {
            $paths.Add($child.FullName.Substring($fullRoot.Length).TrimStart('\', '/'))
            $attributes = $child.Attributes
            if (
                ($attributes -band [IO.FileAttributes]::Directory) -ne 0 -and
                ($attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0
            ) {
                $pendingDirectories.Push([string]$child.FullName)
            }
        }
    }
    return @($paths | Sort-Object)
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw "The runtime reparse guard fixture requires Windows junction support."
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$fixture = Join-Path $tempRoot ("GiftCardRuntimeReparseTest-{0}" -f [Guid]::NewGuid().ToString("N"))
$junctions = [Collections.Generic.List[string]]::new()
$originalLocalAppData = $env:LOCALAPPDATA
$runtimeEnvironmentNames = @("PIP_CACHE_DIR", "PYTHONPYCACHEPREFIX", "TEMP", "TMP")
$originalRuntimeEnvironment = @{}
foreach ($environmentName in $runtimeEnvironmentNames) {
    $originalRuntimeEnvironment[$environmentName] = [Environment]::GetEnvironmentVariable(
        $environmentName,
        [EnvironmentVariableTarget]::Process
    )
}

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

    $leafProbeRuntime = [pscustomobject]@{
        RuntimeRoot = $leafRuntimeRoot
        VenvRoot = $leafJunction
        PythonPath = Join-Path $leafJunction "Scripts\python.exe"
    }
    foreach ($probeFunction in @("Test-GiftCardReconPython", "Test-GiftCardReconRuntime")) {
        $probeMessage = ""
        try {
            & $probeFunction -Runtime $leafProbeRuntime | Out-Null
            throw "$probeFunction unexpectedly suppressed the reparse guard"
        } catch {
            $probeMessage = $_.Exception.Message
        }
        Assert-True (
            $probeMessage -like "*link, junction, or other reparse point*"
        ) "$probeFunction must propagate the exact reparse safety failure"
        Assert-True (
            $probeMessage -like "*$leafJunction*"
        ) "$probeFunction must identify the blocked venv before probing it"
    }

    # An ordinary venv with a redirected Scripts directory is also unsafe. The
    # scanner must identify the junction entry without walking into its target.
    $scriptsRuntimeRoot = Join-Path $fixture "ScriptsDescendantRuntime"
    $scriptsVenvRoot = Join-Path $scriptsRuntimeRoot "venv"
    $externalScripts = Join-Path $fixture "ExternalScripts"
    $scriptsJunction = Join-Path $scriptsVenvRoot "Scripts"
    [void][IO.Directory]::CreateDirectory($scriptsVenvRoot)
    [void][IO.Directory]::CreateDirectory($externalScripts)
    $scriptsSentinel = Join-Path $externalScripts "python.exe"
    [IO.File]::WriteAllText($scriptsSentinel, "must not execute")
    New-Item -ItemType Junction -Path $scriptsJunction -Target $externalScripts | Out-Null
    $junctions.Add($scriptsJunction)
    $scriptsRuntime = [pscustomobject]@{
        RuntimeRoot = $scriptsRuntimeRoot
        VenvRoot = $scriptsVenvRoot
    }
    $scriptsResult = Invoke-VenvGuard -Runtime $scriptsRuntime
    Assert-True $scriptsResult.Blocked "a Scripts descendant junction must be rejected"
    Assert-True (
        $scriptsResult.Message -like "*$scriptsJunction*"
    ) "the blocked Scripts junction must be identified"
    Assert-True (
        [IO.File]::ReadAllText($scriptsSentinel) -eq "must not execute"
    ) "the Scripts junction target must remain untouched"

    # Deep descendant junctions must be caught using lexical traversal as well.
    $packageRuntimeRoot = Join-Path $fixture "PackageDescendantRuntime"
    $packageVenvRoot = Join-Path $packageRuntimeRoot "venv"
    $sitePackages = Join-Path $packageVenvRoot "Lib\site-packages"
    $externalPackage = Join-Path $fixture "ExternalPackage"
    $packageJunction = $sitePackages
    [void][IO.Directory]::CreateDirectory((Join-Path $packageVenvRoot "Lib"))
    [void][IO.Directory]::CreateDirectory($externalPackage)
    $packageSentinel = Join-Path $externalPackage "sentinel.txt"
    [IO.File]::WriteAllText($packageSentinel, "must remain")
    New-Item -ItemType Junction -Path $packageJunction -Target $externalPackage | Out-Null
    $junctions.Add($packageJunction)
    $packageRuntime = [pscustomobject]@{
        RuntimeRoot = $packageRuntimeRoot
        VenvRoot = $packageVenvRoot
    }
    $packageResult = Invoke-VenvGuard -Runtime $packageRuntime
    Assert-True $packageResult.Blocked "a site-packages descendant junction must be rejected"
    Assert-True (
        $packageResult.Message -like "*$packageJunction*"
    ) "the blocked site-packages junction must be identified"
    Assert-True (
        [IO.File]::ReadAllText($packageSentinel) -eq "must remain"
    ) "the site-packages junction target must remain untouched"

    # Exercise the actual initialization entry point. These cases use a real
    # venv junction and make each simulated Python probe a mutation canary, so
    # moving the guard below runtime validation would visibly alter the
    # external sentinel before the create, repair, or refresh decision.
    $script:fixtureRuntime = $null
    $script:fixtureRuntimeValid = $false
    $script:fixturePythonUsable = $false
    $script:fixtureRuntimeProbeInvoked = $false
    $script:fixtureEnvironmentReady = $true
    $script:fixtureProbeMustMutate = $true
    $script:fixtureSwapAfterProbe = ""
    $script:fixtureSwapTarget = ""
    $script:fixtureInstallCommandInvoked = $false
    $script:fixtureSentinel = ""

    function Test-FixtureRuntimeEnvironmentReady {
        param([Parameter(Mandatory = $true)][pscustomobject]$Runtime)

        foreach ($directory in @(
            $Runtime.RuntimeRoot,
            $Runtime.CacheRoot,
            $Runtime.PipCacheDir,
            $Runtime.PycacheDir,
            $Runtime.PytestCacheDir,
            $Runtime.TempRoot,
            $Runtime.MicrosExtractDir
        )) {
            if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
                return $false
            }
        }
        return (
            $env:PIP_CACHE_DIR -eq $Runtime.PipCacheDir -and
            $env:PYTHONPYCACHEPREFIX -eq $Runtime.PycacheDir -and
            $env:TEMP -eq $Runtime.TempRoot -and
            $env:TMP -eq $Runtime.TempRoot
        )
    }

    function Set-FixtureVenvJunction {
        if ([IO.Directory]::Exists($script:fixtureRuntime.VenvRoot)) {
            [IO.Directory]::Delete($script:fixtureRuntime.VenvRoot, $true)
        }
        New-Item `
            -ItemType Junction `
            -Path $script:fixtureRuntime.VenvRoot `
            -Target $script:fixtureSwapTarget | Out-Null
        $junctions.Add([string]$script:fixtureRuntime.VenvRoot)
    }

    function Get-GiftCardReconRuntime { return $script:fixtureRuntime }
    function Get-GiftCardReconDependencyFingerprint {
        param([Parameter(Mandatory = $true)][string]$ProgramRoot)
        return "fixture-fingerprint"
    }
    function Test-GiftCardReconRuntime {
        param([Parameter(Mandatory = $true)][pscustomobject]$Runtime)
        $script:fixtureRuntimeProbeInvoked = $true
        $script:fixtureEnvironmentReady = (
            $script:fixtureEnvironmentReady -and (Test-FixtureRuntimeEnvironmentReady -Runtime $Runtime)
        )
        if ($script:fixtureProbeMustMutate) {
            [IO.File]::WriteAllText($script:fixtureSentinel, "modified by linked Python probe")
        }
        if ($script:fixtureSwapAfterProbe -eq "runtime") {
            Set-FixtureVenvJunction
        }
        return [bool]$script:fixtureRuntimeValid
    }
    function Test-GiftCardReconPython {
        param([Parameter(Mandatory = $true)][pscustomobject]$Runtime)
        $script:fixtureRuntimeProbeInvoked = $true
        $script:fixtureEnvironmentReady = (
            $script:fixtureEnvironmentReady -and (Test-FixtureRuntimeEnvironmentReady -Runtime $Runtime)
        )
        if ($script:fixtureProbeMustMutate) {
            [IO.File]::WriteAllText($script:fixtureSentinel, "modified by linked Python probe")
        }
        if ($script:fixtureSwapAfterProbe -eq "python") {
            Set-FixtureVenvJunction
        }
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
        $script:fixtureRuntimeProbeInvoked = $false
        $script:fixtureEnvironmentReady = $true
        $script:fixtureProbeMustMutate = $true
        $script:fixtureSwapAfterProbe = ""
        $script:fixtureSwapTarget = ""
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
            -not $script:fixtureRuntimeProbeInvoked
        ) "the $($installCase.Name) path must block before executing linked Python or pip"
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

    # A linked runtime that would otherwise be valid and current is rejected
    # before its Python is executed. The matching dependency fingerprint proves
    # this is the no-install path rather than another refresh scenario.
    $noInstallCaseRoot = Join-Path $fixture "NoInstallCase"
    $noInstallLocalAppData = Join-Path $noInstallCaseRoot "LocalAppData"
    $noInstallRuntimeRoot = Join-Path $noInstallLocalAppData "GiftCardRecon"
    $noInstallExternalVenv = Join-Path $noInstallCaseRoot "ExternalVenv"
    $noInstallVenvJunction = Join-Path $noInstallRuntimeRoot "venv"
    [void][IO.Directory]::CreateDirectory($noInstallRuntimeRoot)
    [void][IO.Directory]::CreateDirectory((Join-Path $noInstallExternalVenv "Scripts"))
    [IO.File]::WriteAllText((Join-Path $noInstallExternalVenv "pyvenv.cfg"), "fixture")
    [IO.File]::WriteAllText((Join-Path $noInstallExternalVenv "Scripts\python.exe"), "fixture")
    [IO.File]::WriteAllText(
        (Join-Path $noInstallRuntimeRoot "dependency-fingerprint.sha256"),
        "fixture-fingerprint"
    )
    $noInstallSentinel = Join-Path $noInstallExternalVenv "sentinel.txt"
    [IO.File]::WriteAllText($noInstallSentinel, "must remain")
    New-Item `
        -ItemType Junction `
        -Path $noInstallVenvJunction `
        -Target $noInstallExternalVenv | Out-Null
    $junctions.Add($noInstallVenvJunction)
    $script:fixtureRuntime = [pscustomobject]@{
        RuntimeRoot = $noInstallRuntimeRoot
        VenvRoot = $noInstallVenvJunction
        PythonPath = Join-Path $noInstallVenvJunction "Scripts\python.exe"
        CacheRoot = Join-Path $noInstallRuntimeRoot "cache"
        PipCacheDir = Join-Path $noInstallRuntimeRoot "cache\pip"
        PycacheDir = Join-Path $noInstallRuntimeRoot "cache\pycache"
        PytestCacheDir = Join-Path $noInstallRuntimeRoot "cache\pytest"
        TempRoot = Join-Path $noInstallRuntimeRoot "temp"
        MicrosExtractDir = Join-Path $noInstallRuntimeRoot "temp\micros-extract"
        DependencyFingerprintPath = Join-Path $noInstallRuntimeRoot "dependency-fingerprint.sha256"
    }
    $script:fixtureRuntimeValid = $true
    $script:fixturePythonUsable = $true
    $script:fixtureRuntimeProbeInvoked = $false
    $script:fixtureEnvironmentReady = $true
    $script:fixtureProbeMustMutate = $true
    $script:fixtureSwapAfterProbe = ""
    $script:fixtureSwapTarget = ""
    $script:fixtureInstallCommandInvoked = $false
    $script:fixtureSentinel = $noInstallSentinel
    $noInstallBefore = Get-ChildPathSnapshot -Root $noInstallRuntimeRoot
    $noInstallMessage = ""
    try {
        Invoke-GiftCardReconRuntimeInitialization -ProgramRoot $programRoot | Out-Null
        throw "the valid no-install runtime unexpectedly accepted a reparse-point venv"
    } catch {
        $noInstallMessage = $_.Exception.Message
    }
    Assert-True (
        $noInstallMessage -like "*link, junction, or other reparse point*"
    ) "a valid no-install linked runtime must fail at the reparse guard"
    Assert-True (
        $noInstallMessage -like "*$noInstallVenvJunction*"
    ) "the no-install error must identify the venv junction"
    Assert-True (
        -not $script:fixtureRuntimeProbeInvoked
    ) "the no-install path must block before executing linked Python or pip"
    Assert-True (
        -not $script:fixtureInstallCommandInvoked
    ) "the no-install path must not run an install command"
    Assert-True (
        [IO.File]::ReadAllText($noInstallSentinel) -eq "must remain"
    ) "the no-install path must preserve the external sentinel"
    Assert-True (
        (
            ($noInstallBefore -join "`n") -eq
            ((Get-ChildPathSnapshot -Root $noInstallRuntimeRoot) -join "`n")
        )
    ) "the no-install path must not add runtime directories"

    # On an ordinary runtime, isolated cache/temp directories and environment
    # variables must be ready before the first Python validation probe.
    $environmentCaseRoot = Join-Path $fixture "EnvironmentOrderCase"
    $environmentRuntimeRoot = Join-Path $environmentCaseRoot "GiftCardRecon"
    $environmentVenvRoot = Join-Path $environmentRuntimeRoot "venv"
    [void][IO.Directory]::CreateDirectory((Join-Path $environmentVenvRoot "Scripts"))
    [IO.File]::WriteAllText((Join-Path $environmentVenvRoot "pyvenv.cfg"), "fixture")
    [IO.File]::WriteAllText((Join-Path $environmentVenvRoot "Scripts\python.exe"), "fixture")
    [IO.File]::WriteAllText(
        (Join-Path $environmentRuntimeRoot "dependency-fingerprint.sha256"),
        "fixture-fingerprint"
    )
    $script:fixtureRuntime = [pscustomobject]@{
        RuntimeRoot = $environmentRuntimeRoot
        VenvRoot = $environmentVenvRoot
        PythonPath = Join-Path $environmentVenvRoot "Scripts\python.exe"
        CacheRoot = Join-Path $environmentRuntimeRoot "cache"
        PipCacheDir = Join-Path $environmentRuntimeRoot "cache\pip"
        PycacheDir = Join-Path $environmentRuntimeRoot "cache\pycache"
        PytestCacheDir = Join-Path $environmentRuntimeRoot "cache\pytest"
        TempRoot = Join-Path $environmentRuntimeRoot "temp"
        MicrosExtractDir = Join-Path $environmentRuntimeRoot "temp\micros-extract"
        DependencyFingerprintPath = Join-Path $environmentRuntimeRoot "dependency-fingerprint.sha256"
    }
    $script:fixtureRuntimeValid = $true
    $script:fixturePythonUsable = $true
    $script:fixtureRuntimeProbeInvoked = $false
    $script:fixtureEnvironmentReady = $true
    $script:fixtureProbeMustMutate = $false
    $script:fixtureSwapAfterProbe = ""
    $script:fixtureSwapTarget = ""
    $script:fixtureInstallCommandInvoked = $false
    $script:fixtureSentinel = ""
    Invoke-GiftCardReconRuntimeInitialization -ProgramRoot $programRoot | Out-Null
    Assert-True $script:fixtureRuntimeProbeInvoked "the normal runtime must be validated"
    Assert-True (
        $script:fixtureEnvironmentReady
    ) "runtime cache/temp isolation must be configured before validation"
    Assert-True (
        -not $script:fixtureInstallCommandInvoked
    ) "a valid current ordinary runtime must not run an install command"

    # Simulate an attacker replacing the missing/ordinary venv with a real
    # junction after the early guard and final validation probe. The command-
    # boundary recheck must catch every create, repair, and refresh path before
    # venv --clear or pip can touch the external target.
    foreach ($swapCase in @(
        [pscustomobject]@{
            Name = "create"; RuntimeValid = $false; PythonUsable = $false
            Force = $false; SwapAfterProbe = "python"; CreateInitialVenv = $false
        },
        [pscustomobject]@{
            Name = "repair"; RuntimeValid = $false; PythonUsable = $true
            Force = $false; SwapAfterProbe = "python"; CreateInitialVenv = $true
        },
        [pscustomobject]@{
            Name = "refresh"; RuntimeValid = $true; PythonUsable = $true
            Force = $true; SwapAfterProbe = "runtime"; CreateInitialVenv = $true
        }
    )) {
        $swapCaseRoot = Join-Path $fixture ("SwapCase-{0}" -f $swapCase.Name)
        $swapRuntimeRoot = Join-Path $swapCaseRoot "GiftCardRecon"
        $swapVenvRoot = Join-Path $swapRuntimeRoot "venv"
        $swapExternalVenv = Join-Path $swapCaseRoot "ExternalVenv"
        [void][IO.Directory]::CreateDirectory($swapRuntimeRoot)
        [void][IO.Directory]::CreateDirectory($swapExternalVenv)
        if ($swapCase.CreateInitialVenv) {
            [void][IO.Directory]::CreateDirectory((Join-Path $swapVenvRoot "Scripts"))
            [IO.File]::WriteAllText((Join-Path $swapVenvRoot "pyvenv.cfg"), "fixture")
            [IO.File]::WriteAllText((Join-Path $swapVenvRoot "Scripts\python.exe"), "fixture")
        }
        $swapSentinel = Join-Path $swapExternalVenv "sentinel.txt"
        [IO.File]::WriteAllText($swapSentinel, "must remain")
        $script:fixtureRuntime = [pscustomobject]@{
            RuntimeRoot = $swapRuntimeRoot
            VenvRoot = $swapVenvRoot
            PythonPath = Join-Path $swapVenvRoot "Scripts\python.exe"
            CacheRoot = Join-Path $swapRuntimeRoot "cache"
            PipCacheDir = Join-Path $swapRuntimeRoot "cache\pip"
            PycacheDir = Join-Path $swapRuntimeRoot "cache\pycache"
            PytestCacheDir = Join-Path $swapRuntimeRoot "cache\pytest"
            TempRoot = Join-Path $swapRuntimeRoot "temp"
            MicrosExtractDir = Join-Path $swapRuntimeRoot "temp\micros-extract"
            DependencyFingerprintPath = Join-Path $swapRuntimeRoot "dependency-fingerprint.sha256"
        }
        $script:fixtureRuntimeValid = $swapCase.RuntimeValid
        $script:fixturePythonUsable = $swapCase.PythonUsable
        $script:fixtureRuntimeProbeInvoked = $false
        $script:fixtureEnvironmentReady = $true
        $script:fixtureProbeMustMutate = $false
        $script:fixtureSwapAfterProbe = $swapCase.SwapAfterProbe
        $script:fixtureSwapTarget = $swapExternalVenv
        $script:fixtureInstallCommandInvoked = $false
        $script:fixtureSentinel = $swapSentinel
        $swapMessage = ""
        try {
            Invoke-GiftCardReconRuntimeInitialization `
                -ProgramRoot $programRoot `
                -ForceInstall:$swapCase.Force | Out-Null
            throw "the $($swapCase.Name) swap unexpectedly reached an install command"
        } catch {
            $swapMessage = $_.Exception.Message
        }
        Assert-True (
            $swapMessage -like "*link, junction, or other reparse point*"
        ) "the $($swapCase.Name) swap must fail at a command-boundary guard"
        Assert-True (
            $swapMessage -like "*$swapVenvRoot*"
        ) "the $($swapCase.Name) swap error must identify the replaced venv"
        Assert-True $script:fixtureRuntimeProbeInvoked "the swap must occur after validation begins"
        Assert-True (
            $script:fixtureEnvironmentReady
        ) "the $($swapCase.Name) probes must use isolated cache/temp paths"
        Assert-True (
            -not $script:fixtureInstallCommandInvoked
        ) "the $($swapCase.Name) swap must block before venv or pip execution"
        Assert-True (
            [IO.File]::ReadAllText($swapSentinel) -eq "must remain"
        ) "the $($swapCase.Name) swap must preserve the external sentinel"
    }

    Write-Host "Runtime reparse guard tests passed."
} finally {
    if ($null -eq $originalLocalAppData) {
        Remove-Item Env:LOCALAPPDATA -ErrorAction SilentlyContinue
    } else {
        $env:LOCALAPPDATA = $originalLocalAppData
    }
    foreach ($environmentName in $runtimeEnvironmentNames) {
        [Environment]::SetEnvironmentVariable(
            $environmentName,
            $originalRuntimeEnvironment[$environmentName],
            [EnvironmentVariableTarget]::Process
        )
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
