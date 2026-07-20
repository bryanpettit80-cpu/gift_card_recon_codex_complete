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
        Assert-GiftCardReconVenvRootIsSafeToClear -Runtime $Runtime
        return [pscustomobject]@{ Blocked = $false; Message = "" }
    } catch {
        return [pscustomobject]@{ Blocked = $true; Message = $_.Exception.Message }
    }
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
