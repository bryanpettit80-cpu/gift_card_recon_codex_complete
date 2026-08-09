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

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw "The runtime base-interpreter fixture requires Windows."
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$fixture = Join-Path $tempRoot ("GiftCardRuntimeBasePythonTest-{0}" -f [Guid]::NewGuid().ToString("N"))
$originalLocalAppData = $env:LOCALAPPDATA
$originalPath = $env:PATH

try {
    [void][IO.Directory]::CreateDirectory($fixture)
    $env:LOCALAPPDATA = Join-Path $fixture "LocalAppData"
    [void][IO.Directory]::CreateDirectory($env:LOCALAPPDATA)
    $runtime = Get-GiftCardReconRuntime

    $bootstrapRuntime = [pscustomobject]@{
        VenvRoot = Join-Path $fixture "bootstrap-not-present"
    }
    $basePython = Resolve-GiftCardReconBasePython -Runtime $bootstrapRuntime
    Invoke-GiftCardReconChecked -FilePath $basePython -Arguments @(
        "-m", "venv", "--without-pip", $runtime.VenvRoot
    )

    # Simulate install.ps1 being launched from a shell where this exact target
    # venv is active. Get-Command python must now resolve to the executable that
    # is about to be cleared, reproducing the former self-clear failure.
    $env:PATH = "$(Join-Path $runtime.VenvRoot 'Scripts');$originalPath"
    $pathPythonCommands = @(Get-Command -Name python -CommandType Application -ErrorAction Stop)
    $pathPython = [string]$pathPythonCommands[0].Source
    Assert-True (
        ([IO.Path]::GetFullPath($pathPython)).Equals(
            [IO.Path]::GetFullPath($runtime.PythonPath),
            [StringComparison]::OrdinalIgnoreCase
        )
    ) "the activated target venv must provide python from PATH"

    $selectedBasePython = Resolve-GiftCardReconBasePython -Runtime $runtime
    Assert-True (
        -not (Test-GiftCardReconPathWithinRoot -Path $selectedBasePython -Root $runtime.VenvRoot)
    ) "the venv builder must be outside the target being cleared"

    Reset-GiftCardReconVenv -Runtime $runtime -WithoutPip
    Assert-True (
        Test-Path -LiteralPath (Join-Path $runtime.VenvRoot "pyvenv.cfg") -PathType Leaf
    ) "the target venv must remain complete after an activated-runtime rebuild"
    Assert-True (
        Test-Path -LiteralPath $runtime.PythonPath -PathType Leaf
    ) "the target Python executable must remain after rebuild"

    & $runtime.PythonPath -c "import sys; raise SystemExit(0 if sys.prefix != sys.base_prefix else 1)" *> $null
    Assert-True ($LASTEXITCODE -eq 0) "the rebuilt target must still be a working virtual environment"
    Write-Host "Runtime base-interpreter tests passed."
} finally {
    $env:PATH = $originalPath
    if ($null -eq $originalLocalAppData) {
        Remove-Item Env:LOCALAPPDATA -ErrorAction SilentlyContinue
    } else {
        $env:LOCALAPPDATA = $originalLocalAppData
    }

    $resolvedFixture = [IO.Path]::GetFullPath($fixture)
    if (-not $resolvedFixture.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Test fixture escaped the temporary root: $resolvedFixture"
    }
    if ([IO.Directory]::Exists($resolvedFixture)) {
        [IO.Directory]::Delete($resolvedFixture, $true)
    }
}
