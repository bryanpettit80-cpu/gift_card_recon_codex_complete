#Requires -Version 5.1
[CmdletBinding()]
param([string]$InterruptRoot = "")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "..\runtime.ps1")

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

# Replace only external Python/pip work: exercise runtime selection, candidate
# publication, pointer durability, and failure behavior with isolated files.
function Test-GiftCardReconRuntime([pscustomobject]$Runtime) {
    return Test-Path -LiteralPath (Join-Path $Runtime.VenvRoot "validated") -PathType Leaf
}
function Test-GiftCardReconPython([pscustomobject]$Runtime) { return $true }
function Reset-GiftCardReconVenv([pscustomobject]$Runtime) {
    [void][IO.Directory]::CreateDirectory($Runtime.VenvRoot)
    [IO.File]::WriteAllText((Join-Path $Runtime.VenvRoot "candidate-created"), "candidate")
    if ($InterruptRoot) { Stop-Process -Id $PID -Force }
}
function Invoke-GiftCardReconChecked([string]$FilePath, [string[]]$Arguments) {
    $script:InstallCalls++
    if ($script:FailInstall) { throw "simulated unavailable package source" }
    $candidateRoot = Split-Path -Parent (Split-Path -Parent $FilePath)
    [IO.File]::WriteAllText((Join-Path $candidateRoot "validated"), "ready")
}

$script:InstallCalls = 0
$script:FailInstall = $false
if ($InterruptRoot) {
    Initialize-GiftCardReconRuntime -ProgramRoot (Join-Path $InterruptRoot "program") -RuntimeRoot (Join-Path $InterruptRoot "runtime") -ForceInstall | Out-Null
    throw "interruption canary was not reached"
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$fixture = Join-Path $tempRoot ("GiftCardRuntimePortability-{0}" -f [Guid]::NewGuid().ToString("N"))
$originalEnvironment = @{}
foreach ($name in @("LOCALAPPDATA", "PIP_CACHE_DIR", "PYTHONPYCACHEPREFIX", "PYTHONPATH", "PYTHONNOUSERSITE", "TEMP", "TMP")) {
    $originalEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, "Process")
}
try {
    $env:LOCALAPPDATA = Join-Path $fixture "Other User's LocalAppData"
    $program = Join-Path $fixture "program"
    $moved = Join-Path $fixture "Other User's moved checkout"
    [void][IO.Directory]::CreateDirectory($program)
    [void][IO.Directory]::CreateDirectory($moved)
    foreach ($name in @("requirements.lock", "requirements-bootstrap.in")) {
        [IO.File]::WriteAllText((Join-Path $program $name), "identical dependency specification")
        [IO.File]::WriteAllText((Join-Path $moved $name), "identical dependency specification")
    }
    $fingerprint = Get-GiftCardReconDependencyFingerprint -ProgramRoot $program
    Assert-True ($fingerprint -eq (Get-GiftCardReconDependencyFingerprint -ProgramRoot $moved)) "moving identical dependencies must preserve their identity"

    $runtimeRoot = Join-Path $fixture "runtime"
    $prior = Get-GiftCardReconRuntime -RuntimeRoot $runtimeRoot
    [void][IO.Directory]::CreateDirectory($prior.VenvRoot)
    [IO.File]::WriteAllText((Join-Path $prior.VenvRoot "validated"), "previous good runtime")
    [IO.File]::WriteAllText($prior.DependencyFingerprintPath, $fingerprint)
    $reused = Initialize-GiftCardReconRuntime -ProgramRoot $moved -RuntimeRoot $runtimeRoot -SkipInstall
    Assert-True ($reused.VenvRoot -eq $prior.VenvRoot -and $script:InstallCalls -eq 0) "moved checkout must reuse runtime without installation or network"
    Assert-True ($env:PYTHONPATH -eq (Join-Path $moved "src")) "source imports must follow moved checkout"

    $script:FailInstall = $true
    $message = ""
    try { Initialize-GiftCardReconRuntime -ProgramRoot $program -RuntimeRoot $runtimeRoot -ForceInstall | Out-Null }
    catch { $message = $_.Exception.Message }
    Assert-True ($message -like "*simulated unavailable*") "failed upgrade must surface its cause"
    Assert-True ((Get-GiftCardReconRuntime -RuntimeRoot $runtimeRoot).VenvRoot -eq $prior.VenvRoot) "failed upgrade must keep prior runtime selected"
    Assert-True ([IO.File]::ReadAllText((Join-Path $prior.VenvRoot "validated")) -eq "previous good runtime") "failed upgrade must preserve prior runtime bytes"

    $script:FailInstall = $false
    $candidate = Initialize-GiftCardReconRuntime -ProgramRoot $program -RuntimeRoot $runtimeRoot -ForceInstall
    Assert-True ($candidate.VenvRoot -ne $prior.VenvRoot) "successful upgrade must use a separate generation"
    Assert-True ((Get-GiftCardReconRuntime -RuntimeRoot $runtimeRoot).VenvRoot -eq $candidate.VenvRoot) "successful upgrade must select validated candidate"
    Assert-True (Test-Path -LiteralPath (Join-Path $prior.VenvRoot "validated")) "successful upgrade must retain rollback runtime"
    $pointer = Join-Path $runtimeRoot "active-runtime.json"
    $before = (Get-FileHash -LiteralPath $pointer -Algorithm SHA256).Hash
    $powershell = (Get-Process -Id $PID).Path
    & $powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -InterruptRoot $fixture *> $null
    Assert-True ($LASTEXITCODE -ne 0) "child must terminate during candidate construction"
    Assert-True ((Get-FileHash -LiteralPath $pointer -Algorithm SHA256).Hash -eq $before) "forced termination must leave previous active pointer intact"
    Assert-True ((Get-GiftCardReconRuntime -RuntimeRoot $runtimeRoot).VenvRoot -eq $candidate.VenvRoot) "restart after termination must select previous good generation"

    $operations = Join-Path $fixture "Other User's Reports"
    $config = Join-Path $fixture "paths.json"
    Set-GiftCardReconPathConfiguration -ProgramRoot $program -OperationsRoot $operations -DropboxRoot (Join-Path $fixture "Dropbox") -ConfigPath $config
    $configured = Get-GiftCardReconPathConfiguration -ConfigPath $config
    Assert-True ($configured.OperationsRoot -eq $operations) "local configuration must retain paths with spaces and apostrophes"
    Assert-True ((Get-GiftCardReconTestRuntimeRoot -ProgramRoot $program) -ne (Get-GiftCardReconRuntime).RuntimeRoot) "test runtime must be separate from reporting runtime"

    $repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    foreach ($runnerName in @("run_weekly.ps1", "run_monthly_close.ps1")) {
        $uncreated = Join-Path $fixture "configuration-only-$runnerName"
        $configOutput = & $powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo "_program\$runnerName") -OperationsRoot $uncreated -ShowConfiguration
        Assert-True ($LASTEXITCODE -eq 0) "$runnerName must show configuration without runtime installation"
        $shown = ($configOutput -join "`n") | ConvertFrom-Json
        Assert-True ($shown.OperationsRoot -eq $uncreated) "$runnerName must show the selected operations root"
        Assert-True (-not (Test-Path -LiteralPath $uncreated)) "$runnerName configuration must not create operator folders"
    }
    foreach ($layout in @("flat", "nested", "configured")) {
        $operatorRoot = Join-Path $fixture $layout
        $sourceRepo = switch ($layout) {
            "flat" { $operatorRoot }
            "nested" { Join-Path $operatorRoot "Gift Card Reconciliation Automation" }
            "configured" { Join-Path $fixture "Separate Program" }
        }
        [void][IO.Directory]::CreateDirectory((Join-Path $sourceRepo "_program"))
        [void][IO.Directory]::CreateDirectory($operatorRoot)
        if ($layout -eq "configured") {
            Set-GiftCardReconPathConfiguration -ProgramRoot (Join-Path $sourceRepo "_program") -OperationsRoot $operatorRoot
        }
        foreach ($kind in @("weekly", "monthly_close")) {
            $templateName = if ($kind -eq "weekly") { "Run Weekly Gift Card Reconciliation.cmd" } else { "Run Monthly Gift Card Close.cmd" }
            $launcher = Join-Path $operatorRoot $templateName
            Copy-Item -LiteralPath (Join-Path $repo "templates\$templateName") -Destination $launcher
            $stub = 'param([string]$OperationsRoot); [IO.File]::WriteAllText((Join-Path $OperationsRoot "runner-observed.txt"), [IO.Path]::GetFullPath($OperationsRoot)); exit 0'
            [IO.File]::WriteAllText((Join-Path $sourceRepo "_program\run_$kind.ps1"), $stub)
            & $env:ComSpec /d /c "`"$launcher`" <nul" *> $null
            Assert-True ($LASTEXITCODE -eq 0) "$layout $kind launcher must resolve a runnable program"
            $observed = [IO.File]::ReadAllText((Join-Path $operatorRoot "runner-observed.txt"))
            Assert-True ($observed.TrimEnd('\') -eq $operatorRoot) "$layout $kind launcher must pass its operations root"
        }
    }
    Write-Host "Runtime portability, offline reuse, failed/terminated upgrade, path configuration, and launcher fixtures passed."
} finally {
    foreach ($name in $originalEnvironment.Keys) { [Environment]::SetEnvironmentVariable($name, $originalEnvironment[$name], "Process") }
    $resolved = [IO.Path]::GetFullPath($fixture)
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $resolved) -like "GiftCardRuntimePortability-*") {
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    }
}
