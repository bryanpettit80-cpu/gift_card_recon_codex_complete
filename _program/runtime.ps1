$ErrorActionPreference = "Stop"

function Write-GiftCardReconAtomicJson {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)]$Value)
    $temporary = "$Path.$([Guid]::NewGuid().ToString('N')).tmp"
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($Value | ConvertTo-Json -Depth 8))
    $stream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally { $stream.Dispose() }
    try {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [IO.File]::Replace($temporary, $Path, "$Path.previous")
        } else { [IO.File]::Move($temporary, $Path) }
    } finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Get-GiftCardReconPathConfiguration {
    param([string]$ConfigPath = "")
    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        $ConfigPath = Join-Path (Get-GiftCardReconRuntime).RuntimeRoot "paths.json"
    }
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { return $null }
    $configuration = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    if ($configuration.schema -ne 1) { throw "Unsupported Gift Card Recon path configuration: $ConfigPath" }
    return $configuration
}

function Set-GiftCardReconPathConfiguration {
    param(
        [Parameter(Mandatory = $true)][string]$ProgramRoot,
        [Parameter(Mandatory = $true)][string]$OperationsRoot,
        [string]$DropboxRoot = "",
        [string]$ConfigPath = ""
    )
    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        $ConfigPath = Join-Path (Get-GiftCardReconRuntime).RuntimeRoot "paths.json"
    }
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $ConfigPath))
    Write-GiftCardReconAtomicJson -Path $ConfigPath -Value ([ordered]@{
        schema = 1
        RepositoryRoot = Split-Path -Parent ([IO.Path]::GetFullPath($ProgramRoot))
        OperationsRoot = [IO.Path]::GetFullPath($OperationsRoot)
        DropboxRoot = Resolve-GiftCardReconDropboxRoot -OperationsRoot $OperationsRoot -DropboxRoot $DropboxRoot
    })
}

function Get-GiftCardReconTestRuntimeRoot {
    param([Parameter(Mandatory = $true)][string]$ProgramRoot)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes(([IO.Path]::GetFullPath($ProgramRoot)).ToLowerInvariant())
        $key = (-join ($hasher.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") })).Substring(0, 16)
    } finally { $hasher.Dispose() }
    return Join-Path (Get-GiftCardReconRuntime).RuntimeRoot "tests\$key"
}

function Get-GiftCardReconRuntime {
    [CmdletBinding()]
    param([string]$RuntimeRoot = "")

    $localAppData = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    }
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw "Windows Local AppData could not be resolved for the Gift Card Recon runtime."
    }

    if ([string]::IsNullOrWhiteSpace($RuntimeRoot)) {
        $RuntimeRoot = Join-Path $localAppData "GiftCardRecon"
    }
    $runtimeRoot = [IO.Path]::GetFullPath($RuntimeRoot)
    $cacheRoot = Join-Path $runtimeRoot "cache"
    $tempRoot = Join-Path $runtimeRoot "temp"
    $venvRoot = Join-Path $runtimeRoot "venv"
    $activePath = Join-Path $runtimeRoot "active-runtime.json"
    $fingerprintPath = Join-Path $runtimeRoot "dependency-fingerprint.sha256"
    if (Test-Path -LiteralPath $activePath -PathType Leaf) {
        $active = Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json
        $generationRoot = Join-Path $runtimeRoot "environments"
        $venvRoot = [IO.Path]::GetFullPath([string]$active.venv_root)
        if ($active.schema -ne 1 -or -not (Test-GiftCardReconPathWithinRoot -Path $venvRoot -Root $generationRoot) -or $venvRoot -eq $generationRoot) {
            throw "The local runtime pointer is invalid: $activePath. Restore its .previous copy or rerun setup."
        }
        $fingerprintPath = Join-Path $venvRoot "dependency-fingerprint.sha256"
    }

    return [pscustomobject]@{
        RuntimeRoot = $runtimeRoot
        VenvRoot = $venvRoot
        PythonPath = Join-Path $venvRoot "Scripts\python.exe"
        CacheRoot = $cacheRoot
        PipCacheDir = Join-Path $cacheRoot "pip"
        PycacheDir = Join-Path $cacheRoot "pycache"
        PytestCacheDir = Join-Path $cacheRoot "pytest"
        TempRoot = $tempRoot
        MicrosExtractDir = Join-Path $tempRoot "micros-extract"
        DependencyFingerprintPath = $fingerprintPath
    }
}

function Resolve-GiftCardReconDropboxRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OperationsRoot,

        [string]$DropboxRoot = ""
    )

    $resolvedOperationsRoot = [IO.Path]::GetFullPath($OperationsRoot).TrimEnd('\', '/')
    if (-not [string]::IsNullOrWhiteSpace($DropboxRoot)) {
        if ([IO.Path]::IsPathRooted($DropboxRoot)) {
            return [IO.Path]::GetFullPath($DropboxRoot).TrimEnd('\', '/')
        }
        return [IO.Path]::GetFullPath(
            (Join-Path $resolvedOperationsRoot $DropboxRoot)
        ).TrimEnd('\', '/')
    }

    $operationsParent = [IO.Directory]::GetParent($resolvedOperationsRoot)
    if ($null -eq $operationsParent) {
        throw "Dropbox root could not be inferred from operations root: $resolvedOperationsRoot"
    }

    # Legacy: Dropbox\Gift Card Reconciliation
    # Current: Dropbox\Automations\Gift Card Reconciliation
    if ($operationsParent.Name.Equals("Automations", [StringComparison]::OrdinalIgnoreCase)) {
        if ($null -eq $operationsParent.Parent) {
            throw "Dropbox root could not be inferred above Automations: $resolvedOperationsRoot"
        }
        return $operationsParent.Parent.FullName.TrimEnd('\', '/')
    }
    return $operationsParent.FullName.TrimEnd('\', '/')
}

function Set-GiftCardReconRuntimeEnvironment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    foreach ($path in @(
        $Runtime.RuntimeRoot,
        $Runtime.CacheRoot,
        $Runtime.PipCacheDir,
        $Runtime.PycacheDir,
        $Runtime.PytestCacheDir,
        $Runtime.TempRoot,
        $Runtime.MicrosExtractDir
    )) {
        New-Item -ItemType Directory -Force -Path $path | Out-Null
    }

    $env:PIP_CACHE_DIR = $Runtime.PipCacheDir
    $env:PYTHONPYCACHEPREFIX = $Runtime.PycacheDir
    $env:TEMP = $Runtime.TempRoot
    $env:TMP = $Runtime.TempRoot
}

function Get-GiftCardReconDependencyFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot
    )

    $resolvedProgramRoot = [IO.Path]::GetFullPath($ProgramRoot).TrimEnd('\', '/')
    $specifications = @(
        (Join-Path $resolvedProgramRoot "requirements.lock"),
        (Join-Path $resolvedProgramRoot "requirements-bootstrap.in")
    )
    $parts = @(
        "runtime-schema=2"
    )
    foreach ($path in $specifications) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Runtime dependency specification is missing: $path"
        }
        $digest = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        $parts += "$(Split-Path -Leaf $path)=$digest"
    }

    $payload = [Text.Encoding]::UTF8.GetBytes(($parts -join "`n"))
    $hasher = [Security.Cryptography.SHA256]::Create()
    try {
        return -join ($hasher.ComputeHash($payload) | ForEach-Object { $_.ToString("x2") })
    } finally {
        $hasher.Dispose()
    }
}


function Test-GiftCardReconReparsePoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    } catch [System.Management.Automation.ItemNotFoundException] {
        return $false
    }

    return (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Assert-GiftCardReconVenvRootIsSafeToModify {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    if ([string]::IsNullOrWhiteSpace([string]$Runtime.VenvRoot)) {
        throw "Refusing to modify the Gift Card Recon runtime because its venv path is empty."
    }

    $venvRoot = [IO.Path]::GetFullPath([string]$Runtime.VenvRoot)
    $probe = $venvRoot
    while ($true) {
        # Walk the lexical path rather than resolving it so the reparse entry
        # remains visible even when its target contains an ordinary child.
        if (Test-GiftCardReconReparsePoint -Path $probe) {
            throw (
                "Refusing to modify the Gift Card Recon runtime because '$probe' in the path " +
                "to '$venvRoot' is a link, junction, or other reparse point. Remove that entry " +
                "manually, then rerun setup."
            )
        }

        $parent = [IO.Path]::GetDirectoryName($probe)
        if (
            [string]::IsNullOrWhiteSpace($parent) -or
            $parent.Equals($probe, [StringComparison]::OrdinalIgnoreCase)
        ) {
            break
        }
        $probe = $parent
    }

    if (-not [IO.Directory]::Exists($venvRoot)) {
        return
    }

    # Enumerate one lexical directory at a time. A recursive provider walk can
    # follow a junction before its attributes are inspected, so descendants are
    # checked first and only ordinary directories are added to the stack.
    $pendingDirectories = [Collections.Generic.Stack[string]]::new()
    $pendingDirectories.Push($venvRoot)
    while ($pendingDirectories.Count -gt 0) {
        $directory = $pendingDirectories.Pop()
        # Recheck after dequeueing: an ordinary directory may have been
        # replaced by a junction after it was first inspected.
        if (Test-GiftCardReconReparsePoint -Path $directory) {
            throw (
                "Refusing to modify the Gift Card Recon runtime because '$directory' under " +
                "'$venvRoot' is a link, junction, or other reparse point. Remove that entry " +
                "manually, then rerun setup."
            )
        }
        try {
            $children = @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)
        } catch {
            throw (
                "Refusing to modify the Gift Card Recon runtime because '$directory' under " +
                "'$venvRoot' could not be safely inspected: $($_.Exception.Message)"
            )
        }

        foreach ($child in $children) {
            try {
                $attributes = $child.Attributes
                $childPath = [IO.Path]::GetFullPath([string]$child.FullName)
            } catch {
                throw (
                    "Refusing to modify the Gift Card Recon runtime because an entry under " +
                    "'$directory' could not be safely inspected: $($_.Exception.Message)"
                )
            }

            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw (
                    "Refusing to modify the Gift Card Recon runtime because '$childPath' under " +
                    "'$venvRoot' is a link, junction, or other reparse point. Remove that entry " +
                    "manually, then rerun setup."
                )
            }
            if (($attributes -band [IO.FileAttributes]::Directory) -ne 0) {
                $pendingDirectories.Push($childPath)
            }
        }
    }
}

function Test-GiftCardReconPython {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    if (
        -not (Test-Path -LiteralPath $Runtime.PythonPath -PathType Leaf) -or
        -not (Test-Path -LiteralPath (Join-Path $Runtime.VenvRoot "pyvenv.cfg") -PathType Leaf)
    ) {
        return $false
    }

    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    try {
        & $Runtime.PythonPath --version *> $null
    } catch {
        return $false
    }
    if ($LASTEXITCODE -ne 0) {
        return $false
    }

    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    try {
        & $Runtime.PythonPath -m pip --version *> $null
    } catch {
        return $false
    }
    return $LASTEXITCODE -eq 0
}

function Test-GiftCardReconRuntime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    if (-not (Test-GiftCardReconPython -Runtime $Runtime)) {
        return $false
    }

    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    try {
        & $Runtime.PythonPath -m pip check *> $null
    } catch {
        return $false
    }
    if ($LASTEXITCODE -ne 0) {
        return $false
    }

    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    try {
        & $Runtime.PythonPath -c "import openpyxl, xlrd, pypdf, gift_card_recon" *> $null
    } catch {
        return $false
    }
    return $LASTEXITCODE -eq 0
}

function Invoke-GiftCardReconChecked {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments | Out-Host
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "$FilePath failed with exit code $exitCode."
    }
}

function Test-GiftCardReconPathWithinRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $resolvedPath = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $resolvedRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    if ($resolvedPath.Equals($resolvedRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    $rootPrefix = $resolvedRoot + [IO.Path]::DirectorySeparatorChar
    return $resolvedPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)
}

function Resolve-GiftCardReconBasePython {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    # An activated Gift Card runtime puts its own python.exe first on PATH.
    # That process cannot safely clear the venv containing itself, so query
    # each available interpreter (then the Windows launcher) for the real base
    # executable and validate that executable before using it as the builder.
    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    $candidates = [Collections.Generic.List[object]]::new()
    foreach ($command in @(Get-Command -Name python -CommandType Application -All -ErrorAction SilentlyContinue)) {
        $candidates.Add([pscustomobject]@{
            FilePath = [string]$command.Source
            PrefixArguments = @()
        })
    }
    foreach ($command in @(Get-Command -Name py -CommandType Application -All -ErrorAction SilentlyContinue)) {
        $candidates.Add([pscustomobject]@{
            FilePath = [string]$command.Source
            PrefixArguments = @("-3")
        })
    }

    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $probeScript = (
        "import json, sys; " +
        "print(json.dumps({'base_executable': sys._base_executable}))"
    )
    $validationScript = (
        "import sys; " +
        "raise SystemExit(0 if sys.version_info >= (3, 10) " +
        "and sys.prefix == sys.base_prefix else 1)"
    )

    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate.FilePath)) {
            continue
        }
        $candidateKey = "$($candidate.FilePath)|$($candidate.PrefixArguments -join ' ')"
        if (-not $seen.Add($candidateKey)) {
            continue
        }

        $probeArguments = @($candidate.PrefixArguments) + @("-c", $probeScript)
        try {
            $probeOutput = @(& $candidate.FilePath @probeArguments 2>$null)
            $probeExitCode = $LASTEXITCODE
        } catch {
            continue
        }
        if ($probeExitCode -ne 0 -or $probeOutput.Count -eq 0) {
            continue
        }

        try {
            $probe = $probeOutput[-1] | ConvertFrom-Json -ErrorAction Stop
            $basePython = [IO.Path]::GetFullPath([string]$probe.base_executable)
        } catch {
            continue
        }
        if (
            [string]::IsNullOrWhiteSpace($basePython) -or
            (Test-GiftCardReconPathWithinRoot -Path $basePython -Root $Runtime.VenvRoot) -or
            -not (Test-Path -LiteralPath $basePython -PathType Leaf)
        ) {
            continue
        }

        try {
            & $basePython -c $validationScript *> $null
            $validationExitCode = $LASTEXITCODE
        } catch {
            continue
        }
        if ($validationExitCode -eq 0) {
            return $basePython
        }
    }

    throw (
        "Python 3.10 or newer was not found outside the Gift Card Recon runtime. " +
        "Install Python 3.10 or newer, then rerun setup."
    )
}

function Reset-GiftCardReconVenv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime,

        # Keeps the activated-target regression fixture independent of
        # ensurepip while exercising the same interpreter-selection path.
        [switch]$WithoutPip
    )

    $basePython = Resolve-GiftCardReconBasePython -Runtime $Runtime
    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $Runtime
    $venvArguments = @("-m", "venv", "--clear")
    if ($WithoutPip) {
        $venvArguments += "--without-pip"
    }
    $venvArguments += $Runtime.VenvRoot
    Invoke-GiftCardReconChecked -FilePath $basePython -Arguments $venvArguments
}

function Invoke-GiftCardReconRuntimeInitialization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot,

        [switch]$SkipInstall,

        [switch]$ForceInstall,

        [string]$RuntimeRoot = ""
    )

    $runtime = Get-GiftCardReconRuntime -RuntimeRoot $RuntimeRoot
    # Reject redirected runtime paths before probing the venv's Python or pip.
    # A linked runtime is unsafe even when no install appears necessary because
    # the executables reached through that path are outside our trust boundary.
    Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $runtime
    Set-GiftCardReconRuntimeEnvironment -Runtime $runtime
    # The runtime contains dependencies only. Source follows this invocation,
    # so a moved checkout needs no pip operation or network access.
    $env:PYTHONPATH = Join-Path ([IO.Path]::GetFullPath($ProgramRoot)) "src"
    $env:PYTHONNOUSERSITE = "1"
    $fingerprint = Get-GiftCardReconDependencyFingerprint -ProgramRoot $ProgramRoot
    $installedFingerprint = ""
    if (Test-Path -LiteralPath $runtime.DependencyFingerprintPath -PathType Leaf) {
        $rawFingerprint = Get-Content -LiteralPath $runtime.DependencyFingerprintPath -Raw
        if ($null -ne $rawFingerprint) {
            $installedFingerprint = $rawFingerprint.Trim()
        }
    }

    $runtimeValid = Test-GiftCardReconRuntime -Runtime $runtime
    $pythonUsable = $runtimeValid -or (Test-GiftCardReconPython -Runtime $runtime)
    $installRequired = $ForceInstall -or (-not $runtimeValid) -or ($installedFingerprint -ne $fingerprint)
    if ($SkipInstall -and $installRequired) {
        throw (
            "The local Gift Card Recon runtime is missing or out of date. " +
            "Run _program\install.ps1 explicitly before reporting."
        )
    }

    if ($installRequired) {
        if (-not $pythonUsable) {
            Write-Host "Preparing the local Gift Card Recon runtime beneath $($runtime.RuntimeRoot)..." -ForegroundColor Cyan
        } elseif (-not $runtimeValid) {
            Write-Host "Preparing a new runtime because its package set is incomplete..." -ForegroundColor Cyan
        } elseif ($installedFingerprint -ne $fingerprint) {
            Write-Host "Preparing a new runtime because its dependency specification changed..." -ForegroundColor Cyan
        } else {
            Write-Host "Preparing a new runtime because setup was forced..." -ForegroundColor Cyan
        }

        # Never clear or rename the working environment. Windows entrypoints
        # embed absolute paths, so a candidate is built at its permanent path.
        # Termination before the atomic pointer switch leaves the old runtime
        # active; termination after it leaves the validated candidate active.
        Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $runtime
        $generation = Join-Path $runtime.RuntimeRoot ("environments\{0}-{1}" -f $fingerprint.Substring(0, [Math]::Min(12, $fingerprint.Length)), [Guid]::NewGuid().ToString("N"))
        $candidate = $runtime.PSObject.Copy()
        $candidate.VenvRoot = $generation
        $candidate.PythonPath = Join-Path $generation "Scripts\python.exe"
        $candidate.DependencyFingerprintPath = Join-Path $generation "dependency-fingerprint.sha256"
        Reset-GiftCardReconVenv -Runtime $candidate
        Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $candidate
        Invoke-GiftCardReconChecked -FilePath $candidate.PythonPath -Arguments @(
            "-m", "pip", "install", "--disable-pip-version-check", "--require-hashes", "-r",
            (Join-Path $ProgramRoot "requirements.lock")
        )

        if (-not (Test-GiftCardReconRuntime -Runtime $candidate)) {
            throw "The Gift Card Recon runtime did not validate after installation."
        }
        Assert-GiftCardReconVenvRootIsSafeToModify -Runtime $candidate
        Set-Content -LiteralPath $candidate.DependencyFingerprintPath -Value $fingerprint -NoNewline -Encoding ASCII
        $pointer = Join-Path $runtime.RuntimeRoot "active-runtime.json"
        Write-GiftCardReconAtomicJson -Path $pointer -Value ([ordered]@{
            schema = 1
            venv_root = $candidate.VenvRoot
            previous_venv_root = $runtime.VenvRoot
            dependency_fingerprint = $fingerprint
        })
        $runtime = $candidate
    }

    return $runtime
}

function Initialize-GiftCardReconRuntime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot,

        [switch]$SkipInstall,

        [switch]$ForceInstall,

        [string]$RuntimeRoot = ""
    )

    $mutex = [System.Threading.Mutex]::new($false, "Local\GiftCardReconRuntimeInstall")
    $ownsMutex = $false
    try {
        try {
            $ownsMutex = $mutex.WaitOne([TimeSpan]::FromMinutes(10))
        } catch [System.Threading.AbandonedMutexException] {
            # The prior installer ended unexpectedly; this process now owns the mutex.
            $ownsMutex = $true
        }
        if (-not $ownsMutex) {
            throw "Timed out waiting for another Gift Card Recon setup process to finish."
        }

        return Invoke-GiftCardReconRuntimeInitialization `
            -ProgramRoot $ProgramRoot `
            -SkipInstall:$SkipInstall `
            -ForceInstall:$ForceInstall `
            -RuntimeRoot $RuntimeRoot
    } finally {
        if ($ownsMutex) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
}
