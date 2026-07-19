$ErrorActionPreference = "Stop"

function Get-GiftCardReconRuntime {
    [CmdletBinding()]
    param()

    $localAppData = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    }
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw "Windows Local AppData could not be resolved for the Gift Card Recon runtime."
    }

    $runtimeRoot = Join-Path $localAppData "GiftCardRecon"
    $cacheRoot = Join-Path $runtimeRoot "cache"
    $tempRoot = Join-Path $runtimeRoot "temp"
    $venvRoot = Join-Path $runtimeRoot "venv"

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
        DependencyFingerprintPath = Join-Path $runtimeRoot "dependency-fingerprint.sha256"
    }
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
        (Join-Path $resolvedProgramRoot "requirements.txt"),
        (Join-Path $resolvedProgramRoot "pyproject.toml")
    )
    $parts = @(
        "runtime-schema=1",
        "program-root=$($resolvedProgramRoot.ToLowerInvariant())"
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

    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    $item = Get-Item -LiteralPath $Path -Force
    return (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Assert-GiftCardReconVenvRootIsSafeToClear {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    if (Test-GiftCardReconReparsePoint -Path $Runtime.VenvRoot) {
        throw (
            "Refusing to rebuild the Gift Card Recon runtime because $($Runtime.VenvRoot) " +
            "is a link, junction, or other reparse point. Remove that entry manually, " +
            "then rerun setup."
        )
    }
}

function Test-GiftCardReconPython {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    if (
        -not (Test-Path -LiteralPath $Runtime.PythonPath -PathType Leaf) -or
        -not (Test-Path -LiteralPath (Join-Path $Runtime.VenvRoot "pyvenv.cfg") -PathType Leaf)
    ) {
        return $false
    }
    try {
        & $Runtime.PythonPath --version *> $null
        if ($LASTEXITCODE -ne 0) {
            return $false
        }
        & $Runtime.PythonPath -m pip --version *> $null
        if ($LASTEXITCODE -ne 0) {
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

function Test-GiftCardReconRuntime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime
    )

    if (-not (Test-GiftCardReconPython -Runtime $Runtime)) {
        return $false
    }
    try {
        & $Runtime.PythonPath -m pip check *> $null
        if ($LASTEXITCODE -ne 0) {
            return $false
        }
        & $Runtime.PythonPath -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('gift_card_recon') else 1)" *> $null
        return $LASTEXITCODE -eq 0
    } catch {
        return $false
    }
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

function Invoke-GiftCardReconRuntimeInitialization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot,

        [switch]$SkipInstall,

        [switch]$ForceInstall
    )

    $runtime = Get-GiftCardReconRuntime
    Set-GiftCardReconRuntimeEnvironment -Runtime $runtime
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
            "Run _program\install.ps1, or rerun without -SkipInstall."
        )
    }

    if ($installRequired) {
        if (-not $pythonUsable) {
            $systemPython = Get-Command python -ErrorAction SilentlyContinue
            if ($null -eq $systemPython) {
                throw "Python was not found. Install Python 3.10 or newer, then rerun setup."
            }
            Assert-GiftCardReconVenvRootIsSafeToClear -Runtime $runtime
            Write-Host "Creating the local Gift Card Recon runtime at $($runtime.VenvRoot)..." -ForegroundColor Cyan
            Invoke-GiftCardReconChecked -FilePath $systemPython.Source -Arguments @(
                "-m", "venv", "--clear", $runtime.VenvRoot
            )
        } elseif (-not $runtimeValid) {
            Write-Host "Repairing missing or incomplete local runtime packages..." -ForegroundColor Cyan
        } else {
            Write-Host "Refreshing the local runtime because its dependency specification changed..." -ForegroundColor Cyan
        }

        Invoke-GiftCardReconChecked -FilePath $runtime.PythonPath -Arguments @(
            "-m", "pip", "install", "--disable-pip-version-check", "--upgrade", "pip"
        )
        Invoke-GiftCardReconChecked -FilePath $runtime.PythonPath -Arguments @(
            "-m", "pip", "install", "--disable-pip-version-check", "-r",
            (Join-Path $ProgramRoot "requirements.txt")
        )
        Invoke-GiftCardReconChecked -FilePath $runtime.PythonPath -Arguments @(
            "-m", "pip", "install", "--disable-pip-version-check", "-e", $ProgramRoot
        )

        if (-not (Test-GiftCardReconRuntime -Runtime $runtime)) {
            throw "The Gift Card Recon runtime did not validate after installation."
        }
        $temporaryFingerprint = "$($runtime.DependencyFingerprintPath).$PID.tmp"
        try {
            Set-Content -LiteralPath $temporaryFingerprint -Value $fingerprint -NoNewline -Encoding ASCII
            Move-Item -LiteralPath $temporaryFingerprint -Destination $runtime.DependencyFingerprintPath -Force
        } finally {
            Remove-Item -LiteralPath $temporaryFingerprint -Force -ErrorAction SilentlyContinue
        }
    }

    return $runtime
}

function Initialize-GiftCardReconRuntime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot,

        [switch]$SkipInstall,

        [switch]$ForceInstall
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
            -ForceInstall:$ForceInstall
    } finally {
        if ($ownsMutex) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
}
