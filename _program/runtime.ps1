$ErrorActionPreference = "Stop"

function Get-GiftCardReconRuntime {
    [CmdletBinding()]
    param(
        [ValidateSet("Operator", "Development")]
        [string]$Profile = "Operator"
    )

    $localAppData = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    }
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw "Windows Local AppData could not be resolved for the Gift Card Recon runtime."
    }

    $runtimeBase = Join-Path $localAppData "GiftCardRecon"
    $profileFolder = if ($Profile -eq "Development") { "development" } else { "operator" }
    $runtimeRoot = Join-Path $runtimeBase $profileFolder
    $cacheRoot = Join-Path $runtimeRoot "cache"
    $tempRoot = Join-Path $runtimeRoot "temp"
    $venvRoot = Join-Path $runtimeRoot "venv"

    return [pscustomobject]@{
        Profile = $Profile
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
        [string]$ProgramRoot,

        [ValidateSet("Operator", "Development")]
        [string]$Profile = "Operator"
    )

    $resolvedProgramRoot = [IO.Path]::GetFullPath($ProgramRoot).TrimEnd('\', '/')
    $specifications = @(
        (Join-Path $resolvedProgramRoot "requirements.txt"),
        (Join-Path $resolvedProgramRoot "pyproject.toml")
    )
    $parts = @(
        "runtime-schema=2",
        "profile=$($Profile.ToLowerInvariant())"
    )
    if ($Profile -eq "Development") {
        # Editable development installs are intentionally tied to one checkout.
        $parts += "program-root=$($resolvedProgramRoot.ToLowerInvariant())"
    }
    foreach ($path in $specifications) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Runtime dependency specification is missing: $path"
        }
        $digest = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        $parts += "$(Split-Path -Leaf $path)=$digest"
    }

    if ($Profile -eq "Operator") {
        # Operator installs are copied into site-packages, so fingerprint the
        # application payload without binding the runtime to a Dropbox path.
        $sourceRoot = Join-Path $resolvedProgramRoot "src"
        if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
            throw "Runtime application source is missing: $sourceRoot"
        }
        $sourceParts = foreach ($item in Get-ChildItem -LiteralPath $sourceRoot -Recurse -File | Sort-Object FullName) {
            if (
                $item.Name -like "*.pyc" -or
                $item.FullName -match "[\\/]__pycache__[\\/]" -or
                $item.FullName -match "[\\/][^\\/]+\.egg-info[\\/]"
            ) {
                continue
            }
            $relative = $item.FullName.Substring($resolvedProgramRoot.Length).TrimStart([char[]]"\/").Replace("\", "/")
            $digest = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            "$($relative.ToLowerInvariant())=$digest"
        }
        $parts += $sourceParts
    }

    $payload = [Text.Encoding]::UTF8.GetBytes(($parts -join "`n"))
    $hasher = [Security.Cryptography.SHA256]::Create()
    try {
        return -join ($hasher.ComputeHash($payload) | ForEach-Object { $_.ToString("x2") })
    } finally {
        $hasher.Dispose()
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

function Install-GiftCardReconOperatorPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Runtime,

        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot
    )

    $stagingBase = [IO.Path]::GetFullPath($Runtime.TempRoot).TrimEnd('\', '/')
    $stagingRoot = [IO.Path]::GetFullPath(
        (Join-Path $stagingBase ("operator-package-" + [guid]::NewGuid().ToString("N")))
    )
    $stagingPrefix = $stagingBase + [IO.Path]::DirectorySeparatorChar
    if (-not $stagingRoot.StartsWith($stagingPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing unsafe operator package staging path: $stagingRoot"
    }

    try {
        New-Item -ItemType Directory -Force -Path $stagingRoot | Out-Null
        foreach ($name in @("pyproject.toml", "requirements.txt")) {
            $source = Join-Path $ProgramRoot $name
            if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
                throw "Operator package input is missing: $source"
            }
            Copy-Item -LiteralPath $source -Destination (Join-Path $stagingRoot $name) -Force
        }
        $sourceRoot = Join-Path $ProgramRoot "src"
        if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
            throw "Operator package source is missing: $sourceRoot"
        }
        $stagedSourceRoot = Join-Path $stagingRoot "src"
        New-Item -ItemType Directory -Force -Path $stagedSourceRoot | Out-Null
        foreach ($item in Get-ChildItem -LiteralPath $sourceRoot -Recurse -File | Sort-Object FullName) {
            if (
                $item.Name -like "*.pyc" -or
                $item.FullName -match "[\\/]__pycache__[\\/]" -or
                $item.FullName -match "[\\/][^\\/]+\.egg-info[\\/]"
            ) {
                continue
            }
            $relative = $item.FullName.Substring($sourceRoot.Length).TrimStart([char[]]"\/")
            $destination = Join-Path $stagedSourceRoot $relative
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
            Copy-Item -LiteralPath $item.FullName -Destination $destination -Force
        }

        Invoke-GiftCardReconChecked -FilePath $Runtime.PythonPath -Arguments @(
            "-m", "pip", "install", "--disable-pip-version-check", "--force-reinstall", "--no-deps", $stagingRoot
        )
    }
    finally {
        if (Test-Path -LiteralPath $stagingRoot) {
            if (-not $stagingRoot.StartsWith($stagingPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Refusing unsafe operator package staging cleanup: $stagingRoot"
            }
            Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-GiftCardReconRuntimeInitialization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramRoot,

        [ValidateSet("Operator", "Development")]
        [string]$Profile = "Operator",

        [switch]$SkipInstall,

        [switch]$ForceInstall
    )

    $runtime = Get-GiftCardReconRuntime -Profile $Profile
    Set-GiftCardReconRuntimeEnvironment -Runtime $runtime
    $fingerprint = Get-GiftCardReconDependencyFingerprint -ProgramRoot $ProgramRoot -Profile $Profile
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
        $recovery = if ($Profile -eq "Operator") {
            "Run _program\install.ps1, or rerun without -SkipInstall."
        }
        else {
            "Rerun _program\run_tests.ps1 without -SkipInstall."
        }
        throw (
            "The $($Profile.ToLowerInvariant()) Gift Card Recon runtime is missing or out of date. " +
            $recovery
        )
    }

    if ($installRequired) {
        if (-not $pythonUsable) {
            $systemPython = Get-Command python -ErrorAction SilentlyContinue
            if ($null -eq $systemPython) {
                throw "Python was not found. Install Python 3.10 or newer, then rerun setup."
            }
            Write-Host "Creating the $($Profile.ToLowerInvariant()) Gift Card Recon runtime at $($runtime.VenvRoot)..." -ForegroundColor Cyan
            Invoke-GiftCardReconChecked -FilePath $systemPython.Source -Arguments @(
                "-m", "venv", "--clear", $runtime.VenvRoot
            )
        } elseif (-not $runtimeValid) {
            Write-Host "Repairing missing or incomplete local runtime packages..." -ForegroundColor Cyan
        } else {
            Write-Host "Refreshing the local runtime because its application payload or dependency specification changed..." -ForegroundColor Cyan
        }

        Invoke-GiftCardReconChecked -FilePath $runtime.PythonPath -Arguments @(
            "-m", "pip", "install", "--disable-pip-version-check", "-r",
            (Join-Path $ProgramRoot "requirements.txt")
        )
        if ($Profile -eq "Development") {
            Invoke-GiftCardReconChecked -FilePath $runtime.PythonPath -Arguments @(
                "-m", "pip", "install", "--disable-pip-version-check", "-e", $ProgramRoot
            )
        }
        else {
            Install-GiftCardReconOperatorPackage -Runtime $runtime -ProgramRoot $ProgramRoot
        }

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

        [ValidateSet("Operator", "Development")]
        [string]$Profile = "Operator",

        [switch]$SkipInstall,

        [switch]$ForceInstall
    )

    $mutex = [System.Threading.Mutex]::new($false, "Local\GiftCardReconRuntimeInstall$Profile")
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
            -Profile $Profile `
            -SkipInstall:$SkipInstall `
            -ForceInstall:$ForceInstall
    } finally {
        if ($ownsMutex) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
}
