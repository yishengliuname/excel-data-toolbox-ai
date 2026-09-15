param(
    [switch]$RecreateEnvironment,
    [switch]$UseCurrentEnvironment,
    [switch]$SkipSmokeTest
)

$ErrorActionPreference = 'Stop'

$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$buildEnvironment = Join-Path $projectDir '.build-venv'
$dist = Join-Path $projectDir 'dist'
$appDist = Join-Path $dist 'BiaogeKuaichuAI'
$archive = Join-Path $dist 'Excel-Data-Toolbox-AI-Windows-x64.zip'

function Resolve-PythonExecutable {
    $projectPython = Join-Path $projectDir '.venv\Scripts\python.exe'
    if (Test-Path -LiteralPath $projectPython) {
        return $projectPython
    }

    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        try {
            $resolved = (& $pythonCommand.Source -c "import sys; print(sys.executable)").Trim()
            if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $resolved)) {
                return $resolved
            }
        }
        catch {
            # Microsoft Store can leave a non-executable python.exe alias on PATH.
        }
    }

    $launcher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($launcher) {
        try {
            $resolved = (& $launcher.Source -3 -c "import sys; print(sys.executable)").Trim()
            if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $resolved)) {
                return $resolved
            }
        }
        catch {
            # Continue to the actionable error below.
        }
    }

    throw 'No runnable Python 3 interpreter was found. Install Python 3.11+ or activate a virtual environment.'
}

Push-Location $projectDir
try {
    $env:PYTHONNOUSERSITE = '1'
    $env:PYTHONUTF8 = '1'
    $env:PYTHONIOENCODING = 'utf-8'
    $env:PYTHONUSERBASE = Join-Path $projectDir '.build_userbase'

    if ($UseCurrentEnvironment) {
        $python = Resolve-PythonExecutable
    }
    else {
        $resolvedBuildEnvironment = [System.IO.Path]::GetFullPath($buildEnvironment)
        $resolvedProject = [System.IO.Path]::GetFullPath($projectDir)
        if (-not $resolvedBuildEnvironment.StartsWith($resolvedProject + [System.IO.Path]::DirectorySeparatorChar)) {
            throw 'Refusing to manage a build environment outside the project directory.'
        }
        if ($RecreateEnvironment -and (Test-Path -LiteralPath $resolvedBuildEnvironment)) {
            Remove-Item -LiteralPath $resolvedBuildEnvironment -Recurse -Force
        }
        $python = Join-Path $resolvedBuildEnvironment 'Scripts\python.exe'
        if (-not (Test-Path -LiteralPath $python)) {
            $bootstrap = Resolve-PythonExecutable
            & $bootstrap -m venv $resolvedBuildEnvironment
            if ($LASTEXITCODE -ne 0) { throw 'Failed to create the isolated build environment.' }
        }
        & $python -m pip install --upgrade pip
        if ($LASTEXITCODE -ne 0) { throw 'Failed to upgrade pip in the build environment.' }
        & $python -m pip install -e "${projectDir}[automation,dev]"
        if ($LASTEXITCODE -ne 0) { throw 'Failed to install build dependencies.' }
    }

    $version = (& $python scripts/check_release_version.py --print-version).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Failed to read the project version.' }
    & $python scripts/check_release_version.py --tag "v$version"
    if ($LASTEXITCODE -ne 0) { throw 'Release version validation failed.' }
    & $python -m PyInstaller --noconfirm --clean BiaogeKuaichuAI.spec
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed.' }

    $executable = Join-Path $appDist 'BiaogeKuaichuAI.exe'
    $internal = Join-Path $appDist '_internal'
    $requiredFiles = @(
        $executable,
        (Join-Path $internal 'web\unified.html'),
        (Join-Path $internal 'domain_packs.json')
    )
    foreach ($requiredFile in $requiredFiles) {
        if (-not (Test-Path -LiteralPath $requiredFile)) {
            throw "Windows package is missing required runtime file: $requiredFile"
        }
    }

    Copy-Item -LiteralPath (Join-Path $projectDir 'LICENSE') -Destination $appDist -Force
    Copy-Item -LiteralPath (Join-Path $projectDir 'NOTICE') -Destination $appDist -Force
    Copy-Item -LiteralPath (Join-Path $projectDir 'docs\WINDOWS_RELEASE_README.txt') `
        -Destination (Join-Path $appDist 'README.txt') -Force

    if (-not $SkipSmokeTest) {
        & $python scripts/smoke_test_windows_build.py --app-dir $appDist --expected-version $version
        if ($LASTEXITCODE -ne 0) { throw 'Packaged application smoke test failed.' }
    }

    if (Test-Path -LiteralPath $archive) {
        Remove-Item -LiteralPath $archive -Force
    }
    Compress-Archive -Path (Join-Path $appDist '*') -DestinationPath $archive -CompressionLevel Optimal

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($archive)
    try {
        $forbidden = $zip.Entries | Where-Object {
            $_.FullName -match '(^|/)(\.env|user_data|outputs|logs|\.pytest_cache|\.test_cache)(/|$)' -or
            $_.FullName -match '\.(xlsx|xlsm|csv|sqlite|db|log)$'
        }
        if ($forbidden) {
            throw "Release archive contains forbidden paths: $($forbidden.FullName -join ', ')"
        }
    }
    finally {
        $zip.Dispose()
    }

    Write-Host "Windows release ready: $archive" -ForegroundColor Green
}
finally {
    Pop-Location
}
