$ErrorActionPreference = 'Stop'

# Starts the app from a source checkout.
#
# Interpreter selection is deliberately generic: a local .venv first, then the
# Windows py launcher, then python.exe on PATH.  No machine-specific path and no
# hardcoded minor version is used, so this file works on any Windows machine
# that has Python 3.11+ with the project dependencies installed.

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$serverPath = Join-Path $projectDir 'server.py'
$projectLeaf = Split-Path -Leaf $projectDir
$interactive = -not $env:EDT_NO_PAUSE

function Stop-WithMessage {
    param([string[]]$Lines)
    foreach ($line in $Lines) {
        if ($line -eq '') { Write-Host '' } else { Write-Host $line }
    }
    if ($interactive) { Read-Host 'Press Enter to exit' | Out-Null }
    exit 1
}

# Runs a python probe and returns its exit code.
#
# PowerShell 5.1 turns native stderr into a terminating error while
# $ErrorActionPreference is 'Stop', so a failing 'import pandas' would abort the
# script instead of merely reporting a non-zero exit code.  Probe with 'Continue'
# and restore the caller's preference afterwards.
function Invoke-PythonProbe {
    param(
        [string]$Executable,
        [string[]]$PrefixArgs = @(),
        [string]$Code
    )
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $Executable @PrefixArgs -c $Code 1>$null 2>$null
        $exitCode = $LASTEXITCODE
    }
    catch {
        $exitCode = 1
    }
    finally {
        $ErrorActionPreference = $previous
    }
    return $exitCode
}

function Test-Interpreter {
    param([string]$Executable, [string[]]$PrefixArgs = @())
    if (-not $Executable) { return $false }
    if ($Executable -like '*WindowsApps*') { return $false }

    $versionCode = Invoke-PythonProbe $Executable $PrefixArgs 'import sys; raise SystemExit(0 if sys.version_info[:2] >= (3, 11) else 3)'
    if ($versionCode -ne 0) { return $false }

    $dependencyCode = Invoke-PythonProbe $Executable $PrefixArgs 'import pandas, openpyxl'
    return ($dependencyCode -eq 0)
}

# --- candidate 1: the project virtual environment ---------------------------
$pythonPath = $null
$pythonArgs = @()
$venvPython = Join-Path $projectDir '.venv\Scripts\python.exe'
if (Test-Path -LiteralPath $venvPython) { $pythonPath = $venvPython }

# --- candidate 2: the Windows py launcher (whatever 3.x it resolves) --------
if (-not $pythonPath) {
    $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyLauncher -and (Test-Interpreter $pyLauncher.Source @('-3'))) {
        $pythonPath = $pyLauncher.Source
        $pythonArgs = @('-3')
    }
}

# --- candidate 3: python.exe on PATH ---------------------------------------
if (-not $pythonPath) {
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand -and (Test-Interpreter $pythonCommand.Source)) {
        $pythonPath = $pythonCommand.Source
    }
}

if (-not $pythonPath) {
    Stop-WithMessage @(
        'No usable Python found.',
        '',
        'Either download the packaged Windows build from the Releases page,',
        'or set up the source checkout:',
        '',
        '  py -3 -m venv .venv',
        '  .\.venv\Scripts\python.exe -m pip install -e ".[automation]"',
        '  .\start_app.ps1',
        ''
    )
}

if (-not (Test-Interpreter $pythonPath $pythonArgs)) {
    Stop-WithMessage @(
        "This interpreter is missing the project dependencies: $pythonPath",
        '',
        'Install them with:',
        '',
        "  `"$pythonPath`" -m pip install -e `".[automation]`"",
        ''
    )
}

# Decide how to start: importable package module first, then server.py.
Push-Location $projectDir
try {
    $moduleCode = Invoke-PythonProbe $pythonPath $pythonArgs 'import excel_data_toolbox.server'
}
finally {
    Pop-Location
}
$moduleImportable = ($moduleCode -eq 0)

if ($moduleImportable) {
    $runDir = $projectDir
    $runModule = $true
}
elseif ($projectLeaf -ieq 'excel_data_toolbox') {
    # The checkout folder already carries the importable package name.
    $runDir = Split-Path -Parent $projectDir
    $runModule = $true
}
else {
    $runDir = $projectDir
    $runModule = $false
}

Write-Host ''
Write-Host 'Starting Excel Data Toolbox...'
Write-Host "Interpreter: $pythonPath $($pythonArgs -join ' ')"
Write-Host 'Keep this window open while the app is running.'
Write-Host ''

Push-Location $runDir
try {
    # The server logs to stderr while running; keep native output from aborting us.
    $ErrorActionPreference = 'Continue'
    if ($runModule) {
        & $pythonPath @pythonArgs -m excel_data_toolbox.server @args
    }
    else {
        & $pythonPath @pythonArgs $serverPath @args
    }
    $exitCode = $LASTEXITCODE
}
finally {
    Pop-Location
}

if ($exitCode -ne 0) {
    Write-Host ''
    Write-Host "The application exited with code $exitCode." -ForegroundColor Red
    if (-not $runModule) {
        Write-Host 'If this is an import error, install the package first:' -ForegroundColor Yellow
        Write-Host "  `"$pythonPath`" -m pip install -e `".[automation]`"" -ForegroundColor Yellow
    }
    if ($interactive) { Read-Host 'Press Enter to exit' | Out-Null }
}
exit $exitCode
