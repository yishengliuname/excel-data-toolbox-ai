$ErrorActionPreference = 'Stop'

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$serverPath = Join-Path $projectDir 'server.py'
$bundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$pythonPath = $null

if (Test-Path -LiteralPath $bundledPython) {
    $pythonPath = $bundledPython
}

if (-not $pythonPath) {
    $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        $pythonPath = $pyLauncher.Source
    }
}

if (-not $pythonPath) {
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand -and $pythonCommand.Source -notlike '*WindowsApps*') {
        $pythonPath = $pythonCommand.Source
    }
}

if (-not $pythonPath) {
    Write-Host '没有找到可用的 Python。' -ForegroundColor Red
    Write-Host '请安装 Python 3.12，再在本目录执行：python -m pip install -r requirements.txt'
    Read-Host '按回车键退出'
    exit 1
}

Push-Location $projectDir
try {
    if ((Split-Path -Leaf $pythonPath) -ieq 'py.exe') {
        & $pythonPath -3.12 -c "import pandas, openpyxl" 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Host '缺少 pandas 或 openpyxl。请先运行：py -3.12 -m pip install -r requirements.txt' -ForegroundColor Yellow
            Read-Host '按回车键退出'
            exit 1
        }
        & $pythonPath -3.12 $serverPath
    }
    else {
        & $pythonPath -c "import pandas, openpyxl" 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Host '缺少 pandas 或 openpyxl。请先运行：python -m pip install -r requirements.txt' -ForegroundColor Yellow
            Read-Host '按回车键退出'
            exit 1
        }
        & $pythonPath $serverPath
    }
}
finally {
    Pop-Location
}
