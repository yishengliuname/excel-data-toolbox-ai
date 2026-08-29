$ErrorActionPreference = 'Stop'

$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$entry = Join-Path $projectDir 'server.py'
$dist = Join-Path $projectDir 'dist'
$bundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if (Test-Path -LiteralPath $bundledPython) {
    $python = $bundledPython
}
else {
    $python = (Get-Command python.exe -ErrorAction Stop).Source
}

Push-Location $projectDir
try {
    $env:PYTHONNOUSERSITE = '1'
    # PyInstaller queries Python's user-site path even when it is disabled.
    # Point that read-only probe at a build-local path to avoid locked roaming
    # profile directories on managed Windows machines.
    $env:PYTHONUSERBASE = Join-Path $projectDir '.build_userbase'
    & $python -m pip install -r requirements.lock -r requirements-optional.txt -r requirements-build.txt
    & $python scripts/health_check.py
    & $python -s -m PyInstaller --noconfirm --clean --name 'BiaogeKuaichuAI' --onedir `
        --add-data 'web;web' `
        --hidden-import duckdb --hidden-import pdfplumber --hidden-import pytesseract `
        --hidden-import pyodbc $entry
    $appDist = Join-Path $dist 'BiaogeKuaichuAI'
    $ocrSource = Join-Path ${env:ProgramFiles} 'Tesseract-OCR'
    $ocrTarget = Join-Path $appDist 'tesseract'
    if (Test-Path -LiteralPath $ocrSource) {
        Copy-Item -LiteralPath $ocrSource -Destination $ocrTarget -Recurse -Force
        $localLanguages = Join-Path $env:LOCALAPPDATA 'BiaogeKuaichu\tessdata'
        if (Test-Path -LiteralPath $localLanguages) {
            Copy-Item -Path (Join-Path $localLanguages '*.traineddata') `
                -Destination (Join-Path $ocrTarget 'tessdata') -Force
        }
    }
    else {
        Write-Warning '未找到 Tesseract-OCR；发布包的图片 OCR 需要目标电脑另行安装。'
    }
    Write-Host "构建完成：$dist\BiaogeKuaichuAI" -ForegroundColor Green
}
finally {
    Pop-Location
}
