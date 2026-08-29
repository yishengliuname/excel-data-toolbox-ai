@echo off
setlocal EnableExtensions
title Excel Data Toolbox
cd /d "%~dp0"

set "APP_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "APP_PY_ARGS="
if exist "%APP_PYTHON%" goto python_found

set "APP_PYTHON=C:\Users\liuyisheng\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%APP_PYTHON%" goto python_found

where py.exe >nul 2>nul
if not errorlevel 1 (
  set "APP_PYTHON=py.exe"
  set "APP_PY_ARGS=-3"
  goto python_found
)

where python.exe >nul 2>nul
if not errorlevel 1 (
  set "APP_PYTHON=python.exe"
  goto python_found
)

set "PACKAGED_APP=%~dp0dist\BiaogeKuaichuAI\BiaogeKuaichuAI.exe"
if exist "%PACKAGED_APP%" (
  echo Python was not found. Starting the packaged fallback...
  "%PACKAGED_APP%"
  exit /b %ERRORLEVEL%
)

echo.
echo [START FAILED] Python was not found.
echo Install Python 3.10 or newer, then run:
echo python -m pip install -r requirements.lock -r requirements-optional.txt
echo.
pause
exit /b 1

:python_found
echo.
echo ============================================
echo   Excel Data Toolbox is starting...
echo   Keep this window open while using the app.
echo ============================================
echo.

"%APP_PYTHON%" %APP_PY_ARGS% -c "import pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo [START FAILED] pandas or openpyxl is missing.
  echo Run: "%APP_PYTHON%" %APP_PY_ARGS% -m pip install -r requirements.lock -r requirements-optional.txt
  echo.
  pause
  exit /b 1
)

echo If the browser does not open, visit: http://127.0.0.1:8501
echo.
"%APP_PYTHON%" %APP_PY_ARGS% "%~dp0server.py" 2>>"%~dp0startup-error.log"
set "APP_EXIT_CODE=%ERRORLEVEL%"

if not "%APP_EXIT_CODE%"=="0" (
  echo.
  echo [APP STOPPED] Details were saved to startup-error.log
  echo.
  pause
)
exit /b %APP_EXIT_CODE%
