@echo off
setlocal EnableExtensions
title Excel Data Toolbox
cd /d "%~dp0"

rem Starts the app from a source checkout.
rem Interpreter order: local .venv, then the py launcher, then python.exe.
rem No machine-specific path and no hardcoded minor version is used.

set "APP_PYTHON="
set "APP_PY_ARGS="
set "LEAF="
for %%I in ("%~dp0.") do set "LEAF=%%~nxI"

if exist "%~dp0.venv\Scripts\python.exe" (
  set "APP_PYTHON=%~dp0.venv\Scripts\python.exe"
) else (
  where py.exe >nul 2>nul
  if not errorlevel 1 (
    set "APP_PYTHON=py.exe"
    set "APP_PY_ARGS=-3"
  ) else (
    where python.exe >nul 2>nul
    if not errorlevel 1 set "APP_PYTHON=python.exe"
  )
)

if not defined APP_PYTHON goto no_python

"%APP_PYTHON%" %APP_PY_ARGS% -c "import sys; sys.exit(0 if sys.version_info>=(3,11) else 3)" >nul 2>nul
if errorlevel 1 goto bad_python

"%APP_PYTHON%" %APP_PY_ARGS% -c "import pandas, openpyxl" >nul 2>nul
if errorlevel 1 goto missing_deps

echo.
echo ============================================
echo   Excel Data Toolbox is starting...
echo   Keep this window open while using the app.
echo ============================================
echo.
echo If the browser does not open, visit: http://127.0.0.1:8501
echo.

"%APP_PYTHON%" %APP_PY_ARGS% -c "import excel_data_toolbox.server" >nul 2>nul
if not errorlevel 1 goto run_module
if /i "%LEAF%"=="excel_data_toolbox" goto run_module_parent
goto run_script

:run_module
"%APP_PYTHON%" %APP_PY_ARGS% -m excel_data_toolbox.server 2>>"%~dp0startup-error.log"
goto finished

:run_module_parent
pushd "%~dp0.."
"%APP_PYTHON%" %APP_PY_ARGS% -m excel_data_toolbox.server 2>>"%~dp0startup-error.log"
popd
goto finished

:run_script
"%APP_PYTHON%" %APP_PY_ARGS% "%~dp0server.py" 2>>"%~dp0startup-error.log"
goto finished

:finished
set "APP_EXIT_CODE=%ERRORLEVEL%"
if not "%APP_EXIT_CODE%"=="0" (
  echo.
  echo [APP STOPPED] Exit code %APP_EXIT_CODE%.
  echo Details were appended to startup-error.log.
  echo.
  echo If this is an import or missing-module error, install the package:
  echo   "%APP_PYTHON%" %APP_PY_ARGS% -m pip install -e ".[automation]"
  echo.
  pause
)
exit /b %APP_EXIT_CODE%

:bad_python
echo.
echo [START FAILED] Python 3.11 or newer is required.
echo Found: "%APP_PYTHON%" %APP_PY_ARGS%
echo.
echo Install Python 3.11+ and retry, or download the packaged build
echo from the Releases page.
echo.
pause
exit /b 1

:missing_deps
echo.
echo [START FAILED] pandas or openpyxl is missing for this interpreter:
echo   "%APP_PYTHON%" %APP_PY_ARGS%
echo.
echo Install the project dependencies with:
echo   "%APP_PYTHON%" %APP_PY_ARGS% -m pip install -e ".[automation]"
echo.
pause
exit /b 1

:no_python
if exist "%~dp0BiaogeKuaichuAI.exe" (
  echo Python was not found. Starting the packaged build...
  "%~dp0BiaogeKuaichuAI.exe"
  exit /b %ERRORLEVEL%
)
if exist "%~dp0dist\BiaogeKuaichuAI\BiaogeKuaichuAI.exe" (
  echo Python was not found. Starting the packaged build...
  "%~dp0dist\BiaogeKuaichuAI\BiaogeKuaichuAI.exe"
  exit /b %ERRORLEVEL%
)
echo.
echo [START FAILED] No usable Python found.
echo.
echo Either download the packaged Windows build from the Releases page,
echo or set up the source checkout:
echo.
echo   py -3 -m venv .venv
echo   .venv\Scripts\python.exe -m pip install -e ".[automation]"
echo   "%~nx0"
echo.
pause
exit /b 1
