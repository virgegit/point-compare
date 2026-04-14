@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..") do set "TOOLS_ROOT=%%~fI"
set "ENV_FILE=%TOOLS_ROOT%\.env"
chcp 65001 >nul
set "PYTHONUTF8=1"

if not exist "%ENV_FILE%" (
    echo [ERROR] Workspace .env file was not found:
    echo         %ENV_FILE%
    goto :fail
)

for /f "usebackq tokens=1,* delims==" %%A in ("%ENV_FILE%") do (
    if /I "%%~A"=="PYTHON" set "PYTHON=%%~B"
)

if not defined PYTHON (
    echo [ERROR] PYTHON was not found in:
    echo         %ENV_FILE%
    goto :fail
)

if not exist "%PYTHON%" (
    echo [ERROR] PYTHON points to a missing file:
    echo         %PYTHON%
    goto :fail
)

pushd "%SCRIPT_DIR%"
start "" "%PYTHON%" "%SCRIPT_DIR%compare_points_gui.py"
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" goto :fail_with_code
exit /b 0

:fail_with_code
echo.
echo Launcher finished with exit code %EXIT_CODE%.
pause
exit /b %EXIT_CODE%

:fail
echo.
pause
exit /b 1
