@echo off
setlocal

title Healthcare Research Team Tools Installer

set "SCRIPT_DIR=%~dp0"
set "INSTALLER=%SCRIPT_DIR%Install Healthcare Research Team Tools.ps1"

if not exist "%INSTALLER%" (
    echo.
    echo Installation could not start.
    echo.
    echo The following file is missing:
    echo Install Healthcare Research Team Tools.ps1
    echo.
    echo Extract all files from the ZIP into the same folder,
    echo then run this installer again.
    echo.
    pause
    exit /b 1
)

powershell.exe ^
  -NoLogo ^
  -NoProfile ^
  -ExecutionPolicy Bypass ^
  -File "%INSTALLER%"

set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Installation did not complete.
    echo.
    echo Review the error message and try again.
    echo.
    pause
)

exit /b %EXIT_CODE%