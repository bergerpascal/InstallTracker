@echo off
REM InstallTracker - Launcher Script
REM This batch file starts InstallTracker with proper PowerShell settings
REM
REM Features:
REM - Sets ExecutionPolicy to Bypass for this session
REM - Launches InstallTracker.ps1 from the same directory
REM - Handles errors gracefully

setlocal enabledelayedexpansion

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%InstallTracker.ps1"

REM Check if InstallTracker.ps1 exists
if not exist "%PS_SCRIPT%" (
    cls
    echo.
    echo ===================================================================
    echo ERROR: InstallTracker.ps1 not found!
    echo ===================================================================
    echo.
    echo Expected location: %PS_SCRIPT%
    echo.
    echo Please make sure both files are in the same directory:
    echo   - InstallTracker.bat
    echo   - InstallTracker.ps1
    echo.
    pause
    exit /b 1
)

REM Clear screen and show startup message
cls
echo.
echo Starting InstallTracker...
echo.

REM Start PowerShell hidden (WindowStyle Hidden) and run the script
powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
