@echo off
REM WinToPalette Uninstaller
REM Double-click to uninstall WinToPalette with admin privileges

setlocal enabledelayedexpansion

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    REM Not admin - request elevation
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit /b
)

REM We have admin privileges, run the PowerShell script
title WinToPalette Uninstaller
echo.
echo ================================================================
echo                   WinToPalette Uninstaller
echo ================================================================
echo.

REM Get the directory where this batch file is located
set "scriptDir=%~dp0"

REM Look for the PowerShell script in the same directory
if exist "%scriptDir%Install-WinToPalette.ps1" (
    echo Starting uninstallation...
    echo.
    powershell -NoProfile -ExecutionPolicy Bypass -File "%scriptDir%Install-WinToPalette.ps1" -Uninstall
) else (
    echo ERROR: Install-WinToPalette.ps1 not found!
    echo The PowerShell script must be in the same folder as this batch file.
    echo.
    pause
    exit /b 1
)

if !errorlevel! equ 0 (
    echo.
    echo ================================================================
    echo Uninstallation completed successfully!
    echo Press any key to exit...
    echo ================================================================
    pause
) else (
    echo.
    echo ================================================================
    echo Uninstallation failed. Please check the error messages above.
    echo Press any key to exit...
    echo ================================================================
    pause
    exit /b 1
)
