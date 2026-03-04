@echo off
REM WinToPalette Installer
REM Double-click to install WinToPalette with admin privileges

setlocal enabledelayedexpansion

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    REM Not admin - request elevation
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit /b
)

REM We have admin privileges, run the PowerShell script
title WinToPalette Installer
echo.
echo ================================================================
echo                   WinToPalette Installer
echo ================================================================
echo.

REM Get the directory where this batch file is located
set "scriptDir=%~dp0"

REM Look for the PowerShell script in the same directory
if exist "%scriptDir%Install-WinToPalette.ps1" (
    echo Starting installation...
    echo.
    powershell -NoProfile -ExecutionPolicy Bypass -File "%scriptDir%Install-WinToPalette.ps1"
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
    echo Installation completed successfully!
    echo Press any key to exit...
    echo ================================================================
    pause
) else (
    echo.
    echo ================================================================
    echo Installation failed. Please check the error messages above.
    echo Press any key to exit...
    echo ================================================================
    pause
    exit /b 1
)
