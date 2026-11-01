@echo off
REM ============================================================================
REM  myTech.Today - O&O ShutUp10 Installer Setup
REM  Checks for PowerShell 7+, installs if needed, then runs the script
REM ============================================================================
setlocal enabledelayedexpansion

echo.
echo ============================================================
echo   myTech.Today - O^&O ShutUp10 Installer Setup
echo ============================================================
echo.

where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] PowerShell 7+ is installed
    goto :RUN_INSTALLER
)

echo [INFO] PowerShell 7+ not found. Checking for winget...

where winget >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] winget is not available on this system.
    echo.
    echo Please install PowerShell 7 manually:
    echo   https://github.com/PowerShell/PowerShell/releases
    echo.
    pause
    exit /b 1
)

set "ACCEPT_WINGET_AGREEMENTS="
if /I "!MYTECH_ACCEPT_WINGET_AGREEMENTS!"=="YES" set "ACCEPT_WINGET_AGREEMENTS=1"

if not defined ACCEPT_WINGET_AGREEMENTS (
    powershell.exe -NoProfile -NonInteractive -Command "if ([Console]::IsInputRedirected) { exit 1 }"
    if errorlevel 1 (
        echo Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.
        exit /b 1
    )
)

echo [INFO] Installing PowerShell 7 using winget...

net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [WARN] Not running as administrator. Requesting elevation...
    powershell -Command "$installer = Start-Process '%~f0' -Verb RunAs -Wait -PassThru; exit $installer.ExitCode"
    set "ELEVATED_EXIT_CODE=!ERRORLEVEL!"
    exit /b !ELEVATED_EXIT_CODE!
)

if defined ACCEPT_WINGET_AGREEMENTS (
    winget install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements
) else (
    winget install --id Microsoft.PowerShell --source winget
)
set "WINGET_EXIT_CODE=!ERRORLEVEL!"

REM WinGet agreement rejection HRESULTs: package 0x8A150041, source 0x8A150046 (signed decimal).
if "!WINGET_EXIT_CODE!"=="-1978335167" (
    echo Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.
    exit /b !WINGET_EXIT_CODE!
)
if "!WINGET_EXIT_CODE!"=="-1978335162" (
    echo Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.
    exit /b !WINGET_EXIT_CODE!
)
if not "!WINGET_EXIT_CODE!"=="0" (
    echo [ERROR] WinGet failed with exit code !WINGET_EXIT_CODE!. PowerShell was not installed; O^&O was not started and no retry was attempted.
    exit /b !WINGET_EXIT_CODE!
)

echo [OK] PowerShell 7 installed successfully!

:RUN_INSTALLER
echo.
echo [INFO] Running O^&O ShutUp10 Installer...
echo.

pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-OOShutUp10.ps1" %*
set "INSTALLER_EXIT_CODE=!ERRORLEVEL!"

echo.
if "!INSTALLER_EXIT_CODE!"=="0" (
    echo [OK] Settings and the required update task were verified.
) else if "!INSTALLER_EXIT_CODE!"=="2" (
    echo [WARN] Setup is partial. Repair the update task before relying on automatic reapplication.
) else (
    echo [ERROR] Installer failed with exit code !INSTALLER_EXIT_CODE!.
)

pause
exit /b !INSTALLER_EXIT_CODE!

