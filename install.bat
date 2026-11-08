@echo off
REM ============================================================================
REM  myTech.Today - O&O ShutUp10 Installer Setup
REM  Checks for PowerShell 7+, installs if needed, then runs the script
REM ============================================================================
setlocal DisableDelayedExpansion

echo.
echo ============================================================
echo   myTech.Today - O^&O ShutUp10 Installer Setup
echo ============================================================
echo.

REM The installer requires elevation even when PowerShell 7 is already present.
net session >nul 2>&1
set "ADMIN_CHECK_EXIT_CODE=%ERRORLEVEL%"
if not "%ADMIN_CHECK_EXIT_CODE%"=="0" goto :REQUEST_ELEVATION
goto :CHECK_RUNTIME

:REQUEST_ELEVATION
if defined MYTECH_INSTALL_ELEVATION_ATTEMPTED goto :ELEVATION_RETRY_FAILED

set "ELEVATION_CANCEL_MARKER=%TEMP%\oosu10-uac-cancel-%RANDOM%-%RANDOM%.tmp"
if exist "%ELEVATION_CANCEL_MARKER%" goto :REQUEST_ELEVATION

echo [WARN] Not running as administrator. Requesting elevation...
powershell -NoProfile -ExecutionPolicy AllSigned -File "%~dp0Install-OOShutUp10-Elevation.ps1" "%~f0" "%ELEVATION_CANCEL_MARKER%" %*
set "ELEVATION_EXIT_CODE=%ERRORLEVEL%"
if exist "%ELEVATION_CANCEL_MARKER%" goto :ELEVATION_CANCELLED
exit /b %ELEVATION_EXIT_CODE%

:ELEVATION_CANCELLED
del /q "%ELEVATION_CANCEL_MARKER%" >nul 2>&1
echo [ERROR] UAC elevation was cancelled. No installer was started.
exit /b 1223

:ELEVATION_RETRY_FAILED
echo [ERROR] Administrator access is still unavailable after the elevation attempt. Refusing to retry.
exit /b 1

:CHECK_RUNTIME
where pwsh >nul 2>&1
set "PWSH_LOOKUP_EXIT_CODE=%ERRORLEVEL%"
if not "%PWSH_LOOKUP_EXIT_CODE%"=="0" goto :CHECK_WINGET
echo [OK] PowerShell 7+ is installed
goto :RUN_INSTALLER

:CHECK_WINGET
echo [INFO] PowerShell 7+ not found. Checking for winget...

where winget >nul 2>&1
set "WINGET_LOOKUP_EXIT_CODE=%ERRORLEVEL%"
if not "%WINGET_LOOKUP_EXIT_CODE%"=="0" goto :WINGET_UNAVAILABLE

set "ACCEPT_WINGET_AGREEMENTS="
if /I "%MYTECH_ACCEPT_WINGET_AGREEMENTS%"=="YES" set "ACCEPT_WINGET_AGREEMENTS=1"
if defined ACCEPT_WINGET_AGREEMENTS goto :RUN_WINGET_WITH_CONSENT

powershell -NoProfile -NonInteractive -Command "if ([Console]::IsInputRedirected) { exit 1 }"
set "AGREEMENT_CHECK_EXIT_CODE=%ERRORLEVEL%"
if not "%AGREEMENT_CHECK_EXIT_CODE%"=="0" goto :AGREEMENT_NOT_ACCEPTED

echo [INFO] Installing PowerShell 7 using winget...
winget install --id Microsoft.PowerShell --source winget
set "WINGET_EXIT_CODE=%ERRORLEVEL%"
goto :CHECK_WINGET_RESULT

:RUN_WINGET_WITH_CONSENT
echo [INFO] Installing PowerShell 7 using winget...
winget install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements
set "WINGET_EXIT_CODE=%ERRORLEVEL%"

:CHECK_WINGET_RESULT
REM WinGet agreement rejection HRESULTs: package 0x8A150041, source 0x8A150046 (signed decimal).
if "%WINGET_EXIT_CODE%"=="-1978335167" goto :WINGET_AGREEMENT_REJECTED
if "%WINGET_EXIT_CODE%"=="-1978335162" goto :WINGET_AGREEMENT_REJECTED
if not "%WINGET_EXIT_CODE%"=="0" goto :WINGET_FAILED

echo [OK] PowerShell 7 installed successfully!
goto :RUN_INSTALLER

:WINGET_UNAVAILABLE
echo [ERROR] winget is not available on this system.
echo.
echo Please install PowerShell 7 manually:
echo   https://github.com/PowerShell/PowerShell/releases
echo.
pause
exit /b 1

:AGREEMENT_NOT_ACCEPTED
echo Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.
exit /b 1

:WINGET_AGREEMENT_REJECTED
echo Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.
exit /b %WINGET_EXIT_CODE%

:WINGET_FAILED
echo [ERROR] WinGet failed with exit code %WINGET_EXIT_CODE%. PowerShell was not installed; O^&O was not started and no retry was attempted.
exit /b %WINGET_EXIT_CODE%

:RUN_INSTALLER
echo.
echo [INFO] Running O^&O ShutUp10 Installer...
echo.

pwsh -NoProfile -ExecutionPolicy AllSigned -File "%~dp0Install-OOShutUp10.ps1" %*
set "INSTALLER_EXIT_CODE=%ERRORLEVEL%"

if "%INSTALLER_EXIT_CODE%"=="0" goto :INSTALLER_SUCCESS
if "%INSTALLER_EXIT_CODE%"=="2" goto :INSTALLER_PARTIAL
goto :INSTALLER_FAILURE

:INSTALLER_SUCCESS
echo.
echo [OK] Settings and the required update task were verified.
goto :INSTALLER_FINISHED

:INSTALLER_PARTIAL
echo.
echo [WARN] Setup is partial. Repair the update task before relying on automatic reapplication.
goto :INSTALLER_FINISHED

:INSTALLER_FAILURE
echo.
echo [ERROR] Installer failed with exit code %INSTALLER_EXIT_CODE%.
echo Production launch requires an authorized signed release package and organization-enforced AllSigned policy.
echo For a Git source checkout only, see README.md for the explicit -DevelopmentMode command.

:INSTALLER_FINISHED
pause
exit /b %INSTALLER_EXIT_CODE%
