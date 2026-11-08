@echo off
setlocal DisableDelayedExpansion
REM The PowerShell test creates fake where, net, powershell, pwsh, and winget commands and runs the batch under CMD.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0test-install-winget-consent.ps1"
set "TEST_EXIT_CODE=%ERRORLEVEL%"
exit /b %TEST_EXIT_CODE%
