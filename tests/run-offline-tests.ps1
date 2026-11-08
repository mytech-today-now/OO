[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CommandFixturePath
)

$ErrorActionPreference = 'Stop'

$script:FailureCount = 0
$script:SkippedCount = 0
$script:RunCount = 0
$script:FirstFailureStatus = 0
$script:LastProcessExitCode = 0

$script:PowerShellExecutableName = if ($PSVersionTable.PSEdition -eq 'Core') {
    if ($env:OS -eq 'Windows_NT') { 'pwsh.exe' } else { 'pwsh' }
}
else {
    'powershell.exe'
}
$script:PowerShellExecutable = Join-Path $PSHOME $script:PowerShellExecutableName
$script:CommandProcessor = $env:ComSpec
$script:BashCommand = Get-Command bash -ErrorAction SilentlyContinue
$script:IsWindowsHost = $env:OS -eq 'Windows_NT'
$script:IsLinuxHost = (-not $script:IsWindowsHost) -and (Test-Path -LiteralPath '/proc/sys/kernel/osrelease')

function Write-SkippedCheck {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    Write-Output ("SKIP: {0} ({1})" -f $Label, $Reason)
    $script:SkippedCount++
}

function Invoke-ChildProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $script:LastProcessExitCode = 0
    $resolvedPath = $null

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    }
    else {
        $resolvedCommand = Get-Command -Name $Path -ErrorAction SilentlyContinue
        if ($null -ne $resolvedCommand) {
            $resolvedPath = $resolvedCommand.Source
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
        $script:LastProcessExitCode = 127
        return
    }

    $extension = [System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant()

    try {
        switch ($extension) {
            '.ps1' {
                if (-not (Test-Path -LiteralPath $script:PowerShellExecutable -PathType Leaf)) {
                    $script:LastProcessExitCode = 127
                    return
                }

                & $script:PowerShellExecutable -NoLogo -NoProfile -NonInteractive -File $resolvedPath
                break
            }
            '.cmd' {
                if ([string]::IsNullOrWhiteSpace($script:CommandProcessor) -or
                    -not (Test-Path -LiteralPath $script:CommandProcessor -PathType Leaf)) {
                    $script:LastProcessExitCode = 127
                    return
                }

                $commandLine = '""{0}""' -f $resolvedPath
                & $script:CommandProcessor /d /s /c $commandLine
                break
            }
            '.sh' {
                if ($null -eq $script:BashCommand) {
                    $script:LastProcessExitCode = 127
                    return
                }

                & $script:BashCommand.Source $resolvedPath
                break
            }
            default {
                & $resolvedPath
                break
            }
        }

        if ($null -eq $LASTEXITCODE) {
            $script:LastProcessExitCode = 0
        }
        else {
            $script:LastProcessExitCode = [int]$LASTEXITCODE
        }
    }
    catch {
        Write-Output ("COMMAND ERROR: {0}" -f $_.Exception.Message)
        $script:LastProcessExitCode = 1
    }
}

function Invoke-OfflineCheck {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Write-Output ("RUN: {0}" -f $Label)
    $script:RunCount++
    Invoke-ChildProcess -Path $Path
    $status = $script:LastProcessExitCode

    if ($status -eq 0) {
        Write-Output ("PASS: {0}" -f $Label)
    }
    elseif ($status -eq 77) {
        Write-Output ("SKIP: {0} (exit 77)" -f $Label)
        $script:SkippedCount++
    }
    else {
        Write-Output ("FAIL: {0} (exit {1})" -f $Label, $status)
        $script:FailureCount++
        if ($script:FirstFailureStatus -eq 0) {
            $script:FirstFailureStatus = $status
        }
    }
}

if (-not [string]::IsNullOrWhiteSpace($CommandFixturePath)) {
    if (-not (Test-Path -LiteralPath $CommandFixturePath -PathType Leaf)) {
        Write-Output 'FAIL: -CommandFixturePath does not identify a file (exit 2)'
        exit 2
    }

    try {
        $commandManifest = Get-Content -LiteralPath $CommandFixturePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $commandFixtures = @($commandManifest.Commands)
    }
    catch {
        Write-Output 'FAIL: -CommandFixturePath must contain a JSON array of labeled commands (exit 2)'
        exit 2
    }

    if ($commandFixtures.Count -eq 0) {
        Write-Output 'FAIL: -CommandFixturePath must contain at least one command (exit 2)'
        exit 2
    }

    foreach ($commandFixture in $commandFixtures) {
        if ([string]::IsNullOrWhiteSpace([string]$commandFixture.Label) -or
            [string]::IsNullOrWhiteSpace([string]$commandFixture.Path)) {
            Write-Output 'FAIL: each command fixture requires Label and Path values (exit 2)'
            exit 2
        }
        Invoke-OfflineCheck -Label ([string]$commandFixture.Label) -Path ([string]$commandFixture.Path)
    }

    Write-Output ("SUMMARY: {0} run, {1} skipped, {2} failed" -f $script:RunCount, $script:SkippedCount, $script:FailureCount)
    if ($script:FailureCount -gt 0) {
        exit $script:FirstFailureStatus
    }
    exit 0
}

$runnerTest = Join-Path $PSScriptRoot 'test-offline-runner.ps1'
$windowsPowerShellTests = @(
    'test-execution-policy-docs.ps1',
    'test-install-winget-consent.ps1',
    'test-local-logging.ps1',
    'test-ooshutup-payload-trust.ps1',
    'test-restart-safety.ps1',
    'test-restore-point-readiness.ps1',
    'test-scheduled-task-documentation.ps1'
)
$windowsBatchTest = Join-Path $PSScriptRoot 'test-install-bat-elevation.cmd'
$linuxBashTests = @(
    'test-install-apt.sh',
    'test-install-dnf-pacman.sh'
)

Invoke-OfflineCheck -Label 'runner failure propagation' -Path $runnerTest

if ($script:IsWindowsHost) {
    foreach ($testName in $windowsPowerShellTests) {
        $testPath = Join-Path $PSScriptRoot $testName
        Invoke-OfflineCheck -Label $testName -Path $testPath
    }

    Invoke-OfflineCheck -Label 'test-install-bat-elevation.cmd' -Path $windowsBatchTest

    foreach ($testName in $linuxBashTests) {
        Write-SkippedCheck -Label $testName -Reason 'requires a Linux host'
    }
}
elseif ($script:IsLinuxHost) {
    foreach ($testName in $windowsPowerShellTests) {
        Write-SkippedCheck -Label $testName -Reason 'Windows-only safety harness'
    }
    Write-SkippedCheck -Label 'test-install-bat-elevation.cmd' -Reason 'requires Windows CMD and PowerShell 5.1'

    foreach ($testName in $linuxBashTests) {
        $testPath = Join-Path $PSScriptRoot $testName
        if ($null -eq $script:BashCommand) {
            Write-SkippedCheck -Label $testName -Reason 'Bash is unavailable'
        }
        else {
            Invoke-OfflineCheck -Label $testName -Path $testPath
        }
    }
}
else {
    foreach ($testName in $windowsPowerShellTests) {
        Write-SkippedCheck -Label $testName -Reason 'requires Windows'
    }
    Write-SkippedCheck -Label 'test-install-bat-elevation.cmd' -Reason 'requires Windows CMD and PowerShell 5.1'
    foreach ($testName in $linuxBashTests) {
        Write-SkippedCheck -Label $testName -Reason 'requires Linux'
    }
}

Write-Output ("SUMMARY: {0} run, {1} skipped, {2} failed" -f $script:RunCount, $script:SkippedCount, $script:FailureCount)

if ($script:FailureCount -gt 0) {
    exit $script:FirstFailureStatus
}

exit 0
