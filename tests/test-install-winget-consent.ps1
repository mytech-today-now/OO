[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:FallbackMessage = 'Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.'
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceBatchPath = Join-Path $repoRoot 'install.bat'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oosu10-winget-consent-' + [guid]::NewGuid().ToString('N'))
$batchPath = Join-Path $testRoot 'install-under-test.bat'
$system32 = Join-Path $env:SystemRoot 'System32'
$wingetLog = Join-Path $testRoot 'winget-args.log'
$pwshLog = Join-Path $testRoot 'pwsh-args.log'
$checkLog = Join-Path $testRoot 'console-check.log'
$mutationLog = Join-Path $testRoot 'package-mutation.log'

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }

    $script:AssertionCount++
}

function Assert-Equal {
    param(
        [object]$Expected,
        [object]$Actual,
        [string]$Message
    )

    if ($Expected -cne $Actual) {
        throw "Assertion failed: $Message (expected '$Expected', got '$Actual')"
    }

    $script:AssertionCount++
}

function Reset-Logs {
    foreach ($path in @($wingetLog, $pwshLog, $checkLog, $mutationLog)) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }
}

function Get-LogLines {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    return @(Get-Content -LiteralPath $Path)
}

function Invoke-InstallerMock {
    param(
        [string]$WingetExitCode,
        [switch]$RedirectedInput,
        [string]$ConsentValue,
        [string]$NetExitCode = '0',
        [string]$ElevatedExitCode = '0'
    )

    Reset-Logs
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Join-Path $system32 'cmd.exe'
    $startInfo.Arguments = '/d /c call "' + $batchPath + '"'
    $startInfo.WorkingDirectory = $repoRoot
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables['PATH'] = $testRoot + ';' + $system32
    $startInfo.EnvironmentVariables['PATHEXT'] = '.COM;.EXE;.BAT;.CMD'
    $startInfo.EnvironmentVariables['SYSTEMROOT'] = $env:SystemRoot
    $startInfo.EnvironmentVariables['WINDIR'] = $env:SystemRoot
    $startInfo.EnvironmentVariables['WINGET_LOG'] = $wingetLog
    $startInfo.EnvironmentVariables['PWSH_LOG'] = $pwshLog
    $startInfo.EnvironmentVariables['CHECK_LOG'] = $checkLog
    $startInfo.EnvironmentVariables['MUTATION_LOG'] = $mutationLog
    $startInfo.EnvironmentVariables['FAKE_WINGET_EXIT_CODE'] = $WingetExitCode
    $startInfo.EnvironmentVariables['FAKE_NET_EXIT_CODE'] = $NetExitCode
    $startInfo.EnvironmentVariables['FAKE_ELEVATED_EXIT_CODE'] = $ElevatedExitCode
    if ($RedirectedInput) {
        $startInfo.EnvironmentVariables['FAKE_INPUT_REDIRECTED'] = '1'
    } else {
        [void]$startInfo.EnvironmentVariables.Remove('FAKE_INPUT_REDIRECTED')
    }
    if ($ConsentValue) {
        $startInfo.EnvironmentVariables['MYTECH_ACCEPT_WINGET_AGREEMENTS'] = $ConsentValue
    } else {
        [void]$startInfo.EnvironmentVariables.Remove('MYTECH_ACCEPT_WINGET_AGREEMENTS')
    }

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) {
        throw 'The mock installer process did not start.'
    }

    $process.StandardInput.Write('X')
    $process.StandardInput.Close()
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(20000)) {
        $process.Kill()
        throw 'The mock installer did not exit within 20 seconds.'
    }
    $stdout = $stdoutTask.Result
    $stderr = $stderrTask.Result

    $result = [pscustomobject]@{
        ExitCode = $process.ExitCode
        StdOut = $stdout
        StdErr = $stderr
        WingetArgs = @(Get-LogLines $wingetLog)
        PwshArgs = @(Get-LogLines $pwshLog)
        ConsoleChecks = @(Get-LogLines $checkLog)
        Mutations = @(Get-LogLines $mutationLog)
    }
    $process.Dispose()
    return $result
}

try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

    # Add CALL only around external commands so local .cmd shims return normally; the consent and exit-code logic stays sourced from install.bat.
    $sourceBatch = Get-Content -LiteralPath $sourceBatchPath -Raw
    $testBatch = [System.Text.RegularExpressions.Regex]::Replace($sourceBatch, '(?m)^where (pwsh|winget) ', 'call where $1 ')
    $testBatch = $testBatch.Replace('powershell.exe -NoProfile -NonInteractive -Command', 'call powershell -NoProfile -NonInteractive -Command')
    $testBatch = $testBatch.Replace('net session >nul 2>&1', 'call net session >nul 2>&1')
    $testBatch = [System.Text.RegularExpressions.Regex]::Replace($testBatch, '(?m)^    winget install ', '    call winget install ')
    $testBatch = $testBatch.Replace('powershell -Command ', 'call powershell -Command ')
    $testBatch = $testBatch.Replace('pwsh -NoProfile -ExecutionPolicy Bypass -File', 'call pwsh -NoProfile -ExecutionPolicy Bypass -File')
    [System.IO.File]::WriteAllText($batchPath, $testBatch, [System.Text.Encoding]::ASCII)

    $fakeCommands = @{
        'where.cmd' = @'
@echo off
>>"%CHECK_LOG%" echo where=%*
if /I "%~1"=="pwsh" exit /b 1
if /I "%~1"=="winget" exit /b 0
exit /b 1
'@
        'powershell.cmd' = @'
@echo off
if /I "%~1"=="-Command" (
    >>"%CHECK_LOG%" echo elevated-command
    exit /b %FAKE_ELEVATED_EXIT_CODE%
)
>>"%CHECK_LOG%" echo console-check
if defined FAKE_INPUT_REDIRECTED exit /b 1
exit /b 0
'@
        'net.cmd' = @'
@echo off
exit /b %FAKE_NET_EXIT_CODE%
'@
        'winget.cmd' = @'
@echo off
>>"%WINGET_LOG%" echo %*
if "%FAKE_WINGET_EXIT_CODE%"=="0" >>"%MUTATION_LOG%" echo installed
exit /b %FAKE_WINGET_EXIT_CODE%
'@
        'pwsh.cmd' = @'
@echo off
>>"%PWSH_LOG%" echo %*
exit /b 0
'@
    }

    foreach ($entry in $fakeCommands.GetEnumerator()) {
        $path = Join-Path $testRoot $entry.Key
        [System.IO.File]::WriteAllText($path, $entry.Value.TrimStart("`r", "`n") + "`r`n", [System.Text.Encoding]::ASCII)
    }

    # The fake console is interactive here; fake WinGet success represents the user's acceptance of its normal review.
    $default = Invoke-InstallerMock -WingetExitCode '0'
    if ($default.ExitCode -ne 0) {
        Write-Output ('Mock stdout: ' + $default.StdOut)
        Write-Output ('Mock stderr: ' + $default.StdErr)
        Write-Output ('Mock WinGet args: ' + ($default.WingetArgs -join ' | '))
        Write-Output ('Mock PowerShell args: ' + ($default.PwshArgs -join ' | '))
        Write-Output ('Mock console checks: ' + ($default.ConsoleChecks -join ' | '))
    }
    Assert-Equal 0 $default.ExitCode 'default interactive install exits successfully with the mock'
    Assert-Equal 'install --id Microsoft.PowerShell --source winget' $default.WingetArgs[0] 'default WinGet arguments contain no agreement acceptance'
    Assert-Equal 1 $default.WingetArgs.Count 'the default path invokes WinGet once'
    Assert-Equal 1 $default.PwshArgs.Count 'the default successful path proceeds to the PowerShell installer'
    Assert-Equal 1 $default.Mutations.Count 'the interactive WinGet mock can install after its normal review'

    # Unattended acceptance is a separate exact named input, and only then are both flags supplied.
    $accepted = Invoke-InstallerMock -WingetExitCode '0' -RedirectedInput -ConsentValue 'YES'
    Assert-Equal 0 $accepted.ExitCode 'explicit unattended consent exits successfully with the mock'
    Assert-Equal 'install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements' $accepted.WingetArgs[0] 'explicit consent supplies both WinGet agreement flags'
    Assert-Equal 1 $accepted.PwshArgs.Count 'the accepted successful path proceeds to the PowerShell installer'
    Assert-Equal 1 $accepted.Mutations.Count 'explicitly accepted unattended execution can install'

    # Missing and near-miss unattended consent both fail before WinGet.
    $missing = Invoke-InstallerMock -WingetExitCode '0' -RedirectedInput
    Assert-Equal 1 $missing.ExitCode 'unattended execution without consent fails closed'
    Assert-True ($missing.StdOut.Contains($script:FallbackMessage)) 'the missing-consent path prints the exact fallback'
    Assert-Equal 0 $missing.WingetArgs.Count 'WinGet is not invoked without unattended consent'
    Assert-Equal 0 $missing.PwshArgs.Count 'the O&O installer is not invoked without unattended consent'
    $nearMiss = Invoke-InstallerMock -WingetExitCode '0' -RedirectedInput -ConsentValue 'YES PLEASE'
    Assert-Equal 1 $nearMiss.ExitCode 'a near-miss consent value fails closed'
    Assert-True ($nearMiss.StdOut.Contains($script:FallbackMessage)) 'the near-miss consent path prints the exact fallback'
    Assert-Equal 0 $nearMiss.WingetArgs.Count 'a near-miss consent value does not invoke WinGet'

    $failed = Invoke-InstallerMock -WingetExitCode '37'
    Assert-Equal 37 $failed.ExitCode 'a non-agreement WinGet failure preserves its native status'
    Assert-True ($failed.StdOut.Contains('WinGet failed with exit code 37')) 'a non-agreement failure reports its native status'
    Assert-Equal 1 $failed.WingetArgs.Count 'a non-agreement WinGet failure is not retried'
    Assert-Equal 0 $failed.PwshArgs.Count 'a non-agreement WinGet failure does not start O&O'
    Assert-Equal 0 $failed.Mutations.Count 'a failed WinGet mock records no package mutation'

    # A non-admin launcher returns the elevated copy's rejection status without performing real elevation.
    Assert-True ($sourceBatch.Contains('-Wait -PassThru')) 'the elevation relaunch waits for and captures its child process'
    $elevatedReject = Invoke-InstallerMock -WingetExitCode '0' -NetExitCode '1' -ElevatedExitCode '-1978335167'
    Assert-Equal -1978335167 $elevatedReject.ExitCode 'the non-admin wrapper preserves the elevated rejection status'
    Assert-True ($elevatedReject.ConsoleChecks -contains 'elevated-command') 'the elevation branch used the fake PowerShell shim'
    Assert-Equal 0 $elevatedReject.WingetArgs.Count 'the elevation mock does not invoke real or fake WinGet'
    Assert-Equal 0 $elevatedReject.PwshArgs.Count 'the elevation mock does not start O&O'

    # Agreement rejection preserves WinGet's native HRESULT, performs no retry, and never starts O&O.
    foreach ($rejectionCode in @('-1978335167', '-1978335162')) {
        $rejected = Invoke-InstallerMock -WingetExitCode $rejectionCode
        Assert-Equal ([int]$rejectionCode) $rejected.ExitCode 'agreement rejection preserves WinGet native status'
        Assert-True ($rejected.StdOut.Contains($script:FallbackMessage)) 'agreement rejection prints the exact fallback'
        Assert-Equal 1 $rejected.WingetArgs.Count 'agreement rejection is not retried'
        Assert-Equal 'install --id Microsoft.PowerShell --source winget' $rejected.WingetArgs[0] 'rejection path has no auto-acceptance flags'
        Assert-Equal 0 $rejected.PwshArgs.Count 'the O&O installer is not invoked after agreement rejection'
        Assert-Equal 0 $rejected.Mutations.Count 'the rejected mock path records no package mutation'
    }

    Write-Output ("PASS: WinGet agreement consent mock tests ({0} assertions)." -f $script:AssertionCount)
}
finally {
    $tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    $resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
    if ($resolvedTestRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedTestRoot) -like 'oosu10-winget-consent-*' -and
        (Test-Path -LiteralPath $resolvedTestRoot)) {
        Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
    }
}
