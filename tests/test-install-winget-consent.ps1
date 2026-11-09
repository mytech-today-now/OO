[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:FallbackMessage = 'Package agreements were not accepted. PowerShell was not installed; review the WinGet terms before retrying.'
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceBatchPath = Join-Path $repoRoot 'install.bat'
$elevationHelperPath = Join-Path $repoRoot 'Install-OOShutUp10-Elevation.ps1'
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
        [string]$ElevatedExitCode = '0',
        [string]$InstallerExitCode = '0',
        [string]$InstallerArguments = '',
        [switch]$PowerShellPresent,
        [switch]$WingetPresent = $true,
        [switch]$ElevatedMarker,
        [switch]$UacCancelled
    )

    Reset-Logs
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Join-Path $system32 'cmd.exe'
    $startInfo.Arguments = '/d /c call "' + $batchPath + '"'
    if ($InstallerArguments) {
        $startInfo.Arguments += ' ' + $InstallerArguments
    }
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
    $startInfo.EnvironmentVariables['FAKE_INSTALLER_EXIT_CODE'] = $InstallerExitCode
    if ($PowerShellPresent) {
        $startInfo.EnvironmentVariables['FAKE_PWSH_PRESENT'] = '1'
    } else {
        [void]$startInfo.EnvironmentVariables.Remove('FAKE_PWSH_PRESENT')
    }
    if ($WingetPresent) {
        $startInfo.EnvironmentVariables['FAKE_WINGET_PRESENT'] = '1'
    } else {
        [void]$startInfo.EnvironmentVariables.Remove('FAKE_WINGET_PRESENT')
    }
    if ($ElevatedMarker) {
        $startInfo.EnvironmentVariables['MYTECH_INSTALL_ELEVATION_ATTEMPTED'] = '1'
    } else {
        [void]$startInfo.EnvironmentVariables.Remove('MYTECH_INSTALL_ELEVATION_ATTEMPTED')
    }
    if ($UacCancelled) {
        $startInfo.EnvironmentVariables['FAKE_UAC_CANCELLED'] = '1'
    } else {
        [void]$startInfo.EnvironmentVariables.Remove('FAKE_UAC_CANCELLED')
    }
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
    $testBatch = $testBatch.Replace('powershell -NoProfile -ExecutionPolicy AllSigned -File', 'call powershell -NoProfile -ExecutionPolicy AllSigned -File')
    $testBatch = $testBatch.Replace('powershell -NoProfile -NonInteractive -Command', 'call powershell -NoProfile -NonInteractive -Command')
    $testBatch = $testBatch.Replace('net session >nul 2>&1', 'call net session >nul 2>&1')
    $testBatch = [System.Text.RegularExpressions.Regex]::Replace($testBatch, '(?m)^winget install ', 'call winget install ')
    $testBatch = $testBatch.Replace('pwsh -NoProfile -ExecutionPolicy Bypass -File', 'call pwsh -NoProfile -ExecutionPolicy Bypass -File')
    [System.IO.File]::WriteAllText($batchPath, $testBatch, [System.Text.Encoding]::ASCII)
    Copy-Item -LiteralPath $elevationHelperPath -Destination (Join-Path $testRoot 'Install-OOShutUp10-Elevation.ps1')

    $fakeCommands = @{
        'where.cmd' = @'
@echo off
>>"%CHECK_LOG%" echo where=%*
if /I "%~1"=="pwsh" if defined FAKE_PWSH_PRESENT exit /b 0
if /I "%~1"=="pwsh" exit /b 1
if /I "%~1"=="winget" if defined FAKE_WINGET_PRESENT exit /b 0
if /I "%~1"=="winget" exit /b 1
exit /b 1
'@
        'powershell.cmd' = @'
@echo off
if /I "%~4"=="-File" (
    >>"%CHECK_LOG%" echo elevated-command=%*
    if defined FAKE_UAC_CANCELLED >"%~7" echo UAC_CANCELLED
    exit /b %FAKE_ELEVATED_EXIT_CODE%
)
>>"%CHECK_LOG%" echo console-check
if defined FAKE_INPUT_REDIRECTED exit /b 1
exit /b 0
'@
        'net.cmd' = @'
@echo off
>>"%CHECK_LOG%" echo net-session
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
exit /b %FAKE_INSTALLER_EXIT_CODE%
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

    # Elevation is decided before runtime discovery, and the mocked handoff never requests real UAC.
    $elevationSource = Get-Content -LiteralPath $elevationHelperPath -Raw
    Assert-True ($sourceBatch.IndexOf('net session', [System.StringComparison]::Ordinal) -lt $sourceBatch.IndexOf('where pwsh', [System.StringComparison]::Ordinal)) 'the administrator check precedes both runtime branches'
    Assert-True ($sourceBatch.Contains('Install-OOShutUp10-Elevation.ps1') -and $elevationSource.Contains('-Verb RunAs') -and $elevationSource.Contains('-Wait') -and $elevationSource.Contains('-PassThru')) 'the elevation helper waits for and captures the elevated child'
    Assert-True ($sourceBatch.Contains('MYTECH_INSTALL_ELEVATION_ATTEMPTED')) 'the batch wrapper blocks a recursive elevation attempt'

    $forwardedArguments = '-Profile "alpha beta" "O''Brien & Co"'
    foreach ($childExitCode in @('0', '1', '2')) {
        $elevated = Invoke-InstallerMock -WingetExitCode '0' -NetExitCode '1' -ElevatedExitCode $childExitCode -PowerShellPresent -InstallerArguments $forwardedArguments
        Assert-Equal ([int]$childExitCode) $elevated.ExitCode "the non-admin wrapper preserves elevated child exit code $childExitCode"
        Assert-Equal 1 @($elevated.ConsoleChecks | Where-Object { $_ -like 'elevated-command=*' }).Count 'the non-admin path requests elevation once'
        Assert-Equal 0 $elevated.WingetArgs.Count 'the non-admin path never invokes WinGet before elevation'
        Assert-Equal 0 $elevated.PwshArgs.Count 'the non-admin path never starts O&O before elevation'
        $elevationLogLine = @($elevated.ConsoleChecks | Where-Object { $_ -like 'elevated-command=*' })[0]
        Assert-True ($elevationLogLine.Contains($forwardedArguments)) 'the elevation request forwards installer arguments exactly'
        Assert-True ($elevated.ConsoleChecks.IndexOf('net-session') -lt $elevated.ConsoleChecks.IndexOf($elevationLogLine)) 'the administrator check happens before the elevation request'
    }

    $cancelled = Invoke-InstallerMock -WingetExitCode '0' -NetExitCode '1' -ElevatedExitCode '1223' -PowerShellPresent -UacCancelled
    Assert-Equal 1223 $cancelled.ExitCode 'UAC cancellation returns its nonzero status'
    Assert-True ($cancelled.StdOut.Contains('UAC elevation was cancelled')) 'UAC cancellation is reported clearly'
    Assert-Equal 0 $cancelled.WingetArgs.Count 'UAC cancellation never invokes WinGet'
    Assert-Equal 0 $cancelled.PwshArgs.Count 'UAC cancellation never starts O&O'

    $childExit1223 = Invoke-InstallerMock -WingetExitCode '0' -NetExitCode '1' -ElevatedExitCode '1223' -PowerShellPresent
    Assert-Equal 1223 $childExit1223.ExitCode 'a child exit code that matches the Win32 cancellation code is preserved'
    Assert-True (-not $childExit1223.StdOut.Contains('UAC elevation was cancelled')) 'a child exit code is not mistaken for UAC cancellation'
    Assert-Equal 0 $childExit1223.PwshArgs.Count 'the mocked elevated child status does not start the installer before elevation'

    $missingRuntime = Invoke-InstallerMock -WingetExitCode '0' -NetExitCode '1' -ElevatedExitCode '0' -WingetPresent
    Assert-Equal 0 $missingRuntime.WingetArgs.Count 'a non-admin runtime-missing path does not invoke WinGet before elevation'
    Assert-Equal 0 $missingRuntime.PwshArgs.Count 'a non-admin runtime-missing path does not start the installer before elevation'

    $adminRuntimePresent = Invoke-InstallerMock -WingetExitCode '0' -InstallerExitCode '2' -PowerShellPresent
    Assert-Equal 2 $adminRuntimePresent.ExitCode 'an administrator PowerShell-present path propagates the child status'
    Assert-Equal 0 $adminRuntimePresent.WingetArgs.Count 'an already-installed PowerShell runtime never invokes WinGet'
    Assert-Equal 0 @($adminRuntimePresent.ConsoleChecks | Where-Object { $_ -like 'elevated-command=*' }).Count 'an administrator path skips the elevation request'

    $retryGuard = Invoke-InstallerMock -WingetExitCode '0' -NetExitCode '1' -ElevatedExitCode '0' -PowerShellPresent -ElevatedMarker
    Assert-Equal 1 $retryGuard.ExitCode 'a failed elevation retry guard exits nonzero'
    Assert-Equal 0 @($retryGuard.ConsoleChecks | Where-Object { $_ -like 'elevated-command=*' }).Count 'the retry guard prevents recursive elevation'
    Assert-Equal 0 $retryGuard.PwshArgs.Count 'the retry guard never starts the installer'

    # Exercise the helper's argument serialization and return handling with Start-Process mocked.
    $helperBatchFixture = Join-Path $testRoot 'argument forwarding fixture.bat'
    $helperCapturePath = Join-Path $testRoot 'elevated-arguments.txt'
    $helperCancelMarkerPath = Join-Path $testRoot 'helper-uac-cancelled.marker'
    $helperBatchFixtureContent = @'
@echo off
>"%ELEVATION_ARGUMENT_CAPTURE%" echo %*
exit /b %ELEVATION_CHILD_EXIT_CODE%
'@
    [System.IO.File]::WriteAllText($helperBatchFixture, $helperBatchFixtureContent + "`r`n", [System.Text.Encoding]::ASCII)
    $helperTokens = $null
    $helperParseErrors = $null
    $helperAst = [System.Management.Automation.Language.Parser]::ParseFile($elevationHelperPath, [ref]$helperTokens, [ref]$helperParseErrors)
    Assert-Equal 0 @($helperParseErrors).Count 'the elevation helper parses in Windows PowerShell'
    foreach ($functionAst in $helperAst.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }

    $script:MockElevationExitCode = 0
    $script:MockElevationShouldCancel = $false
    $script:MockElevationUseBatchFixture = $true
    $script:MockElevationArgumentList = $null
    $script:MockElevationMarker = $null
    function Start-Process {
        param(
            [string]$FilePath,
            [string]$ArgumentList,
            [string]$Verb,
            [string]$WorkingDirectory,
            [switch]$Wait,
            [switch]$PassThru,
            [string]$ErrorAction
        )

        $script:MockElevationArgumentList = $ArgumentList
        $script:MockElevationMarker = [Environment]::GetEnvironmentVariable('MYTECH_INSTALL_ELEVATION_ATTEMPTED', 'Process')
        if ($script:MockElevationShouldCancel) {
            throw (New-Object System.ComponentModel.Win32Exception 1223)
        }
        if ($script:MockElevationUseBatchFixture) {
            return Microsoft.PowerShell.Management\Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -WorkingDirectory $WorkingDirectory -Wait -PassThru -ErrorAction Stop
        }
        return [pscustomobject]@{ ExitCode = $script:MockElevationExitCode }
    }

    $previousElevationMarker = [Environment]::GetEnvironmentVariable('MYTECH_INSTALL_ELEVATION_ATTEMPTED', 'Process')
    [Environment]::SetEnvironmentVariable('MYTECH_INSTALL_ELEVATION_ATTEMPTED', $null, 'Process')
    $clearedElevationMarker = [Environment]::GetEnvironmentVariable('MYTECH_INSTALL_ELEVATION_ATTEMPTED', 'Process')
    try {
        $serializedArguments = @('alpha beta', "O'Brien & Co", '', 'C:\path\', 'quote"value')
        foreach ($childExitCode in @(0, 1, 2)) {
            $script:MockElevationExitCode = $childExitCode
            $script:MockElevationShouldCancel = $false
            $script:MockElevationUseBatchFixture = $true
            $env:ELEVATION_ARGUMENT_CAPTURE = $helperCapturePath
            $env:ELEVATION_CHILD_EXIT_CODE = [string]$childExitCode
            $helperExitCode = Invoke-OOShutUp10Elevation -BatchPath $helperBatchFixture -InstallerArguments $serializedArguments
            Assert-Equal $childExitCode $helperExitCode 'the helper returns the elevated batch exit code unchanged'
            Assert-Equal '"alpha beta" "O''Brien & Co" "" "C:\path\\" "quote\"value"' $script:MockElevationArgumentList 'the helper preserves spaces, apostrophes, ampersands, empty values, and quotes'
            Assert-Equal '"alpha beta" "O''Brien & Co" "" "C:\path\\" "quote\"value"' (Get-Content -LiteralPath $helperCapturePath -Raw).TrimEnd("`r", "`n") 'the elevated CMD process receives the original argument boundaries'
            Assert-Equal '1' $script:MockElevationMarker 'the helper marks the elevated child to prevent recursion'
            Assert-Equal $clearedElevationMarker ([Environment]::GetEnvironmentVariable('MYTECH_INSTALL_ELEVATION_ATTEMPTED', 'Process')) 'the parent environment is restored after the handoff'
        }

        $script:MockElevationShouldCancel = $true
        $script:MockElevationUseBatchFixture = $false
        $cancelCode = Invoke-OOShutUp10Elevation -BatchPath $helperBatchFixture -CancellationMarkerPath $helperCancelMarkerPath -InstallerArguments $serializedArguments
        Assert-Equal 1223 $cancelCode 'the helper recognizes UAC cancellation'
        Assert-Equal 'UAC_CANCELLED' (Get-Content -LiteralPath $helperCancelMarkerPath -Raw) 'the helper marks UAC cancellation separately from a child exit code'
    }
    finally {
        Remove-Item Env:\ELEVATION_ARGUMENT_CAPTURE -ErrorAction SilentlyContinue
        Remove-Item Env:\ELEVATION_CHILD_EXIT_CODE -ErrorAction SilentlyContinue
        Remove-Item Function:\Start-Process -ErrorAction SilentlyContinue
        [Environment]::SetEnvironmentVariable('MYTECH_INSTALL_ELEVATION_ATTEMPTED', $previousElevationMarker, 'Process')
    }

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
