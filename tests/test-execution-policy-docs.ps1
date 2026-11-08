[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)

    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }

    $script:AssertionCount++
}

function Assert-Equal {
    param([object]$Expected, [object]$Actual, [string]$Message)

    if ($Expected -cne $Actual) {
        throw "Assertion failed: $Message (expected '$Expected', got '$Actual')"
    }

    $script:AssertionCount++
}

function Get-EngineExecutionPolicySnapshot {
    param([string]$EnginePath)

    $snapshotCommand = '$policies = @(Get-ExecutionPolicy -List | Sort-Object Scope | ForEach-Object { [pscustomobject]@{ Scope = [string]$_.Scope; ExecutionPolicy = [string]$_.ExecutionPolicy } }); $snapshot = [pscustomobject]@{ Policies = $policies }; [Console]::Out.WriteLine((ConvertTo-Json -InputObject $snapshot -Compress))'
    $output = @(& $EnginePath -NoLogo -NoProfile -NonInteractive -Command $snapshotCommand 2>&1)
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "Could not read execution-policy state from '$EnginePath': $([string]::Join([Environment]::NewLine, [string[]]$output))"
    }

    return [string]::Join([Environment]::NewLine, [string[]]$output).Trim()
}

function ConvertTo-PolicySnapshot {
    param(
        [object[]]$Policies,
        [switch]$ExcludeProcess
    )

    $normalized = @($Policies | Where-Object {
        (-not $ExcludeProcess) -or ([string]$_.Scope -ne 'Process')
    } | Sort-Object Scope | ForEach-Object {
        [pscustomobject]@{
            Scope = [string]$_.Scope
            ExecutionPolicy = [string]$_.ExecutionPolicy
        }
    })

    return [string](ConvertTo-Json -InputObject $normalized -Compress)
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$readmePath = Join-Path $repoRoot 'README.md'
$installerPath = Join-Path $repoRoot 'Install-OOShutUp10.ps1'
$elevationHelperPath = Join-Path $repoRoot 'Install-OOShutUp10-Elevation.ps1'
$batchPath = Join-Path $repoRoot 'install.bat'
$readme = [System.IO.File]::ReadAllText($readmePath)
$installer = [System.IO.File]::ReadAllText($installerPath)
$elevationHelper = [System.IO.File]::ReadAllText($elevationHelperPath)
$batch = [System.IO.File]::ReadAllText($batchPath)
$fallback = 'PowerShell policy blocks this production launch. No policy was changed; restore a trusted signed release or use -DevelopmentMode only from a source checkout.'

$instructionExtensions = @('.md', '.ps1', '.bat', '.cmd', '.sh', '.html', '.txt')
$directoryStack = New-Object System.Collections.Stack
$instructionFiles = New-Object System.Collections.ArrayList
$directoryStack.Push($repoRoot)
while ($directoryStack.Count -gt 0) {
    $currentDirectory = [string]$directoryStack.Pop()
    foreach ($entry in Get-ChildItem -LiteralPath $currentDirectory -Force) {
        if ($entry.PSIsContainer) {
            if (@('.git', 'node_modules') -notcontains $entry.Name) {
                $directoryStack.Push($entry.FullName)
            }
            continue
        }

        if ($instructionExtensions -contains $entry.Extension.ToLowerInvariant()) {
            [void]$instructionFiles.Add($entry)
        }
    }
}

foreach ($file in $instructionFiles) {
    $lines = [System.IO.File]::ReadAllLines($file.FullName)
    for ($lineIndex = 0; $lineIndex -lt $lines.Length; $lineIndex++) {
        if ($lines[$lineIndex] -notmatch '(?i)\bSet-ExecutionPolicy\b') {
            continue
        }

        $commandText = $lines[$lineIndex]
        while ($commandText.TrimEnd().EndsWith('`') -and ($lineIndex + 1) -lt $lines.Length) {
            $lineIndex++
            $commandText += ' ' + $lines[$lineIndex]
        }

        $hasPersistentScope = $commandText -match '(?i)-Scope\s+(CurrentUser|LocalMachine)\b'
        $setsUnrestricted = $commandText -match '(?i)-ExecutionPolicy\s+Unrestricted\b'
        Assert-True (-not ($hasPersistentScope -and $setsUnrestricted)) "no instruction sets a persistent scope to Unrestricted ($($file.FullName):$($lineIndex + 1))"
    }
}

Assert-True ($readme.Contains($fallback)) 'README gives the exact safe fallback for a policy-blocked launch'
Assert-True ($installer.Contains($fallback)) 'script help gives the exact safe fallback for a policy-blocked launch'
Assert-True ($readme.Contains('MachinePolicy` and `UserPolicy`') -and $readme.Contains('`Process`, then `CurrentUser`, then `LocalMachine`')) 'README explains execution-policy precedence'
Assert-True ($installer.Contains('MachinePolicy and UserPolicy Group Policy settings') -and $installer.Contains('Process takes precedence over CurrentUser and LocalMachine')) 'script help explains execution-policy precedence'
Assert-True ($readme.Contains('powershell.exe -NoProfile -ExecutionPolicy Bypass -File')) 'README documents the Windows PowerShell 5.1 process-scoped launch'
Assert-True ($readme.Contains('pwsh.exe -NoProfile -ExecutionPolicy Bypass -File')) 'README documents the PowerShell 7 process-scoped launch'
Assert-True ($readme.Contains('That path may still apply the reviewed profile') -and $readme.Contains('never relocates the checkout or registers a scheduled task')) 'README labels the mutation and task limits of development mode'
Assert-True ($installer.Contains('powershell.exe -NoProfile -ExecutionPolicy Bypass -File')) 'script help documents the Windows PowerShell 5.1 process-scoped launch'
Assert-True ($installer.Contains('pwsh.exe -NoProfile -ExecutionPolicy Bypass -File')) 'script help documents the PowerShell 7 process-scoped launch'
Assert-True ($installer.Contains('Development mode is available only for a Git source checkout') -and $installer.Contains('does not relocate the checkout or create a scheduled task')) 'script help limits unsigned execution to explicit checkout development mode'
Assert-True ($batch -match '(?im)^pwsh -NoProfile -ExecutionPolicy AllSigned -File ') 'batch wrapper requires signed scripts for production launch'
Assert-True ($batch -match '(?im)^powershell -NoProfile -ExecutionPolicy AllSigned -File .*Install-OOShutUp10-Elevation\.ps1') 'batch wrapper requires a signed elevation helper'
Assert-True ($elevationHelper.Contains('Start-Process') -and $elevationHelper.Contains('-Verb RunAs') -and $elevationHelper.Contains('return [int]$process.ExitCode')) 'the elevation helper uses UAC and preserves the child status'
Assert-True ($batch.IndexOf('net session', [System.StringComparison]::Ordinal) -lt $batch.IndexOf('where pwsh', [System.StringComparison]::Ordinal)) 'the batch wrapper checks administrator access before runtime discovery'

$windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
$powerShell7 = Get-Command pwsh.exe -ErrorAction SilentlyContinue
Assert-True ($null -ne $windowsPowerShell) 'Windows PowerShell 5.1 is available for launch compatibility checks'
Assert-True ($null -ne $powerShell7) 'PowerShell 7 is available for launch compatibility checks'

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oo-execution-policy-test-' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($tempRoot)
$mockScriptPath = Join-Path $tempRoot 'mock-launch.ps1'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$mockScript = @'
param([Parameter(Mandatory = $true)][string]$SnapshotPath)
$policyList = @(Get-ExecutionPolicy -List)
$persistent = @($policyList | Where-Object { [string]$_.Scope -ne 'Process' } | Sort-Object Scope | ForEach-Object {
    [pscustomobject]@{ Scope = [string]$_.Scope; ExecutionPolicy = [string]$_.ExecutionPolicy }
})
$processPolicy = @($policyList | Where-Object { [string]$_.Scope -eq 'Process' }) | Select-Object -First 1
$result = [pscustomobject]@{
    Persistent = $persistent
    Process = if ($null -eq $processPolicy) { 'Undefined' } else { [string]$processPolicy.ExecutionPolicy }
}
[System.IO.File]::WriteAllText($SnapshotPath, [string](ConvertTo-Json -InputObject $result -Compress))
'@
[System.IO.File]::WriteAllText($mockScriptPath, $mockScript, $utf8NoBom)

try {
    foreach ($engine in @($windowsPowerShell, $powerShell7)) {
        $enginePolicyBeforeJson = Get-EngineExecutionPolicySnapshot -EnginePath $engine.Source
        $enginePolicyBefore = ConvertFrom-Json $enginePolicyBeforeJson
        $persistentPolicyBefore = ConvertTo-PolicySnapshot -Policies @($enginePolicyBefore.Policies) -ExcludeProcess
        $snapshotPath = Join-Path $tempRoot ($engine.Name + '.json')
        $launchOutput = @(& $engine.Source -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $mockScriptPath $snapshotPath 2>&1)
        $launchExitCode = $LASTEXITCODE
        Assert-Equal 0 $launchExitCode "$($engine.Name) runs the mocked process-scoped launch ($([string]::Join(' ', [string[]]$launchOutput)))"

        $launchResult = ConvertFrom-Json ([System.IO.File]::ReadAllText($snapshotPath))
        Assert-Equal 'Bypass' ([string]$launchResult.Process) "$($engine.Name) applies Bypass only to the mock process"
        $childPersistentPolicy = ConvertTo-PolicySnapshot -Policies @($launchResult.Persistent)
        Assert-Equal $persistentPolicyBefore $childPersistentPolicy "$($engine.Name) sees unchanged persistent policy values"

        $enginePolicyAfterJson = Get-EngineExecutionPolicySnapshot -EnginePath $engine.Source
        Assert-Equal $enginePolicyBeforeJson $enginePolicyAfterJson "$($engine.Name) has identical Get-ExecutionPolicy -List values before and after the mocked launch"
    }

    $mockBin = Join-Path $tempRoot 'mock-bin'
    [void][System.IO.Directory]::CreateDirectory($mockBin)
    $fakePwshPath = Join-Path $mockBin 'pwsh.cmd'
    $fakeNetPath = Join-Path $mockBin 'net.cmd'
    $batchCapturePath = Join-Path $tempRoot 'batch-arguments.txt'
    $fakePwsh = [string]::Join([Environment]::NewLine, @(
        '@echo off',
        '>"%POLICY_TEST_CAPTURE%" echo %*',
        'exit /b 0'
    )) + [Environment]::NewLine
    [System.IO.File]::WriteAllText($fakePwshPath, $fakePwsh, $utf8NoBom)

    $fakeNet = [string]::Join([Environment]::NewLine, @(
        '@echo off',
        'exit /b 0'
    )) + [Environment]::NewLine
    [System.IO.File]::WriteAllText($fakeNetPath, $fakeNet, $utf8NoBom)

    $testBatchPath = Join-Path $tempRoot 'install-under-test.bat'
    $testBatch = [System.Text.RegularExpressions.Regex]::Replace($batch, '(?m)^net session ', 'call net session ')
    $testBatch = $testBatch.Replace('pwsh -NoProfile -ExecutionPolicy AllSigned -File', 'call pwsh -NoProfile -ExecutionPolicy AllSigned -File')
    [System.IO.File]::WriteAllText($testBatchPath, $testBatch, [System.Text.Encoding]::ASCII)

    $batchRunnerPath = Join-Path $tempRoot 'run-mocked-batch.cmd'
    $batchRunner = [string]::Join([Environment]::NewLine, @(
        '@echo off',
        ('call "' + $testBatchPath + '" < nul'),
        'exit /b %ERRORLEVEL%'
    )) + [Environment]::NewLine
    [System.IO.File]::WriteAllText($batchRunnerPath, $batchRunner, $utf8NoBom)

    $originalPath = $env:PATH
    $originalCapturePath = $env:POLICY_TEST_CAPTURE
    try {
        $env:PATH = $mockBin + ';' + $originalPath
        $env:POLICY_TEST_CAPTURE = $batchCapturePath
        $resolvedPwsh = @(& where.exe pwsh 2>$null | ForEach-Object { [string]$_ })
        Assert-True (@($resolvedPwsh | Where-Object { $_ -ieq $fakePwshPath }).Count -gt 0) 'the batch mock resolves the fake PowerShell 7 launcher before invoking install.bat'

        $batchStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $batchStartInfo.FileName = $env:ComSpec
        $batchStartInfo.Arguments = '/d /c "' + $batchRunnerPath + '"'
        $batchStartInfo.WorkingDirectory = $repoRoot
        $batchStartInfo.UseShellExecute = $false
        $batchStartInfo.CreateNoWindow = $true
        $batchStartInfo.RedirectStandardOutput = $true
        $batchStartInfo.RedirectStandardError = $true
        $batchProcess = [System.Diagnostics.Process]::Start($batchStartInfo)
        $stdoutTask = $batchProcess.StandardOutput.ReadToEndAsync()
        $stderrTask = $batchProcess.StandardError.ReadToEndAsync()
        $batchProcess.WaitForExit()
        $batchOutput = $stdoutTask.Result + $stderrTask.Result
        Assert-Equal 0 $batchProcess.ExitCode "install.bat exits successfully with the fake pwsh child ($batchOutput)"
        Assert-True (Test-Path -LiteralPath $batchCapturePath) 'install.bat invoked the fake pwsh launcher instead of the installer'
        $batchArguments = [System.IO.File]::ReadAllText($batchCapturePath)
        Assert-True ($batchArguments.Contains('-NoProfile -ExecutionPolicy AllSigned -File')) 'install.bat passes AllSigned to its fake PowerShell 7 child'
        Assert-True ($batchArguments.Contains('Install-OOShutUp10.ps1')) 'install.bat forwards the installer path to the fake PowerShell 7 child'
    }
    finally {
        $env:PATH = $originalPath
        if ($null -eq $originalCapturePath) {
            Remove-Item Env:\POLICY_TEST_CAPTURE -ErrorAction SilentlyContinue
        }
        else {
            $env:POLICY_TEST_CAPTURE = $originalCapturePath
        }
    }

    [Console]::Out.WriteLine("Passed $script:AssertionCount execution-policy documentation and mock-launch assertions in $($PSVersionTable.PSVersion).")
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}
