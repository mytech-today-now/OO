[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:OriginalSystemRoot = $env:SystemRoot
$script:AssertionCount = 0
$script:LogMessages = New-Object System.Collections.ArrayList
$script:RestoreScenario = 'success'
$script:LastRestorePointDescription = $null
$script:LastMockRestoreProviderResult = $null
$script:LastListedRestorePoint = $null
$script:RestorePoints = @()
$script:ProcessBoundaryCalls = 0
$script:CurrentRestoreJob = $null
$script:LastRestoreDrive = $null
$script:RestoreEnableCalls = 0
$script:ConfigurationEvents = New-Object System.Collections.ArrayList
$script:ProfileSnapshotScenario = 'success'
$script:ActiveProfileSnapshots = New-Object System.Collections.ArrayList
$script:CreatedProfileSnapshotPaths = New-Object System.Collections.ArrayList
$script:ProfileFixtureBytes = [System.Text.Encoding]::UTF8.GetBytes('approved profile fixture')
$script:ManifestFixtureBytes = [System.Text.Encoding]::UTF8.GetBytes('approved manifest fixture')
$script:MockStartProcessFailure = $false
$script:FrequencyWriteScenario = 'success'
$script:FrequencyRestoreScenario = 'success'
$script:FrequencyRestoreWriteCount = 0
$script:RegistryFixture = [pscustomobject]@{
    KeyExists = $true
    ValueExists = $true
    Value = [int64]1440
    Kind = [Microsoft.Win32.RegistryValueKind]::QWord
}
$script:OOShutUpRestorePointFailureMessage = 'A restore point could not be verified. No privacy settings were applied; free space or enable restore protection, then retry.'
$script:OOShutUpRestorePointPreparationFailureMessage = 'Restore-point preparation failed on the detected Windows volume. No settings were applied; restore the prior registry value and fix the provider error.'
$script:OOShutUpRestorePointOverrideMessage = 'A restore point could not be verified. System Restore recovery is unavailable; continuing only because -AllowWithoutRestorePoint was supplied. The approved O&O profile export remains separate.'
$script:OOShutUpApprovalFailureMessage = 'The approved profile did not match the test approval.'
$script:OOShutUpApprovalConfirmed = $true
$script:OOShutUpApprovedDigest = 'profile-digest'
$script:OOShutUpApprovedManifestDigest = 'manifest-digest'
$script:OOShutUpLastRunOutcome = $null
$script:OOShutUpChildTimeoutSeconds = 900
$script:OOShutUpTaskExecutionTimeLimit = 'PT30M'
$script:MockOOSChildExitRecorded = $false
$script:MockOOSChildExitCode = 30
$script:InstallPath = $null
$script:OOShutUpMutationMutexName = "Global\myTech.Today.OOShutUp.RestorePointTest.$PID"
$script:OOShutUpMutationLockTimeoutMilliseconds = 600000

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

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    [void]$script:LogMessages.Add("[$Level] $Message")
}

function Get-RestorePointFrequencyState {
    [CmdletBinding()]
    param()

    return [pscustomobject]@{
        KeyExists = [bool]$script:RegistryFixture.KeyExists
        ValueExists = [bool]$script:RegistryFixture.ValueExists
        Value = $script:RegistryFixture.Value
        Kind = $script:RegistryFixture.Kind
    }
}

function Set-RestorePointFrequencyRawValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Value,
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryValueKind]$Kind
    )

    if ($script:FrequencyWriteScenario -eq 'rejected' -and $Kind -eq [Microsoft.Win32.RegistryValueKind]::DWord) {
        throw 'Mock registry access denied.'
    }
    if ($script:FrequencyWriteScenario -eq 'ignored' -and $Kind -eq [Microsoft.Win32.RegistryValueKind]::DWord) {
        return
    }
    if ($script:FrequencyRestoreScenario -eq 'failure' -and $Kind -eq [Microsoft.Win32.RegistryValueKind]::DWord) {
        $script:FrequencyRestoreWriteCount++
        if ($script:FrequencyRestoreWriteCount -gt 1) {
            throw 'Mock registry rollback failure.'
        }
    }

    $script:RegistryFixture.KeyExists = $true
    $script:RegistryFixture.ValueExists = $true
    $script:RegistryFixture.Value = $Value
    $script:RegistryFixture.Kind = $Kind
}

function Remove-RestorePointFrequencyValue {
    [CmdletBinding()]
    param()

    $script:RegistryFixture.ValueExists = $false
    $script:RegistryFixture.Value = $null
    $script:RegistryFixture.Kind = $null
}

function Remove-EmptyRestorePointFrequencyKey {
    [CmdletBinding()]
    param()

    if ($script:RegistryFixture.ValueExists) {
        throw 'Mock registry key is not empty.'
    }
    $script:RegistryFixture.KeyExists = $false
}

function Enable-ComputerRestore {
    [CmdletBinding()]
    param([string]$Drive)

    $script:LastRestoreDrive = $Drive
    $script:RestoreEnableCalls++
    if ($script:RestoreScenario -eq 'enable-failure') {
        throw 'Mock restore protection enable failure.'
    }
}

function Start-Job {
    [CmdletBinding()]
    param([scriptblock]$ScriptBlock, [object[]]$ArgumentList)

    if ($script:RestoreScenario -eq 'job-start-exception') {
        throw 'Mock job startup exception.'
    }

    $script:LastRestorePointDescription = [string]$ArgumentList[0]
    $state = if ($script:RestoreScenario -eq 'timeout') { 'Running' } else { 'Completed' }
    $script:CurrentRestoreJob = [pscustomobject]@{ State = $state; Id = 123 }
    return $script:CurrentRestoreJob
}

function Start-Sleep {
    [CmdletBinding()]
    param([int]$Seconds)
}

function Receive-Job {
    [CmdletBinding()]
    param([object]$Job, [switch]$Wait, [switch]$AutoRemoveJob)

    if ($script:RestoreScenario -eq 'receive-job-exception') {
        throw 'Mock provider receive exception.'
    }
    if ($script:RestoreScenario -eq 'success' -or $script:RestoreScenario -eq 'false-success' -or $script:RestoreScenario -eq 'enumeration-exception') {
        $script:LastMockRestoreProviderResult = @{ Success = $true; Message = 'Mock provider returned success.' }
        return $script:LastMockRestoreProviderResult
    }

    $message = switch -Exact ($script:RestoreScenario) {
        'protection-disabled' { 'System Restore protection is disabled.' }
        'no-space' { 'There is not enough disk space.' }
        'provider-unavailable' { 'The System Restore provider is unavailable.' }
        'provider-exception' { 'The provider threw an exception.' }
        default { 'Mock provider rejected the request.' }
    }
    $script:LastMockRestoreProviderResult = @{ Success = $false; Message = $message }
    return $script:LastMockRestoreProviderResult
}

function Stop-Job {
    [CmdletBinding()]
    param([object]$Job)
}

function Remove-Job {
    [CmdletBinding()]
    param([object]$Job, [switch]$Force)
}

function Write-Progress {
    [CmdletBinding()]
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [string]$CurrentOperation,
        [switch]$Completed
    )
}

function Get-ComputerRestorePoint {
    [CmdletBinding()]
    param()

    if ($script:RestoreScenario -eq 'enumeration-exception') {
        throw 'Mock restore-point enumeration exception.'
    }
    if ($script:RestoreScenario -eq 'success') {
        [void]$script:ConfigurationEvents.Add('restore-point-verified')
        $script:LastListedRestorePoint = [pscustomobject]@{
            Description = $script:LastRestorePointDescription
            SequenceNumber = 42
        }
        return @($script:LastListedRestorePoint)
    }

    $script:LastListedRestorePoint = $null
    return $script:RestorePoints
}

function Update-ScriptProgress {
    param(
        [int]$PercentComplete,
        [string]$Status,
        [string]$CurrentOperation
    )
}

function Get-OOShutUpApprovedProfile {
    [CmdletBinding()]
    param([switch]$RequireMachineScope)

    [void]$script:ConfigurationEvents.Add('profile-validation')
    return [pscustomobject]@{
        Sha256 = $script:OOShutUpApprovedDigest
        ManifestSha256 = $script:OOShutUpApprovedManifestDigest
        Version = 'test-1.0'
        OOShutUpVersionMinimum = '3.6.1135.0'
        OOShutUpVersionMaximum = '3.6.1135.0'
        SettingCount = 1
        Path = 'C:\fixtures\ooshutup10.cfg'
    }
}

function New-OOShutUpProfileSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][pscustomobject]$Profile)

    [void]$script:ConfigurationEvents.Add('snapshot-creation')
    if ($script:ProfileSnapshotScenario -eq 'creation-failure') {
        throw 'Mocked profile snapshot creation failure.'
    }

    $directory = Join-Path 'C:\fixtures' ('ProfileSnapshot-{0}' -f [guid]::NewGuid().ToString('N'))
    if ($script:CreatedProfileSnapshotPaths.Contains($directory)) {
        throw 'Mocked profile snapshot path was reused.'
    }
    [void]$script:CreatedProfileSnapshotPaths.Add($directory)
    [void]$script:ActiveProfileSnapshots.Add($directory)
    [void]$script:ConfigurationEvents.Add('snapshot-verification')

    $profileBytes = [byte[]]$script:ProfileFixtureBytes
    $manifestBytes = [byte[]]$script:ManifestFixtureBytes
    if ($script:ProfileSnapshotScenario -eq 'digest-mismatch') {
        $profileBytes = [System.Text.Encoding]::UTF8.GetBytes('tampered snapshot profile fixture')
    }

    $profileDigest = Get-OOShutUpProfileDigest -Bytes $profileBytes
    $manifestDigest = Get-OOShutUpProfileDigest -Bytes $manifestBytes
    if (-not [string]::Equals($profileDigest, [string]$Profile.Sha256, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($manifestDigest, [string]$Profile.ManifestSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        Remove-OOShutUpProfileSnapshot -Directory $directory
        throw 'Mocked copied profile snapshot digest verification failed.'
    }

    return [pscustomobject]@{
        Directory = $directory
        Path = Join-Path $directory 'ooshutup10.cfg'
        ManifestPath = Join-Path $directory 'ooshutup10.approval.json'
    }
}

function Remove-OOShutUpProfileSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Directory)

    [void]$script:ConfigurationEvents.Add('snapshot-removal')
    if ($script:ProfileSnapshotScenario -eq 'cleanup-failure') {
        throw 'Mocked profile snapshot cleanup failure.'
    }

    [void]$script:ActiveProfileSnapshots.Remove($Directory)
}

function Test-OOShutUpPayload {
    [CmdletBinding()]
    param(
        [string]$Path,
        [switch]$Quiet,
        [string]$ReviewedMinimumVersion,
        [string]$ReviewedMaximumVersion
    )

    return $true
}

function Start-Process {
    [CmdletBinding()]
    param(
        [string]$FilePath,
        [string]$ArgumentList,
        [string]$WorkingDirectory,
        [switch]$Wait,
        [switch]$PassThru,
        [switch]$NoNewWindow
    )

    [void]$script:ConfigurationEvents.Add('process-start')
    $script:ProcessBoundaryCalls++
    if ($script:MockStartProcessFailure) {
        throw 'Mocked Start-Process outcome could not be verified.'
    }
    $process = [pscustomobject]@{
        Id = 123
        ExitCode = $script:MockOOSChildExitCode
        HasExited = $false
    }
    $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value {
        param([int]$Milliseconds = -1)

        if ($Milliseconds -lt 0) {
            return [bool]$this.HasExited
        }
        $this.HasExited = $true
        if (-not $script:MockOOSChildExitRecorded) {
            [void]$script:ConfigurationEvents.Add('process-exited')
            $script:MockOOSChildExitRecorded = $true
        }
        return $true
    } -Force
    $script:MockOOSChildExitRecorded = $false
    return $process
}

try {
    $env:SystemRoot = 'D:\Windows'
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $scriptPath = Join-Path $repoRoot 'Install-OOShutUp10.ps1'
    $tokens = $null
    $parseErrors = $null
    $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath,
        [ref]$tokens,
        [ref]$parseErrors
    )
    Assert-Equal 0 @($parseErrors).Count 'the installer parses in the active PowerShell engine'
    $source = Get-Content -LiteralPath $scriptPath -Raw
    Assert-True ($source.Contains('[switch]$AllowWithoutRestorePoint')) 'the recovery override is an explicit switch parameter'
    Assert-True ($source.Contains("if (`$AllowWithoutRestorePoint) { `$relocateParams += '-AllowWithoutRestorePoint' }")) 'the explicit override survives script relocation'
    Assert-True ($source.Contains($script:OOShutUpRestorePointPreparationFailureMessage)) 'restore preparation has the exact fail-closed status'
    Assert-True ($source -notmatch 'Set-ItemProperty\s+-[^\r\n]*-Type\b') 'Set-ItemProperty is not passed the unsupported Type parameter'
    Assert-True (-not $source.Contains('Set-RestorePointFrequency | Out-Null')) 'the main flow no longer discards restore preparation status'
    Assert-True ($source.Contains("`$script:OOShutUpMutationMutexName = 'Global\myTech.Today.OOShutUp.Mutation'")) 'manual and SYSTEM runs use the same Global named mutex'
    Assert-True ($source.Contains('$script:OOShutUpMutationLockTimeoutMilliseconds = 600000')) 'the machine-wide mutex has a documented ten-minute bounded wait'

    $requiredFunctions = @(
        'New-OOShutUpMutationMutex',
        'Open-OOShutUpMutationMutex',
        'Release-OOShutUpMutationMutex',
        'Invoke-WithOOShutUpMutationLock',
        'Set-OOShutUpLogStage',
        'Set-OOShutUpLogErrorCategory',
        'Set-OOShutUpLogOutcome',
        'Get-OOShutUpRetainedProfileSnapshots',
        'Wait-OOShutUpChildProcess',
        'New-OOShutUpConfigurationOutcome',
        'Test-RestorePointFrequencyStateEqual',
        'Restore-RestorePointFrequencyState',
        'Set-RestorePointFrequencyValue',
        'Set-RestorePointFrequency',
        'Get-WindowsSystemVolume',
        'Get-OOShutUpProfileDigest',
        'Invoke-OOShutUpConfigurationCore',
        'Invoke-OOShutUpConfiguration',
        'New-SystemRestorePointCore',
        'New-SystemRestorePoint',
        'Invoke-OOShutUpConfigurationWithRestorePointGateCore',
        'Invoke-OOShutUpConfigurationWithRestorePointGate'
    )
    $functionAsts = $scriptAst.FindAll({
        param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
        ($requiredFunctions -contains $node.Name)
    }, $true)
    foreach ($name in $requiredFunctions) {
        $functionAst = @($functionAsts | Where-Object { $_.Name -eq $name })[0]
        Assert-True ($null -ne $functionAst) "the $name function exists"
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }

    $script:OOShutUpApprovedDigest = Get-OOShutUpProfileDigest -Bytes ([byte[]]$script:ProfileFixtureBytes)
    $script:OOShutUpApprovedManifestDigest = Get-OOShutUpProfileDigest -Bytes ([byte[]]$script:ManifestFixtureBytes)

    $setFrequencySource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Set-RestorePointFrequency' })[0].Extent.Text)
    $newRestorePointSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'New-SystemRestorePointCore' })[0].Extent.Text)
    Assert-True ($setFrequencySource.Contains('[Microsoft.Win32.RegistryValueKind]::DWord')) 'the temporary value is explicitly written as DWORD'
    Assert-True (-not $newRestorePointSource.Contains('Enable-ComputerRestore -Drive "C:"')) 'restore protection does not assume the C drive'
    Assert-True ($newRestorePointSource.Contains('Restore-RestorePointFrequencyState -State $frequencyPreparation.PreviousState')) 'the captured registry state is restored after restore-point preparation'

    $lockFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-WithOOShutUpMutationLock' })[0].Extent.Text)
    $mutexFactorySource = [string](@($functionAsts | Where-Object { $_.Name -eq 'New-OOShutUpMutationMutex' })[0].Extent.Text)
    $lockWaitPosition = $lockFunctionSource.IndexOf('$mutex.WaitOne($TimeoutMilliseconds)', [System.StringComparison]::Ordinal)
    $operationStartPosition = $lockFunctionSource.IndexOf('$operationStarted = $true', [System.StringComparison]::Ordinal)
    Assert-True ($mutexFactorySource.Contains("'S-1-5-18'") -and $mutexFactorySource.Contains("'S-1-5-32-544'")) 'the named mutex ACL grants Local System and Administrators access'
    Assert-True ($lockWaitPosition -ge 0 -and $operationStartPosition -gt $lockWaitPosition) 'protected registry and O&O work begins only after the bounded mutex wait succeeds'
    Assert-True ($lockFunctionSource.Contains('[System.Threading.AbandonedMutexException]') -and $lockFunctionSource.Contains('$failureType = ''Abandoned''')) 'abandoned ownership fails before the operation runs'
    Assert-True ($lockFunctionSource.Contains('$failureType -ne ''Abandoned''')) 'abandoned ownership remains held until the failing process exits so queued runs cannot proceed'
    $configurationWrapperSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpConfiguration' })[0].Extent.Text)
    $restorePointWrapperSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'New-SystemRestorePoint' })[0].Extent.Text)
    $gateFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpConfigurationWithRestorePointGateCore' })[0].Extent.Text)
    $gateWrapperSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpConfigurationWithRestorePointGate' })[0].Extent.Text)
    Assert-True ($configurationWrapperSource.Contains('Invoke-WithOOShutUpMutationLock') -and $configurationWrapperSource.Contains('Invoke-OOShutUpConfigurationCore')) 'SYSTEM ReapplyOnly reaches O&O only through the shared mutex wrapper'
    Assert-True ($restorePointWrapperSource.Contains('Invoke-WithOOShutUpMutationLock')) 'direct restore-point entry acquires the same mutex before registry work'
    Assert-True ($gateWrapperSource.Contains('Invoke-WithOOShutUpMutationLock') -and $gateWrapperSource.Contains('Invoke-OOShutUpConfigurationWithRestorePointGateCore')) 'manual restore-point and O&O work is covered by one outer mutex ownership interval'
    Assert-True ($gateFunctionSource.Contains('New-SystemRestorePointCore') -and $gateFunctionSource.Contains('Invoke-OOShutUpConfigurationCore')) 'the manual path keeps restore-point registry work and O&O execution in the same lock interval'
    Assert-True ($source.Contains("`$script:OOShutUpUpdateMultipleInstancesPolicy = 'Queue'")) 'the scheduled task Queue policy remains unchanged'

    $env:SystemRoot = 'C:\Windows'
    $detectedVolume = Get-WindowsSystemVolume
    Assert-Equal 'C:\' $detectedVolume 'a C drive Windows installation resolves to C'
    $env:SystemRoot = 'D:\Windows'
    $detectedVolume = Get-WindowsSystemVolume
    Assert-Equal 'D:\' $detectedVolume 'the Windows system volume is detected from SystemRoot'

    $priorRegistryState = [pscustomobject]@{
        KeyExists = $true
        ValueExists = $true
        Value = [int64]1440
        Kind = [Microsoft.Win32.RegistryValueKind]::QWord
    }
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $priorRegistryState.KeyExists
        ValueExists = $priorRegistryState.ValueExists
        Value = $priorRegistryState.Value
        Kind = $priorRegistryState.Kind
    }
    $script:FrequencyWriteScenario = 'success'
    $frequencyResult = Set-RestorePointFrequency
    Assert-Equal $true $frequencyResult.Succeeded 'the supported registry write succeeds in the fixture'
    Assert-Equal ([Microsoft.Win32.RegistryValueKind]::DWord) $script:RegistryFixture.Kind 'the temporary frequency value is DWORD'
    Assert-Equal 0 $script:RegistryFixture.Value 'the temporary frequency value is zero'
    Assert-Equal ([Microsoft.Win32.RegistryValueKind]::QWord) $frequencyResult.PreviousState.Kind 'the previous registry type is captured'
    Assert-Equal ([int64]1440) $frequencyResult.PreviousState.Value 'the previous registry data is captured'
    Assert-True (Restore-RestorePointFrequencyState -State $frequencyResult.PreviousState) 'rollback restores an existing value and its type'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $priorRegistryState -Actual (Get-RestorePointFrequencyState)) 'rollback leaves the existing value and QWORD type unchanged'

    $missingValueState = [pscustomobject]@{ KeyExists = $true; ValueExists = $false; Value = $null; Kind = $null }
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $missingValueState.KeyExists
        ValueExists = $missingValueState.ValueExists
        Value = $missingValueState.Value
        Kind = $missingValueState.Kind
    }
    $frequencyResult = Set-RestorePointFrequency
    Assert-Equal $true $frequencyResult.Succeeded 'a missing frequency value can be created temporarily'
    Assert-True (Restore-RestorePointFrequencyState -State $frequencyResult.PreviousState) 'rollback removes a value that did not previously exist'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $missingValueState -Actual (Get-RestorePointFrequencyState)) 'rollback preserves the pre-existing empty key'

    $missingKeyState = [pscustomobject]@{ KeyExists = $false; ValueExists = $false; Value = $null; Kind = $null }
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $missingKeyState.KeyExists
        ValueExists = $missingKeyState.ValueExists
        Value = $missingKeyState.Value
        Kind = $missingKeyState.Kind
    }
    $frequencyResult = Set-RestorePointFrequency
    Assert-Equal $true $frequencyResult.Succeeded 'a missing registry key can be created temporarily'
    Assert-True (Restore-RestorePointFrequencyState -State $frequencyResult.PreviousState) 'rollback removes the temporary value and key'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $missingKeyState -Actual (Get-RestorePointFrequencyState)) 'rollback restores a previously missing key'

    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $priorRegistryState.KeyExists
        ValueExists = $priorRegistryState.ValueExists
        Value = $priorRegistryState.Value
        Kind = $priorRegistryState.Kind
    }
    $script:FrequencyWriteScenario = 'rejected'
    $frequencyResult = Set-RestorePointFrequency
    Assert-Equal $false $frequencyResult.Succeeded 'a rejected registry write fails closed'
    Assert-Equal $true $frequencyResult.RollbackSucceeded 'a rejected registry write reports successful rollback'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $priorRegistryState -Actual (Get-RestorePointFrequencyState)) 'a rejected registry write leaves the prior data and type intact'

    $script:FrequencyWriteScenario = 'ignored'
    $frequencyResult = Set-RestorePointFrequency
    Assert-Equal $false $frequencyResult.Succeeded 'a registry write that does not take effect fails verification'
    Assert-Equal $true $frequencyResult.RollbackSucceeded 'a failed readback is rolled back'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $priorRegistryState -Actual (Get-RestorePointFrequencyState)) 'failed write verification preserves the prior registry state'
    $script:FrequencyWriteScenario = 'success'

    $pointCreationPosition = $gateFunctionSource.IndexOf('New-SystemRestorePoint', [System.StringComparison]::Ordinal)
    $failureGatePosition = $gateFunctionSource.IndexOf('if (-not $restorePointVerified -and -not $AllowWithoutRestorePoint)', [System.StringComparison]::Ordinal)
    $configurationPosition = $gateFunctionSource.IndexOf('Invoke-OOShutUpConfigurationCore -ExecutablePath', [System.StringComparison]::Ordinal)
    Assert-True ($pointCreationPosition -ge 0 -and $failureGatePosition -gt $pointCreationPosition -and $configurationPosition -gt $failureGatePosition) 'restore verification and the default failure gate precede O&O configuration'

    $mainSource = $source.Substring($source.IndexOf('#region Main Script', [System.StringComparison]::Ordinal))
    $mainGatePosition = $mainSource.IndexOf('if (-not $configurationResult.ConfigurationStarted)', [System.StringComparison]::Ordinal)
    $mainFailureExitPosition = $mainSource.IndexOf('exit (Get-OOShutUpTerminalExitCode -Outcome Failed -ExitCode 1)', $mainGatePosition, [System.StringComparison]::Ordinal)
    Assert-True ($mainGatePosition -ge 0 -and $mainFailureExitPosition -gt $mainGatePosition) 'a blocked restore gate records failure before the installer exits nonzero'

    $restoreCases = @(
        @{ Name = 'protection disabled'; Scenario = 'protection-disabled'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'insufficient space'; Scenario = 'no-space'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'provider unavailable'; Scenario = 'provider-unavailable'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'provider exception'; Scenario = 'provider-exception'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'job startup exception'; Scenario = 'job-start-exception'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'receive exception'; Scenario = 'receive-job-exception'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'provider timeout'; Scenario = 'timeout'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'false-success response'; Scenario = 'false-success'; ExpectedCreated = $false; ExpectedPreparation = $true },
        @{ Name = 'System Restore enable failure'; Scenario = 'enable-failure'; ExpectedCreated = $false; ExpectedPreparation = $false },
        @{ Name = 'verified creation'; Scenario = 'success'; ExpectedCreated = $true; ExpectedPreparation = $true }
    )

    foreach ($case in $restoreCases) {
        $script:RestoreScenario = $case.Scenario
        $script:RestorePoints = @()
        $script:RegistryFixture = [pscustomobject]@{
            KeyExists = $priorRegistryState.KeyExists
            ValueExists = $priorRegistryState.ValueExists
            Value = $priorRegistryState.Value
            Kind = $priorRegistryState.Kind
        }
        $script:LogMessages.Clear()
        $created = New-SystemRestorePoint
        Assert-Equal $case.ExpectedCreated $created.RestorePointVerified "$($case.Name) returns the expected verified status"
        Assert-Equal $case.ExpectedPreparation $created.PreparationSucceeded "$($case.Name) returns the expected preparation status"
        Assert-Equal 'D:\' $created.SystemVolume "$($case.Name) uses the detected non-C Windows volume"
        Assert-Equal 'D:\' $script:LastRestoreDrive "$($case.Name) enables restore protection on the detected volume"
        Assert-True (Test-RestorePointFrequencyStateEqual -Expected $priorRegistryState -Actual (Get-RestorePointFrequencyState)) "$($case.Name) restores the prior registry value and type"
        $messages = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
        if ($case.Scenario -eq 'success') {
            Assert-True ($messages.Contains('System restore point created and verified successfully')) 'success is logged only after Windows lists the exact point'
        }
        else {
            Assert-True (-not $messages.Contains('System restore point created and verified successfully')) "$($case.Name) does not claim a verified point"
        }
        if ($case.Scenario -eq 'timeout') {
            Assert-True ($messages.Contains('Restore point creation timed out after 300 seconds')) 'the timeout status is recorded'
        }
        if ($case.Scenario -eq 'false-success') {
            Assert-True ($messages.Contains('The restore provider reported success, but Windows did not list the requested restore point')) 'a false-success response is rejected when Windows has no matching point'
        }
    }

    $callCountBeforeFailure = $script:ProcessBoundaryCalls
    $blockedScenarios = @('protection-disabled', 'no-space', 'provider-unavailable', 'provider-exception', 'job-start-exception', 'receive-job-exception', 'enumeration-exception', 'timeout', 'false-success', 'enable-failure')
    foreach ($scenario in $blockedScenarios) {
        $script:RestoreScenario = $scenario
        $script:RestorePoints = @()
        $script:RegistryFixture = [pscustomobject]@{
            KeyExists = $priorRegistryState.KeyExists
            ValueExists = $priorRegistryState.ValueExists
            Value = $priorRegistryState.Value
            Kind = $priorRegistryState.Kind
        }
        $script:LogMessages.Clear()
        $eventsBeforeBlockedRestore = $script:ConfigurationEvents.Count
        $blockedResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
        Assert-Equal $false $blockedResult.ConfigurationStarted "$scenario stops before configuration starts"
        Assert-Equal 'Failed' $blockedResult.ConfigurationOutcome.State "$scenario reports a pre-launch settings failure distinctly"
        Assert-Equal $callCountBeforeFailure $script:ProcessBoundaryCalls "$scenario never reaches Start-Process"
        Assert-Equal $eventsBeforeBlockedRestore $script:ConfigurationEvents.Count "$scenario never reaches profile validation or snapshot creation"
        $blockedMessages = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
        if ($scenario -eq 'enable-failure') {
            Assert-True ($blockedMessages.Contains($script:OOShutUpRestorePointPreparationFailureMessage)) "$scenario gives the exact preparation failure status"
        }
        else {
            Assert-True ($blockedMessages.Contains($script:OOShutUpRestorePointFailureMessage)) "$scenario gives the exact actionable retry status"
        }
    }

    $script:RestoreScenario = 'success'
    $script:FrequencyWriteScenario = 'success'
    $script:FrequencyRestoreScenario = 'failure'
    $script:FrequencyRestoreWriteCount = 0
    $rollbackFailurePriorState = [pscustomobject]@{
        KeyExists = $true
        ValueExists = $true
        Value = 0
        Kind = [Microsoft.Win32.RegistryValueKind]::DWord
    }
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $rollbackFailurePriorState.KeyExists
        ValueExists = $rollbackFailurePriorState.ValueExists
        Value = $rollbackFailurePriorState.Value
        Kind = $rollbackFailurePriorState.Kind
    }
    $script:LogMessages.Clear()
    $script:ConfigurationEvents.Clear()
    $processCallsBeforeRollbackFailure = $script:ProcessBoundaryCalls
    $rollbackFailureResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
    Assert-Equal $false $rollbackFailureResult.PreparationSucceeded 'registry restore failure is a failed preparation'
    Assert-Equal $false $rollbackFailureResult.RestorePointVerified 'registry restore failure clears verified status'
    Assert-Equal $false $rollbackFailureResult.ConfigurationStarted 'registry restore failure blocks configuration'
    Assert-Equal $processCallsBeforeRollbackFailure $script:ProcessBoundaryCalls 'registry restore failure never reaches Start-Process'
    Assert-True (-not $script:ConfigurationEvents.Contains('profile-validation')) 'registry restore failure stops before profile validation'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpRestorePointPreparationFailureMessage)) 'registry restore failure reports that no settings were applied'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('Mock registry rollback failure.')) 'registry restore failure detail is retained in diagnostics'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $rollbackFailurePriorState -Actual (Get-RestorePointFrequencyState)) 'registry restore failure leaves the fixture value and type unchanged'
    $script:FrequencyRestoreScenario = 'success'
    $script:FrequencyRestoreWriteCount = 0

    $script:RestoreScenario = 'success'
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $priorRegistryState.KeyExists
        ValueExists = $priorRegistryState.ValueExists
        Value = $priorRegistryState.Value
        Kind = $priorRegistryState.Kind
    }
    $script:RestorePoints = @()
    $script:LogMessages.Clear()
    $script:ConfigurationEvents.Clear()
    $script:LastMockRestoreProviderResult = $null
    $script:LastListedRestorePoint = $null
    $lockedSuccessResult = Invoke-WithOOShutUpMutationLock -Operation {
        Invoke-OOShutUpConfigurationWithRestorePointGateCore -ExecutablePath 'C:\fixtures\OOSU10.exe'
    }
    Assert-Equal $true $lockedSuccessResult.Succeeded 'the successful gate operation completes under the mutation lock'
    Assert-Equal $true $lockedSuccessResult.OperationStarted 'the successful gate operation starts only after acquiring the lock'
    Assert-Equal $true $script:LastMockRestoreProviderResult.Success 'the mocked restore provider reports success'
    Assert-True ($null -ne $script:LastListedRestorePoint) 'Windows lists a restore point after provider success'
    Assert-Equal $script:LastRestorePointDescription $script:LastListedRestorePoint.Description 'the listed point exactly matches the requested restore point'
    Assert-Equal 42 $script:LastListedRestorePoint.SequenceNumber 'the exact listed restore point is retained by the fixture'
    Assert-Equal $true $lockedSuccessResult.Value.RestorePointVerified 'verified point status survives the locked gate result object'
    $processCallsBeforePublicSuccess = $script:ProcessBoundaryCalls
    $script:LogMessages.Clear()
    $script:ConfigurationEvents.Clear()
    $successResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
    Assert-Equal $true $successResult.RestorePointVerified 'verified point status reaches the result'
    Assert-Equal $true $successResult.ConfigurationStarted 'verified point permits configuration'
    Assert-Equal 'Changed' $successResult.ConfigurationOutcome.State 'exit code 30 is retained as an O&O-reported changed outcome'
    Assert-Equal 'Unavailable' $successResult.ConfigurationOutcome.ReadbackState 'the mocked successful process does not imply independent settings readback'
    Assert-Equal ($processCallsBeforePublicSuccess + 1) $script:ProcessBoundaryCalls 'verified point reaches Start-Process exactly once'
    Assert-Equal 'restore-point-verified,profile-validation,snapshot-creation,snapshot-verification,process-start,process-exited,snapshot-removal' ([string]::Join(',', [string[]]$script:ConfigurationEvents.ToArray())) 'the production restore gate, profile validation, snapshot verification, process exit, and cleanup run in order'
    Assert-Equal 0 $script:ActiveProfileSnapshots.Count 'a completed O&O child removes its unique snapshot after process exit is known'

    $script:ConfigurationEvents.Clear()
    $script:OOShutUpApprovalConfirmed = $false
    $profileValidationStartCount = $script:ProcessBoundaryCalls
    $profileValidationFailure = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
    $script:OOShutUpApprovalConfirmed = $true
    Assert-Equal 'Failed' $profileValidationFailure.ConfigurationOutcome.State 'failed profile validation is distinguished from an unknown child outcome'
    Assert-Equal $profileValidationStartCount $script:ProcessBoundaryCalls 'failed approved-profile validation never launches O&O'
    Assert-Equal 'restore-point-verified,profile-validation' ([string]::Join(',', [string[]]$script:ConfigurationEvents.ToArray())) 'profile validation must pass before snapshot creation can begin'

    foreach ($snapshotFailureScenario in @('creation-failure', 'digest-mismatch')) {
        $script:ConfigurationEvents.Clear()
        $script:ProfileSnapshotScenario = $snapshotFailureScenario
        $snapshotFailureStartCount = $script:ProcessBoundaryCalls
        $snapshotFailureResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
        Assert-Equal 'Failed' $snapshotFailureResult.ConfigurationOutcome.State "$snapshotFailureScenario is a known pre-launch failure"
        Assert-Equal $snapshotFailureStartCount $script:ProcessBoundaryCalls "$snapshotFailureScenario never launches O&O"
        Assert-True (-not $script:ConfigurationEvents.Contains('process-start')) "$snapshotFailureScenario does not bypass the snapshot boundary"
        Assert-Equal 0 $script:ActiveProfileSnapshots.Count "$snapshotFailureScenario leaves no unverified snapshot behind"
        if ($snapshotFailureScenario -eq 'digest-mismatch') {
            Assert-True ($script:ConfigurationEvents.Contains('snapshot-removal')) 'a digest-mismatched snapshot is cleaned up before the no-launch result'
            Assert-Equal 0 $script:ActiveProfileSnapshots.Count 'a digest-mismatched snapshot is not retained'
        }
    }

    $script:ProfileSnapshotScenario = 'success'
    $script:ConfigurationEvents.Clear()
    $script:MockStartProcessFailure = $true
    $uncertainStartCount = $script:ProcessBoundaryCalls
    $uncertainStartResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
    $script:MockStartProcessFailure = $false
    Assert-Equal 'Unknown' $uncertainStartResult.ConfigurationOutcome.State 'a missing reliable process result remains unknown'
    Assert-Equal 'Unavailable' $uncertainStartResult.ConfigurationOutcome.ReadbackState 'an unknown process result is not settings verification'
    Assert-Equal ($uncertainStartCount + 1) $script:ProcessBoundaryCalls 'the uncertain process-start scenario reaches Start-Process once'
    Assert-Equal 1 $script:ActiveProfileSnapshots.Count 'an uncertain process-start outcome retains its profile snapshot'
    Assert-True (-not $script:ConfigurationEvents.Contains('snapshot-removal')) 'an uncertain process-start outcome does not clean up the snapshot'
    $uncertainSnapshotDirectory = [string]$script:ActiveProfileSnapshots[0]
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains("The O&O launch outcome could not be verified. The protected profile snapshot was retained at '$uncertainSnapshotDirectory'")) 'an uncertain process-start outcome reports the retained snapshot path'
    Remove-OOShutUpProfileSnapshot -Directory $uncertainSnapshotDirectory

    $script:ConfigurationEvents.Clear()
    $script:ProfileSnapshotScenario = 'cleanup-failure'
    $cleanupFailureStartCount = $script:ProcessBoundaryCalls
    $cleanupFailureResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
    Assert-Equal 'Changed' $cleanupFailureResult.ConfigurationOutcome.State 'cleanup failure preserves the O&O-reported changed status separately from cleanup success'
    Assert-Equal ($cleanupFailureStartCount + 1) $script:ProcessBoundaryCalls 'snapshot cleanup failure happens after the child exits'
    Assert-Equal 1 $script:ActiveProfileSnapshots.Count 'snapshot cleanup failure leaves the snapshot available for recovery'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('The protected profile snapshot could not be removed after O&O exited')) 'snapshot cleanup failure reports that O&O has exited'
    $cleanupFailureSnapshotDirectory = [string]$script:ActiveProfileSnapshots[0]
    $script:ProfileSnapshotScenario = 'success'
    Remove-OOShutUpProfileSnapshot -Directory $cleanupFailureSnapshotDirectory
    $script:ConfigurationEvents.Clear()

    $script:FrequencyWriteScenario = 'rejected'
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $priorRegistryState.KeyExists
        ValueExists = $priorRegistryState.ValueExists
        Value = $priorRegistryState.Value
        Kind = $priorRegistryState.Kind
    }
    $script:LogMessages.Clear()
    $processCallsBeforePreparationFailure = $script:ProcessBoundaryCalls
    $preparationFailureResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe' -AllowWithoutRestorePoint
    Assert-Equal $false $preparationFailureResult.PreparationSucceeded 'the restore override cannot bypass failed preparation'
    Assert-Equal $false $preparationFailureResult.ConfigurationStarted 'failed preparation blocks privacy settings even with the explicit restore-point override'
    Assert-Equal $processCallsBeforePreparationFailure $script:ProcessBoundaryCalls 'failed preparation never reaches Start-Process'
    $preparationFailureMessages = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($preparationFailureMessages.Contains($script:OOShutUpRestorePointPreparationFailureMessage)) 'failed preparation reports the exact fallback status'
    Assert-True (Test-RestorePointFrequencyStateEqual -Expected $priorRegistryState -Actual (Get-RestorePointFrequencyState)) 'failed preparation leaves the original registry state intact'
    $script:FrequencyWriteScenario = 'success'

    $script:RestoreScenario = 'provider-unavailable'
    $script:RestorePoints = @()
    $script:RegistryFixture = [pscustomobject]@{
        KeyExists = $priorRegistryState.KeyExists
        ValueExists = $priorRegistryState.ValueExists
        Value = $priorRegistryState.Value
        Kind = $priorRegistryState.Kind
    }
    $script:LogMessages.Clear()
    $processCallsBeforeOverride = $script:ProcessBoundaryCalls
    $overrideResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe' -AllowWithoutRestorePoint
    Assert-Equal $false $overrideResult.RestorePointVerified 'override does not represent a point as verified'
    Assert-Equal $true $overrideResult.OverrideUsed 'the explicit switch is recorded in the result'
    Assert-Equal $true $overrideResult.ConfigurationStarted 'the explicit override permits configuration'
    Assert-Equal ($processCallsBeforeOverride + 1) $script:ProcessBoundaryCalls 'only the explicit override reaches Start-Process'
    $overrideMessages = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($overrideMessages.Contains($script:OOShutUpRestorePointOverrideMessage)) 'override reports that System Restore recovery is unavailable'
    Assert-True ($overrideMessages -notmatch 'A system restore point was verified|created and verified successfully') 'override never claims a restore point'
    $uniqueSnapshotPaths = @($script:CreatedProfileSnapshotPaths | Select-Object -Unique)
    Assert-Equal $script:CreatedProfileSnapshotPaths.Count $uniqueSnapshotPaths.Count 'each snapshot invocation receives a unique per-run path'

    $mutexHelperSources = @(
        $mutexFactorySource,
        [string](@($functionAsts | Where-Object { $_.Name -eq 'Open-OOShutUpMutationMutex' })[0].Extent.Text),
        [string](@($functionAsts | Where-Object { $_.Name -eq 'Release-OOShutUpMutationMutex' })[0].Extent.Text),
        $lockFunctionSource
    )
    $mutexHelperText = [string]::Join("`r`n", $mutexHelperSources)
    $mutexWorkerBody = @'
$workDir = [string]$args[0]
$mutexName = [string]$args[1]
$role = [string]$args[2]
$eventPath = Join-Path $workDir 'events.txt'
$script:MutationLockWorkerDirectory = $workDir
$script:MutationLockWorkerRole = $role
$script:MutationLockOriginalFactory = (Get-Command New-OOShutUpMutationMutex).ScriptBlock
Set-Item -Path Function:New-OOShutUpMutationMutex -Value {
    param([string]$Name)
    $mutex = & $script:MutationLockOriginalFactory -Name $Name
    [System.IO.File]::WriteAllText((Join-Path $script:MutationLockWorkerDirectory "$script:MutationLockWorkerRole-mutex-opened"), '')
    return $mutex
}
[System.IO.File]::WriteAllText((Join-Path $workDir "$role-lock-name"), $mutexName)
[System.IO.File]::WriteAllText((Join-Path $workDir "$role-attempted"), '')
$lock = Invoke-WithOOShutUpMutationLock -MutexName $mutexName -TimeoutMilliseconds 60000 -Operation {
    [System.IO.File]::AppendAllText($eventPath, "$role-owner-enter`r`n")
    [System.IO.File]::WriteAllText((Join-Path $workDir "$role-registry-capture"), '')
    [System.IO.File]::AppendAllText($eventPath, "$role-registry-capture`r`n")
    if ($role -eq 'first') {
        $deadline = [DateTime]::UtcNow.AddSeconds(60)
        $releaseMarker = Join-Path $workDir 'release-first'
        while (-not [System.IO.File]::Exists($releaseMarker) -and [DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 20
        }
        if (-not [System.IO.File]::Exists($releaseMarker)) { throw 'The first run barrier timed out.' }
    }
    [System.IO.File]::WriteAllText((Join-Path $workDir "$role-ooshutup-start"), '')
    [System.IO.File]::AppendAllText($eventPath, "$role-ooshutup-start`r`n")
    [System.IO.File]::AppendAllText($eventPath, "$role-owner-exit`r`n")
    return "$role-complete"
}
if (-not $lock.Succeeded) {
    [Console]::Error.WriteLine($lock.FailureMessage)
    exit 2
}
exit 0
'@
    $mutexTestName = "$script:OOShutUpMutationMutexName.Barrier"
    $workerRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oosu-mutation-lock-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($workerRoot)
    $workerScriptPath = Join-Path $workerRoot 'mutex-worker.ps1'
    [System.IO.File]::WriteAllText($workerScriptPath, $mutexHelperText + "`r`n" + $mutexWorkerBody, [System.Text.Encoding]::UTF8)
    # Use the available Core host for subprocess barriers when possible; the in-process helpers still run in the caller's edition.
    $coreHost = Get-Command pwsh -ErrorAction SilentlyContinue
    $workerExecutable = if ($null -ne $coreHost) { $coreHost.Source } else { Join-Path $PSHOME 'powershell.exe' }
    $firstWorker = $null
    $secondWorker = $null
    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $workerExecutable
        $startInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" "{1}" "{2}" first' -f $workerScriptPath, $workerRoot, $mutexTestName
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $firstWorker = [System.Diagnostics.Process]::Start($startInfo)

function Wait-ForMutationLockTestFile {
            param([string]$Path, [int]$TimeoutMilliseconds = 30000)
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while (-not [System.IO.File]::Exists($Path) -and $watch.ElapsedMilliseconds -lt $TimeoutMilliseconds) {
                [System.Threading.Thread]::Sleep(20)
            }
            return [System.IO.File]::Exists($Path)
        }

        $firstRegistryMarker = Join-Path $workerRoot 'first-registry-capture'
        $firstReachedRegistry = Wait-ForMutationLockTestFile -Path $firstRegistryMarker
        if (-not $firstReachedRegistry) {
            $firstExited = $firstWorker.WaitForExit(1000)
            if ($firstExited) {
                $firstWorkerOutput = $firstWorker.StandardOutput.ReadToEnd()
                $firstWorkerError = $firstWorker.StandardError.ReadToEnd()
                [Console]::Error.WriteLine("First barrier process: exited=true code=$($firstWorker.ExitCode) stdout=[$firstWorkerOutput] stderr=[$firstWorkerError]")
            }
            else {
                [Console]::Error.WriteLine('First barrier process remained active without reaching the registry marker.')
            }
        }
        Assert-True $firstReachedRegistry 'the first mocked run reaches the registry-capture barrier while holding the mutex'
        $secondStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $secondStartInfo.FileName = $workerExecutable
        $secondStartInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" "{1}" "{2}" second' -f $workerScriptPath, $workerRoot, $mutexTestName
        $secondStartInfo.UseShellExecute = $false
        $secondStartInfo.CreateNoWindow = $true
        $secondStartInfo.RedirectStandardOutput = $true
        $secondStartInfo.RedirectStandardError = $true
        $secondWorker = [System.Diagnostics.Process]::Start($secondStartInfo)
        Assert-True (Wait-ForMutationLockTestFile -Path (Join-Path $workerRoot 'second-attempted')) 'the second mocked entry point attempts to acquire the mutex while the first owns it'
        Assert-Equal $mutexTestName ([System.IO.File]::ReadAllText((Join-Path $workerRoot 'second-lock-name'))) 'the second process receives the exact shared mutex name'
        Assert-Equal $false $firstWorker.HasExited 'the first process still owns the mutex when the second arrives'
        Assert-True (Wait-ForMutationLockTestFile -Path (Join-Path $workerRoot 'second-mutex-opened')) 'the second process opens the shared mutex handle before the first is released'

        $secondRegistryMarker = Join-Path $workerRoot 'second-registry-capture'
        $secondProcessMarker = Join-Path $workerRoot 'second-ooshutup-start'
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($watch.ElapsedMilliseconds -lt 250) { [System.Threading.Thread]::Sleep(20) }
        Assert-Equal $false ([System.IO.File]::Exists($secondRegistryMarker)) 'the second run performs no registry call before acquiring the mutex'
        Assert-Equal $false ([System.IO.File]::Exists($secondProcessMarker)) 'the second run performs no O&O start before acquiring the mutex'
        Assert-Equal $false ([System.IO.File]::Exists((Join-Path $workerRoot 'first-ooshutup-start'))) 'the first run remains at the registry barrier during the overlap check'

        [System.IO.File]::WriteAllText((Join-Path $workerRoot 'release-first'), '')
        Assert-True $firstWorker.WaitForExit(15000) 'the first barrier process completes after its release signal'
        Assert-True $secondWorker.WaitForExit(15000) 'the second barrier process completes after the first releases the mutex'
        $firstWorkerOutput = $firstWorker.StandardOutput.ReadToEnd()
        $firstWorkerError = $firstWorker.StandardError.ReadToEnd()
        $secondWorkerOutput = $secondWorker.StandardOutput.ReadToEnd()
        $secondWorkerError = $secondWorker.StandardError.ReadToEnd()
        Assert-Equal 0 $firstWorker.ExitCode "the first barrier process exits successfully ($firstWorkerError)"
        Assert-Equal 0 $secondWorker.ExitCode "the second barrier process exits successfully ($secondWorkerError)"
        Assert-True ([System.IO.File]::Exists($secondRegistryMarker)) 'the second registry capture starts after acquiring the released mutex'
        Assert-True ([System.IO.File]::Exists($secondProcessMarker)) 'the second O&O start occurs after acquiring the mutex'
        $barrierEvents = [System.IO.File]::ReadAllText((Join-Path $workerRoot 'events.txt')).Trim() -split "`r?`n"
        Assert-Equal 'first-owner-enter,first-registry-capture,first-ooshutup-start,first-owner-exit,second-owner-enter,second-registry-capture,second-ooshutup-start,second-owner-exit' ([string]::Join(',', [string[]]$barrierEvents)) 'the two mocked runs own the lock and perform registry/O&O work in one ordered sequence'
        Assert-Equal '' $firstWorkerOutput 'the first barrier process does not emit unexpected output'
        Assert-Equal '' $secondWorkerOutput 'the second barrier process does not emit unexpected output'
    }
    finally {
        if ($null -ne $firstWorker -and -not $firstWorker.HasExited) {
            [System.IO.File]::WriteAllText((Join-Path $workerRoot 'release-first'), '')
            if (-not $firstWorker.WaitForExit(5000)) { $firstWorker.Kill(); $firstWorker.WaitForExit() }
        }
        if ($null -ne $secondWorker -and -not $secondWorker.HasExited) {
            if (-not $secondWorker.WaitForExit(5000)) { $secondWorker.Kill(); $secondWorker.WaitForExit() }
        }
        if ($null -ne $firstWorker) { $firstWorker.Dispose() }
        if ($null -ne $secondWorker) { $secondWorker.Dispose() }
        if ([System.IO.Directory]::Exists($workerRoot)) {
            Get-ChildItem -LiteralPath $workerRoot -File -Force | Remove-Item -Force
            Remove-Item -LiteralPath $workerRoot -Force
        }
    }

    $timeoutTestName = "$script:OOShutUpMutationMutexName.Timeout"
    $timeoutAnchor = New-OOShutUpMutationMutex -Name $timeoutTestName
    $abandonedTestName = "$script:OOShutUpMutationMutexName.Abandoned"
    $abandonedAnchor = $null
    $failureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oosu-mutation-lock-failures-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($failureRoot)
    $failureWorkerPath = Join-Path $failureRoot 'mutex-failure-worker.ps1'
    $failureWorkerBody = @'
$workDir = [string]$args[0]
$mutexName = [string]$args[1]
$role = [string]$args[2]
if ($role -eq 'probe') {
    $result = Invoke-WithOOShutUpMutationLock -MutexName $mutexName -TimeoutMilliseconds 250 -Operation {
        [System.IO.File]::WriteAllText((Join-Path $workDir 'probe-operation'), '')
    }
    [System.IO.File]::WriteAllText((Join-Path $workDir 'probe-result'), [string]$result.FailureType)
    if ($result.Succeeded) { exit 3 }
    exit 0
}
$mutex = New-OOShutUpMutationMutex -Name $mutexName
if (-not $mutex.WaitOne(10000)) { exit 2 }
[System.IO.File]::WriteAllText((Join-Path $workDir "$role-owned"), '')
if ($role -eq 'holder') {
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    $releaseMarker = Join-Path $workDir 'release-holder'
    while (-not [System.IO.File]::Exists($releaseMarker) -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 20
    }
    if (-not [System.IO.File]::Exists($releaseMarker)) { exit 4 }
    Release-OOShutUpMutationMutex -Mutex $mutex
    $mutex.Dispose()
}
# The abandoner exits here without releasing the owned mutex.
exit 0
'@
    [System.IO.File]::WriteAllText($failureWorkerPath, $mutexHelperText + "`r`n" + $failureWorkerBody, [System.Text.Encoding]::UTF8)
    $timeoutState = [pscustomobject]@{ OperationCalled = $false }
    $timeoutHolder = $null
    $abandonerProcess = $null
    $abandonedProbe = $null
    try {
        $holderInfo = New-Object System.Diagnostics.ProcessStartInfo
        $holderInfo.FileName = $workerExecutable
        $holderInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" "{1}" "{2}" holder' -f $failureWorkerPath, $failureRoot, $timeoutTestName
        $holderInfo.UseShellExecute = $false
        $holderInfo.CreateNoWindow = $true
        $holderInfo.RedirectStandardOutput = $true
        $holderInfo.RedirectStandardError = $true
        $timeoutHolder = [System.Diagnostics.Process]::Start($holderInfo)
        Assert-True (Wait-ForMutationLockTestFile -Path (Join-Path $failureRoot 'holder-owned')) 'the timeout fixture owns the mutex before the contender starts'
        $timeoutResult = Invoke-WithOOShutUpMutationLock -MutexName $timeoutTestName -TimeoutMilliseconds 100 -Operation {
            $timeoutState.OperationCalled = $true
        }
        Assert-Equal $false $timeoutResult.Succeeded 'mutex timeout returns a non-success result'
        Assert-Equal 'Timeout' $timeoutResult.FailureType 'mutex timeout is reported as a bounded wait failure'
        Assert-Equal $false $timeoutResult.OperationStarted 'mutex timeout stops before the protected operation'
        Assert-Equal $false $timeoutState.OperationCalled 'timeout performs no registry or O&O mock call'
        [System.IO.File]::WriteAllText((Join-Path $failureRoot 'release-holder'), '')
        Assert-True $timeoutHolder.WaitForExit(15000) 'the timeout fixture owner releases its lock'
        Assert-Equal 0 $timeoutHolder.ExitCode 'the timeout fixture owner exits cleanly'
        $timeoutAnchor.Dispose()
        $timeoutAnchor = $null

        $abandonedAnchor = New-OOShutUpMutationMutex -Name $abandonedTestName
        $abandonerInfo = New-Object System.Diagnostics.ProcessStartInfo
        $abandonerInfo.FileName = $workerExecutable
        $abandonerInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" "{1}" "{2}" abandoner' -f $failureWorkerPath, $failureRoot, $abandonedTestName
        $abandonerInfo.UseShellExecute = $false
        $abandonerInfo.CreateNoWindow = $true
        $abandonerInfo.RedirectStandardOutput = $true
        $abandonerInfo.RedirectStandardError = $true
        $abandonerProcess = [System.Diagnostics.Process]::Start($abandonerInfo)
        Assert-True (Wait-ForMutationLockTestFile -Path (Join-Path $failureRoot 'abandoner-owned')) 'the abandoned-lock fixture acquired the named mutex'
        Assert-True $abandonerProcess.WaitForExit(15000) 'the fixture process exited while holding the mutex'
        Assert-Equal 0 $abandonerProcess.ExitCode 'the abandoned fixture exits without releasing the mutex'

        $abandonedOperationState = [pscustomobject]@{ Called = $false }
        $abandonedResult = Invoke-WithOOShutUpMutationLock -MutexName $abandonedTestName -TimeoutMilliseconds 1000 -Operation {
            $abandonedOperationState.Called = $true
        }
        Assert-Equal $false $abandonedResult.Succeeded 'abandoned mutex ownership returns a non-success result'
        Assert-Equal 'Abandoned' $abandonedResult.FailureType 'abandoned mutex ownership is identified explicitly'
        Assert-Equal $false $abandonedResult.OperationStarted 'abandoned mutex ownership stops before the protected operation'
        Assert-Equal $false $abandonedOperationState.Called 'abandoned ownership performs no registry or O&O mock call'

        $probeInfo = New-Object System.Diagnostics.ProcessStartInfo
        $probeInfo.FileName = $workerExecutable
        $probeInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" "{1}" "{2}" probe' -f $failureWorkerPath, $failureRoot, $abandonedTestName
        $probeInfo.UseShellExecute = $false
        $probeInfo.CreateNoWindow = $true
        $probeInfo.RedirectStandardOutput = $true
        $probeInfo.RedirectStandardError = $true
        $abandonedProbe = [System.Diagnostics.Process]::Start($probeInfo)
        Assert-True $abandonedProbe.WaitForExit(15000) 'a queued contender returns while the abandoned owner still holds the lock'
        Assert-Equal 0 $abandonedProbe.ExitCode 'the queued contender reports timeout without worker errors'
        Assert-Equal 'Timeout' ([System.IO.File]::ReadAllText((Join-Path $failureRoot 'probe-result'))) 'a queued contender does not proceed after abandonment was detected'
        Assert-Equal $false ([System.IO.File]::Exists((Join-Path $failureRoot 'probe-operation'))) 'no queued registry or O&O mock call starts after abandoned ownership'
        $abandonedAnchor.Dispose()
        $abandonedAnchor = $null
    }
    finally {
        if ($null -ne $timeoutHolder -and -not $timeoutHolder.HasExited) {
            [System.IO.File]::WriteAllText((Join-Path $failureRoot 'release-holder'), '')
            if (-not $timeoutHolder.WaitForExit(5000)) { $timeoutHolder.Kill(); $timeoutHolder.WaitForExit() }
        }
        foreach ($child in @($timeoutHolder, $abandonerProcess, $abandonedProbe)) {
            if ($null -ne $child) { $child.Dispose() }
        }
        if ($null -ne $timeoutAnchor) { $timeoutAnchor.Dispose() }
        if ($null -ne $abandonedAnchor) { $abandonedAnchor.Dispose() }
        if ([System.IO.Directory]::Exists($failureRoot)) {
            Get-ChildItem -LiteralPath $failureRoot -File -Force | Remove-Item -Force
            Remove-Item -LiteralPath $failureRoot -Force
        }
    }

    $releaseTestName = "$script:OOShutUpMutationMutexName.ReleaseFailure"
    $releaseFunction = (Get-Command Release-OOShutUpMutationMutex).ScriptBlock
    $releaseFailureHandler = {
        param([System.Threading.Mutex]$Mutex)
        $Mutex.ReleaseMutex()
        throw 'Mocked mutex release verification failure.'
    }
    Set-Item -Path Function:Release-OOShutUpMutationMutex -Value $releaseFailureHandler
    try {
        $releaseResult = Invoke-WithOOShutUpMutationLock -MutexName $releaseTestName -TimeoutMilliseconds 1000 -Operation {
            return 'mocked-state-change'
        }
    }
    finally {
        Set-Item -Path Function:Release-OOShutUpMutationMutex -Value $releaseFunction
    }
    Assert-Equal $false $releaseResult.Succeeded 'a mutex release verification failure cannot report overall success'
    Assert-Equal 'ReleaseFailed' $releaseResult.FailureType 'mutex release failure has an explicit failure status'
    Assert-Equal $true $releaseResult.OperationStarted 'release failure is reported after the protected operation may have run'
    Assert-Equal 'mocked-state-change' $releaseResult.Value 'release failure preserves the protected operation outcome for diagnosis'
    Assert-True ($releaseResult.FailureMessage.Contains('inspect the System Restore frequency value and O&O settings')) 'release failure gives actionable recovery guidance'

    [Console]::Out.WriteLine("Passed $script:AssertionCount restore-point readiness assertions in $($PSVersionTable.PSVersion).")
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    $env:SystemRoot = $script:OriginalSystemRoot
}
