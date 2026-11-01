[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:LogMessages = New-Object System.Collections.ArrayList
$script:RestoreScenario = 'success'
$script:LastRestorePointDescription = $null
$script:RestorePoints = @()
$script:ProcessBoundaryCalls = 0
$script:CurrentRestoreJob = $null
$script:OOShutUpRestorePointFailureMessage = 'A restore point could not be verified. No privacy settings were applied; free space or enable restore protection, then retry.'
$script:OOShutUpRestorePointOverrideMessage = 'A restore point could not be verified. System Restore recovery is unavailable; continuing only because -AllowWithoutRestorePoint was supplied. The approved O&O profile export remains separate.'
$script:OOShutUpApprovalFailureMessage = 'The approved profile did not match the test approval.'
$script:OOShutUpApprovalConfirmed = $true
$script:OOShutUpApprovedDigest = 'profile-digest'
$script:OOShutUpApprovedManifestDigest = 'manifest-digest'
$script:OOShutUpLastRunChanged = $false

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

function Enable-ComputerRestore {
    [CmdletBinding()]
    param([string]$Drive)
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
        return @{ Success = $true; Message = 'Mock provider returned success.' }
    }

    $message = switch -Exact ($script:RestoreScenario) {
        'protection-disabled' { 'System Restore protection is disabled.' }
        'no-space' { 'There is not enough disk space.' }
        'provider-unavailable' { 'The System Restore provider is unavailable.' }
        'provider-exception' { 'The provider threw an exception.' }
        default { 'Mock provider rejected the request.' }
    }
    return @{ Success = $false; Message = $message }
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
        return @([pscustomobject]@{
            Description = $script:LastRestorePointDescription
            SequenceNumber = 42
        })
    }

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

    return [pscustomobject]@{
        Sha256 = 'profile-digest'
        ManifestSha256 = 'manifest-digest'
        Version = 'test-1.0'
        SettingCount = 1
        Path = 'C:\fixtures\ooshutup10.cfg'
    }
}

function Test-OOShutUpPayload {
    [CmdletBinding()]
    param([string]$Path, [switch]$Quiet)

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

    $script:ProcessBoundaryCalls++
    return [pscustomobject]@{ ExitCode = 30 }
}

try {
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

    $requiredFunctions = @(
        'New-SystemRestorePoint',
        'Invoke-OOShutUpConfiguration',
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

    $gateFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpConfigurationWithRestorePointGate' })[0].Extent.Text)
    $pointCreationPosition = $gateFunctionSource.IndexOf('New-SystemRestorePoint', [System.StringComparison]::Ordinal)
    $failureGatePosition = $gateFunctionSource.IndexOf('if (-not $restorePointVerified -and -not $AllowWithoutRestorePoint)', [System.StringComparison]::Ordinal)
    $configurationPosition = $gateFunctionSource.IndexOf('Invoke-OOShutUpConfiguration -ExecutablePath', [System.StringComparison]::Ordinal)
    Assert-True ($pointCreationPosition -ge 0 -and $failureGatePosition -gt $pointCreationPosition -and $configurationPosition -gt $failureGatePosition) 'restore verification and the default failure gate precede O&O configuration'

    $mainSource = $source.Substring($source.IndexOf('#region Main Script', [System.StringComparison]::Ordinal))
    $mainGatePosition = $mainSource.IndexOf('if (-not $configurationResult.ConfigurationStarted)', [System.StringComparison]::Ordinal)
    $mainFailureExitPosition = $mainSource.IndexOf('exit 1', $mainGatePosition, [System.StringComparison]::Ordinal)
    Assert-True ($mainGatePosition -ge 0 -and $mainFailureExitPosition -gt $mainGatePosition) 'a blocked restore gate exits the installer with a nonzero status'

    $restoreCases = @(
        @{ Name = 'protection disabled'; Scenario = 'protection-disabled'; ExpectedCreated = $false },
        @{ Name = 'insufficient space'; Scenario = 'no-space'; ExpectedCreated = $false },
        @{ Name = 'provider unavailable'; Scenario = 'provider-unavailable'; ExpectedCreated = $false },
        @{ Name = 'provider exception'; Scenario = 'provider-exception'; ExpectedCreated = $false },
        @{ Name = 'job startup exception'; Scenario = 'job-start-exception'; ExpectedCreated = $false },
        @{ Name = 'receive exception'; Scenario = 'receive-job-exception'; ExpectedCreated = $false },
        @{ Name = 'provider timeout'; Scenario = 'timeout'; ExpectedCreated = $false },
        @{ Name = 'false-success response'; Scenario = 'false-success'; ExpectedCreated = $false },
        @{ Name = 'verified creation'; Scenario = 'success'; ExpectedCreated = $true }
    )

    foreach ($case in $restoreCases) {
        $script:RestoreScenario = $case.Scenario
        $script:RestorePoints = @()
        $script:LogMessages.Clear()
        $created = New-SystemRestorePoint
        Assert-Equal $case.ExpectedCreated $created "$($case.Name) returns the expected verified status"
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
            Assert-True ($messages.Contains('Windows did not list the requested restore point')) 'a false-success response is rejected when Windows has no matching point'
        }
    }

    $callCountBeforeFailure = $script:ProcessBoundaryCalls
    $blockedScenarios = @('protection-disabled', 'no-space', 'provider-unavailable', 'provider-exception', 'job-start-exception', 'receive-job-exception', 'enumeration-exception', 'timeout', 'false-success')
    foreach ($scenario in $blockedScenarios) {
        $script:RestoreScenario = $scenario
        $script:RestorePoints = @()
        $script:LogMessages.Clear()
        $blockedResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
        Assert-Equal $false $blockedResult.ConfigurationStarted "$scenario stops before configuration starts"
        Assert-Equal $false $blockedResult.ConfigurationSucceeded "$scenario does not report configuration success"
        Assert-Equal $callCountBeforeFailure $script:ProcessBoundaryCalls "$scenario never reaches Start-Process"
        $blockedMessages = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
        Assert-True ($blockedMessages.Contains($script:OOShutUpRestorePointFailureMessage)) "$scenario gives the exact actionable retry status"
    }

    $script:RestoreScenario = 'success'
    $script:RestorePoints = @()
    $script:LogMessages.Clear()
    $successResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe'
    Assert-Equal $true $successResult.RestorePointVerified 'verified point status reaches the result'
    Assert-Equal $true $successResult.ConfigurationStarted 'verified point permits configuration'
    Assert-Equal $true $successResult.ConfigurationSucceeded 'mocked O&O reports successful application'
    Assert-Equal ($callCountBeforeFailure + 1) $script:ProcessBoundaryCalls 'verified point reaches Start-Process exactly once'

    $script:RestoreScenario = 'provider-unavailable'
    $script:RestorePoints = @()
    $script:LogMessages.Clear()
    $overrideResult = Invoke-OOShutUpConfigurationWithRestorePointGate -ExecutablePath 'C:\fixtures\OOSU10.exe' -AllowWithoutRestorePoint
    Assert-Equal $false $overrideResult.RestorePointVerified 'override does not represent a point as verified'
    Assert-Equal $true $overrideResult.OverrideUsed 'the explicit switch is recorded in the result'
    Assert-Equal $true $overrideResult.ConfigurationStarted 'the explicit override permits configuration'
    Assert-Equal ($callCountBeforeFailure + 2) $script:ProcessBoundaryCalls 'only the explicit override reaches Start-Process'
    $overrideMessages = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($overrideMessages.Contains($script:OOShutUpRestorePointOverrideMessage)) 'override reports that System Restore recovery is unavailable'
    Assert-True ($overrideMessages -notmatch 'A system restore point was verified|created and verified successfully') 'override never claims a restore point'

    [Console]::Out.WriteLine("Passed $script:AssertionCount restore-point readiness assertions in $($PSVersionTable.PSVersion).")
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
