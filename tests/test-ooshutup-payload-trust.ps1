[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:LogMessages = New-Object System.Collections.ArrayList
$script:PromptMessages = New-Object System.Collections.ArrayList
$script:StartedPaths = New-Object System.Collections.ArrayList
$script:StartedArguments = New-Object System.Collections.ArrayList
$script:StartedWorkingDirectories = New-Object System.Collections.ArrayList
$script:StartedWaitFlags = New-Object System.Collections.ArrayList
$script:StartedProcessObjects = New-Object System.Collections.ArrayList
$script:OpenedProfilePaths = New-Object System.Collections.ArrayList
$script:OpenedProfileBytes = New-Object System.Collections.ArrayList
$script:OpenedManifestBytes = New-Object System.Collections.ArrayList
$script:StartProcessCallCount = 0
$script:MockStartProcessFailure = $false
$script:MockStartProcessReturnsNull = $false
$script:SwapProfileBeforeMockOpen = $false
$script:MockReplacementProfileContent = ''
$script:MockReplacementManifestContent = ''
$script:SnapshotAclDenied = $false
$script:MockSnapshotAclDeniedTarget = $null
$script:SnapshotFailureProfile = $null
$script:DownloadScenario = 'valid'
$script:DownloadCount = 0
$script:SignatureCallCount = 0
$script:ReleaseSignatureByPath = @{}
$script:MockTrustedReleaseSignerThumbprint = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
$script:MockEnforcedExecutionPolicy = 'AllSigned'
$script:MockUserExecutionPolicy = 'Undefined'
$script:MockMachineTrustedPublisher = $true
$script:OOShutUpReleaseManifestFileName = 'Install-OOShutUp10.release.psd1'
$script:MockExitCode = 0
$script:MockChildWaitScenario = 'fast-exit'
$script:MockChildWaitTimeoutMilliseconds = 0
$script:PromptDecision = 'YES'
$script:FixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oosu10-trust-tests-' + [guid]::NewGuid().ToString('N'))
$script:TaskRegistry = @{}
$script:TaskSecurityDescriptorByIdentity = @{}
$script:TaskRegistrationCalls = New-Object System.Collections.ArrayList
$script:TaskUnregistrationCalls = New-Object System.Collections.ArrayList
$script:TaskLookupCalls = New-Object System.Collections.ArrayList
$script:TaskSecurityDescriptorReadCalls = New-Object System.Collections.ArrayList
$script:TaskSecurityDescriptorSetCalls = New-Object System.Collections.ArrayList
$script:TaskEnabledCalls = New-Object System.Collections.ArrayList
$script:MockTaskSecurityDescriptorScenario = 'Normal'
$script:MockTaskSecurityDescriptorReplacementInjected = $false
$script:MockTaskSecurityDescriptorSetFailureTaskName = $null
$script:TaskLookupFailure = $null
$script:TaskRegistrationFailure = $null
$script:TaskVerificationMissPending = $false
$script:FailTaskRestoreXml = $null
$script:FailTaskVerification = $false
$script:FailStageVerification = $false
$script:FailStageQueuePolicyVerification = $false
$script:FailStageExecutionTimeLimitVerification = $false
$script:FailTaskRestore = $false
$script:MockInstallerTaskScenario = 'Success'
$script:OOShutUpPartialSetupMessage = 'The privacy settings may be applied, but the update task was not verified. Setup is partial; repair the task before relying on automatic reapplication.'
$script:OOShutUpTaskSecurityDescriptorSddl = 'O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)'
$script:OOShutUpTaskSecurityDescriptorFailureMessage = 'The task security descriptor could not be verified. No new task was enabled; inspect task permissions before retrying.'
$script:OOShutUpTaskRecoveryMessage = 'Recovery: resolve the logged Task Scheduler error, open PowerShell as Administrator, and rerun this installer with the approved profile (or use -NonInteractive -ApprovedProfileSha256 <profile SHA-256> -ApprovedProfileManifestSha256 <manifest SHA-256>). Verify \myTech.Today\Install-OOShutUp10 in Task Scheduler before relying on automatic reapplication. The retry inspects and replaces only this verified task; same-named tasks in other folders are left unchanged.'
$script:OOShutUpUpdateEventSubscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''Microsoft-Windows-WindowsUpdateClient''] and (EventID=19)]]</Select></Query></QueryList>'
$script:OOShutUpUpdateEventDelay = 'PT2M'
$script:OOShutUpUpdateMultipleInstancesPolicy = 'Queue'
$script:OOShutUpChildTimeoutSeconds = 900
$script:OOShutUpTaskExecutionTimeLimit = 'PT30M'
$script:OOShutUpRestorePointFailureMessage = 'A restore point could not be verified. No privacy settings were applied; free space or enable restore protection, then retry.'
$script:MockAclByPath = @{}
$script:MockAclFailurePath = $null
$script:MockWindowsPowerShellSignatureMode = 'Valid'
$script:MockWindowsPowerShellDirectory = $null
$script:MockPowerShellContextIs64BitOperatingSystem = $true
$script:MockPowerShellContextIs64BitProcess = $true
$script:MockPowerShellResolverMode = 'Valid'
$script:MockPowerShellResolverCallCount = 0
$script:MockPowerShellResolverChangedPath = $null
$script:MockLastPowerShellInspectionPath = $null
$script:OOShutUpMutationMutexName = "Global\myTech.Today.OOShutUp.PayloadTrustTest.$PID"
$script:OOShutUpMutationLockTimeoutMilliseconds = 600000

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

function Assert-InvalidProfileStopsBeforeTaskMutation {
    param([string]$Name)

    $lookupsBefore = $script:ScheduledTaskLookupCount
    $registrationsBefore = $script:TaskRegistrationCalls.Count
    Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath (Join-Path $script:FixtureRoot 'unused-invalid-profile.ps1') -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 ('0' * 64)) "$Name profile pair blocks task setup"
    Assert-Equal $lookupsBefore $script:ScheduledTaskLookupCount "$Name profile pair blocks before task lookup"
    Assert-Equal $registrationsBefore $script:TaskRegistrationCalls.Count "$Name profile pair blocks before task registration"
}

function Test-WindowsUpdateClientProviderMetadata {
    try {
        $events = @((Get-WinEvent -ListProvider 'Microsoft-Windows-WindowsUpdateClient' -ErrorAction Stop).Events)
    }
    catch {
        Write-Output ("Skipped host Windows Update provider metadata check: {0}" -f $_.Exception.Message)
        return
    }

    $expectedLabels = @{
        19 = 'Installation Successful'
        20 = 'Installation Failure'
        43 = 'Installation Started'
    }
    foreach ($eventId in @(19, 20, 43)) {
        $matches = @($events | Where-Object { [int]$_.Id -eq $eventId -and -not [string]::IsNullOrWhiteSpace([string]$_.Description) })
        Assert-True ($matches.Count -gt 0) "the host Windows Update provider exposes event $eventId metadata"
        $description = ([string]$matches[0].Description).Replace([char]0, [char]32).Trim()
        if ([System.Globalization.CultureInfo]::CurrentUICulture.Name -like 'en*') {
            Assert-True ($description.Contains($expectedLabels[$eventId])) "the host metadata labels event $eventId as $($expectedLabels[$eventId])"
        }
        Write-Output ("Host WindowsUpdateClient event {0}: {1}" -f $eventId, $description.Trim())
    }
}

function Write-TestPayload {
    param(
        [string]$Path,
        [string]$Content
    )

    [System.IO.File]::WriteAllText($Path, $Content)
}

function New-MockSigner {
    param(
        [string]$Name,
        [string]$Thumbprint = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    )

    $certificate = [pscustomobject]@{ SimpleName = $Name; Thumbprint = $Thumbprint }
    $certificate | Add-Member -MemberType ScriptMethod -Name GetNameInfo -Value { return $this.SimpleName } -Force
    return $certificate
}

function Get-ExecutionPolicy {
    [CmdletBinding()]
    param([switch]$List)

    if (-not $List) {
        return $script:MockEnforcedExecutionPolicy
    }

    return @(
        [pscustomobject]@{ Scope = 'MachinePolicy'; ExecutionPolicy = $script:MockEnforcedExecutionPolicy }
        [pscustomobject]@{ Scope = 'UserPolicy'; ExecutionPolicy = $script:MockUserExecutionPolicy }
    )
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    [void]$script:LogMessages.Add("[$Level] $Message")
}

function Write-Host {
    param(
        [Parameter(Position = 0)]
        [object]$Object,

        [System.ConsoleColor]$ForegroundColor
    )

    [void]$script:PromptMessages.Add([string]$Object)
}

function Read-Host {
    param([string]$Prompt)

    [void]$script:PromptMessages.Add($Prompt)
    return $script:PromptDecision
}

function Get-AuthenticodeSignature {
    [CmdletBinding()]
    param([string]$FilePath)

    $script:SignatureCallCount++
    $canonicalPath = [System.IO.Path]::GetFullPath($FilePath)
    if ($script:ReleaseSignatureByPath.ContainsKey($canonicalPath)) {
        return $script:ReleaseSignatureByPath[$canonicalPath]
    }
    if ([string]$FilePath -match '(?i)\\WindowsPowerShell\\v1\.0\\powershell\.exe$') {
        if ($script:MockWindowsPowerShellSignatureMode -eq 'WrongPublisher') {
            return [pscustomobject]@{ Status = 'Valid'; SignerCertificate = (New-MockSigner 'Other Publisher') }
        }
        if ($script:MockWindowsPowerShellSignatureMode -eq 'Unsigned') {
            return [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null }
        }
        return [pscustomobject]@{ Status = 'Valid'; SignerCertificate = (New-MockSigner 'Microsoft Windows') }
    }

    $content = [System.IO.File]::ReadAllText($FilePath)

    switch -Exact ($content) {
        'UNSIGNED' {
            return [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null }
        }
        'WRONG_PUBLISHER' {
            return [pscustomobject]@{ Status = 'Valid'; SignerCertificate = (New-MockSigner 'Other Publisher') }
        }
        { $_ -in @('TAMPERED', 'CORRUPT', 'PARTIAL') } {
            return [pscustomobject]@{ Status = 'HashMismatch'; SignerCertificate = (New-MockSigner 'O&O Software GmbH') }
        }
        default {
            return [pscustomobject]@{ Status = 'Valid'; SignerCertificate = (New-MockSigner 'O&O Software GmbH') }
        }
    }
}

function New-MockOOSChildProcess {
    $process = [pscustomobject]@{
        Id = 4000 + $script:StartProcessCallCount
        ExitCode = $script:MockExitCode
        HasExited = $false
        WaitCalls = 0
        UntimedWaitCalls = 0
        KillCalls = 0
    }
    $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value {
        param([int]$Milliseconds = -1)

        $this.WaitCalls++
        if ($Milliseconds -lt 0) {
            $this.UntimedWaitCalls++
            return [bool]$this.HasExited
        }

        $script:MockChildWaitTimeoutMilliseconds = $Milliseconds
        switch -Exact ($script:MockChildWaitScenario) {
            'fast-exit' {
                $this.HasExited = $true
                return $true
            }
            'timeout-boundary-race' {
                $this.HasExited = $true
                return $false
            }
            'still-running' {
                return $false
            }
            'wait-error' {
                throw 'Mock process wait could not be verified.'
            }
            default {
                throw "Unexpected mocked O&O child wait scenario: $script:MockChildWaitScenario"
            }
        }
    } -Force
    $process | Add-Member -MemberType ScriptMethod -Name Kill -Value {
        $this.KillCalls++
        $this.HasExited = $true
    } -Force
    return $process
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

    $script:StartProcessCallCount++
    [void]$script:StartedPaths.Add($FilePath)
    [void]$script:StartedArguments.Add($ArgumentList)
    [void]$script:StartedWorkingDirectories.Add($WorkingDirectory)
    [void]$script:StartedWaitFlags.Add([bool]$Wait)
    if ($script:MockStartProcessFailure) {
        throw 'Mocked Start-Process outcome could not be verified.'
    }
    if ($script:MockStartProcessReturnsNull) {
        return $null
    }
    $profileArgument = [System.Text.RegularExpressions.Regex]::Match($ArgumentList, '^/quiet /nosrp "([^"]+)"$')
    if ($profileArgument.Success) {
        if ($script:SwapProfileBeforeMockOpen) {
            [System.IO.File]::WriteAllText($script:OOShutUpProfilePath, $script:MockReplacementProfileContent, [System.Text.Encoding]::UTF8)
            [System.IO.File]::WriteAllText($script:OOShutUpProfileManifestPath, $script:MockReplacementManifestContent, [System.Text.Encoding]::UTF8)
        }
        $openedProfilePath = [System.IO.Path]::GetFullPath($profileArgument.Groups[1].Value)
        [void]$script:OpenedProfilePaths.Add($openedProfilePath)
        [void]$script:OpenedProfileBytes.Add([System.IO.File]::ReadAllBytes($openedProfilePath))
        $openedManifestPath = Join-Path (Split-Path -Parent $openedProfilePath) 'ooshutup10.approval.json'
        [void]$script:OpenedManifestBytes.Add([System.IO.File]::ReadAllBytes($openedManifestPath))
    }
    $process = New-MockOOSChildProcess
    [void]$script:StartedProcessObjects.Add($process)
    return $process
}

function Write-ApprovedProfileFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $true)]
        [string[]]$AllowedSettingIds,

        [Parameter(Mandatory = $false)]
        [string[]]$CurrentUserSettingIds = @(),

        [Parameter(Mandatory = $false)]
        [string]$OOShutUpMinimumVersion = '3.6.1135.0',

        [Parameter(Mandatory = $false)]
        [string]$OOShutUpMaximumVersion = '3.6.1135.0'
    )

    [System.IO.Directory]::CreateDirectory($script:OOShutUpProfileDirectory) | Out-Null
    [System.IO.File]::WriteAllText($script:OOShutUpProfilePath, $Content, [System.Text.Encoding]::UTF8)
    $hash = (Get-FileHash -LiteralPath $script:OOShutUpProfilePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $settingDetails = @(
        foreach ($settingId in $AllowedSettingIds) {
            $entry = [System.Text.RegularExpressions.Regex]::Match(
                $Content,
                ('(?m)^\s*{0}[\t;]([+-])' -f [System.Text.RegularExpressions.Regex]::Escape($settingId))
            )
            $state = if ($entry.Success) { $entry.Groups[1].Value } else { '+' }
            [pscustomobject]@{
                id = $settingId
                state = $state
                category = 'Test category'
                impact = 'Test fixture impact description.'
                scope = if ($CurrentUserSettingIds -ccontains $settingId) { 'current-user' } else { 'machine' }
                scopeEvidence = 'Mock fixture scope evidence; no registry was read or changed.'
            }
        }
    )
    $manifest = [pscustomobject]@{
        schemaVersion = 4
        ooshutupVersionRange = [pscustomobject]@{
            minimum = $OOShutUpMinimumVersion
            maximum = $OOShutUpMaximumVersion
        }
        profileVersion = 'test-1.0'
        sha256 = $hash
        settingIds = $AllowedSettingIds
        settingDetails = $settingDetails
    }
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($manifest | ConvertTo-Json -Depth 3),
        [System.Text.Encoding]::UTF8
    )
}

function New-MockAclRule {
    param(
        [string]$Sid,
        [System.Security.AccessControl.FileSystemRights]$Rights,
        [System.Security.AccessControl.AccessControlType]$Type = [System.Security.AccessControl.AccessControlType]::Allow,
        [System.Security.AccessControl.PropagationFlags]$PropagationFlags = [System.Security.AccessControl.PropagationFlags]::None
    )

    return [pscustomobject]@{
        IdentityReference = [System.Security.Principal.SecurityIdentifier]::new($Sid)
        FileSystemRights = $Rights
        AccessControlType = $Type
        PropagationFlags = $PropagationFlags
    }
}

function New-MockAcl {
    param(
        [string]$OwnerSid = 'S-1-5-18',
        [object[]]$Rules = @()
    )

    return [pscustomobject]@{ Owner = $OwnerSid; Access = @($Rules) }
}

function Get-Acl {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $resolvedPath = [System.IO.Path]::GetFullPath($LiteralPath)
    if ($script:MockAclFailurePath -and [string]::Equals($resolvedPath, $script:MockAclFailurePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Mock ACL access denied.'
    }
    $denySnapshotAclTarget = $false
    if ($script:MockSnapshotAclDeniedTarget -and $script:InstallPath) {
        $snapshotPrefix = Join-Path ([System.IO.Path]::GetFullPath($script:InstallPath)) 'ProfileSnapshot-'
        if ($resolvedPath.StartsWith($snapshotPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $snapshotLeaf = [System.IO.Path]::GetFileName($resolvedPath)
            switch ([string]$script:MockSnapshotAclDeniedTarget) {
                'Directory' { $denySnapshotAclTarget = ($snapshotLeaf -match '^ProfileSnapshot-[A-Fa-f0-9]{32}$') }
                'Profile' { $denySnapshotAclTarget = [string]::Equals($snapshotLeaf, 'ooshutup10.cfg', [System.StringComparison]::OrdinalIgnoreCase) }
                'Manifest' { $denySnapshotAclTarget = [string]::Equals($snapshotLeaf, 'ooshutup10.approval.json', [System.StringComparison]::OrdinalIgnoreCase) }
            }
        }
    }
    if (($script:SnapshotAclDenied -or $denySnapshotAclTarget) -and
        $script:InstallPath -and
        $resolvedPath.StartsWith((Join-Path ([System.IO.Path]::GetFullPath($script:InstallPath)) 'ProfileSnapshot-'), [System.StringComparison]::OrdinalIgnoreCase)) {
        return New-MockAcl -Rules @(
            (New-MockAclRule -Sid 'S-1-5-18' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl)),
            (New-MockAclRule -Sid 'S-1-5-32-544' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl)),
            (New-MockAclRule -Sid 'S-1-5-11' -Rights ([System.Security.AccessControl.FileSystemRights]::CreateFiles))
        )
    }
    if ($script:MockAclByPath.ContainsKey($resolvedPath)) {
        return $script:MockAclByPath[$resolvedPath]
    }

    return $script:MockAclByPath['__default__']
}

function New-MockScheduledTask {
    param(
        [string]$TaskName,
        [string]$TaskPath,
        [string]$Xml,
        [bool]$InvalidAction = $false
    )

    [xml]$definition = $Xml
    $principal = $definition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Principals']/*[local-name()='Principal']")
    $action = $definition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']")
    $enabledNode = $definition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
    $multipleInstancesNode = $definition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='MultipleInstancesPolicy']")
    $executionTimeLimitNode = $definition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='ExecutionTimeLimit']")
    $descriptionNode = $definition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='RegistrationInfo']/*[local-name()='Description']")
    $triggerMocks = @(
        foreach ($triggerNode in $definition.SelectNodes("/*[local-name()='Task']/*[local-name()='Triggers']/*[local-name()='EventTrigger']")) {
            $triggerEnabledNode = $triggerNode.SelectSingleNode("*[local-name()='Enabled']")
            $subscriptionNode = $triggerNode.SelectSingleNode("*[local-name()='Subscription']")
            $delayNode = $triggerNode.SelectSingleNode("*[local-name()='Delay']")
            [pscustomobject]@{
                Enabled = [bool]::Parse([string]$triggerEnabledNode.InnerText)
                Subscription = [string]$subscriptionNode.InnerText
                Delay = [string]$delayNode.InnerText
            }
        }
    )
    $arguments = [string]$action.Arguments
    if ($InvalidAction) {
        $arguments = '-InvalidAction'
    }

    return [pscustomobject]@{
        TaskName = $TaskName
        TaskPath = $TaskPath
        Description = [string]$descriptionNode.InnerText
        State = if ([string]$enabledNode.InnerText -eq 'false') { 'Disabled' } else { 'Ready' }
        Principal = [pscustomobject]@{ UserId = [string]$principal.UserId; RunLevel = [string]$principal.RunLevel }
        Settings = [pscustomobject]@{
            MultipleInstances = [string]$multipleInstancesNode.InnerText
            ExecutionTimeLimit = if ($null -eq $executionTimeLimitNode) { '' } else { [string]$executionTimeLimitNode.InnerText }
        }
        Triggers = $triggerMocks
        Actions = @([pscustomobject]@{ Execute = [string]$action.Command; Arguments = $arguments; WorkingDirectory = [string]$action.WorkingDirectory })
        Definition = $Xml
    }
}

function Invoke-MockUpdateTaskSchedule {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$TaskDefinition,

        [int[]]$EventIds = @(),

        [int]$ActiveRunEventId = 0,

        [int]$ActiveRunEndSeconds = 0,

        [int]$EventOccurredAtSeconds = 10,

        [int]$ReapplyDurationSeconds = 1500
    )

    $trigger = $TaskDefinition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Triggers']/*[local-name()='EventTrigger']")
    $subscription = [string]$trigger.Subscription
    $configuredEventIds = @([regex]::Matches($subscription, 'EventID=(\d+)') | ForEach-Object { [int]$_.Groups[1].Value })
    $delaySeconds = [int][System.Xml.XmlConvert]::ToTimeSpan([string]$trigger.Delay).TotalSeconds
    $policyNode = $TaskDefinition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='MultipleInstancesPolicy']")
    $policy = [string]$policyNode.InnerText
    $runs = New-Object System.Collections.ArrayList
    $nextAvailableSeconds = $ActiveRunEndSeconds

    foreach ($eventId in $EventIds) {
        if ($configuredEventIds -notcontains $eventId) {
            continue
        }

        $startSeconds = $EventOccurredAtSeconds + $delaySeconds
        if ($startSeconds -lt $nextAvailableSeconds) {
            if ($policy -ne 'Queue') {
                continue
            }
            $startSeconds = $nextAvailableSeconds
        }

        [void]$runs.Add([pscustomobject]@{ EventId = $eventId; StartSeconds = $startSeconds; ActiveRunEventId = $ActiveRunEventId })
        $nextAvailableSeconds = $startSeconds + $ReapplyDurationSeconds
    }

    return @($runs.ToArray())
}

function Get-ScheduledTask {
    [CmdletBinding()]
    param(
        [string]$TaskName,
        [string]$TaskPath
    )

    [void]$script:TaskLookupCalls.Add([pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath })
    if ($null -ne $script:TaskLookupFailure) {
        $failureMode = $script:TaskLookupFailure
        $script:TaskLookupFailure = $null
        if ($failureMode -eq 'AccessDenied') {
            throw [System.UnauthorizedAccessException]::new('Mock Task Scheduler access denied.')
        }
        if ($failureMode -eq 'ModuleMissing') {
            throw [System.Management.Automation.CommandNotFoundException]::new('Get-ScheduledTask is unavailable because the ScheduledTasks module is not loaded.')
        }
        throw [System.InvalidOperationException]::new('Mock Task Scheduler RPC error.')
    }

    $tasks = @($script:TaskRegistry.Values | Where-Object {
        (-not $TaskName -or $_.TaskName -eq $TaskName) -and
        (-not $TaskPath -or $_.TaskPath -eq $TaskPath)
    })
    if ($script:TaskVerificationMissPending -and $TaskName -eq 'Install-OOShutUp10' -and $TaskPath -eq '\myTech.Today\') {
        $script:TaskVerificationMissPending = $false
        return @()
    }
    return $tasks
}

function Export-ScheduledTask {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath)

    $key = "$TaskPath|$TaskName"
    if (-not $script:TaskRegistry.ContainsKey($key)) {
        throw "Mock task was not found: $key"
    }
    return [string]$script:TaskRegistry[$key].Definition
}

function Register-ScheduledTask {
    [CmdletBinding()]
    param(
        [string]$TaskName,
        [string]$TaskPath,
        [string]$Xml,
        [switch]$Force
    )

    [void]$script:TaskRegistrationCalls.Add([pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath; Xml = $Xml })
    if ($null -ne $script:TaskRegistrationFailure) {
        $failureMode = $script:TaskRegistrationFailure
        $script:TaskRegistrationFailure = $null
        if ($failureMode -eq 'AccessDenied') {
            throw [System.UnauthorizedAccessException]::new('Mock Task Scheduler registration access denied.')
        }
        throw [System.InvalidOperationException]::new('Mock Task Scheduler registration failure.')
    }
    if ($script:FailTaskRestore -and $Xml -eq $script:FailTaskRestoreXml) {
        throw 'Mock previous-task restore failed.'
    }
    $invalidAction = $script:FailStageVerification -and $TaskName -like 'Install-OOShutUp10-validation-*'
    $task = New-MockScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Xml $Xml -InvalidAction:$invalidAction
    if ($script:FailStageQueuePolicyVerification -and $TaskName -like 'Install-OOShutUp10-validation-*') {
        $task.Settings.MultipleInstances = 'IgnoreNew'
    }
    if ($script:FailStageExecutionTimeLimitVerification -and $TaskName -like 'Install-OOShutUp10-validation-*') {
        $task.Settings.ExecutionTimeLimit = 'PT0S'
    }
    $script:TaskRegistry["$TaskPath|$TaskName"] = $task
    $script:TaskSecurityDescriptorByIdentity["$TaskPath|$TaskName"] = [string]$script:OOShutUpTaskSecurityDescriptorSddl
    if ($script:FailTaskVerification -and $TaskName -eq 'Install-OOShutUp10' -and $TaskPath -eq '\myTech.Today\') {
        $script:TaskVerificationMissPending = $true
        $script:FailTaskVerification = $false
    }
    return $task
}

function Unregister-ScheduledTask {
    [CmdletBinding()]
    param(
        [string]$TaskName,
        [string]$TaskPath,
        [switch]$Confirm
    )

    [void]$script:TaskUnregistrationCalls.Add([pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath })
    $script:TaskRegistry.Remove("$TaskPath|$TaskName") | Out-Null
    $script:TaskSecurityDescriptorByIdentity.Remove("$TaskPath|$TaskName") | Out-Null
}

try {
    [void](New-Item -ItemType Directory -Path $script:FixtureRoot -Force)
    Test-WindowsUpdateClientProviderMetadata
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
    $readmeSource = Get-Content -LiteralPath (Join-Path $repoRoot 'README.md') -Raw
    $installerBatchSource = Get-Content -LiteralPath (Join-Path $repoRoot 'install.bat') -Raw
    $startupIntegrityIndex = $source.IndexOf('$script:OOShutUpSourceIntegrityResult = Test-OOShutUpSourceIntegrity', [System.StringComparison]::Ordinal)
    $startupFailureCheckIndex = $source.IndexOf('if (-not $script:OOShutUpSourceIntegrityResult.Allowed)', $startupIntegrityIndex, [System.StringComparison]::Ordinal)
    $startupFailureExitIndex = if ($startupFailureCheckIndex -ge 0) { $source.IndexOf('exit 1', $startupFailureCheckIndex, [System.StringComparison]::Ordinal) } else { -1 }
    $repositoryHelperLoadIndex = $source.IndexOf('. (Join-Path $script:_RepoRoot', [System.StringComparison]::Ordinal)
    $mainScriptStart = $source.IndexOf('#region Main Script', [System.StringComparison]::Ordinal)
    Assert-True ($startupIntegrityIndex -gt 0 -and $startupFailureCheckIndex -gt $startupIntegrityIndex -and $startupFailureExitIndex -gt $startupFailureCheckIndex -and $startupFailureExitIndex -lt $repositoryHelperLoadIndex -and $mainScriptStart -gt $repositoryHelperLoadIndex) 'untrusted source exits before repository helpers, relocation, or the main installer flow'
    foreach ($protectedOperation in @(
        'function Get-ScriptRelocationPackageTarget',
        'function Set-RestorePointFrequencyRawValue',
        'function Install-OOShutUp',
        'function Invoke-OOShutUpConfigurationWithRestorePointGate',
        'function Invoke-OOShutUpTaskSetup',
        'Register-ScheduledTask -TaskName'
    )) {
        $protectedOperationIndex = $source.IndexOf($protectedOperation, [System.StringComparison]::Ordinal)
        Assert-True ($protectedOperationIndex -gt $startupFailureExitIndex) "the integrity failure exit precedes $protectedOperation"
    }
    Assert-True ($source.Contains("'SOFTWARE\Policies\myTech.Today\OOShutUp10'") -and $source.Contains("'TrustedReleaseSignerThumbprint'")) 'the production signing trust pin is read from the external machine policy key'
    Assert-True ($source.Contains('X509Store') -and $source.Contains('StoreName]::TrustedPublisher') -and $source.Contains('StoreLocation]::LocalMachine')) 'production validation requires the pinned signer in the machine Trusted Publishers store'
    Assert-True ($source.Contains('Import-PowerShellDataFile -Path $ReleaseManifestPath') -and $source.Contains('Get-AuthenticodeSignature -FilePath $ReleaseManifestPath') -and $source.Contains('Get-AuthenticodeSignature -FilePath $ScriptPath')) 'production validation checks signed release metadata and the script signature'
    Assert-True ($source.Contains('if (-not $DevelopmentMode -and -not $SkipRelocate)') -and $source.Contains('Development mode does not register a scheduled task')) 'development mode avoids relocation and scheduled-task registration'
    Assert-True ($installerBatchSource.Contains('-ExecutionPolicy AllSigned') -and $installerBatchSource -notmatch '-ExecutionPolicy Bypass') 'the batch production launcher requests signed scripts'
    Assert-True ($source.Contains('$DownloadUrl = "https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe"')) 'the official vendor download URL is preserved'
    $versionPolicy = '$MinimumOOShutUpVersion = [version]' + [char]39 + '3.6.1135' + [char]39
    Assert-True ($source.Contains($versionPolicy)) 'the supported version floor is explicit'
    Assert-True ($source.Contains('Get-AuthenticodeSignature -FilePath $Path')) 'Authenticode is checked before accepting a payload'
    Assert-True ($source.Contains('The O&O download could not be authenticated. No executable was started; verify the vendor signature before retrying.')) 'the required authentication fallback is present'
    Assert-True ($source.Contains('The approved settings profile is missing or invalid. No privacy settings were applied; place the reviewed profile in the configured directory.')) 'the required profile fallback is present'
    Assert-True ($readmeSource.Contains('Configuration is unavailable from a clean checkout.') -and $readmeSource.Contains('the `profiles/` directory contains only `README.md`') -and $readmeSource.Contains('No default settings are included or implied.')) 'the README clearly labels the clean-checkout profile blocker'
    Assert-True ($source.Contains('The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.')) 'unknown scope has the required fail-closed status'
    Assert-True ($source.Contains('The SYSTEM task target ACL could not be verified. No task was registered; move the script to a protected path and retry.')) 'the required task ACL fallback is present'
    Assert-True ($installerBatchSource.Contains('set "INSTALLER_EXIT_CODE=!ERRORLEVEL!"') -and $installerBatchSource.Contains('exit /b !INSTALLER_EXIT_CODE!')) 'the batch wrapper preserves the PowerShell installer exit code'
    Assert-True ($installerBatchSource.Contains('Setup is partial. Repair the update task before relying on automatic reapplication.')) 'the batch wrapper reports partial setup with the task recovery action'
    Assert-True ($installerBatchSource -notmatch '\[INFO\] Complete\.') 'the batch wrapper does not print unconditional completion'
    Assert-True ($source.Contains('Task ownership is ambiguous. No task was deleted; inspect the listed path and retry after confirming ownership.')) 'ambiguous task ownership has the required fallback'
    Assert-True ($source.Contains('The task security descriptor could not be verified. No new task was enabled; inspect task permissions before retrying.')) 'task descriptor failures use the exact required fallback'
    Assert-True ($source.Contains("O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)")) 'the task descriptor allowlist is explicit for SYSTEM and Administrators'
    Assert-True ($source.Contains('-RequireMachineScope:$ReapplyOnly')) 'ReapplyOnly validates machine scope before later work'
    Assert-True ($readmeSource.Contains('one-shot `-ReapplyOnly` operator command') -and $readmeSource.Contains('There is no startup or periodic trigger')) 'the README documents the bounded manual recovery path and the absence of automatic reconciliation'
    $mainScriptStart = $source.IndexOf('#region Main Script', [System.StringComparison]::Ordinal)
    Assert-True ($mainScriptStart -ge 0) 'the main installer entry point exists'
    $mainScriptSource = $source.Substring($mainScriptStart)
    $profileValidationIndex = $mainScriptSource.IndexOf('$approvedProfile = Get-OOShutUpApprovedProfile', [System.StringComparison]::Ordinal)
    $missingProfileGuardIndex = $mainScriptSource.IndexOf('if ($null -eq $approvedProfile)', [System.StringComparison]::Ordinal)
    $missingProfileExitIndex = if ($missingProfileGuardIndex -ge 0) { $mainScriptSource.IndexOf('exit 1', $missingProfileGuardIndex, [System.StringComparison]::Ordinal) } else { -1 }
    Assert-True ($profileValidationIndex -ge 0 -and $missingProfileGuardIndex -gt $profileValidationIndex -and $missingProfileExitIndex -gt $missingProfileGuardIndex) 'the main entry point exits when profile validation fails'
    foreach ($sideEffectMarker in @(
        'Get-ScriptRelocationPackageTarget',
        'Install-OOShutUp',
        'Invoke-OOShutUpConfigurationWithRestorePointGate',
        'Invoke-OOShutUpTaskSetup'
    )) {
        $sideEffectIndex = $mainScriptSource.IndexOf($sideEffectMarker, [System.StringComparison]::Ordinal)
        Assert-True ($sideEffectIndex -gt $missingProfileExitIndex) "the main entry point defers $sideEffectMarker until after profile validation"
    }
    Assert-True ($source.Contains('Scope declarations and evidence references are reviewer claims; the validator does not authenticate their source or observe registry changes.')) 'interactive approval states the limits of scope and evidence validation'
    Assert-True ($source.Contains('Before approving a SYSTEM profile, verify every setting against current O&O information and its disposable two-account registry record. Omit any entry with unconfirmed scope.')) 'interactive approval requires current product and two-account evidence review'
    Assert-True ($source.Contains("Join-Path `$PSScriptRoot 'profiles'")) 'the profile package path is rooted beside the script'
    Assert-True ($source -notmatch 'Start-Process[^\r\n]*ooshutup10\.cfg') 'the O&O invocation does not pass a bare profile filename'
    Assert-True ($source -notmatch '\$TempPath') 'the managed executable path no longer depends on a TEMP-derived variable'
    Assert-True ($source.Contains('$env:ProgramW6432') -and $source.Contains('$env:ProgramFiles')) 'the machine install directory uses the machine Program Files location'
    Assert-True ($source.Contains("`$script:OOShutUpMutationMutexName = 'Global\myTech.Today.OOShutUp.Mutation'")) 'manual and SYSTEM entry points use the shared machine-wide O&O mutex'
    Assert-True ($source.Contains('$script:OOShutUpMutationLockTimeoutMilliseconds = 600000')) 'the O&O mutex acquisition is bounded at ten minutes'
    Assert-True ($source.Contains('$script:OOShutUpChildTimeoutSeconds = 900')) 'the O&O child wait has a documented fifteen-minute bound'
    Assert-True ($source.Contains("`$script:OOShutUpTaskExecutionTimeLimit = 'PT30M'")) 'the scheduled task has a documented thirty-minute execution limit'
    Assert-True ($source -notmatch '(?m)Start-Process[^\r\n]*-Wait') 'the O&O child is started without an unbounded Start-Process wait'

    $requiredFunctions = @(
        'Get-OOShutUpTrustedReleaseSignerThumbprint',
        'Test-OOShutUpEnforcedAllSignedPolicy',
        'Test-OOShutUpMachineTrustedPublisher',
        'Test-OOShutUpSourceCheckout',
        'Test-OOShutUpSourceIntegrity',
        'New-OOShutUpMutationMutex',
        'Open-OOShutUpMutationMutex',
        'Release-OOShutUpMutationMutex',
        'Invoke-WithOOShutUpMutationLock',
        'Get-OOShutUpFileVersionInfo',
        'Test-OOShutUpPayload',
        'Test-OOShutUpInstalled',
        'Get-OOShutUpPath',
        'Get-OOShutUpApprovedProfile',
        'Get-OOShutUpProfileDigest',
        'Write-OOShutUpProfileSnapshotFile',
        'Remove-OOShutUpProfileSnapshot',
        'Get-OOShutUpRetainedProfileSnapshots',
        'Wait-OOShutUpChildProcess',
        'New-OOShutUpConfigurationOutcome',
        'New-OOShutUpProfileSnapshot',
        'Show-OOShutUpProfileImpactSummary',
        'Confirm-OOShutUpProfile',
        'New-OOShutUpWebClient',
        'Move-OOShutUpFileAtomically',
        'Install-OOShutUp',
        'Invoke-OOShutUpConfigurationCore',
        'Invoke-OOShutUpConfiguration',
        'ConvertTo-SystemTaskSid',
        'Test-SystemTaskTargetAcl',
        'Get-WindowsPowerShellExecutablePaths',
        'Test-WindowsPowerShellPathIdentity',
        'Get-TrustedWindowsPowerShellExecutable',
        'Test-WindowsUpdateTaskDefinition',
        'Test-WindowsUpdateLegacyTaskDefinition',
        'Test-ScheduledTaskAccessDenied',
        'Invoke-WindowsUpdateTaskSchedulerOperation',
        'Get-WindowsUpdateTaskSecurityDescriptor',
        'Set-WindowsUpdateTaskSecurityDescriptor',
        'Set-WindowsUpdateTaskEnabled',
        'Test-WindowsUpdateTaskSecurityDescriptor',
        'Test-WindowsUpdateTaskSecurityDescriptorEqual',
        'Assert-WindowsUpdateTaskSecurityDescriptor',
        'Test-WindowsUpdateTaskMatchesSnapshot',
        'New-WindowsUpdateScheduledTask',
        'Invoke-OOShutUpTaskSetup',
        'Get-OOShutUpInstallOutcome',
        'Write-OOShutUpInstallSummary',
        'Read-OOShutUpRestartKey',
        'Wait-OOShutUpRestartSaveWindow',
        'Invoke-OOShutUpRestart'
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

    function Get-OOShutUpTrustedReleaseSignerThumbprint {
        if ([string]$script:MockTrustedReleaseSignerThumbprint -notmatch '^[A-Fa-f0-9]{40}$') {
            throw 'The mocked machine release signer policy is missing or malformed.'
        }
        return ([string]$script:MockTrustedReleaseSignerThumbprint).ToUpperInvariant()
    }

    function Test-OOShutUpMachineTrustedPublisher {
        param([string]$Thumbprint)
        return [bool]$script:MockMachineTrustedPublisher
    }

    function Write-ReleaseManifestFixture {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$ScriptDigest
        )

        $manifestText = @"
@{
    schemaVersion = 1
    scriptFile = 'Install-OOShutUp10.ps1'
    scriptSha256 = '$ScriptDigest'
}
"@
        [System.IO.File]::WriteAllText($Path, $manifestText, (New-Object System.Text.UTF8Encoding($false)))
    }

    $releaseFixtureRoot = Join-Path $script:FixtureRoot 'release-integrity'
    [System.IO.Directory]::CreateDirectory($releaseFixtureRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory((Join-Path $releaseFixtureRoot '.git')) | Out-Null
    $releaseScriptPath = Join-Path $releaseFixtureRoot 'Install-OOShutUp10.ps1'
    $releaseManifestPath = Join-Path $releaseFixtureRoot 'Install-OOShutUp10.release.psd1'
    [System.IO.File]::WriteAllText($releaseScriptPath, 'trusted installer fixture', (New-Object System.Text.UTF8Encoding($false)))
    $releaseScriptDigest = (Get-FileHash -LiteralPath $releaseScriptPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Write-ReleaseManifestFixture -Path $releaseManifestPath -ScriptDigest $releaseScriptDigest
    $releaseSignerThumbprint = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
    $releaseSigner = New-MockSigner -Name 'Approved Release Publisher' -Thumbprint $releaseSignerThumbprint
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseScriptPath)] = [pscustomobject]@{
        Status = 'Valid'
        SignerCertificate = $releaseSigner
    }
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseManifestPath)] = [pscustomobject]@{
        Status = 'Valid'
        SignerCertificate = $releaseSigner
    }
    $script:MockTrustedReleaseSignerThumbprint = $releaseSignerThumbprint
    $script:MockMachineTrustedPublisher = $true
    $script:MockEnforcedExecutionPolicy = 'AllSigned'

    $trustedRelease = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath
    Assert-Equal $true $trustedRelease.Allowed 'a signed source and matching signed manifest from the configured publisher pass'
    Assert-Equal 'Production' $trustedRelease.Mode 'a trusted release is classified as production'

    $signatureCallsBeforeUnsigned = $script:SignatureCallCount
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseScriptPath)] = [pscustomobject]@{
        Status = 'NotSigned'
        SignerCertificate = $null
    }
    $unsignedRelease = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $unsignedRelease.Allowed 'an unsigned installer source is rejected'
    Assert-True ($unsignedRelease.Message.Contains('signature is missing or invalid')) 'unsigned source rejection explains the failed signature check'
    Assert-True ($script:SignatureCallCount -gt $signatureCallsBeforeUnsigned) 'unsigned fixture reaches only signature validation'

    [System.IO.File]::WriteAllText($releaseScriptPath, 'modified installer fixture', (New-Object System.Text.UTF8Encoding($false)))
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseScriptPath)] = [pscustomobject]@{
        Status = 'HashMismatch'
        SignerCertificate = $releaseSigner
    }
    $modifiedRelease = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $modifiedRelease.Allowed 'a modified installer source with a broken Authenticode signature is rejected'

    [System.IO.File]::WriteAllText($releaseScriptPath, 'trusted installer fixture', (New-Object System.Text.UTF8Encoding($false)))
    $wrongPublisher = New-MockSigner -Name 'Wrong Release Publisher' -Thumbprint ('B' * 40)
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseScriptPath)] = [pscustomobject]@{
        Status = 'Valid'
        SignerCertificate = $wrongPublisher
    }
    $wrongPublisherRelease = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $wrongPublisherRelease.Allowed 'a valid signature from the wrong publisher is rejected'
    Assert-True ($wrongPublisherRelease.Message.Contains('publisher does not match')) 'wrong-publisher rejection identifies the trust mismatch'

    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseScriptPath)] = [pscustomobject]@{
        Status = 'Valid'
        SignerCertificate = $releaseSigner
    }
    $missingManifestPath = Join-Path $releaseFixtureRoot 'missing.release.psd1'
    $missingManifest = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $missingManifestPath
    Assert-Equal $false $missingManifest.Allowed 'a missing release manifest is rejected'
    Assert-True ($missingManifest.Message.Contains('release manifest is missing')) 'missing manifest rejection is actionable'

    [System.IO.File]::WriteAllText($releaseManifestPath, "@{`r`n schemaVersion = 1`r`n", (New-Object System.Text.UTF8Encoding($false)))
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseManifestPath)] = [pscustomobject]@{
        Status = 'HashMismatch'
        SignerCertificate = $releaseSigner
    }
    $tamperedManifest = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $tamperedManifest.Allowed 'tampered signed release metadata is rejected'

    Write-ReleaseManifestFixture -Path $releaseManifestPath -ScriptDigest ('0' * 64)
    $script:ReleaseSignatureByPath[[System.IO.Path]::GetFullPath($releaseManifestPath)] = [pscustomobject]@{
        Status = 'Valid'
        SignerCertificate = $releaseSigner
    }
    $corruptDigest = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $corruptDigest.Allowed 'signed but mismatched release digest metadata is rejected'

    Write-ReleaseManifestFixture -Path $releaseManifestPath -ScriptDigest $releaseScriptDigest
    $script:MockTrustedReleaseSignerThumbprint = ''
    $missingTrustPin = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $missingTrustPin.Allowed 'missing external signer trust metadata is rejected'
    $script:MockTrustedReleaseSignerThumbprint = $releaseSignerThumbprint

    $script:MockMachineTrustedPublisher = $false
    $missingMachinePublisher = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $missingMachinePublisher.Allowed 'a signer missing from LocalMachine Trusted Publishers is rejected'
    $script:MockMachineTrustedPublisher = $true

    $script:MockEnforcedExecutionPolicy = 'RemoteSigned'
    $script:MockUserExecutionPolicy = 'AllSigned'
    $signatureCallsBeforePolicyFailure = $script:SignatureCallCount
    $unenforcedPolicy = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -ReleaseManifestPath $releaseManifestPath
    Assert-Equal $false $unenforcedPolicy.Allowed 'UserPolicy AllSigned without machine-enforced AllSigned is rejected'
    Assert-Equal $signatureCallsBeforePolicyFailure $script:SignatureCallCount 'missing machine policy fails before signature or package operations'
    $script:MockEnforcedExecutionPolicy = 'AllSigned'
    $script:MockUserExecutionPolicy = 'Undefined'

    $developmentRelease = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -DevelopmentMode
    Assert-Equal $true $developmentRelease.Allowed 'explicit development mode is available inside a Git source checkout'
    Assert-Equal 'Development' $developmentRelease.Mode 'an unsigned source checkout is clearly classified as development'
    $nonCheckoutScript = Join-Path $script:FixtureRoot 'outside-checkout.ps1'
    [System.IO.File]::WriteAllText($nonCheckoutScript, 'untrusted fixture', (New-Object System.Text.UTF8Encoding($false)))
    $developmentOutsideCheckout = Test-OOShutUpSourceIntegrity -ScriptPath $nonCheckoutScript -DevelopmentMode
    Assert-Equal $false $developmentOutsideCheckout.Allowed 'development mode is refused outside a Git source checkout'
    $developmentReapply = Test-OOShutUpSourceIntegrity -ScriptPath $releaseScriptPath -DevelopmentMode -ReapplyOnly
    Assert-Equal $false $developmentReapply.Allowed 'development mode cannot be used by the unattended task entry point'

    $childWaitSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Wait-OOShutUpChildProcess' })[0].Extent.Text)
    $configurationSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpConfigurationCore' })[0].Extent.Text)
    Assert-True ($childWaitSource.Contains('$Process.WaitForExit($timeoutMilliseconds)') -and $childWaitSource.Contains('$Process.HasExited')) 'the bounded wait observes only the exact process object returned by Start-Process'
    Assert-True ($childWaitSource -notmatch '\.Kill\(|Stop-Process|taskkill') 'timeout handling never terminates a process by PID or by name'
    Assert-True ($configurationSource.Contains('Get-OOShutUpRetainedProfileSnapshots') -and $configurationSource.Contains('Wait-OOShutUpChildProcess -Process $processInfo')) 'retained snapshots block retries and the exact started child is waited'

    Assert-Equal $true (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $script:OOShutUpTaskSecurityDescriptorSddl) 'the documented SYSTEM and Administrators task descriptor passes verification'
    Assert-Equal $false (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor ($script:OOShutUpTaskSecurityDescriptorSddl + '(A;;FA;;;AU)')) 'an unexpected Authenticated Users write ACE fails descriptor verification'
    Assert-Equal $false (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor ($script:OOShutUpTaskSecurityDescriptorSddl + '(A;;FA;;;BU)')) 'an unexpected BUILTIN Users write ACE fails descriptor verification'
    Assert-Equal $false (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor 'not-an-sddl-descriptor') 'malformed task descriptor data fails verification'
    $reorderedAllowedDescriptor = 'O:BAG:SYD:P(A;;FA;;;BA)(A;;FA;;;SY)'
    Assert-Equal $true (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $reorderedAllowedDescriptor) 'the allowlist check is independent of ACE ordering'
    Assert-Equal $false (Test-WindowsUpdateTaskSecurityDescriptorEqual -ActualSecurityDescriptor $script:OOShutUpTaskSecurityDescriptorSddl -ExpectedSecurityDescriptor $reorderedAllowedDescriptor) 'rollback comparison detects a descriptor that differs from the exact saved descriptor'

    function Get-WindowsUpdateTaskSecurityDescriptor {
        [CmdletBinding()]
        param([string]$TaskPath, [string]$TaskName)

        $key = "$TaskPath|$TaskName"
        [void]$script:TaskSecurityDescriptorReadCalls.Add([pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName })
        if ($script:MockTaskSecurityDescriptorScenario -eq 'ReadDeniedOnOwnedTask' -and $TaskName -eq 'Install-OOShutUp10') {
            throw [System.UnauthorizedAccessException]::new(("{0} Details: Mock task descriptor access denied." -f $script:OOShutUpTaskSecurityDescriptorFailureMessage))
        }
        if ($script:MockTaskSecurityDescriptorScenario -eq 'ReadDeniedOnStage' -and $TaskName -like 'Install-OOShutUp10-validation-*') {
            throw [System.UnauthorizedAccessException]::new(("{0} Details: Mock staged-task descriptor access denied." -f $script:OOShutUpTaskSecurityDescriptorFailureMessage))
        }
        if ($script:MockTaskSecurityDescriptorScenario -eq 'MalformedOnOwnedTask' -and $TaskName -eq 'Install-OOShutUp10') {
            return 'not-an-sddl-descriptor'
        }
        if ($script:MockTaskSecurityDescriptorScenario -eq 'MalformedOnStage' -and $TaskName -like 'Install-OOShutUp10-validation-*') {
            return 'not-an-sddl-descriptor'
        }
        if ($script:TaskSecurityDescriptorByIdentity.ContainsKey($key)) {
            return [string]$script:TaskSecurityDescriptorByIdentity[$key]
        }
        return [string]$script:OOShutUpTaskSecurityDescriptorSddl
    }

    function Set-WindowsUpdateTaskSecurityDescriptor {
        [CmdletBinding()]
        param([string]$TaskPath, [string]$TaskName, [string]$SecurityDescriptor)

        [void]$script:TaskSecurityDescriptorSetCalls.Add([pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; SecurityDescriptor = $SecurityDescriptor })
        if ($script:MockTaskSecurityDescriptorSetFailureTaskName -and $TaskName -like $script:MockTaskSecurityDescriptorSetFailureTaskName) {
            throw [System.InvalidOperationException]::new('Mock task descriptor update failed.')
        }

        $descriptorToStore = $SecurityDescriptor
        $injectReplacementDescriptor = $TaskName -eq 'Install-OOShutUp10' -and
            -not $script:MockTaskSecurityDescriptorReplacementInjected -and
            $script:MockTaskSecurityDescriptorScenario -in @('UnexpectedOnFinal', 'UnprovenOnFinal')
        if ($injectReplacementDescriptor) {
            $script:MockTaskSecurityDescriptorReplacementInjected = $true
            $descriptorToStore = $script:OOShutUpTaskSecurityDescriptorSddl + '(A;;FA;;;AU)'
        }
        $key = "$TaskPath|$TaskName"
        $script:TaskSecurityDescriptorByIdentity[$key] = $descriptorToStore

        if ($injectReplacementDescriptor -and $script:MockTaskSecurityDescriptorScenario -eq 'UnprovenOnFinal') {
            $task = $script:TaskRegistry[$key]
            $document = New-Object System.Xml.XmlDocument
            $document.LoadXml([string]$task.Definition)
            $document.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Arguments']").InnerText = '-UnrecognizedAction'
            $script:TaskRegistry[$key] = New-MockScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Xml $document.OuterXml
        }
    }

    function Set-WindowsUpdateTaskEnabled {
        [CmdletBinding()]
        param([string]$TaskPath, [string]$TaskName, [bool]$Enabled)

        [void]$script:TaskEnabledCalls.Add([pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Enabled = $Enabled })
        $key = "$TaskPath|$TaskName"
        if (-not $script:TaskRegistry.ContainsKey($key)) {
            throw "Mock task was not found: $key"
        }
        $task = $script:TaskRegistry[$key]
        $document = New-Object System.Xml.XmlDocument
        $document.LoadXml([string]$task.Definition)
        $enabledNode = $document.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
        if ($null -eq $enabledNode) {
            throw 'Mock task definition has no Settings/Enabled element.'
        }
        $enabledNode.InnerText = $Enabled.ToString().ToLowerInvariant()
        $script:TaskRegistry[$key] = New-MockScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Xml $document.OuterXml
    }

    $managedPathFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Get-OOShutUpPath' })[0].Extent.Text)
    Assert-True ($managedPathFunctionSource.Contains('Join-Path $InstallPath $ExecutableName')) 'path lookup uses the single managed machine directory'
    Assert-True ($managedPathFunctionSource -notmatch '\$env:(?:TEMP|USERPROFILE|LOCALAPPDATA)|Get-Command') 'path lookup does not search user, temporary, or PATH locations'

    $mainStart = $source.IndexOf('#region Main Script', [System.StringComparison]::Ordinal)
    $mainSource = $source.Substring($mainStart)
    $profileGatePosition = $mainSource.IndexOf('$approvedProfile = Get-OOShutUpApprovedProfile', [System.StringComparison]::Ordinal)
    Assert-True ($profileGatePosition -ge 0) 'main execution performs the profile preflight'
    $profileApprovalGatePosition = $mainSource.IndexOf('$profileApprovalAccepted = Confirm-OOShutUpProfile', [System.StringComparison]::Ordinal)
    Assert-True ($profileApprovalGatePosition -gt $profileGatePosition) 'main execution requires an approval decision after profile validation'
    $relocationTargetPosition = $mainSource.IndexOf('$relocationTarget = Get-ScriptRelocationPackageTarget', [System.StringComparison]::Ordinal)
    $relocationPosition = $mainSource.IndexOf('Invoke-ScriptRelocation -CurrentScriptPath', [System.StringComparison]::Ordinal)
    Assert-True ($relocationTargetPosition -gt $profileApprovalGatePosition -and $relocationPosition -gt $relocationTargetPosition) 'the content-addressed package target is derived only after approval and before relocation'
    Assert-True ($mainSource.Contains('-ScriptFileName $ScriptFileName') -and $mainSource.Contains('$DefaultScriptPath = $relocationTarget.ScriptPath')) 'the package and task use the same installed script entry-point path'
    $reapplyStart = $mainSource.IndexOf('if ($ReapplyOnly)', [System.StringComparison]::Ordinal)
    $reapplyPathPosition = $mainSource.IndexOf('$exePath = Get-OOShutUpPath', $reapplyStart, [System.StringComparison]::Ordinal)
    Assert-True ($reapplyStart -ge 0 -and $reapplyPathPosition -gt $reapplyStart) 'SYSTEM ReapplyOnly resolves the executable through the managed-path lookup'
    Assert-True ($mainSource.Contains('Get-OOShutUpPath -ExpectedPath $ApprovedOOSExecutablePath')) 'SYSTEM ReapplyOnly validates the exact executable path pinned in the task action'
    Assert-True ($source.Contains('The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.') -and $mainSource.Contains('$script:OOShutUpManagedPathFailureMessage')) 'SYSTEM ReapplyOnly reports the exact managed-copy recovery instruction'
    $reapplyBranchEnd = $mainSource.IndexOf('# Normal installation mode', $reapplyStart, [System.StringComparison]::Ordinal)
    $reapplyBranch = $mainSource.Substring($reapplyStart, $reapplyBranchEnd - $reapplyStart)
    Assert-Equal 1 ([regex]::Matches($reapplyBranch, 'Invoke-OOShutUpConfiguration -ExecutablePath \$exePath -SystemReapply').Count) 'one ReapplyOnly invocation makes at most one bounded reapply per operator command'
    Assert-True ($reapplyBranch -notmatch '(?i)\b(Get|Start|Register|Unregister)-ScheduledTask\b') 'manual ReapplyOnly recovery does not depend on Task Scheduler availability'
    Assert-True ($reapplyBranch -notmatch '(?i)Invoke-OOShutUpTaskSetup|Invoke-OOShutUpRestart|\bwhile\s*\(|\bfor\s*\(') 'ReapplyOnly recovery creates no task, requests no restart, and adds no retry loop'
    Assert-True ($reapplyBranch.Contains('exit 0')) 'successful ReapplyOnly recovery exits before normal installation and restart handling'
    foreach ($mutationMarker in @(
        'Invoke-ScriptRelocation -CurrentScriptPath',
        'Invoke-OOShutUpConfigurationWithRestorePointGate',
        'Invoke-OOShutUpConfiguration -ExecutablePath',
        'Invoke-OOShutUpTaskSetup'
    )) {
        $mutationPosition = $mainSource.IndexOf($mutationMarker, [System.StringComparison]::Ordinal)
        Assert-True ($mutationPosition -gt $profileApprovalGatePosition) "profile approval precedes $mutationMarker"
    }
    Assert-True ($mainSource.Contains('-ScriptPath $currentScriptPath') -and $mainSource.Contains('-ApprovedProfileSha256 $approvedProfile.Sha256')) 'scheduled task setup receives the running package and validated profile digest'
    Assert-True ($mainSource.Contains('-ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256')) 'scheduled task setup receives the exact scope manifest digest'
    $taskSetupCallPosition = $mainSource.IndexOf('$taskSetup = Invoke-OOShutUpTaskSetup', [System.StringComparison]::Ordinal)
    $taskScriptPathPosition = $mainSource.IndexOf('-ScriptPath $currentScriptPath', $taskSetupCallPosition, [System.StringComparison]::Ordinal)
    Assert-True ($taskSetupCallPosition -ge 0 -and $taskScriptPathPosition -gt $taskSetupCallPosition) 'task setup uses the current complete package path after relocation or in-place fallback'
    Assert-True ($source.Contains('The privacy settings may be applied, but the update task was not verified. Setup is partial; repair the task before relying on automatic reapplication.')) 'task failure uses the exact partial setup status'
    Assert-True ($source.Contains('Verify \myTech.Today\Install-OOShutUp10 in Task Scheduler')) 'partial setup includes the exact task identity for recovery'
    Assert-True ($mainSource.Contains('$taskSetup = Invoke-OOShutUpTaskSetup') -and $mainSource.Contains('$installerOutcome = Get-OOShutUpInstallOutcome')) 'the normal flow consumes explicit task and configuration outcomes'
    Assert-True ($source -notmatch '\$SettingsVerified|SettingsOutcome\s*=\s*''Verified''') 'native O&O exit results are not converted into settings verification'
    Assert-True ($source.Contains('ReadbackState = ''Unavailable''') -and $source.Contains('SettingsReadback = $settingsReadback')) 'the installer reports that independent settings readback is unavailable'
    Assert-True ($mainSource.Contains('exit $installerOutcome.ExitCode')) 'the normal flow returns the computed complete, partial, or failed exit code'
    Assert-True ($mainSource.Contains('Write-OOShutUpInstallSummary -Outcome $installerOutcome')) 'the normal flow logs the summary from the installer outcome contract'
    Assert-True (-not $mainSource.Contains('=== Configuration Complete ===')) 'full-success wording is not unconditional in the normal flow'
    $summaryFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Write-OOShutUpInstallSummary' })[0].Extent.Text)
    $completeStatusPosition = $summaryFunctionSource.IndexOf('if ($Outcome.Status -eq ''Complete'')', [System.StringComparison]::Ordinal)
    $completeSummaryPosition = $summaryFunctionSource.IndexOf("Write-Log '=== Configuration Complete ===' -Level SUCCESS", [System.StringComparison]::Ordinal)
    $partialStatusPosition = $summaryFunctionSource.IndexOf('elseif ($Outcome.Status -eq ''Partial'')', [System.StringComparison]::Ordinal)
    Assert-True ($completeStatusPosition -ge 0 -and $completeSummaryPosition -gt $completeStatusPosition -and $partialStatusPosition -gt $completeSummaryPosition) 'complete wording is guarded by the computed complete outcome'
    $noChangeRestartGatePosition = $mainSource.IndexOf("if (`$configurationOutcome.State -eq 'Changed')", [System.StringComparison]::Ordinal)
    $restartHandlerCallPosition = $mainSource.IndexOf('Invoke-OOShutUpRestart -NonInteractive:', [System.StringComparison]::Ordinal)
    $restartHandlerSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpRestart' })[0].Extent.Text)
    Assert-True ($noChangeRestartGatePosition -ge 0 -and $restartHandlerCallPosition -gt $noChangeRestartGatePosition) 'the restart handler is called only when O&O reports changed settings'
    Assert-True ($restartHandlerSource.Contains('Press Y to restart, N or Esc to defer') -and $restartHandlerSource.Contains('Restart-Computer -Force')) 'the guarded restart handler owns the confirmation prompt and explicit force command'
    $taskFunctionSource = ([string](@($functionAsts | Where-Object { $_.Name -eq 'New-WindowsUpdateScheduledTask' })[0].Extent.Text))
    $taskApprovalPosition = $taskFunctionSource.IndexOf('$currentApprovedProfile = Get-OOShutUpApprovedProfile', [System.StringComparison]::Ordinal)
    $pathCanonicalizationPosition = $taskFunctionSource.IndexOf('$ScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)', [System.StringComparison]::Ordinal)
    $targetAclPosition = $taskFunctionSource.IndexOf('Test-SystemTaskTargetAcl -ScriptPath $ScriptPath', [System.StringComparison]::Ordinal)
    $taskLookupPosition = $taskFunctionSource.IndexOf('$existingTasks = @(Get-ScheduledTask', [System.StringComparison]::Ordinal)
    Assert-True ($taskApprovalPosition -ge 0 -and $pathCanonicalizationPosition -gt $taskApprovalPosition -and $targetAclPosition -gt $pathCanonicalizationPosition -and $taskLookupPosition -gt $targetAclPosition) 'profile approval, path canonicalization, and target ACL verification precede task inspection or mutation'
    Assert-True ($taskFunctionSource -notmatch 'OOShutUp10Task_Debug\.xml' -and $taskFunctionSource -notmatch '\$env:TEMP') 'task setup does not export diagnostic XML to a predictable TEMP path'
    Assert-True ($taskFunctionSource.Contains('Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath')) 'replacement queries the exact owned task path'
    Assert-True ($taskFunctionSource.Contains('Removing verified owned scheduled task')) 'removal logs the exact verified task path'
    Assert-True ($taskFunctionSource.Contains('Register-ScheduledTask -TaskName $stagingTaskName')) 'a disabled validation task is registered before replacing the existing task'
    Assert-True ($taskFunctionSource.Contains('The previous task definition was {1}')) 'failed replacement reports whether rollback preserved or restored the prior task'
    Assert-True ($taskFunctionSource -notmatch 'schtasks\.exe\s+/Delete') 'task replacement does not delete the previous task before validation'

    $savedTaskCreator = (Get-Command -Name New-WindowsUpdateScheduledTask -CommandType Function).ScriptBlock
    try {
        Set-Item -Path Function:\New-WindowsUpdateScheduledTask -Value {
            param(
                [string]$ScriptPath,
                [string]$ApprovedProfileSha256,
                [string]$ApprovedProfileManifestSha256
            )

            switch ($script:MockInstallerTaskScenario) {
                'Success' { return $true }
                'Failure' { return $false }
                'Exception' { throw 'Mock unexpected scheduler exception.' }
                default { throw 'Unknown mocked task setup scenario.' }
            }
        }

        $changedConfigurationOutcome = New-OOShutUpConfigurationOutcome `
            -State Changed `
            -ExitCode 30 `
            -Evidence 'Mock O&O exit code 30' `
            -Message 'O&O reported changed settings.' `
            -NextStep Proceed
        foreach ($scenario in @('Success', 'Failure', 'Exception')) {
            $script:MockInstallerTaskScenario = $scenario
            $taskSetupResult = Invoke-OOShutUpTaskSetup -ScriptPath 'C:\fixture\Install-OOShutUp10.ps1' -ApprovedProfileSha256 ('A' * 64) -ApprovedProfileManifestSha256 ('B' * 64)
            $installResult = Get-OOShutUpInstallOutcome -ConfigurationOutcome $changedConfigurationOutcome -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $taskSetupResult
            $script:LogMessages.Clear()
            Write-OOShutUpInstallSummary -Outcome $installResult
            $summaryLog = [string]::Join([Environment]::NewLine, [string[]]$script:LogMessages.ToArray())
            if ($scenario -eq 'Success') {
                Assert-Equal $true $taskSetupResult.Verified 'a true task creation and verification result is preserved'
                Assert-Equal 0 $installResult.ExitCode 'O&O-reported settings and a verified task return the success exit code'
                Assert-Equal 'Complete' $installResult.Status 'O&O-reported settings and a verified task report complete status'
                Assert-Equal 'Verified' $installResult.TaskOutcome 'complete status identifies the task as verified'
                Assert-Equal 'Changed' $installResult.SettingsOutcome 'the install outcome preserves the O&O-reported changed state'
                Assert-Equal 'Unavailable' $installResult.SettingsReadback 'the install outcome does not claim independent settings readback'
                Assert-True ($installResult.Summary.Contains('not independently read back') -and -not $installResult.Summary.Contains('privacy settings are verified')) 'the completion summary distinguishes O&O status from independent verification'
                Assert-True ($installResult.Summary.Contains('\myTech.Today\Install-OOShutUp10') -and -not $installResult.Summary.Contains('partial')) 'success summary identifies the verified task'
                Assert-True ($summaryLog.Contains('=== Configuration Complete ===') -and $summaryLog.Contains($installResult.Summary) -and -not $summaryLog.Contains('=== Setup Partial ===')) 'success logging agrees with the verified task outcome'
            }
            else {
                Assert-Equal $false $taskSetupResult.Verified "$scenario task setup is not treated as verified"
                Assert-Equal 2 $installResult.ExitCode "$scenario task setup returns the distinct partial exit code"
                Assert-Equal 'Partial' $installResult.Status "$scenario task setup reports partial status"
                Assert-Equal 'Not verified' $installResult.TaskOutcome "$scenario task setup reports the task as unverified"
                Assert-True ($installResult.Summary.Contains($script:OOShutUpPartialSetupMessage) -and $installResult.Summary.Contains('O&O reported that settings changed') -and $installResult.Summary.Contains('not independently read back')) "$scenario summary preserves the partial wording and identifies vendor-reported settings state"
                Assert-True ($installResult.RecoveryMessage.Contains('\myTech.Today\Install-OOShutUp10') -and $installResult.RecoveryMessage.Contains('same-named tasks in other folders are left unchanged')) "$scenario recovery identifies the owned task and preserves unrelated tasks"
                Assert-True (-not $summaryLog.Contains('=== Configuration Complete ===') -and $summaryLog.Contains('=== Setup Partial ===') -and $summaryLog.Contains($installResult.Summary) -and $summaryLog.Contains($installResult.RecoveryMessage)) "$scenario logging and recovery message agree with the partial task outcome"
            }
            if ($scenario -eq 'Exception') {
                Assert-Equal 'Mock unexpected scheduler exception.' $taskSetupResult.ExceptionMessage 'unexpected task exceptions are returned as details instead of escaping the installer'
            }
            else {
                Assert-Equal $null $taskSetupResult.ExceptionMessage "$scenario has no unexpected exception detail"
            }
        }

        $overrideSetupResult = [pscustomobject]@{ Verified = $true; ExceptionMessage = $null }
        $overrideInstallResult = Get-OOShutUpInstallOutcome -ConfigurationOutcome $changedConfigurationOutcome -RestorePointVerified $false -RestorePointOverrideUsed $true -TaskSetup $overrideSetupResult
        Assert-Equal 0 $overrideInstallResult.ExitCode 'an explicit restore point override still permits success when the task is verified'
        Assert-Equal 'Explicitly overridden' $overrideInstallResult.RestorePointOutcome 'the result retains the separate restore point outcome'

        $unknownConfigurationOutcome = New-OOShutUpConfigurationOutcome `
            -State Unknown `
            -ExitCode 23 `
            -Evidence 'Unrecognized O&O native exit code' `
            -Message 'O&O returned an unrecognized exit code; review O&O Undo History and current settings before retrying.'
        $unknownInstallResult = Get-OOShutUpInstallOutcome `
            -ConfigurationOutcome $unknownConfigurationOutcome `
            -RestorePointVerified $true `
            -RestorePointOverrideUsed $false `
            -TaskSetup $overrideSetupResult
        Assert-Equal 'Unknown' $unknownInstallResult.SettingsOutcome 'unknown native status stays unknown in the installer outcome'
        Assert-Equal 'Unavailable' $unknownInstallResult.SettingsReadback 'unknown status retains unavailable readback'
        Assert-Equal 'Stop' $unknownConfigurationOutcome.NextStep 'unknown native status stops subsequent installer steps'
        Assert-Equal 'Failed' $unknownInstallResult.Status 'unknown settings state cannot produce a complete installer outcome'
        Assert-Equal 1 $unknownInstallResult.ExitCode 'unknown settings state returns failure rather than success'
        Assert-True ($unknownInstallResult.Summary.Contains('Undo History') -and $unknownInstallResult.Summary.Contains('current settings') -and -not $unknownInstallResult.Summary.Contains('settings are verified')) 'unknown status gives actionable recovery without a verified label'

        $failedConfigurationOutcome = New-OOShutUpConfigurationOutcome `
            -State Failed `
            -Evidence 'Pre-launch validation failure' `
            -Message 'The approved profile could not be validated before launch.'
        $failedInstallResult = Get-OOShutUpInstallOutcome `
            -ConfigurationOutcome $failedConfigurationOutcome `
            -RestorePointVerified $true `
            -RestorePointOverrideUsed $false `
            -TaskSetup $overrideSetupResult
        Assert-Equal 'Failed' $failedInstallResult.SettingsOutcome 'pre-launch failure remains distinct from an unknown post-launch state'
        Assert-Equal 'Failed' $failedInstallResult.Status 'a pre-launch failure cannot produce a complete installer result'
        Assert-Equal 'Unavailable' $failedInstallResult.SettingsReadback 'pre-launch failure does not claim settings readback'
    }
    finally {
        Set-Item -Path Function:\New-WindowsUpdateScheduledTask -Value $savedTaskCreator
    }

    # Replace only the metadata reader after loading production functions so the fixture stays offline.
    function Get-OOShutUpFileVersionInfo {
        param([string]$Path)

        $content = [System.IO.File]::ReadAllText($Path)
        $fileVersion = if ($content -eq 'OLD_VERSION') {
            '3.5.1130.0'
        }
        elseif ($content -eq 'NEWER_VERSION') {
            '3.6.1136.0'
        }
        else {
            '3.6.1135.0'
        }
        $productName = if ($content -eq 'WRONG_PRODUCT') { 'O&O DiskImage' } else { 'O&O ShutUp10++' }
        return [pscustomobject]@{ ProductName = $productName; FileVersion = $fileVersion }
    }

    function New-OOShutUpWebClient {
        $webClient = [pscustomobject]@{}
        $webClient | Add-Member -MemberType ScriptMethod -Name DownloadFile -Value {
            param($Uri, $OutFile)

            $script:DownloadCount++
            if ($Uri -ne 'https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe') {
                throw 'Unexpected download source in test.'
            }

            switch ($script:DownloadScenario) {
                'valid' { Write-TestPayload -Path $OutFile -Content 'VALID' }
                'tampered' { Write-TestPayload -Path $OutFile -Content 'TAMPERED' }
                'unsigned' { Write-TestPayload -Path $OutFile -Content 'UNSIGNED' }
                'wrong-publisher' { Write-TestPayload -Path $OutFile -Content 'WRONG_PUBLISHER' }
                'wrong-product' { Write-TestPayload -Path $OutFile -Content 'WRONG_PRODUCT' }
                'old-version' { Write-TestPayload -Path $OutFile -Content 'OLD_VERSION' }
                'empty' { [System.IO.File]::WriteAllBytes($OutFile, [byte[]]@()) }
                'interrupted' {
                    Write-TestPayload -Path $OutFile -Content 'PARTIAL'
                    throw [System.Net.WebException]::new('Mocked interrupted download.')
                }
                default { throw "Unknown mocked download scenario: $script:DownloadScenario" }
            }
        } -Force
        $webClient | Add-Member -MemberType ScriptMethod -Name Dispose -Value { } -Force
        return $webClient
    }

    $script:DownloadUrl = 'https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe'
    $script:ExecutableName = 'OOSU10.exe'
    $script:InstallPath = Join-Path $script:FixtureRoot 'machine controlled O&O directory with spaces'
    $script:ManagedExecutablePath = [System.IO.Path]::GetFullPath((Join-Path $script:InstallPath $script:ExecutableName))
    $script:ExpectedOOShutUpPublisher = 'O&O Software GmbH'
    $script:MinimumOOShutUpVersion = [version]'3.6.1135'
    $script:OOShutUpPayloadFailureMessage = 'The O&O download could not be authenticated. No executable was started; verify the vendor signature before retrying.'
    $script:OOShutUpProfileDirectory = Join-Path $script:FixtureRoot 'approved profiles with spaces'
    $script:OOShutUpProfilePath = Join-Path $script:OOShutUpProfileDirectory 'ooshutup10.cfg'
    $script:OOShutUpProfileManifestPath = Join-Path $script:OOShutUpProfileDirectory 'ooshutup10.approval.json'
    $script:OOShutUpProfileFailureMessage = 'The approved settings profile is missing or invalid. No privacy settings were applied; place the reviewed profile in the configured directory.'
    $script:OOShutUpProfileVersionFailureMessage = 'The O&O executable version is outside the profile review range. No settings were changed; review and reapprove the profile for this release before applying settings or registering a SYSTEM task.'
    $script:OOShutUpProfileSnapshotFailureMessage = 'The approved profile changed after validation. No O&O process was started; restore the reviewed profile and retry.'
    $script:OOShutUpUnknownScopeFailureMessage = 'The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.'
    $script:OOShutUpCurrentUserScopeFailureMessage = 'The SYSTEM task profile contains current-user settings. Reapplication was skipped; run them for the intended signed-in user or remove them from the SYSTEM task profile.'
    $script:OOShutUpApprovalFailureMessage = 'The selected settings profile has no matching approval. No settings were changed; review both file digests in the interactive installer and recreate the task.'
    $script:OOShutUpTaskAclFailureMessage = 'The SYSTEM task target ACL could not be verified. No task was registered; move the script to a protected path and retry.'
    $script:OOShutUpTaskSecurityDescriptorSddl = 'O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)'
    $script:OOShutUpTaskSecurityDescriptorFailureMessage = 'The task security descriptor could not be verified. No new task was enabled; inspect task permissions before retrying.'
    $script:OOShutUpTaskOwnershipFailureMessage = 'Task ownership is ambiguous. No task was deleted; inspect the listed path and retry after confirming ownership.'
    $script:OOShutUpManagedPathFailureMessage = 'The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.'

    $validFixture = Join-Path $script:FixtureRoot 'valid.exe'
    $systemFullControl = New-MockAclRule -Sid 'S-1-5-18' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl)
    $administratorsFullControl = New-MockAclRule -Sid 'S-1-5-32-544' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl)
    $script:MockAclByPath['__default__'] = New-MockAcl -Rules @($systemFullControl, $administratorsFullControl)
    $script:OOShutUpTrustedPowerShellFailureMessage = 'A trusted Windows PowerShell executable could not be resolved. No scheduled task was registered or changed.'

    $script:MockWindowsPowerShellDirectory = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows)
    $native64BitPaths = Get-WindowsPowerShellExecutablePaths -WindowsDirectory $script:MockWindowsPowerShellDirectory -Is64BitOperatingSystem $true -Is64BitProcess $true
    $wow64CallerPaths = Get-WindowsPowerShellExecutablePaths -WindowsDirectory $script:MockWindowsPowerShellDirectory -Is64BitOperatingSystem $true -Is64BitProcess $false
    $native32BitPaths = Get-WindowsPowerShellExecutablePaths -WindowsDirectory $script:MockWindowsPowerShellDirectory -Is64BitOperatingSystem $false -Is64BitProcess $false
    $expectedNativeInterpreterPath = [System.IO.Path]::GetFullPath((Join-Path $script:MockWindowsPowerShellDirectory 'System32\WindowsPowerShell\v1.0\powershell.exe'))
    Assert-Equal $expectedNativeInterpreterPath ([string]$native64BitPaths.TaskPath) 'a 64-bit caller schedules the canonical native Windows PowerShell path'
    Assert-Equal $expectedNativeInterpreterPath ([string]$native64BitPaths.InspectionPath) 'a 64-bit caller validates the native System32 executable directly'
    Assert-Equal $expectedNativeInterpreterPath ([string]$wow64CallerPaths.TaskPath) 'a 32-bit caller schedules the canonical native System32 path'
    Assert-Equal ([System.IO.Path]::GetFullPath((Join-Path $script:MockWindowsPowerShellDirectory 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'))) ([string]$wow64CallerPaths.InspectionPath) 'a 32-bit caller validates the native executable through Sysnative'
    Assert-Equal ([string]$native32BitPaths.TaskPath) ([string]$native32BitPaths.InspectionPath) 'a 32-bit Windows caller validates its native System32 executable'
    $redirectedInterpreterPath = [System.IO.Path]::GetFullPath((Join-Path $script:MockWindowsPowerShellDirectory 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'))
    Assert-Equal $false (Test-WindowsPowerShellPathIdentity -Candidate $wow64CallerPaths -ResolvedPath $redirectedInterpreterPath) 'a SysWOW64 redirection is rejected for a 32-bit caller on 64-bit Windows'
    Assert-Equal $false (Test-WindowsPowerShellPathIdentity -Candidate $native64BitPaths -ResolvedPath (Join-Path $script:FixtureRoot 'powershell.exe')) 'an unexpected executable path is rejected'
    Assert-Equal $null (Get-TrustedWindowsPowerShellExecutable -WindowsDirectory (Join-Path $script:FixtureRoot 'missing-windows-directory') -Is64BitOperatingSystem $true -Is64BitProcess $true) 'a missing Windows PowerShell candidate fails closed'

    $trustedPowerShellPath = Get-TrustedWindowsPowerShellExecutable
    Assert-Equal $expectedNativeInterpreterPath ([string]$trustedPowerShellPath) 'the trusted Windows PowerShell resolver accepts the Microsoft system binary'
    $script:MockWindowsPowerShellSignatureMode = 'WrongPublisher'
    Assert-Equal $null (Get-TrustedWindowsPowerShellExecutable) 'an unexpected Windows PowerShell signer fails closed'
    $script:MockWindowsPowerShellSignatureMode = 'Unsigned'
    Assert-Equal $null (Get-TrustedWindowsPowerShellExecutable) 'an unsigned Windows PowerShell executable fails closed'
    $script:MockWindowsPowerShellSignatureMode = 'Valid'

    $userWritableInterpreterAcl = New-MockAcl -Rules @(
        $systemFullControl,
        $administratorsFullControl,
        (New-MockAclRule -Sid 'S-1-5-11' -Rights ([System.Security.AccessControl.FileSystemRights]::Modify))
    )
    foreach ($interpreterPath in @($native64BitPaths.TaskPath, $wow64CallerPaths.InspectionPath)) {
        $script:MockAclByPath[[System.IO.Path]::GetFullPath([string]$interpreterPath)] = $userWritableInterpreterAcl
    }
    Assert-Equal $null (Get-TrustedWindowsPowerShellExecutable) 'a user-writable Windows PowerShell path fails closed'
    $script:MockAclByPath.Remove([System.IO.Path]::GetFullPath([string]$native64BitPaths.TaskPath)) | Out-Null
    $script:MockAclByPath.Remove([System.IO.Path]::GetFullPath([string]$wow64CallerPaths.InspectionPath)) | Out-Null

    Write-TestPayload -Path $validFixture -Content 'VALID'
    Assert-Equal $true (Test-OOShutUpPayload -Path $validFixture) 'a valid O&O signature, publisher, product, and supported version pass'

    $invalidFixtures = @(
        @{ Name = 'unsigned'; Content = 'UNSIGNED' },
        @{ Name = 'wrong publisher'; Content = 'WRONG_PUBLISHER' },
        @{ Name = 'tampered'; Content = 'TAMPERED' },
        @{ Name = 'corrupt'; Content = 'CORRUPT' },
        @{ Name = 'unsupported product'; Content = 'WRONG_PRODUCT' },
        @{ Name = 'unsupported old version'; Content = 'OLD_VERSION' }
    )
    foreach ($case in $invalidFixtures) {
        $fixture = Join-Path $script:FixtureRoot ($case.Name.Replace(' ', '-') + '.exe')
        Write-TestPayload -Path $fixture -Content $case.Content
        $script:LogMessages.Clear()
        Assert-Equal $false (Test-OOShutUpPayload -Path $fixture) "$($case.Name) fixture is rejected"
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpPayloadFailureMessage)) "$($case.Name) rejection reports the required fallback"
    }

    [System.IO.Directory]::CreateDirectory($script:InstallPath) | Out-Null
    Write-TestPayload -Path $script:ManagedExecutablePath -Content 'TAMPERED'
    Assert-Equal $null (Get-OOShutUpPath) 'path lookup does not return an untrusted executable'
    Assert-Equal $false (Test-OOShutUpInstalled) 'an untrusted executable does not count as installed'
    Remove-Item -LiteralPath $script:ManagedExecutablePath -Force

    $emptyFixture = Join-Path $script:FixtureRoot 'empty.exe'
    [System.IO.File]::WriteAllBytes($emptyFixture, [byte[]]@())
    $signatureCallsBeforeEmpty = $script:SignatureCallCount
    Assert-Equal $false (Test-OOShutUpPayload -Path $emptyFixture) 'empty payload is rejected'
    Assert-Equal $signatureCallsBeforeEmpty $script:SignatureCallCount 'empty payload is rejected before signature lookup'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`tExample setting`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $script:LogMessages.Clear()
    $approvedProfile = Get-OOShutUpApprovedProfile
    Assert-True ($null -ne $approvedProfile) 'a well-formed profile matching its digest and allowlist is accepted'
    Assert-Equal $script:OOShutUpProfilePath ([string]$approvedProfile.Path) 'the approved profile path is absolute and package-relative'
    Assert-Equal 2 $approvedProfile.SettingCount 'the approved profile identity records its setting count'
    Assert-Equal '3.6.1135.0' ([string]$approvedProfile.OOShutUpVersionMinimum) 'an exact profile release policy is retained in the approved profile'
    Assert-Equal '3.6.1135.0' ([string]$approvedProfile.OOShutUpVersionMaximum) 'equal version bounds represent exact-release approval'
    Assert-Equal 'Test category' ([string]$approvedProfile.Settings[0].Category) 'validated impact metadata is included in the interactive profile summary'
    $profileLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($profileLog -notmatch 'S003|M003|Example setting') 'profile identity logging omits setting IDs and values'

    $script:PromptMessages.Clear()
    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive YES accepts the reviewed profile'
    $summaryText = [string]::Join("`n", [string[]]$script:PromptMessages.ToArray())
    Assert-True ($summaryText.Contains($approvedProfile.Sha256)) 'interactive summary displays the exact profile digest'
    Assert-True ($summaryText.Contains('O&O executable file version reviewed: 3.6.1135.0')) 'interactive summary displays the exact reviewed O&O release'
    Assert-True ($summaryText.Contains('Test category') -and $summaryText.Contains('S003') -and $summaryText.Contains('Test fixture impact description.')) 'interactive summary displays reviewed category, ID, and impact text'

    $newerFixture = Join-Path $script:FixtureRoot 'newer-unreviewed.exe'
    Write-TestPayload -Path $newerFixture -Content 'NEWER_VERSION'
    $script:LogMessages.Clear()
    Assert-Equal $false (Test-OOShutUpPayload -Path $newerFixture -ReviewedMinimumVersion $approvedProfile.OOShutUpVersionMinimum -ReviewedMaximumVersion $approvedProfile.OOShutUpVersionMaximum) 'a newer supported O&O release outside the exact approved range is rejected'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpProfileVersionFailureMessage)) 'an unreviewed release reports the profile compatibility recovery instruction'
    $startedBeforeUnreviewedRelease = $script:StartedPaths.Count
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $newerFixture).State 'an unreviewed newer executable cannot apply the approved profile'
    Assert-Equal $startedBeforeUnreviewedRelease $script:StartedPaths.Count 'an unreviewed newer release is rejected before O&O launch'

    $olderFixture = Join-Path $script:FixtureRoot 'unsupported-older.exe'
    Write-TestPayload -Path $olderFixture -Content 'OLD_VERSION'
    Assert-Equal $false (Test-OOShutUpPayload -Path $olderFixture -ReviewedMinimumVersion $approvedProfile.OOShutUpVersionMinimum -ReviewedMaximumVersion $approvedProfile.OOShutUpVersionMaximum) 'an unsupported older O&O release is rejected'

    Write-ApprovedProfileFixture `
        -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" `
        -AllowedSettingIds @('S003', 'M003') `
        -OOShutUpMinimumVersion '3.6.1136.0' `
        -OOShutUpMaximumVersion '3.6.1136.0'
    $newerExactProfile = Get-OOShutUpApprovedProfile
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $newerExactProfile -NonInteractive -ApprovedProfileSha256 $newerExactProfile.Sha256 -ApprovedProfileManifestSha256 $newerExactProfile.ManifestSha256) 'a newer exact-release profile is explicitly approved for the older-release rejection fixture'
    $startedBeforeUnreviewedOlderRelease = $script:StartedPaths.Count
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'an older supported executable outside the approved range cannot apply the profile'
    Assert-Equal $startedBeforeUnreviewedOlderRelease $script:StartedPaths.Count 'an unreviewed older release is rejected before O&O launch'

    Write-ApprovedProfileFixture `
        -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" `
        -AllowedSettingIds @('S003', 'M003') `
        -OOShutUpMinimumVersion '3.6.1135.0' `
        -OOShutUpMaximumVersion '3.6.1136.0'
    $rangeProfile = Get-OOShutUpApprovedProfile
    Assert-Equal '3.6.1135.0 through 3.6.1136.0' ([string]$rangeProfile.OOShutUpVersionRange) 'the reviewed compatibility range is retained for approval'
    Assert-Equal $true (Test-OOShutUpPayload -Path $newerFixture -ReviewedMinimumVersion $rangeProfile.OOShutUpVersionMinimum -ReviewedMaximumVersion $rangeProfile.OOShutUpVersionMaximum) 'a newer executable passes when it lies within the explicitly reviewed range'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $approvedProfile = Get-OOShutUpApprovedProfile
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive -ApprovedProfileSha256 $approvedProfile.Sha256 -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256) 'restored exact-release profile is approved with its new manifest digest'

    Write-ApprovedProfileFixture -Content "# Mixed-scope fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003') -CurrentUserSettingIds @('M003')
    $mixedScopeProfile = Get-OOShutUpApprovedProfile
    $currentUserSetting = @($mixedScopeProfile.Settings | Where-Object { $_.Id -eq 'M003' })[0]
    Assert-Equal 'current-user' ([string]$currentUserSetting.Scope) 'reviewed current-user scope is retained in the approved profile'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $mixedScopeProfile -NonInteractive -ApprovedProfileSha256 $mixedScopeProfile.Sha256 -ApprovedProfileManifestSha256 $mixedScopeProfile.ManifestSha256) 'mixed profile can be approved for an interactive run with both digests'
    $startedBeforeSystemScopeBlock = $script:StartedPaths.Count
    $script:LogMessages.Clear()
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply).State 'SYSTEM reapplication rejects a profile containing current-user settings'
    Assert-Equal $startedBeforeSystemScopeBlock $script:StartedPaths.Count 'current-user scope is rejected before O&O launches'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpCurrentUserScopeFailureMessage)) 'current-user scope rejection explains the safe fallback'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`tExample setting`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $approvedProfile = Get-OOShutUpApprovedProfile
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'restored machine-only fixture requires fresh approval'

    $startedBeforeDecline = $script:StartedPaths.Count
    $script:PromptDecision = 'NO'
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive NO rejects the reviewed profile'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a rejected profile cannot be applied'
    Assert-Equal $startedBeforeDecline $script:StartedPaths.Count 'rejecting approval causes zero O&O launches'

    $script:PromptDecision = 'CANCEL'
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive CANCEL stops approval'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a canceled profile cannot be applied'
    Assert-Equal $startedBeforeDecline $script:StartedPaths.Count 'canceling approval causes zero O&O launches'

    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive -ApprovedProfileSha256 $approvedProfile.Sha256 -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256) 'unattended approval accepts the exact profile and manifest digests'
    Assert-Equal 'Unchanged' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'the exact unattended approval permits the mocked O&O launch'
    Assert-Equal ($startedBeforeDecline + 1) $script:StartedPaths.Count 'an exact approved digest launches O&O once'
    Assert-Equal 'Unchanged' $script:OOShutUpLastRunOutcome.State 'O&O exit code zero records an O&O-reported unchanged state'
    Assert-Equal 'Unavailable' $script:OOShutUpLastRunOutcome.ReadbackState 'exit code zero does not imply independent settings readback'
    Assert-Equal 'Proceed' $script:OOShutUpLastRunOutcome.NextStep 'the recognized unchanged report is represented separately from readback'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('O&O reported that no setting changes were needed')) 'O&O exit code zero reports a no-change result'

    $startedBeforeMissingApproval = $script:StartedPaths.Count
    $script:LogMessages.Clear()
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive) 'unattended mode rejects a missing approval digest'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpApprovalFailureMessage)) 'missing unattended approval reports the required fallback'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a missing unattended approval cannot be applied'
    Assert-Equal $startedBeforeMissingApproval $script:StartedPaths.Count 'missing approval causes zero O&O launches'
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256) 'unattended mode rejects a stale profile digest'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a stale unattended approval cannot be applied'
    Assert-Equal $startedBeforeMissingApproval $script:StartedPaths.Count 'stale approval causes zero O&O launches'

    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive approval can be accepted after rejected attempts'
    $script:MockExitCode = 30
    Assert-Equal 'Changed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'O&O exit code 30 is returned as an explicit changed outcome'
    Assert-Equal 'Changed' $script:OOShutUpLastRunOutcome.State 'O&O exit code 30 records an O&O-reported changed state'
    Assert-Equal 'Unavailable' $script:OOShutUpLastRunOutcome.ReadbackState 'exit code 30 does not imply independent settings readback'
    Assert-Equal 'Proceed' $script:OOShutUpLastRunOutcome.NextStep 'a recognized O&O report can proceed without being called verified'
    $script:MockExitCode = 0

    $script:MockExitCode = 23
    $script:LogMessages.Clear()
    Assert-Equal 'Unknown' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'an unrecognized nonzero exit does not proceed as a known settings outcome'
    Assert-Equal 'Unknown' $script:OOShutUpLastRunOutcome.State 'an unrecognized nonzero exit remains unknown after possible partial changes'
    Assert-Equal 23 $script:OOShutUpLastRunOutcome.ExitCode 'an unknown outcome preserves the exact native exit code'
    Assert-Equal 'Unavailable' $script:OOShutUpLastRunOutcome.ReadbackState 'an unrecognized exit is not treated as readback'
    $unknownExitLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($unknownExitLog.Contains('O&O Undo History') -and $unknownExitLog.Contains('current settings') -and $unknownExitLog.Contains('no rollback was performed')) 'an unrecognized exit gives current-state and Undo History guidance without claiming rollback'
    Assert-True (-not $unknownExitLog.Contains('Verified')) 'an unknown exit is never labelled Verified'
    $script:MockExitCode = 0

    $oldApprovedDigest = $approvedProfile.Sha256
    Write-ApprovedProfileFixture -Content "# Changed reviewed fixture`r`nS003`t-`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $changedProfile = Get-OOShutUpApprovedProfile
    Assert-True ($changedProfile.Sha256 -cne $oldApprovedDigest) 'a changed profile receives a new digest'
    $startedBeforeChangedDigest = $script:StartedPaths.Count
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $changedProfile -NonInteractive -ApprovedProfileSha256 $oldApprovedDigest -ApprovedProfileManifestSha256 $changedProfile.ManifestSha256) 'an approval for the prior profile does not authorize a changed profile'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'the changed profile cannot launch with its prior digest approval'
    Assert-Equal $startedBeforeChangedDigest $script:StartedPaths.Count 'a changed profile with a stale digest causes zero O&O launches'

    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $changedProfile) 'the changed profile receives explicit approval before manifest pinning checks'
    $startedBeforeChangedManifest = $script:StartedPaths.Count
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $changedProfile -NonInteractive -ApprovedProfileSha256 $changedProfile.Sha256 -ApprovedProfileManifestSha256 ('0' * 64)) 'unattended mode rejects a stale manifest digest'
    Assert-Equal $startedBeforeChangedManifest $script:StartedPaths.Count 'a stale manifest digest does not launch O&O'

    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $changedProfile) 'manifest-tamper scenario starts from an interactive approval'
    $tamperedManifest = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $tamperedManifest.settingDetails[0].scopeEvidence = 'Changed after approval.'
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($tamperedManifest | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $startedBeforeManifestTamper = $script:StartedPaths.Count
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'scope metadata changed after approval cannot reach O&O'
    Assert-Equal $startedBeforeManifestTamper $script:StartedPaths.Count 'manifest tampering causes zero O&O process launches'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`tExample setting`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $approvedProfile = Get-OOShutUpApprovedProfile
    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'the restored original profile requires a fresh interactive approval'

    $unrelatedWorkingDirectory = Join-Path $script:FixtureRoot 'unrelated caller directory'
    [System.IO.Directory]::CreateDirectory($unrelatedWorkingDirectory) | Out-Null
    $reviewedProfileBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfilePath)
    $reviewedManifestBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfileManifestPath)
    $script:OpenedProfilePaths.Clear()
    $script:OpenedProfileBytes.Clear()
    $script:OpenedManifestBytes.Clear()
    $script:SwapProfileBeforeMockOpen = $true
    $script:MockReplacementProfileContent = "# Unapproved swapped fixture`r`nS003`t-`r`nM003;+`r`n"
    $script:MockReplacementManifestContent = '{"schemaVersion":3,"tampered":true}'
    Push-Location -LiteralPath $unrelatedWorkingDirectory
    try {
        $script:StartedPaths.Clear()
        $script:StartedArguments.Clear()
        $script:StartedWorkingDirectories.Clear()
        $script:LogMessages.Clear()
        Assert-Equal 'Unchanged' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'configuration returns an O&O-reported unchanged state from an unrelated current directory'
    }
    finally {
        $script:SwapProfileBeforeMockOpen = $false
        Pop-Location
    }
    Assert-Equal 1 $script:StartedPaths.Count 'a valid profile launches the verified executable exactly once'
    Assert-Equal 1 $script:OpenedProfilePaths.Count 'the mocked O&O child opens exactly one profile path'
    $snapshotProfilePath = [string]$script:OpenedProfilePaths[0]
    $snapshotDirectory = Split-Path -Parent $snapshotProfilePath
    Assert-True ($snapshotProfilePath.StartsWith([System.IO.Path]::GetFullPath($script:InstallPath), [System.StringComparison]::OrdinalIgnoreCase)) 'the O&O child receives a profile path under the managed machine directory'
    Assert-True ($snapshotProfilePath -ne $script:OOShutUpProfilePath) 'the O&O child never receives the mutable source profile path'
    Assert-Equal ('/quiet /nosrp "{0}"' -f $snapshotProfilePath) ([string]$script:StartedArguments[0]) 'O&O receives only the quoted absolute snapshot path'
    Assert-Equal $snapshotDirectory ([string]$script:StartedWorkingDirectories[0]) 'O&O starts in the protected snapshot directory'
    Assert-Equal $approvedProfile.Sha256 (Get-OOShutUpProfileDigest -Bytes ([byte[]]$script:OpenedProfileBytes[0])) 'the child opens the exact approved profile bytes after the source swap'
    Assert-Equal $approvedProfile.ManifestSha256 (Get-OOShutUpProfileDigest -Bytes ([byte[]]$script:OpenedManifestBytes[0])) 'the snapshot manifest retains its exact reviewed digest at the child boundary'
    Assert-Equal ([Convert]::ToBase64String($reviewedProfileBytes)) ([Convert]::ToBase64String([byte[]]$script:OpenedProfileBytes[0])) 'unapproved profile bytes never reach the child boundary'
    Assert-Equal ([Convert]::ToBase64String($reviewedManifestBytes)) ([Convert]::ToBase64String([byte[]]$script:OpenedManifestBytes[0])) 'unapproved manifest bytes never reach the child boundary'
    Assert-Equal $script:MockReplacementProfileContent ([System.IO.File]::ReadAllText($script:OOShutUpProfilePath)) 'the barrier-controlled mock replaced the source after snapshot validation'
    Assert-Equal $false (Test-Path -LiteralPath $snapshotDirectory) 'a completed O&O child releases its per-run snapshot'
    $launchProfileLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($launchProfileLog.Contains("version test-1.0 (SHA-256 $($approvedProfile.Sha256)")) 'profile version and digest are recorded at launch'
    Assert-True ($launchProfileLog -notmatch 'S003|M003|Example setting') 'launch logging omits setting IDs and values'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $approvedProfile = Get-OOShutUpApprovedProfile
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'the restored source profile receives fresh approval after the race fixture'

    $approvedDigestBeforeSnapshotWriteFailure = $script:OOShutUpApprovedDigest
    $approvedManifestDigestBeforeSnapshotWriteFailure = $script:OOShutUpApprovedManifestDigest
    $writeSnapshotFunction = (Get-Command -Name Write-OOShutUpProfileSnapshotFile -CommandType Function).ScriptBlock
    Set-Item -Path Function:\Write-OOShutUpProfileSnapshotFile -Value {
        param([string]$Path, [byte[]]$Bytes)
        throw 'Mocked access denied while writing the protected snapshot.'
    }
    $startedBeforeSnapshotWriteFailure = $script:StartedPaths.Count
    $script:LogMessages.Clear()
    try {
        Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a denied protected snapshot write fails closed'
    }
    finally {
        Set-Item -Path Function:\Write-OOShutUpProfileSnapshotFile -Value $writeSnapshotFunction
    }
    Assert-Equal $startedBeforeSnapshotWriteFailure $script:StartedPaths.Count 'a denied snapshot write never launches O&O'
    Assert-Equal $approvedDigestBeforeSnapshotWriteFailure $script:OOShutUpApprovedDigest 'a denied snapshot write does not widen or replace the approved profile digest'
    Assert-Equal $approvedManifestDigestBeforeSnapshotWriteFailure $script:OOShutUpApprovedManifestDigest 'a denied snapshot write does not widen or replace the approved manifest digest'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpProfileSnapshotFailureMessage)) 'a denied snapshot write reports the exact no-launch fallback'

    $originalSourceProfileBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfilePath)
    $originalSourceManifestBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfileManifestPath)
    $getApprovedProfileFunction = (Get-Command -Name Get-OOShutUpApprovedProfile -CommandType Function).ScriptBlock
    $script:SnapshotFailureProfile = $approvedProfile
    Set-Item -Path Function:\Get-OOShutUpApprovedProfile -Value {
        param([switch]$RequireMachineScope)
        return $script:SnapshotFailureProfile
    }
    try {
        foreach ($sourceDigestCase in @(
            @{ Name = 'source profile'; Path = $script:OOShutUpProfilePath },
            @{ Name = 'source manifest'; Path = $script:OOShutUpProfileManifestPath }
        )) {
            [System.IO.File]::WriteAllBytes([string]$sourceDigestCase.Path, [System.Text.Encoding]::UTF8.GetBytes('changed after approval'))
            $startProcessCallsBeforeSourceDigestFailure = $script:StartProcessCallCount
            $script:LogMessages.Clear()
            try {
                Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State "a changed $($sourceDigestCase.Name) digest fails closed"
            }
            finally {
                if ($sourceDigestCase.Name -eq 'source profile') {
                    [System.IO.File]::WriteAllBytes($script:OOShutUpProfilePath, $originalSourceProfileBytes)
                }
                else {
                    [System.IO.File]::WriteAllBytes($script:OOShutUpProfileManifestPath, $originalSourceManifestBytes)
                }
            }
            Assert-Equal $startProcessCallsBeforeSourceDigestFailure $script:StartProcessCallCount "a changed $($sourceDigestCase.Name) digest never calls Start-Process"
            Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpProfileSnapshotFailureMessage)) "a changed $($sourceDigestCase.Name) digest reports the exact no-launch fallback"
        }
    }
    finally {
        Set-Item -Path Function:\Get-OOShutUpApprovedProfile -Value $getApprovedProfileFunction
        [System.IO.File]::WriteAllBytes($script:OOShutUpProfilePath, $originalSourceProfileBytes)
        [System.IO.File]::WriteAllBytes($script:OOShutUpProfileManifestPath, $originalSourceManifestBytes)
        $script:SnapshotFailureProfile = $null
    }

    $writeSnapshotFunction = (Get-Command -Name Write-OOShutUpProfileSnapshotFile -CommandType Function).ScriptBlock
    $script:RealWriteSnapshotFunction = $writeSnapshotFunction
    foreach ($copiedDigestFileName in @('ooshutup10.cfg', 'ooshutup10.approval.json')) {
        $script:InjectedSnapshotDigestFileName = $copiedDigestFileName
        Set-Item -Path Function:\Write-OOShutUpProfileSnapshotFile -Value {
            param([string]$Path, [byte[]]$Bytes)
            if ([string]::Equals([System.IO.Path]::GetFileName($Path), $script:InjectedSnapshotDigestFileName, [System.StringComparison]::OrdinalIgnoreCase)) {
                $Bytes = [System.Text.Encoding]::UTF8.GetBytes('tampered snapshot copy')
            }
            & $script:RealWriteSnapshotFunction -Path $Path -Bytes $Bytes
        }
        $startProcessCallsBeforeCopiedDigestFailure = $script:StartProcessCallCount
        $script:LogMessages.Clear()
        try {
            Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State "a copied $copiedDigestFileName digest mismatch fails closed"
        }
        finally {
            Set-Item -Path Function:\Write-OOShutUpProfileSnapshotFile -Value $writeSnapshotFunction
        }
        Assert-Equal $startProcessCallsBeforeCopiedDigestFailure $script:StartProcessCallCount "a copied $copiedDigestFileName digest mismatch never calls Start-Process"
        Assert-Equal $approvedDigestBeforeSnapshotWriteFailure $script:OOShutUpApprovedDigest "a copied $copiedDigestFileName mismatch leaves the approved profile digest unchanged"
        Assert-Equal $approvedManifestDigestBeforeSnapshotWriteFailure $script:OOShutUpApprovedManifestDigest "a copied $copiedDigestFileName mismatch leaves the approved manifest digest unchanged"
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpProfileSnapshotFailureMessage)) "a copied $copiedDigestFileName mismatch reports the exact no-launch fallback"
    }
    $script:RealWriteSnapshotFunction = $null
    $script:InjectedSnapshotDigestFileName = $null

    $aclFailureCases = @(
        @{ Name = 'managed O&O directory ACL inspection'; Target = 'ManagedDirectory' },
        @{ Name = 'snapshot directory ACL'; Target = 'Directory' },
        @{ Name = 'snapshot profile ACL'; Target = 'Profile' },
        @{ Name = 'snapshot manifest ACL'; Target = 'Manifest' }
    )
    foreach ($aclFailureCase in $aclFailureCases) {
        $script:MockAclFailurePath = $null
        $script:MockSnapshotAclDeniedTarget = $null
        if ($aclFailureCase.Target -eq 'ManagedDirectory') {
            $script:MockAclFailurePath = [System.IO.Path]::GetFullPath($script:InstallPath)
        }
        else {
            $script:MockSnapshotAclDeniedTarget = [string]$aclFailureCase.Target
        }
        $startProcessCallsBeforeAclFailure = $script:StartProcessCallCount
        $script:LogMessages.Clear()
        try {
            Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State "$($aclFailureCase.Name) fails closed"
        }
        finally {
            $script:MockAclFailurePath = $null
            $script:MockSnapshotAclDeniedTarget = $null
        }
        Assert-Equal $startProcessCallsBeforeAclFailure $script:StartProcessCallCount "$($aclFailureCase.Name) never calls Start-Process"
        Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:InstallPath -Filter 'ProfileSnapshot-*' -Directory).Count "$($aclFailureCase.Name) cleans its known-invalid snapshot"
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpProfileSnapshotFailureMessage)) "$($aclFailureCase.Name) reports the exact no-launch fallback"
    }

    $removeSnapshotFunction = (Get-Command -Name Remove-OOShutUpProfileSnapshot -CommandType Function).ScriptBlock
    $script:MockStartProcessReturnsNull = $true
    $startProcessCallsBeforeMissingResult = $script:StartProcessCallCount
    $workingDirectoriesBeforeMissingResult = $script:StartedWorkingDirectories.Count
    try {
        Assert-Equal 'Unknown' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'an absent Start-Process result leaves the settings outcome unknown'
    }
    finally {
        $script:MockStartProcessReturnsNull = $false
    }
    Assert-Equal 'Unknown' $script:OOShutUpLastRunOutcome.State 'an absent process object is an unknown settings outcome'
    Assert-Equal $null $script:OOShutUpLastRunOutcome.ExitCode 'an absent process object does not fabricate an exit code'
    Assert-Equal ($startProcessCallsBeforeMissingResult + 1) $script:StartProcessCallCount 'the absent-result scenario reaches the process boundary once'
    $missingResultSnapshotDirectory = [string]$script:StartedWorkingDirectories[$workingDirectoriesBeforeMissingResult]
    Assert-True (Test-Path -LiteralPath $missingResultSnapshotDirectory -PathType Container) 'the absent-result scenario retains the protected profile snapshot'
    Assert-True ($script:OOShutUpLastRunOutcome.Message.Contains('Undo History') -and $script:OOShutUpLastRunOutcome.Message.Contains('current settings')) 'an absent process result gives actionable state-inspection guidance'
    & $removeSnapshotFunction -Directory $missingResultSnapshotDirectory

    $script:MockStartProcessFailure = $true
    $startProcessCallsBeforeUncertainLaunch = $script:StartProcessCallCount
    $workingDirectoriesBeforeUncertainLaunch = $script:StartedWorkingDirectories.Count
    $script:LogMessages.Clear()
    try {
        Assert-Equal 'Unknown' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'an unverified process-start outcome reports an unknown result'
    }
    finally {
        $script:MockStartProcessFailure = $false
    }
    Assert-Equal ($startProcessCallsBeforeUncertainLaunch + 1) $script:StartProcessCallCount 'the unverified process-start scenario reaches Start-Process once'
    $uncertainSnapshotDirectory = [string]$script:StartedWorkingDirectories[$workingDirectoriesBeforeUncertainLaunch]
    Assert-True (Test-Path -LiteralPath $uncertainSnapshotDirectory -PathType Container) 'an uncertain launch retains its protected snapshot for diagnosis'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains("The O&O launch outcome could not be verified. The protected profile snapshot was retained at '$uncertainSnapshotDirectory'")) 'an uncertain launch reports the retained snapshot path'
    Assert-Equal 'Unknown' $script:OOShutUpLastRunOutcome.State 'an ambiguous process launch is not classified as a clean failure'
    Assert-Equal 'Unavailable' $script:OOShutUpLastRunOutcome.ReadbackState 'an ambiguous process launch has no independent state readback'
    Assert-True (-not $script:OOShutUpLastRunOutcome.Message.Contains('Verified')) 'an unknown process result is never labelled Verified'
    & $removeSnapshotFunction -Directory $uncertainSnapshotDirectory

    Set-Item -Path Function:\Remove-OOShutUpProfileSnapshot -Value {
        param([string]$Directory)
        throw 'Mocked snapshot cleanup access denied.'
    }
    $script:OpenedProfilePaths.Clear()
    $startedBeforeSnapshotCleanupFailure = $script:StartedPaths.Count
    $script:LogMessages.Clear()
    $script:MockExitCode = 30
    try {
        $cleanupFailureOutcome = Invoke-OOShutUpConfiguration -ExecutablePath $validFixture
        Assert-Equal 'Changed' $cleanupFailureOutcome.State 'a snapshot cleanup failure preserves the explicit O&O-reported changed state'
    }
    finally {
        $script:MockExitCode = 0
        Set-Item -Path Function:\Remove-OOShutUpProfileSnapshot -Value $removeSnapshotFunction
    }
    Assert-Equal ($startedBeforeSnapshotCleanupFailure + 1) $script:StartedPaths.Count 'cleanup failure occurs only after the child received a protected snapshot'
    Assert-Equal 'Changed' $script:OOShutUpLastRunOutcome.State 'a cleanup failure preserves O&O exit code 30 as a vendor-reported changed result'
    Assert-Equal 'Stop' $script:OOShutUpLastRunOutcome.NextStep 'snapshot cleanup failure stops later installer work despite the reported changed state'
    Assert-Equal $false ([string]$script:OpenedProfilePaths[0] -eq $script:OOShutUpProfilePath) 'cleanup failure never passes the source profile path to O&O'
    $orphanedSnapshotDirectory = Split-Path -Parent ([string]$script:OpenedProfilePaths[0])
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains("remove '$orphanedSnapshotDirectory' manually")) 'cleanup failure reports the exact protected snapshot path for manual recovery'
    & $removeSnapshotFunction -Directory $orphanedSnapshotDirectory

    $script:MockChildWaitScenario = 'timeout-boundary-race'
    $script:MockExitCode = 30
    $script:LogMessages.Clear()
    $boundaryStartCount = $script:StartProcessCallCount
    $boundaryResult = Invoke-OOShutUpConfiguration -ExecutablePath $validFixture
    $boundaryProcess = $script:StartedProcessObjects[$script:StartedProcessObjects.Count - 1]
    $boundarySnapshotDirectory = [string]$script:StartedWorkingDirectories[$script:StartedWorkingDirectories.Count - 1]
    Assert-Equal 'Unknown' $boundaryResult.State 'a child that exits only during the timeout-boundary check remains unknown'
    Assert-Equal 'Unknown' $script:OOShutUpLastRunOutcome.State 'the timeout-boundary race preserves an unknown settings outcome'
    Assert-Equal ($boundaryStartCount + 1) $script:StartProcessCallCount 'the timeout-boundary scenario starts only its own O&O child'
    Assert-Equal $true $boundaryProcess.HasExited 'the timeout-boundary mock confirms the exact child stopped after the wait expired'
    Assert-Equal 0 $boundaryProcess.KillCalls 'timeout handling does not terminate the child by PID or process name'
    Assert-True (Test-Path -LiteralPath $boundarySnapshotDirectory -PathType Container) 'the timeout-boundary result retains its snapshot for inspection'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('O&O exceeded its 900-second wait limit')) 'the timeout-boundary result is explicitly reported as incomplete'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('No success was reported')) 'the timeout-boundary result never reports success'

    $startsBeforeBoundaryRetry = $script:StartProcessCallCount
    $script:MockChildWaitScenario = 'fast-exit'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a retained timeout snapshot blocks retry until inspection'
    Assert-Equal $startsBeforeBoundaryRetry $script:StartProcessCallCount 'a retained timeout snapshot prevents a second O&O child from starting'
    Assert-True (Test-Path -LiteralPath $boundarySnapshotDirectory -PathType Container) 'a blocked retry preserves the original timeout snapshot'
    & $removeSnapshotFunction -Directory $boundarySnapshotDirectory

    $script:MockChildWaitScenario = 'still-running'
    $script:LogMessages.Clear()
    $runningStartCount = $script:StartProcessCallCount
    $runningResult = Invoke-OOShutUpConfiguration -ExecutablePath $validFixture
    $runningProcess = $script:StartedProcessObjects[$script:StartedProcessObjects.Count - 1]
    $runningSnapshotDirectory = [string]$script:StartedWorkingDirectories[$script:StartedWorkingDirectories.Count - 1]
    Assert-Equal 'Unknown' $runningResult.State 'a still-running O&O child returns an unknown timeout result'
    Assert-Equal 'Unknown' $script:OOShutUpLastRunOutcome.State 'a still-running child cannot report a known settings state'
    Assert-Equal ($runningStartCount + 1) $script:StartProcessCallCount 'the still-running scenario starts exactly one owned O&O child'
    Assert-Equal $false $runningProcess.HasExited 'the controllable child remains alive after timeout'
    Assert-Equal 0 $runningProcess.KillCalls 'the timeout does not kill any process or process tree'
    Assert-True (Test-Path -LiteralPath $runningSnapshotDirectory -PathType Container) 'the still-running child retains the profile it may still read'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('may still be reading its snapshot')) 'the timeout log identifies that the child may still read the retained snapshot'

    $startsBeforeRunningRetry = $script:StartProcessCallCount
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a still-running child requires inspection before retry'
    Assert-Equal $startsBeforeRunningRetry $script:StartProcessCallCount 'the retained still-running snapshot prevents a queued reapplication'
    Assert-Equal $false $runningProcess.HasExited 'the blocked retry does not affect the original child'
    $runningProcess.HasExited = $true
    & $removeSnapshotFunction -Directory $runningSnapshotDirectory
    Assert-Equal $false (Test-Path -LiteralPath $runningSnapshotDirectory) 'the fixture snapshot is removed only after the mock confirms the child stopped'
    $script:MockExitCode = 0
    $script:MockChildWaitScenario = 'fast-exit'

    $script:StartedPaths.Clear()
    $script:StartedArguments.Clear()
    $script:StartedWorkingDirectories.Clear()
    $invalidProfileCases = @(
        @{ Name = 'malformed'; Content = "S003 +`r`n"; Allowed = @('S003') },
        @{ Name = 'unknown id'; Content = "X999`t+`r`n"; Allowed = @('S003') },
        @{ Name = 'duplicate id'; Content = "S003`t+`r`nS003`t-`r`n"; Allowed = @('S003', 'M003') },
        @{ Name = 'empty'; Content = "# comments only`r`n`r`n"; Allowed = @('S003') }
    )
    foreach ($case in $invalidProfileCases) {
        Write-ApprovedProfileFixture -Content $case.Content -AllowedSettingIds $case.Allowed
        $startedBefore = $script:StartedPaths.Count
        Assert-Equal $null (Get-OOShutUpApprovedProfile) "$($case.Name) profile is rejected"
        Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State "$($case.Name) profile prevents configuration"
        Assert-Equal $startedBefore $script:StartedPaths.Count "$($case.Name) profile causes no process launch"
        Assert-InvalidProfileStopsBeforeTaskMutation -Name $case.Name
    }

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $manifestWithoutVersionRange = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $manifestWithoutVersionRange.PSObject.Properties.Remove('ooshutupVersionRange')
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($manifestWithoutVersionRange | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $startedBeforeMissingVersionRange = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a schema 4 manifest without a reviewed O&O version range is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a missing O&O version range prevents configuration'
    Assert-Equal $startedBeforeMissingVersionRange $script:StartedPaths.Count 'a missing O&O version range causes no process launch'
    $lookupBeforeMissingVersionRangeTask = $script:ScheduledTaskLookupCount
    $registrationBeforeMissingVersionRangeTask = $script:TaskRegistrationCalls.Count
    Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath (Join-Path $script:FixtureRoot 'unused-missing-version-range.ps1') -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 ('0' * 64)) 'a missing O&O version range blocks SYSTEM task registration'
    Assert-Equal $lookupBeforeMissingVersionRangeTask $script:ScheduledTaskLookupCount 'a missing O&O version range blocks before task lookup'
    Assert-Equal $registrationBeforeMissingVersionRangeTask $script:TaskRegistrationCalls.Count 'a missing O&O version range blocks before task registration'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $legacyVersionManifest = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $legacyVersionManifest.schemaVersion = 3
    $legacyVersionManifest.PSObject.Properties.Remove('ooshutupVersionRange')
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($legacyVersionManifest | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a prior schema is rejected instead of silently widened'

    $legacyVersionManifest.schemaVersion = 4
    $legacyVersionManifest.profileVersion = 'test-migrated-1.0'
    $legacyVersionManifest | Add-Member -MemberType NoteProperty -Name ooshutupVersionRange -Value ([pscustomobject]@{
        minimum = '3.6.1135.0'
        maximum = '3.6.1135.0'
    })
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($legacyVersionManifest | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $migratedProfile = Get-OOShutUpApprovedProfile
    Assert-True ($null -ne $migratedProfile) 'an owner-migrated prior profile is accepted only after adding an exact reviewed release'
    Assert-Equal 'test-migrated-1.0' ([string]$migratedProfile.Version) 'deliberate migration records a new reviewed profile version'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $migratedProfile -NonInteractive -ApprovedProfileSha256 $migratedProfile.Sha256 -ApprovedProfileManifestSha256 $migratedProfile.ManifestSha256) 'deliberate migration requires approval of the migrated manifest digest'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $legacyManifest = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $legacyManifest.schemaVersion = 2
    $legacyManifest.settingDetails[0].PSObject.Properties.Remove('scope')
    $legacyManifest.settingDetails[0].PSObject.Properties.Remove('scopeEvidence')
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($legacyManifest | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $script:LogMessages.Clear()
    $startedBeforeUnknownScope = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile -RequireMachineScope) 'legacy metadata without explicit scope is rejected by the SYSTEM scope gate'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpUnknownScopeFailureMessage)) 'unknown scope logs the required exact fallback'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply).State 'unknown scope prevents SYSTEM reapplication'
    Assert-Equal $startedBeforeUnknownScope $script:StartedPaths.Count 'unknown scope causes zero O&O process launches'
    $lookupBeforeAbsentScopeTask = $script:ScheduledTaskLookupCount
    $registrationBeforeAbsentScopeTask = $script:TaskRegistrationCalls.Count
    Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath (Join-Path $script:FixtureRoot 'unused-absent-scope.ps1') -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 ('0' * 64)) 'an absent scope claim blocks SYSTEM task registration'
    Assert-Equal $lookupBeforeAbsentScopeTask $script:ScheduledTaskLookupCount 'an absent scope claim blocks before task lookup'
    Assert-Equal $registrationBeforeAbsentScopeTask $script:TaskRegistrationCalls.Count 'an absent scope claim blocks before task registration'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $missingScopeEvidenceManifest = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $missingScopeEvidenceManifest.settingDetails[0].PSObject.Properties.Remove('scopeEvidence')
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($missingScopeEvidenceManifest | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $lookupBeforeMissingEvidenceTask = $script:ScheduledTaskLookupCount
    $registrationBeforeMissingEvidenceTask = $script:TaskRegistrationCalls.Count
    Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath (Join-Path $script:FixtureRoot 'unused-missing-evidence.ps1') -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 ('0' * 64)) 'missing scope evidence blocks SYSTEM task registration'
    Assert-Equal $lookupBeforeMissingEvidenceTask $script:ScheduledTaskLookupCount 'missing scope evidence blocks before task lookup'
    Assert-Equal $registrationBeforeMissingEvidenceTask $script:TaskRegistrationCalls.Count 'missing scope evidence blocks before task registration'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $unknownScopeManifest = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $unknownScopeManifest.settingDetails[0].scope = 'unknown'
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($unknownScopeManifest | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $script:LogMessages.Clear()
    Assert-Equal $null (Get-OOShutUpApprovedProfile -RequireMachineScope) 'an unsupported scope value is rejected'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpUnknownScopeFailureMessage)) 'an unsupported scope value logs the required fallback'
    $startedBeforeUnsupportedScope = $script:StartedPaths.Count
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply).State 'an unsupported scope prevents SYSTEM reapplication'
    Assert-Equal $startedBeforeUnsupportedScope $script:StartedPaths.Count 'an unsupported scope causes zero O&O process launches'
    $lookupBeforeUnsupportedScopeTask = $script:ScheduledTaskLookupCount
    $registrationBeforeUnsupportedScopeTask = $script:TaskRegistrationCalls.Count
    Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath (Join-Path $script:FixtureRoot 'unused-unsupported-scope.ps1') -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 ('0' * 64)) 'an unsupported scope blocks SYSTEM task registration'
    Assert-Equal $lookupBeforeUnsupportedScopeTask $script:ScheduledTaskLookupCount 'an unsupported scope blocks before task lookup'
    Assert-Equal $registrationBeforeUnsupportedScopeTask $script:TaskRegistrationCalls.Count 'an unsupported scope blocks before task registration'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $manifestWithMismatchedImpactState = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $manifestWithMismatchedImpactState.settingDetails[0].state = '-'
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($manifestWithMismatchedImpactState | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $startedBeforeMismatchedImpact = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'impact metadata with a mismatched setting state is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'mismatched impact metadata prevents configuration'
    Assert-Equal $startedBeforeMismatchedImpact $script:StartedPaths.Count 'mismatched impact metadata causes zero O&O launches'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    $manifestWithoutImpactDetails = Get-Content -LiteralPath $script:OOShutUpProfileManifestPath -Raw | ConvertFrom-Json
    $manifestWithoutImpactDetails.PSObject.Properties.Remove('settingDetails')
    [System.IO.File]::WriteAllText(
        $script:OOShutUpProfileManifestPath,
        ($manifestWithoutImpactDetails | ConvertTo-Json -Depth 4),
        [System.Text.Encoding]::UTF8
    )
    $startedBeforeMissingImpact = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a manifest without reviewed impact details is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'missing impact details prevent configuration'
    Assert-Equal $startedBeforeMissingImpact $script:StartedPaths.Count 'missing impact details cause zero O&O launches'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    [System.IO.File]::AppendAllText($script:OOShutUpProfilePath, "M003`t-`r`n")
    $startedBeforeDigestMismatch = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a digest-mismatched profile is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a digest mismatch prevents configuration'
    Assert-Equal $startedBeforeDigestMismatch $script:StartedPaths.Count 'a digest mismatch causes no process launch'
    Assert-InvalidProfileStopsBeforeTaskMutation -Name 'digest-mismatched'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    [System.IO.File]::WriteAllText($script:OOShutUpProfileManifestPath, '{ invalid json')
    $startedBeforeMalformedManifest = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a malformed approval manifest is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a malformed manifest prevents configuration'
    Assert-Equal $startedBeforeMalformedManifest $script:StartedPaths.Count 'a malformed manifest causes no process launch'
    Assert-InvalidProfileStopsBeforeTaskMutation -Name 'malformed manifest'

    Remove-Item -LiteralPath $script:OOShutUpProfilePath -Force
    $startedBeforeMissingProfile = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a missing profile is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a missing profile prevents configuration'
    Assert-Equal $startedBeforeMissingProfile $script:StartedPaths.Count 'a missing profile causes no process launch'
    Assert-InvalidProfileStopsBeforeTaskMutation -Name 'missing profile'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    Remove-Item -LiteralPath $script:OOShutUpProfileManifestPath -Force
    $startedBeforeMissingManifest = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a missing approval manifest is rejected'
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture).State 'a missing manifest prevents configuration'
    Assert-Equal $startedBeforeMissingManifest $script:StartedPaths.Count 'a missing manifest causes no process launch'
    Assert-InvalidProfileStopsBeforeTaskMutation -Name 'missing manifest'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')

    $oldTemp = $env:TEMP
    $env:TEMP = $script:FixtureRoot
    $script:ScheduledTaskLookupCount = 0
    $savedTrustedPowerShellResolver = (Get-Command -Name Get-TrustedWindowsPowerShellExecutable -CommandType Function).ScriptBlock
    Set-Item -Path Function:\Get-TrustedWindowsPowerShellExecutable -Value {
        $script:MockPowerShellResolverCallCount++
        if ($script:MockPowerShellResolverMode -eq 'Missing') {
            return $null
        }
        $candidate = Get-WindowsPowerShellExecutablePaths -WindowsDirectory $script:MockWindowsPowerShellDirectory -Is64BitOperatingSystem $script:MockPowerShellContextIs64BitOperatingSystem -Is64BitProcess $script:MockPowerShellContextIs64BitProcess
        if ($null -eq $candidate) {
            return $null
        }
        $script:MockLastPowerShellInspectionPath = [string]$candidate.InspectionPath
        if ($script:MockPowerShellResolverMode -eq 'Unexpected' -or
            ($script:MockPowerShellResolverMode -eq 'ChangeAfterResolution' -and $script:MockPowerShellResolverCallCount -ge 2)) {
            return [string]$script:MockPowerShellResolverChangedPath
        }
        return [string]$candidate.TaskPath
    }
    try {
        [System.IO.Directory]::CreateDirectory($script:InstallPath) | Out-Null
        Write-TestPayload -Path $script:ManagedExecutablePath -Content 'VALID'
        $taskScriptDirectory = Join-Path $script:FixtureRoot 'package with spaces'
        [System.IO.Directory]::CreateDirectory($taskScriptDirectory) | Out-Null
        $taskScriptPath = Join-Path $taskScriptDirectory 'Install-OOShutUp10.ps1'
        [System.IO.File]::WriteAllText($taskScriptPath, '# ACL target fixture')
        $systemFullControl = New-MockAclRule -Sid 'S-1-5-18' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl)
        $administratorsFullControl = New-MockAclRule -Sid 'S-1-5-32-544' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl)
        $script:MockAclByPath = @{
            '__default__' = New-MockAcl -Rules @($systemFullControl, $administratorsFullControl)
        }
        $script:MockAclFailurePath = $null
        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        Write-ApprovedProfileFixture -Content "# Mixed-scope fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003') -CurrentUserSettingIds @('M003')
        $mixedTaskProfile = Get-OOShutUpApprovedProfile
        Assert-Equal $true (Confirm-OOShutUpProfile -Profile $mixedTaskProfile -NonInteractive -ApprovedProfileSha256 $mixedTaskProfile.Sha256 -ApprovedProfileManifestSha256 $mixedTaskProfile.ManifestSha256) 'the mixed profile receives an exact test approval before task registration is checked'
        $lookupBeforeScopeBlock = $script:ScheduledTaskLookupCount
        $registrationsBeforeScopeBlock = $script:TaskRegistrationCalls.Count
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $mixedTaskProfile.Sha256 -ApprovedProfileManifestSha256 $mixedTaskProfile.ManifestSha256) 'a mixed-scope profile cannot create a SYSTEM task'
        Assert-Equal $lookupBeforeScopeBlock $script:ScheduledTaskLookupCount 'scope rejection does not inspect or replace existing tasks'
        Assert-Equal $registrationsBeforeScopeBlock $script:TaskRegistrationCalls.Count 'scope rejection does not register a candidate task'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpCurrentUserScopeFailureMessage)) 'task registration reports why current-user settings are excluded'

        Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
        $taskProfile = Get-OOShutUpApprovedProfile
        Assert-Equal $false (Confirm-OOShutUpProfile -Profile $taskProfile -NonInteractive) 'task approval rejects a missing digest'
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256) 'task registration rejects a missing manifest approval'
        Assert-Equal 0 $script:ScheduledTaskLookupCount 'missing task approval does not inspect or change existing tasks'
        Assert-Equal 0 $script:TaskRegistrationCalls.Count 'missing task approval does not register a candidate task'
        Assert-Equal $true (Confirm-OOShutUpProfile -Profile $taskProfile -NonInteractive -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'task registration uses the explicitly approved profile and manifest digests'

        $lookupsBeforeMissingPayload = $script:ScheduledTaskLookupCount
        $registrationsBeforeMissingPayload = $script:TaskRegistrationCalls.Count
        Remove-Item -LiteralPath $script:ManagedExecutablePath -Force
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'task registration refuses to pin a missing managed O&O executable'
        Assert-Equal $lookupsBeforeMissingPayload $script:ScheduledTaskLookupCount 'a missing managed executable does not inspect existing tasks'
        Assert-Equal $registrationsBeforeMissingPayload $script:TaskRegistrationCalls.Count 'a missing managed executable does not register or replace tasks'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpManagedPathFailureMessage)) 'a missing managed executable reports its recovery instruction'
        Write-TestPayload -Path $script:ManagedExecutablePath -Content 'VALID'

        $lookupsBeforeUnreviewedTaskVersion = $script:ScheduledTaskLookupCount
        $registrationsBeforeUnreviewedTaskVersion = $script:TaskRegistrationCalls.Count
        Write-TestPayload -Path $script:ManagedExecutablePath -Content 'NEWER_VERSION'
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'task creation refuses an O&O executable newer than the reviewed exact release'
        Assert-Equal $lookupsBeforeUnreviewedTaskVersion $script:ScheduledTaskLookupCount 'an unreviewed task executable blocks before task lookup'
        Assert-Equal $registrationsBeforeUnreviewedTaskVersion $script:TaskRegistrationCalls.Count 'an unreviewed task executable blocks before task registration'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpProfileVersionFailureMessage)) 'task creation reports the unreviewed O&O version rejection'
        Write-TestPayload -Path $script:ManagedExecutablePath -Content 'VALID'

        $registrationsBeforeMissingFile = $script:TaskRegistrationCalls.Count
        Remove-Item -LiteralPath $taskScriptPath -Force
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a missing SYSTEM task target fails closed'
        Assert-Equal $registrationsBeforeMissingFile $script:TaskRegistrationCalls.Count 'a missing task target never registers a candidate'
        [System.IO.File]::WriteAllText($taskScriptPath, '# ACL target fixture')

        $taskParentPath = [System.IO.Path]::GetFullPath($taskScriptDirectory)
        $untrustedWrite = New-MockAclRule -Sid 'S-1-5-11' -Rights ([System.Security.AccessControl.FileSystemRights]::Modify)
        $script:MockAclByPath[$taskParentPath] = New-MockAcl -Rules @($systemFullControl, $administratorsFullControl, $untrustedWrite)
        $registrationsBeforeWritableParent = $script:TaskRegistrationCalls.Count
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a standard-user-writable parent fails closed'
        Assert-Equal $registrationsBeforeWritableParent $script:TaskRegistrationCalls.Count 'a writable parent never registers a candidate task'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpTaskAclFailureMessage)) 'a writable parent reports the required ACL fallback'

        $script:MockAclByPath.Remove($taskParentPath) | Out-Null
        $taskTargetPath = [System.IO.Path]::GetFullPath($taskScriptPath)
        $script:MockAclByPath[$taskTargetPath] = New-MockAcl -Rules @($systemFullControl, $administratorsFullControl, $untrustedWrite)
        $registrationsBeforeWritableFile = $script:TaskRegistrationCalls.Count
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a standard-user-writable script file fails closed'
        Assert-Equal $registrationsBeforeWritableFile $script:TaskRegistrationCalls.Count 'a writable script file never registers a candidate task'
        $script:MockAclByPath.Remove($taskTargetPath) | Out-Null

        $driveRootPath = [System.IO.Path]::GetPathRoot($taskScriptPath)
        $script:MockAclByPath[[System.IO.Path]::GetFullPath($driveRootPath)] = New-MockAcl -Rules @(
            $systemFullControl,
            $administratorsFullControl,
            (New-MockAclRule -Sid 'S-1-5-11' -Rights ([System.Security.AccessControl.FileSystemRights]::FullControl) -PropagationFlags ([System.Security.AccessControl.PropagationFlags]::InheritOnly))
        )
        Assert-Equal $true (Test-SystemTaskTargetAcl -ScriptPath $taskScriptPath) 'an inherit-only write entry on an ancestor does not count as access to that ancestor'
        $script:MockAclByPath.Remove([System.IO.Path]::GetFullPath($driveRootPath)) | Out-Null

        $script:MockAclFailurePath = [System.IO.Path]::GetFullPath($taskScriptPath)
        $registrationsBeforeDeniedAcl = $script:TaskRegistrationCalls.Count
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'an ACL inspection error fails closed'
        Assert-Equal $registrationsBeforeDeniedAcl $script:TaskRegistrationCalls.Count 'an ACL inspection error never registers a candidate task'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpTaskAclFailureMessage)) 'ACL inspection failure reports the required fallback'
        $script:MockAclFailurePath = $null

        Assert-Equal $true (Test-SystemTaskTargetAcl -ScriptPath $taskScriptPath) 'a safe temporary ACL fixture permits the target'
        foreach ($mockedHost in @(
            @{ Name = '64-bit'; Is64BitProcess = $true },
            @{ Name = '32-bit'; Is64BitProcess = $false }
        )) {
            if ($mockedHost.Name -eq '64-bit') {
                $legacyTaskXmlPath = Join-Path $script:FixtureRoot 'OOShutUp10Task_Debug.xml'
                $legacyTaskXmlBytes = [System.Text.Encoding]::UTF8.GetBytes('pre-existing diagnostic file must remain unchanged')
                [System.IO.File]::WriteAllBytes($legacyTaskXmlPath, $legacyTaskXmlBytes)
                $legacyTaskXmlFingerprint = [System.Convert]::ToBase64String($legacyTaskXmlBytes)
                # A file cannot contain a child path, so this also makes an attempted TEMP export fail.
                $env:TEMP = $legacyTaskXmlPath
            }
            $script:MockPowerShellContextIs64BitOperatingSystem = $true
            $script:MockPowerShellContextIs64BitProcess = [bool]$mockedHost.Is64BitProcess
            $script:MockPowerShellResolverMode = 'Valid'
            $script:MockPowerShellResolverCallCount = 0
            $expectedMockedPaths = Get-WindowsPowerShellExecutablePaths -WindowsDirectory $script:MockWindowsPowerShellDirectory -Is64BitOperatingSystem $true -Is64BitProcess ([bool]$mockedHost.Is64BitProcess)
            Assert-Equal $true (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) "a safe ACL fixture registers the task under a mocked $($mockedHost.Name) host context"
            [xml]$taskXml = [string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition
            Assert-Equal $legacyTaskXmlFingerprint ([System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($legacyTaskXmlPath))) 'task setup leaves the pre-existing legacy diagnostic target byte-identical'
            Assert-Equal ([string]$expectedMockedPaths.TaskPath) ([string]$taskXml.Task.Actions.Exec.Command) "the $($mockedHost.Name) host stores the canonical trusted Windows PowerShell path"
            Assert-Equal ([string]$expectedMockedPaths.InspectionPath) ([string]$script:MockLastPowerShellInspectionPath) "the $($mockedHost.Name) host validates the architecture-correct system path"
        }
        Assert-True ([System.IO.File]::Exists($legacyTaskXmlPath)) 'repeated task setup requests do not replace or remove the old diagnostic target'
        $diagnosticArtifacts = @(Get-ChildItem -LiteralPath $script:FixtureRoot -Filter 'OOShutUp10Task_Debug*.xml' -File)
        Assert-Equal 1 $diagnosticArtifacts.Count 'task setup creates no additional diagnostic XML containing action details or approved digests'
        Assert-Equal 'Ready' ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].State) 'task registration remains verified and enabled when TEMP cannot accept a child write'
        $env:TEMP = $script:FixtureRoot
        Assert-True ($script:TaskRegistrationCalls.Count -ge 2) 'candidate and final registration use mocked Task Scheduler commands'
        $taskWorkingDirectory = $taskXml.Task.Actions.Exec.WorkingDirectory
        Assert-Equal (Split-Path -Parent $taskScriptPath) ([string]$taskWorkingDirectory) 'scheduled task uses the script package as its working directory'
        Assert-Equal 'S-1-5-18' ([string]$taskXml.Task.Principals.Principal.UserId) 'scheduled task runs under LocalSystem'
        Assert-Equal $expectedNativeInterpreterPath ([string]$taskXml.Task.Actions.Exec.Command) 'the task command is a canonical absolute Windows PowerShell path'
        $taskArguments = [string]$taskXml.Task.Actions.Exec.Arguments
        Assert-True ($taskArguments.Contains("-File `"$taskScriptPath`"")) 'scheduled task invokes the script from the same package'
        Assert-True ($taskArguments -match '-ReapplyOnly') 'scheduled task preserves ReapplyOnly mode'
        Assert-True ($taskArguments -match '-SkipRelocate') 'scheduled task preserves the no-relocation switch'
        Assert-True ($taskArguments -match '-NonInteractive') 'scheduled task uses the unattended approval path'
        Assert-True ($taskArguments.Contains("-ApprovedOOSExecutablePath `"$script:ManagedExecutablePath`"")) 'scheduled task pins the exact absolute managed O&O executable path'
        Assert-True ($taskArguments.Contains([string]$taskProfile.Sha256)) 'scheduled task carries the approved profile digest'
        Assert-True ($taskArguments.Contains([string]$taskProfile.ManifestSha256)) 'scheduled task carries the scope manifest digest'
        $taskLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
        Assert-True ($taskLog.Contains('S-1-5-18') -and $taskLog.Contains($taskProfile.Sha256) -and $taskLog.Contains($taskProfile.ManifestSha256)) 'task logging records principal and digests without setting values'
        Assert-True ($taskLog.Contains('installation-success event 19') -and $taskLog.Contains('queued sequentially')) 'task logging records the success-only serialized reapply policy'
        $trigger = $taskXml.Task.Triggers.EventTrigger
        Assert-Equal 1 $taskXml.Task.Triggers.ChildNodes.Count 'scheduled task has no startup or periodic reconciliation trigger'
        $triggerSubscriptionText = [string]$trigger.Subscription
        Assert-Equal 'true' ([string]$trigger.Enabled) 'scheduled task preserves its enabled Windows Update event trigger'
        Assert-Equal 'PT2M' ([string]$trigger.Delay) 'scheduled task preserves its two-minute update delay'
        Assert-True ($triggerSubscriptionText.Contains('WindowsUpdateClient') -and $triggerSubscriptionText.Contains('EventID=19') -and -not $triggerSubscriptionText.Contains('EventID=20') -and -not $triggerSubscriptionText.Contains('EventID=43')) 'scheduled task triggers only on Windows Update installation success'
        Assert-Equal 'Queue' ([string]$taskXml.Task.Settings.MultipleInstancesPolicy) 'scheduled task queues successful-update triggers behind an active reapply'
        Assert-Equal 'PT30M' ([string]$taskXml.Task.Settings.ExecutionTimeLimit) 'scheduled task has the explicit thirty-minute execution limit'

        $verifiedCurrentTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $taskXml.OuterXml
        $taskVerifierParameters = @{
            Task = $verifiedCurrentTask
            TaskPath = '\myTech.Today\'
            TaskName = 'Install-OOShutUp10'
            TaskDescription = [string]$taskXml.Task.RegistrationInfo.Description
            ScriptPath = $taskScriptPath
            ApprovedProfileSha256 = $taskProfile.Sha256
            ApprovedProfileManifestSha256 = $taskProfile.ManifestSha256
            ApprovedOOSExecutablePath = $script:ManagedExecutablePath
            TrustedPowerShellPath = $expectedNativeInterpreterPath
        }
        Assert-Equal $true (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification checks the canonical interpreter, success trigger, queue policy, and execution limit'

        $extraTriggerDocument = New-Object System.Xml.XmlDocument
        $extraTriggerDocument.LoadXml($taskXml.OuterXml)
        $extraTriggerParent = $extraTriggerDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Triggers']")
        $secondTrigger = $extraTriggerDocument.ImportNode($extraTriggerParent.FirstChild, $true)
        [void]$extraTriggerParent.AppendChild($secondTrigger)
        $taskVerifierParameters.Task = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $extraTriggerDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification rejects unapproved additional or reconciliation triggers'

        $timespanTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $taskXml.OuterXml
        $timespanTask.Settings.ExecutionTimeLimit = [TimeSpan]::FromMinutes(30)
        $taskVerifierParameters.Task = $timespanTask
        Assert-Equal $true (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification accepts the ScheduledTasks TimeSpan representation of the configured limit'
        $taskVerifierParameters.Task = $verifiedCurrentTask

        $bareInterpreterDocument = New-Object System.Xml.XmlDocument
        $bareInterpreterDocument.LoadXml($taskXml.OuterXml)
        $bareInterpreterDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Command']").InnerText = 'PowerShell.exe'
        $taskVerifierParameters.Task = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $bareInterpreterDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification rejects a bare PowerShell executable name'

        $userWritableInterpreterDirectory = Join-Path $script:FixtureRoot 'user writable interpreter'
        [System.IO.Directory]::CreateDirectory($userWritableInterpreterDirectory) | Out-Null
        $userWritableInterpreterPath = Join-Path $userWritableInterpreterDirectory 'powershell.exe'
        [System.IO.File]::WriteAllText($userWritableInterpreterPath, 'untrusted interpreter')
        $script:MockAclByPath[[System.IO.Path]::GetFullPath($userWritableInterpreterPath)] = New-MockAcl -Rules @(
            $systemFullControl,
            $administratorsFullControl,
            (New-MockAclRule -Sid 'S-1-5-11' -Rights ([System.Security.AccessControl.FileSystemRights]::Modify))
        )
        Assert-Equal $false (Test-SystemTaskTargetAcl -ScriptPath $userWritableInterpreterPath) 'the user-writable interpreter fixture fails the SYSTEM target ACL check'
        $userWritableInterpreterDocument = New-Object System.Xml.XmlDocument
        $userWritableInterpreterDocument.LoadXml($taskXml.OuterXml)
        $userWritableInterpreterDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Command']").InnerText = [System.IO.Path]::GetFullPath($userWritableInterpreterPath)
        $taskVerifierParameters.Task = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $userWritableInterpreterDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification rejects a user-writable interpreter path'
        $script:MockAclByPath.Remove([System.IO.Path]::GetFullPath($userWritableInterpreterPath)) | Out-Null

        $script:MockPowerShellResolverMode = 'Unexpected'
        $script:MockPowerShellResolverChangedPath = $redirectedInterpreterPath
        $taskVerifierParameters.Task = $verifiedCurrentTask
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification rejects an interpreter path that changed after resolution'
        $script:MockPowerShellResolverMode = 'Valid'
        $script:MockPowerShellResolverChangedPath = $null

        $priorTaskXmlBeforeInterpreterFailures = [string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition
        $registrationsBeforeInterpreterFailure = $script:TaskRegistrationCalls.Count
        $unregistrationsBeforeInterpreterFailure = $script:TaskUnregistrationCalls.Count
        $lookupsBeforeInterpreterFailure = $script:ScheduledTaskLookupCount
        $script:LogMessages.Clear()
        $script:MockPowerShellResolverMode = 'Missing'
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a missing trusted interpreter prevents task setup'
        Assert-Equal $registrationsBeforeInterpreterFailure $script:TaskRegistrationCalls.Count 'a missing trusted interpreter does not register a validation task'
        Assert-Equal $unregistrationsBeforeInterpreterFailure $script:TaskUnregistrationCalls.Count 'a missing trusted interpreter does not remove or replace the existing task'
        Assert-Equal $lookupsBeforeInterpreterFailure $script:ScheduledTaskLookupCount 'a missing trusted interpreter fails before Task Scheduler inspection'
        Assert-Equal $priorTaskXmlBeforeInterpreterFailures ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'a missing trusted interpreter preserves the existing approved task'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains($script:OOShutUpTrustedPowerShellFailureMessage)) 'a missing trusted interpreter reports the exact no-change fallback'

        $script:MockPowerShellResolverMode = 'ChangeAfterResolution'
        $script:MockPowerShellResolverCallCount = 0
        $script:MockPowerShellResolverChangedPath = $redirectedInterpreterPath
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'an interpreter path change after resolution prevents task staging'
        Assert-Equal $registrationsBeforeInterpreterFailure $script:TaskRegistrationCalls.Count 'an interpreter path change is detected before validation-task registration'
        Assert-Equal $unregistrationsBeforeInterpreterFailure $script:TaskUnregistrationCalls.Count 'an interpreter path change performs no task rollback or removal'
        Assert-Equal $priorTaskXmlBeforeInterpreterFailures ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'an interpreter path change preserves the existing approved task'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains($script:OOShutUpTrustedPowerShellFailureMessage)) 'an interpreter path change reports the exact no-change fallback'

        $script:MockPowerShellResolverMode = 'Unexpected'
        $script:MockPowerShellResolverCallCount = 0
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a redirected interpreter resolution result is rejected before task staging'
        Assert-Equal $registrationsBeforeInterpreterFailure $script:TaskRegistrationCalls.Count 'a redirected interpreter is never registered as a validation task'
        Assert-Equal $priorTaskXmlBeforeInterpreterFailures ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'a redirected interpreter result leaves the approved task unchanged'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains($script:OOShutUpTrustedPowerShellFailureMessage)) 'a redirected interpreter result reports the exact no-change fallback'
        $script:MockPowerShellResolverMode = 'Valid'
        $script:MockPowerShellResolverChangedPath = $null

        $ignoreNewDocument = New-Object System.Xml.XmlDocument
        $ignoreNewDocument.LoadXml($taskXml.OuterXml)
        $ignoreNewDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='MultipleInstancesPolicy']").InnerText = 'IgnoreNew'
        $ignoreNewTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $ignoreNewDocument.OuterXml
        $taskVerifierParameters.Task = $ignoreNewTask
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'IgnoreNew is rejected as an unsafe current task definition'

        $unboundedTaskDocument = New-Object System.Xml.XmlDocument
        $unboundedTaskDocument.LoadXml($taskXml.OuterXml)
        $unboundedLimitNode = $unboundedTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='ExecutionTimeLimit']")
        [void]$unboundedLimitNode.ParentNode.RemoveChild($unboundedLimitNode)
        $taskVerifierParameters.Task = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $unboundedTaskDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification rejects a missing execution limit'

        $wrongLimitDocument = New-Object System.Xml.XmlDocument
        $wrongLimitDocument.LoadXml($taskXml.OuterXml)
        $wrongLimitDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='ExecutionTimeLimit']").InnerText = 'PT0S'
        $taskVerifierParameters.Task = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $wrongLimitDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'task verification rejects an unlimited execution limit'

        $startAndSuccessDocument = New-Object System.Xml.XmlDocument
        $startAndSuccessDocument.LoadXml($taskXml.OuterXml)
        $startAndSuccessDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Triggers']/*[local-name()='EventTrigger']/*[local-name()='Subscription']").InnerText = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''Microsoft-Windows-WindowsUpdateClient''] and (EventID=19 or EventID=43)]]</Select></Query></QueryList>'
        $startAndSuccessTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $startAndSuccessDocument.OuterXml
        $taskVerifierParameters.Task = $startAndSuccessTask
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition @taskVerifierParameters) 'the old start-and-success trigger is rejected as a current task definition'

        $noEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @())
        $successEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(19))
        $failureEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(20))
        $startEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(43))
        Assert-Equal 0 $noEventRuns.Count 'no Windows Update event leaves the task idle'
        Assert-Equal 1 $successEventRuns.Count 'one successful Windows Update event produces one reapply'
        Assert-Equal 19 $successEventRuns[0].EventId 'the scheduled reapply is tied to installation success'
        Assert-Equal 130 $successEventRuns[0].StartSeconds 'a successful event keeps its two-minute delay'
        Assert-Equal 0 $failureEventRuns.Count 'a Windows Update installation failure does not start or report a successful reapply'
        Assert-Equal 0 $startEventRuns.Count 'an installation-start event does not start an early reapply'

        $duplicateEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(19, 19))
        Assert-Equal 2 $duplicateEventRuns.Count 'two delivered success events each retain one bounded queued run'
        Assert-Equal 19 $duplicateEventRuns[0].EventId 'the first duplicate-event run uses the success trigger'
        Assert-Equal 19 $duplicateEventRuns[1].EventId 'the second duplicate-event run uses the success trigger'
        Assert-True ($duplicateEventRuns[1].StartSeconds -ge ($duplicateEventRuns[0].StartSeconds + 1500)) 'duplicate-event runs are queued without overlapping the bounded mutation wait and O&O child wait'

        foreach ($runDurationAfterSuccess in @(119, 120, 121)) {
            $activeRunEndSeconds = 10 + $runDurationAfterSuccess
            $completionRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(19) -ActiveRunEventId 43 -ActiveRunEndSeconds $activeRunEndSeconds)
            $expectedCompletionStartSeconds = [Math]::Max(130, $activeRunEndSeconds)
            Assert-Equal 1 $completionRuns.Count "a success event during an event-43 run at ${runDurationAfterSuccess}s retains one completion reapply"
            Assert-Equal 19 $completionRuns[0].EventId "the ${runDurationAfterSuccess}s follow-up is caused by the success event"
            Assert-Equal 43 $completionRuns[0].ActiveRunEventId "the ${runDurationAfterSuccess}s scenario holds the event-43 run open"
            Assert-Equal $expectedCompletionStartSeconds $completionRuns[0].StartSeconds "the ${runDurationAfterSuccess}s follow-up waits for both the PT2M delay and active run"
            Assert-True ($completionRuns[0].StartSeconds -ge $activeRunEndSeconds) "the ${runDurationAfterSuccess}s follow-up does not overlap the active O&O reapply"
        }

        $legacyTaskDocument = New-Object System.Xml.XmlDocument
        $legacyTaskDocument.LoadXml($taskXml.OuterXml)
        $legacyCommand = $legacyTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Command']")
        $legacyCommand.InnerText = 'PowerShell.exe'
        $legacyArguments = $legacyTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Arguments']")
        $legacyArguments.InnerText = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $taskScriptPath + '" -ReapplyOnly -SkipRelocate'
        $legacyWorkingDirectory = $legacyTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='WorkingDirectory']")
        if ($null -ne $legacyWorkingDirectory) {
            [void]$legacyWorkingDirectory.ParentNode.RemoveChild($legacyWorkingDirectory)
        }
        $oldTaskXml = $legacyTaskDocument.OuterXml
        $oldTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $oldTaskXml

        $previousApprovedTaskDocument = New-Object System.Xml.XmlDocument
        $previousApprovedTaskDocument.LoadXml($taskXml.OuterXml)
        $previousApprovedCommand = $previousApprovedTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Command']")
        $previousApprovedCommand.InnerText = 'PowerShell.exe'
        $previousApprovedTaskArguments = $previousApprovedTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Arguments']")
        $previousApprovedTaskArguments.InnerText = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $taskScriptPath + '" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedProfileSha256 ' + $taskProfile.Sha256 + ' -ApprovedProfileManifestSha256 ' + $taskProfile.ManifestSha256
        $previousApprovedTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $previousApprovedTaskDocument.OuterXml
        Assert-Equal $true (Test-WindowsUpdateLegacyTaskDefinition -Task $previousApprovedTask -TaskPath '\myTech.Today\' -TaskName 'Install-OOShutUp10' -TaskDescription ([string]$taskXml.Task.RegistrationInfo.Description) -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'the previous approved task definition remains eligible for safe replacement with the pinned executable path'

        $unexpectedLegacyDocument = New-Object System.Xml.XmlDocument
        $unexpectedLegacyDocument.LoadXml($previousApprovedTaskDocument.OuterXml)
        $unexpectedLegacyDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Command']").InnerText = $redirectedInterpreterPath
        $unexpectedLegacyTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $unexpectedLegacyDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateLegacyTaskDefinition -Task $unexpectedLegacyTask -TaskPath '\myTech.Today\' -TaskName 'Install-OOShutUp10' -TaskDescription ([string]$taskXml.Task.RegistrationInfo.Description) -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'legacy migration rejects an unexpected interpreter path even when the task identity matches'

        $foreignTaskDocument = New-Object System.Xml.XmlDocument
        $foreignTaskDocument.LoadXml($taskXml.OuterXml)
        $foreignDescription = $foreignTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='RegistrationInfo']/*[local-name()='Description']")
        $foreignDescription.InnerText = 'Business backup automation'
        $foreignArguments = $foreignTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Arguments']")
        $foreignArguments.InnerText = '-NoProfile -File "C:\Business\Backup.ps1"'
        $foreignTaskXml = $foreignTaskDocument.OuterXml
        $foreignTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\Backup\' -Xml $foreignTaskXml
        $securityTaskDocument = New-Object System.Xml.XmlDocument
        $securityTaskDocument.LoadXml($foreignTaskXml)
        $securityDescription = $securityTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='RegistrationInfo']/*[local-name()='Description']")
        $securityDescription.InnerText = 'Security maintenance automation'
        $securityTaskXml = $securityTaskDocument.OuterXml
        $securityTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\Security\' -Xml $securityTaskXml

        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskLookupCalls.Clear()
        $ownedTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $taskXml.OuterXml
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $ownedTask
        $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
        $script:TaskRegistry['\Security\|Install-OOShutUp10'] = $securityTask
        $script:LogMessages.Clear()
        Assert-Equal $true (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'an owned task can be updated while foreign same-name tasks remain untouched'
        Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) 'the backup-folder task remains unchanged'
        Assert-Equal $securityTaskXml ([string]$script:TaskRegistry['\Security\|Install-OOShutUp10'].Definition) 'the security-folder task remains unchanged'
        $ownedTaskRemovals = @($script:TaskUnregistrationCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' })
        Assert-Equal 1 $ownedTaskRemovals.Count 'only the verified owned task is removed during replacement'
        Assert-Equal 'Install-OOShutUp10' ([string]$ownedTaskRemovals[0].TaskName) 'the exact owned task name is passed for removal'
        Assert-Equal '\myTech.Today\' ([string]$ownedTaskRemovals[0].TaskPath) 'the exact owned task path is passed for removal'
        Assert-Equal 0 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskPath -ne '\myTech.Today\' }).Count 'no task outside the owned folder is passed for removal'
        Assert-True (@($script:TaskLookupCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' -and $_.TaskPath -eq '\myTech.Today\' }).Count -gt 0) 'the replacement lookup includes the exact owned folder'
        $ownershipLog = @($script:LogMessages.ToArray()) -join [Environment]::NewLine
        Assert-True ($ownershipLog.Contains("Removing verified owned scheduled task '\myTech.Today\Install-OOShutUp10'")) 'the exact owned path is logged before removal'
        Assert-True ($ownershipLog.Contains("Leaving same-named task '\Backup\Install-OOShutUp10' unchanged")) 'the foreign backup path is logged with an operator instruction'
        Assert-True ($ownershipLog.Contains("Leaving same-named task '\Security\Install-OOShutUp10' unchanged")) 'the foreign security path is logged with an operator instruction'

        $script:TaskRegistry.Clear()
        $script:TaskSecurityDescriptorByIdentity.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskSecurityDescriptorReadCalls.Clear()
        $script:TaskSecurityDescriptorSetCalls.Clear()
        $script:TaskEnabledCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
        $script:TaskSecurityDescriptorByIdentity['\myTech.Today\|Install-OOShutUp10'] = $script:OOShutUpTaskSecurityDescriptorSddl
        $script:MockTaskSecurityDescriptorScenario = 'Normal'
        Assert-Equal $true (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a prior task with the documented descriptor can be replaced'
        $stagedDescriptorSet = @($script:TaskSecurityDescriptorSetCalls | Where-Object { $_.TaskName -like 'Install-OOShutUp10-validation-*' })
        $stagedDescriptorRead = @($script:TaskSecurityDescriptorReadCalls | Where-Object { $_.TaskName -like 'Install-OOShutUp10-validation-*' })
        $finalDescriptorSet = @($script:TaskSecurityDescriptorSetCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' })
        $finalDescriptorRead = @($script:TaskSecurityDescriptorReadCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' })
        Assert-Equal 1 $stagedDescriptorSet.Count 'the disabled staging task receives the documented ACL'
        Assert-True ($stagedDescriptorRead.Count -ge 1) 'the disabled staging task descriptor is read back'
        Assert-Equal 1 $finalDescriptorSet.Count 'the disabled final task receives the documented ACL'
        Assert-True ($finalDescriptorRead.Count -ge 3) 'the final task descriptor is read back before and after enabling, in addition to the prior-task snapshot'
        Assert-Equal 1 @($script:TaskEnabledCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' -and $_.Enabled }).Count 'the verified final task is enabled only after descriptor checks'
        Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) 'descriptor verification leaves a same-named task outside the owned path unchanged'

        foreach ($descriptorFailure in @('ReadDeniedOnOwnedTask', 'MalformedOnOwnedTask')) {
            $script:TaskRegistry.Clear()
            $script:TaskSecurityDescriptorByIdentity.Clear()
            $script:TaskRegistrationCalls.Clear()
            $script:TaskUnregistrationCalls.Clear()
            $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
            $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
            $script:MockTaskSecurityDescriptorScenario = $descriptorFailure
            $script:LogMessages.Clear()
            Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) "$descriptorFailure on the prior task fails closed"
            Assert-Equal 0 $script:TaskRegistrationCalls.Count "$descriptorFailure does not register a candidate"
            Assert-Equal 0 $script:TaskUnregistrationCalls.Count "$descriptorFailure does not remove any task"
            Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) "$descriptorFailure preserves the existing owned task"
            Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) "$descriptorFailure preserves the same-named task outside the exact owned path"
            Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains($script:OOShutUpTaskSecurityDescriptorFailureMessage)) "$descriptorFailure uses the exact descriptor fallback"
        }

        foreach ($descriptorFailure in @('ReadDeniedOnStage', 'MalformedOnStage')) {
            $script:TaskRegistry.Clear()
            $script:TaskSecurityDescriptorByIdentity.Clear()
            $script:TaskRegistrationCalls.Clear()
            $script:TaskUnregistrationCalls.Clear()
            $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
            $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
            $script:TaskSecurityDescriptorByIdentity['\myTech.Today\|Install-OOShutUp10'] = $script:OOShutUpTaskSecurityDescriptorSddl
            $script:MockTaskSecurityDescriptorScenario = $descriptorFailure
            $script:MockTaskSecurityDescriptorReplacementInjected = $false
            $script:LogMessages.Clear()
            Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) "$descriptorFailure on the staging task fails closed"
            Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) "$descriptorFailure on staging leaves the prior task unchanged"
            Assert-Equal 1 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskPath -eq '\myTech.Today\' -and $_.TaskName -like 'Install-OOShutUp10-validation-*' }).Count "$descriptorFailure cleans only the exact disabled staging task"
            Assert-Equal 0 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskPath -ne '\myTech.Today\' }).Count "$descriptorFailure never unregisters a same-named task outside the owned folder"
            Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) "$descriptorFailure preserves the foreign same-named task"
            Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains($script:OOShutUpTaskSecurityDescriptorFailureMessage)) "$descriptorFailure on staging uses the exact descriptor fallback"
        }

        foreach ($descriptorFailure in @('UnexpectedOnFinal', 'UnprovenOnFinal')) {
            $script:TaskRegistry.Clear()
            $script:TaskSecurityDescriptorByIdentity.Clear()
            $script:TaskRegistrationCalls.Clear()
            $script:TaskUnregistrationCalls.Clear()
            $script:TaskEnabledCalls.Clear()
            $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
            $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
            $script:TaskSecurityDescriptorByIdentity['\myTech.Today\|Install-OOShutUp10'] = $script:OOShutUpTaskSecurityDescriptorSddl
            $script:MockTaskSecurityDescriptorScenario = $descriptorFailure
            $script:MockTaskSecurityDescriptorReplacementInjected = $false
            $script:LogMessages.Clear()
            $descriptorSetup = Invoke-OOShutUpTaskSetup -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256
            $descriptorOutcome = Get-OOShutUpInstallOutcome -ConfigurationOutcome $changedConfigurationOutcome -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $descriptorSetup
            Assert-Equal $false $descriptorSetup.Verified "$descriptorFailure cannot produce verified task setup"
            Assert-Equal 'Partial' $descriptorOutcome.Status "$descriptorFailure cannot report Complete"
            Assert-Equal 2 $descriptorOutcome.ExitCode "$descriptorFailure returns the existing partial exit code"
            Assert-True ($descriptorOutcome.Summary.Contains($script:OOShutUpPartialSetupMessage) -and $descriptorOutcome.Summary.Contains('O&O reported that settings changed') -and $descriptorOutcome.Summary.Contains('not independently read back')) "$descriptorFailure preserves the partial summary and vendor-reported settings detail"
            Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains($script:OOShutUpTaskSecurityDescriptorFailureMessage)) "$descriptorFailure uses the exact descriptor fallback"
            Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) "$descriptorFailure leaves the foreign same-named task unchanged"
            if ($descriptorFailure -eq 'UnexpectedOnFinal') {
                Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'a permissive final descriptor restores the exact prior task definition'
                Assert-Equal $script:OOShutUpTaskSecurityDescriptorSddl ([string]$script:TaskSecurityDescriptorByIdentity['\myTech.Today\|Install-OOShutUp10']) 'rollback restores the exact prior task descriptor'
                Assert-Equal 1 @($script:TaskEnabledCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' -and $_.Enabled }).Count 'rollback re-enables only the prior verified task'
                Assert-Equal 0 @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.sddl').Count 'verified rollback removes its protected descriptor backup'
            }
            else {
                Assert-Equal '-UnrecognizedAction' ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Actions[0].Arguments) 'rollback preserves a task whose ownership became ambiguous'
                Assert-Equal 1 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskPath -eq '\myTech.Today\' -and $_.TaskName -eq 'Install-OOShutUp10' }).Count 'ambiguous ownership is not unregistered after the original replacement removal'
                Assert-Equal 0 $script:TaskEnabledCalls.Count 'an unverified replacement remains disabled when ownership is ambiguous'
                Assert-Equal 1 @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.xml').Count 'ambiguous rollback retains the protected prior task XML'
                Assert-Equal 1 @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.sddl').Count 'ambiguous rollback retains the protected prior descriptor'
            }
        }
        $script:MockTaskSecurityDescriptorScenario = 'Normal'
        $script:TaskRegistry.Clear()
        $script:TaskSecurityDescriptorByIdentity.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskEnabledCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $taskXml.OuterXml
        $script:TaskSecurityDescriptorByIdentity['\myTech.Today\|Install-OOShutUp10'] = $script:OOShutUpTaskSecurityDescriptorSddl
        $script:MockTaskSecurityDescriptorScenario = 'UnexpectedOnFinal'
        $script:MockTaskSecurityDescriptorReplacementInjected = $false
        $sameXmlFailureSetup = Invoke-OOShutUpTaskSetup -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256
        Assert-Equal $false $sameXmlFailureSetup.Verified 'an unexpected descriptor is rejected when prior and replacement task XML are identical'
        Assert-Equal $taskXml.OuterXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'same-XML rollback restores the prior task definition'
        Assert-Equal $script:OOShutUpTaskSecurityDescriptorSddl ([string]$script:TaskSecurityDescriptorByIdentity['\myTech.Today\|Install-OOShutUp10']) 'same-XML rollback restores the prior descriptor instead of accepting the permissive replacement'
        Assert-Equal 2 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskPath -eq '\myTech.Today\' -and $_.TaskName -eq 'Install-OOShutUp10' }).Count 'same-XML rollback removes the old task and then the rejected replacement at the exact path'
        Assert-Equal 1 @($script:TaskEnabledCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' -and $_.Enabled }).Count 'same-XML rollback enables only the restored prior task'
        $script:MockTaskSecurityDescriptorScenario = 'Normal'
        $script:MockTaskSecurityDescriptorReplacementInjected = $false
        $script:TaskSecurityDescriptorByIdentity.Clear()
        foreach ($rollbackFile in @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.xml')) {
            Remove-Item -LiteralPath $rollbackFile.FullName -Force
        }
        foreach ($rollbackFile in @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.sddl')) {
            Remove-Item -LiteralPath $rollbackFile.FullName -Force
        }

        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:LogMessages.Clear()
        Assert-Equal $true (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'the exact approved legacy definition is migrated'
        Assert-True ($script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Actions[0].Arguments.Contains($taskProfile.Sha256)) 'the migrated task receives the current approved profile action'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains("Migrating approved legacy scheduled task '\myTech.Today\Install-OOShutUp10'")) 'the exact legacy migration path is logged'
        Assert-Equal 1 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskName -eq 'Install-OOShutUp10' -and $_.TaskPath -eq '\myTech.Today\' }).Count 'legacy migration removes only the exact approved path'

        $ambiguousDocument = New-Object System.Xml.XmlDocument
        $ambiguousDocument.LoadXml($taskXml.OuterXml)
        $ambiguousDescription = $ambiguousDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='RegistrationInfo']/*[local-name()='Description']")
        $ambiguousDescription.InnerText = 'Unknown task owner'
        $ambiguousXml = $ambiguousDocument.OuterXml
        $ambiguousTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $ambiguousXml
        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $ambiguousTask
        $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'an unrecognized task definition at the owned path stops replacement'
        Assert-Equal 0 $script:TaskRegistrationCalls.Count 'ambiguous ownership stops before candidate registration'
        Assert-Equal 0 $script:TaskUnregistrationCalls.Count 'ambiguous ownership stops before any task removal'
        Assert-Equal $ambiguousXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'the ambiguous owned-path task remains unchanged'
        Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) 'the foreign same-name task remains unchanged after refusal'
        $ambiguousLog = @($script:LogMessages.ToArray()) -join [Environment]::NewLine
        Assert-True ($ambiguousLog.Contains($script:OOShutUpTaskOwnershipFailureMessage) -and $ambiguousLog.Contains('\myTech.Today\Install-OOShutUp10')) 'ambiguous refusal reports the exact listed path and operator instruction'

        foreach ($failureMode in @('AccessDenied', 'SchedulerError', 'ModuleMissing')) {
            $script:TaskRegistry.Clear()
            $script:TaskRegistrationCalls.Clear()
            $script:TaskUnregistrationCalls.Clear()
            $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $ownedTask
            $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
            $script:TaskLookupFailure = $failureMode
            $script:LogMessages.Clear()
            Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) "$failureMode while inspecting tasks fails closed"
            Assert-Equal 0 $script:TaskRegistrationCalls.Count "$failureMode does not register a candidate"
            Assert-Equal 0 $script:TaskUnregistrationCalls.Count "$failureMode does not remove any task"
            Assert-Equal ([string]$ownedTask.Definition) ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) "$failureMode leaves the old owned task unchanged"
            Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) "$failureMode leaves the foreign task unchanged"
            $failureLog = @($script:LogMessages.ToArray()) -join [Environment]::NewLine
            if ($failureMode -eq 'AccessDenied') {
                Assert-True ($failureLog.Contains('Access denied while inspecting Task Scheduler')) 'access denied is reported distinctly'
            }
            elseif ($failureMode -eq 'ModuleMissing') {
                Assert-True ($failureLog.Contains('Task Scheduler could not inspect') -and $failureLog.Contains('ScheduledTasks module is not loaded')) 'a missing ScheduledTasks module is reported as an unavailable scheduler command'
            }
            else {
                Assert-True ($failureLog.Contains('Task Scheduler could not inspect')) 'scheduler errors are reported distinctly from access denied'
            }
        }

        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:TaskRegistry['\Backup\|Install-OOShutUp10'] = $foreignTask
        $script:TaskRegistry['\Security\|Install-OOShutUp10'] = $securityTask
        $script:TaskRegistrationFailure = 'AccessDenied'
        $script:LogMessages.Clear()
        $deniedRegistrationSetup = Invoke-OOShutUpTaskSetup -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256
        $deniedRegistrationOutcome = Get-OOShutUpInstallOutcome -ConfigurationOutcome $changedConfigurationOutcome -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $deniedRegistrationSetup
        Assert-Equal $false $deniedRegistrationSetup.Verified 'access denied during task registration is returned as unverified'
        Assert-Equal 2 $deniedRegistrationOutcome.ExitCode 'access denied during task registration yields the partial exit code'
        Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'denied candidate registration preserves the existing approved task'
        Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) 'denied registration preserves the unrelated backup task'
        Assert-Equal $securityTaskXml ([string]$script:TaskRegistry['\Security\|Install-OOShutUp10'].Definition) 'denied registration preserves the unrelated security task'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains('Access denied by Task Scheduler')) 'denied task registration gives the access recovery category'

        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:LogMessages.Clear()
        $retrySetup = Invoke-OOShutUpTaskSetup -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256
        $retryOutcome = Get-OOShutUpInstallOutcome -ConfigurationOutcome $changedConfigurationOutcome -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $retrySetup
        Assert-Equal $true $retrySetup.Verified 'retry verifies the exact approved task after scheduler access is restored'
        Assert-Equal 0 $retryOutcome.ExitCode 'verified task retry returns the success exit code'
        Assert-True ($script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Actions[0].Arguments.Contains($taskProfile.Sha256)) 'retry installs the approved profile digest at the owned path'
        Assert-Equal $foreignTaskXml ([string]$script:TaskRegistry['\Backup\|Install-OOShutUp10'].Definition) 'successful retry leaves the unrelated backup task unchanged'
        Assert-Equal $securityTaskXml ([string]$script:TaskRegistry['\Security\|Install-OOShutUp10'].Definition) 'successful retry leaves the unrelated security task unchanged'
        Assert-Equal 0 @($script:TaskUnregistrationCalls | Where-Object { $_.TaskPath -ne '\myTech.Today\' }).Count 'retry never unregisters a task outside the owned folder'

        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $registrationsBeforeStageFailure = $script:TaskRegistrationCalls.Count
        $unregistrationsBeforeStageFailure = $script:TaskUnregistrationCalls.Count
        $script:FailStageVerification = $true
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a candidate that fails verification is rejected before replacement'
        Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'candidate verification failure leaves the previous task unchanged'
        Assert-True ($script:TaskRegistry.ContainsKey('\myTech.Today\|Install-OOShutUp10')) 'candidate verification failure preserves the owned production task'
        Assert-Equal ($registrationsBeforeStageFailure + 1) $script:TaskRegistrationCalls.Count 'candidate failure registers only the disabled validation task'
        Assert-Equal ($unregistrationsBeforeStageFailure + 1) $script:TaskUnregistrationCalls.Count 'failed candidate cleanup removes only the disabled validation task'
        Assert-Equal 0 @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.xml').Count 'candidate verification failure removes its temporary rollback copy'
        $script:FailStageVerification = $false

        $script:TaskRegistry.Clear()
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:LogMessages.Clear()
        $script:FailStageQueuePolicyVerification = $true
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a candidate with IgnoreNew fails the queue-policy verification'
        Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'queue-policy verification failure preserves the prior task definition'
        Assert-Equal 1 $script:TaskRegistrationCalls.Count 'queue-policy verification failure registers only a disabled candidate'
        Assert-Equal 1 $script:TaskUnregistrationCalls.Count 'queue-policy verification failure removes only the rejected candidate'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains('trigger, queue, and execution-time settings')) 'queue-policy verification failure reports the exact rejected task criteria'
        $script:FailStageQueuePolicyVerification = $false

        $script:TaskRegistry.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:TaskRegistrationCalls.Clear()
        $script:TaskUnregistrationCalls.Clear()
        $script:LogMessages.Clear()
        $script:FailStageExecutionTimeLimitVerification = $true
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a candidate without the required bounded execution time is rejected'
        Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'an invalid execution limit preserves the prior task definition'
        Assert-Equal 1 $script:TaskRegistrationCalls.Count 'execution-limit verification failure registers only a disabled candidate'
        Assert-Equal 1 $script:TaskUnregistrationCalls.Count 'execution-limit verification failure removes only the rejected candidate'
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains('trigger, queue, and execution-time settings')) 'execution-limit verification failure reports the exact rejected task criteria'
        $script:FailStageExecutionTimeLimitVerification = $false

        $script:TaskRegistry.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:FailTaskVerification = $true
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a failed replacement verification reports failure'
        Assert-True ($script:TaskRegistry.ContainsKey('\myTech.Today\|Install-OOShutUp10')) 'a failed replacement restores the previous task definition'
        Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'rollback restores the previously exported task XML'
        Assert-Equal $oldTaskXml ([string]$script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'].Definition) 'a failed replacement removes the candidate and restores the approved legacy task'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('The previous task definition was restored')) 'rollback reports that the previous task was restored'
        $script:FailTaskVerification = $false

        $script:TaskRegistry.Clear()
        $script:TaskRegistry['\myTech.Today\|Install-OOShutUp10'] = $oldTask
        $script:FailTaskVerification = $true
        $script:FailTaskRestore = $true
        $restoreFailureDocument = New-Object System.Xml.XmlDocument
        $restoreFailureDocument.LoadXml($oldTaskXml)
        $restoreFailureDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']").InnerText = 'false'
        $script:FailTaskRestoreXml = $restoreFailureDocument.OuterXml
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a failed automatic rollback reports failure'
        $retainedRollbackFiles = @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.xml')
        Assert-Equal 1 $retainedRollbackFiles.Count 'failed automatic rollback retains the protected task definition export'
        $retainedDescriptorFiles = @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.sddl')
        Assert-Equal 1 $retainedDescriptorFiles.Count 'failed automatic rollback retains the protected task descriptor export'
        $rollbackFailureLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
        Assert-True ($rollbackFailureLog.Contains('RegisteredTask.SetSecurityDescriptor')) 'failed automatic rollback reports how to restore the task descriptor'
        Assert-True ($rollbackFailureLog.Contains($retainedRollbackFiles[0].FullName) -and $rollbackFailureLog.Contains($retainedDescriptorFiles[0].FullName) -and $rollbackFailureLog.Contains('\myTech.Today\Install-OOShutUp10')) 'failed automatic rollback logs both backups and the original task identity'
        Assert-True (-not $script:TaskRegistry.ContainsKey('\myTech.Today\|Install-OOShutUp10')) 'failed automatic rollback does not leave an unverified SYSTEM task registered'
        $script:FailTaskVerification = $false
        $script:FailTaskRestore = $false
        $script:FailTaskRestoreXml = $null
    }
    finally {
        Set-Item -Path Function:\Get-TrustedWindowsPowerShellExecutable -Value $savedTrustedPowerShellResolver
        $env:TEMP = $oldTemp
    }

    $script:StartedPaths.Clear()
    $approvedProfile = Get-OOShutUpApprovedProfile -RequireMachineScope
    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'the current profile is confirmed before configuration'
    Assert-Equal 'Unchanged' (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply).State 'a one-shot operator recovery after a missed event reuses machine-only profile validation and reports no change'
    Assert-Equal 1 $script:StartedPaths.Count 'the one-shot recovery starts one verified payload'
    Assert-Equal $validFixture ([string]$script:StartedPaths[0]) 'Start-Process receives the verified path'
    Assert-Equal $false $script:StartedWaitFlags[0] 'Start-Process returns immediately so the explicit bounded child wait controls liveness'
    Assert-Equal 900000 $script:MockChildWaitTimeoutMilliseconds 'the exact started child receives the fifteen-minute wait bound'

    foreach ($case in $invalidFixtures) {
        $fixture = Join-Path $script:FixtureRoot ($case.Name.Replace(' ', '-') + '.exe')
        $startedBefore = $script:StartedPaths.Count
        Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $fixture).State "$($case.Name) executable is not launched"
        Assert-Equal $startedBefore $script:StartedPaths.Count "$($case.Name) fixture makes zero process launches"
    }
    $startedBeforeEmpty = $script:StartedPaths.Count
    Assert-Equal 'Failed' (Invoke-OOShutUpConfiguration -ExecutablePath $emptyFixture).State 'empty executable is not launched'
    Assert-Equal $startedBeforeEmpty $script:StartedPaths.Count 'empty fixture makes zero process launches'

    $oldTemp = $env:TEMP
    $oldUserProfile = $env:USERPROFILE
    $oldLocalAppData = $env:LOCALAPPDATA
    $installerTemp = Join-Path $script:FixtureRoot 'installer temp with spaces'
    $systemTemp = Join-Path $script:FixtureRoot 'system temp with spaces'
    $installerLocalAppData = Join-Path $script:FixtureRoot 'installer local appdata'
    $systemLocalAppData = Join-Path $script:FixtureRoot 'system local appdata'
    $unrelatedArtifact = Join-Path $script:InstallPath 'preserve-this-file.txt'
    [System.IO.Directory]::CreateDirectory($installerTemp) | Out-Null
    [System.IO.Directory]::CreateDirectory($systemTemp) | Out-Null
    [System.IO.Directory]::CreateDirectory((Join-Path $installerLocalAppData 'Programs\OOShutUp10')) | Out-Null
    [System.IO.Directory]::CreateDirectory((Join-Path $systemLocalAppData 'Programs\OOShutUp10')) | Out-Null
    [System.IO.File]::WriteAllText($unrelatedArtifact, 'unrelated')
    $installerTempExecutable = Join-Path $installerTemp 'OOSU10.exe'
    $systemTempExecutable = Join-Path $systemTemp 'OOSU10.exe'
    $installerLocalExecutable = Join-Path $installerLocalAppData 'Programs\OOShutUp10\OOSU10.exe'
    $systemLocalExecutable = Join-Path $systemLocalAppData 'Programs\OOShutUp10\OOSU10.exe'
    Write-TestPayload -Path $installerTempExecutable -Content 'VALID'
    Write-TestPayload -Path $systemTempExecutable -Content 'VALID'
    Write-TestPayload -Path $installerLocalExecutable -Content 'VALID'
    Write-TestPayload -Path $systemLocalExecutable -Content 'VALID'

    try {
        $env:USERPROFILE = Join-Path $script:FixtureRoot 'installer profile'
        $env:TEMP = $installerTemp
        $env:LOCALAPPDATA = $installerLocalAppData
        $script:DownloadScenario = 'valid'
        Assert-Equal $true (Install-OOShutUp) 'fresh mocked installation promotes the verified executable into the managed machine directory'
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($script:ManagedExecutablePath)) 'fresh installation writes the verified portable executable to its absolute managed path'
        Assert-Equal $script:ManagedExecutablePath (Get-OOShutUpPath) 'installer lookup returns the absolute managed executable path'
        Assert-True ([System.IO.Path]::IsPathRooted((Get-OOShutUpPath))) 'the managed executable path is absolute'
        Assert-True ($script:ManagedExecutablePath.Contains('with spaces')) 'the managed machine path supports spaces'
        Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:InstallPath -Filter 'OOSU10.exe.*.download').Count 'fresh promotion leaves no staging file'

        $taskArguments = [string]$taskXml.Task.Actions.Exec.Arguments
        Assert-True ($taskArguments.Contains('-ReapplyOnly') -and $taskArguments.Contains("-File `"$taskScriptPath`"")) 'the LocalSystem task invokes the guarded ReapplyOnly script'
        Assert-True ($taskArguments.Contains("-ApprovedOOSExecutablePath `"$script:ManagedExecutablePath`"")) 'the LocalSystem task carries the exact managed executable path'
        $env:USERPROFILE = Join-Path $script:FixtureRoot 'system profile'
        $env:TEMP = $systemTemp
        $env:LOCALAPPDATA = $systemLocalAppData
        $systemResolvedPath = Get-OOShutUpPath -ExpectedPath $script:ManagedExecutablePath
        Assert-Equal $script:ManagedExecutablePath $systemResolvedPath 'SYSTEM environment changes do not change the managed executable path'
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($systemResolvedPath)) 'SYSTEM ReapplyOnly resolves the same verified machine copy installed by the user run'
        Assert-Equal $null (Get-OOShutUpPath -ExpectedPath $systemTempExecutable) 'SYSTEM ReapplyOnly rejects an executable path that differs from the task-pinned managed file'

        $script:StartedPaths.Clear()
        Assert-Equal 'Unchanged' (Invoke-OOShutUpConfiguration -ExecutablePath $systemResolvedPath -SystemReapply).State 'SYSTEM reapplication returns an explicit outcome through the mocked process boundary'
        Assert-Equal 1 $script:StartedPaths.Count 'SYSTEM reapplication starts one verified O&O payload'
        Assert-Equal $script:ManagedExecutablePath ([string]$script:StartedPaths[0]) 'the task reapplication path references the managed O&O executable'

        $replacementResult = Install-OOShutUp
        Assert-Equal $true $replacementResult ("a valid replacement is atomically promoted over the prior approved copy; logs: " + [string]::Join(' | ', [string[]]$script:LogMessages.ToArray()))
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($script:ManagedExecutablePath)) 'the approved portable path remains usable after replacement'

        $rejectedDownloadScenarios = @('tampered', 'unsigned', 'wrong-publisher', 'wrong-product', 'old-version', 'empty')
        foreach ($scenario in $rejectedDownloadScenarios) {
            Write-TestPayload -Path $script:ManagedExecutablePath -Content 'APPROVED'
            $script:DownloadScenario = $scenario
            $script:LogMessages.Clear()
            Assert-Equal $false (Install-OOShutUp) "$scenario download is rejected"
            Assert-Equal 'APPROVED' ([System.IO.File]::ReadAllText($script:ManagedExecutablePath)) "$scenario rejection preserves the previously approved executable"
            Assert-Equal $script:ManagedExecutablePath (Get-OOShutUpPath) "$scenario rejection keeps the approved copy selected"
            Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:InstallPath -Filter 'OOSU10.exe.*.download').Count "$scenario rejection removes only its staging file"
        }

        Write-TestPayload -Path $script:ManagedExecutablePath -Content 'APPROVED'
        $script:DownloadScenario = 'interrupted'
        $script:LogMessages.Clear()
        Assert-Equal $false (Install-OOShutUp) 'interrupted download fails with no promotion'
        Assert-Equal 'APPROVED' ([System.IO.File]::ReadAllText($script:ManagedExecutablePath)) 'interrupted download preserves the approved executable'
        Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:InstallPath -Filter 'OOSU10.exe.*.download').Count 'interrupted partial file is removed'
        Assert-Equal $script:ManagedExecutablePath (Get-OOShutUpPath) 'a later lookup cannot select the interrupted partial file'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('The O&O download failed. Check network connectivity and access to the managed machine directory, then retry.')) 'interrupted download reports network or machine-directory failure'

        $savedPromotion = (Get-Command -Name Move-OOShutUpFileAtomically -CommandType Function).ScriptBlock
        try {
            Set-Item -Path Function:\Move-OOShutUpFileAtomically -Value { throw 'Mocked interrupted promotion.' }
            $script:DownloadScenario = 'valid'
            $script:LogMessages.Clear()
            Assert-Equal $false (Install-OOShutUp) 'interrupted promotion fails closed'
            Assert-Equal 'APPROVED' ([System.IO.File]::ReadAllText($script:ManagedExecutablePath)) 'interrupted promotion preserves the prior verified executable'
            Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:InstallPath -Filter 'OOSU10.exe.*.download').Count 'interrupted promotion removes only its staging file'
        }
        finally {
            Set-Item -Path Function:\Move-OOShutUpFileAtomically -Value $savedPromotion
        }

        Assert-Equal 'unrelated' ([System.IO.File]::ReadAllText($unrelatedArtifact)) 'cleanup preserves unrelated files in the managed directory'
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($installerTempExecutable)) 'cleanup leaves installer TEMP files untouched'
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($systemTempExecutable)) 'cleanup leaves SYSTEM TEMP files untouched'
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($installerLocalExecutable)) 'cleanup leaves installer LOCALAPPDATA files untouched'
        Assert-Equal 'VALID' ([System.IO.File]::ReadAllText($systemLocalExecutable)) 'cleanup leaves SYSTEM LOCALAPPDATA files untouched'
    }
    finally {
        $env:TEMP = $oldTemp
        $env:USERPROFILE = $oldUserProfile
        $env:LOCALAPPDATA = $oldLocalAppData
    }

    $script:StartedPaths.Clear()
    Assert-Equal 'Unchanged' (Invoke-OOShutUpConfiguration -ExecutablePath $script:ManagedExecutablePath -SystemReapply).State 'the preserved approved copy can be passed to the mocked process boundary'
    Assert-Equal 1 $script:StartedPaths.Count 'the preserved approved copy launches exactly once through the mock'
    Assert-Equal $script:ManagedExecutablePath ([string]$script:StartedPaths[0]) 'the preserved approved copy is the launched path'

    [Console]::Out.WriteLine("Passed $script:AssertionCount O&O payload trust assertions in $($PSVersionTable.PSVersion).")
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    $tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    $resolvedFixtureRoot = [System.IO.Path]::GetFullPath($script:FixtureRoot)
    if ($resolvedFixtureRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) -and
        [System.IO.Path]::GetFileName($resolvedFixtureRoot).StartsWith('oosu10-trust-tests-', [System.StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $resolvedFixtureRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
