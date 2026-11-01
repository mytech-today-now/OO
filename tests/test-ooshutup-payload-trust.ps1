[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:LogMessages = New-Object System.Collections.ArrayList
$script:PromptMessages = New-Object System.Collections.ArrayList
$script:StartedPaths = New-Object System.Collections.ArrayList
$script:StartedArguments = New-Object System.Collections.ArrayList
$script:StartedWorkingDirectories = New-Object System.Collections.ArrayList
$script:DownloadScenario = 'valid'
$script:DownloadCount = 0
$script:SignatureCallCount = 0
$script:MockExitCode = 0
$script:PromptDecision = 'YES'
$script:FixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oosu10-trust-tests-' + [guid]::NewGuid().ToString('N'))
$script:TaskRegistry = @{}
$script:TaskRegistrationCalls = New-Object System.Collections.ArrayList
$script:TaskUnregistrationCalls = New-Object System.Collections.ArrayList
$script:TaskLookupCalls = New-Object System.Collections.ArrayList
$script:TaskLookupFailure = $null
$script:TaskRegistrationFailure = $null
$script:TaskVerificationMissPending = $false
$script:FailTaskRestoreXml = $null
$script:FailTaskVerification = $false
$script:FailStageVerification = $false
$script:FailStageQueuePolicyVerification = $false
$script:FailTaskRestore = $false
$script:MockInstallerTaskScenario = 'Success'
$script:OOShutUpPartialSetupMessage = 'The privacy settings may be applied, but the update task was not verified. Setup is partial; repair the task before relying on automatic reapplication.'
$script:OOShutUpTaskRecoveryMessage = 'Recovery: resolve the logged Task Scheduler error, open PowerShell as Administrator, and rerun this installer with the approved profile (or use -NonInteractive -ApprovedProfileSha256 <profile SHA-256> -ApprovedProfileManifestSha256 <manifest SHA-256>). Verify \myTech.Today\Install-OOShutUp10 in Task Scheduler before relying on automatic reapplication. The retry inspects and replaces only this verified task; same-named tasks in other folders are left unchanged.'
$script:OOShutUpUpdateEventSubscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''Microsoft-Windows-WindowsUpdateClient''] and (EventID=19)]]</Select></Query></QueryList>'
$script:OOShutUpUpdateEventDelay = 'PT2M'
$script:OOShutUpUpdateMultipleInstancesPolicy = 'Queue'
$script:OOShutUpRestorePointFailureMessage = 'A restore point could not be verified. No privacy settings were applied; free space or enable restore protection, then retry.'
$script:MockAclByPath = @{}
$script:MockAclFailurePath = $null

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
    param([string]$Name)

    $certificate = [pscustomobject]@{ SimpleName = $Name }
    $certificate | Add-Member -MemberType ScriptMethod -Name GetNameInfo -Value { return $this.SimpleName } -Force
    return $certificate
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

    [void]$script:StartedPaths.Add($FilePath)
    [void]$script:StartedArguments.Add($ArgumentList)
    [void]$script:StartedWorkingDirectories.Add($WorkingDirectory)
    return [pscustomobject]@{ ExitCode = $script:MockExitCode }
}

function Write-ApprovedProfileFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $true)]
        [string[]]$AllowedSettingIds,

        [Parameter(Mandatory = $false)]
        [string[]]$CurrentUserSettingIds = @()
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
        schemaVersion = 3
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
        [System.Security.AccessControl.AccessControlType]$Type = [System.Security.AccessControl.AccessControlType]::Allow
    )

    return [pscustomobject]@{
        IdentityReference = [System.Security.Principal.SecurityIdentifier]::new($Sid)
        FileSystemRights = $Rights
        AccessControlType = $Type
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
        Settings = [pscustomobject]@{ MultipleInstances = [string]$multipleInstancesNode.InnerText }
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

        [int]$EventOccurredAtSeconds = 10
    )

    $trigger = $TaskDefinition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Triggers']/*[local-name()='EventTrigger']")
    $subscription = [string]$trigger.Subscription
    $configuredEventIds = @([regex]::Matches($subscription, 'EventID=(\d+)') | ForEach-Object { [int]$_.Groups[1].Value })
    $delaySeconds = [int][System.Xml.XmlConvert]::ToTimeSpan([string]$trigger.Delay).TotalSeconds
    $policyNode = $TaskDefinition.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='MultipleInstancesPolicy']")
    $policy = [string]$policyNode.InnerText
    $runs = New-Object System.Collections.ArrayList

    foreach ($eventId in $EventIds) {
        if ($configuredEventIds -notcontains $eventId) {
            continue
        }

        $startSeconds = $EventOccurredAtSeconds + $delaySeconds
        if ($startSeconds -lt $ActiveRunEndSeconds) {
            if ($policy -ne 'Queue') {
                continue
            }
            $startSeconds = $ActiveRunEndSeconds
        }

        [void]$runs.Add([pscustomobject]@{ EventId = $eventId; StartSeconds = $startSeconds; ActiveRunEventId = $ActiveRunEventId })
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
    $script:TaskRegistry["$TaskPath|$TaskName"] = $task
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
    $installerBatchSource = Get-Content -LiteralPath (Join-Path $repoRoot 'install.bat') -Raw
    Assert-True ($source.Contains('$DownloadUrl = "https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe"')) 'the official vendor download URL is preserved'
    $versionPolicy = '$MinimumOOShutUpVersion = [version]' + [char]39 + '3.6.1135' + [char]39
    Assert-True ($source.Contains($versionPolicy)) 'the supported version floor is explicit'
    Assert-True ($source.Contains('Get-AuthenticodeSignature -FilePath $Path')) 'Authenticode is checked before accepting a payload'
    Assert-True ($source.Contains('The O&O download could not be authenticated. No executable was started; verify the vendor signature before retrying.')) 'the required authentication fallback is present'
    Assert-True ($source.Contains('The approved settings profile is missing or invalid. No privacy settings were applied; place the reviewed profile in the configured directory.')) 'the required profile fallback is present'
    Assert-True ($source.Contains('The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.')) 'unknown scope has the required fail-closed status'
    Assert-True ($source.Contains('The SYSTEM task target ACL could not be verified. No task was registered; move the script to a protected path and retry.')) 'the required task ACL fallback is present'
    Assert-True ($installerBatchSource.Contains('set "INSTALLER_EXIT_CODE=!ERRORLEVEL!"') -and $installerBatchSource.Contains('exit /b !INSTALLER_EXIT_CODE!')) 'the batch wrapper preserves the PowerShell installer exit code'
    Assert-True ($installerBatchSource.Contains('Setup is partial. Repair the update task before relying on automatic reapplication.')) 'the batch wrapper reports partial setup with the task recovery action'
    Assert-True ($installerBatchSource -notmatch '\[INFO\] Complete\.') 'the batch wrapper does not print unconditional completion'
    Assert-True ($source.Contains('Task ownership is ambiguous. No task was deleted; inspect the listed path and retry after confirming ownership.')) 'ambiguous task ownership has the required fallback'
    Assert-True ($source.Contains('-RequireMachineScope:$ReapplyOnly')) 'ReapplyOnly validates machine scope before later work'
    Assert-True ($source.Contains("Join-Path `$PSScriptRoot 'profiles'")) 'the profile package path is rooted beside the script'
    Assert-True ($source -notmatch 'Start-Process[^\r\n]*ooshutup10\.cfg') 'the O&O invocation does not pass a bare profile filename'
    Assert-True ($source -notmatch '\$TempPath') 'the managed executable path no longer depends on a TEMP-derived variable'
    Assert-True ($source.Contains('$env:ProgramW6432') -and $source.Contains('$env:ProgramFiles')) 'the machine install directory uses the machine Program Files location'

    $requiredFunctions = @(
        'Get-OOShutUpFileVersionInfo',
        'Test-OOShutUpPayload',
        'Test-OOShutUpInstalled',
        'Get-OOShutUpPath',
        'Get-OOShutUpApprovedProfile',
        'Show-OOShutUpProfileImpactSummary',
        'Confirm-OOShutUpProfile',
        'New-OOShutUpWebClient',
        'Move-OOShutUpFileAtomically',
        'Install-OOShutUp',
        'Invoke-OOShutUpConfiguration',
        'ConvertTo-SystemTaskSid',
        'Test-SystemTaskTargetAcl',
        'Test-WindowsUpdateTaskDefinition',
        'Test-WindowsUpdateLegacyTaskDefinition',
        'Test-ScheduledTaskAccessDenied',
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

    $managedPathFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Get-OOShutUpPath' })[0].Extent.Text)
    Assert-True ($managedPathFunctionSource.Contains('Join-Path $InstallPath $ExecutableName')) 'path lookup uses the single managed machine directory'
    Assert-True ($managedPathFunctionSource -notmatch '\$env:(?:TEMP|USERPROFILE|LOCALAPPDATA)|Get-Command') 'path lookup does not search user, temporary, or PATH locations'

    $mainStart = $source.IndexOf('#region Main Script', [System.StringComparison]::Ordinal)
    $mainSource = $source.Substring($mainStart)
    $profileGatePosition = $mainSource.IndexOf('$approvedProfile = Get-OOShutUpApprovedProfile', [System.StringComparison]::Ordinal)
    Assert-True ($profileGatePosition -ge 0) 'main execution performs the profile preflight'
    $profileApprovalGatePosition = $mainSource.IndexOf('$profileApprovalAccepted = Confirm-OOShutUpProfile', [System.StringComparison]::Ordinal)
    Assert-True ($profileApprovalGatePosition -gt $profileGatePosition) 'main execution requires an approval decision after profile validation'
    $reapplyStart = $mainSource.IndexOf('if ($ReapplyOnly)', [System.StringComparison]::Ordinal)
    $reapplyPathPosition = $mainSource.IndexOf('$exePath = Get-OOShutUpPath', $reapplyStart, [System.StringComparison]::Ordinal)
    Assert-True ($reapplyStart -ge 0 -and $reapplyPathPosition -gt $reapplyStart) 'SYSTEM ReapplyOnly resolves the executable through the managed-path lookup'
    Assert-True ($mainSource.Contains('Get-OOShutUpPath -ExpectedPath $ApprovedOOSExecutablePath')) 'SYSTEM ReapplyOnly validates the exact executable path pinned in the task action'
    Assert-True ($source.Contains('The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.') -and $mainSource.Contains('$script:OOShutUpManagedPathFailureMessage')) 'SYSTEM ReapplyOnly reports the exact managed-copy recovery instruction'
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
    Assert-True ($source.Contains('The privacy settings may be applied, but the update task was not verified. Setup is partial; repair the task before relying on automatic reapplication.')) 'task failure uses the exact partial setup status'
    Assert-True ($source.Contains('Verify \myTech.Today\Install-OOShutUp10 in Task Scheduler')) 'partial setup includes the exact task identity for recovery'
    Assert-True ($mainSource.Contains('$taskSetup = Invoke-OOShutUpTaskSetup') -and $mainSource.Contains('$installerOutcome = Get-OOShutUpInstallOutcome')) 'the normal flow consumes explicit task and installer outcomes'
    Assert-True ($mainSource.Contains('exit $installerOutcome.ExitCode')) 'the normal flow returns the computed complete, partial, or failed exit code'
    Assert-True ($mainSource.Contains('Write-OOShutUpInstallSummary -Outcome $installerOutcome')) 'the normal flow logs the summary from the installer outcome contract'
    Assert-True (-not $mainSource.Contains('=== Configuration Complete ===')) 'full-success wording is not unconditional in the normal flow'
    $summaryFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Write-OOShutUpInstallSummary' })[0].Extent.Text)
    $completeStatusPosition = $summaryFunctionSource.IndexOf('if ($Outcome.Status -eq ''Complete'')', [System.StringComparison]::Ordinal)
    $completeSummaryPosition = $summaryFunctionSource.IndexOf("Write-Log '=== Configuration Complete ===' -Level SUCCESS", [System.StringComparison]::Ordinal)
    $partialStatusPosition = $summaryFunctionSource.IndexOf('elseif ($Outcome.Status -eq ''Partial'')', [System.StringComparison]::Ordinal)
    Assert-True ($completeStatusPosition -ge 0 -and $completeSummaryPosition -gt $completeStatusPosition -and $partialStatusPosition -gt $completeSummaryPosition) 'full-success wording is guarded by the verified complete outcome'
    $noChangeRestartGatePosition = $mainSource.IndexOf('if (-not $script:OOShutUpLastRunChanged)', [System.StringComparison]::Ordinal)
    $restartHandlerCallPosition = $mainSource.IndexOf('Invoke-OOShutUpRestart -NonInteractive:', [System.StringComparison]::Ordinal)
    $restartHandlerSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpRestart' })[0].Extent.Text)
    Assert-True ($noChangeRestartGatePosition -ge 0 -and $restartHandlerCallPosition -gt $noChangeRestartGatePosition) 'the restart handler is called only after the no-change guard'
    Assert-True ($restartHandlerSource.Contains('Press Y to restart, N or Esc to defer') -and $restartHandlerSource.Contains('Restart-Computer -Force')) 'the guarded restart handler owns the confirmation prompt and explicit force command'
    $taskFunctionSource = ([string](@($functionAsts | Where-Object { $_.Name -eq 'New-WindowsUpdateScheduledTask' })[0].Extent.Text))
    $taskApprovalPosition = $taskFunctionSource.IndexOf('$currentApprovedProfile = Get-OOShutUpApprovedProfile', [System.StringComparison]::Ordinal)
    $pathCanonicalizationPosition = $taskFunctionSource.IndexOf('$ScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)', [System.StringComparison]::Ordinal)
    $targetAclPosition = $taskFunctionSource.IndexOf('Test-SystemTaskTargetAcl -ScriptPath $ScriptPath', [System.StringComparison]::Ordinal)
    $taskLookupPosition = $taskFunctionSource.IndexOf('$existingTasks = @(Get-ScheduledTask', [System.StringComparison]::Ordinal)
    Assert-True ($taskApprovalPosition -ge 0 -and $pathCanonicalizationPosition -gt $taskApprovalPosition -and $targetAclPosition -gt $pathCanonicalizationPosition -and $taskLookupPosition -gt $targetAclPosition) 'profile approval, path canonicalization, and target ACL verification precede task inspection or mutation'
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

        foreach ($scenario in @('Success', 'Failure', 'Exception')) {
            $script:MockInstallerTaskScenario = $scenario
            $taskSetupResult = Invoke-OOShutUpTaskSetup -ScriptPath 'C:\fixture\Install-OOShutUp10.ps1' -ApprovedProfileSha256 ('A' * 64) -ApprovedProfileManifestSha256 ('B' * 64)
            $installResult = Get-OOShutUpInstallOutcome -SettingsVerified $true -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $taskSetupResult
            $script:LogMessages.Clear()
            Write-OOShutUpInstallSummary -Outcome $installResult
            $summaryLog = [string]::Join([Environment]::NewLine, [string[]]$script:LogMessages.ToArray())
            if ($scenario -eq 'Success') {
                Assert-Equal $true $taskSetupResult.Verified 'a true task creation and verification result is preserved'
                Assert-Equal 0 $installResult.ExitCode 'verified settings and task return the success exit code'
                Assert-Equal 'Complete' $installResult.Status 'verified settings and task report complete status'
                Assert-Equal 'Verified' $installResult.TaskOutcome 'complete status identifies the task as verified'
                Assert-True ($installResult.Summary.Contains('\myTech.Today\Install-OOShutUp10') -and -not $installResult.Summary.Contains('partial')) 'success summary identifies the verified task'
                Assert-True ($summaryLog.Contains('=== Configuration Complete ===') -and $summaryLog.Contains($installResult.Summary) -and -not $summaryLog.Contains('=== Setup Partial ===')) 'success logging agrees with the verified task outcome'
            }
            else {
                Assert-Equal $false $taskSetupResult.Verified "$scenario task setup is not treated as verified"
                Assert-Equal 2 $installResult.ExitCode "$scenario task setup returns the distinct partial exit code"
                Assert-Equal 'Partial' $installResult.Status "$scenario task setup reports partial status"
                Assert-Equal 'Not verified' $installResult.TaskOutcome "$scenario task setup reports the task as unverified"
                Assert-Equal $script:OOShutUpPartialSetupMessage $installResult.Summary "$scenario summary matches the required partial wording"
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
        $overrideInstallResult = Get-OOShutUpInstallOutcome -SettingsVerified $true -RestorePointVerified $false -RestorePointOverrideUsed $true -TaskSetup $overrideSetupResult
        Assert-Equal 0 $overrideInstallResult.ExitCode 'an explicit restore point override still permits success when the task is verified'
        Assert-Equal 'Explicitly overridden' $overrideInstallResult.RestorePointOutcome 'the result retains the separate restore point outcome'
    }
    finally {
        Set-Item -Path Function:\New-WindowsUpdateScheduledTask -Value $savedTaskCreator
    }

    # Replace only the metadata reader after loading production functions so the fixture stays offline.
    function Get-OOShutUpFileVersionInfo {
        param([string]$Path)

        $content = [System.IO.File]::ReadAllText($Path)
        $fileVersion = if ($content -eq 'OLD_VERSION') { '3.5.1130.0' } else { '3.6.1135.0' }
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
    $script:OOShutUpUnknownScopeFailureMessage = 'The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.'
    $script:OOShutUpCurrentUserScopeFailureMessage = 'The SYSTEM task profile contains current-user settings. Reapplication was skipped; run them for the intended signed-in user or remove them from the SYSTEM task profile.'
    $script:OOShutUpApprovalFailureMessage = 'The selected settings profile has no matching approval. No settings were changed; review both file digests in the interactive installer and recreate the task.'
    $script:OOShutUpTaskAclFailureMessage = 'The SYSTEM task target ACL could not be verified. No task was registered; move the script to a protected path and retry.'
    $script:OOShutUpTaskOwnershipFailureMessage = 'Task ownership is ambiguous. No task was deleted; inspect the listed path and retry after confirming ownership.'
    $script:OOShutUpManagedPathFailureMessage = 'The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.'

    $validFixture = Join-Path $script:FixtureRoot 'valid.exe'
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
    Assert-Equal 'Test category' ([string]$approvedProfile.Settings[0].Category) 'validated impact metadata is included in the interactive profile summary'
    $profileLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($profileLog -notmatch 'S003|M003|Example setting') 'profile identity logging omits setting IDs and values'

    $script:PromptMessages.Clear()
    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive YES accepts the reviewed profile'
    $summaryText = [string]::Join("`n", [string[]]$script:PromptMessages.ToArray())
    Assert-True ($summaryText.Contains($approvedProfile.Sha256)) 'interactive summary displays the exact profile digest'
    Assert-True ($summaryText.Contains('Test category') -and $summaryText.Contains('S003') -and $summaryText.Contains('Test fixture impact description.')) 'interactive summary displays reviewed category, ID, and impact text'

    Write-ApprovedProfileFixture -Content "# Mixed-scope fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003') -CurrentUserSettingIds @('M003')
    $mixedScopeProfile = Get-OOShutUpApprovedProfile
    $currentUserSetting = @($mixedScopeProfile.Settings | Where-Object { $_.Id -eq 'M003' })[0]
    Assert-Equal 'current-user' ([string]$currentUserSetting.Scope) 'reviewed current-user scope is retained in the approved profile'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $mixedScopeProfile -NonInteractive -ApprovedProfileSha256 $mixedScopeProfile.Sha256 -ApprovedProfileManifestSha256 $mixedScopeProfile.ManifestSha256) 'mixed profile can be approved for an interactive run with both digests'
    $startedBeforeSystemScopeBlock = $script:StartedPaths.Count
    $script:LogMessages.Clear()
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply) 'SYSTEM reapplication rejects a profile containing current-user settings'
    Assert-Equal $startedBeforeSystemScopeBlock $script:StartedPaths.Count 'current-user scope is rejected before O&O launches'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpCurrentUserScopeFailureMessage)) 'current-user scope rejection explains the safe fallback'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`tExample setting`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $approvedProfile = Get-OOShutUpApprovedProfile
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'restored machine-only fixture requires fresh approval'

    $startedBeforeDecline = $script:StartedPaths.Count
    $script:PromptDecision = 'NO'
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive NO rejects the reviewed profile'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a rejected profile cannot be applied'
    Assert-Equal $startedBeforeDecline $script:StartedPaths.Count 'rejecting approval causes zero O&O launches'

    $script:PromptDecision = 'CANCEL'
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive CANCEL stops approval'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a canceled profile cannot be applied'
    Assert-Equal $startedBeforeDecline $script:StartedPaths.Count 'canceling approval causes zero O&O launches'

    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive -ApprovedProfileSha256 $approvedProfile.Sha256 -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256) 'unattended approval accepts the exact profile and manifest digests'
    Assert-Equal $true (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'the exact unattended approval permits the mocked O&O launch'
    Assert-Equal ($startedBeforeDecline + 1) $script:StartedPaths.Count 'an exact approved digest launches O&O once'
    Assert-Equal $false $script:OOShutUpLastRunChanged 'O&O exit code zero records that no setting changed'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains('O&O reported that no setting changes were needed')) 'O&O exit code zero reports a no-change result'

    $startedBeforeMissingApproval = $script:StartedPaths.Count
    $script:LogMessages.Clear()
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive) 'unattended mode rejects a missing approval digest'
    Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpApprovalFailureMessage)) 'missing unattended approval reports the required fallback'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a missing unattended approval cannot be applied'
    Assert-Equal $startedBeforeMissingApproval $script:StartedPaths.Count 'missing approval causes zero O&O launches'
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $approvedProfile -NonInteractive -ApprovedProfileSha256 ('0' * 64) -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256) 'unattended mode rejects a stale profile digest'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a stale unattended approval cannot be applied'
    Assert-Equal $startedBeforeMissingApproval $script:StartedPaths.Count 'stale approval causes zero O&O launches'

    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'interactive approval can be accepted after rejected attempts'
    $script:MockExitCode = 30
    Assert-Equal $true (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'O&O exit code 30 succeeds through the mocked process boundary'
    Assert-Equal $true $script:OOShutUpLastRunChanged 'O&O exit code 30 records that settings changed'
    $script:MockExitCode = 0

    $oldApprovedDigest = $approvedProfile.Sha256
    Write-ApprovedProfileFixture -Content "# Changed reviewed fixture`r`nS003`t-`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $changedProfile = Get-OOShutUpApprovedProfile
    Assert-True ($changedProfile.Sha256 -cne $oldApprovedDigest) 'a changed profile receives a new digest'
    $startedBeforeChangedDigest = $script:StartedPaths.Count
    Assert-Equal $false (Confirm-OOShutUpProfile -Profile $changedProfile -NonInteractive -ApprovedProfileSha256 $oldApprovedDigest -ApprovedProfileManifestSha256 $changedProfile.ManifestSha256) 'an approval for the prior profile does not authorize a changed profile'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'the changed profile cannot launch with its prior digest approval'
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
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'scope metadata changed after approval cannot reach O&O'
    Assert-Equal $startedBeforeManifestTamper $script:StartedPaths.Count 'manifest tampering causes zero O&O process launches'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`tExample setting`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')
    $approvedProfile = Get-OOShutUpApprovedProfile
    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'the restored original profile requires a fresh interactive approval'

    $unrelatedWorkingDirectory = Join-Path $script:FixtureRoot 'unrelated caller directory'
    [System.IO.Directory]::CreateDirectory($unrelatedWorkingDirectory) | Out-Null
    Push-Location -LiteralPath $unrelatedWorkingDirectory
    try {
        $script:StartedPaths.Clear()
        $script:StartedArguments.Clear()
        $script:StartedWorkingDirectories.Clear()
        $script:LogMessages.Clear()
        Assert-Equal $true (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'configuration succeeds from an unrelated current directory'
    }
    finally {
        Pop-Location
    }
    Assert-Equal 1 $script:StartedPaths.Count 'a valid profile launches the verified executable exactly once'
    Assert-Equal ('/quiet /nosrp "{0}"' -f $script:OOShutUpProfilePath) ([string]$script:StartedArguments[0]) 'O&O receives the quoted absolute profile path'
    Assert-Equal $script:OOShutUpProfileDirectory ([string]$script:StartedWorkingDirectories[0]) 'O&O starts in the configured profile directory'
    $launchProfileLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
    Assert-True ($launchProfileLog.Contains("version test-1.0 (SHA-256 $($approvedProfile.Sha256)")) 'profile version and digest are recorded at launch'
    Assert-True ($launchProfileLog -notmatch 'S003|M003|Example setting') 'launch logging omits setting IDs and values'

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
        Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) "$($case.Name) profile prevents configuration"
        Assert-Equal $startedBefore $script:StartedPaths.Count "$($case.Name) profile causes no process launch"
    }

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
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply) 'unknown scope prevents SYSTEM reapplication'
    Assert-Equal $startedBeforeUnknownScope $script:StartedPaths.Count 'unknown scope causes zero O&O process launches'

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
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'mismatched impact metadata prevents configuration'
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
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'missing impact details prevent configuration'
    Assert-Equal $startedBeforeMissingImpact $script:StartedPaths.Count 'missing impact details cause zero O&O launches'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    [System.IO.File]::AppendAllText($script:OOShutUpProfilePath, "M003`t-`r`n")
    $startedBeforeDigestMismatch = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a digest-mismatched profile is rejected'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a digest mismatch prevents configuration'
    Assert-Equal $startedBeforeDigestMismatch $script:StartedPaths.Count 'a digest mismatch causes no process launch'

    Write-ApprovedProfileFixture -Content "S003`t+`r`n" -AllowedSettingIds @('S003')
    [System.IO.File]::WriteAllText($script:OOShutUpProfileManifestPath, '{ invalid json')
    $startedBeforeMalformedManifest = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a malformed approval manifest is rejected'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a malformed manifest prevents configuration'
    Assert-Equal $startedBeforeMalformedManifest $script:StartedPaths.Count 'a malformed manifest causes no process launch'

    Remove-Item -LiteralPath $script:OOShutUpProfilePath -Force
    $startedBeforeMissingProfile = $script:StartedPaths.Count
    Assert-Equal $null (Get-OOShutUpApprovedProfile) 'a missing profile is rejected'
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture) 'a missing profile prevents configuration'
    Assert-Equal $startedBeforeMissingProfile $script:StartedPaths.Count 'a missing profile causes no process launch'

    Write-ApprovedProfileFixture -Content "# Reviewed fixture`r`nS003`t+`r`nM003;-`r`n" -AllowedSettingIds @('S003', 'M003')

    $oldTemp = $env:TEMP
    $env:TEMP = $script:FixtureRoot
    $script:ScheduledTaskLookupCount = 0
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

        $script:MockAclFailurePath = [System.IO.Path]::GetFullPath($taskScriptPath)
        $registrationsBeforeDeniedAcl = $script:TaskRegistrationCalls.Count
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'an ACL inspection error fails closed'
        Assert-Equal $registrationsBeforeDeniedAcl $script:TaskRegistrationCalls.Count 'an ACL inspection error never registers a candidate task'
        Assert-True ([string]::Join("`n", [string[]]$script:LogMessages.ToArray()).Contains($script:OOShutUpTaskAclFailureMessage)) 'ACL inspection failure reports the required fallback'
        $script:MockAclFailurePath = $null

        Assert-Equal $true (Test-SystemTaskTargetAcl -ScriptPath $taskScriptPath) 'a safe temporary ACL fixture permits the target'
        Assert-Equal $true (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a safe ACL fixture reaches mocked candidate and final task registration'
        Assert-True ($script:TaskRegistrationCalls.Count -ge 2) 'candidate and final registration use mocked Task Scheduler commands'
        $taskXmlPath = Join-Path $script:FixtureRoot 'OOShutUp10Task_Debug.xml'
        [xml]$taskXml = [System.IO.File]::ReadAllText($taskXmlPath)
        $taskWorkingDirectory = $taskXml.Task.Actions.Exec.WorkingDirectory
        Assert-Equal (Split-Path -Parent $taskScriptPath) ([string]$taskWorkingDirectory) 'scheduled task uses the script package as its working directory'
        Assert-Equal 'S-1-5-18' ([string]$taskXml.Task.Principals.Principal.UserId) 'scheduled task runs under LocalSystem'
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
        $triggerSubscriptionText = [string]$trigger.Subscription
        Assert-Equal 'true' ([string]$trigger.Enabled) 'scheduled task preserves its enabled Windows Update event trigger'
        Assert-Equal 'PT2M' ([string]$trigger.Delay) 'scheduled task preserves its two-minute update delay'
        Assert-True ($triggerSubscriptionText.Contains('WindowsUpdateClient') -and $triggerSubscriptionText.Contains('EventID=19') -and -not $triggerSubscriptionText.Contains('EventID=20') -and -not $triggerSubscriptionText.Contains('EventID=43')) 'scheduled task triggers only on Windows Update installation success'
        Assert-Equal 'Queue' ([string]$taskXml.Task.Settings.MultipleInstancesPolicy) 'scheduled task queues successful-update triggers behind an active reapply'

        $verifiedCurrentTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $taskXml.OuterXml
        Assert-Equal $true (Test-WindowsUpdateTaskDefinition -Task $verifiedCurrentTask -TaskPath '\myTech.Today\' -TaskName 'Install-OOShutUp10' -TaskDescription ([string]$taskXml.Task.RegistrationInfo.Description) -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256 -ApprovedOOSExecutablePath $script:ManagedExecutablePath) 'task verification checks the success trigger and queue policy'

        $ignoreNewDocument = New-Object System.Xml.XmlDocument
        $ignoreNewDocument.LoadXml($taskXml.OuterXml)
        $ignoreNewDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='MultipleInstancesPolicy']").InnerText = 'IgnoreNew'
        $ignoreNewTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $ignoreNewDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition -Task $ignoreNewTask -TaskPath '\myTech.Today\' -TaskName 'Install-OOShutUp10' -TaskDescription ([string]$taskXml.Task.RegistrationInfo.Description) -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256 -ApprovedOOSExecutablePath $script:ManagedExecutablePath) 'IgnoreNew is rejected as an unsafe current task definition'

        $startAndSuccessDocument = New-Object System.Xml.XmlDocument
        $startAndSuccessDocument.LoadXml($taskXml.OuterXml)
        $startAndSuccessDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Triggers']/*[local-name()='EventTrigger']/*[local-name()='Subscription']").InnerText = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''Microsoft-Windows-WindowsUpdateClient''] and (EventID=19 or EventID=43)]]</Select></Query></QueryList>'
        $startAndSuccessTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $startAndSuccessDocument.OuterXml
        Assert-Equal $false (Test-WindowsUpdateTaskDefinition -Task $startAndSuccessTask -TaskPath '\myTech.Today\' -TaskName 'Install-OOShutUp10' -TaskDescription ([string]$taskXml.Task.RegistrationInfo.Description) -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256 -ApprovedOOSExecutablePath $script:ManagedExecutablePath) 'the old start-and-success trigger is rejected as a current task definition'

        $noEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @())
        $failureEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(20))
        $startEventRuns = @(Invoke-MockUpdateTaskSchedule -TaskDefinition $taskXml -EventIds @(43))
        Assert-Equal 0 $noEventRuns.Count 'no Windows Update event leaves the task idle'
        Assert-Equal 0 $failureEventRuns.Count 'a Windows Update installation failure does not start or report a successful reapply'
        Assert-Equal 0 $startEventRuns.Count 'an installation-start event does not start an early reapply'

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
        $previousApprovedTaskArguments = $previousApprovedTaskDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Actions']/*[local-name()='Exec']/*[local-name()='Arguments']")
        $previousApprovedTaskArguments.InnerText = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $taskScriptPath + '" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedProfileSha256 ' + $taskProfile.Sha256 + ' -ApprovedProfileManifestSha256 ' + $taskProfile.ManifestSha256
        $previousApprovedTask = New-MockScheduledTask -TaskName 'Install-OOShutUp10' -TaskPath '\myTech.Today\' -Xml $previousApprovedTaskDocument.OuterXml
        Assert-Equal $true (Test-WindowsUpdateLegacyTaskDefinition -Task $previousApprovedTask -TaskPath '\myTech.Today\' -TaskName 'Install-OOShutUp10' -TaskDescription ([string]$taskXml.Task.RegistrationInfo.Description) -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'the previous approved task definition remains eligible for safe replacement with the pinned executable path'

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
        $deniedRegistrationOutcome = Get-OOShutUpInstallOutcome -SettingsVerified $true -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $deniedRegistrationSetup
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
        $retryOutcome = Get-OOShutUpInstallOutcome -SettingsVerified $true -RestorePointVerified $true -RestorePointOverrideUsed $false -TaskSetup $retrySetup
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
        Assert-True ((@($script:LogMessages.ToArray()) -join [Environment]::NewLine).Contains('trigger, and queue settings')) 'queue-policy verification failure reports the exact rejected task criteria'
        $script:FailStageQueuePolicyVerification = $false

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
        $script:FailTaskRestoreXml = $oldTaskXml
        $script:LogMessages.Clear()
        Assert-Equal $false (New-WindowsUpdateScheduledTask -ScriptPath $taskScriptPath -ApprovedProfileSha256 $taskProfile.Sha256 -ApprovedProfileManifestSha256 $taskProfile.ManifestSha256) 'a failed automatic rollback reports failure'
        $retainedRollbackFiles = @(Get-ChildItem -LiteralPath $taskScriptDirectory -Filter '.Install-OOShutUp10-task-rollback-*.xml')
        Assert-Equal 1 $retainedRollbackFiles.Count 'failed automatic rollback retains the protected task definition export'
        $rollbackFailureLog = [string]::Join("`n", [string[]]$script:LogMessages.ToArray())
        Assert-True ($rollbackFailureLog.Contains('Register-ScheduledTask -TaskName <name> -TaskPath <path>')) 'failed automatic rollback reports a manual recovery command'
        Assert-True ($rollbackFailureLog.Contains($retainedRollbackFiles[0].FullName) -and $rollbackFailureLog.Contains('\myTech.Today\Install-OOShutUp10')) 'failed automatic rollback logs the backup and original task identity'
        Assert-True (-not $script:TaskRegistry.ContainsKey('\myTech.Today\|Install-OOShutUp10')) 'failed automatic rollback does not leave an unverified SYSTEM task registered'
        $script:FailTaskVerification = $false
        $script:FailTaskRestore = $false
        $script:FailTaskRestoreXml = $null
    }
    finally {
        $env:TEMP = $oldTemp
    }

    $script:StartedPaths.Clear()
    $approvedProfile = Get-OOShutUpApprovedProfile
    $script:PromptDecision = 'YES'
    Assert-Equal $true (Confirm-OOShutUpProfile -Profile $approvedProfile) 'the current profile is confirmed before configuration'
    Assert-Equal $true (Invoke-OOShutUpConfiguration -ExecutablePath $validFixture -SystemReapply) 'verified machine-only profile proceeds through the mocked SYSTEM reapplication call'
    Assert-Equal 1 $script:StartedPaths.Count 'a verified payload starts once'
    Assert-Equal $validFixture ([string]$script:StartedPaths[0]) 'Start-Process receives the verified path'

    foreach ($case in $invalidFixtures) {
        $fixture = Join-Path $script:FixtureRoot ($case.Name.Replace(' ', '-') + '.exe')
        $startedBefore = $script:StartedPaths.Count
        Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $fixture) "$($case.Name) executable is not launched"
        Assert-Equal $startedBefore $script:StartedPaths.Count "$($case.Name) fixture makes zero process launches"
    }
    $startedBeforeEmpty = $script:StartedPaths.Count
    Assert-Equal $false (Invoke-OOShutUpConfiguration -ExecutablePath $emptyFixture) 'empty executable is not launched'
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
        Assert-Equal $true (Invoke-OOShutUpConfiguration -ExecutablePath $systemResolvedPath -SystemReapply) 'SYSTEM reapplication uses the resolved managed executable through the mocked process boundary'
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
    Assert-Equal $true (Invoke-OOShutUpConfiguration -ExecutablePath $script:ManagedExecutablePath -SystemReapply) 'the preserved approved copy can be passed to the mocked process boundary'
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
