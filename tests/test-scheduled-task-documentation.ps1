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

function Get-SourceStringAssignment {
    param([string]$VariableName)

    $pattern = '(?im)^\s*\$' + [regex]::Escape($VariableName) + '\s*=\s*["''](?<value>[^"'']+)["'']\s*$'
    $matches = [regex]::Matches($installer, $pattern)
    Assert-Equal 1 $matches.Count "installer defines one literal `$${VariableName} task identity value"
    return [string]$matches[0].Groups['value'].Value
}

function Get-MarkdownSection {
    param([string]$Heading)

    $pattern = '(?ms)^###\s+' + [regex]::Escape($Heading) + '\s*\r?\n(?<body>.*?)(?=^###\s+|\z)'
    $match = [regex]::Match($readme, $pattern)
    Assert-True $match.Success "README has the '$Heading' section"
    return [string]$match.Groups['body'].Value
}

function Get-FirstPowerShellBlock {
    param([string]$Heading)

    $section = Get-MarkdownSection -Heading $Heading
    $match = [regex]::Match($section, '(?ms)^```powershell\s*\r?\n(?<code>.*?)^```')
    Assert-True $match.Success "README '$Heading' section has a PowerShell example"
    return [scriptblock]::Create([string]$match.Groups['code'].Value)
}

function Get-CommandParameterValue {
    param(
        [System.Management.Automation.Language.CommandAst]$Command,
        [string]$ParameterName
    )

    for ($index = 0; $index -lt ($Command.CommandElements.Count - 1); $index++) {
        $element = $Command.CommandElements[$index]
        if (($element -is [System.Management.Automation.Language.CommandParameterAst]) -and
            ($element.ParameterName -eq $ParameterName)) {
            $next = $Command.CommandElements[$index + 1]
            if ($next -is [System.Management.Automation.Language.CommandParameterAst]) {
                return $null
            }

            $value = [string]$next.Extent.Text
            if (($value.Length -ge 2) -and
                (($value[0] -eq '"') -or ($value[0] -eq "'")) -and
                ($value[$value.Length - 1] -eq $value[0])) {
                return $value.Substring(1, $value.Length - 2)
            }

            return $value
        }
    }

    return $null
}

function Test-CommandHasParameter {
    param(
        [System.Management.Automation.Language.CommandAst]$Command,
        [string]$ParameterName
    )

    foreach ($element in $Command.CommandElements) {
        if (($element -is [System.Management.Automation.Language.CommandParameterAst]) -and
            ($element.ParameterName -eq $ParameterName)) {
            return $true
        }
    }

    return $false
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$readme = [System.IO.File]::ReadAllText((Join-Path $repoRoot 'README.md'))
$installer = [System.IO.File]::ReadAllText((Join-Path $repoRoot 'Install-OOShutUp10.ps1'))
$taskName = Get-SourceStringAssignment -VariableName 'taskName'
$taskFolder = Get-SourceStringAssignment -VariableName 'taskFolder'
$taskPath = '\' + $taskFolder + '\'
$taskFullPath = $taskPath + $taskName
$descriptorAssignment = [regex]::Match($installer, '(?im)^\s*\$script:OOShutUpTaskSecurityDescriptorSddl\s*=\s*''(?<value>[^'']+)''\s*$')
Assert-True $descriptorAssignment.Success 'installer defines a literal task security descriptor allowlist'
$taskDescriptor = [string]$descriptorAssignment.Groups['value'].Value
$descriptorSection = Get-MarkdownSection -Heading 'Task Security Descriptor'
Assert-True ($descriptorSection.Contains(('`{0}`' -f $taskDescriptor))) 'README documents the exact installer task descriptor format'
$recoveryBlockMatch = [regex]::Match($descriptorSection, '(?ms)^```powershell\s*\r?\n(?<code>.*?)^```')
Assert-True $recoveryBlockMatch.Success 'README task security section includes a manual recovery command block'
$recoveryBlock = [string]$recoveryBlockMatch.Groups['code'].Value
$recoveryNameMatch = [regex]::Match($recoveryBlock, '(?im)^\s*\$taskName\s*=\s*''(?<value>[^'']+)''\s*$')
$recoveryPathMatch = [regex]::Match($recoveryBlock, '(?im)^\s*\$taskPath\s*=\s*''(?<value>[^'']+)''\s*$')
Assert-True $recoveryNameMatch.Success 'README recovery declares an exact task name'
Assert-True $recoveryPathMatch.Success 'README recovery declares an exact task path'
Assert-Equal $taskName ([string]$recoveryNameMatch.Groups['value'].Value) 'README recovery uses the implementation task name'
Assert-Equal $taskPath ([string]$recoveryPathMatch.Groups['value'].Value) 'README recovery uses the implementation task path'
Assert-True ($recoveryBlock.Contains('$registeredTask.SetSecurityDescriptor($savedSddl, 0x10)')) 'README recovery writes the saved owner/group/DACL descriptor'
Assert-True ($recoveryBlock.Contains('$registeredTask.GetSecurityDescriptor(7)')) 'README recovery reads owner/group/DACL back from Task Scheduler'
$vmProcedurePath = Join-Path $repoRoot 'docs/task-security-descriptor-vm-acceptance.md'
Assert-True (Test-Path -LiteralPath $vmProcedurePath -PathType Leaf) 'README disposable VM acceptance procedure exists'
$vmProcedure = [System.IO.File]::ReadAllText($vmProcedurePath)
$vmCodeBlocks = [regex]::Matches($vmProcedure, '(?ms)^```powershell\s*\r?\n(?<code>.*?)^```')
Assert-True ($vmCodeBlocks.Count -ge 3) 'VM procedure includes task registration, descriptor, mismatch, and enablement examples'
for ($blockIndex = 0; $blockIndex -lt $vmCodeBlocks.Count; $blockIndex++) {
    $vmTokens = $null
    $vmParseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput(
        [string]$vmCodeBlocks[$blockIndex].Groups['code'].Value,
        [ref]$vmTokens,
        [ref]$vmParseErrors
    )
    Assert-Equal 0 @($vmParseErrors).Count "VM procedure PowerShell block $($blockIndex + 1) parses"
}
foreach ($requiredVmEvidence in @(
    $taskFullPath,
    $taskDescriptor,
    '\Issue03Validation\',
    '\Issue03Foreign\',
    'Windows build',
    'before enablement',
    'after enablement',
    'descriptor mismatch',
    'uncertain rollback',
    'foreign same-named task',
    'Do not run this procedure on a workstation',
    'standard test account',
    'SetSecurityDescriptor',
    'DeleteTask',
    'permissive positive control',
    'does not verify a production installer run'
)) {
    Assert-True ($vmProcedure.Contains($requiredVmEvidence)) "VM procedure documents '$requiredVmEvidence'"
}

Assert-True $readme.Contains(('`{0}`' -f $taskFullPath)) 'README documents the implementation task path'
Assert-True ($readme.Contains(('The exact registered identity is `TaskPath = {0}` and `TaskName = {1}`.' -f $taskPath, $taskName))) 'README states the exact TaskPath and TaskName'
Assert-True (-not $readme.Contains('OOShutUp10-PostWindowsUpdate')) 'README does not use the unregistered report name'
Assert-True ($readme.Contains('The documented task was not found. No removal was attempted; discover the exact task path before retrying.')) 'README gives the exact not-found fallback'
Assert-True ($readme.Contains('does not start or remove tasks')) 'README marks candidate discovery as read-only'
Assert-True ($readme.Contains('Do not remove a legacy or same-named task based on a wildcard match.')) 'README forbids wildcard removal of legacy or foreign tasks'
Assert-True ($readme.Contains('This removes only the registration')) 'README explains the removal effect'
$historySection = Get-MarkdownSection -Heading 'Task History'
Assert-True ($historySection.Contains(('`{0}`' -f $taskFullPath))) 'Task Scheduler history instructions name the exact registered task path'

$commandCounts = @{
    'Get-ScheduledTask' = 0
    'Get-ScheduledTaskInfo' = 0
    'Start-ScheduledTask' = 0
    'Unregister-ScheduledTask' = 0
}
$codeBlocks = [regex]::Matches($readme, '(?ms)^```powershell\s*\r?\n(?<code>.*?)^```')
foreach ($codeBlock in $codeBlocks) {
    $prefix = $readme.Substring(0, $codeBlock.Index)
    $headings = [regex]::Matches($prefix, '(?m)^#{2,3}\s+(?<heading>.+?)\s*$')
    $heading = if ($headings.Count -eq 0) { '' } else { [string]$headings[$headings.Count - 1].Groups['heading'].Value }
    $code = [string]$codeBlock.Groups['code'].Value
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$parseErrors)
    $commands = $ast.FindAll({
        param($node)
        ($node -is [System.Management.Automation.Language.CommandAst]) -and
        ($commandCounts.ContainsKey([string]$node.GetCommandName()))
    }, $true)

    if ($commands.Count -gt 0) {
        Assert-Equal 0 @($parseErrors).Count "README task example under '$heading' parses"
    }

    foreach ($command in $commands) {
        $name = [string]$command.GetCommandName()
        $commandCounts[$name]++

        if (($heading -eq 'Safe Task Discovery') -and ($name -eq 'Get-ScheduledTask')) {
            Assert-True ($code -match '(?is)\$_.TaskName\s+-like\s+''\*OOShutUp\*''') 'candidate discovery uses a read-only name filter'
            Assert-True ($code -match '(?is)Select-Object\s+TaskPath\s*,\s*TaskName\s*,\s*State') 'candidate discovery displays both identity fields and state'
            Assert-True ($code -notmatch '(?i)\b(Start-ScheduledTask|Unregister-ScheduledTask)\b') 'candidate discovery cannot start or remove a task'
            continue
        }

        if (($heading -eq 'Removing the Scheduled Task') -and ($name -eq 'Get-ScheduledTask')) {
            Assert-True ($code -match '(?im)^\s*\$taskName\s*=\s*''(?<value>[^'']+)''\s*$') 'removal declares a literal task name'
            $removalName = [regex]::Match($code, '(?im)^\s*\$taskName\s*=\s*''(?<value>[^'']+)''\s*$').Groups['value'].Value
            Assert-Equal $taskName $removalName 'removal query uses the implementation task name'
            Assert-True ($code -match '(?im)^\s*\$taskPath\s*=\s*''(?<value>[^'']+)''\s*$') 'removal declares a literal task path'
            $removalPath = [regex]::Match($code, '(?im)^\s*\$taskPath\s*=\s*''(?<value>[^'']+)''\s*$').Groups['value'].Value
            Assert-Equal $taskPath $removalPath 'removal query uses the implementation task path'
            Assert-True ($code -match '(?is)\$_.TaskName\s+-eq\s+\$taskName\s+-and\s+\$_.TaskPath\s+-eq\s+\$taskPath') 'removal query filters by both exact identity fields'
            continue
        }

        if (($heading -eq 'Removing the Scheduled Task') -and ($name -eq 'Unregister-ScheduledTask')) {
            Assert-Equal '$taskName' (Get-CommandParameterValue -Command $command -ParameterName 'TaskName') 'removal passes the exact task name'
            Assert-Equal '$taskPath' (Get-CommandParameterValue -Command $command -ParameterName 'TaskPath') 'removal passes the exact task path'
            Assert-True (Test-CommandHasParameter -Command $command -ParameterName 'Confirm') 'removal requires explicit confirmation'
            continue
        }

        if ($name -ne 'Get-ScheduledTask') {
            Assert-Equal $taskName (Get-CommandParameterValue -Command $command -ParameterName 'TaskName') "$name uses the implementation task name under '$heading'"
            Assert-Equal $taskPath (Get-CommandParameterValue -Command $command -ParameterName 'TaskPath') "$name uses the implementation task path under '$heading'"
        }
        else {
            Assert-Equal $taskName (Get-CommandParameterValue -Command $command -ParameterName 'TaskName') "task query uses the implementation task name under '$heading'"
            Assert-Equal $taskPath (Get-CommandParameterValue -Command $command -ParameterName 'TaskPath') "task query uses the implementation task path under '$heading'"
        }
    }
}

foreach ($requiredCommand in @('Get-ScheduledTask', 'Get-ScheduledTaskInfo', 'Start-ScheduledTask', 'Unregister-ScheduledTask')) {
    Assert-True ($commandCounts[$requiredCommand] -gt 0) "README has a $requiredCommand example"
}

$script:MockSchedulerTasks = New-Object System.Collections.ArrayList
$script:MockTaskQueries = New-Object System.Collections.ArrayList
$script:MockStartCalls = New-Object System.Collections.ArrayList
$script:MockHistoryCalls = New-Object System.Collections.ArrayList
$script:MockRemovalCalls = New-Object System.Collections.ArrayList
$script:MockWarnings = New-Object System.Collections.ArrayList
$script:MockTaskQueryThrows = $false
$script:MockUnregisterThrows = $false

function Get-ScheduledTask {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath)

    if ($script:MockTaskQueryThrows) {
        throw 'mock task query failure'
    }

    [void]$script:MockTaskQueries.Add([pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath })
    $hasTaskName = $PSBoundParameters.ContainsKey('TaskName')
    $hasTaskPath = $PSBoundParameters.ContainsKey('TaskPath')
    return @($script:MockSchedulerTasks | Where-Object {
        (-not $hasTaskName -or $_.TaskName -eq $TaskName) -and
        (-not $hasTaskPath -or $_.TaskPath -eq $TaskPath)
    })
}

function Get-ScheduledTaskInfo {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath)

    [void]$script:MockHistoryCalls.Add([pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath })
    $matches = @($script:MockSchedulerTasks | Where-Object { $_.TaskName -eq $TaskName -and $_.TaskPath -eq $TaskPath })
    if ($matches.Count -ne 1) {
        throw 'mock exact task history query failed'
    }

    return [pscustomobject]@{ LastRunTime = [datetime]'2026-01-01T00:00:00'; LastTaskResult = 0 }
}

function Start-ScheduledTask {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath)

    [void]$script:MockStartCalls.Add([pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath })
    $matches = @($script:MockSchedulerTasks | Where-Object { $_.TaskName -eq $TaskName -and $_.TaskPath -eq $TaskPath })
    if ($matches.Count -ne 1) {
        throw 'mock exact task start query failed'
    }
}

function Unregister-ScheduledTask {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$TaskName, [string]$TaskPath)

    $confirmWasRequested = $PSBoundParameters.ContainsKey('Confirm')
    [void]$script:MockRemovalCalls.Add([pscustomobject]@{
        TaskName = $TaskName
        TaskPath = $TaskPath
        Confirm = $confirmWasRequested
    })

    if ($script:MockUnregisterThrows) {
        throw 'mock unregister failure'
    }

    $matches = @($script:MockSchedulerTasks | Where-Object { $_.TaskName -eq $TaskName -and $_.TaskPath -eq $TaskPath })
    if ($matches.Count -ne 1) {
        throw 'mock unregister refused a non-unique task identity'
    }

    [void]$script:MockSchedulerTasks.Remove($matches[0])
}

function Write-Warning {
    [CmdletBinding()]
    param([string]$Message)

    [void]$script:MockWarnings.Add($Message)
}

function Reset-MockSchedulerTasks {
    $script:MockSchedulerTasks.Clear()
    [void]$script:MockSchedulerTasks.Add([pscustomobject]@{
        TaskName = $taskName
        TaskPath = $taskPath
        State = 'Ready'
        Marker = 'owned'
    })
    [void]$script:MockSchedulerTasks.Add([pscustomobject]@{
        TaskName = $taskName
        TaskPath = '\Backup\'
        State = 'Ready'
        Marker = 'foreign'
    })
}

$discoveryBlock = Get-FirstPowerShellBlock -Heading 'Safe Task Discovery'
$viewBlock = Get-FirstPowerShellBlock -Heading 'Viewing the Scheduled Task'
$startBlock = Get-FirstPowerShellBlock -Heading 'Manually Running the Scheduled Task'
$historyBlock = Get-FirstPowerShellBlock -Heading 'Task History'
$removalBlock = Get-FirstPowerShellBlock -Heading 'Removing the Scheduled Task'
Reset-MockSchedulerTasks

$discovered = @(& $discoveryBlock)
Assert-Equal 2 $discovered.Count 'safe discovery lists candidates without filtering away a same-named foreign path'
Assert-Equal 0 $script:MockStartCalls.Count 'safe discovery does not start tasks'
Assert-Equal 0 $script:MockRemovalCalls.Count 'safe discovery does not remove tasks'

$viewed = @(& $viewBlock)
Assert-Equal 1 $viewed.Count 'exact task query finds one registered task'
Assert-Equal $taskName ([string]$viewed[0].TaskName) 'exact task query returns the implementation task name'
Assert-Equal $taskPath ([string]$viewed[0].TaskPath) 'exact task query returns the implementation task path'

& $startBlock | Out-Null
Assert-Equal 1 $script:MockStartCalls.Count 'manual start command calls the scheduler mock once'
Assert-Equal $taskName ([string]$script:MockStartCalls[0].TaskName) 'manual start targets the implementation task name'
Assert-Equal $taskPath ([string]$script:MockStartCalls[0].TaskPath) 'manual start targets the implementation task path'

$history = @(& $historyBlock)
Assert-Equal 1 $history.Count 'task history query returns mocked last-run information'
Assert-Equal $taskName ([string]$script:MockHistoryCalls[0].TaskName) 'history query targets the implementation task name'
Assert-Equal $taskPath ([string]$script:MockHistoryCalls[0].TaskPath) 'history query targets the implementation task path'

& $removalBlock | Out-Null
Assert-Equal 1 $script:MockRemovalCalls.Count 'removal unregisters one exact owned task'
Assert-Equal $taskName ([string]$script:MockRemovalCalls[0].TaskName) 'removal passes the implementation task name'
Assert-Equal $taskPath ([string]$script:MockRemovalCalls[0].TaskPath) 'removal passes the implementation task path'
Assert-True ([bool]$script:MockRemovalCalls[0].Confirm) 'removal requires confirmation'
Assert-Equal 1 $script:MockSchedulerTasks.Count 'removal leaves the foreign same-named task registered'
Assert-Equal '\Backup\' ([string]$script:MockSchedulerTasks[0].TaskPath) 'the task with the same name in another folder remains untouched'

$script:MockSchedulerTasks.Clear()
[void]$script:MockSchedulerTasks.Add([pscustomobject]@{ TaskName = $taskName; TaskPath = '\Backup\'; State = 'Ready'; Marker = 'foreign' })
$script:MockWarnings.Clear()
$removeCallsBeforeNotFound = $script:MockRemovalCalls.Count
& $removalBlock | Out-Null
Assert-Equal 'The documented task was not found. No removal was attempted; discover the exact task path before retrying.' ([string]$script:MockWarnings[0]) 'missing task gets the exact fallback and no removal'
Assert-Equal $removeCallsBeforeNotFound $script:MockRemovalCalls.Count 'missing exact path performs no removal'
Assert-Equal 1 $script:MockSchedulerTasks.Count 'missing exact path preserves the same-named foreign task'

Reset-MockSchedulerTasks
$script:MockTaskQueryThrows = $true
$queryFailure = $null
try {
    & $removalBlock | Out-Null
}
catch {
    $queryFailure = [string]$_.Exception.Message
}
finally {
    $script:MockTaskQueryThrows = $false
}
Assert-Equal 'Task lookup failed. No removal was attempted; verify Task Scheduler access and retry.' $queryFailure 'unexpected lookup failure stops before removal with an actionable status'
Assert-Equal 2 $script:MockSchedulerTasks.Count 'lookup failure preserves both prior mock task entries'

$script:MockUnregisterThrows = $true
$unregisterFailure = $null
try {
    & $removalBlock | Out-Null
}
catch {
    $unregisterFailure = [string]$_.Exception.Message
}
finally {
    $script:MockUnregisterThrows = $false
}
Assert-Equal 'Task removal could not be confirmed. Inspect the exact task path in Task Scheduler before retrying.' $unregisterFailure 'unexpected removal failure requires inspection before retrying'
Assert-Equal 2 $script:MockSchedulerTasks.Count 'mock removal failure preserves both prior task entries'

[Console]::Out.WriteLine("Passed $script:AssertionCount scheduled-task documentation contract and mock scheduler assertions in $($PSVersionTable.PSVersion).")
