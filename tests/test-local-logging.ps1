[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:CapturedLines = New-Object System.Collections.ArrayList
$script:FixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oosu10-local-logging-' + [guid]::NewGuid().ToString('N'))
$script:MockAclDenied = $false
$script:OOShutUpRunId = [guid]::NewGuid().ToString('N')
$script:OOShutUpApprovedDigest = $null
$script:OOShutUpApprovedManifestDigest = $null

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
    $script:AssertionCount++
}

function Assert-Equal {
    param([object]$Expected, [object]$Actual, [string]$Message)
    if ($Expected -cne $Actual) { throw "Assertion failed: $Message (expected '$Expected', got '$Actual')" }
    $script:AssertionCount++
}

function Write-Host {
    param(
        [Parameter(Position = 0)] [object]$Object,
        [switch]$NoNewline,
        [string]$Separator = ' ',
        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor
    )
    [void]$script:CapturedLines.Add([string]$Object)
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'Install-OOShutUp10.ps1'
$tokens = $null
$parseErrors = $null
$scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
Assert-Equal 0 @($parseErrors).Count 'the installer parses in the active PowerShell engine'
$source = Get-Content -LiteralPath $scriptPath -Raw

Assert-True ($source -notmatch '(?i)Invoke-Expression|\biex\b|\[scriptblock\]::Create') 'the installer contains no dynamic evaluator'
Assert-True ($source -notmatch '(?i)Invoke-WebRequest|logging\.ps1|raw\.githubusercontent\.com') 'the protected logger has no mutable remote loader'
Assert-True ($source -notmatch '(?i)\.ScriptStackTrace|Stack Trace:|Stack trace:') 'the installer does not persist raw exception stacks'
Assert-True ($source.Contains("'S-1-5-18'") -and $source.Contains("'S-1-5-32-544'")) 'the log ACL names only SYSTEM and Builtin Administrators'
Assert-True ($source.Contains('SetAccessRuleProtection($true, $false)')) 'log ACLs disable inheritance and discard inherited access rules'
Assert-True ($source.Contains("'myTech.Today-OOShutUp10\Logs'")) 'the log directory is under machine-wide application data'
Assert-True ($source.Contains('$script:OOShutUpLogMaxBytes = 1048576') -and $source.Contains('$script:OOShutUpLogRetainedFileCount = 3')) 'the log rotation has a fixed byte and file bound'
Assert-True ($source.Contains("-Outcome Complete -ExitCode 0")) 'the hidden reapply success path writes a terminal result before exit 0'
Assert-True ($source.Contains('if ($script:OOShutUpLogWriteFailed -or -not $script:OOShutUpLogReady)')) 'terminal completion is blocked after any persistent logger failure'
Assert-True ($source.Contains('-WindowStyle Hidden')) 'the scheduled task action remains hidden'

$functionNames = @(
    'Get-OOShutUpLogDirectoryPath',
    'Set-OOShutUpProtectedLogAcl',
    'Test-OOShutUpProtectedLogAcl',
    'Initialize-OOShutUpLogDirectory',
    'Open-OOShutUpLogWriterLock',
    'Set-OOShutUpLogStage',
    'Set-OOShutUpLogErrorCategory',
    'Set-OOShutUpLogOutcome',
    'Initialize-OOShutUpLogFile',
    'Invoke-OOShutUpLogRotation',
    'Write-OOShutUpAuditRecord',
    'Write-Log',
    'Initialize-LogFile',
    'Get-OOShutUpTerminalExitCode'
)
$functionAsts = $scriptAst.FindAll({
    param($node)
    ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($functionNames -contains $node.Name)
}, $true)
foreach ($name in $functionNames) {
    $functionAst = @($functionAsts | Where-Object { $_.Name -eq $name })[0]
    Assert-True ($null -ne $functionAst) "the $name function exists"
    . ([scriptblock]::Create($functionAst.Extent.Text))
}

# The fixture uses a temporary directory and mocks ACL setup. Effective Windows ACL behavior is a separate VM check.
function Initialize-OOShutUpLogDirectory {
    if ($script:MockAclDenied) {
        throw [System.UnauthorizedAccessException]::new('SENTINEL_ACL_DENIAL_DETAIL')
    }
    [void][System.IO.Directory]::CreateDirectory($script:OOShutUpLogDirectoryPath)
    $script:OOShutUpLogFilePath = Join-Path $script:OOShutUpLogDirectoryPath 'Install-OOShutUp10.jsonl'
}

function Set-OOShutUpProtectedLogAcl {
    param([string]$Path, [switch]$Directory)
}

function Test-OOShutUpProtectedLogAcl {
    param([string]$Path, [switch]$Directory)
    return $true
}

function Reset-TestLog {
    param([string]$Name)
    $script:OOShutUpLogDirectoryPath = Join-Path $script:FixtureRoot $Name
    $script:OOShutUpLogFilePath = Join-Path $script:OOShutUpLogDirectoryPath 'Install-OOShutUp10.jsonl'
    $script:OOShutUpLogLockFilePath = Join-Path $script:OOShutUpLogDirectoryPath 'OOShutUp10Log.lock'
    $script:OOShutUpLogMaxBytes = 900
    $script:OOShutUpLogRetainedFileCount = 2
    $script:OOShutUpLogReady = $false
    $script:OOShutUpLogWriteFailed = $false
    $script:OOShutUpLogStage = 'Startup'
    $script:OOShutUpLogErrorCategory = 'None'
    $script:OOShutUpLogNativeResult = $null
    $script:OOShutUpLogOutcome = 'InProgress'
    $script:OOShutUpApprovedDigest = $null
    $script:OOShutUpApprovedManifestDigest = $null
    $script:CapturedLines.Clear()
    $script:MockAclDenied = $false
}

function Get-CurrentLogRecords {
    if (-not [System.IO.File]::Exists($script:OOShutUpLogFilePath)) { return @() }
    $lines = @([System.IO.File]::ReadAllLines($script:OOShutUpLogFilePath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return @($lines | ForEach-Object { ConvertFrom-Json -InputObject $_ -ErrorAction Stop })
}

[void][System.IO.Directory]::CreateDirectory($script:FixtureRoot)
Reset-TestLog -Name 'success-and-redaction'
Assert-Equal $true (Initialize-LogFile) 'a protected structured log initializes in a temporary directory'
Assert-True ([System.IO.File]::Exists($script:OOShutUpLogFilePath)) 'initialization creates the current structured log'

$script:OOShutUpApprovedDigest = 'A' * 64
$script:OOShutUpApprovedManifestDigest = 'B' * 64
Set-OOShutUpLogStage -Stage ProfileApproval
Write-Log -Message 'SENTINEL_SETTING_VALUE cfg-content=SENTINEL_CFG token=SENTINEL_TOKEN credential=SENTINEL_CREDENTIAL' -Level INFO
$records = Get-CurrentLogRecords
$record = $records[-1]
Assert-Equal 'ProfileApproval' $record.stage 'the log records the current safe stage'
Assert-Equal 'InProgress' $record.outcome 'ordinary stage records do not claim run completion'
Assert-Equal $script:OOShutUpRunId $record.runId 'the log records the per-run identifier'
Assert-Equal ('A' * 64) $record.approvedDigests.profileSha256 'the log records the approved profile digest'
Assert-Equal ('B' * 64) $record.approvedDigests.manifestSha256 'the log records the approved manifest digest'
$logText = [System.IO.File]::ReadAllText($script:OOShutUpLogFilePath)
Assert-True ($logText -notmatch 'SENTINEL_SETTING_VALUE|SENTINEL_CFG|SENTINEL_TOKEN|SENTINEL_CREDENTIAL|cfg-content') 'profile values, cfg content, tokens, and credentials are omitted'
Assert-True ($record.PSObject.Properties.Name -contains 'timestampUtc') 'the log records a UTC timestamp'
Assert-True ($record.PSObject.Properties.Name -contains 'nativeResult') 'the log record has a native result field'

Set-OOShutUpLogStage -Stage ProfileValidation
Set-OOShutUpLogErrorCategory -Category ProfileRejected
Set-OOShutUpLogOutcome -Outcome Failed
Write-Log -Message 'SENTINEL_REJECTED_PROFILE_VALUE' -Level ERROR
$record = (Get-CurrentLogRecords)[-1]
Assert-Equal 'ProfileValidation' $record.stage 'profile rejection identifies the validation stage'
Assert-Equal 'ProfileRejected' $record.errorCategory 'profile rejection uses a safe category'
Assert-Equal 'Failed' $record.outcome 'profile rejection records a failed outcome'
Assert-True (-not [System.IO.File]::ReadAllText($script:OOShutUpLogFilePath).Contains('SENTINEL_REJECTED_PROFILE_VALUE')) 'profile rejection does not persist the profile value'

Set-OOShutUpLogStage -Stage OOSProcessWait
Set-OOShutUpLogErrorCategory -Category Timeout
Set-OOShutUpLogOutcome -Outcome Unknown
Write-Log -Message 'SENTINEL_TIMEOUT_DETAIL' -Level ERROR
$record = (Get-CurrentLogRecords)[-1]
Assert-Equal 'OOSProcessWait' $record.stage 'timeout identifies the child wait stage'
Assert-Equal 'Timeout' $record.errorCategory 'timeout uses a safe category'
Assert-Equal 'Unknown' $record.outcome 'timeout records an unknown settings outcome'
Assert-Equal $null $record.nativeResult 'timeout has no fabricated native result'
Assert-True (-not [System.IO.File]::ReadAllText($script:OOShutUpLogFilePath).Contains('SENTINEL_TIMEOUT_DETAIL')) 'timeout detail is not persisted'

Set-OOShutUpLogStage -Stage OOSProcessResult
Set-OOShutUpLogErrorCategory -Category None -NativeResult 30
Set-OOShutUpLogOutcome -Outcome InProgress
Write-Log -Message 'O&O returned an approved native status.' -Level INFO
$record = (Get-CurrentLogRecords)[-1]
Assert-Equal 30 $record.nativeResult 'the log records a numeric native process result'

Reset-TestLog -Name 'rotation'
Assert-Equal $true (Initialize-LogFile) 'rotation fixture initializes'
$script:OOShutUpApprovedDigest = 'C' * 64
$script:OOShutUpApprovedManifestDigest = 'D' * 64
for ($index = 1; $index -le 20; $index++) {
    Set-OOShutUpLogStage -Stage ProfileValidation
    Write-Log -Message ("bounded rotation entry {0}" -f $index) -Level INFO
}
Set-OOShutUpLogStage -Stage TaskSetup
Set-OOShutUpLogErrorCategory -Category TaskScheduler
Set-OOShutUpLogOutcome -Outcome Failed
Write-Log -Message 'SENTINEL_CURRENT_FAILURE_DETAIL' -Level ERROR
$retainedLogs = @(Get-ChildItem -LiteralPath $script:OOShutUpLogDirectoryPath -Filter 'Install-OOShutUp10.jsonl*' -File)
Assert-True ($retainedLogs.Count -le 3) 'rotation retains only the current log and two backups'
foreach ($retainedLog in $retainedLogs) {
    Assert-True ($retainedLog.Length -le $script:OOShutUpLogMaxBytes) 'every retained log stays within the configured byte bound'
}
$record = (Get-CurrentLogRecords)[-1]
Assert-Equal 'TaskSetup' $record.stage 'the newest failure remains in the current log'
Assert-Equal 'TaskScheduler' $record.errorCategory 'the current failure category survives rotation'
Assert-Equal 'Failed' $record.outcome 'the current failure outcome survives rotation'
Assert-True (-not [System.IO.File]::ReadAllText($script:OOShutUpLogFilePath).Contains('SENTINEL_CURRENT_FAILURE_DETAIL')) 'rotation does not retain arbitrary error text'

$terminalCode = Get-OOShutUpTerminalExitCode -Outcome Complete -ExitCode 0
Assert-Equal 0 $terminalCode 'a successful terminal record returns exit code zero'
$record = (Get-CurrentLogRecords)[-1]
Assert-Equal 'Complete' $record.outcome 'the final durable record records completion only after successful logging'
Assert-Equal $null $record.nativeResult 'the terminal record does not label an installer exit status as a native result'

Reset-TestLog -Name 'disk-full'
Assert-Equal $true (Initialize-LogFile) 'disk-full fixture initializes before failure injection'
$savedAuditWriter = (Get-Command -Name Write-OOShutUpAuditRecord -CommandType Function).ScriptBlock
Set-Item -Path Function:\Write-OOShutUpAuditRecord -Value { throw [System.IO.IOException]::new('SENTINEL_DISK_FULL_DETAIL') }
Write-Log -Message 'SENTINEL_DISK_FULL_DETAIL' -Level SUCCESS
Assert-Equal $true $script:OOShutUpLogWriteFailed 'an append failure marks the logger unavailable'
Assert-Equal 1 (Get-OOShutUpTerminalExitCode -Outcome Complete -ExitCode 0) 'a disk-full logging failure cannot return a complete result'
$records = Get-CurrentLogRecords
Assert-True (@($records | Where-Object { $_.outcome -eq 'Complete' }).Count -eq 0) 'a disk-full failure does not create a Complete record'
$fallbackText = [string]::Join("`n", [string[]]$script:CapturedLines.ToArray())
Assert-True ($fallbackText.Contains('Details suppressed')) 'the console fallback is safe and suppresses error details'
Assert-True ($fallbackText -notmatch 'SENTINEL_DISK_FULL_DETAIL') 'the console fallback omits disk error details'
Set-Item -Path Function:\Write-OOShutUpAuditRecord -Value $savedAuditWriter

Reset-TestLog -Name 'acl-denied'
$script:MockAclDenied = $true
Assert-Equal $false (Initialize-LogFile) 'an ACL-denied logger fails initialization'
Assert-Equal $true $script:OOShutUpLogWriteFailed 'an ACL denial marks logging as failed'
$fallbackText = [string]::Join("`n", [string[]]$script:CapturedLines.ToArray())
Assert-True ($fallbackText.Contains('No configuration changes were started')) 'ACL denial stops before configuration changes'
Assert-True ($fallbackText -notmatch 'SENTINEL_ACL_DENIAL_DETAIL') 'ACL fallback omits exception details'

$guardStart = $source.IndexOf('if (-not (Initialize-LogFile))')
$firstOperationalStage = $source.IndexOf('# Step 0: Ensure script is running from default location')
Assert-True ($guardStart -ge 0 -and $firstOperationalStage -gt $guardStart) 'the protected logger gate precedes relocation'
$guardText = $source.Substring($guardStart, $firstOperationalStage - $guardStart)
Assert-True ($guardText.Contains('exit 1')) 'logger initialization failure exits before operational work'
foreach ($operation in @('Install-OOShutUp', 'Invoke-OOShutUpConfigurationWithRestorePointGate', 'Invoke-OOShutUpConfiguration', 'Invoke-OOShutUpTaskSetup')) {
    $operationPosition = $source.IndexOf($operation, $firstOperationalStage)
    Assert-True ($operationPosition -gt $firstOperationalStage) "$operation remains after the logger gate"
}

[Console]::Out.WriteLine("Passed $script:AssertionCount local logging assertions.")
if (Test-Path -LiteralPath $script:FixtureRoot) {
    $resolvedFixtureRoot = [System.IO.Path]::GetFullPath($script:FixtureRoot)
    $resolvedTempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $resolvedFixtureRoot.StartsWith($resolvedTempRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'The logging fixture cleanup path is outside the system temporary directory.'
    }
    Remove-Item -LiteralPath $resolvedFixtureRoot -Recurse -Force
}
