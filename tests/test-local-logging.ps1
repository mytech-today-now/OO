[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:CapturedLines = New-Object System.Collections.ArrayList

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

# Capture local console output without running the elevated script entrypoint.
function Write-Host {
    param(
        [Parameter(Position = 0)]
        [object]$Object,
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
$scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
    $scriptPath,
    [ref]$tokens,
    [ref]$parseErrors
)

Assert-Equal 0 @($parseErrors).Count 'the installer parses in the active PowerShell engine'
$source = Get-Content -LiteralPath $scriptPath -Raw

Assert-True ($source -notmatch '(?i)Invoke-Expression|\biex\b|\[scriptblock\]::Create') 'the installer contains no dynamic evaluator'
Assert-True ($source -notmatch '(?i)Invoke-WebRequest|logging\.ps1|raw\.githubusercontent\.com') 'the logger has no mutable remote loader'
Assert-True ($source -notmatch '(?i)\.ScriptStackTrace|Stack Trace:|Stack trace:') 'the installer does not log raw exception stacks'
Assert-True ($source.Contains('Trusted logging validation failed. No configuration changes were started; use the packaged local logger and review the logged stage.')) 'the safe initialization failure message is present'

$functionAsts = $scriptAst.FindAll({
    param($node)
    ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
    (@('Write-Log', 'Initialize-LogFile') -contains $node.Name)
}, $true)
$writeLogAst = @($functionAsts | Where-Object { $_.Name -eq 'Write-Log' })[0]
$initializeAst = @($functionAsts | Where-Object { $_.Name -eq 'Initialize-LogFile' })[0]
Assert-True ($null -ne $writeLogAst) 'the packaged Write-Log function exists'
Assert-True ($null -ne $initializeAst) 'the packaged Initialize-LogFile function exists'

$usedLevels = New-Object System.Collections.ArrayList
$logCalls = $scriptAst.FindAll({
    param($node)
    ($node -is [System.Management.Automation.Language.CommandAst]) -and
    ($node.GetCommandName() -eq 'Write-Log')
}, $true)
foreach ($call in $logCalls) {
    $elements = $call.CommandElements
    for ($index = 0; $index -lt ($elements.Count - 1); $index++) {
        if (($elements[$index] -is [System.Management.Automation.Language.CommandParameterAst]) -and
            ($elements[$index].ParameterName -eq 'Level')) {
            [void]$usedLevels.Add($elements[$index + 1].Extent.Text.Trim("'", '"'))
        }
    }
}

foreach ($requiredLevel in @('INFO', 'SUCCESS', 'WARNING', 'ERROR')) {
    Assert-True ($usedLevels -contains $requiredLevel) "Write-Log supports the used level $requiredLevel"
}

# Exercise only the two local function definitions. The script's admin-only main flow is never run.
. ([scriptblock]::Create($writeLogAst.Extent.Text))
. ([scriptblock]::Create($initializeAst.Extent.Text))

foreach ($level in @('INFO', 'SUCCESS', 'WARNING', 'ERROR')) {
    Write-Log -Message "Local logger check $level" -Level $level
}
foreach ($level in @('INFO', 'SUCCESS', 'WARNING', 'ERROR')) {
    Assert-True (@($script:CapturedLines | Where-Object { $_ -match "\[$level\] Local logger check $level" }).Count -eq 1) "Write-Log emits $level messages"
}

# These response mocks cover the former remote boundary. The packaged logger must never request or evaluate them.
$responseBodySentinel = '$script:SentinelExecutions++'
$script:SentinelExecutions = 0
$script:FetchCount = 0
$script:MockScenario = $null
function Invoke-WebRequest {
    $script:FetchCount++
    if ($script:MockScenario.Throws) {
        throw $script:MockScenario.Error
    }
    return $script:MockScenario.Response
}
function Invoke-Expression {
    param([string]$Command)
    $script:SentinelExecutions++
    ([scriptblock]::Create($Command)).Invoke()
}

$responseCases = @(
    @{ Name = 'approved'; Throws = $false; Response = [pscustomobject]@{ Content = $responseBodySentinel } },
    @{ Name = 'tampered'; Throws = $false; Response = [pscustomobject]@{ Content = 'SENTINEL_RESPONSE_BODY_DO_NOT_LOG; SENTINEL_TOKEN_DO_NOT_LOG' } },
    @{ Name = 'missing'; Throws = $false; Response = $null },
    @{ Name = 'throwing'; Throws = $true; Error = 'SENTINEL_RAW_STACK_DO_NOT_LOG' }
)

foreach ($responseCase in $responseCases) {
    $script:CapturedLines.Clear()
    $script:MockScenario = $responseCase
    Assert-Equal $true (Initialize-LogFile) "local logging initializes with a $($responseCase.Name) response mock"
    Assert-Equal 0 $script:FetchCount "the logger does not fetch a $($responseCase.Name) response"
    Assert-Equal 0 $script:SentinelExecutions "the logger never executes a $($responseCase.Name) response body"
    $loggedText = [string]::Join("`n", [string[]]$script:CapturedLines.ToArray())
    Assert-True ($loggedText -notmatch 'SENTINEL_RESPONSE_BODY_DO_NOT_LOG|SENTINEL_TOKEN_DO_NOT_LOG|SENTINEL_RAW_STACK_DO_NOT_LOG') "the logger does not expose $($responseCase.Name) response data"
}

# A local logger failure is reported safely, and the entrypoint gate precedes relocation and configuration work.
$savedWriteLog = (Get-Command -Name Write-Log -CommandType Function).ScriptBlock
Set-Item -Path Function:\Write-Log -Value { throw 'SENTINEL_TOKEN_DO_NOT_LOG SENTINEL_RAW_STACK_DO_NOT_LOG' }
$script:CapturedLines.Clear()
$initializeResult = Initialize-LogFile
Assert-Equal $false $initializeResult 'a throwing local logger reports initialization failure'
$failureMessage = 'Trusted logging validation failed. No configuration changes were started; use the packaged local logger and review the logged stage.'
$failureOutput = [string]::Join("`n", [string[]]$script:CapturedLines.ToArray())
Assert-True ($failureOutput.Contains($failureMessage)) 'logger initialization failure reports the safe fallback message'
Assert-True ($failureOutput -notmatch 'SENTINEL_TOKEN_DO_NOT_LOG|SENTINEL_RAW_STACK_DO_NOT_LOG') 'logger initialization failure omits exception details'
Set-Item -Path Function:\Write-Log -Value $savedWriteLog

$guardStart = $source.IndexOf('if (-not (Initialize-LogFile))')
$firstOperationalStage = $source.IndexOf('# Step 0: Ensure script is running from default location')
Assert-True ($guardStart -ge 0 -and $firstOperationalStage -gt $guardStart) 'the local logger gate precedes relocation'
$guardText = $source.Substring($guardStart, $firstOperationalStage - $guardStart)
Assert-True ($guardText.Contains('exit 1')) 'logger initialization failure exits before operational work'
foreach ($operation in @('Install-OOShutUp', 'Invoke-OOShutUpConfigurationWithRestorePointGate', 'Invoke-OOShutUpConfiguration', 'Invoke-OOShutUpTaskSetup')) {
    $operationPosition = $source.IndexOf($operation, $firstOperationalStage)
    Assert-True ($operationPosition -gt $firstOperationalStage) "$operation remains after the logger gate"
}

[Console]::Out.WriteLine("Passed $script:AssertionCount local logging assertions.")
