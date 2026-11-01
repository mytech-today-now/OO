[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:LogMessages = New-Object System.Collections.ArrayList
$script:PromptLines = New-Object System.Collections.ArrayList
$script:RestartInvocations = New-Object System.Collections.ArrayList
$script:OOShutUpRestartDeferredMessage = 'Restart was deferred. Setup status remains unchanged; save your work and restart when ready.'
$script:OOShutUpRestartPromptTimeoutSeconds = 60
$script:OOShutUpRestartSaveWindowSeconds = 60

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
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Level = 'INFO'
    )

    [void]$script:LogMessages.Add("[$Level] $Message")
}

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object,
        [Parameter(Mandatory = $false)][switch]$NoNewline,
        [Parameter(Mandatory = $false)][System.ConsoleColor]$ForegroundColor
    )

    if ($null -ne $Object) {
        [void]$script:PromptLines.Add([string]$Object)
    }
}

function Reset-RestartScenario {
    param(
        [Parameter(Mandatory = $false)][AllowNull()][string]$Choice = $null,
        [Parameter(Mandatory = $false)][bool]$SaveWindowElapsed = $true,
        [Parameter(Mandatory = $false)][bool]$ReadShouldThrow = $false,
        [Parameter(Mandatory = $false)][bool]$RestartShouldThrow = $false
    )

    $script:MockRestartChoice = $Choice
    $script:MockSaveWindowElapsed = $SaveWindowElapsed
    $script:MockReadShouldThrow = $ReadShouldThrow
    $script:MockRestartShouldThrow = $RestartShouldThrow
    $script:LastPromptTimeoutSeconds = $null
    $script:LastSaveWindowSeconds = $null
    $script:LastSaveWindowCallCount = 0
    $script:LogMessages.Clear()
    $script:PromptLines.Clear()
    $script:RestartInvocations.Clear()
    $script:InstallerOutcome = [pscustomobject]@{ Status = 'Complete'; ExitCode = 0 }
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
    Assert-True ($source.Contains('[switch]$ForceRestart')) 'force restart is a separate explicit script option'
    Assert-True ($source.Contains($script:OOShutUpRestartDeferredMessage)) 'the exact defer wording preserves completed configuration status'
    Assert-True ($source.Contains("if (`$ForceRestart) { `$relocateParams += '-ForceRestart' }")) 'force opt-in survives script relocation'
    Assert-True ($source.Contains('Restart-Computer -Force')) 'the force flag is used only in the explicit force branch'
    Assert-True ($source.Contains('Configuration status is shown above. Save your work before restarting.')) 'the prompt explains the action and save opportunity'
    Assert-True ($source.Contains('Press Y to restart, N or Esc to defer. No key within {0} seconds defers the restart.')) 'the prompt is keyboard-operable and documents its timeout behavior'
    Assert-True ($source.Contains('The restart begins in {0} seconds; press any key to cancel.')) 'the save window offers a keyboard cancellation path'

    $requiredFunctions = @(
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

    $readFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Read-OOShutUpRestartKey' })[0].Extent.Text)
    $waitFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Wait-OOShutUpRestartSaveWindow' })[0].Extent.Text)
    Assert-True ($readFunctionSource.Contains('[System.Console]::KeyAvailable') -and $readFunctionSource.Contains('[System.Console]::ReadKey($true)')) 'the prompt reads keyboard input without requiring a GUI'
    Assert-True ($waitFunctionSource.Contains('[System.Console]::KeyAvailable') -and $waitFunctionSource.Contains('return $false')) 'the save window can be canceled from the keyboard'

    $restartCommands = @($scriptAst.FindAll({
        param($node)
        ($node -is [System.Management.Automation.Language.CommandAst]) -and
        ($node.GetCommandName() -eq 'Restart-Computer')
    }, $true))
    Assert-Equal 2 $restartCommands.Count 'the graceful and force restart paths are explicit separate calls'
    $forceCalls = @($restartCommands | Where-Object {
        @($_.CommandElements | Where-Object {
            ($_ -is [System.Management.Automation.Language.CommandParameterAst]) -and
            ($_.ParameterName -eq 'Force')
        }).Count -gt 0
    })
    Assert-Equal 1 $forceCalls.Count 'only one restart call supplies -Force'

    $taskSetupIndex = $source.IndexOf('$taskSetup = Invoke-OOShutUpTaskSetup', [System.StringComparison]::Ordinal)
    $outcomeIndex = $source.IndexOf('$installerOutcome = Get-OOShutUpInstallOutcome', [System.StringComparison]::Ordinal)
    $summaryIndex = $source.IndexOf('Write-OOShutUpInstallSummary -Outcome $installerOutcome', [System.StringComparison]::Ordinal)
    $restartIndex = $source.IndexOf('Invoke-OOShutUpRestart -NonInteractive:', [System.StringComparison]::Ordinal)
    Assert-True ($taskSetupIndex -ge 0 -and $outcomeIndex -gt $taskSetupIndex -and $summaryIndex -gt $outcomeIndex -and $restartIndex -gt $summaryIndex) 'settings and task outcomes are recorded before the restart offer'

    Assert-Equal $null (Read-OOShutUpRestartKey -TimeoutSeconds 0) 'a zero-length confirmation timeout returns no input without restarting'
    Assert-Equal $true (Wait-OOShutUpRestartSaveWindow -Seconds 0) 'an elapsed save window permits the already-confirmed restart'

    Set-Item -Path Function:Read-OOShutUpRestartKey -Value {
        param([int]$TimeoutSeconds)
        $script:LastPromptTimeoutSeconds = $TimeoutSeconds
        if ($script:MockReadShouldThrow) {
            throw 'Mock input timeout.'
        }
        return $script:MockRestartChoice
    }
    Set-Item -Path Function:Wait-OOShutUpRestartSaveWindow -Value {
        param([int]$Seconds)
        $script:LastSaveWindowCallCount++
        $script:LastSaveWindowSeconds = $Seconds
        return $script:MockSaveWindowElapsed
    }
    Set-Item -Path Function:Restart-Computer -Value {
        [CmdletBinding()]
        param([switch]$Force)

        $boundKeys = @($PSBoundParameters.Keys | Sort-Object)
        [void]$script:RestartInvocations.Add([pscustomobject]@{
                Force = [bool]$Force
                BoundKeys = $boundKeys
            })
        if ($script:MockRestartShouldThrow) {
            throw 'Mock graceful restart was blocked.'
        }
    }

    Reset-RestartScenario -Choice 'Y'
    Invoke-OOShutUpRestart
    Assert-Equal 1 $script:RestartInvocations.Count 'ordinary Yes requests one restart'
    Assert-Equal $false $script:RestartInvocations[0].Force 'ordinary Yes never force-closes applications'
    Assert-Equal 0 $script:RestartInvocations[0].BoundKeys.Count 'ordinary Yes supplies no Restart-Computer arguments'
    Assert-Equal 60 $script:LastPromptTimeoutSeconds 'the confirmation prompt has a 60-second timeout'
    Assert-Equal 60 $script:LastSaveWindowSeconds 'the save window lasts 60 seconds'
    Assert-True ($script:PromptLines -contains 'Press Y to restart, N or Esc to defer. No key within 60 seconds defers the restart.') 'the visible prompt explains the keyboard choices and timeout'
    Assert-True ($script:PromptLines -contains 'Save your work now. The restart begins in 60 seconds; press any key to cancel.') 'the visible save window explains how to cancel'
    Assert-Equal 'Complete' $script:InstallerOutcome.Status 'the completed configuration status remains intact after a restart request'

    Reset-RestartScenario -Choice 'N'
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'No defers without calling Restart-Computer'
    Assert-Equal 0 $script:LastSaveWindowCallCount 'No does not start the save window'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'No reports the exact deferred status'

    Reset-RestartScenario -Choice 'Escape'
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'Escape defers without calling Restart-Computer'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'Escape reports the exact deferred status'

    Reset-RestartScenario -Choice 'Y' -SaveWindowElapsed $false
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'cancellation during the save window prevents restart'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'cancellation reports the exact deferred status'
    Assert-Equal 'Complete' $script:InstallerOutcome.Status 'cancellation leaves completed configuration status intact'

    Reset-RestartScenario -Choice $null
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'missing input or prompt timeout never restarts'
    Assert-Equal 0 $script:LastSaveWindowCallCount 'missing input never starts the save window'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'timeout reports the exact deferred status'

    Reset-RestartScenario -Choice $null -ReadShouldThrow $true
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'unavailable keyboard input never restarts'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'unavailable keyboard input reports the exact deferred status'

    Reset-RestartScenario -Choice 'Y'
    Invoke-OOShutUpRestart -ForceRestart
    Assert-Equal 1 $script:RestartInvocations.Count 'explicit force opt-in requests one restart'
    Assert-Equal $true $script:RestartInvocations[0].Force 'explicit -ForceRestart supplies the Force switch'
    Assert-Equal 'Force' ([string]$script:RestartInvocations[0].BoundKeys[0]) 'force restart supplies exactly the -Force argument'
    Assert-True ($script:LogMessages -contains '[WARNING] The explicit -ForceRestart option can close applications that have unsaved work.') 'force opt-in warns that unsaved work can be closed'

    Reset-RestartScenario -Choice 'Y'
    Invoke-OOShutUpRestart -NonInteractive -ForceRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'non-interactive mode never restarts'
    Assert-Equal $null $script:LastPromptTimeoutSeconds 'non-interactive mode does not request input'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'non-interactive mode reports the exact deferred status'

    Reset-RestartScenario -Choice 'Y' -RestartShouldThrow $true
    Invoke-OOShutUpRestart
    Assert-Equal 1 $script:RestartInvocations.Count 'a graceful restart attempt is made after confirmation and the save window'
    Assert-True ($script:LogMessages -contains '[WARNING] Restart could not be started: Mock graceful restart was blocked.') 'restart failure has an exact user-facing status'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'restart failure leaves a clear recovery reminder'
    Assert-Equal 0 $script:InstallerOutcome.ExitCode 'restart failure does not change the completed installer outcome'

    [Console]::Out.WriteLine("PASS: Restart safety mock tests ($script:AssertionCount assertions) in $($PSVersionTable.PSVersion).")
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
