[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:AssertionCount = 0
$script:LogMessages = New-Object System.Collections.ArrayList
$script:PromptLines = New-Object System.Collections.ArrayList
$script:RestartInvocations = New-Object System.Collections.ArrayList
$script:OOShutUpRestartDeferredMessage = 'Restart was deferred. Setup status remains unchanged; save your work and restart when ready.'
$script:OOShutUpRestartPromptUnavailableMessage = 'No interactive restart prompt is available. Configuration status is retained; restart manually when convenient.'
$script:OOShutUpRestartPromptTimeoutSeconds = 60
$script:OOShutUpRestartSaveWindowSeconds = 60
$script:OOShutUpReleaseManifestFileName = 'Install-OOShutUp10.release.psd1'

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
    Assert-True ($source.Contains($script:OOShutUpRestartPromptUnavailableMessage)) 'unavailable restart input has the required recovery wording'
    Assert-True ($source.Contains("if (`$ForceRestart) { `$relocateParams += '-ForceRestart' }")) 'force opt-in survives script relocation'
    foreach ($forwardedArgument in @(
        "`$relocateParams = @('-SkipRelocate')",
        "if (`$ReapplyOnly) { `$relocateParams += '-ReapplyOnly' }",
        "if (`$NonInteractive) { `$relocateParams += '-NonInteractive' }",
        "if (`$AllowWithoutRestorePoint) { `$relocateParams += '-AllowWithoutRestorePoint' }",
        "'-ApprovedProfileSha256', `$script:OOShutUpApprovedDigest",
        "'-ApprovedProfileManifestSha256', `$script:OOShutUpApprovedManifestDigest",
        "'-ProfileApprovalConfirmed'"
    )) {
        Assert-True ($source.Contains($forwardedArgument)) "relocation preserves $forwardedArgument"
    }
    Assert-True ($source.Contains('Restart-Computer -Force')) 'the force flag is used only in the explicit force branch'
    Assert-True ($source.Contains('Configuration status is shown above. Save your work before restarting.')) 'the prompt explains the action and save opportunity'
    Assert-True ($source.Contains('Press Y to restart, N or Esc to defer. No key within {0} seconds defers the restart.')) 'the prompt is keyboard-operable and documents its timeout behavior'
    Assert-True ($source.Contains('The restart begins in {0} seconds; press any key to cancel.')) 'the save window offers a keyboard cancellation path'

    $requiredFunctions = @(
        'Test-OOShutUpInteractiveInput',
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

    $relocationFunctionNames = @(
        'Get-ScriptRelocationPackageTarget',
        'Test-ScriptRelocationPackage',
        'Move-ScriptRelocationDirectory',
        'Remove-ScriptRelocationStage',
        'Install-ScriptRelocationPackage',
        'Invoke-ScriptRelocation'
    )
    $relocationFunctionAsts = @($scriptAst.FindAll({
        param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
        ($relocationFunctionNames -contains $node.Name)
    }, $true))
    foreach ($name in $relocationFunctionNames) {
        $functionAst = @($relocationFunctionAsts | Where-Object { $_.Name -eq $name })[0]
        Assert-True ($null -ne $functionAst) "the $name relocation function exists"
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }
    $relocationSource = [string](@($relocationFunctionAsts | Where-Object { $_.Name -eq 'Invoke-ScriptRelocation' })[0].Extent.Text)
    $packageInstallerSource = [string](@($relocationFunctionAsts | Where-Object { $_.Name -eq 'Install-ScriptRelocationPackage' })[0].Extent.Text)
    Assert-True ($packageInstallerSource.Contains("'.stage-{0}'") -and $packageInstallerSource.Contains('Move-ScriptRelocationDirectory -Source $stageRoot -Destination $targetRoot')) 'relocation stages a unique sibling package and promotes it with one directory rename'
    Assert-True ($packageInstallerSource.Contains('Test-ScriptRelocationPackage -PackageRoot $stageRoot') -and $packageInstallerSource.Contains('Test-ScriptRelocationPackage -PackageRoot $targetRoot')) 'staged and promoted packages both require digest and ACL verification'
    Assert-True (-not $packageInstallerSource.Contains('Register-ScheduledTask')) 'package staging cannot register a task before activation'
    Assert-True ($relocationSource.Contains("`$relocateParams = @('-SkipRelocate')") -and $relocationSource.Contains("'-ApprovedProfileManifestSha256', `$script:OOShutUpApprovedManifestDigest")) 'the relocated child keeps SkipRelocate and both approved profile digests'

    $interactiveInputSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Test-OOShutUpInteractiveInput' })[0].Extent.Text)
    $readFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Read-OOShutUpRestartKey' })[0].Extent.Text)
    $restartHandlerSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Invoke-OOShutUpRestart' })[0].Extent.Text)
    $waitFunctionSource = [string](@($functionAsts | Where-Object { $_.Name -eq 'Wait-OOShutUpRestartSaveWindow' })[0].Extent.Text)
    Assert-True ($interactiveInputSource.Contains('[System.Console]::IsInputRedirected')) 'redirected input is detected before profile approval'
    Assert-True (-not $restartHandlerSource.Contains('Read-Host')) 'restart handling does not use a potentially unbounded Read-Host call'
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
    $nonInteractiveDetectionIndex = $source.IndexOf('Test-OOShutUpInteractiveInput)', [System.StringComparison]::Ordinal)
    $profileApprovalIndex = $source.IndexOf('$profileApprovalAccepted = Confirm-OOShutUpProfile', [System.StringComparison]::Ordinal)
    $outcomeIndex = $source.IndexOf('$installerOutcome = Get-OOShutUpInstallOutcome', [System.StringComparison]::Ordinal)
    $summaryIndex = $source.IndexOf('Write-OOShutUpInstallSummary -Outcome $installerOutcome', [System.StringComparison]::Ordinal)
    $restartIndex = $source.IndexOf('Invoke-OOShutUpRestart -NonInteractive:', [System.StringComparison]::Ordinal)
    Assert-True ($nonInteractiveDetectionIndex -ge 0 -and $profileApprovalIndex -gt $nonInteractiveDetectionIndex) 'redirected input selects non-interactive approval before any profile prompt'
    Assert-True ($taskSetupIndex -ge 0 -and $outcomeIndex -gt $taskSetupIndex -and $summaryIndex -gt $outcomeIndex -and $restartIndex -gt $summaryIndex) 'settings and task outcomes are recorded before the restart offer'

    Assert-Equal $null (Read-OOShutUpRestartKey -TimeoutSeconds 0) 'a zero-length confirmation timeout returns no input without restarting'
    Assert-Equal $true (Wait-OOShutUpRestartSaveWindow -Seconds 0) 'an elapsed save window permits the already-confirmed restart'
    Assert-Equal ([bool](-not [System.Console]::IsInputRedirected)) ([bool](Test-OOShutUpInteractiveInput)) 'interactive capability follows redirected standard input and fails closed'
    if ([System.Console]::IsInputRedirected) {
        $redirectedInputRejected = $false
        try {
            Read-OOShutUpRestartKey -TimeoutSeconds 1 | Out-Null
        }
        catch {
            $redirectedInputRejected = $true
        }
        Assert-True $redirectedInputRejected 'redirected input or EOF exits the keyboard prompt immediately for the recovery fallback'
    }
    else {
        [Console]::Out.WriteLine('SKIP: Direct redirected-input prompt check requires a process with redirected standard input.')
    }

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

    Reset-RestartScenario -Choice ''
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'empty input never restarts'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'empty input reports that restart was deferred'

    Reset-RestartScenario -Choice $null -ReadShouldThrow $true
    $script:InstallerOutcome = [pscustomobject]@{
        Status = 'Partial'
        ExitCode = 2
        Summary = 'Settings were verified; the update task was not verified.'
    }
    Invoke-OOShutUpRestart
    Assert-Equal 0 $script:RestartInvocations.Count 'unavailable keyboard input never restarts'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartPromptUnavailableMessage") 'a thrown prompt reports the exact restart fallback'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'unavailable keyboard input reports the exact deferred status'
    Assert-Equal 'Partial' $script:InstallerOutcome.Status 'a prompt failure retains the previously reported partial setup status'
    Assert-Equal 2 $script:InstallerOutcome.ExitCode 'a prompt failure retains the setup recovery exit code'
    Assert-Equal 'Settings were verified; the update task was not verified.' $script:InstallerOutcome.Summary 'a prompt failure retains the setup summary'

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
    Assert-Equal 0 $script:PromptLines.Count 'non-interactive mode does not display a restart prompt'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartPromptUnavailableMessage") 'non-interactive mode reports the exact restart fallback'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'non-interactive mode reports the exact deferred status'

    Reset-RestartScenario -Choice 'Y' -RestartShouldThrow $true
    Invoke-OOShutUpRestart
    Assert-Equal 1 $script:RestartInvocations.Count 'a graceful restart attempt is made after confirmation and the save window'
    Assert-True ($script:LogMessages -contains '[WARNING] Restart could not be started: Mock graceful restart was blocked.') 'restart failure has an exact user-facing status'
    Assert-True ($script:LogMessages -contains "[WARNING] $script:OOShutUpRestartDeferredMessage") 'restart failure leaves a clear recovery reminder'
    Assert-Equal 0 $script:InstallerOutcome.ExitCode 'restart failure does not change the completed installer outcome'

    function Reset-RelocationFaults {
        param(
            [string]$FailureStage = '',
            [int]$FailureAt = 0,
            [switch]$MoveFailureAfterOperation
        )
        $script:RelocationFailureStage = $FailureStage
        $script:RelocationFailureAt = $FailureAt
        $script:RelocationMoveFailureAfterOperation = [bool]$MoveFailureAfterOperation
        $script:RelocationCopyCount = 0
        $script:RelocationHashCount = 0
        $script:RelocationAclCount = 0
        $script:RelocationMoveCount = 0
    }

    function Get-RelocationTreeSnapshot {
        param([string]$Root)
        $canonicalRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
        $entries = New-Object System.Collections.ArrayList
        foreach ($path in @([System.IO.Directory]::GetFiles($canonicalRoot, '*', [System.IO.SearchOption]::AllDirectories) | Sort-Object)) {
            $relativePath = $path.Substring($canonicalRoot.Length).TrimStart([char[]]@('\', '/'))
            $digest = (Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            [void]$entries.Add("$relativePath|$digest")
        }
        return [string]::Join("`n", [string[]]$entries.ToArray())
    }

    function New-RelocationTestScenario {
        param([string]$Name)
        $scenarioRoot = Join-Path $script:RelocationTestRoot $Name
        $sourceRoot = Join-Path $scenarioRoot 'source'
        $profileRoot = Join-Path $sourceRoot 'profiles'
        [System.IO.Directory]::CreateDirectory($profileRoot) | Out-Null
        $sourceScriptPath = Join-Path $sourceRoot 'Install-OOShutUp10.ps1'
        $sourceReleaseManifestPath = Join-Path $sourceRoot $script:OOShutUpReleaseManifestFileName
        $sourceProfilePath = Join-Path $profileRoot 'ooshutup10.cfg'
        $sourceManifestPath = Join-Path $profileRoot 'ooshutup10.approval.json'
        [System.IO.File]::WriteAllText($sourceScriptPath, "script-$Name")
        [System.IO.File]::WriteAllText($sourceReleaseManifestPath, "release-manifest-$Name")
        [System.IO.File]::WriteAllText($sourceProfilePath, "profile-$Name")
        [System.IO.File]::WriteAllText($sourceManifestPath, "manifest-$Name")

        $approvedProfileDigest = (Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $sourceProfilePath -Algorithm SHA256).Hash
        $approvedManifestDigest = (Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $sourceManifestPath -Algorithm SHA256).Hash
        $target = Get-ScriptRelocationPackageTarget `
            -CurrentScriptPath $sourceScriptPath `
            -ProfileDirectory $profileRoot `
            -ScriptFileName 'Install-OOShutUp10.ps1' `
            -PackageContainer (Join-Path $scenarioRoot 'package-container') `
            -ApprovedProfileSha256 $approvedProfileDigest `
            -ApprovedProfileManifestSha256 $approvedManifestDigest

        $previousRoot = Join-Path ([System.IO.Path]::GetDirectoryName($target.Root)) 'previous-release'
        $previousProfiles = Join-Path $previousRoot 'profiles'
        [System.IO.Directory]::CreateDirectory($previousProfiles) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $previousRoot 'Install-OOShutUp10.ps1'), "previous-script-$Name")
        [System.IO.File]::WriteAllText((Join-Path $previousRoot $script:OOShutUpReleaseManifestFileName), "previous-release-manifest-$Name")
        [System.IO.File]::WriteAllText((Join-Path $previousProfiles 'ooshutup10.cfg'), "previous-profile-$Name")
        [System.IO.File]::WriteAllText((Join-Path $previousProfiles 'ooshutup10.approval.json'), "previous-manifest-$Name")
        [System.IO.File]::WriteAllText((Join-Path $previousRoot 'retained.marker'), "previous-release-$Name")

        return [pscustomobject]@{
            SourceScriptPath = $sourceScriptPath
            ProfileDirectory = $profileRoot
            Target = $target
            PreviousRoot = $previousRoot
            PreviousSnapshot = Get-RelocationTreeSnapshot -Root $previousRoot
        }
    }

    function Assert-RelocationFailureScenario {
        param(
            [string]$Name,
            [string]$FailureStage,
            [int]$FailureAt,
            [switch]$MoveFailureAfterOperation
        )
        Reset-RelocationFaults
        $scenario = New-RelocationTestScenario -Name $Name
        Reset-RelocationFaults -FailureStage $FailureStage -FailureAt $FailureAt -MoveFailureAfterOperation:$MoveFailureAfterOperation
        $failed = $false
        try {
            Install-ScriptRelocationPackage `
                -CurrentScriptPath $scenario.SourceScriptPath `
                -TargetRoot $scenario.Target.Root `
                -ScriptFileName 'Install-OOShutUp10.ps1' `
                -ProfileDirectory $scenario.ProfileDirectory `
                -ExpectedDigests $scenario.Target.Digests | Out-Null
        }
        catch {
            $failed = $true
        }
        Assert-True $failed "$Name rejects its injected $FailureStage failure"
        Assert-True (-not [System.IO.Directory]::Exists($scenario.Target.Root)) "$Name leaves no active content-addressed target after failure"
        Assert-Equal $scenario.PreviousSnapshot (Get-RelocationTreeSnapshot -Root $scenario.PreviousRoot) "$Name retains the previous complete package byte-for-byte"
        $stageDirectories = @([System.IO.Directory]::GetDirectories([System.IO.Path]::GetDirectoryName($scenario.Target.Root), '.stage-*'))
        Assert-Equal 0 $stageDirectories.Count "$Name removes its inactive staging directory after failure"
    }

    $script:RelocationTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oo-relocation-mocks-' + [guid]::NewGuid().ToString('N'))
    [System.IO.Directory]::CreateDirectory($script:RelocationTestRoot) | Out-Null
    Reset-RelocationFaults

    Set-Item -Path Function:\Get-FileHash -Value {
        [CmdletBinding()]
        param([string]$Path, [string]$LiteralPath, [string]$Algorithm = 'SHA256')
        $hashPath = if ($LiteralPath) { $LiteralPath } else { $Path }
        $script:RelocationHashCount++
        if ($script:RelocationFailureStage -eq 'Hash' -and $script:RelocationHashCount -eq $script:RelocationFailureAt) {
            throw "Injected hash failure at call $script:RelocationHashCount."
        }
        Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $hashPath -Algorithm $Algorithm -ErrorAction Stop
    }
    Set-Item -Path Function:\Copy-Item -Value {
        [CmdletBinding()]
        param([string]$Path, [string]$LiteralPath, [string]$Destination, [switch]$Force)
        $copySource = if ($LiteralPath) { $LiteralPath } else { $Path }
        $script:RelocationCopyCount++
        Microsoft.PowerShell.Management\Copy-Item -LiteralPath $copySource -Destination $Destination -Force:$Force -ErrorAction Stop
        if ($script:RelocationFailureStage -eq 'Copy' -and $script:RelocationCopyCount -eq $script:RelocationFailureAt) {
            throw "Injected copy failure after call $script:RelocationCopyCount."
        }
    }
    Set-Item -Path Function:\Test-SystemTaskTargetAcl -Value {
        [CmdletBinding()]
        param([string]$ScriptPath, [switch]$AllowDirectory)
        $script:RelocationAclCount++
        return -not ($script:RelocationFailureStage -eq 'Acl' -and $script:RelocationAclCount -eq $script:RelocationFailureAt)
    }
    Set-Item -Path Function:\Move-ScriptRelocationDirectory -Value {
        [CmdletBinding()]
        param([string]$Source, [string]$Destination)
        $script:RelocationMoveCount++
        $injectFailure = $script:RelocationFailureStage -eq 'Move' -and $script:RelocationMoveCount -eq $script:RelocationFailureAt
        if ($injectFailure -and -not $script:RelocationMoveFailureAfterOperation) {
            throw "Injected directory move failure before call $script:RelocationMoveCount."
        }
        [System.IO.Directory]::Move($Source, $Destination)
        if ($injectFailure) {
            throw "Injected directory move failure after call $script:RelocationMoveCount."
        }
    }

    Reset-RelocationFaults
    $successfulScenario = New-RelocationTestScenario -Name 'success'
    Assert-Equal (Join-Path $successfulScenario.Target.Root 'Install-OOShutUp10.ps1') $successfulScenario.Target.ScriptPath 'the task script path names the promoted package entry point'
    Reset-RelocationFaults
    $successfulPromotion = Install-ScriptRelocationPackage `
        -CurrentScriptPath $successfulScenario.SourceScriptPath `
        -TargetRoot $successfulScenario.Target.Root `
        -ScriptFileName 'Install-OOShutUp10.ps1' `
        -ProfileDirectory $successfulScenario.ProfileDirectory `
        -ExpectedDigests $successfulScenario.Target.Digests
    Assert-Equal $true $successfulPromotion 'the verified four-file package is promoted successfully'
    Assert-True (Test-ScriptRelocationPackage -PackageRoot $successfulScenario.Target.Root -ScriptFileName 'Install-OOShutUp10.ps1' -ExpectedDigests $successfulScenario.Target.Digests) 'the active release contains the matching script, release manifest, cfg, profile manifest, and ACL checks'
    Assert-Equal $successfulScenario.PreviousSnapshot (Get-RelocationTreeSnapshot -Root $successfulScenario.PreviousRoot) 'a successful new release leaves the prior complete package byte-for-byte intact'
    Assert-Equal 0 @([System.IO.Directory]::GetDirectories([System.IO.Path]::GetDirectoryName($successfulScenario.Target.Root), '.stage-*')).Count 'successful promotion leaves no active staging directory'

    Reset-RelocationFaults
    $reusedPromotion = Install-ScriptRelocationPackage `
        -CurrentScriptPath $successfulScenario.SourceScriptPath `
        -TargetRoot $successfulScenario.Target.Root `
        -ScriptFileName 'Install-OOShutUp10.ps1' `
        -ProfileDirectory $successfulScenario.ProfileDirectory `
        -ExpectedDigests $successfulScenario.Target.Digests
    Assert-Equal $true $reusedPromotion 'a previously promoted identical package is safely reused'
    Assert-Equal 0 $script:RelocationCopyCount 'reusing a complete release does not copy files'
    Assert-Equal 0 $script:RelocationMoveCount 'reusing a complete release does not move directories'

    Reset-RelocationFaults
    $occupiedScenario = New-RelocationTestScenario -Name 'occupied-target'
    $occupiedProfiles = Join-Path $occupiedScenario.Target.Root 'profiles'
    [System.IO.Directory]::CreateDirectory($occupiedProfiles) | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $occupiedScenario.Target.Root 'Install-OOShutUp10.ps1'), 'partial prior target')
    [System.IO.File]::WriteAllText((Join-Path $occupiedProfiles 'ooshutup10.cfg'), 'partial prior profile')
    $occupiedSnapshot = Get-RelocationTreeSnapshot -Root $occupiedScenario.Target.Root
    Reset-RelocationFaults
    $occupiedTargetRejected = $false
    try {
        Install-ScriptRelocationPackage `
            -CurrentScriptPath $occupiedScenario.SourceScriptPath `
            -TargetRoot $occupiedScenario.Target.Root `
            -ScriptFileName 'Install-OOShutUp10.ps1' `
            -ProfileDirectory $occupiedScenario.ProfileDirectory `
            -ExpectedDigests $occupiedScenario.Target.Digests | Out-Null
    }
    catch {
        $occupiedTargetRejected = $true
    }
    Assert-True $occupiedTargetRejected 'an existing incomplete target fails closed instead of being overwritten'
    Assert-Equal $occupiedSnapshot (Get-RelocationTreeSnapshot -Root $occupiedScenario.Target.Root) 'an invalid pre-existing target remains byte-identical after refusal'
    Assert-Equal 0 $script:RelocationCopyCount 'an invalid pre-existing target is rejected before staging copies'
    Assert-Equal 0 $script:RelocationMoveCount 'an invalid pre-existing target is rejected before promotion'

    foreach ($copyNumber in 1..4) {
        Assert-RelocationFailureScenario -Name "copy-$copyNumber" -FailureStage 'Copy' -FailureAt $copyNumber
    }
    foreach ($hashNumber in 1..16) {
        Assert-RelocationFailureScenario -Name "hash-$hashNumber" -FailureStage 'Hash' -FailureAt $hashNumber
    }
    foreach ($aclNumber in 1..10) {
        Assert-RelocationFailureScenario -Name "acl-$aclNumber" -FailureStage 'Acl' -FailureAt $aclNumber
    }
    Assert-RelocationFailureScenario -Name 'move-before' -FailureStage 'Move' -FailureAt 1

    Reset-RelocationFaults
    $verifiedAfterMoveScenario = New-RelocationTestScenario -Name 'move-after'
    Reset-RelocationFaults -FailureStage 'Move' -FailureAt 1 -MoveFailureAfterOperation
    $verifiedAfterMove = Install-ScriptRelocationPackage `
        -CurrentScriptPath $verifiedAfterMoveScenario.SourceScriptPath `
        -TargetRoot $verifiedAfterMoveScenario.Target.Root `
        -ScriptFileName 'Install-OOShutUp10.ps1' `
        -ProfileDirectory $verifiedAfterMoveScenario.ProfileDirectory `
        -ExpectedDigests $verifiedAfterMoveScenario.Target.Digests
    Assert-Equal $true $verifiedAfterMove 'a move that reports late is accepted only after the promoted package and ACLs verify'
    Assert-True (Test-ScriptRelocationPackage -PackageRoot $verifiedAfterMoveScenario.Target.Root -ScriptFileName 'Install-OOShutUp10.ps1' -ExpectedDigests $verifiedAfterMoveScenario.Target.Digests) 'a late move result never activates an unverified package'
    Assert-Equal $verifiedAfterMoveScenario.PreviousSnapshot (Get-RelocationTreeSnapshot -Root $verifiedAfterMoveScenario.PreviousRoot) 'a late move result retains the previous complete release'

    Reset-RelocationFaults
    $fallbackScenario = New-RelocationTestScenario -Name 'fallback'
    $DefaultScriptRoot = $fallbackScenario.Target.Root
    $DefaultScriptPath = $fallbackScenario.Target.ScriptPath
    $ScriptFileName = 'Install-OOShutUp10.ps1'
    $script:OOShutUpProfileDirectory = $fallbackScenario.ProfileDirectory
    $script:OOShutUpRelocationPackageDigests = $fallbackScenario.Target.Digests
    Reset-RelocationFaults -FailureStage 'Move' -FailureAt 1
    $fallbackResult = Invoke-ScriptRelocation -CurrentScriptPath $fallbackScenario.SourceScriptPath
    Assert-Equal $false $fallbackResult 'an uncertain promotion returns to the caller running the original script'
    Assert-True (-not [System.IO.Directory]::Exists($fallbackScenario.Target.Root)) 'fallback leaves no mixed or uncertain package at the task target path'
    Assert-Equal $fallbackScenario.PreviousSnapshot (Get-RelocationTreeSnapshot -Root $fallbackScenario.PreviousRoot) 'fallback retains the previous complete package byte-for-byte'

    Remove-Item -LiteralPath $script:RelocationTestRoot -Recurse -Force -ErrorAction Stop

    [Console]::Out.WriteLine("PASS: Restart safety mock tests ($script:AssertionCount assertions) in $($PSVersionTable.PSVersion).")
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
