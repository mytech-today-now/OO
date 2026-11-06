#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs O&O ShutUp10++ and applies the approved settings profile packaged beside the script.

.DESCRIPTION
    This script will:
    1. Validate the reviewed profile and obtain interactive approval or a matching unattended digest
    2. Check if O&O ShutUp10++ is already installed and download it if not found
    3. Attempt a system restore point before applying the profile
    4. Apply the approved privacy settings profile quietly
    5. Create a scheduled task pinned to the approved profile and manifest digests
    6. Offer a graceful, cancellable restart only when O&O reports changed settings; force requires -ForceRestart

.PARAMETER ReapplyOnly
    If specified, only reapplies privacy settings without installation or task creation.
    This is used by the scheduled task after Windows updates.

.PARAMETER NonInteractive
    Suppresses prompts only when both approved profile and manifest SHA-256 digests match. Redirected or unavailable
    console input automatically selects this exact-digest approval path before any configuration changes.

.PARAMETER ForceRestart
    Explicitly permits Restart-Computer to force-close applications when a restart is confirmed.
    Unsaved work may be lost; use only when that risk is accepted.

.PARAMETER ApprovedProfileSha256
    The exact SHA-256 digest of the reviewed O&O settings export approved for unattended use.

.PARAMETER ApprovedProfileManifestSha256
    The exact SHA-256 digest of the reviewed profile manifest approved for unattended use.

.PARAMETER ApprovedOOSExecutablePath
    The exact absolute managed O&O executable path registered for SYSTEM reapplication.

.PARAMETER ProfileApprovalConfirmed
    Internal relocation handoff used after the interactive profile approval has already been accepted.

.PARAMETER AllowWithoutRestorePoint
    Explicitly permits applying the approved profile when a new System Restore point cannot be verified.
    System Restore recovery is unavailable when this override is used.

.PARAMETER SkipRelocate
    If specified, skips the script relocation to the default location.
    Used internally to prevent infinite loops.

.NOTES
    File Name      : Install-OOShutUp10.ps1
    Author         : Kyle C. Rode / myTech.Today
    Company        : myTech.Today
    Version        : 2.3.0
    DateCreated    : 2025-01-08
    LastModified   : 2026-10-31
    Copyright      : (c) 2025 myTech.Today. All rights reserved.
    Prerequisite   : PowerShell 5.1+, Administrator privileges

    Execution policy:
    Prefer an organization-approved signed copy. The -ExecutionPolicy Bypass examples below are process-scoped
    and must only be used when approved by your policy owner. MachinePolicy and UserPolicy Group Policy settings
    take precedence; otherwise Process takes precedence over CurrentUser and LocalMachine. A blocked launch must
    stop without changing stored policy. PowerShell policy blocks this script. No policy was changed; use an approved signed or process-scoped launch.

    About myTech.Today:
    myTech.Today is a Chicagoland-based MSP and IT consulting firm specializing in
    PowerShell automation, infrastructure optimization, and custom development solutions.
    With 20+ years of experience, we deliver tech solutions that improve efficiency
    and drive business value.

    Contact Information:
    - Email: sales@mytech.today
    - Website: https://mytech.today
    - Phone: (847) 767-4914
    - Location: Lake Zurich, IL (Chicagoland area)
    - GitHub: https://github.com/mytech-today-now

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Install-OOShutUp10.ps1"
    Starts Windows PowerShell 5.1 with a process-scoped policy. Use only when approved by your policy owner.

.EXAMPLE
    pwsh.exe -NoProfile -ExecutionPolicy Bypass -File ".\Install-OOShutUp10.ps1"
    Starts PowerShell 7 with a process-scoped policy. Use only when approved by your policy owner.

.EXAMPLE
    .\Install-OOShutUp10.ps1
    Installs O&O ShutUp10++ and applies the approved settings profile

.EXAMPLE
    .\Install-OOShutUp10.ps1 -ReapplyOnly
    Only reapplies privacy settings (used by scheduled task)

.LINK
    https://mytech.today

.LINK
    https://github.com/mytech-today-now/OO
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$ReapplyOnly,

    [Parameter(Mandatory = $false)]
    [switch]$SkipRelocate,

    [Parameter(Mandatory = $false)]
    [switch]$NonInteractive,

    [Parameter(Mandatory = $false)]
    [switch]$ForceRestart,

    [Parameter(Mandatory = $false)]
    [string]$ApprovedProfileSha256,

    [Parameter(Mandatory = $false)]
    [string]$ApprovedProfileManifestSha256,

    [Parameter(Mandatory = $false)]
    [string]$ApprovedOOSExecutablePath,

    [Parameter(Mandatory = $false)]
    [switch]$ProfileApprovalConfirmed,

    [Parameter(Mandatory = $false)]
    [switch]$AllowWithoutRestorePoint
)

# PowerShell 7+ Version Check - myTech.Today standard
$script:PS7ContinueOnPS51 = $true  # Allow running on PS 5.1 with warning
$script:PS7Silent = $false
$script:_RepoRoot = $PSScriptRoot
while ($script:_RepoRoot -and -not (Test-Path (Join-Path $script:_RepoRoot 'scripts\Require-PowerShell7.ps1'))) {
    $script:_RepoRoot = Split-Path $script:_RepoRoot -Parent
}
if ($script:_RepoRoot -and (Test-Path (Join-Path $script:_RepoRoot 'scripts\Require-PowerShell7.ps1'))) {
    . (Join-Path $script:_RepoRoot 'scripts\Require-PowerShell7.ps1')
}

# Script configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Script metadata
$script:ScriptVersion = '2.3.0'

# Default installation paths
$DefaultScriptRoot = "$env:USERPROFILE\myTech.Today\scripts\Install-OOShutUp10"
$script:DefaultScriptPackageContainer = Join-Path $DefaultScriptRoot 'packages'
$ScriptFileName = "Install-OOShutUp10.ps1"
$DefaultScriptPath = Join-Path $DefaultScriptRoot $ScriptFileName

# O&O ShutUp10++ configuration
$AppName = "O&O ShutUp10++"
$DownloadUrl = "https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe"
$ExecutableName = "OOSU10.exe"
$MachineProgramFilesPath = $env:ProgramW6432
if ([string]::IsNullOrWhiteSpace($MachineProgramFilesPath)) {
    $MachineProgramFilesPath = $env:ProgramFiles
}
if ([string]::IsNullOrWhiteSpace($MachineProgramFilesPath) -or -not [System.IO.Path]::IsPathRooted($MachineProgramFilesPath)) {
    throw 'The machine Program Files directory is unavailable; no O&O executable path can be established.'
}
$InstallPath = [System.IO.Path]::GetFullPath((Join-Path $MachineProgramFilesPath 'OO Software\O&O ShutUp10'))
$script:OOShutUpProfileDirectory = Join-Path $PSScriptRoot 'profiles'
$script:OOShutUpProfilePath = Join-Path $script:OOShutUpProfileDirectory 'ooshutup10.cfg'
$script:OOShutUpProfileManifestPath = Join-Path $script:OOShutUpProfileDirectory 'ooshutup10.approval.json'
$script:OOShutUpProfileFailureMessage = 'The approved settings profile is missing or invalid. No privacy settings were applied; place the reviewed profile in the configured directory.'
$script:OOShutUpProfileSnapshotFailureMessage = 'The approved profile changed after validation. No O&O process was started; restore the reviewed profile and retry.'
$script:OOShutUpUnknownScopeFailureMessage = 'The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.'
$script:OOShutUpCurrentUserScopeFailureMessage = 'The SYSTEM task profile contains current-user settings. Reapplication was skipped; run them for the intended signed-in user or remove them from the SYSTEM task profile.'
$script:OOShutUpApprovalFailureMessage = 'The selected settings profile has no matching approval. No settings were changed; review both file digests in the interactive installer and recreate the task.'
$script:OOShutUpRestorePointFailureMessage = 'A restore point could not be verified. No privacy settings were applied; free space or enable restore protection, then retry.'
$script:OOShutUpRestorePointPreparationFailureMessage = 'Restore-point preparation failed on the detected Windows volume. No settings were applied; restore the prior registry value and fix the provider error.'
$script:OOShutUpRestorePointOverrideMessage = 'A restore point could not be verified. System Restore recovery is unavailable; continuing only because -AllowWithoutRestorePoint was supplied. The approved O&O profile export remains separate.'
$script:OOShutUpRestorePointFrequencyRegistryPath = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
$script:OOShutUpRestorePointFrequencyRegistryName = 'SystemRestorePointCreationFrequency'
# Manual and SYSTEM entry points share this Global mutex. Wait at most ten minutes before failing closed.
$script:OOShutUpMutationMutexName = 'Global\myTech.Today.OOShutUp.Mutation'
$script:OOShutUpMutationLockTimeoutMilliseconds = 600000
$script:OOShutUpTaskAclFailureMessage = 'The SYSTEM task target ACL could not be verified. No task was registered; move the script to a protected path and retry.'
$script:OOShutUpTaskSecurityDescriptorSddl = 'O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)'
$script:OOShutUpTaskSecurityDescriptorFailureMessage = 'The task security descriptor could not be verified. No new task was enabled; inspect task permissions before retrying.'
$script:OOShutUpTrustedPowerShellFailureMessage = 'A trusted Windows PowerShell executable could not be resolved. No scheduled task was registered or changed.'
$script:OOShutUpTaskOwnershipFailureMessage = 'Task ownership is ambiguous. No task was deleted; inspect the listed path and retry after confirming ownership.'
$script:OOShutUpPartialSetupMessage = 'The privacy settings may be applied, but the update task was not verified. Setup is partial; repair the task before relying on automatic reapplication.'
$script:OOShutUpTaskRecoveryMessage = 'Recovery: resolve the logged Task Scheduler error, open PowerShell as Administrator, and rerun this installer with the approved profile (or use -NonInteractive -ApprovedProfileSha256 <profile SHA-256> -ApprovedProfileManifestSha256 <manifest SHA-256>). Verify \myTech.Today\Install-OOShutUp10 in Task Scheduler before relying on automatic reapplication. The retry inspects and replaces only this verified task; same-named tasks in other folders are left unchanged.'
$script:OOShutUpManagedPathFailureMessage = 'The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.'
$script:OOShutUpUpdateEventSubscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''Microsoft-Windows-WindowsUpdateClient''] and (EventID=19)]]</Select></Query></QueryList>'
$script:OOShutUpUpdateEventDelay = 'PT2M'
$script:OOShutUpUpdateMultipleInstancesPolicy = 'Queue'
$script:OOShutUpApprovedDigest = $null
$script:OOShutUpApprovedManifestDigest = $null
$script:OOShutUpLastRunChanged = $false
$script:OOShutUpRestartDeferredMessage = 'Restart was deferred. Setup status remains unchanged; save your work and restart when ready.'
$script:OOShutUpRestartPromptUnavailableMessage = 'No interactive restart prompt is available. Configuration status is retained; restart manually when convenient.'
$script:OOShutUpRestartPromptTimeoutSeconds = 60
$script:OOShutUpRestartSaveWindowSeconds = 60
# Accept only Authenticode-valid O&O Software GmbH signatures for O&O ShutUp10 at the current supported release floor.
# Vendor release policy reviewed 2026-10-31: https://www.oo-software.com/en/shutup10/update
$ExpectedOOShutUpPublisher = 'O&O Software GmbH'
$MinimumOOShutUpVersion = [version]'3.6.1135'
$OOShutUpPayloadFailureMessage = 'The O&O download could not be authenticated. No executable was started; verify the vendor signature before retrying.'

#region Local Logging

# Keep logging packaged with this privileged script; never evaluate remote text.
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] [$Level] $Message"
}

#endregion

#region Restart Handling

function Test-OOShutUpInteractiveInput {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        return (-not [System.Console]::IsInputRedirected)
    }
    catch {
        return $false
    }
}

function Read-OOShutUpRestartKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 60
    )

    if ($TimeoutSeconds -le 0) {
        return $null
    }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ([System.Console]::KeyAvailable) {
            $key = [System.Console]::ReadKey($true)
            return [string]$key.Key
        }

        Start-Sleep -Milliseconds 100
    }

    return $null
}

function Wait-OOShutUpRestartSaveWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$Seconds = 60
    )

    if ($Seconds -le 0) {
        return $true
    }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $Seconds) {
        try {
            if ([System.Console]::KeyAvailable) {
                [System.Console]::ReadKey($true) | Out-Null
                return $false
            }

            Start-Sleep -Milliseconds 100
        }
        catch {
            return $false
        }
    }

    return $true
}

function Invoke-OOShutUpRestart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$NonInteractive,

        [Parameter(Mandatory = $false)]
        [switch]$ForceRestart
    )

    if ($NonInteractive) {
        Write-Log $script:OOShutUpRestartPromptUnavailableMessage -Level WARNING
        Write-Log $script:OOShutUpRestartDeferredMessage -Level WARNING
        return
    }

    if ($ForceRestart) {
        Write-Log 'The explicit -ForceRestart option can close applications that have unsaved work.' -Level WARNING
    }

    Write-Host 'Configuration status is shown above. Save your work before restarting.'
    Write-Host ('Press Y to restart, N or Esc to defer. No key within {0} seconds defers the restart.' -f $script:OOShutUpRestartPromptTimeoutSeconds)

    $restartPromptUnavailable = $false
    try {
        $restartKey = Read-OOShutUpRestartKey -TimeoutSeconds $script:OOShutUpRestartPromptTimeoutSeconds
    }
    catch {
        $restartKey = $null
        $restartPromptUnavailable = $true
    }

    if ($restartPromptUnavailable) {
        Write-Log $script:OOShutUpRestartPromptUnavailableMessage -Level WARNING
    }

    if ($null -eq $restartKey) {
        Write-Log $script:OOShutUpRestartDeferredMessage -Level WARNING
        return
    }

    if ($restartKey -ne 'Y') {
        Write-Log $script:OOShutUpRestartDeferredMessage -Level WARNING
        return
    }

    Write-Host ('Save your work now. The restart begins in {0} seconds; press any key to cancel.' -f $script:OOShutUpRestartSaveWindowSeconds)
    try {
        $saveWindowElapsed = Wait-OOShutUpRestartSaveWindow -Seconds $script:OOShutUpRestartSaveWindowSeconds
    }
    catch {
        $saveWindowElapsed = $false
    }

    if (-not $saveWindowElapsed) {
        Write-Log $script:OOShutUpRestartDeferredMessage -Level WARNING
        return
    }

    try {
        if ($ForceRestart) {
            Restart-Computer -Force
        }
        else {
            Restart-Computer
        }
    }
    catch {
        Write-Log ("Restart could not be started: {0}" -f $_.Exception.Message) -Level WARNING
        Write-Log $script:OOShutUpRestartDeferredMessage -Level WARNING
    }
}

#endregion

#region Helper Functions

function Show-MyTechTodayBanner {
    <#
    .SYNOPSIS
        Displays the myTech.Today marketing banner
    #>
    param()

    # Each line must be exactly 78 characters between the pipes
    $lines = @(
        "",
        "+==============================================================================+",
        "|                                                                              |",
        "|                            myTech.Today                                      |",
        "|                   Professional IT Solutions & Automation                     |",
        "|                                                                              |",
        "+==============================================================================+",
        "|                                                                              |",
        "|  This script is brought to you by myTech.Today - Your trusted partner       |",
        "|  for IT consulting, managed services, and custom automation solutions.      |",
        "|                                                                              |",
        "|  Services:                                                                   |",
        "|    - PowerShell Automation & Scripting                                      |",
        "|    - Infrastructure Optimization                                            |",
        "|    - Custom Application Development                                         |",
        "|    - Cloud Integration (AWS, Azure, Google Cloud)                           |",
        "|    - Managed IT Services (MSP)                                              |",
        "|    - IT Consulting & Support                                                |",
        "|                                                                              |",
        "|  Service Area: Chicagoland (Lake Zurich, IL) & Remote Support Available     |",
        "|                                                                              |",
        "|  Contact Us:                                                                 |",
        "|    Email:    sales@mytech.today                                             |",
        "|    Website:  https://mytech.today                                           |",
        "|    Phone:    (847) 767-4914                                                 |",
        "|    GitHub:   https://github.com/mytech-today-now                            |",
        "|                                                                              |",
        "|  20+ Years of Experience | 190+ Satisfied Clients | 5-Star Rating           |",
        "|                                                                              |",
        "+==============================================================================+",
        ""
    )

    foreach ($line in $lines) {
        Write-Host $line -ForegroundColor Cyan
    }
}

function Update-ScriptProgress {
    <#
    .SYNOPSIS
        Updates the overall script progress bar
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$PercentComplete,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [string]$CurrentOperation = ""
    )

    Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" `
                   -Status $Status `
                   -PercentComplete $PercentComplete `
                   -CurrentOperation $CurrentOperation
}

#region Logging Wrapper

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Initializes the packaged local logger for this script.
    #>
    param()

    try {
        Write-Log -Message 'Packaged local console logger initialized.' -Level INFO
        return $true
    }
    catch {
        Write-Host "Trusted logging validation failed. No configuration changes were started; use the packaged local logger and review the logged stage." -ForegroundColor Red
        return $false
    }
}

#endregion

function Get-ScriptRelocationPackageTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$ProfileDirectory,

        [Parameter(Mandatory = $true)]
        [string]$ScriptFileName,

        [Parameter(Mandatory = $true)]
        [string]$PackageContainer,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileSha256,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileManifestSha256
    )

    $profilePath = Join-Path $ProfileDirectory 'ooshutup10.cfg'
    $manifestPath = Join-Path $ProfileDirectory 'ooshutup10.approval.json'
    foreach ($path in @($CurrentScriptPath, $profilePath, $manifestPath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "The relocation package source file is missing: $path"
        }
    }

    $digests = [ordered]@{
        Script   = (Get-FileHash -LiteralPath $CurrentScriptPath -Algorithm SHA256 -ErrorAction Stop).Hash
        Profile  = (Get-FileHash -LiteralPath $profilePath -Algorithm SHA256 -ErrorAction Stop).Hash
        Manifest = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    if ($ApprovedProfileSha256 -notmatch '^[A-Fa-f0-9]{64}$' -or
        $ApprovedProfileManifestSha256 -notmatch '^[A-Fa-f0-9]{64}$' -or
        -not [string]::Equals($digests.Profile, $ApprovedProfileSha256, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($digests.Manifest, $ApprovedProfileManifestSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'The relocation profile source no longer matches both approved SHA-256 digests.'
    }

    $packageMaterial = '{0}:{1}:{2}' -f $digests.Script, $digests.Profile, $digests.Manifest
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $packageId = ([System.BitConverter]::ToString(
            $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($packageMaterial))
        )).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }

    $targetRoot = Join-Path (Join-Path $PackageContainer 'releases') $packageId
    return [pscustomobject]@{
        PackageId = $packageId
        Digests = $digests
        Root = $targetRoot
        ScriptPath = Join-Path $targetRoot $ScriptFileName
    }
}

function Test-ScriptRelocationPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,

        [Parameter(Mandatory = $true)]
        [string]$ScriptFileName,

        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$ExpectedDigests
    )

    try {
        if (-not (Test-SystemTaskTargetAcl -ScriptPath $PackageRoot -AllowDirectory)) {
            return $false
        }

        $filePaths = [ordered]@{
            Script   = Join-Path $PackageRoot $ScriptFileName
            Profile  = Join-Path (Join-Path $PackageRoot 'profiles') 'ooshutup10.cfg'
            Manifest = Join-Path (Join-Path $PackageRoot 'profiles') 'ooshutup10.approval.json'
        }
        foreach ($key in @('Script', 'Profile', 'Manifest')) {
            $path = [string]$filePaths[$key]
            if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or
                -not (Test-SystemTaskTargetAcl -ScriptPath $path)) {
                return $false
            }
            $actualDigest = (Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not [string]::Equals(
                $actualDigest,
                [string]$ExpectedDigests[$key],
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                return $false
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

function Move-ScriptRelocationDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    # Source and destination are unique sibling directories on the same volume.
    [System.IO.Directory]::Move($Source, $Destination)
}

function Remove-ScriptRelocationStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StageRoot,

        [Parameter(Mandatory = $true)]
        [string]$ScriptFileName
    )

    if (-not [System.IO.Directory]::Exists($StageRoot)) {
        return
    }

    $stageItem = Get-Item -LiteralPath $StageRoot -Force -ErrorAction Stop
    if (([int]$stageItem.Attributes -band [int][System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'The relocation staging path became a reparse point; it was left untouched.'
    }

    $profileDirectory = Join-Path $StageRoot 'profiles'
    if ([System.IO.Directory]::Exists($profileDirectory)) {
        $profileItem = Get-Item -LiteralPath $profileDirectory -Force -ErrorAction Stop
        if (([int]$profileItem.Attributes -band [int][System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'The staged profile directory became a reparse point; it was left untouched.'
        }
        foreach ($fileName in @('ooshutup10.cfg', 'ooshutup10.approval.json')) {
            $filePath = Join-Path $profileDirectory $fileName
            if ([System.IO.File]::Exists($filePath)) {
                [System.IO.File]::Delete($filePath)
            }
        }
        [System.IO.Directory]::Delete($profileDirectory, $false)
    }

    $scriptPath = Join-Path $StageRoot $ScriptFileName
    if ([System.IO.File]::Exists($scriptPath)) {
        [System.IO.File]::Delete($scriptPath)
    }
    [System.IO.Directory]::Delete($StageRoot, $false)
}

function Install-ScriptRelocationPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetRoot,

        [Parameter(Mandatory = $true)]
        [string]$ScriptFileName,

        [Parameter(Mandatory = $true)]
        [string]$ProfileDirectory,

        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$ExpectedDigests
    )

    $targetRoot = [System.IO.Path]::GetFullPath($TargetRoot)
    $targetParent = [System.IO.Path]::GetDirectoryName($targetRoot)
    if ([string]::IsNullOrWhiteSpace($targetParent)) {
        throw 'The relocation package target has no parent directory.'
    }

    if ([System.IO.Directory]::Exists($targetRoot)) {
        if (Test-ScriptRelocationPackage -PackageRoot $targetRoot -ScriptFileName $ScriptFileName -ExpectedDigests $ExpectedDigests) {
            return $true
        }
        throw 'The content-addressed relocation target exists but its files or ACLs do not verify; it was left unchanged.'
    }
    if ([System.IO.File]::Exists($targetRoot)) {
        throw 'The content-addressed relocation target is occupied by a file; it was left unchanged.'
    }

    [System.IO.Directory]::CreateDirectory($targetParent) | Out-Null
    $stageRoot = Join-Path $targetParent ('.stage-{0}' -f [guid]::NewGuid().ToString('N'))
    $stageProfileDirectory = Join-Path $stageRoot 'profiles'
    try {
        [System.IO.Directory]::CreateDirectory($stageRoot) | Out-Null
        [System.IO.Directory]::CreateDirectory($stageProfileDirectory) | Out-Null

        $sourcePaths = [ordered]@{
            Script   = $CurrentScriptPath
            Profile  = Join-Path $ProfileDirectory 'ooshutup10.cfg'
            Manifest = Join-Path $ProfileDirectory 'ooshutup10.approval.json'
        }
        foreach ($key in @('Script', 'Profile', 'Manifest')) {
            $sourcePath = [string]$sourcePaths[$key]
            if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                throw "The relocation source file is missing: $sourcePath"
            }
            $sourceDigest = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not [string]::Equals(
                $sourceDigest,
                [string]$ExpectedDigests[$key],
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                throw "The relocation source file changed after approval: $sourcePath"
            }
        }

        Copy-Item -LiteralPath $CurrentScriptPath -Destination (Join-Path $stageRoot $ScriptFileName) -Force -ErrorAction Stop
        Copy-Item -LiteralPath (Join-Path $ProfileDirectory 'ooshutup10.cfg') -Destination (Join-Path $stageProfileDirectory 'ooshutup10.cfg') -Force -ErrorAction Stop
        Copy-Item -LiteralPath (Join-Path $ProfileDirectory 'ooshutup10.approval.json') -Destination (Join-Path $stageProfileDirectory 'ooshutup10.approval.json') -Force -ErrorAction Stop

        foreach ($key in @('Script', 'Profile', 'Manifest')) {
            $sourcePath = [string]$sourcePaths[$key]
            $sourceDigest = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256 -ErrorAction Stop).Hash
            if (-not [string]::Equals(
                $sourceDigest,
                [string]$ExpectedDigests[$key],
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                throw "The relocation source file changed while staging: $sourcePath"
            }
        }

        if (-not (Test-ScriptRelocationPackage -PackageRoot $stageRoot -ScriptFileName $ScriptFileName -ExpectedDigests $ExpectedDigests)) {
            throw 'The staged relocation package digests or target ACLs could not be verified.'
        }

        # A concurrent run may have promoted this same content-addressed package.
        if ([System.IO.Directory]::Exists($targetRoot)) {
            if (Test-ScriptRelocationPackage -PackageRoot $targetRoot -ScriptFileName $ScriptFileName -ExpectedDigests $ExpectedDigests) {
                Remove-ScriptRelocationStage -StageRoot $stageRoot -ScriptFileName $ScriptFileName
                return $true
            }
            throw 'Another relocation created an invalid content-addressed target; it was left unchanged.'
        }

        try {
            Move-ScriptRelocationDirectory -Source $stageRoot -Destination $targetRoot
        }
        catch {
            # A move API can report an error after completing the same-volume rename. Only accept it after full verification.
            $moveFailure = $_.Exception.Message
            if (-not [System.IO.Directory]::Exists($stageRoot)) {
                if (Test-ScriptRelocationPackage -PackageRoot $targetRoot -ScriptFileName $ScriptFileName -ExpectedDigests $ExpectedDigests) {
                    return $true
                }
                if ([System.IO.Directory]::Exists($targetRoot)) {
                    $rejectedRoot = Join-Path $targetParent ('.rejected-{0}' -f [guid]::NewGuid().ToString('N'))
                    try {
                        Move-ScriptRelocationDirectory -Source $targetRoot -Destination $rejectedRoot
                    }
                    catch {
                        throw ("The directory move failed and the unverified package could not be moved aside from '{0}': {1}. Original move error: {2}" -f $targetRoot, $_.Exception.Message, $moveFailure)
                    }
                }
            }
            throw "The package promotion move failed: $moveFailure"
        }

        if (-not (Test-ScriptRelocationPackage -PackageRoot $targetRoot -ScriptFileName $ScriptFileName -ExpectedDigests $ExpectedDigests)) {
            $rejectedRoot = Join-Path $targetParent ('.rejected-{0}' -f [guid]::NewGuid().ToString('N'))
            try {
                Move-ScriptRelocationDirectory -Source $targetRoot -Destination $rejectedRoot
            }
            catch {
                throw ("The promoted package failed verification and could not be moved aside from '{0}': {1}" -f $targetRoot, $_.Exception.Message)
            }
            throw "The promoted package failed its post-promotion digest or ACL check and was retained at '$rejectedRoot'."
        }

        return $true
    }
    catch {
        $failure = $_.Exception.Message
        if ([System.IO.Directory]::Exists($stageRoot)) {
            try {
                Remove-ScriptRelocationStage -StageRoot $stageRoot -ScriptFileName $ScriptFileName
            }
            catch {
                $failure = '{0} Staging cleanup failed and the inactive directory remains at {1}: {2}' -f $failure, $stageRoot, $_.Exception.Message
            }
        }
        throw "Script relocation package was not activated: $failure"
    }
}

function Invoke-ScriptRelocation {
    <#
    .SYNOPSIS
        Ensures the script is running from the default location
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentScriptPath
    )

    try {
        $canonicalCurrentPath = [System.IO.Path]::GetFullPath($CurrentScriptPath)
        $canonicalDefaultPath = [System.IO.Path]::GetFullPath($DefaultScriptPath)
        # Check if already running from this content-addressed package location.
        if ([string]::Equals($canonicalCurrentPath, $canonicalDefaultPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Script is already running from default location" -ForegroundColor Cyan
            return $true
        }

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Staging the approved script and profile package..." -ForegroundColor Cyan
        Install-ScriptRelocationPackage `
            -CurrentScriptPath $CurrentScriptPath `
            -TargetRoot $DefaultScriptRoot `
            -ScriptFileName $ScriptFileName `
            -ProfileDirectory $script:OOShutUpProfileDirectory `
            -ExpectedDigests $script:OOShutUpRelocationPackageDigests | Out-Null

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SUCCESS] Verified package promoted to: $DefaultScriptRoot" -ForegroundColor Green

        # Re-execute from default location with same parameters
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Re-executing from default location..." -ForegroundColor Cyan
        Write-Host ""

        # Execute the relocated script directly using the call operator
        $relocateParams = @('-SkipRelocate')
        if ($ReapplyOnly) { $relocateParams += '-ReapplyOnly' }
        if ($NonInteractive) { $relocateParams += '-NonInteractive' }
        if ($ForceRestart) { $relocateParams += '-ForceRestart' }
        if ($AllowWithoutRestorePoint) { $relocateParams += '-AllowWithoutRestorePoint' }
        if ($script:OOShutUpApprovalConfirmed -and $script:OOShutUpApprovedDigest -and $script:OOShutUpApprovedManifestDigest) {
            $relocateParams += @(
                '-ApprovedProfileSha256', $script:OOShutUpApprovedDigest,
                '-ApprovedProfileManifestSha256', $script:OOShutUpApprovedManifestDigest,
                '-ProfileApprovalConfirmed'
            )
        }

        try {
            & $DefaultScriptPath @relocateParams
        }
        catch {
            # Older copies at the destination may not understand -SkipRelocate; fall back to in-place run.
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARN] Relocated script failed to launch (likely older version at destination): $($_.Exception.Message). Continuing in-place." -ForegroundColor Yellow
            return $false
        }

        # Exit this instance with the relocated child's exit code
        exit $LASTEXITCODE
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARN] Could not relocate script ($($_.Exception.Message)). Continuing in-place." -ForegroundColor Yellow
        return $false
    }
}

function New-OOShutUpMutationMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $security = [System.Security.AccessControl.MutexSecurity]::new()
    $allowedSids = @(
        [System.Security.Principal.SecurityIdentifier]::new('S-1-5-18'),
        [System.Security.Principal.SecurityIdentifier]::new('S-1-5-32-544'),
        [System.Security.Principal.WindowsIdentity]::GetCurrent().User
    )
    foreach ($sid in $allowedSids) {
        if ($null -eq $sid) {
            throw 'The current Windows identity could not be resolved for the mutation lock.'
        }
        $rule = [System.Security.AccessControl.MutexAccessRule]::new(
            $sid,
            [System.Security.AccessControl.MutexRights]::FullControl,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        [void]$security.AddAccessRule($rule)
    }

    $createdNew = $false
    $constructorTypes = [type[]]@(
        [bool],
        [string],
        [bool].MakeByRefType(),
        [System.Security.AccessControl.MutexSecurity]
    )
    $securityAwareConstructor = [System.Threading.Mutex].GetConstructor($constructorTypes)
    if ($null -ne $securityAwareConstructor) {
        $arguments = [object[]]::new(4)
        $arguments[0] = $false
        $arguments[1] = $Name
        $arguments[2] = $createdNew
        $arguments[3] = $security
        try {
            $creationMutex = $securityAwareConstructor.Invoke($arguments)
        }
        catch {
            return Open-OOShutUpMutationMutex -Name $Name
        }
    }
    else {
        $mutexAclType = [type]::GetType('System.Threading.MutexAcl, System.Threading.AccessControl', $false)
        if ($null -eq $mutexAclType) {
            throw 'This PowerShell runtime cannot create a named mutex with the required SYSTEM and Administrators access rules.'
        }
        $createMethod = $mutexAclType.GetMethod(
            'Create',
            [System.Reflection.BindingFlags]'Public,Static',
            $null,
            [type[]]@(
                [bool],
                [string],
                [bool].MakeByRefType(),
                [System.Security.AccessControl.MutexSecurity]
            ),
            $null
        )
        if ($null -eq $createMethod) {
            throw 'The Windows named-mutex ACL creation method is unavailable.'
        }
        $arguments = [object[]]::new(4)
        $arguments[0] = $false
        $arguments[1] = $Name
        $arguments[2] = $createdNew
        $arguments[3] = $security
        try {
            $creationMutex = $createMethod.Invoke($null, $arguments)
        }
        catch {
            return Open-OOShutUpMutationMutex -Name $Name
        }
    }

    try {
        # Reopen and inspect even a successful create-or-open result; an existing name may have an unsafe DACL.
        return Open-OOShutUpMutationMutex -Name $Name
    }
    finally {
        if ($null -ne $creationMutex) {
            $creationMutex.Dispose()
        }
    }
}

function Release-OOShutUpMutationMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Threading.Mutex]$Mutex
    )

    $Mutex.ReleaseMutex()
}

function Open-OOShutUpMutationMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $requiredRights = [System.Security.AccessControl.MutexRights]::Synchronize -bor `
        [System.Security.AccessControl.MutexRights]::Modify -bor `
        [System.Security.AccessControl.MutexRights]::ReadPermissions
    $mutex = $null
    $mutexAclType = [type]::GetType('System.Threading.MutexAcl, System.Threading.AccessControl', $false)
    if ($null -ne $mutexAclType) {
        $openMethod = $mutexAclType.GetMethod(
            'OpenExisting',
            [System.Reflection.BindingFlags]'Public,Static',
            $null,
            [type[]]@([string], [System.Security.AccessControl.MutexRights]),
            $null
        )
        if ($null -ne $openMethod) {
            $arguments = [object[]]::new(2)
            $arguments[0] = $Name
            $arguments[1] = $requiredRights
            $mutex = $openMethod.Invoke($null, $arguments)
        }
    }
    if ($null -eq $mutex) {
        $mutex = [System.Threading.Mutex]::OpenExisting($Name, $requiredRights)
    }
    try {
        $security = [System.Security.AccessControl.MutexSecurity]::new(
            $Name,
            [System.Security.AccessControl.AccessControlSections]::Access
        )
        $rules = @($security.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        $requiredSids = @(
            'S-1-5-18',
            'S-1-5-32-544',
            [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        ) | Select-Object -Unique
        foreach ($sid in $requiredSids) {
            $hasRequiredAccess = $false
            foreach ($rule in $rules) {
                if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow -or
                    $rule.IdentityReference.Value -ne $sid) {
                    continue
                }
                $grantedRights = [long][int]$rule.MutexRights
                $rightsToCheck = [long][int]$requiredRights
                if (($grantedRights -band $rightsToCheck) -eq $rightsToCheck) {
                    $hasRequiredAccess = $true
                    break
                }
            }
            if (-not $hasRequiredAccess) {
                throw "The existing machine-wide O&O mutation lock does not grant required access to '$sid'."
            }
        }
        return $mutex
    }
    catch {
        $mutex.Dispose()
        throw
    }
}

function Invoke-WithOOShutUpMutationLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,

        [Parameter(Mandatory = $false)]
        [string]$MutexName = $script:OOShutUpMutationMutexName,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutMilliseconds = $script:OOShutUpMutationLockTimeoutMilliseconds
    )

    $mutex = $null
    $ownsMutex = $false
    $operationStarted = $false
    $succeeded = $false
    $value = $null
    $failureType = $null
    $failureMessage = $null

    try {
        $mutex = New-OOShutUpMutationMutex -Name $MutexName
        try {
            $ownsMutex = $mutex.WaitOne($TimeoutMilliseconds)
        }
        catch [System.Threading.AbandonedMutexException] {
            # Keep abandoned ownership until this process exits so queued callers also fail closed.
            $ownsMutex = $true
            $failureType = 'Abandoned'
            $failureMessage = 'The machine-wide O&O mutation lock was abandoned by an earlier run. This run made no System Restore registry or O&O changes. Inspect the SystemRestorePointCreationFrequency value and O&O settings, then retry.'
        }

        if (-not $failureType -and -not $ownsMutex) {
            $failureType = 'Timeout'
            $failureMessage = 'The machine-wide O&O mutation lock was not acquired within 600 seconds. This run made no System Restore registry or O&O changes. Wait for the active run to finish, then retry.'
        }
        elseif (-not $failureType) {
            $operationStarted = $true
            $value = & $Operation
            $succeeded = $true
        }
    }
    catch {
        if (-not $failureType) {
            $failureType = if ($operationStarted) { 'OperationFailed' } else { 'Unavailable' }
            if ($operationStarted) {
                $failureMessage = "The protected O&O operation failed: $($_.Exception.Message)"
            }
            else {
                $failureMessage = "The machine-wide O&O mutation lock could not be opened or ownership could not be verified. No protected operation was started. Run the installer as Administrator and check that the Global lock is available to Local System and Administrators: $($_.Exception.Message)"
            }
        }
        $succeeded = $false
    }
    finally {
        if ($ownsMutex -and $null -ne $mutex -and $failureType -ne 'Abandoned') {
            try {
                Release-OOShutUpMutationMutex -Mutex $mutex
            }
            catch {
                $succeeded = $false
                $failureType = 'ReleaseFailed'
                $releaseDetail = "The machine-wide O&O mutation lock could not be released or verified. The protected operation may have changed system state; inspect the System Restore frequency value and O&O settings before retrying: $($_.Exception.Message)"
                if (-not [string]::IsNullOrWhiteSpace($failureMessage)) {
                    $failureMessage = "$failureMessage $releaseDetail"
                }
                else {
                    $failureMessage = $releaseDetail
                }
            }
        }
        if ($null -ne $mutex) {
            try {
                $mutex.Dispose()
            }
            catch {
                $succeeded = $false
                $failureType = 'ReleaseFailed'
                $disposeDetail = "The machine-wide O&O mutation lock handle could not be closed. Inspect the System Restore frequency value and O&O settings before retrying: $($_.Exception.Message)"
                if (-not [string]::IsNullOrWhiteSpace($failureMessage)) {
                    $failureMessage = "$failureMessage $disposeDetail"
                }
                else {
                    $failureMessage = $disposeDetail
                }
            }
        }
    }

    return [pscustomobject]@{
        Succeeded = [bool]$succeeded
        OperationStarted = [bool]$operationStarted
        Value = $value
        FailureType = $failureType
        FailureMessage = $failureMessage
    }
}

function Get-RestorePointFrequencyState {
    [CmdletBinding()]
    param()

    $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($script:OOShutUpRestorePointFrequencyRegistryPath, $false)
    if ($null -eq $key) {
        return [pscustomobject]@{
            KeyExists = $false
            ValueExists = $false
            Value = $null
            Kind = $null
        }
    }

    try {
        $valueExists = $key.GetValueNames() -contains $script:OOShutUpRestorePointFrequencyRegistryName
        if (-not $valueExists) {
            return [pscustomobject]@{
                KeyExists = $true
                ValueExists = $false
                Value = $null
                Kind = $null
            }
        }

        return [pscustomobject]@{
            KeyExists = $true
            ValueExists = $true
            Value = $key.GetValue(
                $script:OOShutUpRestorePointFrequencyRegistryName,
                $null,
                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
            )
            Kind = $key.GetValueKind($script:OOShutUpRestorePointFrequencyRegistryName)
        }
    }
    finally {
        $key.Dispose()
    }
}

function Set-RestorePointFrequencyRawValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryValueKind]$Kind
    )

    $key = [Microsoft.Win32.Registry]::LocalMachine.CreateSubKey($script:OOShutUpRestorePointFrequencyRegistryPath)
    if ($null -eq $key) {
        throw 'The System Restore registry key could not be opened for writing.'
    }

    try {
        $key.SetValue($script:OOShutUpRestorePointFrequencyRegistryName, $Value, $Kind)
        $key.Flush()
    }
    finally {
        $key.Dispose()
    }
}

function Set-RestorePointFrequencyValue {
    [CmdletBinding()]
    param()

    Set-RestorePointFrequencyRawValue -Value 0 -Kind ([Microsoft.Win32.RegistryValueKind]::DWord)
}

function Remove-RestorePointFrequencyValue {
    [CmdletBinding()]
    param()

    $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($script:OOShutUpRestorePointFrequencyRegistryPath, $true)
    if ($null -eq $key) {
        return
    }

    try {
        if ($key.GetValueNames() -contains $script:OOShutUpRestorePointFrequencyRegistryName) {
            $key.DeleteValue($script:OOShutUpRestorePointFrequencyRegistryName, $false)
            $key.Flush()
        }
    }
    finally {
        $key.Dispose()
    }
}

function Remove-EmptyRestorePointFrequencyKey {
    [CmdletBinding()]
    param()

    $separator = $script:OOShutUpRestorePointFrequencyRegistryPath.LastIndexOf('\')
    if ($separator -le 0) {
        throw 'The System Restore registry path is invalid.'
    }

    $parentPath = $script:OOShutUpRestorePointFrequencyRegistryPath.Substring(0, $separator)
    $leafName = $script:OOShutUpRestorePointFrequencyRegistryPath.Substring($separator + 1)
    $parentKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($parentPath, $true)
    if ($null -eq $parentKey) {
        return
    }

    try {
        $key = $parentKey.OpenSubKey($leafName, $false)
        if ($null -eq $key) {
            return
        }

        try {
            if ($key.GetValueNames().Count -gt 0 -or $key.GetSubKeyNames().Count -gt 0) {
                throw 'The System Restore registry key contains data beyond the temporary frequency value.'
            }
        }
        finally {
            $key.Dispose()
        }

        $parentKey.DeleteSubKey($leafName, $false)
        $parentKey.Flush()
    }
    finally {
        $parentKey.Dispose()
    }
}

function Test-RestorePointFrequencyStateEqual {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Expected,

        [Parameter(Mandatory = $true)]
        [psobject]$Actual
    )

    if ([bool]$Expected.KeyExists -ne [bool]$Actual.KeyExists -or
        [bool]$Expected.ValueExists -ne [bool]$Actual.ValueExists) {
        return $false
    }
    if (-not $Expected.ValueExists) {
        return $true
    }
    if ($Expected.Kind -ne $Actual.Kind) {
        return $false
    }

    if ($Expected.Value -is [Array] -or $Actual.Value -is [Array]) {
        return [System.Collections.StructuralComparisons]::StructuralEqualityComparer.Equals($Expected.Value, $Actual.Value)
    }
    return [object]::Equals($Expected.Value, $Actual.Value)
}

function Restore-RestorePointFrequencyState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$State
    )

    try {
        if ([bool]$State.KeyExists) {
            if ([bool]$State.ValueExists) {
                Set-RestorePointFrequencyRawValue -Value $State.Value -Kind $State.Kind
            }
            else {
                Remove-RestorePointFrequencyValue
            }
        }
        else {
            Remove-RestorePointFrequencyValue
            Remove-EmptyRestorePointFrequencyKey
        }

        $actualState = Get-RestorePointFrequencyState
        if (-not (Test-RestorePointFrequencyStateEqual -Expected $State -Actual $actualState)) {
            throw 'The prior System Restore registry state did not match after rollback.'
        }

        return $true
    }
    catch {
        Write-Log "Warning: Could not restore the prior System Restore registry state: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Set-RestorePointFrequency {
    <#
    .SYNOPSIS
        Temporarily permits another restore point and returns the prior registry state.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $previousState = $null
    try {
        Write-Log 'Configuring registry to allow frequent restore points...' -Level INFO
        $previousState = Get-RestorePointFrequencyState

        Set-RestorePointFrequencyValue

        $expectedState = [pscustomobject]@{
            KeyExists = $true
            ValueExists = $true
            Value = 0
            Kind = [Microsoft.Win32.RegistryValueKind]::DWord
        }
        $actualState = Get-RestorePointFrequencyState
        if (-not (Test-RestorePointFrequencyStateEqual -Expected $expectedState -Actual $actualState)) {
            throw 'The System Restore frequency registry write could not be verified.'
        }

        Write-Log 'Registry configured and verified: restore points can be created temporarily' -Level SUCCESS
        return [pscustomobject]@{
            Succeeded = $true
            PreviousState = $previousState
            RollbackSucceeded = $null
            ErrorMessage = $null
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        $rollbackSucceeded = $false
        if ($null -ne $previousState) {
            $rollbackSucceeded = Restore-RestorePointFrequencyState -State $previousState
        }

        Write-Log "Warning: Could not configure restore-point frequency: $errorMessage" -Level WARNING
        return [pscustomobject]@{
            Succeeded = $false
            PreviousState = $previousState
            RollbackSucceeded = $rollbackSucceeded
            ErrorMessage = $errorMessage
        }
    }
}

function Get-OOShutUpFileVersionInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
}

function Test-OOShutUpPayload {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [switch]$Quiet
    )

    $isVerified = $false

    try {
        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw 'The payload file is missing.'
        }

        $payloadFile = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($payloadFile.Length -le 0) {
            throw 'The payload file is empty.'
        }

        $signature = Get-AuthenticodeSignature -FilePath $Path -ErrorAction Stop
        if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate) {
            throw 'The payload does not have a valid Authenticode signature.'
        }

        $publisher = $signature.SignerCertificate.GetNameInfo(
            [System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,
            $false
        )
        if ($publisher -cne $ExpectedOOShutUpPublisher) {
            throw 'The payload publisher is not the approved O&O publisher.'
        }

        $versionInfo = Get-OOShutUpFileVersionInfo -Path $Path
        if ($versionInfo.ProductName -notmatch '^O&O ShutUp10(\+\+)?( Free( Edition)?)?$') {
            throw 'The payload is not the O&O ShutUp10 product.'
        }

        $fileVersion = $null
        if (-not [version]::TryParse([string]$versionInfo.FileVersion, [ref]$fileVersion) -or
            $fileVersion -lt $MinimumOOShutUpVersion) {
            throw 'The payload version is missing or older than the supported release.'
        }

        $isVerified = $true
    }
    catch {
        $isVerified = $false
    }

    if (-not $isVerified -and -not $Quiet) {
        Write-Log $OOShutUpPayloadFailureMessage -Level ERROR
    }

    return $isVerified
}

function Test-OOShutUpInstalled {
    <#
    .SYNOPSIS
        Checks if a verified O&O ShutUp10++ executable is available.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $executablePath = Get-OOShutUpPath
    if ($executablePath) {
        Write-Log "Found verified O&O ShutUp10++ at: $executablePath" -Level SUCCESS
        return $true
    }

    return $false
}

function Get-OOShutUpPath {
    <#
    .SYNOPSIS
        Gets a verified O&O ShutUp10++ executable path.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExpectedPath
    )

    $path = [System.IO.Path]::GetFullPath((Join-Path $InstallPath $ExecutableName))
    if (-not [string]::IsNullOrWhiteSpace($ExpectedPath)) {
        try {
            $canonicalExpectedPath = [System.IO.Path]::GetFullPath($ExpectedPath)
        }
        catch {
            return $null
        }
        if (-not [string]::Equals($path, $canonicalExpectedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $null
        }
    }

    if (Test-Path -LiteralPath $path -PathType Leaf) {
        if (Test-OOShutUpPayload -Path $path -Quiet) {
            return $path
        }

        Write-Log $OOShutUpPayloadFailureMessage -Level ERROR
    }

    return $null
}

function Get-OOShutUpApprovedProfile {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$RequireMachineScope
    )

    $failureMessage = $script:OOShutUpProfileFailureMessage

    try {
        if (-not (Test-Path -LiteralPath $script:OOShutUpProfilePath -PathType Leaf) -or
            -not (Test-Path -LiteralPath $script:OOShutUpProfileManifestPath -PathType Leaf)) {
            throw 'The profile package is incomplete.'
        }

        $manifestBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfileManifestPath)
        $manifestHashAlgorithm = [System.Security.Cryptography.SHA256]::Create()
        try {
            $manifestDigest = ([System.BitConverter]::ToString($manifestHashAlgorithm.ComputeHash($manifestBytes))).Replace('-', '').ToLowerInvariant()
        }
        finally {
            $manifestHashAlgorithm.Dispose()
        }
        $manifestEncoding = New-Object System.Text.UTF8Encoding($false, $true)
        $manifestText = $manifestEncoding.GetString($manifestBytes)
        if ($manifestText.Length -gt 0 -and $manifestText[0] -eq [char]0xFEFF) {
            $manifestText = $manifestText.Substring(1)
        }
        $manifest = ConvertFrom-Json -InputObject $manifestText -ErrorAction Stop
        $manifestPropertyNames = @($manifest.PSObject.Properties.Name)
        $requiredManifestProperties = @('schemaVersion', 'profileVersion', 'sha256', 'settingIds', 'settingDetails')
        if ($RequireMachineScope -and
            (($manifestPropertyNames -cnotcontains 'schemaVersion') -or $manifest.schemaVersion -ne 3)) {
            $failureMessage = $script:OOShutUpUnknownScopeFailureMessage
        }
        if ($manifestPropertyNames.Count -ne $requiredManifestProperties.Count) {
            throw 'The profile approval manifest has an unexpected shape.'
        }
        foreach ($propertyName in $requiredManifestProperties) {
            if ($manifestPropertyNames -cnotcontains $propertyName) {
                throw 'The profile approval manifest is incomplete.'
            }
        }

        if (($manifest.schemaVersion -isnot [int] -and $manifest.schemaVersion -isnot [long]) -or
            [long]$manifest.schemaVersion -ne 3 -or
            $manifest.profileVersion -isnot [string] -or
            $manifest.profileVersion -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$' -or
            $manifest.sha256 -isnot [string] -or
            $manifest.sha256 -notmatch '^[A-Fa-f0-9]{64}$' -or
            $manifest.settingIds -is [string] -or
            $null -eq $manifest.settingIds -or
            $manifest.settingDetails -isnot [System.Array] -or
            $null -eq $manifest.settingDetails) {
            throw 'The profile approval manifest contains invalid metadata.'
        }

        $allowedSettingIds = @($manifest.settingIds)
        if ($allowedSettingIds.Count -eq 0) {
            throw 'The profile approval manifest does not allow any settings.'
        }
        foreach ($settingId in $allowedSettingIds) {
            if ($settingId -isnot [string] -or $settingId -cnotmatch '^[A-Z][A-Z0-9]{1,15}$') {
                throw 'The profile approval manifest contains an invalid setting ID.'
            }
        }
        if (@($allowedSettingIds | Select-Object -Unique).Count -ne $allowedSettingIds.Count) {
            throw 'The profile approval manifest contains duplicate setting IDs.'
        }

        $settingDetails = @($manifest.settingDetails)
        if ($settingDetails.Count -eq 0 -or $settingDetails.Count -ne $allowedSettingIds.Count) {
            throw 'The profile approval manifest must describe every approved setting exactly once.'
        }
        foreach ($settingDetail in $settingDetails) {
            $detailPropertyNames = @($settingDetail.PSObject.Properties.Name)
            $requiredDetailProperties = @('id', 'state', 'category', 'impact', 'scope', 'scopeEvidence')
            if ($RequireMachineScope -and
                ($detailPropertyNames -cnotcontains 'scope' -or $detailPropertyNames -cnotcontains 'scopeEvidence')) {
                $failureMessage = $script:OOShutUpUnknownScopeFailureMessage
            }
            if ($detailPropertyNames.Count -ne $requiredDetailProperties.Count) {
                throw 'A reviewed setting detail has an unexpected shape.'
            }
            foreach ($propertyName in $requiredDetailProperties) {
                if ($detailPropertyNames -cnotcontains $propertyName) {
                    throw 'A reviewed setting detail is incomplete.'
                }
            }
            if ($settingDetail.id -isnot [string] -or
                $settingDetail.id -cnotmatch '^[A-Z][A-Z0-9]{1,15}$' -or
                $settingDetail.state -isnot [string] -or
                $settingDetail.state -cnotmatch '^[+-]$' -or
                $settingDetail.category -isnot [string] -or
                [string]::IsNullOrWhiteSpace($settingDetail.category) -or
                $settingDetail.category.Length -gt 64 -or
                $settingDetail.category -match '[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]' -or
                $settingDetail.impact -isnot [string] -or
                [string]::IsNullOrWhiteSpace($settingDetail.impact) -or
                $settingDetail.impact.Length -gt 240 -or
                $settingDetail.impact -match '[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]' -or
                $settingDetail.scope -isnot [string] -or
                $settingDetail.scope -cnotin @('machine', 'current-user') -or
                $settingDetail.scopeEvidence -isnot [string] -or
                [string]::IsNullOrWhiteSpace($settingDetail.scopeEvidence) -or
                $settingDetail.scopeEvidence.Length -gt 256 -or
                $settingDetail.scopeEvidence -match '[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]') {
                if ($RequireMachineScope) {
                    $failureMessage = $script:OOShutUpUnknownScopeFailureMessage
                }
                throw 'A reviewed setting detail contains invalid impact metadata.'
            }
            if ($RequireMachineScope -and $settingDetail.scope -ceq 'current-user') {
                $failureMessage = $script:OOShutUpCurrentUserScopeFailureMessage
                throw 'The profile contains current-user settings and cannot run as SYSTEM.'
            }
        }
        $detailIds = @($settingDetails | ForEach-Object { [string]$_.id })
        if (@($detailIds | Select-Object -Unique).Count -ne $detailIds.Count) {
            throw 'The profile approval manifest contains duplicate reviewed setting details.'
        }
        foreach ($settingId in $allowedSettingIds) {
            if ($detailIds -cnotcontains $settingId) {
                throw 'The profile approval manifest does not describe every approved setting.'
            }
        }

        $profileBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfilePath)
        if ($profileBytes.Length -eq 0) {
            throw 'The settings profile is empty.'
        }

        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        try {
            $actualDigest = ([System.BitConverter]::ToString($sha256.ComputeHash($profileBytes))).Replace('-', '').ToLowerInvariant()
        }
        finally {
            $sha256.Dispose()
        }
        if (-not [string]::Equals($actualDigest, [string]$manifest.sha256, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'The settings profile digest does not match its approval manifest.'
        }

        # StreamReader detects UTF-8/UTF-16 byte-order marks; strict UTF-8 rejects malformed byte sequences.
        $utf8Encoding = New-Object System.Text.UTF8Encoding($false, $true)
        $reader = New-Object System.IO.StreamReader($script:OOShutUpProfilePath, $utf8Encoding, $true)
        try {
            $profileText = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $profileSettingIds = New-Object System.Collections.Generic.List[string]
        $profileSettingStates = @{}
        foreach ($line in ($profileText -split "`r?`n")) {
            if ([string]::IsNullOrWhiteSpace($line) -or $line.TrimStart().StartsWith('#')) {
                continue
            }

            $lineMatch = [System.Text.RegularExpressions.Regex]::Match(
                $line,
                '^\s*([A-Z][A-Z0-9]{1,15})[\t;]([+-])(?:[\t;].*)?\s*$'
            )
            if (-not $lineMatch.Success) {
                throw 'The settings profile contains a malformed entry.'
            }

            $settingId = $lineMatch.Groups[1].Value
            $settingState = $lineMatch.Groups[2].Value
            if ($profileSettingIds.Contains($settingId)) {
                throw 'The settings profile contains a duplicate setting ID.'
            }
            if ($allowedSettingIds -cnotcontains $settingId) {
                throw 'The settings profile contains an unapproved setting ID.'
            }
            $profileSettingIds.Add($settingId)
            $profileSettingStates[$settingId] = $settingState
        }

        if ($profileSettingIds.Count -eq 0 -or $profileSettingIds.Count -ne $allowedSettingIds.Count) {
            throw 'The settings profile is empty or does not match its approved setting IDs.'
        }
        foreach ($settingId in $allowedSettingIds) {
            if (-not $profileSettingIds.Contains($settingId)) {
                throw 'The settings profile does not match its approved setting IDs.'
            }
        }
        foreach ($settingDetail in $settingDetails) {
            if ($profileSettingStates[[string]$settingDetail.id] -cne [string]$settingDetail.state) {
                throw 'A reviewed setting state does not match the settings profile.'
            }
        }

        return [pscustomobject]@{
            Path          = [System.IO.Path]::GetFullPath($script:OOShutUpProfilePath)
            Version       = [string]$manifest.profileVersion
            Sha256        = $actualDigest
            ManifestSha256 = $manifestDigest
            SettingCount  = $profileSettingIds.Count
            Settings      = @($settingDetails | Sort-Object -Property category, id | ForEach-Object {
                [pscustomobject]@{
                    Id       = [string]$_.id
                    State    = [string]$_.state
                    Category = [string]$_.category
                    Impact   = [string]$_.impact
                    Scope    = [string]$_.scope
                    ScopeEvidence = [string]$_.scopeEvidence
                }
            })
        }
    }
    catch {
        Write-Log $failureMessage -Level ERROR
        return $null
    }
}

function Show-OOShutUpProfileImpactSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Profile
    )

    Write-Host ''
    Write-Host 'Review this O&O ShutUp10 settings profile before any system changes:'
    Write-Host ('Profile version: {0}' -f $Profile.Version)
    Write-Host ('Profile SHA-256: {0}' -f $Profile.Sha256)
    Write-Host ('Profile manifest SHA-256: {0}' -f $Profile.ManifestSha256)
    Write-Host ('Settings: {0}' -f $Profile.SettingCount)
    Write-Host 'Affected categories and reviewed setting descriptions:'
    foreach ($setting in $Profile.Settings) {
        $stateLabel = if ($setting.State -eq '+') { 'on' } else { 'off' }
        Write-Host ('- {0}: {1} ({2}; {3} scope): {4} [Scope evidence: {5}]' -f $setting.Category, $setting.Id, $stateLabel, $setting.Scope, $setting.Impact, $setting.ScopeEvidence)
    }
    Write-Host 'The category and impact text comes from the reviewed manifest; no setting names are inferred from IDs.'
    Write-Host 'O&O settings can affect Windows features. Review every entry before accepting.'
}

function Confirm-OOShutUpProfile {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Profile,

        [Parameter(Mandatory = $false)]
        [string]$ApprovedProfileSha256 = '',

        [Parameter(Mandatory = $false)]
        [string]$ApprovedProfileManifestSha256 = '',

        [Parameter(Mandatory = $false)]
        [switch]$NonInteractive,

        [Parameter(Mandatory = $false)]
        [switch]$ProfileApprovalConfirmed
    )

    $script:OOShutUpApprovalConfirmed = $false
    $script:OOShutUpApprovedDigest = $null
    $script:OOShutUpApprovedManifestDigest = $null
    $hasMatchingDigest = $ApprovedProfileSha256 -match '^[A-Fa-f0-9]{64}$' -and
        [string]::Equals($ApprovedProfileSha256, $Profile.Sha256, [System.StringComparison]::OrdinalIgnoreCase)
    $hasMatchingManifestDigest = $ApprovedProfileManifestSha256 -match '^[A-Fa-f0-9]{64}$' -and
        [string]::Equals($ApprovedProfileManifestSha256, $Profile.ManifestSha256, [System.StringComparison]::OrdinalIgnoreCase)

    if ($NonInteractive -or $ProfileApprovalConfirmed) {
        if (-not $hasMatchingDigest -or -not $hasMatchingManifestDigest) {
            Write-Log $script:OOShutUpApprovalFailureMessage -Level ERROR
            return $false
        }

        if (-not $ProfileApprovalConfirmed) {
            Write-Log ('Unattended approval matches profile SHA-256 {0} and manifest SHA-256 {1}.' -f $Profile.Sha256, $Profile.ManifestSha256) -Level INFO
        }
        $script:OOShutUpApprovalConfirmed = $true
        $script:OOShutUpApprovedDigest = $Profile.Sha256
        $script:OOShutUpApprovedManifestDigest = $Profile.ManifestSha256
        return $true
    }

    if ((-not [string]::IsNullOrWhiteSpace($ApprovedProfileSha256) -and -not $hasMatchingDigest) -or
        (-not [string]::IsNullOrWhiteSpace($ApprovedProfileManifestSha256) -and -not $hasMatchingManifestDigest)) {
        Write-Log $script:OOShutUpApprovalFailureMessage -Level ERROR
        return $false
    }

    Show-OOShutUpProfileImpactSummary -Profile $Profile
    while ($true) {
        try {
            $decision = Read-Host 'Type YES to apply, NO to reject, or CANCEL to stop'
        }
        catch [System.Management.Automation.PipelineStoppedException] {
            Write-Log 'Profile approval was canceled. No settings were changed.' -Level WARNING
            return $false
        }

        switch -Regex (($decision -as [string]).Trim().ToUpperInvariant()) {
            '^YES$' {
                $script:OOShutUpApprovalConfirmed = $true
                $script:OOShutUpApprovedDigest = $Profile.Sha256
                $script:OOShutUpApprovedManifestDigest = $Profile.ManifestSha256
                return $true
            }
            '^NO$' {
                Write-Log 'Profile approval was declined. No settings were changed.' -Level INFO
                return $false
            }
            '^(CANCEL)?$' {
                Write-Log 'Profile approval was canceled. No settings were changed.' -Level INFO
                return $false
            }
            default {
                Write-Host 'Enter YES, NO, or CANCEL.'
            }
        }
    }
}

function New-OOShutUpWebClient {
    return New-Object System.Net.WebClient
}

function Move-OOShutUpFileAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $moveWithOverwrite = [System.IO.File].GetMethod('Move', [type[]]@([string], [string], [bool]))
    if ($null -ne $moveWithOverwrite) {
        [System.IO.File]::Move($SourcePath, $DestinationPath, $true)
        return
    }

    if (-not ('OOShutUp10AtomicFilePromotion' -as [type])) {
        $interopSource = @'
using System.Runtime.InteropServices;

public static class OOShutUp10AtomicFilePromotion
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MoveFileEx(string existingFileName, string newFileName, uint flags);
}
'@
        Add-Type -TypeDefinition $interopSource -ErrorAction Stop
    }

    # MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH; source and destination share a volume.
    $promoted = [OOShutUp10AtomicFilePromotion]::MoveFileEx($SourcePath, $DestinationPath, [uint32]9)
    if (-not $promoted) {
        $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw [System.ComponentModel.Win32Exception]::new($errorCode)
    }
}

function Install-OOShutUp {
    <#
    .SYNOPSIS
        Downloads and verifies O&O ShutUp10++ into its managed machine directory.
    #>
    [CmdletBinding()]
    param()

    $managedExecutablePath = [System.IO.Path]::GetFullPath((Join-Path $InstallPath $ExecutableName))
    $stagingPath = Join-Path $InstallPath ('{0}.{1}.download' -f $ExecutableName, [guid]::NewGuid().ToString('N'))

    try {
        if (-not (Test-Path -LiteralPath $InstallPath -PathType Container)) {
            New-Item -LiteralPath $InstallPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        Write-Log "Downloading O&O ShutUp10++ from: $DownloadUrl" -Level INFO

        try {
            $webClient = New-OOShutUpWebClient
            try {
                $webClient.DownloadFile($DownloadUrl, $stagingPath)
            }
            finally {
                if ($null -ne $webClient) {
                    $webClient.Dispose()
                }
            }
        }
        catch {
            Write-Log 'The O&O download failed. Check network connectivity and access to the managed machine directory, then retry.' -Level ERROR
            return $false
        }

        if (-not (Test-OOShutUpPayload -Path $stagingPath)) {
            return $false
        }

        try {
            Move-OOShutUpFileAtomically -SourcePath $stagingPath -DestinationPath $managedExecutablePath
        }
        catch {
            Write-Log 'The verified O&O download could not be promoted; the existing executable was left in place. Check managed directory access and retry.' -Level ERROR
            return $false
        }

        $fileSize = (Get-Item -LiteralPath $managedExecutablePath -ErrorAction Stop).Length / 1MB
        Write-Log "Successfully downloaded and verified O&O ShutUp10++ ($([Math]::Round($fileSize, 2)) MB)" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log 'The O&O executable could not be stored in the approved machine directory. The previous verified copy was left in place; check Program Files permissions and retry.' -Level ERROR
        return $false
    }
    finally {
        if (Test-Path -LiteralPath $stagingPath -PathType Leaf) {
            Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-OOShutUpProfileDigest {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha256.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Write-OOShutUpProfileSnapshotFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $stream = $null
    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None
        )
        $stream.Write($Bytes, 0, $Bytes.Length)
        $stream.Flush()
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Remove-OOShutUpProfileSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $installDirectory = [System.IO.Path]::GetFullPath($script:InstallPath)
    $snapshotDirectory = [System.IO.Path]::GetFullPath($Directory)
    $expectedNamePattern = '^ProfileSnapshot-[A-Fa-f0-9]{32}$'
    if (-not [string]::Equals(
        [System.IO.Path]::GetDirectoryName($snapshotDirectory),
        $installDirectory,
        [System.StringComparison]::OrdinalIgnoreCase
    ) -or [System.IO.Path]::GetFileName($snapshotDirectory) -notmatch $expectedNamePattern) {
        throw 'The snapshot path is outside the managed O&O directory or is not a per-run snapshot.'
    }

    if (-not [System.IO.Directory]::Exists($snapshotDirectory)) {
        return
    }
    $snapshotItem = Get-Item -LiteralPath $snapshotDirectory -Force -ErrorAction Stop
    if (-not $snapshotItem.PSIsContainer -or
        (([int]$snapshotItem.Attributes -band [int][System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
        throw 'The profile snapshot directory is not a regular directory.'
    }

    foreach ($fileName in @('ooshutup10.cfg', 'ooshutup10.approval.json')) {
        $filePath = Join-Path $snapshotDirectory $fileName
        if ([System.IO.File]::Exists($filePath)) {
            [System.IO.File]::Delete($filePath)
        }
    }
    [System.IO.Directory]::Delete($snapshotDirectory, $false)
}

function New-OOShutUpProfileSnapshot {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Profile
    )

    $snapshotDirectory = $null
    $stage = 'checking the managed O&O directory ACL'
    try {
        if ([string]$Profile.Sha256 -notmatch '^[A-Fa-f0-9]{64}$' -or
            [string]$Profile.ManifestSha256 -notmatch '^[A-Fa-f0-9]{64}$') {
            throw 'The approved profile digests are missing or malformed.'
        }
        if (-not (Test-SystemTaskTargetAcl -ScriptPath $script:InstallPath -AllowDirectory)) {
            throw 'The managed O&O directory ACL could not be verified for the profile snapshot.'
        }

        $snapshotDirectory = Join-Path $script:InstallPath ('ProfileSnapshot-{0}' -f [guid]::NewGuid().ToString('N'))
        $stage = 'creating the per-run snapshot directory'
        [System.IO.Directory]::CreateDirectory($snapshotDirectory) | Out-Null
        if (-not (Test-SystemTaskTargetAcl -ScriptPath $snapshotDirectory -AllowDirectory)) {
            throw 'The profile snapshot directory ACL could not be verified.'
        }

        $stage = 'reading the approved source files'
        $profileBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfilePath)
        $manifestBytes = [System.IO.File]::ReadAllBytes($script:OOShutUpProfileManifestPath)
        if (-not [string]::Equals(
            (Get-OOShutUpProfileDigest -Bytes $profileBytes),
            [string]$Profile.Sha256,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -or -not [string]::Equals(
            (Get-OOShutUpProfileDigest -Bytes $manifestBytes),
            [string]$Profile.ManifestSha256,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            throw 'The source profile package no longer matches both reviewed digests.'
        }

        $snapshotProfilePath = Join-Path $snapshotDirectory 'ooshutup10.cfg'
        $snapshotManifestPath = Join-Path $snapshotDirectory 'ooshutup10.approval.json'
        $stage = 'writing the exclusive profile snapshot'
        Write-OOShutUpProfileSnapshotFile -Path $snapshotProfilePath -Bytes $profileBytes
        Write-OOShutUpProfileSnapshotFile -Path $snapshotManifestPath -Bytes $manifestBytes

        $stage = 'verifying the snapshot ACL and copied digests'
        if (-not (Test-SystemTaskTargetAcl -ScriptPath $snapshotProfilePath) -or
            -not (Test-SystemTaskTargetAcl -ScriptPath $snapshotManifestPath)) {
            throw 'The copied profile package ACL could not be verified.'
        }
        if (-not [string]::Equals(
            (Get-OOShutUpProfileDigest -Bytes ([System.IO.File]::ReadAllBytes($snapshotProfilePath))),
            [string]$Profile.Sha256,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -or -not [string]::Equals(
            (Get-OOShutUpProfileDigest -Bytes ([System.IO.File]::ReadAllBytes($snapshotManifestPath))),
            [string]$Profile.ManifestSha256,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            throw 'The copied profile package does not match both reviewed digests.'
        }

        return [pscustomobject]@{
            Directory    = [System.IO.Path]::GetFullPath($snapshotDirectory)
            Path          = [System.IO.Path]::GetFullPath($snapshotProfilePath)
            ManifestPath  = [System.IO.Path]::GetFullPath($snapshotManifestPath)
        }
    }
    catch {
        $failure = $_.Exception.Message
        if ($snapshotDirectory -and [System.IO.Directory]::Exists($snapshotDirectory)) {
            try {
                Remove-OOShutUpProfileSnapshot -Directory $snapshotDirectory
            }
            catch {
                throw ("Snapshot stage '{0}' failed: {1} Cleanup also failed for '{2}': {3}" -f $stage, $failure, $snapshotDirectory, $_.Exception.Message)
            }
        }
        throw ("Snapshot stage '{0}' failed: {1}" -f $stage, $failure)
    }
}

function Invoke-OOShutUpConfigurationCore {
    <#
    .SYNOPSIS
        Applies an approved, digest-verified profile using a verified O&O ShutUp10++ executable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [Parameter(Mandatory = $false)]
        [switch]$SystemReapply
    )

    $script:OOShutUpLastRunChanged = $false
    $snapshot = $null
    $stage = 'profile validation'
    $launchAttempted = $false

    try {
        $approvedProfile = Get-OOShutUpApprovedProfile -RequireMachineScope:$SystemReapply
        if ($null -eq $approvedProfile) {
            return $false
        }

        if (-not $script:OOShutUpApprovalConfirmed -or
            -not [string]::Equals($script:OOShutUpApprovedDigest, $approvedProfile.Sha256, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($script:OOShutUpApprovedManifestDigest, $approvedProfile.ManifestSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Log $script:OOShutUpApprovalFailureMessage -Level ERROR
            return $false
        }

        if (-not (Test-OOShutUpPayload -Path $ExecutablePath)) {
            return $false
        }

        $stage = 'protected profile snapshot'
        $snapshot = New-OOShutUpProfileSnapshot -Profile $approvedProfile

        # O&O ShutUp10++ command line arguments:
        # /quiet - Silent mode
        # /nosrp - Don't create restore point (we'll create our own)
        # Protected per-run profile snapshot - independent of the caller's working directory
        Write-Log ("Applying approved settings profile version {0} (SHA-256 {1}; {2} settings). Setting values are omitted from logs." -f $approvedProfile.Version, $approvedProfile.Sha256, $approvedProfile.SettingCount) -Level INFO

        $arguments = '/quiet /nosrp "{0}"' -f $snapshot.Path
        $profileWorkingDirectory = $snapshot.Directory
        $stage = 'O&O process launch'
        $launchAttempted = $true
        $processInfo = Start-Process -FilePath $ExecutablePath -ArgumentList $arguments -WorkingDirectory $profileWorkingDirectory -Wait -PassThru -NoNewWindow
        if ($null -eq $processInfo -or $null -eq $processInfo.PSObject.Properties['ExitCode']) {
            throw 'Start-Process did not return a completed O&O process result.'
        }
    }
    catch {
        Write-Log ("Error during {0}: {1}" -f $stage, $_.Exception.Message) -Level ERROR
        if ($snapshot) {
            if ($launchAttempted) {
                Write-Log ("The O&O launch outcome could not be verified. The protected profile snapshot was retained at '{0}'; check for a running OOSU10.exe process before removing it." -f $snapshot.Directory) -Level ERROR
            }
            else {
                try {
                    Remove-OOShutUpProfileSnapshot -Directory $snapshot.Directory
                }
                catch {
                    Write-Log ("The protected profile snapshot could not be removed at '{0}'. Check for OOSU10.exe before removing it manually: {1}" -f $snapshot.Directory, $_.Exception.Message) -Level ERROR
                }
            }
        }
        if ($stage -eq 'protected profile snapshot') {
            Write-Log $script:OOShutUpProfileSnapshotFailureMessage -Level ERROR
        }
        return $false
    }

    $processExitCode = [int]$processInfo.ExitCode
    if ($processExitCode -eq 30) {
        $script:OOShutUpLastRunChanged = $true
    }

    try {
        Remove-OOShutUpProfileSnapshot -Directory $snapshot.Directory
    }
    catch {
        Write-Log ("The protected profile snapshot could not be removed after O&O exited. Verify O&O's applied settings, then remove '{0}' manually after confirming no OOSU10.exe process still uses it: {1}" -f $snapshot.Directory, $_.Exception.Message) -Level ERROR
        return $false
    }

    # O&O ShutUp10++ exit codes: 0 = no changes, 30 = settings applied, other codes are failures.
    if ($processExitCode -eq 0 -or $processExitCode -eq 30) {
        if ($processExitCode -eq 30) {
            Write-Log "Successfully applied the approved privacy settings profile" -Level SUCCESS
            Write-Log "Settings were applied and changes were made (exit code 30 is normal)" -Level INFO
        }
        else {
            Write-Log 'O&O reported that no setting changes were needed for the approved profile.' -Level INFO
        }
        return $true
    }
    else {
        Write-Log "O&O ShutUp10++ exited with unexpected code: $processExitCode" -Level WARNING
        Write-Log "Settings may not have been applied correctly. Please check manually." -Level WARNING
        return $false
    }
}

function Invoke-OOShutUpConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [Parameter(Mandatory = $false)]
        [switch]$SystemReapply
    )

    $script:OOShutUpLastRunChanged = $false
    $lockResult = Invoke-WithOOShutUpMutationLock -Operation {
        Invoke-OOShutUpConfigurationCore -ExecutablePath $ExecutablePath -SystemReapply:$SystemReapply
    }
    if (-not $lockResult.Succeeded) {
        Write-Log $lockResult.FailureMessage -Level ERROR
        return $false
    }

    return [bool]$lockResult.Value
}

function Get-WindowsSystemVolume {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $windowsPath = $env:SystemRoot
    if ([string]::IsNullOrWhiteSpace($windowsPath)) {
        throw 'The Windows system directory could not be detected.'
    }

    $windowsPath = [System.IO.Path]::GetFullPath($windowsPath)
    $volume = [System.IO.Path]::GetPathRoot($windowsPath)
    if ([string]::IsNullOrWhiteSpace($volume) -or $volume -notmatch '^[A-Za-z]:\\$') {
        throw 'The Windows system volume is not a supported local drive.'
    }

    return $volume
}

function New-SystemRestorePointCore {
    <#
    .SYNOPSIS
        Creates and verifies a system restore point, restoring the prior frequency registry state afterward.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $preparationSucceeded = $false
    $restorePointVerified = $false
    $failureDetail = $null
    $systemVolume = $null
    $frequencyPreparation = $null
    $registryRollbackSucceeded = $true

    try {
        Write-Log 'Creating system restore point...' -Level INFO
        Write-Log 'This may take a few minutes. Please wait...' -Level INFO

        $systemVolume = Get-WindowsSystemVolume
        $frequencyPreparation = Set-RestorePointFrequency
        if (-not $frequencyPreparation.Succeeded) {
            $failureDetail = $frequencyPreparation.ErrorMessage
        }
        else {
            Enable-ComputerRestore -Drive $systemVolume -ErrorAction Stop
            $preparationSucceeded = $true

            $description = "Before O&O ShutUp10++ Privacy Settings - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$([guid]::NewGuid().ToString('N'))]"
            $job = Start-Job -ScriptBlock {
                param($desc)
                try {
                    Checkpoint-Computer -Description $desc -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
                    return @{ Success = $true; Message = 'Restore point created successfully' }
                }
                catch {
                    return @{ Success = $false; Message = $_.Exception.Message }
                }
            } -ArgumentList $description

            $elapsed = 0
            $maxWait = 300
            while ($job.State -eq 'Running' -and $elapsed -lt $maxWait) {
                $seconds = $elapsed % 60
                $minutes = [Math]::Floor($elapsed / 60)
                $timeString = if ($minutes -gt 0) { "$minutes min $seconds sec" } else { "$seconds seconds" }

                Write-Progress -Activity 'Creating System Restore Point' `
                    -Status "Elapsed time: $timeString" `
                    -PercentComplete (($elapsed / $maxWait) * 100) `
                    -CurrentOperation 'Please wait while Windows creates a restore point...'

                Start-Sleep -Seconds 1
                $elapsed++
            }

            if ($job.State -eq 'Running') {
                $failureDetail = "Restore point creation timed out after $maxWait seconds."
                Write-Log $failureDetail -Level WARNING
                Stop-Job -Job $job
                Remove-Job -Job $job -Force
            }
            else {
                $result = Receive-Job -Job $job -Wait -AutoRemoveJob
                if ($result -and $result.Success) {
                    $matchingRestorePoint = @(Get-ComputerRestorePoint -ErrorAction Stop |
                        Where-Object { [string]$_.Description -ceq $description } |
                        Select-Object -First 1)
                    if ($matchingRestorePoint.Count -gt 0) {
                        $restorePointVerified = $true
                    }
                    else {
                        $failureDetail = 'The restore provider reported success, but Windows did not list the requested restore point.'
                    }
                }
                else {
                    $errorMessage = if ($result.Message) { [string]$result.Message } else { 'Unknown provider error.' }
                    $failureDetail = "Could not create the restore point: $errorMessage"
                }
            }
        }
    }
    catch {
        $failureDetail = $_.Exception.Message
    }
    finally {
        if ($null -ne $frequencyPreparation -and $null -ne $frequencyPreparation.PreviousState) {
            try {
                $registryRollbackSucceeded = Restore-RestorePointFrequencyState -State $frequencyPreparation.PreviousState
            }
            catch {
                $registryRollbackSucceeded = $false
                $failureDetail = "The prior System Restore registry value could not be restored: $($_.Exception.Message)"
            }
            if (-not $registryRollbackSucceeded) {
                $preparationSucceeded = $false
                $restorePointVerified = $false
                if ([string]::IsNullOrWhiteSpace($failureDetail)) {
                    $failureDetail = 'The prior System Restore registry value could not be restored.'
                }
            }
        }

        try {
            Write-Progress -Activity 'Creating System Restore Point' -Completed
        }
        catch {
            Write-Log "Warning: Could not clear restore-point progress display: $($_.Exception.Message)" -Level WARNING
        }
    }

    if (-not $preparationSucceeded) {
        if (-not [string]::IsNullOrWhiteSpace($failureDetail)) {
            Write-Log "Restore-point preparation detail: $failureDetail" -Level ERROR
        }
    }
    elseif ($restorePointVerified) {
        Write-Log 'System restore point created and verified successfully' -Level SUCCESS
    }
    elseif (-not [string]::IsNullOrWhiteSpace($failureDetail)) {
        Write-Log "Warning: Could not create restore point: $failureDetail" -Level WARNING
    }

    return [pscustomobject]@{
        PreparationSucceeded = [bool]$preparationSucceeded
        RestorePointVerified = [bool]$restorePointVerified
        RegistryRollbackSucceeded = [bool]$registryRollbackSucceeded
        SystemVolume = $systemVolume
        FailureDetail = $failureDetail
    }
}

function New-SystemRestorePoint {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $lockResult = Invoke-WithOOShutUpMutationLock -Operation {
        New-SystemRestorePointCore
    }
    if (-not $lockResult.Succeeded) {
        Write-Log $lockResult.FailureMessage -Level ERROR
        return [pscustomobject]@{
            PreparationSucceeded = $false
            RestorePointVerified = $false
            RegistryRollbackSucceeded = $false
            SystemVolume = $null
            FailureDetail = $lockResult.FailureMessage
        }
    }

    return $lockResult.Value
}

function Invoke-OOShutUpConfigurationWithRestorePointGateCore {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [Parameter(Mandatory = $false)]
        [switch]$AllowWithoutRestorePoint
    )

    $restorePointResult = New-SystemRestorePointCore
    $restorePointVerified = [bool]$restorePointResult.RestorePointVerified
    if (-not $restorePointResult.PreparationSucceeded) {
        Write-Log $script:OOShutUpRestorePointPreparationFailureMessage -Level ERROR
        return [pscustomobject]@{
            PreparationSucceeded = $false
            RestorePointVerified = $false
            OverrideUsed = $false
            ConfigurationStarted = $false
            ConfigurationSucceeded = $false
        }
    }

    if (-not $restorePointVerified -and -not $AllowWithoutRestorePoint) {
        Write-Log $script:OOShutUpRestorePointFailureMessage -Level ERROR
        return [pscustomobject]@{
            PreparationSucceeded = $true
            RestorePointVerified = $false
            OverrideUsed = $false
            ConfigurationStarted = $false
            ConfigurationSucceeded = $false
        }
    }

    if (-not $restorePointVerified) {
        Write-Log $script:OOShutUpRestorePointOverrideMessage -Level WARNING
    }

    Update-ScriptProgress -PercentComplete 70 -Status "Step 4 of 5: Applying Privacy Settings" -CurrentOperation "Applying the approved privacy settings profile"
    Write-Log "Step 4: Applying the approved privacy settings profile..." -Level INFO
    $configurationSucceeded = Invoke-OOShutUpConfigurationCore -ExecutablePath $ExecutablePath

    return [pscustomobject]@{
        PreparationSucceeded = $true
        RestorePointVerified = [bool]$restorePointVerified
        OverrideUsed = [bool](-not $restorePointVerified -and $AllowWithoutRestorePoint)
        ConfigurationStarted = $true
        ConfigurationSucceeded = [bool]$configurationSucceeded
    }
}

function Invoke-OOShutUpConfigurationWithRestorePointGate {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [Parameter(Mandatory = $false)]
        [switch]$AllowWithoutRestorePoint
    )

    $script:OOShutUpLastRunChanged = $false
    $lockResult = Invoke-WithOOShutUpMutationLock -Operation {
        Invoke-OOShutUpConfigurationWithRestorePointGateCore `
            -ExecutablePath $ExecutablePath `
            -AllowWithoutRestorePoint:$AllowWithoutRestorePoint
    }
    if (-not $lockResult.Succeeded) {
        Write-Log $lockResult.FailureMessage -Level ERROR
        if ($lockResult.OperationStarted -and $null -ne $lockResult.Value) {
            $result = $lockResult.Value
            $result.ConfigurationSucceeded = $false
            $result | Add-Member -NotePropertyName FailureDetail -NotePropertyValue $lockResult.FailureMessage -Force
            return $result
        }

        return [pscustomobject]@{
            PreparationSucceeded = $false
            RestorePointVerified = $false
            OverrideUsed = $false
            ConfigurationStarted = $false
            ConfigurationSucceeded = $false
            FailureDetail = $lockResult.FailureMessage
        }
    }

    return $lockResult.Value
}

function ConvertTo-SystemTaskSid {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Identity
    )

    if ($null -eq $Identity) {
        return $null
    }

    $identityText = [string]$Identity
    if ($Identity -is [System.Security.Principal.SecurityIdentifier]) {
        return $Identity.Value
    }
    if ($identityText -match '^S-1-') {
        return $identityText
    }

    switch -Regex ($identityText) {
        '^(?i:SYSTEM|NT AUTHORITY\\SYSTEM)$' { return 'S-1-5-18' }
        '^(?i:BUILTIN\\Administrators|Administrators)$' { return 'S-1-5-32-544' }
        '^(?i:NT SERVICE\\TrustedInstaller)$' { return 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464' }
    }

    try {
        $account = New-Object System.Security.Principal.NTAccount -ArgumentList $identityText
        return $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        return $null
    }
}

function Test-SystemTaskTargetAcl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $false)]
        [switch]$AllowDirectory
    )

    try {
        $scriptItem = Get-Item -LiteralPath $ScriptPath -Force -ErrorAction Stop
        if ($scriptItem.PSIsContainer -and -not $AllowDirectory) {
            return $false
        }

        $canonicalScriptPath = [System.IO.Path]::GetFullPath($scriptItem.FullName)
        if (-not [string]::Equals($canonicalScriptPath, $scriptItem.FullName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }

        $trustedSids = @(
            'S-1-5-18',
            'S-1-5-32-544',
            'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464'
        )
        try {
            $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            $currentPrincipal = New-Object System.Security.Principal.WindowsPrincipal -ArgumentList $currentIdentity
            if ($currentIdentity.User -and $currentPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
                $trustedSids += $currentIdentity.User.Value
            }
        }
        catch {
            return $false
        }
        $fileWriteMask = [int](
            [System.Security.AccessControl.FileSystemRights]::Write -bor
            [System.Security.AccessControl.FileSystemRights]::Delete -bor
            [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
            [System.Security.AccessControl.FileSystemRights]::TakeOwnership
        )
        # Parent-directory create rights alone cannot replace this existing file; reject rights that can alter or remove the path.
        $directoryWriteMask = [int](
            [System.Security.AccessControl.FileSystemRights]::Delete -bor
            [System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor
            [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
            [System.Security.AccessControl.FileSystemRights]::TakeOwnership -bor
            0x10000000
        )
        $fileWriteMask = $fileWriteMask -bor 0x40000000 -bor 0x10000000
        $readDataMask = [int][System.Security.AccessControl.FileSystemRights]::ReadData
        $hasSystemRead = $false
        $hasAdministratorsRead = $false

        $pathsToInspect = New-Object System.Collections.ArrayList
        $currentPath = $canonicalScriptPath
        while (-not [string]::IsNullOrWhiteSpace($currentPath)) {
            [void]$pathsToInspect.Add($currentPath)
            $parent = [System.IO.Directory]::GetParent($currentPath)
            if ($null -eq $parent) {
                break
            }
            $currentPath = $parent.FullName
        }

        foreach ($pathToInspect in $pathsToInspect) {
            $item = Get-Item -LiteralPath $pathToInspect -Force -ErrorAction Stop
            if (([int]$item.Attributes -band [int][System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                return $false
            }

            $acl = Get-Acl -LiteralPath $pathToInspect -ErrorAction Stop
            $ownerSid = ConvertTo-SystemTaskSid -Identity $acl.Owner
            if ($trustedSids -notcontains $ownerSid) {
                return $false
            }

            foreach ($rule in @($acl.Access)) {
                if (([int]$rule.PropagationFlags -band [int][System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) {
                    # Inherit-only entries apply to descendants, which are checked separately as the path walk continues.
                    continue
                }
                $ruleSid = ConvertTo-SystemTaskSid -Identity $rule.IdentityReference
                $rights = [int]$rule.FileSystemRights
                $ruleType = [string]$rule.AccessControlType
                $writeMask = if ($item.PSIsContainer) { $directoryWriteMask } else { $fileWriteMask }
                if ($AllowDirectory -and $item.PSIsContainer -and
                    [string]::Equals($pathToInspect, $canonicalScriptPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    # A protected snapshot directory must not permit an untrusted identity to create entries in it.
                    $writeMask = $writeMask -bor [int][System.Security.AccessControl.FileSystemRights]::CreateFiles -bor
                        [int][System.Security.AccessControl.FileSystemRights]::CreateDirectories -bor 0x40000000
                }
                if ($item.PSIsContainer -and [string]::Equals(
                    [System.IO.Path]::GetFullPath($item.FullName),
                    [System.IO.Path]::GetPathRoot($item.FullName),
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    # A delete right on a volume/share root cannot rename the root or replace its children.
                    $writeMask = $directoryWriteMask -band (-bnot [int][System.Security.AccessControl.FileSystemRights]::Delete)
                }

                if ($ruleType -eq 'Allow' -and (($rights -band $writeMask) -ne 0) -and $trustedSids -notcontains $ruleSid) {
                    return $false
                }

                if ([string]::Equals($pathToInspect, $canonicalScriptPath, [System.StringComparison]::OrdinalIgnoreCase) -and
                    (($rights -band $readDataMask) -ne 0)) {
                    if ($ruleType -eq 'Deny') {
                        # A read deny may apply through group membership, so do not guess its effective result.
                        return $false
                    }
                    if ($ruleType -eq 'Allow' -and $ruleSid -eq 'S-1-5-18') {
                        $hasSystemRead = $true
                    }
                    if ($ruleType -eq 'Allow' -and $ruleSid -eq 'S-1-5-32-544') {
                        $hasAdministratorsRead = $true
                    }
                }
            }
        }

        return ($hasSystemRead -and $hasAdministratorsRead)
    }
    catch {
        # ACL, identity, canonical-path, and filesystem inspection errors all fail closed.
        return $false
    }
}

function Get-WindowsPowerShellExecutablePaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$WindowsDirectory = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows),

        [Parameter(Mandatory = $false)]
        [bool]$Is64BitOperatingSystem = [System.Environment]::Is64BitOperatingSystem,

        [Parameter(Mandatory = $false)]
        [bool]$Is64BitProcess = [System.Environment]::Is64BitProcess
    )

    try {
        if ([string]::IsNullOrWhiteSpace($WindowsDirectory) -or
            -not [System.IO.Path]::IsPathRooted($WindowsDirectory)) {
            return $null
        }

        $canonicalWindowsDirectory = [System.IO.Path]::GetFullPath($WindowsDirectory)
        if ($canonicalWindowsDirectory.StartsWith('\\', [System.StringComparison]::Ordinal) -or
            -not (Test-Path -LiteralPath $canonicalWindowsDirectory -PathType Container)) {
            return $null
        }

        $relativeInterpreterPath = 'WindowsPowerShell\v1.0\powershell.exe'
        $taskPath = [System.IO.Path]::GetFullPath((Join-Path $canonicalWindowsDirectory ("System32\{0}" -f $relativeInterpreterPath)))
        $inspectionPath = $taskPath
        if ($Is64BitOperatingSystem -and -not $Is64BitProcess) {
            $inspectionPath = [System.IO.Path]::GetFullPath((Join-Path $canonicalWindowsDirectory ("Sysnative\{0}" -f $relativeInterpreterPath)))
        }

        return [pscustomobject]@{
            TaskPath = $taskPath
            InspectionPath = $inspectionPath
        }
    }
    catch {
        return $null
    }
}

function Test-WindowsPowerShellPathIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Candidate,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedPath
    )

    try {
        $canonicalResolvedPath = [System.IO.Path]::GetFullPath($ResolvedPath)
        foreach ($allowedPath in @([string]$Candidate.TaskPath, [string]$Candidate.InspectionPath)) {
            if ([string]::Equals(
                $canonicalResolvedPath,
                [System.IO.Path]::GetFullPath($allowedPath),
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                return $true
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

function Get-TrustedWindowsPowerShellExecutable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$WindowsDirectory = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows),

        [Parameter(Mandatory = $false)]
        [bool]$Is64BitOperatingSystem = [System.Environment]::Is64BitOperatingSystem,

        [Parameter(Mandatory = $false)]
        [bool]$Is64BitProcess = [System.Environment]::Is64BitProcess
    )

    try {
        $candidate = Get-WindowsPowerShellExecutablePaths -WindowsDirectory $WindowsDirectory -Is64BitOperatingSystem $Is64BitOperatingSystem -Is64BitProcess $Is64BitProcess
        if ($null -eq $candidate) {
            return $null
        }

        $interpreter = Get-Item -LiteralPath $candidate.InspectionPath -Force -ErrorAction Stop
        if ($interpreter.PSIsContainer -or
            (([int]$interpreter.Attributes -band [int][System.IO.FileAttributes]::ReparsePoint) -ne 0) -or
            -not (Test-WindowsPowerShellPathIdentity -Candidate $candidate -ResolvedPath ([string]$interpreter.FullName)) -or
            -not (Test-SystemTaskTargetAcl -ScriptPath $candidate.InspectionPath)) {
            return $null
        }

        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($candidate.InspectionPath)
        if ([string]$versionInfo.CompanyName -cne 'Microsoft Corporation' -or
            [string]$versionInfo.FileDescription -notmatch '^(?i:Windows PowerShell)(?:\s|$)' -or
            [string]$versionInfo.ProductName -notmatch '(?i)Windows.*Operating System') {
            return $null
        }

        $signature = Get-AuthenticodeSignature -FilePath $candidate.InspectionPath -ErrorAction Stop
        if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate) {
            return $null
        }
        $publisher = $signature.SignerCertificate.GetNameInfo(
            [System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,
            $false
        )
        if ([string]$publisher -notmatch '^(?i:Microsoft Windows|Microsoft Windows Publisher)$') {
            return $null
        }

        return [string]$candidate.TaskPath
    }
    catch {
        return $null
    }
}

function Test-WindowsUpdateTaskDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Task,

        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $true)]
        [string]$TaskDescription,

        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileSha256,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileManifestSha256,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedOOSExecutablePath,

        [Parameter(Mandatory = $true)]
        [string]$TrustedPowerShellPath
    )

    try {
        if (-not [string]::Equals([string]$Task.TaskPath, $TaskPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.TaskName, $TaskName, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.Description, $TaskDescription, [System.StringComparison]::Ordinal)) {
            return $false
        }

        $principalSid = ConvertTo-SystemTaskSid -Identity $Task.Principal.UserId
        if ($principalSid -ne 'S-1-5-18' -or [string]$Task.Principal.RunLevel -notin @('Highest', 'HighestAvailable')) {
            return $false
        }

        $actions = @($Task.Actions)
        if ($actions.Count -ne 1) {
            return $false
        }

        $expectedPowerShellPaths = Get-WindowsPowerShellExecutablePaths
        $currentTrustedPowerShellPath = Get-TrustedWindowsPowerShellExecutable
        if ($null -eq $expectedPowerShellPaths -or
            [string]::IsNullOrWhiteSpace($currentTrustedPowerShellPath) -or
            -not [string]::Equals([string]$currentTrustedPowerShellPath, [string]$expectedPowerShellPaths.TaskPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$TrustedPowerShellPath, [string]$currentTrustedPowerShellPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$actions[0].Execute, [string]$currentTrustedPowerShellPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }

        $canonicalScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)
        $arguments = [string]$actions[0].Arguments
        $canonicalExecutablePath = [System.IO.Path]::GetFullPath($ApprovedOOSExecutablePath)
        $expectedArguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedOOSExecutablePath "{1}" -ApprovedProfileSha256 {2} -ApprovedProfileManifestSha256 {3}' -f $canonicalScriptPath, $canonicalExecutablePath, $ApprovedProfileSha256, $ApprovedProfileManifestSha256
        $expectedWorkingDirectory = [System.IO.Path]::GetDirectoryName($canonicalScriptPath)
        if (-not [string]::Equals($arguments, $expectedArguments, [System.StringComparison]::Ordinal) -or
            -not [string]::Equals([string]$actions[0].WorkingDirectory, $expectedWorkingDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }

        $settings = $Task.Settings
        if ($null -eq $settings -or
            -not [string]::Equals([string]$settings.MultipleInstances, $script:OOShutUpUpdateMultipleInstancesPolicy, [System.StringComparison]::Ordinal)) {
            return $false
        }

        $triggers = @($Task.Triggers)
        if ($triggers.Count -ne 1 -or
            -not [bool]$triggers[0].Enabled -or
            -not [string]::Equals([string]$triggers[0].Delay, $script:OOShutUpUpdateEventDelay, [System.StringComparison]::Ordinal) -or
            -not [string]::Equals([string]$triggers[0].Subscription, $script:OOShutUpUpdateEventSubscription, [System.StringComparison]::Ordinal)) {
            return $false
        }

        return $true
    }
    catch {
        return $false
    }
}

function Test-WindowsUpdateLegacyTaskDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Task,

        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $true)]
        [string]$TaskDescription,

        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $false)]
        [string]$ApprovedProfileSha256,

        [Parameter(Mandatory = $false)]
        [string]$ApprovedProfileManifestSha256
    )

    try {
        if (-not [string]::Equals([string]$Task.TaskPath, $TaskPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.TaskName, $TaskName, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.Description, $TaskDescription, [System.StringComparison]::Ordinal)) {
            return $false
        }

        $principalSid = ConvertTo-SystemTaskSid -Identity $Task.Principal.UserId
        if ($principalSid -ne 'S-1-5-18' -or [string]$Task.Principal.RunLevel -notin @('Highest', 'HighestAvailable')) {
            return $false
        }

        $actions = @($Task.Actions)
        if ($actions.Count -ne 1 -or
            -not [string]::Equals([string]$actions[0].Execute, 'PowerShell.exe', [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }

        $canonicalScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)
        $arguments = [string]$actions[0].Arguments
        $expectedLegacyArguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -ReapplyOnly -SkipRelocate' -f $canonicalScriptPath
        $expectedPreviousApprovedArguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedProfileSha256 {1} -ApprovedProfileManifestSha256 {2}' -f $canonicalScriptPath, $ApprovedProfileSha256, $ApprovedProfileManifestSha256
        $isLegacyDefinition = [string]::Equals($arguments, $expectedLegacyArguments, [System.StringComparison]::Ordinal) -and
            [string]::IsNullOrEmpty([string]$actions[0].WorkingDirectory)
        $isPreviousApprovedDefinition = -not [string]::IsNullOrWhiteSpace($ApprovedProfileSha256) -and
            -not [string]::IsNullOrWhiteSpace($ApprovedProfileManifestSha256) -and
            [string]::Equals($arguments, $expectedPreviousApprovedArguments, [System.StringComparison]::Ordinal) -and
            [string]::Equals([string]$actions[0].WorkingDirectory, [System.IO.Path]::GetDirectoryName($canonicalScriptPath), [System.StringComparison]::OrdinalIgnoreCase)
        if (-not ($isLegacyDefinition -or $isPreviousApprovedDefinition)) {
            return $false
        }

        return $true
    }
    catch {
        return $false
    }
}

function Test-ScheduledTaskAccessDenied {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $exception = $ErrorRecord.Exception
    while ($null -ne $exception) {
        if ($exception -is [System.UnauthorizedAccessException] -or
            $exception -is [System.Security.SecurityException] -or
            $exception.HResult -eq -2147024891 -or
            $exception.Message -match '(?i)(access is denied|access denied|permission denied)') {
            return $true
        }
        $exception = $exception.InnerException
    }

    return $false
}

function Invoke-WindowsUpdateTaskSchedulerOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('GetSecurityDescriptor', 'SetSecurityDescriptor', 'SetEnabled')]
        [string]$Operation,

        [Parameter(Mandatory = $false)]
        [string]$SecurityDescriptor,

        [Parameter(Mandatory = $false)]
        [bool]$Enabled
    )

    $service = $null
    $folder = $null
    $registeredTask = $null
    try {
        $service = New-Object -ComObject 'Schedule.Service'
        $service.Connect()
        $folder = $service.GetFolder($TaskPath)
        $registeredTask = $folder.GetTask($TaskName)

        switch ($Operation) {
            'GetSecurityDescriptor' {
                # OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION.
                return [string]$registeredTask.GetSecurityDescriptor(0x7)
            }
            'SetSecurityDescriptor' {
                # TASK_DONT_ADD_PRINCIPAL_ACE; SYSTEM is already in the explicit allowlist below.
                $registeredTask.SetSecurityDescriptor($SecurityDescriptor, 0x10)
                return $true
            }
            'SetEnabled' {
                $registeredTask.Enabled = $Enabled
                return $true
            }
        }
    }
    finally {
        foreach ($comObject in @($registeredTask, $folder, $service)) {
            if ($null -ne $comObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }
    }
}

function Get-WindowsUpdateTaskSecurityDescriptor {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName
    )

    try {
        $descriptor = Invoke-WindowsUpdateTaskSchedulerOperation -TaskPath $TaskPath -TaskName $TaskName -Operation GetSecurityDescriptor
        if ([string]::IsNullOrWhiteSpace([string]$descriptor)) {
            throw 'Task Scheduler returned an empty security descriptor.'
        }
        return [string]$descriptor
    }
    catch {
        throw ('{0} Details: {1}' -f $script:OOShutUpTaskSecurityDescriptorFailureMessage, $_.Exception.Message)
    }
}

function Set-WindowsUpdateTaskSecurityDescriptor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $true)]
        [string]$SecurityDescriptor
    )

    try {
        [void](Invoke-WindowsUpdateTaskSchedulerOperation -TaskPath $TaskPath -TaskName $TaskName -Operation SetSecurityDescriptor -SecurityDescriptor $SecurityDescriptor)
    }
    catch {
        throw ('{0} Details: {1}' -f $script:OOShutUpTaskSecurityDescriptorFailureMessage, $_.Exception.Message)
    }
}

function Set-WindowsUpdateTaskEnabled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    [void](Invoke-WindowsUpdateTaskSchedulerOperation -TaskPath $TaskPath -TaskName $TaskName -Operation SetEnabled -Enabled $Enabled)
}

function Test-WindowsUpdateTaskSecurityDescriptor {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SecurityDescriptor
    )

    try {
        $expected = [System.Security.AccessControl.RawSecurityDescriptor]::new($script:OOShutUpTaskSecurityDescriptorSddl)
        $actual = [System.Security.AccessControl.RawSecurityDescriptor]::new($SecurityDescriptor)
        if ($null -eq $actual.Owner -or $actual.Owner.Value -ne 'S-1-5-32-544' -or
            $null -eq $actual.Group -or $actual.Group.Value -ne 'S-1-5-18' -or
            $null -eq $actual.DiscretionaryAcl -or
            ($actual.ControlFlags -band [System.Security.AccessControl.ControlFlags]::DiscretionaryAclProtected) -eq 0 -or
            $actual.DiscretionaryAcl.Count -ne $expected.DiscretionaryAcl.Count) {
            return $false
        }

        $actualAces = @()
        foreach ($actualAce in $actual.DiscretionaryAcl) {
            if ($actualAce -isnot [System.Security.AccessControl.CommonAce] -or
                $actualAce.IsCallback -or
                $actualAce.AceType -ne [System.Security.AccessControl.AceType]::AccessAllowed -or
                $actualAce.AceFlags -ne [System.Security.AccessControl.AceFlags]::None) {
                return $false
            }
            $actualAces += $actualAce
        }

        foreach ($expectedAce in $expected.DiscretionaryAcl) {
            $matchingAces = @($actualAces | Where-Object {
                $_.AccessMask -eq $expectedAce.AccessMask -and
                $_.SecurityIdentifier.Value -eq $expectedAce.SecurityIdentifier.Value
            })
            if ($matchingAces.Count -ne 1) {
                return $false
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

function Test-WindowsUpdateTaskSecurityDescriptorEqual {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ActualSecurityDescriptor,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedSecurityDescriptor
    )

    try {
        $actual = [System.Security.AccessControl.RawSecurityDescriptor]::new($ActualSecurityDescriptor)
        $expected = [System.Security.AccessControl.RawSecurityDescriptor]::new($ExpectedSecurityDescriptor)
        if ($actual.BinaryLength -ne $expected.BinaryLength) {
            return $false
        }

        $actualBytes = [byte[]]::new($actual.BinaryLength)
        $expectedBytes = [byte[]]::new($expected.BinaryLength)
        $actual.GetBinaryForm($actualBytes, 0)
        $expected.GetBinaryForm($expectedBytes, 0)
        for ($index = 0; $index -lt $actualBytes.Length; $index++) {
            if ($actualBytes[$index] -ne $expectedBytes[$index]) {
                return $false
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

function Assert-WindowsUpdateTaskSecurityDescriptor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskPath,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $false)]
        [string]$ExpectedSecurityDescriptor
    )

    $descriptor = Get-WindowsUpdateTaskSecurityDescriptor -TaskPath $TaskPath -TaskName $TaskName
    if (-not (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $descriptor)) {
        throw $script:OOShutUpTaskSecurityDescriptorFailureMessage
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedSecurityDescriptor)) {
        if (-not (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $ExpectedSecurityDescriptor) -or
            -not (Test-WindowsUpdateTaskSecurityDescriptorEqual -ActualSecurityDescriptor $descriptor -ExpectedSecurityDescriptor $ExpectedSecurityDescriptor)) {
            throw $script:OOShutUpTaskSecurityDescriptorFailureMessage
        }
    }

    return $descriptor
}

function Test-WindowsUpdateTaskMatchesSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Task,

        [Parameter(Mandatory = $true)]
        [object]$Snapshot
    )

    try {
        if (-not [string]::Equals([string]$Task.TaskPath, [string]$Snapshot.TaskPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.TaskName, [string]$Snapshot.TaskName, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.Description, [string]$Snapshot.Description, [System.StringComparison]::Ordinal) -or
            -not [string]::Equals([string]$Task.Principal.UserId, [string]$Snapshot.PrincipalUserId, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$Task.Principal.RunLevel, [string]$Snapshot.PrincipalRunLevel, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }

        $actions = @($Task.Actions)
        if ($actions.Count -ne @($Snapshot.Actions).Count) {
            return $false
        }
        for ($index = 0; $index -lt $actions.Count; $index++) {
            if (-not [string]::Equals([string]$actions[$index].Execute, [string]$Snapshot.Actions[$index].Execute, [System.StringComparison]::OrdinalIgnoreCase) -or
                -not [string]::Equals([string]$actions[$index].Arguments, [string]$Snapshot.Actions[$index].Arguments, [System.StringComparison]::Ordinal) -or
                -not [string]::Equals([string]$actions[$index].WorkingDirectory, [string]$Snapshot.Actions[$index].WorkingDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $false
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

function New-WindowsUpdateScheduledTask {
    <#
    .SYNOPSIS
        Creates a scheduled task to reapply O&O ShutUp10++ settings after Windows updates
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileSha256,

        [Parameter(Mandatory = $false)]
        [string]$ApprovedProfileManifestSha256
    )

    $stagingTaskName = $null
    $stagingTaskAttempted = $false
    $replacementStarted = $false
    $taskSchedulerOperationStarted = $false
    $previousDefinitions = New-Object System.Collections.ArrayList

    try {
        if ($ApprovedProfileSha256 -notmatch '^[A-Fa-f0-9]{64}$' -or
            $ApprovedProfileManifestSha256 -notmatch '^[A-Fa-f0-9]{64}$') {
            Write-Log $script:OOShutUpApprovalFailureMessage -Level ERROR
            return $false
        }

        $currentApprovedProfile = Get-OOShutUpApprovedProfile -RequireMachineScope
        if ($null -eq $currentApprovedProfile) {
            return $false
        }
        if (-not $script:OOShutUpApprovalConfirmed -or
            -not [string]::Equals($currentApprovedProfile.Sha256, $ApprovedProfileSha256, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($script:OOShutUpApprovedDigest, $ApprovedProfileSha256, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($currentApprovedProfile.ManifestSha256, $ApprovedProfileManifestSha256, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals($script:OOShutUpApprovedManifestDigest, $ApprovedProfileManifestSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Log $script:OOShutUpApprovalFailureMessage -Level ERROR
            return $false
        }

        try {
            $ScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)
        }
        catch {
            Write-Log $script:OOShutUpTaskAclFailureMessage -Level ERROR
            return $false
        }
        if (-not (Test-SystemTaskTargetAcl -ScriptPath $ScriptPath)) {
            Write-Log $script:OOShutUpTaskAclFailureMessage -Level ERROR
            return $false
        }

        $approvedOOSExecutablePath = Get-OOShutUpPath
        if (-not $approvedOOSExecutablePath) {
            Write-Log $script:OOShutUpManagedPathFailureMessage -Level ERROR
            return $false
        }

        $expectedPowerShellPaths = Get-WindowsPowerShellExecutablePaths
        $trustedPowerShellPath = Get-TrustedWindowsPowerShellExecutable
        if ($null -eq $expectedPowerShellPaths -or
            [string]::IsNullOrWhiteSpace($trustedPowerShellPath) -or
            -not [string]::Equals([string]$trustedPowerShellPath, [string]$expectedPowerShellPaths.TaskPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Log $script:OOShutUpTrustedPowerShellFailureMessage -Level ERROR
            return $false
        }

        Write-Log ("Scheduled task principal SID S-1-5-18 passed the machine-scope gate for {0} settings (profile SHA-256 {1}; manifest SHA-256 {2}); setting values are omitted." -f $currentApprovedProfile.SettingCount, $currentApprovedProfile.Sha256, $currentApprovedProfile.ManifestSha256) -Level INFO

        Write-Log "Creating scheduled task for post-Windows Update execution..." -Level INFO
        Write-Log 'The task listens for Windows Update installation-success event 19 with a two-minute delay; overlapping reapplications are queued sequentially.' -Level INFO

        $taskName     = "Install-OOShutUp10"
        $taskFolder   = "myTech.Today"
        $taskPath     = "\$taskFolder\"
        $taskFullPath = "$taskPath$taskName"
        $taskDescription = "Reapplies O&O ShutUp10++ privacy settings after Windows updates to prevent Microsoft from resetting privacy configurations"

        # Build the task XML using XmlDocument for proper escaping
        $xml = New-Object System.Xml.XmlDocument

        # Create XML declaration
        $xmlDeclaration = $xml.CreateXmlDeclaration("1.0", "UTF-16", $null)
        $xml.AppendChild($xmlDeclaration) | Out-Null

        # Create root Task element
        $task = $xml.CreateElement("Task")
        $task.SetAttribute("version", "1.2")
        $task.SetAttribute("xmlns", "http://schemas.microsoft.com/windows/2004/02/mit/task")
        $xml.AppendChild($task) | Out-Null

        # RegistrationInfo
        $regInfo = $xml.CreateElement("RegistrationInfo")
        $desc = $xml.CreateElement("Description")
        $desc.InnerText = $taskDescription
        $regInfo.AppendChild($desc) | Out-Null
        $task.AppendChild($regInfo) | Out-Null

        # Triggers
        $triggers = $xml.CreateElement("Triggers")
        $eventTrigger = $xml.CreateElement("EventTrigger")

        $enabled = $xml.CreateElement("Enabled")
        $enabled.InnerText = "true"
        $eventTrigger.AppendChild($enabled) | Out-Null

        $subscription = $xml.CreateElement("Subscription")
        $subscription.InnerText = $script:OOShutUpUpdateEventSubscription
        $eventTrigger.AppendChild($subscription) | Out-Null

        $delay = $xml.CreateElement("Delay")
        $delay.InnerText = $script:OOShutUpUpdateEventDelay
        $eventTrigger.AppendChild($delay) | Out-Null

        $triggers.AppendChild($eventTrigger) | Out-Null
        $task.AppendChild($triggers) | Out-Null

        # Principals
        $principals = $xml.CreateElement("Principals")
        $principal = $xml.CreateElement("Principal")
        $principal.SetAttribute("id", "Author")

        $userId = $xml.CreateElement("UserId")
        $userId.InnerText = "S-1-5-18"
        $principal.AppendChild($userId) | Out-Null

        $runLevel = $xml.CreateElement("RunLevel")
        $runLevel.InnerText = "HighestAvailable"
        $principal.AppendChild($runLevel) | Out-Null

        $principals.AppendChild($principal) | Out-Null
        $task.AppendChild($principals) | Out-Null

        # Settings
        $settings = $xml.CreateElement("Settings")

        $multipleInstances = $xml.CreateElement("MultipleInstancesPolicy")
        $multipleInstances.InnerText = $script:OOShutUpUpdateMultipleInstancesPolicy
        $settings.AppendChild($multipleInstances) | Out-Null

        $disallowBattery = $xml.CreateElement("DisallowStartIfOnBatteries")
        $disallowBattery.InnerText = "false"
        $settings.AppendChild($disallowBattery) | Out-Null

        $stopBattery = $xml.CreateElement("StopIfGoingOnBatteries")
        $stopBattery.InnerText = "false"
        $settings.AppendChild($stopBattery) | Out-Null

        $startAvailable = $xml.CreateElement("StartWhenAvailable")
        $startAvailable.InnerText = "true"
        $settings.AppendChild($startAvailable) | Out-Null

        $enabled2 = $xml.CreateElement("Enabled")
        $enabled2.InnerText = "true"
        $settings.AppendChild($enabled2) | Out-Null

        $task.AppendChild($settings) | Out-Null

        # Actions
        $actions = $xml.CreateElement("Actions")
        $actions.SetAttribute("Context", "Author")

        $exec = $xml.CreateElement("Exec")

        $command = $xml.CreateElement("Command")
        $command.InnerText = $trustedPowerShellPath
        $exec.AppendChild($command) | Out-Null

        $arguments = $xml.CreateElement("Arguments")
        $arguments.InnerText = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedOOSExecutablePath `"$approvedOOSExecutablePath`" -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256"
        $exec.AppendChild($arguments) | Out-Null

        $workingDirectory = $xml.CreateElement("WorkingDirectory")
        $workingDirectory.InnerText = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ScriptPath))
        $exec.AppendChild($workingDirectory) | Out-Null

        $actions.AppendChild($exec) | Out-Null
        $task.AppendChild($actions) | Out-Null

        $taskXml = $xml.OuterXml

        $currentPowerShellPaths = Get-WindowsPowerShellExecutablePaths
        $currentTrustedPowerShellPath = Get-TrustedWindowsPowerShellExecutable
        if ($null -eq $currentPowerShellPaths -or
            [string]::IsNullOrWhiteSpace($currentTrustedPowerShellPath) -or
            -not [string]::Equals([string]$currentTrustedPowerShellPath, [string]$currentPowerShellPaths.TaskPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$currentTrustedPowerShellPath, [string]$trustedPowerShellPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Log $script:OOShutUpTrustedPowerShellFailureMessage -Level ERROR
            return $false
        }

        $stageDocument = New-Object System.Xml.XmlDocument
        $stageDocument.LoadXml($taskXml)
        $stageEnabled = $stageDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
        if ($null -eq $stageEnabled) {
            throw 'The candidate task XML has no Settings/Enabled element.'
        }
        $stageEnabled.InnerText = 'false'
        $stagingTaskXml = $stageDocument.OuterXml

        $finalDocument = New-Object System.Xml.XmlDocument
        $finalDocument.LoadXml($taskXml)
        $finalEnabled = $finalDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
        if ($null -eq $finalEnabled) {
            throw 'The replacement task XML has no Settings/Enabled element.'
        }
        $finalEnabled.InnerText = 'false'
        $finalTaskXml = $finalDocument.OuterXml

        # Keep task XML in memory; a temporary diagnostic copy is unnecessary and may expose approved digests.

        # Report same-name tasks elsewhere, but only inspect and replace the exact owned task path.
        try {
            $sameNameTasks = @(Get-ScheduledTask -TaskName $taskName -ErrorAction Stop)
            $existingTasks = @(Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction Stop)
        }
        catch {
            if (Test-ScheduledTaskAccessDenied -ErrorRecord $_) {
                Write-Log ("Access denied while inspecting Task Scheduler for '{0}'. No task was deleted; retry with an account that can inspect this path." -f $taskFullPath) -Level ERROR
            }
            else {
                Write-Log ("Task Scheduler could not inspect '{0}'. No task was deleted; inspect Task Scheduler and retry. Details: {1}" -f $taskFullPath, $_.Exception.Message) -Level ERROR
            }
            return $false
        }

        foreach ($sameNameTask in $sameNameTasks) {
            if ([string]::IsNullOrWhiteSpace([string]$sameNameTask.TaskPath) -or
                [string]::IsNullOrWhiteSpace([string]$sameNameTask.TaskName)) {
                Write-Log ("{0} Listed path: <unknown>. Task Scheduler returned incomplete identity metadata; no task was changed." -f $script:OOShutUpTaskOwnershipFailureMessage) -Level ERROR
                return $false
            }

            if (-not [string]::Equals([string]$sameNameTask.TaskPath, $taskPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $unownedPath = '{0}{1}' -f $sameNameTask.TaskPath, $sameNameTask.TaskName
                Write-Log ("Leaving same-named task '{0}' unchanged because it is outside the owned folder. Inspect that path in Task Scheduler and confirm ownership before taking any manual action." -f $unownedPath) -Level WARNING
            }
        }

        if ($existingTasks.Count -gt 1) {
            Write-Log ("{0} Listed path: {1}. Multiple tasks matched the exact owned identity; inspect this path and retry after confirming ownership." -f $script:OOShutUpTaskOwnershipFailureMessage, $taskFullPath) -Level ERROR
            return $false
        }

        foreach ($existingTask in $existingTasks) {
            $isCurrentDefinition = Test-WindowsUpdateTaskDefinition -Task $existingTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath -TrustedPowerShellPath $trustedPowerShellPath
            $isApprovedLegacyDefinition = Test-WindowsUpdateLegacyTaskDefinition -Task $existingTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256
            if (-not $isCurrentDefinition -and -not $isApprovedLegacyDefinition) {
                Write-Log ("{0} Listed path: {1}. The action or description does not match an approved definition." -f $script:OOShutUpTaskOwnershipFailureMessage, $taskFullPath) -Level ERROR
                return $false
            }

            if ($isApprovedLegacyDefinition -and -not $isCurrentDefinition) {
                Write-Log ("Migrating approved legacy scheduled task '{0}' after verifying its path, action, principal, and description." -f $taskFullPath) -Level INFO
            }
            else {
                Write-Log ("Verified existing owned scheduled task '{0}' before replacement." -f $taskFullPath) -Level INFO
            }

            try {
                $previousSecurityDescriptor = Get-WindowsUpdateTaskSecurityDescriptor -TaskPath $existingTask.TaskPath -TaskName $existingTask.TaskName
                if (-not (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $previousSecurityDescriptor)) {
                    throw $script:OOShutUpTaskSecurityDescriptorFailureMessage
                }
            }
            catch {
                Write-Log ("{0} Existing task '{1}' was left unchanged. Details: {2}" -f $script:OOShutUpTaskSecurityDescriptorFailureMessage, $taskFullPath, $_.Exception.Message) -Level ERROR
                return $false
            }

            try {
                $previousXml = Export-ScheduledTask -TaskName $existingTask.TaskName -TaskPath $existingTask.TaskPath -ErrorAction Stop
            }
            catch {
                if (Test-ScheduledTaskAccessDenied -ErrorRecord $_) {
                    Write-Log ("Access denied while exporting scheduled task '{0}'. No task was deleted; inspect this path with an account that can read its definition." -f $taskFullPath) -Level ERROR
                }
                else {
                    Write-Log ("Task Scheduler could not export '{0}'. No task was deleted; inspect Task Scheduler and retry. Details: {1}" -f $taskFullPath, $_.Exception.Message) -Level ERROR
                }
                return $false
            }
            if ([string]::IsNullOrWhiteSpace([string]$previousXml)) {
                Write-Log ("Task Scheduler returned an empty definition for '{0}'. No task was deleted; inspect Task Scheduler and retry." -f $taskFullPath) -Level ERROR
                return $false
            }

            $backupIndex = $previousDefinitions.Count + 1
            $backupFileName = '.Install-OOShutUp10-task-rollback-{0}-{1}.xml' -f [guid]::NewGuid().ToString('N'), $backupIndex
            $backupPath = Join-Path ([System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ScriptPath))) $backupFileName
            $descriptorBackupPath = [System.IO.Path]::ChangeExtension($backupPath, '.sddl')
            try {
                [System.IO.File]::WriteAllText($backupPath, [string]$previousXml, [System.Text.Encoding]::UTF8)
                if (-not (Test-SystemTaskTargetAcl -ScriptPath $backupPath)) {
                    throw 'The exported task rollback file did not inherit a verifiable protected ACL.'
                }
                [System.IO.File]::WriteAllText($descriptorBackupPath, [string]$previousSecurityDescriptor, [System.Text.Encoding]::ASCII)
                if (-not (Test-SystemTaskTargetAcl -ScriptPath $descriptorBackupPath)) {
                    throw 'The task security-descriptor rollback file did not inherit a verifiable protected ACL.'
                }
            }
            catch {
                Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $descriptorBackupPath -Force -ErrorAction SilentlyContinue
                throw
            }

            [void]$previousDefinitions.Add([pscustomobject]@{
                TaskName = [string]$existingTask.TaskName
                TaskPath = [string]$existingTask.TaskPath
                Description = [string]$existingTask.Description
                PrincipalUserId = [string]$existingTask.Principal.UserId
                PrincipalRunLevel = [string]$existingTask.Principal.RunLevel
                Actions = @($existingTask.Actions | ForEach-Object {
                    [pscustomobject]@{
                        Execute = [string]$_.Execute
                        Arguments = [string]$_.Arguments
                        WorkingDirectory = [string]$_.WorkingDirectory
                    }
                })
                Xml = [string]$previousXml
                SecurityDescriptor = [string]$previousSecurityDescriptor
                BackupPath = $backupPath
                DescriptorBackupPath = $descriptorBackupPath
            })
        }

        # Register a disabled candidate first. The current task remains untouched until its action, trigger, settings, and principal verify.
        $stagingTaskName = '{0}-validation-{1}' -f $taskName, [guid]::NewGuid().ToString('N')
        $stagingTaskAttempted = $true
        $taskSchedulerOperationStarted = $true
        Register-ScheduledTask -TaskName $stagingTaskName -TaskPath $taskPath -Xml $stagingTaskXml -Force -ErrorAction Stop | Out-Null
        Set-WindowsUpdateTaskSecurityDescriptor -TaskPath $taskPath -TaskName $stagingTaskName -SecurityDescriptor $script:OOShutUpTaskSecurityDescriptorSddl
        [void](Assert-WindowsUpdateTaskSecurityDescriptor -TaskPath $taskPath -TaskName $stagingTaskName)
        $stagedTask = Get-ScheduledTask -TaskName $stagingTaskName -TaskPath $taskPath -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskPath -eq $taskPath -and $_.TaskName -eq $stagingTaskName } |
            Select-Object -First 1
        if ($null -eq $stagedTask -or [string]$stagedTask.State -ne 'Disabled' -or
            -not (Test-WindowsUpdateTaskDefinition -Task $stagedTask -TaskPath $taskPath -TaskName $stagingTaskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath -TrustedPowerShellPath $trustedPowerShellPath)) {
            throw 'The disabled candidate task did not match the approved SYSTEM action, trigger, and queue settings; the existing task was left unchanged.'
        }

        # The verified candidate is ready. Replace prior tasks and roll them back if final verification fails.
        $replacementStarted = $true
        foreach ($previous in $previousDefinitions) {
            Write-Log ("Removing verified owned scheduled task '{0}{1}' before replacement." -f $previous.TaskPath, $previous.TaskName) -Level INFO
            Unregister-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -Confirm:$false -ErrorAction Stop | Out-Null
        }

        Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Xml $finalTaskXml -Force -ErrorAction Stop | Out-Null
        Set-WindowsUpdateTaskSecurityDescriptor -TaskPath $taskPath -TaskName $taskName -SecurityDescriptor $script:OOShutUpTaskSecurityDescriptorSddl
        [void](Assert-WindowsUpdateTaskSecurityDescriptor -TaskPath $taskPath -TaskName $taskName)
        $verifiedTask = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskPath -eq $taskPath -and $_.TaskName -eq $taskName } |
            Select-Object -First 1
        if ($null -eq $verifiedTask -or [string]$verifiedTask.State -ne 'Disabled' -or
            -not (Test-WindowsUpdateTaskDefinition -Task $verifiedTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath -TrustedPowerShellPath $trustedPowerShellPath)) {
            throw 'The disabled replacement task did not match the approved SYSTEM action, success trigger, and queue settings after registration.'
        }

        Set-WindowsUpdateTaskEnabled -TaskPath $taskPath -TaskName $taskName -Enabled $true
        $verifiedTask = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskPath -eq $taskPath -and $_.TaskName -eq $taskName } |
            Select-Object -First 1
        if ($null -eq $verifiedTask -or [string]$verifiedTask.State -eq 'Disabled' -or
            -not (Test-WindowsUpdateTaskDefinition -Task $verifiedTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath -TrustedPowerShellPath $trustedPowerShellPath)) {
            throw 'The replacement task was not enabled with the approved SYSTEM action, success trigger, and queue settings.'
        }
        [void](Assert-WindowsUpdateTaskSecurityDescriptor -TaskPath $taskPath -TaskName $taskName)

        try {
            Unregister-ScheduledTask -TaskName $stagingTaskName -TaskPath $taskPath -Confirm:$false -ErrorAction Stop | Out-Null
            $stagingTaskAttempted = $false
        }
        catch {
            Write-Log 'The replacement task was verified, but its disabled validation task could not be removed. Remove the Install-OOShutUp10-validation-* task from Task Scheduler.' -Level WARNING
        }

        foreach ($previous in $previousDefinitions) {
            try {
                Remove-Item -LiteralPath $previous.BackupPath -Force -ErrorAction Stop
                Remove-Item -LiteralPath $previous.DescriptorBackupPath -Force -ErrorAction Stop
            }
            catch {
                Write-Log "The replacement task was verified, but its protected rollback files remain at $($previous.BackupPath) and $($previous.DescriptorBackupPath). Remove them after confirming the task is healthy." -Level WARNING
            }
        }

        Write-Log "Successfully created and verified scheduled task '$taskFullPath'" -Level SUCCESS
        Write-Log 'Task will run 2 minutes after Windows Update events (Event IDs 19, 43).' -Level INFO
        return $true
    }
    catch {
        $rollbackSucceeded = $true
        $rollbackOutcome = 'preserved'
        if ($replacementStarted) {
            try {
                $registeredTasks = @(Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction Stop)
                if ($registeredTasks.Count -gt 1) {
                    throw 'Multiple tasks matched the exact owned identity during rollback.'
                }

                $currentTargetTask = $null
                if ($registeredTasks.Count -eq 1) {
                    $currentTargetTask = $registeredTasks[0]
                    $matchesPreviousSnapshot = $false
                    foreach ($previous in $previousDefinitions) {
                        if (Test-WindowsUpdateTaskMatchesSnapshot -Task $currentTargetTask -Snapshot $previous) {
                            try {
                                $currentDescriptor = Get-WindowsUpdateTaskSecurityDescriptor -TaskPath $previous.TaskPath -TaskName $previous.TaskName
                                $matchesPreviousSnapshot = (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $currentDescriptor) -and
                                    (Test-WindowsUpdateTaskSecurityDescriptor -SecurityDescriptor $previous.SecurityDescriptor) -and
                                    (Test-WindowsUpdateTaskSecurityDescriptorEqual -ActualSecurityDescriptor $currentDescriptor -ExpectedSecurityDescriptor $previous.SecurityDescriptor)
                            }
                            catch {
                                $matchesPreviousSnapshot = $false
                            }
                            if ($matchesPreviousSnapshot) {
                                break
                            }
                        }
                    }

                    if (-not $matchesPreviousSnapshot) {
                        if (-not (Test-WindowsUpdateTaskDefinition -Task $currentTargetTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath -TrustedPowerShellPath $trustedPowerShellPath)) {
                            throw "The task at '$taskFullPath' no longer matches a verified previous or replacement definition; rollback stopped without deleting it."
                        }

                        Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false -ErrorAction Stop | Out-Null
                        $currentTargetTask = $null
                    }
                }

                if ($null -eq $currentTargetTask) {
                    foreach ($previous in $previousDefinitions) {
                        $restoreDocument = New-Object System.Xml.XmlDocument
                        $restoreDocument.LoadXml($previous.Xml)
                        $restoreEnabledNode = $restoreDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
                        if ($null -eq $restoreEnabledNode) {
                            throw "The previous task definition at '$($previous.TaskPath)$($previous.TaskName)' has no Settings/Enabled element."
                        }
                        $restoreEnabled = [bool]::Parse([string]$restoreEnabledNode.InnerText)
                        $restoreEnabledNode.InnerText = 'false'
                        Register-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -Xml $restoreDocument.OuterXml -Force -ErrorAction Stop | Out-Null
                        Set-WindowsUpdateTaskSecurityDescriptor -TaskPath $previous.TaskPath -TaskName $previous.TaskName -SecurityDescriptor $previous.SecurityDescriptor
                        [void](Assert-WindowsUpdateTaskSecurityDescriptor -TaskPath $previous.TaskPath -TaskName $previous.TaskName -ExpectedSecurityDescriptor $previous.SecurityDescriptor)
                        $restoredTask = Get-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -ErrorAction Stop |
                            Where-Object { $_.TaskPath -eq $previous.TaskPath -and $_.TaskName -eq $previous.TaskName } |
                            Select-Object -First 1
                        if ($null -eq $restoredTask -or [string]$restoredTask.State -ne 'Disabled' -or -not (Test-WindowsUpdateTaskMatchesSnapshot -Task $restoredTask -Snapshot $previous)) {
                            throw "The previous task definition at '$($previous.TaskPath)$($previous.TaskName)' could not be verified after rollback."
                        }
                        if ($restoreEnabled) {
                            Set-WindowsUpdateTaskEnabled -TaskPath $previous.TaskPath -TaskName $previous.TaskName -Enabled $true
                            $restoredTask = Get-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -ErrorAction Stop |
                                Where-Object { $_.TaskPath -eq $previous.TaskPath -and $_.TaskName -eq $previous.TaskName } |
                                Select-Object -First 1
                            if ($null -eq $restoredTask -or [string]$restoredTask.State -eq 'Disabled' -or -not (Test-WindowsUpdateTaskMatchesSnapshot -Task $restoredTask -Snapshot $previous)) {
                                throw "The previous task at '$($previous.TaskPath)$($previous.TaskName)' did not return to its prior enabled state after rollback."
                            }
                        }
                        [void](Assert-WindowsUpdateTaskSecurityDescriptor -TaskPath $previous.TaskPath -TaskName $previous.TaskName -ExpectedSecurityDescriptor $previous.SecurityDescriptor)
                    }
                    if ($previousDefinitions.Count -gt 0) {
                        $rollbackOutcome = 'restored'
                    }
                    else {
                        $rollbackOutcome = 'cleaned'
                    }
                }
            }
            catch {
                $rollbackSucceeded = $false
            }
        }

        if ($stagingTaskAttempted) {
            try {
                Unregister-ScheduledTask -TaskName $stagingTaskName -TaskPath $taskPath -Confirm:$false -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Log "A disabled validation task may remain at $taskPath$stagingTaskName; remove it from Task Scheduler after reviewing the approved target." -Level WARNING
            }
        }

        if (-not ($replacementStarted -and -not $rollbackSucceeded)) {
            foreach ($previous in $previousDefinitions) {
                try {
                    Remove-Item -LiteralPath $previous.BackupPath -Force -ErrorAction Stop
                    Remove-Item -LiteralPath $previous.DescriptorBackupPath -Force -ErrorAction Stop
                }
                catch {
                    Write-Log "Temporary task rollback files remain at $($previous.BackupPath) and $($previous.DescriptorBackupPath); remove them after confirming the previous task state." -Level WARNING
                }
            }
        }

        $failureKind = if (Test-ScheduledTaskAccessDenied -ErrorRecord $_) {
            'Access denied by Task Scheduler'
        }
        elseif ($taskSchedulerOperationStarted) {
            'Task Scheduler operation failed'
        }
        else {
            'Scheduled-task setup failed'
        }

        if ($replacementStarted -and $rollbackSucceeded -and $previousDefinitions.Count -gt 0) {
            Write-Log ("{0}. Scheduled-task replacement failed. The previous task definition was {1}; retry after reviewing the approved target. Details: {2}" -f $failureKind, $rollbackOutcome, $_.Exception.Message) -Level ERROR
        }
        elseif ($replacementStarted -and -not $rollbackSucceeded) {
            if ($previousDefinitions.Count -gt 0) {
                $rollbackFiles = [string]::Join('; ', [string[]]@($previousDefinitions | ForEach-Object { '{0}{1} => XML {2}; SDDL {3}' -f $_.TaskPath, $_.TaskName, $_.BackupPath, $_.DescriptorBackupPath }))
                Write-Log ("{0} and automatic rollback could not be verified. Restore the previous task from its protected XML and SDDL rollback files; register it disabled, apply the saved descriptor with RegisteredTask.SetSecurityDescriptor, verify it, then restore its prior enabled state. Backups: {1}. Details: {2}" -f $failureKind, $rollbackFiles, $_.Exception.Message) -Level ERROR
            }
            else {
                Write-Log ("{0} and cleanup of a partial task could not be verified. Inspect Task Scheduler before retrying. Details: {1}" -f $failureKind, $_.Exception.Message) -Level ERROR
            }
        }
        elseif ($taskSchedulerOperationStarted) {
            Write-Log ("{0}. No existing task was deleted; inspect Task Scheduler and retry. Details: {1}" -f $failureKind, $_.Exception.Message) -Level ERROR
        }
        else {
            Write-Log "Error creating scheduled task: $($_.Exception.Message)" -Level ERROR
        }
        return $false
    }
}

#endregion

function Invoke-OOShutUpTaskSetup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileSha256,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedProfileManifestSha256
    )

    $taskVerified = $false
    $exceptionMessage = $null
    try {
        $taskResult = New-WindowsUpdateScheduledTask `
            -ScriptPath $ScriptPath `
            -ApprovedProfileSha256 $ApprovedProfileSha256 `
            -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256
        $taskVerified = [object]::Equals($taskResult, [bool]$true)
    }
    catch {
        $exceptionMessage = $_.Exception.Message
    }

    return [pscustomobject]@{
        Verified = $taskVerified
        ExceptionMessage = $exceptionMessage
    }
}

function Get-OOShutUpInstallOutcome {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$SettingsVerified,

        [Parameter(Mandatory = $true)]
        [bool]$RestorePointVerified,

        [Parameter(Mandatory = $true)]
        [bool]$RestorePointOverrideUsed,

        [Parameter(Mandatory = $true)]
        [psobject]$TaskSetup
    )

    $taskVerified = [object]::Equals($TaskSetup.Verified, [bool]$true)
    $settingsOutcome = if ($SettingsVerified) { 'Verified' } else { 'Not verified' }
    $restorePointOutcome = if ($RestorePointVerified) {
        'Verified'
    }
    elseif ($RestorePointOverrideUsed) {
        'Explicitly overridden'
    }
    else {
        'Not verified'
    }
    $taskOutcome = if ($taskVerified) { 'Verified' } else { 'Not verified' }
    $status = 'Complete'
    $exitCode = 0
    $summary = 'The approved privacy settings are verified and the update task \myTech.Today\Install-OOShutUp10 is registered and verified.'
    $recoveryMessage = $null

    if (-not $SettingsVerified) {
        $status = 'Failed'
        $exitCode = 1
        $summary = 'Setup failed because the approved privacy settings were not verified. Review the preceding error and retry after resolving it.'
    }
    elseif (-not $taskVerified) {
        $status = 'Partial'
        $exitCode = 2
        $summary = $script:OOShutUpPartialSetupMessage
        $recoveryMessage = $script:OOShutUpTaskRecoveryMessage
    }
    elseif (-not $RestorePointVerified -and -not $RestorePointOverrideUsed) {
        $status = 'Failed'
        $exitCode = 1
        $summary = 'Setup failed because the required restore point was not verified.'
        $recoveryMessage = $script:OOShutUpRestorePointFailureMessage
    }

    return [pscustomobject]@{
        Status = $status
        ExitCode = $exitCode
        SettingsOutcome = $settingsOutcome
        RestorePointOutcome = $restorePointOutcome
        TaskOutcome = $taskOutcome
        TaskVerified = $taskVerified
        TaskExceptionMessage = $TaskSetup.ExceptionMessage
        Summary = $summary
        RecoveryMessage = $recoveryMessage
    }
}

function Write-OOShutUpInstallSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Outcome
    )

    if ($Outcome.Status -eq 'Complete') {
        Write-Log '=== Configuration Complete ===' -Level SUCCESS
        Write-Log $Outcome.Summary -Level SUCCESS
    }
    elseif ($Outcome.Status -eq 'Partial') {
        Write-Log '=== Setup Partial ===' -Level WARNING
        Write-Log $Outcome.Summary -Level WARNING
        Write-Log $Outcome.RecoveryMessage -Level WARNING
    }
    else {
        Write-Log '=== Setup Failed ===' -Level ERROR
        Write-Log $Outcome.Summary -Level ERROR
        if (-not [string]::IsNullOrWhiteSpace([string]$Outcome.RecoveryMessage)) {
            Write-Log $Outcome.RecoveryMessage -Level ERROR
        }
    }
}

#region Main Script

try {
    # Logging must be usable before relocation or any configuration work begins.
    if (-not (Initialize-LogFile)) {
        exit 1
    }

    # Treat redirected or unavailable console input as unattended before profile approval can prompt.
    if (-not $NonInteractive -and -not (Test-OOShutUpInteractiveInput)) {
        $NonInteractive = $true
        Write-Log 'Console input is redirected or unavailable. Using the non-interactive approval path; exact approved profile and manifest SHA-256 digests are required.' -Level WARNING
    }

    # Validate the reviewed profile before relocation, registry changes, restore points, task changes, or O&O launch.
    $approvedProfile = Get-OOShutUpApprovedProfile -RequireMachineScope:$ReapplyOnly
    if ($null -eq $approvedProfile) {
        exit 1
    }
    Write-Log ("Validated approved settings profile version {0} (profile SHA-256 {1}; manifest SHA-256 {2}; {3} settings). Setting values are omitted from logs." -f $approvedProfile.Version, $approvedProfile.Sha256, $approvedProfile.ManifestSha256, $approvedProfile.SettingCount) -Level INFO

    $profileApprovalAccepted = Confirm-OOShutUpProfile `
        -Profile $approvedProfile `
        -ApprovedProfileSha256 $ApprovedProfileSha256 `
        -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 `
        -NonInteractive:$NonInteractive `
        -ProfileApprovalConfirmed:$ProfileApprovalConfirmed
    if (-not $profileApprovalAccepted) {
        exit 1
    }

    $currentPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($currentPath)) {
        $currentPath = $MyInvocation.MyCommand.Path
    }
    $relocationTarget = Get-ScriptRelocationPackageTarget `
        -CurrentScriptPath $currentPath `
        -ProfileDirectory $script:OOShutUpProfileDirectory `
        -ScriptFileName $ScriptFileName `
        -PackageContainer $script:DefaultScriptPackageContainer `
        -ApprovedProfileSha256 $script:OOShutUpApprovedDigest `
        -ApprovedProfileManifestSha256 $script:OOShutUpApprovedManifestDigest
    $DefaultScriptRoot = $relocationTarget.Root
    $DefaultScriptPath = $relocationTarget.ScriptPath
    $script:OOShutUpRelocationPackageDigests = $relocationTarget.Digests

    # Step 0: Ensure script is running from default location (unless SkipRelocate is set)
    if (-not $SkipRelocate) {
        Update-ScriptProgress -PercentComplete 5 -Status "Initializing..." -CurrentOperation "Checking script location"

        if (-not [string]::Equals(
            [System.IO.Path]::GetFullPath($currentPath),
            [System.IO.Path]::GetFullPath($DefaultScriptPath),
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            Invoke-ScriptRelocation -CurrentScriptPath $currentPath
            # If we reach here, relocation failed, but continue anyway
        }
    }

    # Determine if we're running from the default location
    $currentScriptPath = $PSCommandPath
    if ([string]::IsNullOrEmpty($currentScriptPath)) {
        $currentScriptPath = $MyInvocation.MyCommand.Path
    }
    $isRunningFromDefaultLocation = ($currentScriptPath -eq $DefaultScriptPath)

    # Display myTech.Today banner only when running from default location
    if ($isRunningFromDefaultLocation) {
        Show-MyTechTodayBanner
    }

    Update-ScriptProgress -PercentComplete 10 -Status "Initializing..." -CurrentOperation "Preparing configuration"

    # If ReapplyOnly mode, just reapply settings and exit
    if ($ReapplyOnly) {
        Write-Log "=== O&O ShutUp10++ Post-Windows Update Reapplication ===" -Level INFO
        Write-Log " " -Level INFO
        Write-Log "Running in ReapplyOnly mode (triggered by Windows Update)" -Level INFO

        # Get executable path
        $exePath = Get-OOShutUpPath -ExpectedPath $ApprovedOOSExecutablePath
        if (-not $exePath) {
            Write-Log $script:OOShutUpManagedPathFailureMessage -Level ERROR
            exit 1
        }

        Write-Log "Using O&O ShutUp10++ at: $exePath" -Level INFO

        # Apply the validated profile
        Write-Log "Reapplying the approved privacy settings profile..." -Level INFO
        $configSuccess = Invoke-OOShutUpConfiguration -ExecutablePath $exePath -SystemReapply

        if ($configSuccess) {
            if ($script:OOShutUpLastRunChanged) {
                Write-Log "Successfully reapplied privacy settings after Windows Update" -Level SUCCESS
            }
            else {
                Write-Log 'The approved profile already matched the current settings; no setting changes were made.' -Level INFO
            }
        }
        else {
            Write-Log "Failed to reapply privacy settings" -Level ERROR
            exit 1
        }

        if ($script:OOShutUpLastRunChanged) {
            Write-Log "Privacy settings have been restored after Windows Update" -Level SUCCESS
        }
        else {
            Write-Log 'No settings needed restoration after Windows Update.' -Level INFO
        }
        exit 0
    }

    # Normal installation mode
    Write-Log "=== O&O ShutUp10++ Automated Installation and Configuration ===" -Level INFO
    Write-Log " " -Level INFO
    Write-Log "Script Location: $DefaultScriptPath" -Level INFO
    Write-Log "Log Location: console output (packaged local logger)" -Level INFO
    Write-Log " " -Level INFO

    # Step 1: Check if already installed
    Update-ScriptProgress -PercentComplete 15 -Status "Step 1 of 5: Checking Installation" -CurrentOperation "Checking if O&O ShutUp10++ is already installed"
    Write-Log "Step 1: Checking if O&O ShutUp10++ is already installed..." -Level INFO
    $isInstalled = Test-OOShutUpInstalled

    if (-not $isInstalled) {
        Write-Log "O&O ShutUp10++ not found. Proceeding with download..." -Level INFO

        # Step 2: Download the application
        Update-ScriptProgress -PercentComplete 20 -Status "Step 1 of 5: Installing Application" -CurrentOperation "Downloading O&O ShutUp10++"
        Write-Log "Step 2: Downloading O&O ShutUp10++..." -Level INFO
        $downloadSuccess = Install-OOShutUp

        if (-not $downloadSuccess) {
            Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
            Write-Log "Failed to download O&O ShutUp10++. Exiting." -Level ERROR
            exit 1
        }
    }
    else {
        Write-Log "O&O ShutUp10++ is already installed. Skipping download." -Level SUCCESS
    }

    # Step 3: Get executable path
    Update-ScriptProgress -PercentComplete 25 -Status "Step 1 of 5: Verifying Installation" -CurrentOperation "Locating O&O ShutUp10++ executable"
    $exePath = Get-OOShutUpPath
    if (-not $exePath) {
        Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
        Write-Log "Could not locate O&O ShutUp10++ executable. Exiting." -Level ERROR
        exit 1
    }

    Write-Log "Using O&O ShutUp10++ at: $exePath" -Level INFO

    # Step 2: Prepare System Restore. The temporary frequency value is applied and restored inside New-SystemRestorePoint.
    Update-ScriptProgress -PercentComplete 30 -Status "Step 2 of 5: Preparing System Restore" -CurrentOperation "Detecting the Windows volume and checking restore-point readiness"
    Write-Log "Step 2: Preparing Windows System Restore..." -Level INFO

    # Step 3: Create restore point BEFORE applying settings
    Update-ScriptProgress -PercentComplete 40 -Status "Step 3 of 5: Creating Restore Point" -CurrentOperation "Creating system restore point (this may take a few minutes)"
    Write-Log "Step 3: Creating system restore point..." -Level INFO
    $configurationResult = Invoke-OOShutUpConfigurationWithRestorePointGate `
        -ExecutablePath $exePath `
        -AllowWithoutRestorePoint:$AllowWithoutRestorePoint
    $restorePointVerified = $configurationResult.RestorePointVerified
    $restorePointOverrideUsed = $configurationResult.OverrideUsed

    if (-not $configurationResult.ConfigurationStarted) {
        Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
        exit 1
    }

    $configSuccess = $configurationResult.ConfigurationSucceeded

    if (-not $configSuccess) {
        Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
        Write-Log "Failed to apply privacy settings. Please run O&O ShutUp10++ manually." -Level ERROR
        exit 1
    }

    # Step 5: Create scheduled task for post-Windows Update execution
    Update-ScriptProgress -PercentComplete 85 -Status "Step 5 of 5: Creating Scheduled Task" -CurrentOperation "Setting up automatic reapplication after Windows updates"
    Write-Log "Step 5: Creating scheduled task for post-Windows Update execution..." -Level INFO
    # Bind the task to the same running package that passed profile validation, including when relocation failed.
    $taskSetup = Invoke-OOShutUpTaskSetup `
        -ScriptPath $currentScriptPath `
        -ApprovedProfileSha256 $approvedProfile.Sha256 `
        -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256
    $installerOutcome = Get-OOShutUpInstallOutcome `
        -SettingsVerified $configSuccess `
        -RestorePointVerified $restorePointVerified `
        -RestorePointOverrideUsed $restorePointOverrideUsed `
        -TaskSetup $taskSetup
    $taskCreated = $installerOutcome.TaskVerified
    if ($taskCreated) {
        Write-Log "Scheduled task created successfully" -Level SUCCESS
        Write-Log "Privacy settings will be automatically reapplied after Windows updates" -Level INFO
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$installerOutcome.TaskExceptionMessage)) {
        Write-Log ("Unexpected scheduled-task setup exception; the task state is unverified. Details: {0}" -f $installerOutcome.TaskExceptionMessage) -Level ERROR
    }

    # Step 6: Completion
    Write-Log " " -Level INFO
    if ($installerOutcome.Status -eq 'Complete') {
        Update-ScriptProgress -PercentComplete 100 -Status "Complete!" -CurrentOperation "Settings and scheduled task verified"
    }
    elseif ($installerOutcome.Status -eq 'Partial') {
        Update-ScriptProgress -PercentComplete 100 -Status "Partial setup" -CurrentOperation "Settings applied; update task not verified"
    }
    else {
        Update-ScriptProgress -PercentComplete 100 -Status "Setup failed" -CurrentOperation "A required outcome was not verified"
    }
    Write-OOShutUpInstallSummary -Outcome $installerOutcome
    Write-Log " " -Level INFO
    if ($script:OOShutUpLastRunChanged) {
        Write-Log "O&O ShutUp10++ applied the approved privacy settings profile." -Level SUCCESS
    }
    elseif ($configSuccess) {
        Write-Log 'O&O reported that no setting changes were needed for the approved profile.' -Level INFO
    }
    else {
        Write-Log 'The O&O configuration did not complete successfully. Review the preceding error and O&O Undo History before retrying.' -Level WARNING
    }
    if ($restorePointVerified) {
        Write-Log "A system restore point was verified before applying settings; it is separate from the exported O&O profile." -Level SUCCESS
    }
    elseif ($restorePointOverrideUsed) {
        Write-Log "System Restore recovery is unavailable because -AllowWithoutRestorePoint was used. Review O&O Undo History; the approved profile export remains separate." -Level WARNING
    }
    else {
        Write-Log $script:OOShutUpRestorePointFailureMessage -Level ERROR
    }
    Write-Log " " -Level INFO

    # Clear the progress bar
    Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed

    # Display marketing footer only when running from default location
    if ($isRunningFromDefaultLocation) {
        $footerLines = @(
            "",
            "+==============================================================================+",
            "|                                                                              |",
            "|                   Thank You for Using myTech.Today Scripts!                 |",
            "|                                                                              |",
            "+==============================================================================+",
            "|                                                                              |",
            "|  Need Professional IT Support or Custom Automation?                         |",
            "|                                                                              |",
            "|  myTech.Today offers comprehensive IT services including:                   |",
            "|    - Custom PowerShell automation and scripting solutions                   |",
            "|    - Infrastructure optimization and management                             |",
            "|    - Cloud migration and integration services                               |",
            "|    - Managed IT services (MSP) for businesses                               |",
            "|    - 24/7 IT support and consulting                                         |",
            "|                                                                              |",
            "|  Why Choose myTech.Today?                                                   |",
            "|    - 20+ years of IT experience                                             |",
            "|    - 190+ satisfied clients                                                 |",
            "|    - 5-star customer rating                                                 |",
            "|    - Fast 24-hour solution delivery                                         |",
            "|    - Serving Chicagoland and remote clients nationwide                      |",
            "|                                                                              |",
            "|  Get a Free Consultation:                                                   |",
            "|    Call: (847) 767-4914                                                     |",
            "|    Email: sales@mytech.today                                                |",
            "|    Visit: https://mytech.today                                              |",
            "|                                                                              |",
            "|  Follow Us:                                                                  |",
            "|    GitHub: https://github.com/mytech-today-now                              |",
            "|    LinkedIn: https://linkedin.com/in/kylerode                               |",
            "|                                                                              |",
            "+==============================================================================+",
            ""
        )

        foreach ($line in $footerLines) {
            Write-Host $line -ForegroundColor Green
        }
    }

    # Request a restart only after O&O confirms that it changed settings.
    if (-not $script:OOShutUpLastRunChanged) {
        Write-Log 'No setting changes were made; a restart is not required.' -Level INFO
    }
    else {
        Invoke-OOShutUpRestart -NonInteractive:$NonInteractive -ForceRestart:$ForceRestart
    }

    # Exit 0 is reserved for verified settings, an acceptable restore-point outcome, and a verified update task.
    exit $installerOutcome.ExitCode
}
catch {
    Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
    Write-Log "An unexpected error occurred: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    # Keep the last approved portable executable in place for subsequent runs and recovery.
}

#endregion

