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
    6. Ask for a restart only when O&O reports that settings changed

.PARAMETER ReapplyOnly
    If specified, only reapplies privacy settings without installation or task creation.
    This is used by the scheduled task after Windows updates.

.PARAMETER NonInteractive
    Suppresses prompts only when both approved profile and manifest SHA-256 digests match.

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
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
    Set the execution policy to allow running scripts (run once before first use)

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
$script:OOShutUpUnknownScopeFailureMessage = 'The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.'
$script:OOShutUpCurrentUserScopeFailureMessage = 'The SYSTEM task profile contains current-user settings. Reapplication was skipped; run them for the intended signed-in user or remove them from the SYSTEM task profile.'
$script:OOShutUpApprovalFailureMessage = 'The selected settings profile has no matching approval. No settings were changed; review both file digests in the interactive installer and recreate the task.'
$script:OOShutUpRestorePointFailureMessage = 'A restore point could not be verified. No privacy settings were applied; free space or enable restore protection, then retry.'
$script:OOShutUpRestorePointPreparationFailureMessage = 'Restore-point preparation failed on the detected Windows volume. No settings were applied; restore the prior registry value and fix the provider error.'
$script:OOShutUpRestorePointOverrideMessage = 'A restore point could not be verified. System Restore recovery is unavailable; continuing only because -AllowWithoutRestorePoint was supplied. The approved O&O profile export remains separate.'
$script:OOShutUpRestorePointFrequencyRegistryPath = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
$script:OOShutUpRestorePointFrequencyRegistryName = 'SystemRestorePointCreationFrequency'
$script:OOShutUpTaskAclFailureMessage = 'The SYSTEM task target ACL could not be verified. No task was registered; move the script to a protected path and retry.'
$script:OOShutUpTaskOwnershipFailureMessage = 'Task ownership is ambiguous. No task was deleted; inspect the listed path and retry after confirming ownership.'
$script:OOShutUpManagedPathFailureMessage = 'The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.'
$script:OOShutUpApprovedDigest = $null
$script:OOShutUpApprovedManifestDigest = $null
$script:OOShutUpLastRunChanged = $false
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
        # Check if already running from default location
        if ($CurrentScriptPath -eq $DefaultScriptPath) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Script is already running from default location" -ForegroundColor Cyan
            return $true
        }

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Relocating script to default location..." -ForegroundColor Cyan

        # Create default directory if it doesn't exist
        if (-not (Test-Path $DefaultScriptRoot)) {
            New-Item -Path $DefaultScriptRoot -ItemType Directory -Force | Out-Null
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SUCCESS] Created directory: $DefaultScriptRoot" -ForegroundColor Green
        }

        # Copy script to default location (terminating on failure so we don't print a misleading success line)
        Copy-Item -Path $CurrentScriptPath -Destination $DefaultScriptPath -Force -ErrorAction Stop

        # Verify the copy actually produced an identical file (guards against locked/stale destinations)
        $sourceHash = (Get-FileHash -Path $CurrentScriptPath -Algorithm SHA256 -ErrorAction Stop).Hash
        $destHash = (Get-FileHash -Path $DefaultScriptPath -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($sourceHash -ne $destHash) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARN] Relocated script hash mismatch (destination may be locked or stale). Continuing in-place." -ForegroundColor Yellow
            return $false
        }

        # Keep the reviewed profile package beside the relocated script for interactive and SYSTEM runs.
        $destinationProfileDirectory = Join-Path $DefaultScriptRoot 'profiles'
        if (-not (Test-Path -LiteralPath $destinationProfileDirectory -PathType Container)) {
            New-Item -Path $destinationProfileDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        foreach ($profileFileName in @('ooshutup10.cfg', 'ooshutup10.approval.json')) {
            $sourceProfileFile = Join-Path $script:OOShutUpProfileDirectory $profileFileName
            $destinationProfileFile = Join-Path $destinationProfileDirectory $profileFileName
            Copy-Item -LiteralPath $sourceProfileFile -Destination $destinationProfileFile -Force -ErrorAction Stop
            $sourceProfileHash = (Get-FileHash -LiteralPath $sourceProfileFile -Algorithm SHA256 -ErrorAction Stop).Hash
            $destinationProfileHash = (Get-FileHash -LiteralPath $destinationProfileFile -Algorithm SHA256 -ErrorAction Stop).Hash
            if ($sourceProfileHash -ne $destinationProfileHash) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARN] Relocated settings profile hash mismatch. Continuing in-place." -ForegroundColor Yellow
                return $false
            }
        }

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SUCCESS] Script copied to: $DefaultScriptPath" -ForegroundColor Green

        # Re-execute from default location with same parameters
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Re-executing from default location..." -ForegroundColor Cyan
        Write-Host ""

        # Execute the relocated script directly using the call operator
        $relocateParams = @('-SkipRelocate')
        if ($ReapplyOnly) { $relocateParams += '-ReapplyOnly' }
        if ($NonInteractive) { $relocateParams += '-NonInteractive' }
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

function Invoke-OOShutUpConfiguration {
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

        # O&O ShutUp10++ command line arguments:
        # /quiet - Silent mode
        # /nosrp - Don't create restore point (we'll create our own)
        # Absolute profile path - independent of the caller's working directory
        Write-Log ("Applying approved settings profile version {0} (SHA-256 {1}; {2} settings). Setting values are omitted from logs." -f $approvedProfile.Version, $approvedProfile.Sha256, $approvedProfile.SettingCount) -Level INFO

        $arguments = '/quiet /nosrp "{0}"' -f $approvedProfile.Path
        $profileWorkingDirectory = Split-Path -Parent $approvedProfile.Path
        $processInfo = Start-Process -FilePath $ExecutablePath -ArgumentList $arguments -WorkingDirectory $profileWorkingDirectory -Wait -PassThru -NoNewWindow

        # O&O ShutUp10++ exit codes:
        # 0 = Success (no changes needed)
        # 30 = Success (settings applied successfully) - This is the NORMAL success code
        # Other codes may indicate issues

        if ($processInfo.ExitCode -eq 0 -or $processInfo.ExitCode -eq 30) {
            if ($processInfo.ExitCode -eq 30) {
                $script:OOShutUpLastRunChanged = $true
                Write-Log "Successfully applied the approved privacy settings profile" -Level SUCCESS
                Write-Log "Settings were applied and changes were made (exit code 30 is normal)" -Level INFO
            }
            else {
                Write-Log 'O&O reported that no setting changes were needed for the approved profile.' -Level INFO
            }
            return $true
        }
        else {
            Write-Log "O&O ShutUp10++ exited with unexpected code: $($processInfo.ExitCode)" -Level WARNING
            Write-Log "Settings may not have been applied correctly. Please check manually." -Level WARNING
            return $false
        }
    }
    catch {
        Write-Log "Error applying settings: $($_.Exception.Message)" -Level ERROR
        return $false
    }
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

function New-SystemRestorePoint {
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

function Invoke-OOShutUpConfigurationWithRestorePointGate {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [Parameter(Mandatory = $false)]
        [switch]$AllowWithoutRestorePoint
    )

    $restorePointResult = New-SystemRestorePoint
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
    $configurationSucceeded = Invoke-OOShutUpConfiguration -ExecutablePath $ExecutablePath

    return [pscustomobject]@{
        PreparationSucceeded = $true
        RestorePointVerified = [bool]$restorePointVerified
        OverrideUsed = [bool](-not $restorePointVerified -and $AllowWithoutRestorePoint)
        ConfigurationStarted = $true
        ConfigurationSucceeded = [bool]$configurationSucceeded
    }
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
        [string]$ScriptPath
    )

    try {
        $scriptItem = Get-Item -LiteralPath $ScriptPath -Force -ErrorAction Stop
        if ($scriptItem.PSIsContainer) {
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
                $ruleSid = ConvertTo-SystemTaskSid -Identity $rule.IdentityReference
                $rights = [int]$rule.FileSystemRights
                $ruleType = [string]$rule.AccessControlType
                $writeMask = if ($item.PSIsContainer) { $directoryWriteMask } else { $fileWriteMask }
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
        [string]$ApprovedOOSExecutablePath
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
        if ($actions.Count -ne 1 -or [System.IO.Path]::GetFileName([string]$actions[0].Execute) -notmatch '^(?i:powershell\.exe)$') {
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
        if ($actions.Count -ne 1 -or [System.IO.Path]::GetFileName([string]$actions[0].Execute) -notmatch '^(?i:powershell\.exe)$') {
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

        Write-Log ("Scheduled task principal SID S-1-5-18 passed the machine-scope gate for {0} settings (profile SHA-256 {1}; manifest SHA-256 {2}); setting values are omitted." -f $currentApprovedProfile.SettingCount, $currentApprovedProfile.Sha256, $currentApprovedProfile.ManifestSha256) -Level INFO

        Write-Log "Creating scheduled task for post-Windows Update execution..." -Level INFO

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
        $subscription.InnerText = "<QueryList><Query Id=`"0`" Path=`"System`"><Select Path=`"System`">*[System[Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and (EventID=19 or EventID=43)]]</Select></Query></QueryList>"
        $eventTrigger.AppendChild($subscription) | Out-Null

        $delay = $xml.CreateElement("Delay")
        $delay.InnerText = "PT2M"
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
        $multipleInstances.InnerText = "IgnoreNew"
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
        $command.InnerText = "PowerShell.exe"
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
        $stageDocument = New-Object System.Xml.XmlDocument
        $stageDocument.LoadXml($taskXml)
        $stageEnabled = $stageDocument.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
        if ($null -eq $stageEnabled) {
            throw 'The candidate task XML has no Settings/Enabled element.'
        }
        $stageEnabled.InnerText = 'false'
        $stagingTaskXml = $stageDocument.OuterXml

        # Keep the diagnostic copy for local troubleshooting. Registration uses the in-memory XML.
        $debugXmlPath = Join-Path $env:TEMP 'OOShutUp10Task_Debug.xml'
        [System.IO.File]::WriteAllText($debugXmlPath, $taskXml, [System.Text.Encoding]::Unicode)
        Write-Log "Task XML saved to: $debugXmlPath (for debugging)" -Level INFO

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
            $isCurrentDefinition = Test-WindowsUpdateTaskDefinition -Task $existingTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath
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
            try {
                [System.IO.File]::WriteAllText($backupPath, [string]$previousXml, [System.Text.Encoding]::UTF8)
                if (-not (Test-SystemTaskTargetAcl -ScriptPath $backupPath)) {
                    throw 'The exported task rollback file did not inherit a verifiable protected ACL.'
                }
            }
            catch {
                Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
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
                BackupPath = $backupPath
            })
        }

        # Register a disabled candidate first. The current task remains untouched until its action and principal verify.
        $stagingTaskName = '{0}-validation-{1}' -f $taskName, [guid]::NewGuid().ToString('N')
        $stagingTaskAttempted = $true
        $taskSchedulerOperationStarted = $true
        Register-ScheduledTask -TaskName $stagingTaskName -TaskPath $taskPath -Xml $stagingTaskXml -Force -ErrorAction Stop | Out-Null
        $stagedTask = Get-ScheduledTask -TaskName $stagingTaskName -TaskPath $taskPath -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskPath -eq $taskPath -and $_.TaskName -eq $stagingTaskName } |
            Select-Object -First 1
        if ($null -eq $stagedTask -or [string]$stagedTask.State -ne 'Disabled' -or
            -not (Test-WindowsUpdateTaskDefinition -Task $stagedTask -TaskPath $taskPath -TaskName $stagingTaskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath)) {
            throw 'The disabled candidate task did not match the approved SYSTEM action; the existing task was left unchanged.'
        }

        # The verified candidate is ready. Replace prior tasks and roll them back if final verification fails.
        $replacementStarted = $true
        foreach ($previous in $previousDefinitions) {
            Write-Log ("Removing verified owned scheduled task '{0}{1}' before replacement." -f $previous.TaskPath, $previous.TaskName) -Level INFO
            Unregister-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -Confirm:$false -ErrorAction Stop | Out-Null
        }

        Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Xml $taskXml -Force -ErrorAction Stop | Out-Null
        $verifiedTask = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskPath -eq $taskPath -and $_.TaskName -eq $taskName } |
            Select-Object -First 1
        if ($null -eq $verifiedTask -or
            -not (Test-WindowsUpdateTaskDefinition -Task $verifiedTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath)) {
            throw 'The replacement task did not match the approved SYSTEM action after registration.'
        }

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
            }
            catch {
                Write-Log "The replacement task was verified, but its protected rollback file remains at $($previous.BackupPath). Remove it after confirming the task is healthy." -Level WARNING
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
                            $matchesPreviousSnapshot = $true
                            break
                        }
                    }

                    if (-not $matchesPreviousSnapshot) {
                        if (-not (Test-WindowsUpdateTaskDefinition -Task $currentTargetTask -TaskPath $taskPath -TaskName $taskName -TaskDescription $taskDescription -ScriptPath $ScriptPath -ApprovedProfileSha256 $ApprovedProfileSha256 -ApprovedProfileManifestSha256 $ApprovedProfileManifestSha256 -ApprovedOOSExecutablePath $approvedOOSExecutablePath)) {
                            throw "The task at '$taskFullPath' no longer matches a verified previous or replacement definition; rollback stopped without deleting it."
                        }

                        Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false -ErrorAction Stop | Out-Null
                        $currentTargetTask = $null
                    }
                }

                if ($null -eq $currentTargetTask) {
                    foreach ($previous in $previousDefinitions) {
                        Register-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -Xml $previous.Xml -Force -ErrorAction Stop | Out-Null
                        $restoredTask = Get-ScheduledTask -TaskName $previous.TaskName -TaskPath $previous.TaskPath -ErrorAction Stop |
                            Where-Object { $_.TaskPath -eq $previous.TaskPath -and $_.TaskName -eq $previous.TaskName } |
                            Select-Object -First 1
                        if ($null -eq $restoredTask -or -not (Test-WindowsUpdateTaskMatchesSnapshot -Task $restoredTask -Snapshot $previous)) {
                            throw "The previous task definition at '$($previous.TaskPath)$($previous.TaskName)' could not be verified after rollback."
                        }
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
                }
                catch {
                    Write-Log "A temporary task rollback file remains at $($previous.BackupPath); remove it after confirming the previous task state." -Level WARNING
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
                $rollbackFiles = [string]::Join('; ', [string[]]@($previousDefinitions | ForEach-Object { '{0}{1} => {2}' -f $_.TaskPath, $_.TaskName, $_.BackupPath }))
                Write-Log ("{0} and automatic rollback could not be verified. Restore the previous task definition from its protected rollback file using Register-ScheduledTask -TaskName <name> -TaskPath <path> -Xml (Get-Content -LiteralPath <file> -Raw) -Force. Backups: {1}. Details: {2}" -f $failureKind, $rollbackFiles, $_.Exception.Message) -Level ERROR
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

#region Main Script

try {
    # Logging must be usable before relocation or any configuration work begins.
    if (-not (Initialize-LogFile)) {
        exit 1
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

    # Step 0: Ensure script is running from default location (unless SkipRelocate is set)
    if (-not $SkipRelocate) {
        Update-ScriptProgress -PercentComplete 5 -Status "Initializing..." -CurrentOperation "Checking script location"

        $currentPath = $PSCommandPath
        if ([string]::IsNullOrEmpty($currentPath)) {
            $currentPath = $MyInvocation.MyCommand.Path
        }

        if ($currentPath -ne $DefaultScriptPath) {
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
    $taskCreated = New-WindowsUpdateScheduledTask -ScriptPath $currentScriptPath -ApprovedProfileSha256 $approvedProfile.Sha256 -ApprovedProfileManifestSha256 $approvedProfile.ManifestSha256
    if ($taskCreated) {
        Write-Log "Scheduled task created successfully" -Level SUCCESS
        Write-Log "Privacy settings will be automatically reapplied after Windows updates" -Level INFO
    }
    else {
        Write-Log "Warning: Could not create scheduled task. Privacy settings may need manual reapplication after Windows updates." -Level WARNING
    }

    # Step 6: Completion
    Update-ScriptProgress -PercentComplete 100 -Status "Complete!" -CurrentOperation "Configuration completed successfully"
    Write-Log " " -Level INFO
    Write-Log "=== Configuration Complete ===" -Level SUCCESS
    Write-Log " " -Level INFO
    if ($script:OOShutUpLastRunChanged) {
        Write-Log "O&O ShutUp10++ applied the approved privacy settings profile." -Level SUCCESS
    }
    else {
        Write-Log 'O&O reported that no setting changes were needed for the approved profile.' -Level INFO
    }
    if ($restorePointVerified) {
        Write-Log "A system restore point was verified before applying settings; it is separate from the exported O&O profile." -Level SUCCESS
    }
    elseif ($restorePointOverrideUsed) {
        Write-Log "System Restore recovery is unavailable because -AllowWithoutRestorePoint was used. Review O&O Undo History; the approved profile export remains separate." -Level WARNING
    }
    else {
        Write-Log $script:OOShutUpRestorePointFailureMessage -Level ERROR
        exit 1
    }
    if ($taskCreated) {
        Write-Log "A scheduled task has been created to reapply settings after Windows updates." -Level SUCCESS
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
    elseif ($NonInteractive) {
        Write-Log 'Settings changed. Non-interactive mode skipped the restart prompt; restart manually to activate them.' -Level WARNING
    }
    else {
        Write-Log "IMPORTANT: Please restart your computer to activate the changed privacy settings." -Level WARNING
        $restart = Read-Host "Would you like to restart now? (Y/N)"
        if ($restart -eq 'Y' -or $restart -eq 'y') {
            Write-Log "Restarting computer in 10 seconds..." -Level WARNING
            Start-Sleep -Seconds 10
            Restart-Computer -Force
        }
        else {
            Write-Log "Please remember to restart your computer manually." -Level WARNING
        }
    }

    # Explicit success exit code
    exit 0
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

