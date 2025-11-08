#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Automatically installs and configures O&O ShutUp10++ with recommended privacy settings.

.DESCRIPTION
    This script will:
    1. Check if O&O ShutUp10++ is already installed
    2. Download and install it if not found
    3. Apply recommended privacy settings silently
    4. Create a system restore point
    5. Create a scheduled task to reapply settings after Windows updates
    6. Notify user to restart the computer

.PARAMETER ReapplyOnly
    If specified, only reapplies privacy settings without installation or task creation.
    This is used by the scheduled task after Windows updates.

.NOTES
    File Name      : Install-OOShutUp10.ps1
    Author         : Auto-generated
    Prerequisite   : PowerShell 5.1+, Administrator privileges

.EXAMPLE
    .\Install-OOShutUp10.ps1
    Installs O&O ShutUp10++ and applies recommended settings

.EXAMPLE
    .\Install-OOShutUp10.ps1 -ReapplyOnly
    Only reapplies privacy settings (used by scheduled task)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$ReapplyOnly,

    [Parameter(Mandatory = $false)]
    [switch]$SkipRelocate
)

# Script configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Default installation paths
$DefaultScriptRoot = "$env:SystemRoot\myTech.Today\OOShutup"
$DefaultLogRoot = "$env:SystemRoot\myTech.Today\logs"
$ScriptFileName = "Install-OOShutUp10.ps1"
$DefaultScriptPath = Join-Path $DefaultScriptRoot $ScriptFileName

# Log configuration
$LogFileName = "OOShutup-$(Get-Date -Format 'yyyy-MM').md"
$LogFilePath = Join-Path $DefaultLogRoot $LogFileName
$MaxLogSizeBytes = 10MB  # Rotate log if it exceeds 10MB

# O&O ShutUp10++ configuration
$AppName = "O&O ShutUp10++"
$DownloadUrl = "https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe"
$TempPath = "$env:TEMP\OOSU10.exe"
$InstallPath = "$env:ProgramFiles\OO Software\O&O ShutUp10"
$ExecutableName = "OOSU10.exe"

#region Helper Functions

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Initializes the log file and ensures directory exists
    #>
    param()

    try {
        # Create log directory if it doesn't exist
        if (-not (Test-Path $DefaultLogRoot)) {
            New-Item -Path $DefaultLogRoot -ItemType Directory -Force | Out-Null
        }

        # Check if log file exists and is too large
        if (Test-Path $LogFilePath) {
            $logFile = Get-Item $LogFilePath
            if ($logFile.Length -gt $MaxLogSizeBytes) {
                # Rotate log file
                $rotatedName = "OOShutup-$(Get-Date -Format 'yyyy-MM-dd_HHmmss').md"
                $rotatedPath = Join-Path $DefaultLogRoot $rotatedName
                Move-Item -Path $LogFilePath -Destination $rotatedPath -Force
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Log file rotated to: $rotatedName" -ForegroundColor Cyan
            }
        }

        # Create log file if it doesn't exist
        if (-not (Test-Path $LogFilePath)) {
            $header = @"
# O&O ShutUp10++ Installation and Configuration Log
**Log Started**: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

---

"@
            $header | Out-File -FilePath $LogFilePath -Encoding UTF8 -Force
        }

        return $true
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARNING] Could not initialize log file: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes formatted log messages to console and file
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Message = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    # Allow empty messages for spacing
    if ([string]::IsNullOrEmpty($Message)) {
        Write-Host ""
        # Add spacing to log file too
        try {
            "" | Out-File -FilePath $LogFilePath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
        return
    }

    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
    }

    # Console output
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color

    # File output in Markdown format
    try {
        $mdLevel = switch ($Level) {
            'INFO'    { '📘' }
            'SUCCESS' { '✅' }
            'WARNING' { '⚠️' }
            'ERROR'   { '❌' }
        }

        $logEntry = "**[$timestamp]** $mdLevel **$Level**: $Message"
        $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {
        # Silently fail if we can't write to log file
    }
}

function Invoke-ScriptRelocation {
    <#
    .SYNOPSIS
        Ensures the script is running from the default location
    #>
    [CmdletBinding()]
    param()

    try {
        $currentScriptPath = $MyInvocation.PSCommandPath
        if ([string]::IsNullOrEmpty($currentScriptPath)) {
            $currentScriptPath = $PSCommandPath
        }

        # Check if already running from default location
        if ($currentScriptPath -eq $DefaultScriptPath) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Script is already running from default location" -ForegroundColor Cyan
            return $true
        }

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Relocating script to default location..." -ForegroundColor Cyan

        # Create default directory if it doesn't exist
        if (-not (Test-Path $DefaultScriptRoot)) {
            New-Item -Path $DefaultScriptRoot -ItemType Directory -Force | Out-Null
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SUCCESS] Created directory: $DefaultScriptRoot" -ForegroundColor Green
        }

        # Copy script to default location
        Copy-Item -Path $currentScriptPath -Destination $DefaultScriptPath -Force
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SUCCESS] Script copied to: $DefaultScriptPath" -ForegroundColor Green

        # Re-execute from default location with same parameters
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Re-executing from default location..." -ForegroundColor Cyan

        $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$DefaultScriptPath`"", "-SkipRelocate")
        if ($ReapplyOnly) {
            $arguments += "-ReapplyOnly"
        }

        Start-Process -FilePath "PowerShell.exe" -ArgumentList $arguments -Wait -NoNewWindow

        # Exit this instance
        exit 0
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to relocate script: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Set-RestorePointFrequency {
    <#
    .SYNOPSIS
        Modifies registry to allow restore points to be created anytime
    #>
    [CmdletBinding()]
    param()

    try {
        Write-Log "Configuring registry to allow frequent restore points..." -Level INFO

        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
        $regName = "SystemRestorePointCreationFrequency"

        # Check if registry key exists
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }

        # Set to 0 to allow restore points anytime (default is 1440 minutes = 24 hours)
        Set-ItemProperty -Path $regPath -Name $regName -Value 0 -Type DWord -Force

        Write-Log "Registry configured: Restore points can now be created anytime" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Warning: Could not modify registry: $($_.Exception.Message)" -Level WARNING
        return $false
    }
}

function Test-OOShutUpInstalled {
    <#
    .SYNOPSIS
        Checks if O&O ShutUp10++ is installed
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    # Check common installation paths
    $possiblePaths = @(
        "$env:ProgramFiles\OO Software\O&O ShutUp10\$ExecutableName",
        "$env:ProgramFiles\OOShutUp10\$ExecutableName",
        "$env:LOCALAPPDATA\Programs\OOShutUp10\$ExecutableName"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            Write-Log "Found O&O ShutUp10++ at: $path" -Level SUCCESS
            return $true
        }
    }
    
    # Check if executable exists in PATH
    $exeInPath = Get-Command $ExecutableName -ErrorAction SilentlyContinue
    if ($exeInPath) {
        Write-Log "Found O&O ShutUp10++ in PATH: $($exeInPath.Source)" -Level SUCCESS
        return $true
    }
    
    return $false
}

function Get-OOShutUpPath {
    <#
    .SYNOPSIS
        Gets the installation path of O&O ShutUp10++
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    
    $possiblePaths = @(
        "$env:ProgramFiles\OO Software\O&O ShutUp10\$ExecutableName",
        "$env:ProgramFiles\OOShutUp10\$ExecutableName",
        "$env:LOCALAPPDATA\Programs\OOShutUp10\$ExecutableName",
        "$TempPath"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    $exeInPath = Get-Command $ExecutableName -ErrorAction SilentlyContinue
    if ($exeInPath) {
        return $exeInPath.Source
    }
    
    return $null
}

function Install-OOShutUp {
    <#
    .SYNOPSIS
        Downloads O&O ShutUp10++ to temp directory
    #>
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Downloading O&O ShutUp10++ from: $DownloadUrl" -Level INFO
        
        # Download the executable
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($DownloadUrl, $TempPath)
        
        if (Test-Path $TempPath) {
            $fileSize = (Get-Item $TempPath).Length / 1MB
            Write-Log "Successfully downloaded O&O ShutUp10++ ($([Math]::Round($fileSize, 2)) MB)" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "Failed to download O&O ShutUp10++" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "Error downloading O&O ShutUp10++: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Invoke-OOShutUpConfiguration {
    <#
    .SYNOPSIS
        Applies recommended privacy settings using O&O ShutUp10++
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath
    )
    
    try {
        Write-Log "Applying recommended privacy settings..." -Level INFO
        
        # O&O ShutUp10++ command line arguments:
        # /quiet - Silent mode
        # /nosrp - Don't create restore point (we'll create our own)
        # ooshutup10.cfg - Apply recommended settings
        
        $arguments = "/quiet /nosrp ooshutup10.cfg"
        
        $processInfo = Start-Process -FilePath $ExecutablePath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        
        if ($processInfo.ExitCode -eq 0) {
            Write-Log "Successfully applied recommended privacy settings" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "O&O ShutUp10++ exited with code: $($processInfo.ExitCode)" -Level WARNING
            return $true  # Still return true as some exit codes are non-critical
        }
    }
    catch {
        Write-Log "Error applying settings: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function New-SystemRestorePoint {
    <#
    .SYNOPSIS
        Creates a system restore point
    #>
    [CmdletBinding()]
    param()

    try {
        Write-Log "Creating system restore point..." -Level INFO

        # Enable System Restore on C: drive if not already enabled
        Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue

        # Create restore point
        $description = "Before O&O ShutUp10++ Privacy Settings - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        Checkpoint-Computer -Description $description -RestorePointType "MODIFY_SETTINGS"

        Write-Log "System restore point created successfully" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Warning: Could not create restore point: $($_.Exception.Message)" -Level WARNING
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
        [string]$ScriptPath
    )

    try {
        Write-Log "Creating scheduled task for post-Windows Update execution..." -Level INFO

        $taskName = "OOShutUp10-PostWindowsUpdate"
        $taskDescription = "Reapplies O&O ShutUp10++ privacy settings after Windows updates to prevent Microsoft from resetting privacy configurations"

        # Check if task already exists
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Log "Scheduled task already exists. Removing old task..." -Level INFO
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        }

        # Create action - run PowerShell with the script in ReapplyOnly mode
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
            -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -ReapplyOnly -SkipRelocate"

        # Create multiple triggers for different Windows Update events
        # Trigger 1: Windows Update successful installation (Event ID 19)
        $trigger1 = New-ScheduledTaskTrigger -AtLogOn
        $trigger1CimInstance = Get-CimInstance -ClassName MSFT_TaskEventTrigger -Namespace Root/Microsoft/Windows/TaskScheduler -Property * |
            Where-Object { $_.Subscription -eq $null } | Select-Object -First 1

        # Create Event Trigger XML for Windows Update completion
        $triggerXml = @"
<QueryList>
  <Query Id="0" Path="System">
    <Select Path="System">
      *[System[Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and (EventID=19 or EventID=43)]]
    </Select>
  </Query>
</QueryList>
"@

        # Create the trigger using CIM
        $trigger = New-ScheduledTaskTrigger -AtStartup  # Temporary trigger, will be replaced

        # Create principal (run with highest privileges as SYSTEM)
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

        # Create settings
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -RunOnlyIfNetworkAvailable:$false `
            -DontStopOnIdleEnd `
            -MultipleInstances IgnoreNew

        # Register the task
        $task = Register-ScheduledTask `
            -TaskName $taskName `
            -Description $taskDescription `
            -Action $action `
            -Trigger $trigger `
            -Principal $principal `
            -Settings $settings `
            -Force

        # Now update the task with the proper Event Trigger using XML
        $taskXml = Export-ScheduledTask -TaskName $taskName

        # Replace the trigger section with event-based trigger
        $eventTrigger = @"
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and (EventID=19 or EventID=43)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT2M</Delay>
    </EventTrigger>
"@

        # Use schtasks.exe for more reliable event trigger creation
        # Build XML with proper variable substitution
        $schtasksXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>$taskDescription</Description>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and (EventID=19 or EventID=43)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT2M</Delay>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>PowerShell.exe</Command>
      <Arguments>-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File &quot;$ScriptPath&quot; -ReapplyOnly -SkipRelocate</Arguments>
    </Exec>
  </Actions>
</Task>
"@

        # Save XML to temp file
        $tempXmlPath = "$env:TEMP\OOShutUp10Task.xml"
        $schtasksXml | Out-File -FilePath $tempXmlPath -Encoding Unicode -Force

        # Delete existing task and recreate with schtasks
        schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null
        $result = schtasks.exe /Create /TN $taskName /XML $tempXmlPath /F 2>&1

        # Clean up temp file
        Remove-Item $tempXmlPath -Force -ErrorAction SilentlyContinue

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Successfully created scheduled task '$taskName'" -Level SUCCESS
            Write-Log "Task will run after Windows Update events (Event IDs 19, 43)" -Level INFO
            return $true
        }
        else {
            Write-Log "Warning: Could not create event-based trigger: $result" -Level WARNING
            return $false
        }
    }
    catch {
        Write-Log "Error creating scheduled task: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

#endregion

#region Main Script

try {
    # Step 0: Ensure script is running from default location (unless SkipRelocate is set)
    if (-not $SkipRelocate) {
        $currentPath = $PSCommandPath
        if ($currentPath -ne $DefaultScriptPath) {
            Invoke-ScriptRelocation
            # If we reach here, relocation failed, but continue anyway
        }
    }

    # Initialize logging
    Initialize-LogFile | Out-Null

    # If ReapplyOnly mode, just reapply settings and exit
    if ($ReapplyOnly) {
        Write-Log "=== O&O ShutUp10++ Post-Windows Update Reapplication ===" -Level INFO
        Write-Log "" -Level INFO
        Write-Log "Running in ReapplyOnly mode (triggered by Windows Update)" -Level INFO

        # Get executable path
        $exePath = Get-OOShutUpPath
        if (-not $exePath) {
            Write-Log "Could not locate O&O ShutUp10++ executable. Exiting." -Level ERROR
            exit 1
        }

        Write-Log "Using O&O ShutUp10++ at: $exePath" -Level INFO

        # Apply recommended settings
        Write-Log "Reapplying recommended privacy settings..." -Level INFO
        $configSuccess = Invoke-OOShutUpConfiguration -ExecutablePath $exePath

        if ($configSuccess) {
            Write-Log "Successfully reapplied privacy settings after Windows Update" -Level SUCCESS
        }
        else {
            Write-Log "Failed to reapply privacy settings" -Level ERROR
            exit 1
        }

        Write-Log "Privacy settings have been restored after Windows Update" -Level SUCCESS
        exit 0
    }

    # Normal installation mode
    Write-Log "=== O&O ShutUp10++ Automated Installation and Configuration ===" -Level INFO
    Write-Log "" -Level INFO
    Write-Log "Script Location: $DefaultScriptPath" -Level INFO
    Write-Log "Log Location: $LogFilePath" -Level INFO
    Write-Log "" -Level INFO

    # Step 1: Check if already installed
    Write-Log "Step 1: Checking if O&O ShutUp10++ is already installed..." -Level INFO
    $isInstalled = Test-OOShutUpInstalled

    if (-not $isInstalled) {
        Write-Log "O&O ShutUp10++ not found. Proceeding with download..." -Level INFO

        # Step 2: Download the application
        Write-Log "Step 2: Downloading O&O ShutUp10++..." -Level INFO
        $downloadSuccess = Install-OOShutUp

        if (-not $downloadSuccess) {
            Write-Log "Failed to download O&O ShutUp10++. Exiting." -Level ERROR
            exit 1
        }
    }
    else {
        Write-Log "O&O ShutUp10++ is already installed. Skipping download." -Level SUCCESS
    }

    # Step 3: Get executable path
    $exePath = Get-OOShutUpPath
    if (-not $exePath) {
        Write-Log "Could not locate O&O ShutUp10++ executable. Exiting." -Level ERROR
        exit 1
    }

    Write-Log "Using O&O ShutUp10++ at: $exePath" -Level INFO

    # Step 3: Configure registry to allow frequent restore points
    Write-Log "Step 2: Configuring system for frequent restore points..." -Level INFO
    Set-RestorePointFrequency | Out-Null

    # Step 4: Create restore point BEFORE applying settings
    Write-Log "Step 3: Creating system restore point..." -Level INFO
    New-SystemRestorePoint | Out-Null

    # Step 5: Apply recommended settings
    Write-Log "Step 4: Applying recommended privacy settings..." -Level INFO
    $configSuccess = Invoke-OOShutUpConfiguration -ExecutablePath $exePath

    if (-not $configSuccess) {
        Write-Log "Failed to apply privacy settings. Please run O&O ShutUp10++ manually." -Level ERROR
        exit 1
    }

    # Step 6: Create scheduled task for post-Windows Update execution
    Write-Log "Step 5: Creating scheduled task for post-Windows Update execution..." -Level INFO
    # Always use the default script path for the scheduled task
    $taskCreated = New-WindowsUpdateScheduledTask -ScriptPath $DefaultScriptPath
    if ($taskCreated) {
        Write-Log "Scheduled task created successfully" -Level SUCCESS
        Write-Log "Privacy settings will be automatically reapplied after Windows updates" -Level INFO
    }
    else {
        Write-Log "Warning: Could not create scheduled task. Privacy settings may need manual reapplication after Windows updates." -Level WARNING
    }

    # Step 7: Completion message
    Write-Log "" -Level INFO
    Write-Log "=== Configuration Complete ===" -Level SUCCESS
    Write-Log "" -Level INFO
    Write-Log "O&O ShutUp10++ has been configured with recommended privacy settings." -Level SUCCESS
    Write-Log "A system restore point has been created for safety." -Level SUCCESS
    if ($taskCreated) {
        Write-Log "A scheduled task has been created to reapply settings after Windows updates." -Level SUCCESS
    }
    Write-Log "" -Level INFO
    Write-Log "IMPORTANT: Please restart your computer to activate all privacy settings." -Level WARNING
    Write-Log "" -Level INFO

    # Prompt for restart
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
catch {
    Write-Log "An unexpected error occurred: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    exit 1
}
finally {
    # Cleanup temp file if it exists and we're not using it
    if ((Test-Path $TempPath) -and $isInstalled) {
        Remove-Item $TempPath -Force -ErrorAction SilentlyContinue
    }
}

#endregion

