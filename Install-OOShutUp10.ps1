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

.PARAMETER SkipRelocate
    If specified, skips the script relocation to the default location.
    Used internally to prevent infinite loops.

.NOTES
    File Name      : Install-OOShutUp10.ps1
    Author         : Kyle C. Rode / myTech.Today
    Company        : myTech.Today
    Version        : 2.0.0
    DateCreated    : 2025-01-08
    LastModified   : 2025-01-08
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
    .\Install-OOShutUp10.ps1
    Installs O&O ShutUp10++ and applies recommended settings

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

**Created by**: myTech.Today
**Author**: Kyle C. Rode
**Website**: https://mytech.today
**Email**: sales@mytech.today
**Phone**: (847) 767-4914
**Location**: Lake Zurich, IL (Chicagoland area)

**Log Started**: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

---

## About myTech.Today

myTech.Today is a professional IT consulting and managed services provider (MSP)
specializing in PowerShell automation, infrastructure optimization, and custom
development solutions. With over 20 years of experience, we deliver technology
solutions that improve efficiency and drive business value.

**Our Services:**
- PowerShell Automation & Scripting
- Infrastructure Optimization
- Custom Application Development
- Cloud Integration (AWS, Azure, Google Cloud)
- Managed IT Services (MSP)
- IT Consulting & Support

**Service Area**: Chicagoland (Lake Zurich, IL) with remote support available nationwide

**Contact Us**: For professional IT services and custom automation solutions,
visit https://mytech.today or call (847) 767-4914

---

## Installation Log

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

        # Copy script to default location
        Copy-Item -Path $CurrentScriptPath -Destination $DefaultScriptPath -Force
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [SUCCESS] Script copied to: $DefaultScriptPath" -ForegroundColor Green

        # Re-execute from default location with same parameters
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Re-executing from default location..." -ForegroundColor Cyan
        Write-Host ""

        # Execute the relocated script directly using the call operator
        if ($ReapplyOnly) {
            & $DefaultScriptPath -SkipRelocate -ReapplyOnly
        }
        else {
            & $DefaultScriptPath -SkipRelocate
        }

        # Exit this instance
        exit $LASTEXITCODE
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

        # O&O ShutUp10++ exit codes:
        # 0 = Success (no changes needed)
        # 30 = Success (settings applied successfully) - This is the NORMAL success code
        # Other codes may indicate issues

        if ($processInfo.ExitCode -eq 0 -or $processInfo.ExitCode -eq 30) {
            Write-Log "Successfully applied recommended privacy settings" -Level SUCCESS
            if ($processInfo.ExitCode -eq 30) {
                Write-Log "Settings were applied and changes were made (exit code 30 is normal)" -Level INFO
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

function New-SystemRestorePoint {
    <#
    .SYNOPSIS
        Creates a system restore point with progress indication
    #>
    [CmdletBinding()]
    param()

    try {
        Write-Log "Creating system restore point..." -Level INFO
        Write-Log "This may take a few minutes. Please wait..." -Level INFO

        # Enable System Restore on C: drive if not already enabled
        Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue

        # Create restore point with progress indication
        $description = "Before O&O ShutUp10++ Privacy Settings - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

        # Start a background job to create the restore point
        $job = Start-Job -ScriptBlock {
            param($desc)
            try {
                Checkpoint-Computer -Description $desc -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
                return @{ Success = $true; Message = "Restore point created successfully" }
            }
            catch {
                return @{ Success = $false; Message = $_.Exception.Message }
            }
        } -ArgumentList $description

        # Show progress while waiting for the job to complete
        $elapsed = 0
        $maxWait = 300 # 5 minutes max

        while ($job.State -eq 'Running' -and $elapsed -lt $maxWait) {
            $seconds = $elapsed % 60
            $minutes = [Math]::Floor($elapsed / 60)

            if ($minutes -gt 0) {
                $timeStr = "$minutes min $seconds sec"
            } else {
                $timeStr = "$seconds seconds"
            }

            Write-Progress -Activity "Creating System Restore Point" `
                          -Status "Elapsed time: $timeStr" `
                          -PercentComplete (($elapsed / $maxWait) * 100) `
                          -CurrentOperation "Please wait while Windows creates a restore point..."

            Start-Sleep -Seconds 1
            $elapsed++
        }

        # Check if job completed or timed out
        if ($job.State -eq 'Running') {
            # Job timed out
            Write-Log "Restore point creation timed out after $maxWait seconds" -Level WARNING
            Stop-Job -Job $job
            Remove-Job -Job $job -Force
            Write-Progress -Activity "Creating System Restore Point" -Completed
            return $false
        }

        # Job completed - get the result
        $result = Receive-Job -Job $job -Wait -AutoRemoveJob

        # Clear the progress bar
        Write-Progress -Activity "Creating System Restore Point" -Completed

        if ($result -and $result.Success) {
            Write-Log "System restore point created successfully" -Level SUCCESS
            return $true
        }
        else {
            $errorMsg = if ($result.Message) { $result.Message } else { "Unknown error" }
            Write-Log "Warning: Could not create restore point: $errorMsg" -Level WARNING
            return $false
        }
    }
    catch {
        Write-Progress -Activity "Creating System Restore Point" -Completed
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

        # Delete existing task if it exists
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Log "Scheduled task already exists. Removing old task..." -Level INFO
            schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null
        }

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
        $arguments.InnerText = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -ReapplyOnly -SkipRelocate"
        $exec.AppendChild($arguments) | Out-Null

        $actions.AppendChild($exec) | Out-Null
        $task.AppendChild($actions) | Out-Null

        # Convert XML to string
        $stringWriter = New-Object System.IO.StringWriter
        $xmlWriter = New-Object System.Xml.XmlTextWriter($stringWriter)
        $xmlWriter.Formatting = "Indented"
        $xml.WriteContentTo($xmlWriter)
        $xmlWriter.Flush()
        $stringWriter.Flush()
        $taskXml = $stringWriter.ToString()

        # Save XML to temp file with UTF-16 encoding (required by schtasks)
        $tempXmlPath = "$env:TEMP\OOShutUp10Task.xml"

        # Also save a copy for debugging
        $debugXmlPath = "$env:TEMP\OOShutUp10Task_Debug.xml"

        # Write the XML file with proper encoding
        [System.IO.File]::WriteAllText($tempXmlPath, $taskXml, [System.Text.Encoding]::Unicode)
        [System.IO.File]::WriteAllText($debugXmlPath, $taskXml, [System.Text.Encoding]::Unicode)

        Write-Log "Task XML saved to: $debugXmlPath (for debugging)" -Level INFO

        # Create the task using schtasks.exe
        Write-Log "Registering scheduled task with Windows Task Scheduler..." -Level INFO
        $result = schtasks.exe /Create /TN $taskName /XML $tempXmlPath /F 2>&1

        # Clean up temp file (but keep debug file)
        Remove-Item $tempXmlPath -Force -ErrorAction SilentlyContinue

        # Check if task creation was successful
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Successfully created scheduled task '$taskName'" -Level SUCCESS
            Write-Log "Task will run 2 minutes after Windows Update events (Event IDs 19, 43)" -Level INFO

            # Verify the task was created
            $verifyTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($verifyTask) {
                Write-Log "Task verification: Scheduled task is registered and ready" -Level SUCCESS
                return $true
            }
            else {
                Write-Log "Warning: Task creation reported success but task not found in scheduler" -Level WARNING
                return $false
            }
        }
        else {
            # Parse the error message
            $errorMsg = $result | Out-String
            Write-Log "Error creating scheduled task: $errorMsg" -Level ERROR
            Write-Log "Exit code: $LASTEXITCODE" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "Error creating scheduled task: $($_.Exception.Message)" -Level ERROR
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
        return $false
    }
}

#endregion

#region Main Script

try {
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

    # Initialize logging
    Update-ScriptProgress -PercentComplete 10 -Status "Initializing..." -CurrentOperation "Setting up logging"
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

    # Step 2: Configure registry to allow frequent restore points
    Update-ScriptProgress -PercentComplete 30 -Status "Step 2 of 5: Configuring System" -CurrentOperation "Configuring registry for frequent restore points"
    Write-Log "Step 2: Configuring system for frequent restore points..." -Level INFO
    Set-RestorePointFrequency | Out-Null

    # Step 3: Create restore point BEFORE applying settings
    Update-ScriptProgress -PercentComplete 40 -Status "Step 3 of 5: Creating Restore Point" -CurrentOperation "Creating system restore point (this may take a few minutes)"
    Write-Log "Step 3: Creating system restore point..." -Level INFO
    New-SystemRestorePoint | Out-Null

    # Step 4: Apply recommended settings
    Update-ScriptProgress -PercentComplete 70 -Status "Step 4 of 5: Applying Privacy Settings" -CurrentOperation "Applying recommended privacy settings"
    Write-Log "Step 4: Applying recommended privacy settings..." -Level INFO
    $configSuccess = Invoke-OOShutUpConfiguration -ExecutablePath $exePath

    if (-not $configSuccess) {
        Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
        Write-Log "Failed to apply privacy settings. Please run O&O ShutUp10++ manually." -Level ERROR
        exit 1
    }

    # Step 5: Create scheduled task for post-Windows Update execution
    Update-ScriptProgress -PercentComplete 85 -Status "Step 5 of 5: Creating Scheduled Task" -CurrentOperation "Setting up automatic reapplication after Windows updates"
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

    # Step 6: Completion
    Update-ScriptProgress -PercentComplete 100 -Status "Complete!" -CurrentOperation "Configuration completed successfully"
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
    Write-Progress -Activity "O&O ShutUp10++ Installation and Configuration" -Completed
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

