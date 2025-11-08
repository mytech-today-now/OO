# O&O ShutUp10++ Automated Installation and Configuration

---

## Developed by myTech.Today

**Professional IT Services - Serving the Midwest**

This PowerShell automation script is developed and maintained by **myTech.Today**, a professional IT consulting and managed services provider with 20+ years of experience delivering customized technology solutions.

### About myTech.Today

**myTech.Today** specializes in providing expert IT consulting, automation, and support services to businesses and individuals across the Midwest. We deliver personalized solutions that improve efficiency, enhance security, and drive business value.

**Our Services:**
- IT Consulting and Support
- Custom PowerShell Automation and Scripting
- Infrastructure Optimization and System Administration
- Cloud Integration (Azure, AWS, Microsoft 365)
- Network Design and Management
- Cybersecurity and Compliance Solutions
- Database Management and Custom Development
- System Setup, Upgrades, and Troubleshooting
- Virus and Spyware Removal
- Hardware Procurement, Installation, and Configuration

**Service Area:**
- Chicagoland, IL (Lake Zurich and surrounding areas)
- Northern Illinois
- Southern Wisconsin
- Northern Indiana
- Southern Michigan

**Contact Information:**
- **Email:** sales@mytech.today
- **Phone:** (847) 767-4914
- **Website:** https://mytech.today
- **GitHub:** [@mytech-today-now](https://github.com/mytech-today-now)
- **LinkedIn:** [linkedin.com/in/kylerode](https://linkedin.com/in/kylerode)

**Why Choose myTech.Today?**
- 20+ years of IT consulting and software development experience
- Personalized service to 190+ satisfied clients
- 5-star review rating with consistent 24-hour solution delivery
- Payment-upon-success terms available
- Expertise in Windows automation, PowerShell scripting, and system optimization

---

## Overview

This PowerShell script automates the installation and configuration of **O&O ShutUp10++**, a privacy tool that helps protect Windows users from Microsoft telemetry and spyware by applying recommended privacy settings.

## Features

### ✅ **Automated Installation**
- Scans for existing O&O ShutUp10++ installation
- Downloads from official O&O Software website if not found
- Silent installation with no user interaction required

### ✅ **Privacy Settings Application**
- Applies recommended privacy settings automatically
- Uses O&O ShutUp10++'s default configuration for optimal privacy
- Runs completely silently without opening the GUI

### ✅ **System Protection**
- Creates a system restore point before applying settings
- Allows easy rollback if needed

### ✅ **Post-Windows Update Protection** 🆕
- **Automatically creates a scheduled task** that runs after Windows updates
- Reapplies privacy settings after major, minor, or patch updates
- Prevents Microsoft from resetting privacy configurations during updates
- Triggers on Windows Update Event IDs 19 and 43 (successful installations)

### ✅ **Comprehensive Logging**
- Color-coded console output with timestamps
- Detailed error messages and stack traces
- Clear status indicators (INFO, SUCCESS, WARNING, ERROR)

## Requirements

- **PowerShell 5.1 or higher**
- **Administrator privileges** (script requires elevation)
- **Windows 10 or Windows 11**
- **Internet connection** (for initial download)

## Usage

### Initial Installation and Setup

Run the script as Administrator:

```powershell
# Using call operator (recommended for paths with special characters)
& "Q:\_kyle\temp_documents\GitHub\PowerShellScripts\OO\Install-OOShutUp10.ps1"

# Or navigate to directory first
cd "Q:\_kyle\temp_documents\GitHub\PowerShellScripts\OO"
.\Install-OOShutUp10.ps1
```

### What Happens During Installation

1. ✅ Checks if O&O ShutUp10++ is already installed
2. ✅ Downloads the application if not found
3. ✅ Creates a system restore point
4. ✅ Applies recommended privacy settings silently
5. ✅ **Creates a scheduled task** to reapply settings after Windows updates
6. ✅ Prompts to restart the computer

### Reapply Settings Only (Used by Scheduled Task)

The script can also run in "ReapplyOnly" mode, which only reapplies privacy settings without installation or task creation:

```powershell
.\Install-OOShutUp10.ps1 -ReapplyOnly
```

This mode is automatically used by the scheduled task after Windows updates.

## Scheduled Task Details

### Task Name
`OOShutUp10-PostWindowsUpdate`

### Task Description
Reapplies O&O ShutUp10++ privacy settings after Windows updates to prevent Microsoft from resetting privacy configurations

### Trigger Events
The task triggers on the following Windows Update events:
- **Event ID 19**: Windows Update successful installation
- **Event ID 43**: Windows Update installation completed

### Execution Details
- **Runs as**: SYSTEM account
- **Privileges**: Highest available
- **Delay**: 2 minutes after Windows Update completion
- **Execution**: Hidden PowerShell window (no user interaction)
- **Command**: `PowerShell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<ScriptPath>" -ReapplyOnly`

### Viewing the Scheduled Task

You can view the scheduled task in Task Scheduler:

1. Open **Task Scheduler** (`taskschd.msc`)
2. Navigate to **Task Scheduler Library**
3. Look for **OOShutUp10-PostWindowsUpdate**

Or use PowerShell:

```powershell
Get-ScheduledTask -TaskName "OOShutUp10-PostWindowsUpdate"
```

### Manually Running the Scheduled Task

To test the scheduled task manually:

```powershell
Start-ScheduledTask -TaskName "OOShutUp10-PostWindowsUpdate"
```

### Removing the Scheduled Task

If you want to remove the scheduled task:

```powershell
Unregister-ScheduledTask -TaskName "OOShutUp10-PostWindowsUpdate" -Confirm:$false
```

## How It Works

### Initial Installation Flow

```
┌─────────────────────────────────────┐
│  Check if O&O ShutUp10++ installed  │
└──────────────┬──────────────────────┘
               │
               ├─ Not Installed ──> Download from official source
               │
               └─ Already Installed ──> Skip download
               │
               ▼
┌─────────────────────────────────────┐
│   Create System Restore Point       │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Apply Recommended Privacy Settings │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Create Scheduled Task for Updates  │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│      Prompt for Restart              │
└─────────────────────────────────────┘
```

### Post-Windows Update Flow

```
┌─────────────────────────────────────┐
│    Windows Update Completes         │
│  (Event ID 19 or 43 triggered)      │
└──────────────┬──────────────────────┘
               │
               ▼ (2 minute delay)
┌─────────────────────────────────────┐
│  Scheduled Task Executes Script     │
│  with -ReapplyOnly parameter        │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Locate O&O ShutUp10++ Executable   │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Reapply Privacy Settings Silently  │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│  Privacy Settings Restored!         │
└─────────────────────────────────────┘
```

## Command Line Arguments

O&O ShutUp10++ is executed with the following arguments:

- `/quiet` - Silent mode (no UI interaction)
- `/nosrp` - Don't create restore point (we create our own)
- `ooshutup10.cfg` - Apply recommended privacy settings

## Troubleshooting

### Script Won't Run - Execution Policy Error

If you get an execution policy error, run PowerShell as Administrator and execute:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Path Contains Special Characters Error

If you get an ampersand error, use the call operator:

```powershell
& "Q:\_kyle\temp_documents\GitHub\PowerShellScripts\OO\Install-OOShutUp10.ps1"
```

### Scheduled Task Not Triggering

1. Verify the task exists:
   ```powershell
   Get-ScheduledTask -TaskName "OOShutUp10-PostWindowsUpdate"
   ```

2. Check task history in Task Scheduler:
   - Open Task Scheduler
   - Find the task
   - Click on the "History" tab

3. Manually test the task:
   ```powershell
   Start-ScheduledTask -TaskName "OOShutUp10-PostWindowsUpdate"
   ```

### O&O ShutUp10++ Not Found After Installation

The script checks multiple installation paths:
- `%ProgramFiles%\OO Software\O&O ShutUp10\OOSU10.exe`
- `%ProgramFiles%\OOShutUp10\OOSU10.exe`
- `%LOCALAPPDATA%\Programs\OOShutUp10\OOSU10.exe`
- `%TEMP%\OOSU10.exe`

If the executable is in a different location, the script will attempt to find it in the PATH.

## Security Considerations

- The script requires **Administrator privileges** to:
  - Install software
  - Create system restore points
  - Create scheduled tasks
  - Modify system privacy settings

- The scheduled task runs as **SYSTEM** account to ensure it can apply privacy settings even when no user is logged in

- All downloads are from the **official O&O Software website**: `https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe`

## Benefits

### Why This Script is Useful

1. **Saves Time**: Automates the entire installation and configuration process
2. **Consistency**: Ensures recommended privacy settings are always applied
3. **Protection**: Creates restore points before making changes
4. **Persistence**: Automatically reapplies settings after Windows updates
5. **No User Interaction**: Runs completely silently in the background
6. **Peace of Mind**: Privacy settings are maintained even after major Windows updates

### Why Reapply After Windows Updates?

Microsoft has been known to reset privacy settings during Windows updates, especially major feature updates. This script ensures that your privacy preferences are automatically restored after any Windows update, maintaining your desired level of privacy protection.

## References

- [O&O ShutUp10++ Official Website](https://www.oo-software.com/en/shutup10)
- [O&O ShutUp10++ Documentation](https://www.oo-software.com/en/shutup10/help)
- [Windows Update Event IDs](https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-logs)

## License

This script is provided as-is for educational and personal use. O&O ShutUp10++ is developed by O&O Software GmbH and is subject to their license terms.

## Author

**Developed by:** myTech.Today
**Author:** Kyle C. Rode
**Copyright:** (c) 2025 myTech.Today. All rights reserved.

This script was developed based on requirements for automated O&O ShutUp10++ installation and configuration with post-Windows Update protection, leveraging 20+ years of IT consulting and PowerShell automation expertise.

## Version History

- **v1.0** - Initial release with basic installation and configuration
- **v2.0** - Added scheduled task for post-Windows Update reapplication
- **v3.0** - Added script relocation, file-based logging, and registry modification for restore points

---

## Need Professional IT Support?

**myTech.Today** is here to help with all your technology needs!

### Our Expertise

With over 20 years of experience in IT consulting and software development, we provide comprehensive technology solutions tailored to your specific needs. From simple troubleshooting to complex infrastructure optimization, we deliver results that drive business value.

### Services We Offer

**IT Consulting & Support:**
- Expert consultation on purchasing decisions and system setup
- Troubleshooting and problem resolution
- System upgrades and optimization
- Virus and spyware removal
- Security solutions and compliance

**Custom Development & Automation:**
- PowerShell scripting and automation
- Custom application development (C#, Python, React/Node)
- Database management and integration
- Workflow optimization
- EDI and ERP systems integration

**Infrastructure & Cloud Services:**
- Infrastructure optimization and system administration
- Cloud integration (Azure, AWS, Microsoft 365)
- Network design and management
- Server administration (Windows & Linux)
- Backup and disaster recovery solutions

**Specialized Services:**
- Active Directory and Group Policy management
- Email server setup and configuration
- Website design and maintenance
- SEO and digital marketing
- Remote support and VPN solutions

### Why Choose myTech.Today?

- **Experienced:** 20+ years delivering technology solutions
- **Trusted:** Serving 190+ satisfied clients with a 5-star rating
- **Responsive:** Consistent 24-hour solution delivery
- **Flexible:** Payment-upon-success terms available
- **Results-Driven:** Proven track record of exceeding client goals

### Service Area

We proudly serve businesses and individuals throughout the Midwest:
- **Chicagoland, IL** - Lake Zurich and surrounding areas
- **Northern Illinois**
- **Southern Wisconsin**
- **Northern Indiana**
- **Southern Michigan**

### Get in Touch

Ready to optimize your technology infrastructure? Contact us today!

**Email:** sales@mytech.today
**Phone:** (847) 767-4914
**Website:** https://mytech.today
**GitHub:** [@mytech-today-now](https://github.com/mytech-today-now)
**LinkedIn:** [linkedin.com/in/kylerode](https://linkedin.com/in/kylerode)

---

**myTech.Today** - *Professional IT Services Delivering Results Since 2016*

