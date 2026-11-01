# O&O ShutUp10++ Automated Installation and Configuration

---

## Developed by myTech.Today

**Professional IT Services - Serving the Midwest**

This PowerShell automation script is developed and maintained by **myTech.Today**, a professional IT consulting and managed services provider with 20+ years of experience delivering customized technology solutions.

<div align="center">

<!-- ── Header ── -->
<table>
<tr>
<td colspan="2" align="center">
<h3>🏢 About myTech.Today</h3>
<p><strong>myTech.Today</strong> specializes in providing expert IT consulting, automation, and support services to businesses and individuals across the Midwest. We deliver personalized solutions that improve efficiency, enhance security, and drive business value.</p>
</td>
</tr>
</table>

<!-- ── Info Cards ── -->
<table>
<tr>

<!-- Left Column -->
<td width="50%" valign="top">

<h4>🛠️ Our Services</h4>
<ul>
<li>IT Consulting and Support</li>
<li>Custom PowerShell Automation and Scripting</li>
<li>Infrastructure Optimization and System Administration</li>
<li>Cloud Integration (Azure, AWS, Microsoft 365)</li>
<li>Network Design and Management</li>
<li>Cybersecurity and Compliance Solutions</li>
<li>Database Management and Custom Development</li>
<li>System Setup, Upgrades, and Troubleshooting</li>
<li>Virus and Spyware Removal</li>
<li>Hardware Procurement, Installation, and Configuration</li>
</ul>

<h4>📍 Service Area</h4>
<ul>
<li>Chicagoland, IL</li>
<li>Illinois</li>
<li>Southern Wisconsin</li>
<li>Northern Indiana</li>
<li>Southern Michigan</li>
</ul>

</td>

<!-- Right Column -->
<td width="50%" valign="top">

<h4>⭐ Why Choose myTech.Today?</h4>
<ul>
<li>20+ years of IT consulting and software development experience</li>
<li>Personalized service to 190+ satisfied clients</li>
<li>5-star review rating with consistent 24-hour solution delivery</li>
<li>Payment-upon-success terms available</li>
<li>Expertise in Windows automation, PowerShell scripting, and system optimization</li>
</ul>

<h4>📬 Contact Information</h4>
<table>
<tr><td>📧 <strong>Email</strong></td><td><a href="mailto:sales@mytech.today">sales@mytech.today</a></td></tr>
<tr><td>📞 <strong>Phone</strong></td><td>(847) 767-4914</td></tr>
<tr><td>🌐 <strong>Website</strong></td><td><a href="https://mytech.today">mytech.today</a></td></tr>
<tr><td>💻 <strong>GitHub</strong></td><td><a href="https://github.com/mytech-today-now">@mytech-today-now</a></td></tr>
<tr><td>🔗 <strong>LinkedIn</strong></td><td><a href="https://linkedin.com/in/kylerode">linkedin.com/in/kylerode</a></td></tr>
</table>

</td>

</tr>
</table>

<!-- ── Footer ── -->
<table>
<tr>
<td colspan="2" align="center">
<sub><strong>myTech.Today</strong> · Professional IT Services · Serving the Midwest · <a href="https://mytech.today">mytech.today</a> · (847) 767-4914</sub>
</td>
</tr>
</table>

</div>

---

## Overview

This PowerShell script automates the installation and configuration of **O&O ShutUp10++** using a reviewed settings profile supplied with the script package.

## Features

### ✅ **Automated Installation**
- Scans for existing O&O ShutUp10++ installation
- Downloads from official O&O Software website if not found
- Silent installation after the profile approval gate

### ✅ **Privacy Settings Application**
- Applies only the settings in the approved, digest-verified profile
- Stops before system or privacy changes when the approved profile is missing or invalid
- Shows a keyboard-readable impact summary before interactive approval
- Allows unattended application only when the exact profile SHA-256 digest is supplied and matches
- Runs O&O in quiet mode after the approval gate

### ✅ **System Protection**
- Attempts a system restore point before applying settings and reports the outcome
- Documents O&O's per-session undo history as the primary settings rollback path

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

## Set Execution Policy

Before running the script, set the PowerShell execution policy (one-time setup):

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
```

## Usage

### Initial Installation and Setup

Run the script as Administrator in CMD:

```cmd (elevated)
<drive>:\path\to\script\install.bat
```

Run the script as Administrator:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
cd "<drive>:\path\to\script\" & .\Install-OOShutUp10.ps1
```

### What Happens During Installation

1. ✅ Validates the selected profile and confirms the impact summary
2. ✅ Checks if O&O ShutUp10++ is already installed and downloads it if needed
3. ✅ Attempts a system restore point
4. ✅ Applies the approved profile quietly
5. ✅ **Creates a scheduled task** to reapply settings after Windows updates
6. ✅ Prompts for restart only when O&O reports that settings changed

### Reapply Settings Only (Used by Scheduled Task)

The script can also run in "ReapplyOnly" mode, which only reapplies privacy settings without installation or task creation:

```powershell
.\Install-OOShutUp10.ps1 -ReapplyOnly
```

This mode is automatically used by the scheduled task after Windows updates.

## Scheduled Task Details

### Task Name
`\myTech.Today\Install-OOShutUp10`

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
- **Command**: `PowerShell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<ScriptPath>" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedProfileSha256 "<PinnedProfileSHA256>" -ApprovedProfileManifestSha256 "<PinnedManifestSHA256>"`
- **Working directory**: the directory containing `<ScriptPath>`

### Viewing the Scheduled Task

You can view the scheduled task in Task Scheduler:

1. Open **Task Scheduler** (`taskschd.msc`)
2. Navigate to **Task Scheduler Library**
3. Open the **myTech.Today** folder and look for **Install-OOShutUp10**

Or use PowerShell:

```powershell
Get-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
```

### Manually Running the Scheduled Task

To test the scheduled task manually:

```powershell
Start-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
```

### Removing the Scheduled Task

If you want to remove the scheduled task:

```powershell
Unregister-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\" -Confirm:$false
```

## How It Works

### Initial Installation Flow

1. Validate the profile and show its impact summary. A rejection or cancellation stops before relocation or system changes.
2. Check for O&O ShutUp10++ and download it from the official source if needed.
3. Attempt a system restore point.
4. Apply the approved profile quietly.
5. Create the SYSTEM task only when every profile entry is reviewed as machine-scoped; pin both profile and manifest digests.
6. Prompt for a restart only when O&O reports that settings changed.

### Post-Windows Update Flow

1. Windows Update event 19 or 43 starts the task after a two-minute delay.
2. The task runs `-ReapplyOnly -NonInteractive` with its pinned profile and manifest digests.
3. A digest mismatch, unknown scope, or current-user entry stops execution before O&O launches.
4. Only an entirely machine-scoped, digest-matched profile can be reapplied by LocalSystem; the script reports whether settings changed.

## Command Line Arguments

O&O ShutUp10++ is executed with the following arguments:

- `/quiet` - Silent mode (no UI interaction)
- `/nosrp` - Don't create restore point (we create our own)
- The absolute path to the approved `ooshutup10.cfg` profile

Before relocation, registry changes, restore-point creation, scheduled-task changes, or O&O launch, interactive runs display the profile and manifest SHA-256 digests, setting IDs and states, reviewer-provided categories and impact descriptions, and each setting's scope and scope evidence. Type `YES` to accept, `NO` to reject, or `CANCEL` to stop. The summary uses the reviewed manifest text and does not infer setting names, categories, or scope from IDs.

For an approved unattended deployment, pass the exact `.cfg` and manifest digests explicitly:

```powershell
.\Install-OOShutUp10.ps1 -NonInteractive -ApprovedProfileSha256 'REPLACE_WITH_REVIEWED_CFG_SHA256' -ApprovedProfileManifestSha256 'REPLACE_WITH_REVIEWED_MANIFEST_SHA256'
```

Omitting either digest or using a stale digest fails closed with:

> The selected settings profile has no matching approval. No settings were changed; review both file digests in the interactive installer and recreate the task.

Unknown scope fails closed with:

> The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.

The scheduled task pins both exact file digests. If either file changes after approval, the task stops without launching O&O until its approval is reviewed and updated.

### Required reviewed settings profile

Place these two versioned files in the script package's `profiles` directory:

- `profiles/ooshutup10.cfg` - O&O ShutUp10 Free Edition settings export
- `profiles/ooshutup10.approval.json` - Reviewed profile version, SHA-256 digest, exact setting IDs, impact metadata, and scope evidence

The schema version 3 manifest must be JSON with exactly these properties: `schemaVersion` set to `3`, a non-empty `profileVersion`, the lowercase or uppercase 64-character `sha256` digest of the `.cfg` file bytes, a `settingIds` array listing every profile ID exactly once, and a `settingDetails` array with one entry per ID. Each detail has exactly `id`, `state`, `category`, `impact`, `scope`, and `scopeEvidence` properties. `state` must match the profile's `+` or `-`; `category`, `impact`, and `scopeEvidence` must contain reviewed plain text; `scope` must be `machine` or `current-user`. Do not guess setting names, categories, impacts, or scope from an ID. The validator requires the IDs and states to match exactly and rejects absent, empty, malformed, duplicate, unknown, or digest-mismatched data.

Export the `.cfg` from O&O ShutUp10, review every setting ID and state, and check each setting's current description, availability, and scope in O&O. O&O documents Current User settings as HKCU and Local Machine settings as HKLM; some IDs have separate variants in both scopes. Verify each entry in a disposable Windows VM with two accounts and assertions that only the intended hive changes. A SYSTEM task is created only if every entry is machine-scoped; current-user entries block that task and must be applied under their intended user or removed from the SYSTEM profile. The manifest is not a digital signature; interactive approval is bound to the displayed `.cfg` and manifest digests, and unattended approval must supply both exact digests separately. A `-` state is an explicit setting choice; it does not restore the value that existed before the setting was changed. See [O&O's profile file structure documentation](https://manuals.oo-software.com/ooshutup10/en/docs/common-features/profile-structure/), [O&O's scope documentation](https://manuals.oo-software.com/ooshutup10/en/docs/features/overview/), and [the package notes](profiles/README.md).

This checkout currently contains no approved profile, so it exits with this message until both reviewed files are supplied:

> The approved settings profile is missing or invalid. No privacy settings were applied; place the reviewed profile in the configured directory.

The profile is resolved beneath the script's package directory regardless of the caller's current directory. Relocation copies the reviewed profile package beside the relocated script and preserves both accepted digests. O&O receives the absolute profile path and starts with the profile directory as its working directory. The scheduled task runs PowerShell with the script directory as its working directory, retains `-ReapplyOnly`, and supplies the exact absolute managed O&O executable path and both pinned profile digests. O&O exit code `0` is reported as no setting changes and does not prompt for a restart; exit code `30` is treated as settings changed.

### Rollback

Before applying the profile, the script attempts to create a Windows restore point and reports whether that attempt succeeded. In O&O ShutUp10, use **Actions → Undo → History** to review and undo the session that applied the profile. O&O keeps a stack of recent sessions, so undoing an older entry also rolls back newer entries above it. See [O&O's undo-history instructions](https://manuals.oo-software.com/ooshutup10/en/docs/common-features/undo-history/). A restore point is a broader recovery option when available; neither recovery path is guaranteed by this script.

## Troubleshooting

### Script Won't Run - Execution Policy Error

If you get an execution policy error, run PowerShell as Administrator and execute:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
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

The installer stores the portable executable at one machine-managed path: `%ProgramW6432%\OO Software\O&O ShutUp10\OOSU10.exe`. If `%ProgramW6432%` is unavailable, it uses `%ProgramFiles%` as the machine root. The installer and LocalSystem reapply task resolve this same absolute path; they do not search a user's TEMP, profile, LOCALAPPDATA, or PATH.

If the managed copy is missing or fails signature and version validation, reapplication is skipped with: “The approved O&O executable path is unavailable to the scheduled task. Reapplication was skipped; restore the verified machine copy.” Run the installer again with administrator privileges to restore the verified portable copy.

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
2. **Consistency**: Reapplies the reviewed, digest-verified settings profile
3. **Recovery**: Attempts a restore point and documents O&O's undo history
4. **Persistence**: Automatically reapplies settings after Windows updates
5. **Reviewed Operation**: Interactive runs show the profile's reviewed categories and impact descriptions before applying settings
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
**Copyright:** (c) 2025 myTech.Today. All rights reserved.

This script was developed based on requirements for automated O&O ShutUp10++ installation and configuration with post-Windows Update protection, leveraging 20+ years of IT consulting and PowerShell automation expertise.

## Version History

- **v1.0** - Initial release with basic installation and configuration
- **v2.0** - Added scheduled task for post-Windows Update reapplication
- **v3.0** - Added script relocation, file-based logging, and registry modification for restore points
- **v2.1.0** - Added `-NonInteractive` switch for GUI/automated contexts
- **v2.2.0** - Requires an interactive profile review or an exact approved digest for unattended runs
- **v2.3.0** - Adds reviewed per-setting scope, machine-only SYSTEM reapplication, and a pinned manifest digest

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

---

**myTech.Today** - *Professional IT Services Delivering Results Since 2016*

