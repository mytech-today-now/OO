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
- Allows unattended application only when both exact profile and manifest SHA-256 digests are supplied and match
- Runs O&O in quiet mode after the approval gate

### ✅ **System Protection**
- Attempts a system restore point before applying settings and reports the outcome
- Documents O&O's per-session undo history as the primary settings rollback path

### ✅ **Post-Windows Update Protection** 🆕
- **Automatically creates a scheduled task** that runs after Windows updates
- Reapplies the approved profile after Windows Update reports a successful installation
- Triggers only on Windows Update Client event 19, with a two-minute delay
- Queues overlapping reapplications so a later successful update is not discarded

### ✅ **Comprehensive Logging**
- Color-coded console output with timestamps
- Detailed error messages and stack traces
- Clear status indicators (INFO, SUCCESS, WARNING, ERROR)

## Requirements

- **PowerShell 5.1 or higher**
- **Administrator privileges** (script requires elevation)
- **Windows 10 or Windows 11**
- **Internet connection** (for initial download)

## Execution Policy and Safe Launch

Do not change the stored `CurrentUser` or `LocalMachine` execution policy to run this installer. Prefer an organization-approved signed copy and follow your organization's launch policy. If your policy owner approves a process-scoped launch for this unsigned or local copy, the `-ExecutionPolicy Bypass` examples below apply only to the new PowerShell process; they do not change stored policy values.

On Windows, `MachinePolicy` and `UserPolicy` Group Policy settings take precedence. When neither is set, precedence is `Process`, then `CurrentUser`, then `LocalMachine`. A process-scoped launch cannot override Group Policy. Execution policy is a safety feature, not a security boundary. See [Microsoft's execution policy documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies).

## Usage

### PowerShell Setup on Linux and macOS

`install.sh` installs PowerShell 7 only; O&O ShutUp10 remains Windows-only. For Debian or Ubuntu systems, the script runs both `apt-get update` and `apt-get install -y powershell` as root or through `sudo` when run by a non-root user. It stops when either command fails and returns that command's exit status. A non-root run requires `sudo` to be available.

### Initial Installation and Setup

Run the script as Administrator in CMD:

```cmd (elevated)
<drive>:\path\to\script\install.bat
```

The batch wrapper starts PowerShell 7 (`pwsh`) with `-ExecutionPolicy Bypass` for that child process. Use it only when your organization's policy permits that launch. The wrapper does not change stored policy values, and Group Policy can still block the launch.

If PowerShell 7 is missing, an interactive launch calls WinGet without agreement-acceptance flags so WinGet can show its source and package terms. A launch with redirected input fails closed unless you separately reviewed both agreements and set `MYTECH_ACCEPT_WINGET_AGREEMENTS=YES`. When set on a launch that needs WinGet, this variable passes both acceptance flags and skips those WinGet prompts:

```cmd
set "MYTECH_ACCEPT_WINGET_AGREEMENTS=YES"
call install.bat
set "MYTECH_ACCEPT_WINGET_AGREEMENTS="
```

The consent variable is independent of the selected package. Without it, unattended installation stops before WinGet runs. Agreement rejection is not retried, and WinGet's native failure code is returned. If a non-agreement WinGet failure leaves package state uncertain, inspect it first with `winget list --id Microsoft.PowerShell --exact`; if removal is appropriate, uninstall it manually with `winget uninstall --id Microsoft.PowerShell --exact --source winget`. The launcher does not automatically remove a package.

To launch directly from an elevated CMD prompt, use the engine approved for your environment. These process-scoped examples work with Windows PowerShell 5.1 and PowerShell 7, respectively:

```cmd
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<drive>:\path\to\script\Install-OOShutUp10.ps1"
```

```cmd
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "<drive>:\path\to\script\Install-OOShutUp10.ps1"
```

### What Happens During Installation

1. ✅ Validates the selected profile and confirms the impact summary
2. ✅ Checks if O&O ShutUp10++ is already installed and downloads it if needed
3. ✅ Attempts a system restore point
4. ✅ Applies the approved profile quietly
5. ✅ **Creates a scheduled task** to reapply settings after Windows updates
6. ✅ Offers a graceful, cancellable restart only when O&O reports that settings changed

### Restart Behavior

The restart prompt is shown after the settings and scheduled-task outcome is reported. Press **Y** to request a graceful restart, or **N** / **Esc** to defer. No response within 60 seconds defers the restart. After **Y**, a 60-second save window begins; press any key during that window to cancel. If restart is deferred, the reported setup status remains unchanged; save your work and restart later.

If the restart prompt cannot read console input, the installer logs that manual restart is needed and keeps the reported setup status unchanged.

The default restart does not use `Restart-Computer -Force`, so Windows and applications can block or cancel shutdown while handling unsaved work. Force-closing applications requires the separate `-ForceRestart` option and may lose unsaved changes:

```powershell
.\Install-OOShutUp10.ps1 -ForceRestart
```

`-NonInteractive` always skips restarting, including when `-ForceRestart` is supplied.

### Reapply Settings Only (Used by Scheduled Task)

The script can also run in "ReapplyOnly" mode, which only reapplies privacy settings without installation or task creation:

```powershell
.\Install-OOShutUp10.ps1 -ReapplyOnly
```

This mode is automatically used by the scheduled task after Windows updates.

## Scheduled Task Details

### Task Name
`\myTech.Today\Install-OOShutUp10`

### Safe Task Discovery

If the exact task lookup below returns nothing, use this read-only command to list current or legacy candidates and their paths:

```powershell
Get-ScheduledTask |
    Where-Object { $_.TaskName -like '*OOShutUp*' } |
    Select-Object TaskPath, TaskName, State
```

Review both `TaskPath` and `TaskName` before retrying. This command only lists candidates; it does not start or remove tasks. Do not remove a legacy or same-named task based on a wildcard match.

### Task Description
Reapplies the approved O&O ShutUp10++ profile after successful Windows Update installation events, when executable and profile checks pass.

### Trigger Events
The task triggers only on Windows Update Client **event ID 19** (installation successful). Event ID 20 (installation failure) and event ID 43 (installation started) do not trigger reapplication.

### Execution Details
- **Runs as**: SYSTEM account
- **Privileges**: Highest available
- **Delay**: 2 minutes after the successful installation event
- **Overlapping runs**: Queued and run sequentially
- **Execution**: Hidden PowerShell window (no user interaction)
- **Command**: the canonical absolute Windows PowerShell 5.1 executable under `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`, followed by `-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<ScriptPath>" -ReapplyOnly -SkipRelocate -NonInteractive -ApprovedOOSExecutablePath "<ManagedOOSPath>" -ApprovedProfileSha256 "<PinnedProfileSHA256>" -ApprovedProfileManifestSha256 "<PinnedManifestSHA256>"`
- **Working directory**: the directory containing `<ScriptPath>`

Before creating or replacing the task, the installer confirms the Windows PowerShell file is in the expected system directory, has Microsoft Windows PowerShell product metadata, a valid Microsoft Windows Authenticode signature, and protected system ACLs. A 32-bit caller on 64-bit Windows checks the native binary through `Sysnative`, then stores its canonical `System32` path in the task. The verifier repeats these checks and requires the task action to match that exact path.

If the trusted interpreter is absent, redirected, or changes before candidate staging, the installer logs: `A trusted Windows PowerShell executable could not be resolved. No scheduled task was registered or changed.` The previous task is left as found. Repair or restore the Windows PowerShell 5.1 system executable and its protected ACLs, then rerun the installer with the approved profile. Do not create a replacement task using a bare `PowerShell.exe` command.

If a later verification detects an interpreter change after a disabled validation task is staged, the candidate is rejected before replacing the existing task. If candidate cleanup cannot be confirmed, inspect the exact validation-task path named in the installer log and remove it only after confirming that it is disabled.

If the approved settings are applied but the update task cannot be verified, setup is reported as partial and the installer exits with code `2`. The installer prints recovery instructions for repairing `\myTech.Today\Install-OOShutUp10`. Exit code `0` requires verified settings, an acceptable restore-point outcome, and a verified task.

### Task Security Descriptor

The registered task's owner, group, and DACL are read through Task Scheduler's `RegisteredTask.GetSecurityDescriptor(7)` API. The allowlist is `O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)`: Builtin Administrators own the task, LocalSystem is the group, the DACL is protected from inherited entries, and only LocalSystem and Builtin Administrators receive full task access. The installer applies this descriptor to the disabled validation task and disabled replacement task, reads each back, and enables the replacement only after its action and descriptor verify. It reads the descriptor again after enabling. An existing task with a descriptor outside this allowlist is left unchanged for review.

Descriptor read, write, or verification failures use this fallback and leave the new task disabled: `The task security descriptor could not be verified. No new task was enabled; inspect task permissions before retrying.`

When a replacement fails after the prior task is removed, the installer retains protected `.xml` and `.sddl` rollback files until it verifies both the restored task and descriptor. If automatic rollback is uncertain, do not enable the partial task. Run the following as an administrator after reviewing the exact task identity and both backup paths. The sequence restores the definition disabled, restores and reads back its saved owner/group/DACL, then returns the task to its prior enabled state only when that readback matches the saved descriptor:

```powershell
$taskPath = '\myTech.Today\'
$taskName = 'Install-OOShutUp10'
$xmlPath = '<protected rollback XML path>'
$sddlPath = '<matching protected rollback SDDL path>'
$document = New-Object System.Xml.XmlDocument
$document.Load($xmlPath)
$enabledNode = $document.SelectSingleNode("/*[local-name()='Task']/*[local-name()='Settings']/*[local-name()='Enabled']")
if ($null -eq $enabledNode) { throw 'Rollback XML has no Settings/Enabled element.' }
$wasEnabled = [bool]::Parse($enabledNode.InnerText)
$enabledNode.InnerText = 'false'
Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Xml $document.OuterXml -Force
$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$registeredTask = $service.GetFolder($taskPath).GetTask($taskName)
$savedSddl = [System.IO.File]::ReadAllText($sddlPath)
$registeredTask.SetSecurityDescriptor($savedSddl, 0x10)
$readBackSddl = $registeredTask.GetSecurityDescriptor(7)
if ($readBackSddl -ne $savedSddl) { throw 'Task descriptor readback differs from the saved descriptor; leave the task disabled and inspect it.' }
if ($wasEnabled) { $registeredTask.Enabled = $true }
```

If the readback differs or any command fails, leave the task disabled and inspect the exact task in Task Scheduler before retrying. The descriptor backup covers owner, group, and DACL; it does not request the audit SACL.

#### Disposable VM Task Scheduler acceptance gate

Run the [Issue 02 disposable VM procedure](docs/task-security-descriptor-vm-acceptance.md) on a checkpointed Windows VM before release. It creates only harmless, on-demand `cmd.exe /d /c exit 0` tasks beneath dedicated validation folders; it does not run the installer or use the production task path. Record the actual Task Scheduler descriptor, task action, and enabled state before and after enablement, and test descriptor writes with a standard account against both the allowlisted and permissive fixtures. Keep this live VM result separate from the mocked offline harness result. The fixture procedure does not verify a production installer run; record that live result separately. Do not register these tasks on a workstation.

### Viewing the Scheduled Task

You can view the scheduled task in Task Scheduler:

1. Open **Task Scheduler** (`taskschd.msc`)
2. Navigate to **Task Scheduler Library**
3. Open the **myTech.Today** folder and look for **Install-OOShutUp10**

Or use PowerShell:

```powershell
Get-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
```

The exact registered identity is `TaskPath = \myTech.Today\` and `TaskName = Install-OOShutUp10`.

### Manually Running the Scheduled Task

Running this task manually reapplies the approved settings as SYSTEM. Do this only when that reapplication is intended:

```powershell
Start-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
```

### Task History

To check the task's last-run time and result with PowerShell:

```powershell
Get-ScheduledTaskInfo -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
```

For the full Task Scheduler history, open `\myTech.Today\Install-OOShutUp10` in Task Scheduler and select its **History** tab.

### Removing the Scheduled Task

This removes only the registration at `\myTech.Today\Install-OOShutUp10` after confirmation. It does not delete the installer, approved profile, O&O executable, or logs. The exact-path lookup must return one task before removal is attempted:

```powershell
$taskName = 'Install-OOShutUp10'
$taskPath = '\myTech.Today\'

try {
    $matches = @(Get-ScheduledTask -ErrorAction Stop | Where-Object {
        $_.TaskName -eq $taskName -and $_.TaskPath -eq $taskPath
    })
}
catch {
    throw 'Task lookup failed. No removal was attempted; verify Task Scheduler access and retry.'
}

if ($matches.Count -eq 0) {
    Write-Warning 'The documented task was not found. No removal was attempted; discover the exact task path before retrying.'
}
elseif ($matches.Count -ne 1) {
    throw 'Multiple tasks matched the exact identity. No removal was attempted; inspect Task Scheduler before retrying.'
}
else {
    try {
        Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm -ErrorAction Stop
    }
    catch {
        throw 'Task removal could not be confirmed. Inspect the exact task path in Task Scheduler before retrying.'
    }
}
```

If the lookup fails, the task is absent, or multiple matches are returned, the command does not unregister anything. If unregistering reports an error, inspect the exact path in Task Scheduler before retrying.

## How It Works

### Initial Installation Flow

1. Validate the profile and show its impact summary. A rejection or cancellation stops before relocation or system changes.
2. Check for O&O ShutUp10++ and download it from the official source if needed.
3. Attempt a system restore point.
4. Apply the approved profile quietly.
5. Create the SYSTEM task only when every profile entry is reviewed as machine-scoped; pin both profile and manifest digests.
6. Offer a graceful, cancellable restart only when O&O reports that settings changed; force-closing applications requires `-ForceRestart`.

### Post-Windows Update Flow

1. Windows Update Client event 19 starts the task after a two-minute delay. Event 20 and event 43 do not trigger it.
2. The task runs `-ReapplyOnly -NonInteractive` with its pinned profile and manifest digests.
3. A digest mismatch, unknown scope, or current-user entry stops execution before O&O launches.
4. Only an entirely machine-scoped, digest-matched profile can be reapplied by LocalSystem; the script reports whether settings changed.

## Command Line Arguments

O&O ShutUp10++ is executed with the following arguments:

- `/quiet` - Silent mode (no UI interaction)
- `/nosrp` - Don't create restore point (we create our own)
- The absolute path to the approved `ooshutup10.cfg` profile

Before relocation, registry changes, restore-point creation, scheduled-task changes, or O&O launch, interactive runs display the profile and manifest SHA-256 digests, setting IDs and states, reviewer-provided categories and impact descriptions, and each setting's scope and scope evidence. Type `YES` to accept, `NO` to reject, or `CANCEL` to stop. The summary uses the reviewed manifest text and does not infer setting names, categories, or scope from IDs.

Redirected or unavailable console input uses the non-interactive path and requires both exact approved digests before configuration changes.

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

The profile is resolved beneath the script's package directory regardless of the caller's current directory. Relocation copies the reviewed profile package beside the relocated script and preserves both accepted digests. Before O&O starts, the installer copies both files with exclusive creation into a per-run directory beneath the managed O&O directory, verifies its ACL for SYSTEM and Administrators, and rechecks both reviewed digests on the snapshot. O&O receives only the snapshot `.cfg` path and runs with that directory as its working directory; the snapshot remains in place until O&O exits. The scheduled task runs PowerShell with the script directory as its working directory, retains `-ReapplyOnly`, and supplies the exact absolute managed O&O executable path and both pinned profile digests. O&O exit code `0` is reported as no setting changes and does not prompt for a restart; exit code `30` is treated as settings changed.

If snapshot creation, ACL verification, or digest verification fails, O&O is not started and the installer reports: `The approved profile changed after validation. No O&O process was started; restore the reviewed profile and retry.` If O&O exits but snapshot cleanup fails, verify O&O's applied settings and manually remove the reported `ProfileSnapshot-<guid>` directory beneath the managed O&O directory after confirming no `OOSU10.exe` process still uses it. If the process start outcome cannot be verified, the installer retains the snapshot and reports its path; check for a running O&O process before removing it manually.

The offline harness uses mocked ACL records and does not establish effective Windows permissions. Treat the following disposable-VM check as a release gate; do not run it on a workstation or against a profile that can make unwanted changes.

#### Disposable VM snapshot ACL release gate

1. Create a checkpointed disposable Windows VM with a local NTFS volume. Record the Windows edition, version, build and revision, PowerShell version, VM name, test date, and the exact installer revision or source-file hash. Do not use a network share, synced folder, host-mounted folder, or production profile.
2. Create two local test accounts: one elevated administrator and one standard user with no membership in Administrators. Record each account name, SID, and local group membership. Sign in as the administrator and verify that the standard account cannot elevate.
3. Prepare a separate validation copy with both approved profile files and a harmless, reviewed fixture whose settings are already at their intended values. Verify the fixture digests and O&O executable trust using the normal checks. Do not bypass missing trust pins or profile approval; if a trusted executable and harmless approved fixture are unavailable, record the VM gate as **Not run**.
4. Start the validation run and suspend the O&O child as soon as it appears, before it reads the profile. Confirm the per-run `ProfileSnapshot-<guid>` directory and both files exist. In a second elevated PowerShell process, open both files for read while allowing other handles to share read, write, and delete access, then keep that process open until the lower-privilege probes finish. This makes the snapshot stay in use without a file-sharing lock masking the NTFS ACL result. Record the child PID and snapshot paths. If the snapshot cannot be held for the full check, record the gate as **Not run**.
5. While the files are held open, record the owner and DACL for the snapshot directory, `ooshutup10.cfg`, and `ooshutup10.approval.json`. Save both `Get-Acl` owner/access output and `icacls` output. As the elevated administrator, read both files and record their SHA-256 digests. Use a temporary SYSTEM scheduled task to read and hash the same files; record its result and remove the task after the probe.
6. Sign in to a separate session as the standard user and attempt each of these operations while the snapshot is still held open: create a new file in the snapshot directory; write the original bytes back to each snapshot file; replace each snapshot file with a harmless temporary file; and delete each snapshot file. Record `Denied` or `Allowed` for every attempt and confirm the original file hashes remain unchanged. The test fixture contains no sensitive or production data. If any operation is allowed or a hash changes, stop the O&O child without resuming it, preserve the snapshot and logs for diagnosis, and mark the gate **Failed**.
7. If all mutation attempts are denied and both trusted reads succeed, release the helper handles and resume the O&O child. Wait for the child and installer to finish, then confirm the snapshot was cleaned up. Remove the temporary SYSTEM task, test accounts, fixture files, and validation copy, and restore the VM checkpoint.
8. Save one VM record containing the tested OS/build, installer revision, both account names/SIDs/groups, snapshot directory and file owners/DACLs, SYSTEM and administrator read results, every lower-privilege access attempt and result, before/after file hashes, child/helper PIDs, cleanup result, and final **Pass**, **Failed**, or **Not run** status. Mark **Pass** only when every mutation is denied, both SYSTEM and Administrators can read both files, hashes remain unchanged, and cleanup is recorded.

The hold-open helper must share read, write, and delete access; a `FileShare.None` lock can make access attempts fail for the wrong reason. Keep the helper open only for the test window. The offline harness result and this live VM record are separate evidence: mocked ACL results never substitute for the VM gate.

### Rollback

Before applying the profile, the script attempts to create a Windows restore point and reports whether that attempt succeeded. In O&O ShutUp10, use **Actions → Undo → History** to review and undo the session that applied the profile. O&O keeps a stack of recent sessions, so undoing an older entry also rolls back newer entries above it. See [O&O's undo-history instructions](https://manuals.oo-software.com/ooshutup10/en/docs/common-features/undo-history/). A restore point is a broader recovery option when available; neither recovery path is guaranteed by this script.

## Troubleshooting

### Script Won't Run - Execution Policy Error

If the launch is blocked, stop. Do not change a stored execution policy or retry with a broader scope. Use this guidance:

> PowerShell policy blocks this script. No policy was changed; use an approved signed or process-scoped launch.

If Group Policy blocks the script, contact your administrator for an approved signed script or launch method. A process-scoped option cannot override Group Policy.

### Path Contains Special Characters Error

If you get an ampersand error, use the call operator:

```powershell
& "Q:\_kyle\temp_documents\GitHub\PowerShellScripts\OO\Install-OOShutUp10.ps1"
```

### Scheduled Task Not Triggering

1. Verify the task exists:
   ```powershell
   Get-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
   ```

2. Check task history in Task Scheduler:
   - Open Task Scheduler
   - Open `\myTech.Today\Install-OOShutUp10`
   - Click on the "History" tab

3. Manually test the task:
   ```powershell
   Start-ScheduledTask -TaskName "Install-OOShutUp10" -TaskPath "\myTech.Today\"
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

For unversioned changes, see [CHANGELOG.md](CHANGELOG.md).

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

