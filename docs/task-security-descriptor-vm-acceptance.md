# Issue 02: Task Scheduler security descriptor acceptance on a disposable VM

## Scope

This acceptance checks the Windows Task Scheduler round trip for harmless SYSTEM tasks and tests access from a separate standard account. It does not run `Install-OOShutUp10.ps1`, O&O ShutUp10++, Windows Update, or the production task at `\myTech.Today\Install-OOShutUp10`. Use only the dedicated paths `\Issue02Validation\` and `\Issue02Foreign\` below. The foreign task has the same leaf name as the validation task so the record can show that exact-path operations leave it unchanged.

The offline test `tests/test-ooshutup-payload-trust.ps1` remains the evidence for the installer's mismatch handling, replacement, rollback, and retained recovery files. This VM procedure separately records live Task Scheduler API behavior for harmless tasks. It does not verify a production installer run; that result must be recorded separately. A mocked pass does not count as a live VM result.

## VM setup and evidence

1. Create a checkpointed, disposable Windows VM with no production profile or data. Use an elevated PowerShell session inside the guest. Do not attach a host task store. Do not run this procedure on a workstation.
2. Record the VM name, Windows edition, version, build and revision, PowerShell version, test date, and exact repository revision or installer SHA-256. Record whether the guest is isolated from production networks.
3. Create one local standard test account that is not a member of Builtin Administrators. Record its name, SID, and `whoami /groups` output. Do not use an administrator token for its access probes.
4. Create an evidence directory accessible only to SYSTEM and Builtin Administrators. Capture the directory ACL with `Get-Acl` and `icacls`, then set `$evidenceDirectory` in the elevated session to that verified path. Keep the task XML and SDDL rollback files, command output, and screenshots there until the record is reviewed.
5. Confirm the two dedicated task folders do not exist. If either folder already exists, stop and restore the VM checkpoint rather than reusing or modifying it.

## Register harmless validation and foreign tasks

Use Task Scheduler's scripting API from elevated PowerShell. The task has no trigger and its sole action exits without changing system state. The registration constants are `TASK_CREATE = 0x2`, `TASK_LOGON_SERVICE_ACCOUNT = 5`, and `TASK_ACTION_EXEC = 0`.

```powershell
$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$rootFolder = $service.GetFolder('\')
$validationFolder = $rootFolder.CreateFolder('Issue02Validation', $null)
$foreignFolder = $rootFolder.CreateFolder('Issue02Foreign', $null)
$taskName = 'HarmlessDescriptorProbe'
$expectedSddl = 'O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)'
$mismatchedSddl = 'O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FA;;;AU)'

function New-HarmlessTaskDefinition {
    param([string]$Description)

    $definition = $service.NewTask(0)
    $definition.RegistrationInfo.Description = $Description
    $definition.Principal.UserId = 'SYSTEM'
    $definition.Principal.LogonType = 5
    $definition.Principal.RunLevel = 1
    $definition.Settings.Enabled = $false
    $definition.Settings.AllowDemandStart = $true
    $definition.Settings.DisallowStartIfOnBatteries = $false
    $definition.Settings.StopIfGoingOnBatteries = $false
    $definition.Settings.ExecutionTimeLimit = 'PT1M'
    $action = $definition.Actions.Create(0)
    $action.Path = Join-Path $env:SystemRoot 'System32\cmd.exe'
    $action.Arguments = '/d /c exit 0'
    return $definition
}

$validationTask = $validationFolder.RegisterTaskDefinition(
    $taskName, (New-HarmlessTaskDefinition 'Issue 02 harmless descriptor validation'),
    0x2, 'SYSTEM', $null, 5, $null
)
$foreignTask = $foreignFolder.RegisterTaskDefinition(
    $taskName, (New-HarmlessTaskDefinition 'Issue 02 foreign same-name sentinel'),
    0x2, 'SYSTEM', $null, 5, $null
)
$foreignBefore = [pscustomobject]@{
    Xml = [string]$foreignTask.Definition.XmlText
    Sddl = [string]$foreignTask.GetSecurityDescriptor(7)
    Enabled = [bool]$foreignTask.Enabled
}
```

Save the foreign task's initial `Definition.XmlText`, `GetSecurityDescriptor(7)` result, and `Enabled` value. Confirm its action is also `cmd.exe /d /c exit 0`. Keep its saved values for the final unchanged check.

## Descriptor, action, and enablement round trip

The owner/group/DACL contract is `O:BAG:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)`: owner SID `S-1-5-32-544` (Builtin Administrators), group SID `S-1-5-18` (LocalSystem), protected DACL, and exactly two non-inherited full-access allow ACEs for those SIDs. Record the raw descriptor string as returned by Task Scheduler; also record the parsed owner, group, DACL protection flag, ACE types, masks, flags, and principal SIDs. Treat a parse failure, extra ACE, missing ACE, or unprotected DACL as a failure.

Use this structural check for each readback. It allows the API to format the SDDL string differently while requiring the exact owner, group, protected DACL, and ACE set enforced by the installer:

```powershell
function Test-Issue02TaskDescriptor {
    param([string]$Sddl)

    try {
        $expected = [System.Security.AccessControl.RawSecurityDescriptor]::new($expectedSddl)
        $actual = [System.Security.AccessControl.RawSecurityDescriptor]::new($Sddl)
        if ($null -eq $actual.Owner -or $actual.Owner.Value -ne 'S-1-5-32-544' -or
            $null -eq $actual.Group -or $actual.Group.Value -ne 'S-1-5-18' -or
            $null -eq $actual.DiscretionaryAcl -or
            ($actual.ControlFlags -band [System.Security.AccessControl.ControlFlags]::DiscretionaryAclProtected) -eq 0 -or
            $actual.DiscretionaryAcl.Count -ne $expected.DiscretionaryAcl.Count) {
            return $false
        }

        $actualAces = @()
        foreach ($ace in $actual.DiscretionaryAcl) {
            if ($ace -isnot [System.Security.AccessControl.CommonAce] -or $ace.IsCallback -or
                $ace.AceType -ne [System.Security.AccessControl.AceType]::AccessAllowed -or
                $ace.AceFlags -ne [System.Security.AccessControl.AceFlags]::None) {
                return $false
            }
            $actualAces += $ace
        }

        foreach ($expectedAce in $expected.DiscretionaryAcl) {
            $matches = @($actualAces | Where-Object {
                $_.AccessMask -eq $expectedAce.AccessMask -and
                $_.SecurityIdentifier.Value -eq $expectedAce.SecurityIdentifier.Value
            })
            if ($matches.Count -ne 1) { return $false }
        }
        return $true
    }
    catch {
        return $false
    }
}
```

Before enabling, confirm the registered task is disabled and its definition contains exactly one action, the expected absolute `cmd.exe` path and `/d /c exit 0` arguments, and the LocalSystem principal. Set the descriptor and read it back using the same Task Scheduler API used by the installer:

```powershell
$validationTask.SetSecurityDescriptor($expectedSddl, 0x10)
$beforeEnableSddl = [string]$validationTask.GetSecurityDescriptor(7)
$beforeEnableEnabled = [bool]$validationTask.Enabled
$beforeEnableXml = [string]$validationTask.Definition.XmlText
if (-not (Test-Issue02TaskDescriptor $beforeEnableSddl) -or $beforeEnableEnabled) {
    throw 'The descriptor or disabled-state check failed; keep the validation task disabled.'
}

[xml]$registeredXml = $beforeEnableXml
$namespace = New-Object System.Xml.XmlNamespaceManager($registeredXml.NameTable)
$namespace.AddNamespace('t', 'http://schemas.microsoft.com/windows/2004/02/mit/task')
$actions = $registeredXml.SelectNodes('/t:Task/t:Actions/*', $namespace)
$command = $registeredXml.SelectSingleNode('/t:Task/t:Actions/t:Exec/t:Command', $namespace)
$arguments = $registeredXml.SelectSingleNode('/t:Task/t:Actions/t:Exec/t:Arguments', $namespace)
$principal = $registeredXml.SelectSingleNode('/t:Task/t:Principals/t:Principal/t:UserId', $namespace)
$logonType = $registeredXml.SelectSingleNode('/t:Task/t:Principals/t:Principal/t:LogonType', $namespace)
$runLevel = $registeredXml.SelectSingleNode('/t:Task/t:Principals/t:Principal/t:RunLevel', $namespace)
$expectedCommand = Join-Path $env:SystemRoot 'System32\cmd.exe'
if ($actions.Count -ne 1 -or $null -eq $command -or $command.InnerText -ine $expectedCommand -or
    $null -eq $arguments -or $arguments.InnerText -cne '/d /c exit 0' -or
    $null -eq $principal -or $principal.InnerText -notin @('SYSTEM', 'S-1-5-18') -or
    $null -eq $logonType -or $logonType.InnerText -cne 'ServiceAccount' -or
    $null -eq $runLevel -or $runLevel.InnerText -cne 'HighestAvailable') {
    throw 'The disabled task action or LocalSystem principal does not match the harmless fixture.'
}
```

Record all three values immediately before enablement. Do not enable unless the readback has the expected owner, group, protected DACL, and exact allowed ACE set, and the action and principal match the harmless fixture.

### Descriptor mismatch case

While the validation task is disabled, write `$mismatchedSddl` and read it back. Record the returned SDDL and parsed ACEs. The extra Authenticated Users ACE must fail the descriptor check. Confirm the task remains disabled, preserve its XML and descriptor output, and do not enable it. If the Task Scheduler API does not preserve the injected mismatch, record the actual returned descriptor and mark this injected case **Not run** rather than claiming a mismatch was tested.

```powershell
$validationTask.SetSecurityDescriptor($mismatchedSddl, 0x10)
$mismatchReadback = [string]$validationTask.GetSecurityDescriptor(7)
if ($validationTask.Enabled) {
    throw 'The candidate became enabled during the mismatch case; mark the VM case Failed.'
}
if (Test-Issue02TaskDescriptor $mismatchReadback) {
    Write-Warning 'Task Scheduler did not preserve the injected mismatch; record this mismatch case as Not run.'
}
else {
    Write-Output 'The mismatch was observed and the candidate remains disabled.'
}
```

After recording that result, restore `$expectedSddl`, read it back, and repeat the full descriptor, action, principal, and disabled-state checks. Only then enable the validation task and read the descriptor and enabled state again:

```powershell
$validationTask.SetSecurityDescriptor($expectedSddl, 0x10)
$preEnableReadback = [string]$validationTask.GetSecurityDescriptor(7)
if (-not (Test-Issue02TaskDescriptor $preEnableReadback) -or $validationTask.Enabled) {
    throw 'The restored descriptor or disabled-state check failed; keep the task disabled.'
}
# Confirm the action and LocalSystem principal again before this line.
$validationTask.Enabled = $true
$afterEnableSddl = [string]$validationTask.GetSecurityDescriptor(7)
$afterEnableEnabled = [bool]$validationTask.Enabled
if (-not (Test-Issue02TaskDescriptor $afterEnableSddl) -or -not $afterEnableEnabled) {
    throw 'The after-enable descriptor or enabled-state check failed.'
}
```

The after-enable descriptor must pass the same parsed checks as the before-enable descriptor, and `Enabled` must be true. Save the before/after SDDL, enabled values, task XML, and action/principal details immediately after enablement in the VM record.

Disable the validation task again before the standard-account probes. Confirm it is disabled and its descriptor still passes the allowlist check:

```powershell
$validationTask.Enabled = $false
$postEnableDisabledSddl = [string]$validationTask.GetSecurityDescriptor(7)
if ($validationTask.Enabled -or -not (Test-Issue02TaskDescriptor $postEnableDisabledSddl)) {
    throw 'The validation task did not return to a disabled state with the approved descriptor.'
}
```

## Standard-user effective-access check

This test checks real access on the guest instead of inferring effective access from SDDL text. Keep both tasks disabled with no triggers. Run the first probe from a separate logon for the recorded standard account, without elevation. If `GetTask` itself is denied, record that result and the HRESULT; do not claim the descriptor write was independently attempted. The admin session must verify both tasks still exist and their descriptors are unchanged after the probe.

```powershell
$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$validationFolder = $service.GetFolder('\Issue02Validation\')
$foreignFolder = $service.GetFolder('\Issue02Foreign\')
$taskName = 'HarmlessDescriptorProbe'
$userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$userWriteAttempt = "O:${userSid}G:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FA;;;AU)"

try {
    $validationTask = $validationFolder.GetTask($taskName)
    $validationTask.SetSecurityDescriptor($userWriteAttempt, 0x10)
    Write-Output 'UNEXPECTED: the standard account changed the allowlisted task descriptor.'
}
catch {
    Write-Output ('Descriptor write denied or task unreadable: HRESULT 0x{0:X8}; {1}' -f $_.Exception.HResult, $_.Exception.Message)
}

try {
    $foreignFolder.DeleteTask($taskName, 0)
    Write-Output 'UNEXPECTED: the standard account deleted the foreign sentinel task.'
}
catch {
    Write-Output ('Task deletion denied: HRESULT 0x{0:X8}; {1}' -f $_.Exception.HResult, $_.Exception.Message)
}
```

Record each result independently. A **Pass** requires access-denied results for both probes and an elevated readback showing both tasks still disabled with their original allowlisted descriptors. Any successful write or deletion is **Failed**; preserve the VM and evidence, then restore the checkpoint. Do not infer denial from the account name alone: save its SID and group membership.

For a permissive positive control, return to the elevated session and apply `$mismatchedSddl` to the disabled validation task. Confirm the extra Authenticated Users ACE is present. Then sign in as the same standard account and set the descriptor to one that retains the SYSTEM and Administrators ACEs but removes the Authenticated Users ACE:

```powershell
$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$validationTask = $service.GetFolder('\Issue02Validation\').GetTask('HarmlessDescriptorProbe')
$userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$userChangedSddl = "O:${userSid}G:SYD:P(A;;FA;;;SY)(A;;FA;;;BA)"
$validationTask.SetSecurityDescriptor($userChangedSddl, 0x10)
Write-Output 'The standard account changed the deliberately permissive fixture descriptor.'
```

Return to the elevated session, capture the resulting descriptor, confirm the task stayed disabled, then restore `$expectedSddl` and verify it with `Test-Issue02TaskDescriptor`. This positive control must succeed for the access probe to count as validated. If it is denied, record the standard-account access result as **Not run** and investigate the guest policy. Never enable either fixture to test access.

## Rollback and foreign-task checks

1. Before a rollback exercise, save the validation task's exact XML and owner/group/DACL SDDL to the protected evidence directory. Save its prior enabled state too. These are disposable fixture backups, not production rollback files.

   ```powershell
   $rollbackXml = [string]$validationTask.Definition.XmlText
   $rollbackSddl = [string]$validationTask.GetSecurityDescriptor(7)
   $rollbackWasEnabled = [bool]$validationTask.Enabled
   $rollbackXmlPath = Join-Path $evidenceDirectory 'HarmlessDescriptorProbe.rollback.xml'
   $rollbackSddlPath = Join-Path $evidenceDirectory 'HarmlessDescriptorProbe.rollback.sddl'
   [System.IO.File]::WriteAllText($rollbackXmlPath, $rollbackXml)
   [System.IO.File]::WriteAllText($rollbackSddlPath, $rollbackSddl)
   ```

2. Replace only `\Issue02Validation\HarmlessDescriptorProbe` with a disabled harmless candidate using the Task Scheduler API. Inject the mismatched descriptor and verify the candidate remains disabled.

   ```powershell
   $candidateDefinition = New-HarmlessTaskDefinition 'Issue 02 injected rollback candidate'
   $candidateTask = $validationFolder.RegisterTaskDefinition(
       $taskName, $candidateDefinition, 0x6, 'SYSTEM', $null, 5, $null
   )
   $candidateTask.SetSecurityDescriptor($mismatchedSddl, 0x10)
   $candidateMismatch = [string]$candidateTask.GetSecurityDescriptor(7)
   if ($candidateTask.Enabled) {
       throw 'The rollback candidate became enabled during the mismatch case.'
   }
   if (Test-Issue02TaskDescriptor $candidateMismatch) {
       Write-Warning 'Task Scheduler did not preserve the rollback-candidate mismatch; record this case as Not run.'
   }
   else {
       Write-Output 'The rollback candidate mismatch was observed and it remains disabled.'
   }
   ```

3. Restore the saved definition with its `Enabled` element set to false, apply the saved descriptor, and read it back. The following forced mismatch simulates an uncertain rollback readback; it is an injected failure, not evidence that Windows normalized the SDDL this way:

   ```powershell
   [xml]$restoreDocument = [System.IO.File]::ReadAllText($rollbackXmlPath)
   $restoreNamespace = New-Object System.Xml.XmlNamespaceManager($restoreDocument.NameTable)
   $restoreNamespace.AddNamespace('t', 'http://schemas.microsoft.com/windows/2004/02/mit/task')
   $enabledNode = $restoreDocument.SelectSingleNode('/t:Task/t:Settings/t:Enabled', $restoreNamespace)
   if ($null -eq $enabledNode) { throw 'Saved task XML has no Settings/Enabled element.' }
   $enabledNode.InnerText = 'false'
   $restoredTask = $validationFolder.RegisterTask(
       $taskName, $restoreDocument.OuterXml, 0x6, 'SYSTEM', $null, 5, $null
   )
   $restoredTask.SetSecurityDescriptor($rollbackSddl, 0x10)
   $restoredTask.SetSecurityDescriptor($mismatchedSddl, 0x10) # forced readback mismatch
   $rollbackReadback = [string]$restoredTask.GetSecurityDescriptor(7)
   $uncertainRollback = -not (Test-Issue02TaskDescriptor $rollbackReadback)
   $backupFilesUnchanged =
       [System.IO.File]::ReadAllText($rollbackXmlPath) -ceq $rollbackXml -and
       [System.IO.File]::ReadAllText($rollbackSddlPath) -ceq $rollbackSddl
   if ($restoredTask.Enabled -or -not $backupFilesUnchanged) {
       throw 'Uncertain rollback did not retain its evidence with the task disabled.'
   }
   if (-not $uncertainRollback) {
       Write-Warning 'Task Scheduler did not preserve the forced rollback mismatch; record this rollback case as Not run.'
   }
   else {
       Write-Output 'Uncertain rollback is confirmed; the task remains disabled and evidence is retained.'
   }
   ```

   Record that rollback is uncertain, verify the restored task remains disabled, and verify both saved evidence files remain present and unchanged. Do not delete the evidence or enable the task while the descriptor is uncertain.
4. Compare the foreign same-named task's final `Definition.XmlText`, descriptor, and enabled state with the values captured before the exercise. They must be identical:

   ```powershell
   $foreignAfterTask = $foreignFolder.GetTask($taskName)
   $foreignAfter = [pscustomobject]@{
       Xml = [string]$foreignAfterTask.Definition.XmlText
       Sddl = [string]$foreignAfterTask.GetSecurityDescriptor(7)
       Enabled = [bool]$foreignAfterTask.Enabled
   }
   if ($foreignAfter.Xml -cne $foreignBefore.Xml -or
       $foreignAfter.Sddl -cne $foreignBefore.Sddl -or
       $foreignAfter.Enabled -ne $foreignBefore.Enabled) {
       throw 'The foreign same-named task changed; preserve the VM and evidence.'
   }
   ```

   If any value differs, mark the run **Failed** and preserve the VM and evidence for investigation.
5. Save one record with the VM metadata, exact Task Scheduler paths, registered action and principal, before/after descriptors and enabled states, mismatch result, rollback result, foreign-task comparison, evidence file paths and ACLs, and final **Pass**, **Failed**, or **Not run** status. Mark **Pass** only if the valid round trip succeeds, mismatch never enables a candidate, uncertain rollback retains both evidence files and leaves the task disabled, and the foreign task is unchanged.
6. After exporting and reviewing the record, remove only the two test tasks and their now-empty dedicated folders, then restore the VM checkpoint. Never use wildcard task deletion.

The live procedure verifies Task Scheduler behavior on the recorded Windows build. It does not prove behavior on other Windows builds or replace the offline installer harness. The installer harness must separately pass its staged-task, descriptor mismatch, rollback-retention, and foreign same-named task assertions.
