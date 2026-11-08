# Issue 13: Disposable Windows VM recovery acceptance

**Current evidence state: NOT VERIFIED.** This procedure is not a completed test record. No approved profile or recovery VM record is included in this checkout. Keep release recovery unverified until a completed record is reviewed.

## Safety and scope

Run this acceptance gate only in a checkpointed, disposable Windows VM. Never apply the fixture on a developer workstation, a production device, or a VM containing personal or business data. Restore the VM checkpoint after the exercise.

The gate tests O&O ShutUp10++ Undo History against one harmless, owner-reviewed profile fixture. It records System Restore availability separately; it does not test a full System Restore rollback. Neither recovery route is guaranteed by the installer or by this procedure.

## Prerequisites

1. Create a clean Windows VM checkpoint and confirm the VM is isolated from production systems and data. Do not use host-mounted, synced, or shared production folders for the fixture.
2. Prepare a profile and approval manifest that pass the repository's normal review, digest, source-integrity, and O&O version checks. Select at least one harmless, reversible setting whose current state differs from the reviewed target state. The fixture must not disable security, networking, sign-in, recovery, or VM management features.
3. Record the exact setting IDs, intended states and scopes, source evidence, profile and manifest digests, and the registry paths, value names, types, and user SIDs used to verify each setting. Use a dedicated disposable account for any current-user setting.
4. Use the normal supported installer and O&O launch path for the tested release. Do not bypass approval, signature, digest, or payload trust checks. If a trusted O&O executable or reviewed fixture is unavailable, mark the gate **Not run**.

## Procedure

1. Record the test date and time with time zone, VM name, Windows edition, version, build and revision, PowerShell version, installer revision or source hash, O&O version and executable hash, profile and manifest digests, and fixture review reference.
2. Before applying the profile, query and save the exact state of each reviewed setting. Include the registry hive, path, value name, value type, value data, and account SID where applicable. Save the evidence outside the VM checkpoint so it can be retained after cleanup.
3. Run the approved installer once in the disposable VM. Record the action, installer and O&O result, log or evidence paths, and any restore-point creation status. Do not retry an uncertain O&O launch. If the installer does not complete the approved profile application, record **Not run** for the Undo History exercise and preserve the failure details.
4. Query the same setting state immediately after application. The exercise has not tested recovery if no reviewed setting changed. Record the exact O&O session shown in **Actions → Undo → History**, including its time and profile identity.
5. In O&O ShutUp10++, select the session that applied the fixture and use Undo History to undo it. Record the exact UI actions and result. Query the same setting state again after undo. Compare each value with its recorded pre-application value and confirm that no out-of-scope setting changed.
6. Separately record System Restore protection status and whether Windows lists the exact restore point attempted by the installer. Save the query used, point description and sequence number when present, or the exact error or unavailable result. A listed point means only that Windows listed it; do not report System Restore rollback as tested or successful.
7. Export the completed evidence record and logs, remove the fixture and temporary test data, then restore the VM checkpoint. Record whether cleanup and checkpoint restoration completed. Keep the record outside the restored VM.

## Result rules

- Mark the Undo History exercise **Pass** only when a real reviewed setting changed, the matching O&O session was undone, every captured target value matches its exact pre-application value, no out-of-scope value changed, and cleanup is recorded.
- Mark it **Failed** if the profile was applied but Undo History is unavailable, fails, or leaves any target value different from its recorded pre-application value; also fail if an out-of-scope setting changes.
- Mark it **Not run** if prerequisites fail, the profile does not change a setting, the O&O session cannot be identified, or evidence is incomplete. Do not treat an offline mock, a successful provider response, or a listed restore point as a passing recovery exercise.
- Report System Restore separately as **point listed**, **not listed**, or **unknown**. This field does not change the Undo History result and does not prove that a restore operation would succeed.

## Recovery VM record

Create one record per run with these fields:

```text
Run date and time with time zone:
VM name and checkpoint identifier:
Windows edition, version, build, and revision:
PowerShell version:
Installer revision or source hash:
O&O version, executable hash, and trust result:
Profile and manifest digests:
Fixture review reference and setting IDs:
Before state evidence path:
Profile application action and result:
After-application state evidence path:
O&O Undo History session identity and actions:
Undo result and after-undo state evidence path:
System Restore protection status:
System Restore point query, description, sequence number, and result:
Cleanup actions and result:
VM checkpoint restoration result:
Undo History outcome: Pass / Failed / Not run
Evidence reviewer and review date:
```

Keep O&O Undo History evidence and System Restore availability evidence as separate results. A passed VM record demonstrates only the tested fixture on the recorded Windows and O&O versions; it does not make either recovery route guaranteed.
