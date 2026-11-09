# O&O ShutUp10 settings profile package

The installer expects this directory to contain two reviewed files:

- `ooshutup10.cfg`, exported from O&O ShutUp10 Free Edition.
- `ooshutup10.approval.json`, a versioned manifest for that exact export.

The schema version 4 manifest has exactly these properties:

```json
{
  "schemaVersion": 4,
  "ooshutupVersionRange": {
    "minimum": "<reviewed four-part O&O FileVersion>",
    "maximum": "<reviewed four-part O&O FileVersion>"
  },
  "profileVersion": "<reviewed version>",
  "sha256": "<SHA-256 of the exact cfg file bytes>",
  "settingIds": ["<every actual setting ID in the cfg, once>"],
  "settingDetails": [
    {
      "id": "<same actual setting ID>",
      "state": "<+ or ->",
      "category": "<reviewed O&O category>",
      "impact": "<verified description of the setting's effect>",
      "scope": "<machine or current-user>",
      "scopeEvidence": "<current O&O source and per-setting two-account VM record reference>"
    }
  ]
}
```

The inclusive `ooshutupVersionRange` is compared with the executable's four-part Windows `FileVersion` value. Set `minimum` and `maximum` to the same value for exact-release approval. A wider range is allowed only when the owner has reviewed the setting descriptions, availability, and scope throughout that entire range. The supported-version floor in the installer is only a general payload safety check; it never expands this profile's reviewed range.

Do not add a profile until an owner has reviewed every setting ID, its `+` or `-` state, its category, impact description, availability, and scope against current O&O setting information for the declared version range. Do not infer names, categories, availability, or scope from IDs or `profileVersion`. For every setting proposed for a SYSTEM profile, record the exact O&O version, current official source URL and review date, setting ID and description, plus a disposable Windows VM test record. For a range, retain source and VM evidence for every release included in the range before approving its upper bound. The record must identify the Windows build, O&O version, two disposable accounts and SIDs, the setting and state tested, and the exact relevant registry paths and values observed before and after for the machine hive and both user hives. It must state whether only the intended hive changed. Store one record per setting under `profiles/scope-evidence/<SETTING-ID>.md` and have `scopeEvidence` identify both the O&O source and that record.

The manifest validator checks the declared version range, scope, and evidence text and binds those claims to the manifest digest. It cannot authenticate the source, inspect the VM record, or prove which hive the setting changes. A reviewer must open and verify the current O&O information and the matching two-account record before approving the profile digests or creating the SYSTEM task. A nonempty evidence string alone is not proof. If a record is absent, stale, incomplete, or disagrees with the declared scope or release range, omit that setting from the SYSTEM profile. Do not create a per-user task automatically.

The installer displays the reviewed O&O file-version range, setting details, and both file digests before interactive approval. It checks the executable against the inclusive range before applying settings and again before registering or running the SYSTEM task. Unattended use requires `-NonInteractive -ApprovedProfileSha256 <exact-cfg-sha256> -ApprovedProfileManifestSha256 <exact-manifest-sha256>`; the scheduled task carries both digests so scope and version metadata cannot be changed independently of approval. Commit changes to the profile and manifest together so the review covers the selected settings and their digest. The manifest is not a digital signature; approval relies on reviewing the pair in source control and controlling write access to the installed package. The installer rejects absent files, empty or malformed metadata, duplicate or unapproved IDs, unsupported scope, mismatched IDs or states, and digest mismatches before system changes or O&O launch.

Schema versions 1 through 3 are rejected. To migrate an older manifest, an owner must repeat the setting review against the exact O&O release or fully reviewed range, refresh each source and VM evidence record, set `schemaVersion` to `4`, add `ooshutupVersionRange`, and recalculate the cfg and manifest approval digests. Do not carry forward `profileVersion` or old review text as evidence of compatibility. Owners refresh evidence for a new O&O release by inspecting its current setting descriptions, availability, and scope, repeating the disposable two-account VM checks, then updating the range only after every included release is reviewed. A release outside the approved range requires this review and digest reapproval before application or task registration.

The Windows Update task runs as LocalSystem. It is created and reapplies settings only when every profile entry is explicitly classified as `machine` with scope evidence. If any entry is `current-user`, task creation or reapplication is skipped; run those settings under the intended signed-in user or remove them from the SYSTEM profile. Missing, unknown, or unsupported scope blocks reapplication with: `The task profile contains settings with unknown user scope. Reapplication was skipped; review the profile scope before retrying.` Current-user settings have a separate status that explains how to handle them. No per-user task is created automatically.

O&O `.cfg` lines contain an ID, a tab or semicolon, and an explicit `+` or `-` state. Blank lines, `#` comments, and optional explanatory text are accepted. A `-` state is an explicit choice, not a record of a previous value. Consult [O&O's profile file structure documentation](https://manuals.oo-software.com/ooshutup10/en/docs/common-features/profile-structure/) and review the setting descriptions before importing an export.

`ooshutup10.cfg` is included as an estimated candidate containing two telemetry restrictions. It is not an O&O export, is not owner-reviewed, and its setting states, recommendation levels, availability, and scope are not verified. Do not import or apply it as-is. No approval manifest or per-setting scope evidence record is included, so the installer rejects the incomplete package before O&O launch or privacy/registry changes. The scope gate and mock tests do not establish registry behavior; validate every real setting in a disposable Windows VM before approving a profile.

Recovery validation is a separate release gate. No recovery VM record is included in this checkout. Follow [the disposable-VM recovery acceptance procedure](../docs/recovery-vm-acceptance.md) with a harmless reviewed fixture before reporting O&O Undo History behavior as verified. Record System Restore availability separately; neither route is guaranteed.
