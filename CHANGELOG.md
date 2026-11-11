# Changelog

## Unreleased

### Changed

- Align the README company profile and installer help, banner, and completion text with the Omaha and Lincoln service area, listed services, and contact links.
- Automatically select source-checkout development mode when no release manifest or enforced `AllSigned` policy is present; use the development elevation helper only for that checkout path.
- Add a setup-only first run that verifies or downloads O&O and opens its UI in interactive sessions when a reviewed profile pair is absent, without applying settings or creating a restore point or scheduled task.
- Install a new setup-only O&O copy under the current user's LocalAppData when no verified machine copy exists, avoiding a Program Files write on first run; keep approved-profile installation machine-wide.
- Fall back to the writable elevated machine directory when setup-only LocalAppData cannot be prepared, and include the underlying folder, download, or promotion error in the protected log.
- Report the setup-only outcome separately from successful profile configuration and task registration.
- Locate PowerShell 7 from PATH, standard machine paths, and the Windows app alias, including immediately after WinGet installation.
- Add a protected, rotating SYSTEM application log with safe stage diagnostics; logging failures stop initialization or prevent a successful completion result.
- Add an aggregate offline safety runner and Windows PowerShell 5.1, PowerShell 7, and Linux Bash CI jobs.
- Check administrator access before runtime discovery and request UAC elevation through the signed helper when needed. Preserve installer arguments and exit status, and stop when elevation is canceled.
- Run `dnf` and `pacman` installation directly as root or through `sudo`, preserving package-manager failures without retrying.
- Add a checkpointed disposable-VM acceptance procedure for O&O Undo History recovery. Record System Restore availability separately and keep recovery **NOT VERIFIED** until a reviewed VM record is complete.
- Expand mock coverage for installer elevation, argument forwarding, WinGet consent, restore-point outcomes, and `dnf`/`pacman` behavior.
- Reapply the approved profile only after Windows Update Client event 19, wait two minutes, and queue overlapping task runs.
- Verify the task trigger and queue policy before replacing the owned task. Report partial setup with exit code `2` and task recovery instructions when settings were applied but the task is unverified.
- Offer a timed, cancellable graceful restart after settings change. Force-closing applications requires the explicit `-ForceRestart` option, and non-interactive runs skip restart.
- Keep execution-policy examples process-scoped and require explicit WinGet agreement consent when PowerShell 7 must be installed without an interactive console.
- Preserve WinGet and installer exit codes through the batch launcher instead of reporting unconditional completion.
- Add the dated pre-mortem report and its originating prompt under `ai-prompts/`.
- Treat redirected or unavailable console input as non-interactive before profile approval, require exact profile and manifest digests, and defer restart when no prompt can be used.
- Run apt update and install with root privileges, require `sudo` for non-root apt setup, and preserve the failing command's exit status without an unprivileged retry.
- Document exact scheduled-task identity, read-only candidate discovery, task history, and confirmation before exact-path removal.
- Add the November 1 repository pre-mortem report under `ai-prompts/`.
