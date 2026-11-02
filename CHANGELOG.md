# Changelog

## Unreleased

### Changed

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
