#!/usr/bin/env bash
set -euo pipefail

if [[ "$OSTYPE" != linux-gnu* ]]; then
    printf 'This mock suite requires Linux; no package commands were run.\n' >&2
    exit 77
fi

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
installer="$repo_root/install.sh"
temp_root="$(mktemp -d)"
trap 'rm -rf "$temp_root"' EXIT
assertion_count=0

assert_equal() {
    local expected="$1"
    local actual="$2"
    local description="$3"

    assertion_count=$((assertion_count + 1))
    if [[ "$expected" != "$actual" ]]; then
        printf 'FAIL: %s\nExpected: <%s>\nActual:   <%s>\n' "$description" "$expected" "$actual" >&2
        exit 1
    fi
}

assert_contains() {
    local expected="$1"
    local actual="$2"
    local description="$3"

    assertion_count=$((assertion_count + 1))
    if [[ "$actual" != *"$expected"* ]]; then
        printf 'FAIL: %s\nMissing: <%s>\n' "$description" "$expected" >&2
        exit 1
    fi
}

setup_case() {
    local name="$1"
    case_uid="$2"
    case_sudo="$3"
    case_fail_stage="$4"
    case_fail_status="$5"
    case_install_pwsh="$6"
    case_dir="$temp_root/$name"
    case_bin="$case_dir/bin"
    case_apt_log="$case_dir/apt.log"
    case_sudo_log="$case_dir/sudo.log"

    mkdir -p "$case_bin"
    : > "$case_apt_log"
    : > "$case_sudo_log"

    cat > "$case_bin/dirname" <<'MOCK'
#!/bin/sh
case "$1" in
    */*) printf '%s\n' "${1%/*}" ;;
    *) printf '.\n' ;;
esac
MOCK

    cat > "$case_bin/id" <<'MOCK'
#!/bin/sh
if [ "$1" = '-u' ]; then
    printf '%s\n' "$MOCK_UID"
    exit 0
fi
exit 2
MOCK

    cat > "$case_bin/apt-get" <<'MOCK'
#!/bin/sh
printf '%s apt-get %s\n' "${MOCK_PRIVILEGE:-direct}" "$*" >> "$APT_LOG"
if [ "$1" = "$MOCK_FAIL_STAGE" ] && [ -n "$MOCK_FAIL_STAGE" ]; then
    exit "$MOCK_FAIL_STATUS"
fi
if [ "$1" = 'install' ] && [ "$MOCK_INSTALL_PWSH" = '1' ]; then
    printf '#!/bin/sh\nexit 0\n' > "$MOCK_BIN/pwsh"
    /bin/chmod +x "$MOCK_BIN/pwsh"
fi
exit 0
MOCK

    if [[ "$case_sudo" == 'yes' ]]; then
        cat > "$case_bin/sudo" <<'MOCK'
#!/bin/sh
printf '%s\n' "$*" >> "$SUDO_LOG"
MOCK_PRIVILEGE=sudo "$@"
MOCK
    fi

    chmod +x "$case_bin/dirname" "$case_bin/id" "$case_bin/apt-get"
    if [[ -f "$case_bin/sudo" ]]; then
        chmod +x "$case_bin/sudo"
    fi
}

run_installer() {
    set +e
    actual_output="$(
        MOCK_UID="$case_uid" \
        MOCK_FAIL_STAGE="$case_fail_stage" \
        MOCK_FAIL_STATUS="$case_fail_status" \
        MOCK_INSTALL_PWSH="$case_install_pwsh" \
        MOCK_BIN="$case_bin" \
        APT_LOG="$case_apt_log" \
        SUDO_LOG="$case_sudo_log" \
        PATH="$case_bin" \
        "$BASH" "$installer" 2>&1
    )"
    actual_status=$?
    set -e
}

file_text() {
    if [[ -f "$1" ]]; then
        printf '%s' "$(<"$1")"
    fi
}

setup_case root 0 no '' 0 1
run_installer
assert_equal 0 "$actual_status" 'root apt install exits successfully'
assert_equal $'direct apt-get update\ndirect apt-get install -y powershell' "$(file_text "$case_apt_log")" 'root runs both exact apt commands directly'
assert_equal '' "$(file_text "$case_sudo_log")" 'root does not invoke sudo'
assert_contains 'O&O ShutUp10 is a Windows-only application.' "$actual_output" 'Windows-only warning remains present'
assert_contains '[OK] PowerShell 7 installed successfully!' "$actual_output" 'successful install satisfies the pwsh postcondition'

setup_case sudo-success 1000 yes '' 0 1
run_installer
assert_equal 0 "$actual_status" 'non-root sudo apt install exits successfully'
assert_equal $'sudo apt-get update\nsudo apt-get install -y powershell' "$(file_text "$case_apt_log")" 'non-root apt calls both execute through sudo without a direct retry'
assert_equal $'apt-get update\napt-get install -y powershell' "$(file_text "$case_sudo_log")" 'sudo receives the exact apt arguments'
assert_contains '[OK] PowerShell 7 installed successfully!' "$actual_output" 'sudo install satisfies the pwsh postcondition'

setup_case missing-sudo 1000 no '' 0 1
run_installer
assert_equal 127 "$actual_status" 'missing sudo returns the missing-command status'
assert_equal '' "$(file_text "$case_apt_log")" 'missing sudo does not run apt directly'
assert_equal '' "$(file_text "$case_sudo_log")" 'missing sudo has no sudo invocation'
assert_contains 'sudo is required for both apt-get update and apt-get install when not running as root.' "$actual_output" 'missing sudo reports the required privilege'

setup_case update-failure 1000 yes update 23 1
run_installer
assert_equal 23 "$actual_status" 'sudo update failure preserves its exit status'
assert_equal 'sudo apt-get update' "$(file_text "$case_apt_log")" 'update failure does not fall back to direct apt'
assert_equal 'apt-get update' "$(file_text "$case_sudo_log")" 'update failure is reported at the sudo update stage'
assert_contains '[ERROR] apt-get update failed (exit 23).' "$actual_output" 'update failure reports its stage and status'
assert_contains 'The apt-get operation failed at the recorded stage. Verify package state and run the next required command with sudo.' "$actual_output" 'update failure includes the recovery wording'

setup_case install-failure 1000 yes install 42 1
run_installer
assert_equal 42 "$actual_status" 'sudo install failure preserves its exit status'
assert_equal $'sudo apt-get update\nsudo apt-get install -y powershell' "$(file_text "$case_apt_log")" 'install failure runs no unprivileged fallback'
assert_equal $'apt-get update\napt-get install -y powershell' "$(file_text "$case_sudo_log")" 'install failure is reported at the sudo install stage'
assert_contains '[ERROR] apt-get install failed (exit 42).' "$actual_output" 'install failure reports its stage and status'
assert_contains 'The apt-get operation failed at the recorded stage. Verify package state and run the next required command with sudo.' "$actual_output" 'install failure includes the recovery wording'

setup_case unsupported-manager 1000 no '' 0 1
rm -f "$case_bin/apt-get"
run_installer
assert_equal 1 "$actual_status" 'unsupported package manager exits with failure'
assert_equal '' "$(file_text "$case_apt_log")" 'unsupported package manager does not invoke apt'
assert_contains '[ERROR] Could not detect package manager.' "$actual_output" 'unsupported manager reports the existing status'

setup_case missing-pwsh-postcondition 0 no '' 0 0
run_installer
assert_equal 1 "$actual_status" 'a successful package command without pwsh fails the existing postcondition'
assert_equal $'direct apt-get update\ndirect apt-get install -y powershell' "$(file_text "$case_apt_log")" 'postcondition check follows the expected root package commands'
assert_contains '[ERROR] PowerShell installation failed.' "$actual_output" 'postcondition failure is reported'

printf 'Passed %s apt installer mock assertions.\n' "$assertion_count"
