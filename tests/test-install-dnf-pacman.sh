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

file_text() {
    if [[ -f "$1" ]]; then
        cat "$1"
    fi
}

setup_case() {
    local name="$1"
    case_manager="$2"
    case_uid="$3"
    case_sudo="$4"
    case_sudo_status="$5"
    case_package_status="$6"
    case_install_pwsh="$7"
    case_dir="$temp_root/$name"
    case_bin="$case_dir/bin"
    case_package_log="$case_dir/package.log"
    case_sudo_log="$case_dir/sudo.log"
    case_stdout="$case_dir/stdout.log"
    case_stderr="$case_dir/stderr.log"

    mkdir -p "$case_bin"
    : > "$case_package_log"
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
if [ "$#" -eq 1 ] && [ "$1" = '-u' ]; then
    printf '%s\n' "$MOCK_UID"
    exit 0
fi
exit 2
MOCK

    cat > "$case_bin/$case_manager" <<'MOCK'
#!/bin/sh
printf '%s|%s\n' "${MOCK_PRIVILEGE:-direct}" "$*" >> "$PACKAGE_LOG"
printf '%s stdout marker\n' "$MOCK_MANAGER"
printf '%s stderr marker\n' "$MOCK_MANAGER" >&2
if [ "$MOCK_PACKAGE_STATUS" -ne 0 ]; then
    exit "$MOCK_PACKAGE_STATUS"
fi
if [ "$MOCK_INSTALL_PWSH" = '1' ]; then
    printf '#!/bin/sh\nexit 0\n' > "$MOCK_BIN/pwsh"
    /bin/chmod +x "$MOCK_BIN/pwsh"
fi
exit 0
MOCK

    if [[ "$case_sudo" == 'yes' ]]; then
        cat > "$case_bin/sudo" <<'MOCK'
#!/bin/sh
printf '%s\n' "$*" >> "$SUDO_LOG"
printf 'sudo stdout marker\n'
printf 'sudo stderr marker\n' >&2
if [ "$MOCK_SUDO_STATUS" -ne 0 ]; then
    exit "$MOCK_SUDO_STATUS"
fi
MOCK_PRIVILEGE=sudo "$@"
MOCK
    fi

    chmod +x "$case_bin/dirname" "$case_bin/id" "$case_bin/$case_manager"
    if [[ -f "$case_bin/sudo" ]]; then
        chmod +x "$case_bin/sudo"
    fi
}

run_installer() {
    set +e
    MOCK_UID="$case_uid" \
    MOCK_SUDO_STATUS="$case_sudo_status" \
    MOCK_PACKAGE_STATUS="$case_package_status" \
    MOCK_INSTALL_PWSH="$case_install_pwsh" \
    MOCK_MANAGER="$case_manager" \
    MOCK_BIN="$case_bin" \
    PACKAGE_LOG="$case_package_log" \
    SUDO_LOG="$case_sudo_log" \
    PATH="$case_bin" \
    "$BASH" "$installer" > "$case_stdout" 2> "$case_stderr"
    actual_status=$?
    set -e

    actual_stdout="$(file_text "$case_stdout")"
    actual_stderr="$(file_text "$case_stderr")"
}

assert_manager_cases() {
    local manager="$1"
    local package_args
    if [[ "$manager" == 'dnf' ]]; then
        package_args='install -y powershell'
    else
        package_args='-S --noconfirm powershell'
    fi

    setup_case "$manager-root-success" "$manager" 0 no 0 0 1
    run_installer
    assert_equal 0 "$actual_status" "$manager root install succeeds"
    assert_equal "direct|$package_args" "$(file_text "$case_package_log")" "$manager root invokes the exact package command once directly"
    assert_equal '' "$(file_text "$case_sudo_log")" "$manager root does not invoke sudo"
    assert_contains "$manager stdout marker" "$actual_stdout" "$manager root stdout remains visible"
    assert_equal "$manager stderr marker" "$actual_stderr" "$manager root stderr remains unchanged"

    setup_case "$manager-root-failure" "$manager" 0 no 0 42 0
    run_installer
    assert_equal 42 "$actual_status" "$manager root package failure preserves its status"
    assert_equal "direct|$package_args" "$(file_text "$case_package_log")" "$manager root failure has no second package call"
    assert_equal '' "$(file_text "$case_sudo_log")" "$manager root failure does not invoke sudo"
    assert_contains "$manager stdout marker" "$actual_stdout" "$manager root failure preserves package stdout"
    assert_equal "$manager stderr marker" "$actual_stderr" "$manager root failure preserves package stderr"

    setup_case "$manager-sudo-success" "$manager" 1000 yes 0 0 1
    run_installer
    assert_equal 0 "$actual_status" "$manager non-root sudo install succeeds"
    assert_equal "sudo|$package_args" "$(file_text "$case_package_log")" "$manager non-root package command runs once through sudo"
    assert_equal "$manager $package_args" "$(file_text "$case_sudo_log")" "$manager sudo receives the exact package arguments"
    assert_contains "$manager stdout marker" "$actual_stdout" "$manager sudo package stdout remains visible"
    assert_equal "sudo stderr marker
$manager stderr marker" "$actual_stderr" "$manager sudo and package stderr remain visible"

    setup_case "$manager-sudo-failure" "$manager" 1000 yes 31 0 0
    run_installer
    assert_equal 31 "$actual_status" "$manager sudo failure preserves the sudo status"
    assert_equal '' "$(file_text "$case_package_log")" "$manager sudo failure does not run a package command"
    assert_equal "$manager $package_args" "$(file_text "$case_sudo_log")" "$manager sudo failure makes one sudo call"
    assert_contains 'sudo stdout marker' "$actual_stdout" "$manager sudo failure preserves sudo stdout"
    assert_equal 'sudo stderr marker' "$actual_stderr" "$manager sudo failure preserves sudo stderr"

    setup_case "$manager-package-failure" "$manager" 1000 yes 0 42 0
    run_installer
    assert_equal 42 "$actual_status" "$manager package failure through sudo preserves the package status"
    assert_equal "sudo|$package_args" "$(file_text "$case_package_log")" "$manager package failure invokes the package command once through sudo"
    assert_equal "$manager $package_args" "$(file_text "$case_sudo_log")" "$manager package failure makes one sudo call"
    assert_contains 'sudo stdout marker' "$actual_stdout" "$manager package failure preserves sudo stdout"
    assert_contains "$manager stdout marker" "$actual_stdout" "$manager package failure preserves package stdout"
    assert_equal "sudo stderr marker
$manager stderr marker" "$actual_stderr" "$manager package failure preserves sudo and package stderr"

    setup_case "$manager-missing-sudo" "$manager" 1000 no 0 0 0
    run_installer
    assert_equal 127 "$actual_status" "$manager without sudo fails with status 127"
    assert_equal '' "$(file_text "$case_package_log")" "$manager without sudo does not invoke the package command"
    assert_equal '' "$(file_text "$case_sudo_log")" "$manager without sudo has no sudo invocation"
    assert_contains "sudo is required to install PowerShell with $manager when not running as root." "$actual_stderr" "$manager without sudo reports the required identity"
}

assert_manager_cases dnf
assert_manager_cases pacman

printf 'Passed %s dnf/pacman installer mock assertions.\n' "$assertion_count"
