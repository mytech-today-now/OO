#!/bin/bash
# ============================================================================
#  myTech.Today - O&O ShutUp10 Installer Setup
#  Note: O&O ShutUp10 is Windows-only, but this script installs PowerShell 7
# ============================================================================

set -e

echo ""
echo "============================================================"
echo "  myTech.Today - O&O ShutUp10 Installer Setup"
echo "============================================================"
echo ""

echo "[WARN] O&O ShutUp10 is a Windows-only application."
echo "       This script is provided for PowerShell 7 installation only."
echo ""

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
command_exists() { command -v "$1" >/dev/null 2>&1; }

if command_exists pwsh; then
    echo "[OK] PowerShell 7+ is installed"
else
    echo "[INFO] PowerShell 7+ not found. Installing..."
    
    if [[ "$OSTYPE" == "darwin"* ]]; then
        command_exists brew && brew install --cask powershell || { echo "[ERROR] Install Homebrew first: https://brew.sh"; exit 1; }
    elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
        if command_exists apt-get; then
            if [ "$(id -u)" -eq 0 ]; then
                apt_prefix=()
            else
                if ! command_exists sudo; then
                    echo "[ERROR] sudo is required for both apt-get update and apt-get install when not running as root." >&2
                    exit 127
                fi
                apt_prefix=(sudo)
            fi

            if "${apt_prefix[@]}" apt-get update; then
                :
            else
                status=$?
                echo "[ERROR] apt-get update failed (exit ${status})." >&2
                echo "The apt-get operation failed at the recorded stage. Verify package state and run the next required command with sudo." >&2
                exit "$status"
            fi

            if "${apt_prefix[@]}" apt-get install -y powershell; then
                :
            else
                status=$?
                echo "[ERROR] apt-get install failed (exit ${status})." >&2
                echo "The apt-get operation failed at the recorded stage. Verify package state and run the next required command with sudo." >&2
                exit "$status"
            fi
        elif command_exists dnf; then
            if [ "$(id -u)" -eq 0 ]; then
                if dnf install -y powershell; then
                    :
                else
                    status=$?
                    exit "$status"
                fi
            else
                if ! command_exists sudo; then
                    echo "[ERROR] sudo is required to install PowerShell with dnf when not running as root." >&2
                    exit 127
                fi

                if sudo dnf install -y powershell; then
                    :
                else
                    status=$?
                    exit "$status"
                fi
            fi
        elif command_exists pacman; then
            if [ "$(id -u)" -eq 0 ]; then
                if pacman -S --noconfirm powershell; then
                    :
                else
                    status=$?
                    exit "$status"
                fi
            else
                if ! command_exists sudo; then
                    echo "[ERROR] sudo is required to install PowerShell with pacman when not running as root." >&2
                    exit 127
                fi

                if sudo pacman -S --noconfirm powershell; then
                    :
                else
                    status=$?
                    exit "$status"
                fi
            fi
        else
            echo "[ERROR] Could not detect package manager."; exit 1
        fi
    else
        echo "[ERROR] Unsupported OS: $OSTYPE"; exit 1
    fi
    
    command_exists pwsh || { echo "[ERROR] PowerShell installation failed."; exit 1; }
    echo "[OK] PowerShell 7 installed successfully!"
fi

echo ""
echo "[INFO] PowerShell 7 is ready. Run this on Windows for O&O ShutUp10."

