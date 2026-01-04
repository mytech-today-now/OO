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
            [ "$EUID" -ne 0 ] && sudo apt-get update && sudo apt-get install -y powershell || apt-get update && apt-get install -y powershell
        elif command_exists dnf; then
            [ "$EUID" -ne 0 ] && sudo dnf install -y powershell || dnf install -y powershell
        elif command_exists pacman; then
            [ "$EUID" -ne 0 ] && sudo pacman -S --noconfirm powershell || pacman -S --noconfirm powershell
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

