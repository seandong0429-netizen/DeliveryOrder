#!/bin/bash
cd "$(dirname "$0")"

echo "========================================"
echo "   DeliveryOrder FINAL Build Script"
echo "   Strategy: Clean Venv + No Cache + Explicit Target"
echo "========================================"

# 1. Cleanup Caches (Nuclear Option)
echo "[1/5] Nuking caches..."
rm -rf build dist venv_build
echo "Cleaning PyInstaller User Cache..."
rm -rf ~/Library/Application\ Support/pyinstaller

# 2. Setup Isolated Venv
echo "[2/5] Creating clean Virtual Environment..."
# Use system python3 to avoid Homebrew leakage if possible, or whatever is in path if verified.
# The user asked to check for official python. /usr/bin/python3 is generally safe.
# 2. Setup Isolated Venv
echo "[2/5] Creating clean Virtual Environment..."
# CRITICAL FIX: Use Homebrew Python to bypass System/SDK version mismatch
PY_BIN="/opt/homebrew/bin/python3"

echo "Using Python: $PY_BIN"
"$PY_BIN" -m venv venv_build
source venv_build/bin/activate

# 3. Install Dependencies
echo "[3/5] Installing Dependencies in Venv..."
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

# 4. Set Environment Variables
# Using Homebrew Python allows us to target a standard, stable macOS version.
export MACOSX_DEPLOYMENT_TARGET=13.0
echo "MACOSX_DEPLOYMENT_TARGET set to: 13.0 (Standard Compatibility)"

# 5. Build
echo "[4/5] Running PyInstaller..."
# We use the DeliveryOrderGen.spec file which allows customization (arm64 etc)
pyinstaller --clean --noconfirm DeliveryOrderGen.spec

echo "========================================"
echo "   Build Complete!"
echo "========================================"
echo "File located at: dist/DeliveryOrderGen.app"

# Open folder
open dist

# Keep terminal open to see errors
echo "Process finished. You can close this window."
read -p "Press Enter to exit..."
