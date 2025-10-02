# Run Jarvis: create venv if missing, activate, install deps, and start
# Usage: Right-click this file > Run with PowerShell (or run from a PS prompt)

# 1) Allow scripts for this session only (safe, temporary)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# 2) Move to script directory
Set-Location -Path $PSScriptRoot

# 3) Create venv if needed
if (-not (Test-Path ".venv/Scripts/python.exe")) {
    Write-Host "Creating virtual environment..."
    py -3 -m venv .venv
}

# 4) Activate venv
. .\.venv\Scripts\Activate.ps1

# 5) Install dependencies
Write-Host "Installing dependencies from requirements.txt..."
pip install -r requirements.txt

# 6) Launch Jarvis
Write-Host "Starting Jarvis..."
python jarvis.py
