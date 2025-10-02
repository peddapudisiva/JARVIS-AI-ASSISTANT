# Create Desktop and Startup shortcuts to run Jarvis via run_jarvis.ps1
# Usage:
#   Right-click > Run with PowerShell
#   or from PS: powershell -ExecutionPolicy Bypass -File .\install_shortcuts.ps1

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ErrorActionPreference = 'Stop'

# Paths
$scriptDir   = $PSScriptRoot
$runner      = Join-Path $scriptDir 'run_jarvis.ps1'
$desktop     = [Environment]::GetFolderPath('Desktop')
$startup     = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$ws          = New-Object -ComObject WScript.Shell

function New-Shortcut($path) {
    $sc = $ws.CreateShortcut($path)
    $sc.TargetPath  = 'powershell.exe'
    $sc.Arguments   = "-ExecutionPolicy Bypass -File `"$runner`""
    $sc.WorkingDirectory = $scriptDir
    # try to use python icon from venv; fallback to powershell icon
    $venvPy = Join-Path $scriptDir '.venv\Scripts\python.exe'
    if (Test-Path $venvPy) {
        $sc.IconLocation = "$venvPy,0"
    } else {
        $sc.IconLocation = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe,0"
    }
    $sc.Description = 'Launch Jarvis Assistant'
    $sc.Save()
}

# Create Desktop shortcut
$desktopLnk = Join-Path $desktop 'Jarvis.lnk'
New-Shortcut -path $desktopLnk
Write-Host "Created Desktop shortcut: $desktopLnk"

# Create Startup shortcut
$startupLnk = Join-Path $startup 'Jarvis.lnk'
New-Shortcut -path $startupLnk
Write-Host "Created Startup shortcut: $startupLnk (Jarvis will start with Windows)"
