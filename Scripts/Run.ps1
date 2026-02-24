# Ensure running as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Set location to script directory
Set-Location -Path $PSScriptRoot

# Create Folder and Copy Script
mkdir "C:\ProgramData\Winhance\Scripts" -Force
Copy-Item -Path $PSScriptRoot\PauseWindowsUpdate.ps1 -Destination "C:\ProgramData\Winhance\Scripts\PauseWindowsUpdate.ps1"
Copy-Item -Path $PSScriptRoot\PauseWindowsUpdate.xml -Destination "C:\PauseWindowsUpdate.xml"
Get-ChildItem "C:\ProgramData\Winhance\Scripts"
Get-ChildItem "C:\PauseWindowsUpdate.xml"

# Create the Task Using Command Prompt (as Administrator)
schtasks /create /tn "\Winhance\PauseWindowsUpdate" /xml "C:\PauseWindowsUpdate.xml" /f

# Pause to view results
Pause
