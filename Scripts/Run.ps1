# Ensure running as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Create Folder and Copy Script
mkdir "C:\ProgramData\Winhance\Scripts"
Copy-Item -Path "$PSScriptRoot\BloatRemoval.ps1" -Destination "C:\ProgramData\Winhance\Scripts\BloatRemoval.ps1"
Copy-Item -Path "$PSScriptRoot\PauseWindowsUpdate.ps1" -Destination "C:\ProgramData\Winhance\Scripts\PauseWindowsUpdate.ps1"
Get-ChildItem "C:\ProgramData\Winhance\Scripts"

# Create the Task Using Command Prompt (as Administrator)
schtasks /create /tn "\Winhance\BloatRemoval" /xml "$PSScriptRoot\BloatRemoval.xml" /f
schtasks /create /tn "\Winhance\PauseWindowsUpdate" /xml "$PSScriptRoot\PauseWindowsUpdate.xml" /f

# Install Applications
powershell -ExecutionPolicy Bypass -File "$PSScriptRoot\Program_Install.ps1" -InstallersPath "$PSScriptRoot"

# Extract QuickLook Plugins
powershell -ExecutionPolicy Bypass -File "$PSScriptRoot\Install_QuickLookPlugins.ps1"

# Auto Start QuickLook
& "$PSScriptRoot\QuickLookStartup.cmd"

# Pause to view results
Pause
