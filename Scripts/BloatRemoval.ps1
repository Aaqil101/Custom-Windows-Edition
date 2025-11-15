<#
  .SYNOPSIS
      Removes Windows bloatware apps, legacy capabilities, and optional features from Windows 10/11 systems.

  .DESCRIPTION
      This script removes selected Windows components including:
      - Appx packages (UWP apps like Calculator, Weather, etc.)
      - Legacy Windows capabilities
      - Optional Windows features
      - Special apps requiring custom uninstall procedures (e.g., OneNote)

      The script includes retry logic and verification to ensure complete removal.
      This script is designed to run in any context: user sessions, SYSTEM account, or scheduled tasks.

  .NOTES
      Source: https://github.com/memstechtips/Winhance

      Requirements:
      - Windows 10/11
      - Administrator privileges (script will auto-elevate)
      - PowerShell 5.1 or higher
#>

# Check if script is running as Administrator
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Try {
        Start-Process PowerShell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
        Exit
    }
    Catch {
        Write-Host "Failed to run as Administrator. Please rerun with elevated privileges."
        Exit
    }
}

# Setup logging
$logFolder = "C:\ProgramData\Winhance\Logs"
$logFile = "$logFolder\BloatRemovalLog.txt"

# Create log directory if it doesn't exist
if (!(Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

# Function to write to log file
function Write-Log {
    param (
        [string]$Message
    )
    
    # Check if log file exists and is over 500KB (512000 bytes)
    if ((Test-Path $logFile) -and (Get-Item $logFile).Length -gt 512000) {
        Remove-Item $logFile -Force -ErrorAction SilentlyContinue
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$timestamp - Log rotated - previous log exceeded 500KB" | Out-File -FilePath $logFile
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
    
    # Also output to console for real-time progress
    Write-Host $Message
}
Write-Log "Starting bloat removal process"

# Enable Remove-AppX -AllUsers compatibility aliases for this session
Write-Log "Setting up AppX compatibility aliases for this session..."
try {
    Set-Alias Get-AppPackageAutoUpdateSettings Get-AppxPackageAutoUpdateSettings -Scope Global -Force
    Set-Alias Remove-AppPackageAutoUpdateSettings Remove-AppxPackageAutoUpdateSettings -Scope Global -Force
    Set-Alias Set-AppPackageAutoUpdateSettings Set-AppxPackageAutoUpdateSettings -Scope Global -Force
    Set-Alias Reset-AppPackage Reset-AppxPackage -Scope Global -Force
    Set-Alias Add-MsixPackage Add-AppxPackage -Scope Global -Force
    Set-Alias Get-MsixPackage Get-AppxPackage -Scope Global -Force
    Set-Alias Remove-MsixPackage Remove-AppxPackage -Scope Global -Force
    Write-Log "AppX compatibility aliases created successfully"
} catch {
    Write-Log "Warning: Could not create some AppX aliases: $($_.Exception.Message)"
}

# Packages to remove
$packages = @(
    'Microsoft.MixedReality.Portal'
    'Microsoft.BingNews'
    'Microsoft.BingWeather'
    'Microsoft.GetHelp'
    'Microsoft.Windows.DevHome'
    'MicrosoftCorporationII.MicrosoftFamily'
    'Microsoft.WindowsFeedbackHub'
    'Microsoft.People'
    'Microsoft.XboxIdentityProvider'
    'Microsoft.Xbox.TCUI'
    'Microsoft.Getstarted'
    'Microsoft.Copilot'
    'Microsoft.Windows.Ai.Copilot.Provider'
    'Microsoft.Copilot_8wekyb3d8bbwe'
    'Microsoft.MicrosoftOfficeHub'
)

# Capabilities to remove
$capabilities = @(
)

# Optional Features to disable
$optionalFeatures = @(
    'Recall'
)

# Special apps requiring uninstall string execution
$specialApps = @(
)

$maxRetries = 3
$retryCount = 0

do {
    $retryCount++
    Write-Log "Standard removal attempt $retryCount of $maxRetries"

    Write-Log "Discovering all packages..."
    $allInstalledPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    $allProvisionedPackages = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue

    Write-Log "Processing packages..."
    $packagesToRemove = @()
    $provisionedPackagesToRemove = @()
    $notFoundPackages = @()

    foreach ($package in $packages) {
        $foundAny = $false

        $installedPackages = $allInstalledPackages | Where-Object { $_.Name -eq $package }
        if ($installedPackages) {
            Write-Log "Found installed package: $package"
            foreach ($pkg in $installedPackages) {
                Write-Log "Queuing installed package for removal: $($pkg.PackageFullName)"
                $packagesToRemove += $pkg.PackageFullName
            }
            $foundAny = $true
        }

        $provisionedPackages = $allProvisionedPackages | Where-Object { $_.DisplayName -eq $package }
        if ($provisionedPackages) {
            Write-Log "Found provisioned package: $package"
            foreach ($pkg in $provisionedPackages) {
                Write-Log "Queuing provisioned package for removal: $($pkg.PackageName)"
                $provisionedPackagesToRemove += $pkg.PackageName
            }
            $foundAny = $true
        }

        if (-not $foundAny) {
            $notFoundPackages += $package
        }
    }

    if ($notFoundPackages.Count -gt 0) {
        Write-Log "Packages not found: $($notFoundPackages -join ', ')"
    }

    if ($packagesToRemove.Count -gt 0) {
        Write-Log "Removing $($packagesToRemove.Count) installed packages in batch..."
        try {
            $packagesToRemove | ForEach-Object {
                Write-Log "Removing installed package: $_"
                Remove-AppxPackage -Package $_ -AllUsers -ErrorAction SilentlyContinue
            }
            Write-Log "Batch removal of installed packages completed"
        } catch {
            Write-Log "Error in batch removal of installed packages: $($_.Exception.Message)"
        }
    }

    if ($provisionedPackagesToRemove.Count -gt 0) {
        Write-Log "Removing $($provisionedPackagesToRemove.Count) provisioned packages..."
        foreach ($pkgName in $provisionedPackagesToRemove) {
            try {
                Write-Log "Removing provisioned package: $pkgName"
                Remove-AppxProvisionedPackage -Online -PackageName $pkgName -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Error removing provisioned package $pkgName : $($_.Exception.Message)"
            }
        }
        Write-Log "Provisioned packages removal completed"
    }

    Write-Log "Processing capabilities..."
    foreach ($capability in $capabilities) {
        Write-Log "Checking capability: $capability"
        try {
            $matchingCapabilities = Get-WindowsCapability -Online | Where-Object { $_.Name -like "$capability*" -or $_.Name -like "$capability~~~~*" }

            if ($matchingCapabilities) {
                $foundInstalled = $false
                foreach ($existingCapability in $matchingCapabilities) {
                    if ($existingCapability.State -eq "Installed") {
                        $foundInstalled = $true
                        Write-Log "Removing capability: $($existingCapability.Name)"
                        Remove-WindowsCapability -Online -Name $existingCapability.Name -ErrorAction SilentlyContinue | Out-Null
                    }
                }

                if (-not $foundInstalled) {
                    Write-Log "Found capability $capability but it is not installed"
                }
            }
            else {
                Write-Log "No matching capabilities found for: $capability"
            }
        }
        catch {
            Write-Log "Error checking capability: $capability - $($_.Exception.Message)"
        }
    }

    Write-Log "Processing optional features..."
    if ($optionalFeatures.Count -gt 0) {
        $enabledFeatures = @()
        foreach ($feature in $optionalFeatures) {
            Write-Log "Checking feature: $feature"
            $existingFeature = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction SilentlyContinue
            if ($existingFeature -and $existingFeature.State -eq "Enabled") {
                $enabledFeatures += $feature
            } else {
                Write-Log "Feature not found or not enabled: $feature"
            }
        }

        if ($enabledFeatures.Count -gt 0) {
            Write-Log "Disabling features: $($enabledFeatures -join ', ')"
            Disable-WindowsOptionalFeature -Online -FeatureName $enabledFeatures -NoRestart -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Write-Log "Verifying removal results..."
    $remainingItems = @()

    $currentPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    foreach ($package in $packages) {
        if ($currentPackages | Where-Object { $_.Name -eq $package }) {
            $remainingItems += $package
            Write-Log "Package still installed: $package"
        }
    }

    $currentCapabilities = Get-WindowsCapability -Online -ErrorAction SilentlyContinue | Where-Object State -eq 'Installed'
    foreach ($capability in $capabilities) {
        if ($currentCapabilities | Where-Object { $_.Name -like "$capability*" }) {
            $remainingItems += $capability
            Write-Log "Capability still installed: $capability"
        }
    }

    $currentFeatures = Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object State -eq 'Enabled'
    foreach ($feature in $optionalFeatures) {
        if ($currentFeatures | Where-Object { $_.FeatureName -eq $feature }) {
            $remainingItems += $feature
            Write-Log "Feature still enabled: $feature"
        }
    }

    if ($remainingItems.Count -eq 0) {
        Write-Log "All standard items successfully removed!"
        break
    } else {
        Write-Log "Retry needed. $($remainingItems.Count) items remain: $($remainingItems -join ', ')"
        if ($retryCount -lt $maxRetries) {
            Write-Log "Waiting 2 seconds before retry..."
            Start-Sleep -Seconds 2
        }
    }

} while ($retryCount -lt $maxRetries -and $remainingItems.Count -gt 0)

if ($remainingItems.Count -gt 0) {
    Write-Log "Warning: $($remainingItems.Count) standard items could not be removed after $maxRetries attempts: $($remainingItems -join ', ')"
}

if ($specialApps.Count -gt 0) {
    Write-Log "Processing special apps that require custom uninstall procedures..."

    $maxSpecialRetries = 2
    $specialRetryCount = 0
    $specialAppsRemaining = @()

    do {
        $specialRetryCount++
        if ($specialRetryCount -gt 1) {
            Write-Log "Special apps retry attempt $specialRetryCount of $maxSpecialRetries"
        }

        $uninstallBasePaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
        )

        foreach ($specialApp in $specialApps) {
            Write-Log "Processing special app: $specialApp"

            switch ($specialApp) {
                'OneNote' {
                    $processesToStop = @('OneNote', 'ONENOTE', 'ONENOTEM')
                    $searchPattern = 'OneNote*'
                    $packagePattern = '*OneNote*'
                }
                default {
                    Write-Log "Unknown or unsupported special app: $specialApp"
                    continue
                }
            }

            foreach ($processName in $processesToStop) {
                $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
                if ($processes) {
                    $processes | Stop-Process -Force -ErrorAction SilentlyContinue
                    Write-Log "Stopped process: $processName"
                }
            }

            $uninstallExecuted = $false
            foreach ($uninstallBasePath in $uninstallBasePaths) {
                try {
                    Write-Log "Searching for $searchPattern in $uninstallBasePath"
                    $uninstallKeys = Get-ChildItem -Path $uninstallBasePath -ErrorAction SilentlyContinue |
                                    Where-Object { $_.PSChildName -like $searchPattern }

                    foreach ($key in $uninstallKeys) {
                        try {
                            $uninstallString = (Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue).UninstallString
                            if ($uninstallString) {
                                Write-Log "Found uninstall string: $uninstallString"

                                if ($uninstallString -match '^\"([^\"]+)\"(.*)$') {
                                    $exePath = $matches[1]
                                    $args = $matches[2].Trim()

                                    if ($exePath -like '*OfficeClickToRun.exe') {
                                        $args += ' DisplayLevel=False'
                                    } else {
                                        $args += ' /silent'
                                    }

                                    Write-Log "Executing: $exePath with args: $args"
                                    Start-Process -FilePath $exePath -ArgumentList $args -NoNewWindow -Wait -ErrorAction SilentlyContinue
                                } else {
                                    if ($uninstallString -like '*OfficeClickToRun.exe*') {
                                        Start-Process -FilePath $uninstallString -ArgumentList 'DisplayLevel=False' -NoNewWindow -Wait -ErrorAction SilentlyContinue
                                    } else {
                                        Start-Process -FilePath $uninstallString -ArgumentList '/silent' -NoNewWindow -Wait -ErrorAction SilentlyContinue
                                    }
                                }

                                $uninstallExecuted = $true
                                Write-Log "Completed uninstall execution for $specialApp"
                            }
                        }
                        catch {
                            Write-Log "Error processing uninstall key: $($_.Exception.Message)"
                        }
                    }
                }
                catch {
                    Write-Log "Error searching for uninstall keys: $($_.Exception.Message)"
                }
            }

            if (-not $uninstallExecuted) {
                Write-Log "No uninstall strings found for $specialApp"
            }
        }

        if ($specialRetryCount -eq 1) {
            Write-Log "Waiting 3 seconds for uninstallers to complete..."
            Start-Sleep -Seconds 3
        }

        Write-Log "Verifying special apps removal..."
        $specialAppsRemaining = @()

        foreach ($specialApp in $specialApps) {
            $stillExists = $false

            switch ($specialApp) {
                'OneNote' {
                    $appxPackage = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
                                   Where-Object { $_.Name -like '*OneNote*' }
                    if ($appxPackage) {
                        $stillExists = $true
                        Write-Log "OneNote AppxPackage still exists: $($appxPackage.PackageFullName)"
                    }

                    $uninstallKeys = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' -ErrorAction SilentlyContinue |
                                     Where-Object { $_.PSChildName -like 'OneNote*' }
                    if ($uninstallKeys) {
                        $stillExists = $true
                        Write-Log "OneNote registry uninstall keys still exist"
                    }
                }
            }

            if ($stillExists) {
                $specialAppsRemaining += $specialApp
            }
        }

        if ($specialAppsRemaining.Count -eq 0) {
            Write-Log "All special apps successfully removed!"
            break
        } else {
            Write-Log "$($specialAppsRemaining.Count) special apps remain: $($specialAppsRemaining -join ', ')"
            if ($specialRetryCount -lt $maxSpecialRetries) {
                Write-Log "Waiting 3 seconds before retry..."
                Start-Sleep -Seconds 3
            }
        }

    } while ($specialRetryCount -lt $maxSpecialRetries -and $specialAppsRemaining.Count -gt 0)

    if ($specialAppsRemaining.Count -gt 0) {
        Write-Log "Warning: $($specialAppsRemaining.Count) special apps could not be removed after $maxSpecialRetries attempts: $($specialAppsRemaining -join ', ')"
    }
}

Write-Log "Bloat removal process completed"
