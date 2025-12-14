# Winget Installer Updater with Version Check
# Checks for latest versions and updates installers if needed

# Set your target folder path here
$TargetFolder = "D:\ISO FIles\Custom ISO Files\SetupFiles\Software"

# Create folder if it doesn't exist
if (!(Test-Path $TargetFolder)) {
    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
}

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Winget Installer Updater" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host "Target: $TargetFolder`n" -ForegroundColor Yellow

# Check if Winget is available
try {
    $null = winget --version
    Write-Host "[OK] Winget is available`n" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Winget is not installed!" -ForegroundColor Red
    Write-Host "Install from: https://aka.ms/getwinget" -ForegroundColor Yellow
    exit
}

# Define software to download
$Apps = @(
    @{
        Name = "WinRAR";
        ID = "RARLab.WinRAR";
        WantMSI = $false
    },
    @{
        Name = "VLC media player";
        ID = "VideoLAN.VLC";
        WantMSI = $true;
        MSIUrlTemplate = "https://download.videolan.org/pub/videolan/vlc/{version}/win64/vlc-{version}-win64.msi"
    },
    @{
        Name = "QuickLook";
        ID = "QL-Win.QuickLook";
        WantMSI = $true
    },
    @{
        Name = "K-Lite Codec Pack Mega";
        ID = "CodecGuide.K-LiteCodecPack.Mega";
        WantMSI = $false
    },
    @{
        Name = "Google Chrome";
        ID = "Google.Chrome";
        WantMSI = $true
    },
    @{
        Name = "File Converter";
        ID = "AdrienAllard.FileConverter";
        WantMSI = $true
    },
    @{
        Name = "Everything";
        ID = "voidtools.Everything";
        WantMSI = $false
    },
    @{
        Name = "AnyDesk";
        ID = "AnyDesk.AnyDesk";
        WantMSI = $false
    },
    @{
        Name = "7-Zip";
        ID = "7zip.7zip";
        WantMSI = $true;
        MSIUrlTemplate = "https://github.com/ip7z/7zip/releases/download/{version}/7z{versionshort}-x64.msi"
    }
)

# Function to get latest version from Winget
function Get-LatestVersion {
    param($PackageID)

    try {
        $info = winget show --id $PackageID --accept-source-agreements 2>&1 | Out-String
        if ($info -match "Version:\s+(\d+\.[\d\.]+)") {
            return $matches[1].Trim()
        }
    }
    catch {
        return $null
    }
    return $null
}

# Function to extract version from filename
function Get-FileVersion {
    param($FileName)

    if ($FileName -match "(\d+\.[\d\.]+)") {
        return $matches[1]
    }
    return $null
}

# Function to get version from installer metadata
function Get-InstallerVersion {
    param($FilePath)

    try {
        # Use VersionInfo property for both EXE and MSI files
        $fileInfo = Get-Item $FilePath
        $versionInfo = $fileInfo.VersionInfo

        # Try ProductVersion first, then FileVersion
        $version = $null
        if ($versionInfo.ProductVersion) {
            $version = $versionInfo.ProductVersion.Trim()
        }
        elseif ($versionInfo.FileVersion) {
            $version = $versionInfo.FileVersion.Trim()
        }

        # Extract version number from string
        if ($version -and $version -match "(\d+\.[\d\.]+)") {
            return $matches[1]
        }

        return $version
    }
    catch {
        # If metadata extraction fails, return null
        return $null
    }
}

# Function to normalize version (remove trailing zeros)
function Normalize-Version {
    param($Version)

    if (-not $Version) { return $null }

    # Remove any whitespace
    $Version = $Version.Trim()

    # Remove leading/trailing whitespace that might be in the string
    if ([string]::IsNullOrWhiteSpace($Version)) {
        return $null
    }

    try {
        # Convert to version object and back to remove trailing zeros
        $v = [version]$Version
        # Rebuild version string without trailing zeros
        $parts = @($v.Major, $v.Minor)
        if ($v.Build -gt 0 -or $v.Revision -gt 0) {
            $parts += $v.Build
        }
        if ($v.Revision -gt 0) {
            $parts += $v.Revision
        }
        return ($parts -join '.')
    }
    catch {
        # If conversion fails, return as-is
        return $Version
    }
}

# Function to compare versions
function Compare-Versions {
    param($Version1, $Version2)

    # Check for null or empty versions
    if ([string]::IsNullOrWhiteSpace($Version1) -or [string]::IsNullOrWhiteSpace($Version2)) {
        return $null
    }

    # Normalize both versions
    $v1Normalized = Normalize-Version -Version $Version1
    $v2Normalized = Normalize-Version -Version $Version2

    if (-not $v1Normalized -or -not $v2Normalized) {
        return $null
    }

    try {
        $v1 = [version]$v1Normalized
        $v2 = [version]$v2Normalized

        if ($v1 -gt $v2) { return 1 }
        elseif ($v1 -lt $v2) { return -1 }
        else { return 0 }
    }
    catch {
        # If version comparison fails, return null
        return $null
    }
}

# Counters
$updated = 0
$skipped = 0
$failed = 0

# Process each app
foreach ($app in $Apps) {
    Write-Host "Checking: $($app.Name)" -ForegroundColor Cyan

    try {
        # Get latest available version
        $latestVersion = Get-LatestVersion -PackageID $app.ID

        if ($latestVersion) {
            Write-Host "  Latest version: $latestVersion" -ForegroundColor Yellow
        }
        else {
            Write-Host "  [WARN] Could not determine latest version" -ForegroundColor Yellow
        }

        # Look for existing installer in target folder
        $existingFiles = Get-ChildItem -Path $TargetFolder -File | Where-Object {
            $_.Name -match [regex]::Escape($app.Name) -or
            $_.Name -match ($app.ID -split '\.' | Select-Object -Last 1)
        }

        $needsUpdate = $true
        $oldFileName = $null
        $oldFile = $null

        if ($existingFiles) {
            $existingFile = $existingFiles[0]
            $oldFileName = $existingFile.Name
            $oldFile = $existingFile

            Write-Host "  Existing file: $oldFileName" -ForegroundColor Gray

            # Try to get version from filename first
            $existingVersion = Get-FileVersion -FileName $existingFile.Name

            # If no version in filename, try to extract from file metadata
            if (-not $existingVersion) {
                Write-Host "  Reading file metadata..." -ForegroundColor Gray
                $existingVersion = Get-InstallerVersion -FilePath $existingFile.FullName
            }

            # Ensure version is properly trimmed (critical fix for whitespace issues)
            if ($existingVersion) {
                $existingVersion = $existingVersion.ToString().Trim()
            }

            if ($existingVersion -and $latestVersion) {
                Write-Host "  Installed version: $existingVersion" -ForegroundColor Gray

                $comparison = Compare-Versions -Version1 $existingVersion -Version2 $latestVersion

                if ($null -eq $comparison) {
                    Write-Host "  [WARN] Cannot compare versions, keeping existing" -ForegroundColor Yellow
                    $needsUpdate = $false
                    $skipped++
                }
                elseif ($comparison -eq 0) {
                    Write-Host "  [OK] Already up to date!" -ForegroundColor Green
                    $needsUpdate = $false
                    $skipped++
                }
                elseif ($comparison -gt 0) {
                    Write-Host "  [OK] Installed version is newer!" -ForegroundColor Green
                    $needsUpdate = $false
                    $skipped++
                }
                else {
                    Write-Host "  [UPDATE] Newer version available ($existingVersion -> $latestVersion)" -ForegroundColor Yellow
                }
            }
            elseif ($existingVersion) {
                # Have existing version but no latest version info
                Write-Host "  Installed version: $existingVersion" -ForegroundColor Gray
                Write-Host "  [OK] Cannot verify latest version, keeping existing" -ForegroundColor Yellow
                $needsUpdate = $false
                $skipped++
            }
            else {
                # No version in filename or metadata - check file age instead
                $fileAge = (Get-Date) - $existingFile.LastWriteTime

                if ($fileAge.TotalDays -lt 7) {
                    Write-Host "  [OK] File is recent (modified $([math]::Round($fileAge.TotalDays, 1)) days ago)" -ForegroundColor Green
                    $needsUpdate = $false
                    $skipped++
                }
                else {
                    Write-Host "  [UPDATE] File is old (modified $([math]::Round($fileAge.TotalDays, 0)) days ago)" -ForegroundColor Yellow
                }
            }
        }
        else {
            Write-Host "  [NEW] No existing installer found" -ForegroundColor Yellow
        }

        # Download if needed
        if ($needsUpdate) {
            # Create temp directory
            $tempDir = Join-Path $env:TEMP "winget_$([guid]::NewGuid())"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

            Write-Host "  Downloading..." -ForegroundColor Cyan
            Write-Host ""

            # Check if direct MSI download is preferred
            if ($app.MSIUrlTemplate -and $latestVersion) {
                # Build the MSI URL dynamically using the latest version
                $msiUrl = $app.MSIUrlTemplate -replace "\{version\}", $latestVersion

                # Handle special version format (e.g., 25.01 -> 2501 for 7-Zip)
                if ($msiUrl -match "\{versionshort\}") {
                    $versionShort = $latestVersion -replace '\.', ''
                    $msiUrl = $msiUrl -replace "\{versionshort\}", $versionShort
                }

                Write-Host "  Downloading MSI directly from: $msiUrl" -ForegroundColor Cyan

                try {
                    $msiFileName = [System.IO.Path]::GetFileName($msiUrl)
                    $msiPath = Join-Path $tempDir $msiFileName

                    # Download using Invoke-WebRequest with progress
                    $ProgressPreference = 'SilentlyContinue'
                    Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing
                    $ProgressPreference = 'Continue'

                    Write-Host "  [OK] MSI downloaded successfully" -ForegroundColor Green
                    $installer = Get-Item $msiPath
                }
                catch {
                    Write-Host "  [WARN] Direct MSI download failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Host "  Falling back to Winget..." -ForegroundColor Yellow
                    Write-Host ""

                    # Fall back to Winget
                    winget download --id $($app.ID) --download-directory $tempDir --accept-source-agreements --accept-package-agreements

                    # Find the installer file based on WantMSI preference
                    if ($app.WantMSI) {
                        $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.msi | Select-Object -First 1
                        if (-not $installer) {
                            $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.exe | Select-Object -First 1
                        }
                    }
                    else {
                        $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.exe | Select-Object -First 1
                        if (-not $installer) {
                            $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.msi | Select-Object -First 1
                        }
                    }
                }
            }
            else {
                # Download using winget (show output)
                winget download --id $($app.ID) --download-directory $tempDir --accept-source-agreements --accept-package-agreements

                Write-Host ""

                # Find the installer file based on WantMSI preference
                if ($app.WantMSI) {
                    $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.msi | Select-Object -First 1
                    if (-not $installer) {
                        Write-Host "  [WARN] MSI not found, using EXE instead" -ForegroundColor Yellow
                        $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.exe | Select-Object -First 1
                    }
                }
                else {
                    $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.exe | Select-Object -First 1
                    if (-not $installer) {
                        Write-Host "  [WARN] EXE not found, using MSI instead" -ForegroundColor Yellow
                        $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.msi | Select-Object -First 1
                    }
                }
            }

            Write-Host ""

            if ($installer) {
                # Determine final filename using the app Name
                $extension = $installer.Extension
                $finalFileName = "$($app.Name)$extension"
                $destination = Join-Path $TargetFolder $finalFileName

                # Remove old file if it exists and is different from destination
                if ($oldFile -and (Test-Path $oldFile.FullName)) {
                    Remove-Item $oldFile.FullName -Force
                }

                # Copy installer to target folder
                Copy-Item $installer.FullName $destination -Force

                $sizeMB = [math]::Round((Get-Item $destination).Length / 1MB, 2)
                Write-Host "  [OK] Downloaded: $finalFileName ($sizeMB MB)" -ForegroundColor Green
                $updated++
            }
            else {
                Write-Host "  [FAIL] No installer found after download" -ForegroundColor Red
                $failed++
            }

            # Cleanup temp directory
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host "  [FAIL] Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Debug: Version1='$existingVersion', Version2='$latestVersion'" -ForegroundColor DarkGray
        $failed++
    }

    Write-Host ""
}

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Update Complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Updated: $updated" -ForegroundColor Green
Write-Host "  Already up to date: $skipped" -ForegroundColor Yellow
Write-Host "  Failed: $failed" -ForegroundColor Red
Write-Host ""
Write-Host "Files location: $TargetFolder" -ForegroundColor Yellow
Write-Host "================================" -ForegroundColor Cyan
