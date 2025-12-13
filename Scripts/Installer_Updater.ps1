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
    @{ Name = "WinRAR"; ID = "RARLab.WinRAR" },
    @{ Name = "VLC"; ID = "VideoLAN.VLC" },
    @{ Name = "QuickLook"; ID = "QL-Win.QuickLook" },
    @{ Name = "K-Lite Codec Pack Mega"; ID = "CodecGuide.K-LiteCodecPack.Mega" },
    @{ Name = "Google Chrome"; ID = "Google.Chrome" },
    @{ Name = "File Converter"; ID = "AdrienAllard.FileConverter" },
    @{ Name = "Everything"; ID = "voidtools.Everything" },
    @{ Name = "AnyDesk"; ID = "AnyDeskSoftwareGmbH.AnyDesk" },
    @{ Name = "7-Zip"; ID = "7zip.7zip" }
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

# Function to compare versions
function Compare-Versions {
    param($Version1, $Version2)

    try {
        $v1 = [version]$Version1
        $v2 = [version]$Version2

        if ($v1 -gt $v2) { return 1 }
        elseif ($v1 -lt $v2) { return -1 }
        else { return 0 }
    }
    catch {
        return -1  # If comparison fails, assume update needed
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

        if ($existingFiles) {
            $existingFile = $existingFiles[0]
            $oldFileName = $existingFile.Name
            $existingVersion = Get-FileVersion -FileName $existingFile.Name

            Write-Host "  Existing file: $oldFileName" -ForegroundColor Gray

            if ($existingVersion -and $latestVersion) {
                Write-Host "  Installed version: $existingVersion" -ForegroundColor Gray

                $comparison = Compare-Versions -Version1 $existingVersion -Version2 $latestVersion

                if ($comparison -ge 0) {
                    Write-Host "  [OK] Already up to date!" -ForegroundColor Green
                    $needsUpdate = $false
                    $skipped++
                }
                else {
                    Write-Host "  [UPDATE] Newer version available" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "  [UPDATE] Downloading latest version" -ForegroundColor Yellow
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

            # Download using winget (show output)
            winget download --id $($app.ID) --download-directory $tempDir --accept-source-agreements --accept-package-agreements

            Write-Host ""

            # Find the installer file
            $installer = Get-ChildItem -Path $tempDir -Recurse -Include *.exe, *.msi | Select-Object -First 1

            if ($installer) {
                # Determine final filename
                if ($oldFileName) {
                    # Use the old filename to maintain naming convention
                    $finalFileName = $oldFileName
                }
                else {
                    # Use the downloaded filename
                    $finalFileName = $installer.Name
                }

                $destination = Join-Path $TargetFolder $finalFileName

                # Remove old file if it exists
                if (Test-Path $destination) {
                    Remove-Item $destination -Force
                }

                # Move new file to destination
                Move-Item $installer.FullName $destination -Force

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
