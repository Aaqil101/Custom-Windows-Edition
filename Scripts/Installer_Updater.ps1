# Software Installer Updater Script
# Uses Winget first, falls back to Chocolatey, downloads latest installers

# Set your target folder path here
$TargetFolder = "D:\ISO FIles\Custom ISO Files\SetupFiles\Software"  # Change this to your folder path

# Create folder if it doesn't exist
if (!(Test-Path $TargetFolder)) {
    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Software Installer Updater" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Target folder: $TargetFolder`n" -ForegroundColor Yellow

# Check if Winget is installed
Write-Host "Checking for Winget..." -ForegroundColor Cyan
$wingetInstalled = $false
try {
    $null = winget --version
    $wingetInstalled = $true
    Write-Host "✓ Winget is installed`n" -ForegroundColor Green
}
catch {
    Write-Host "✗ Winget is NOT installed" -ForegroundColor Red
    Write-Host "  Install from: https://aka.ms/getwinget`n" -ForegroundColor Yellow
}

# Check if Chocolatey is installed
Write-Host "Checking for Chocolatey..." -ForegroundColor Cyan
$chocoInstalled = $false
try {
    $null = choco --version
    $chocoInstalled = $true
    Write-Host "✓ Chocolatey is installed`n" -ForegroundColor Green
}
catch {
    Write-Host "✗ Chocolatey is NOT installed" -ForegroundColor Red
    Write-Host "  Install with: Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))`n" -ForegroundColor Yellow
}

if (!$wingetInstalled -and !$chocoInstalled) {
    Write-Host "ERROR: Neither Winget nor Chocolatey is installed. Please install at least one." -ForegroundColor Red
    exit
}

Write-Host "========================================`n" -ForegroundColor Cyan

# Define software with Winget and Chocolatey package names
$Software = @(
    @{
        Name = "WinRAR"
        WingetID = "RARLab.WinRAR"
        ChocoID = "winrar"
        DesiredName = "winrar"
    },
    @{
        Name = "VLC"
        WingetID = "VideoLAN.VLC"
        ChocoID = "vlc"
        DesiredName = "vlc-{version}-win64"
        DynamicVersion = $true
    },
    @{
        Name = "QuickLook"
        WingetID = "QL-Win.QuickLook"
        ChocoID = "quicklook"
        DesiredName = "QuickLook"
    },
    @{
        Name = "K-Lite Codec Pack Mega"
        WingetID = "CodecGuide.K-LiteCodecPack.Mega"
        ChocoID = "k-litecodecpackmega"
        DesiredName = "K-Lite_Codec_Pack_Mega"
    },
    @{
        Name = "Google Chrome"
        WingetID = "Google.Chrome"
        ChocoID = "googlechrome"
        DesiredName = "googlechromestandaloneenterprise64"
    },
    @{
        Name = "File Converter"
        WingetID = "AdrienAllard.FileConverter"
        ChocoID = "file-converter"
        DesiredName = "FileConverter"
    },
    @{
        Name = "Everything"
        WingetID = "voidtools.Everything"
        ChocoID = "everything"
        DesiredName = "Everything"
    },
    @{
        Name = "AnyDesk"
        WingetID = "AnyDeskSoftwareGmbH.AnyDesk"
        ChocoID = "anydesk"
        DesiredName = "AnyDesk"
    },
    @{
        Name = "7-Zip"
        WingetID = "7zip.7zip"
        ChocoID = "7zip"
        DesiredName = "7-zip"
    }
)

# Function to download with Winget
function Download-WithWinget {
    param($ID, $Name, $Destination, $DynamicVersion)

    Write-Host "  Trying Winget..." -ForegroundColor Yellow
    try {
        $tempDir = Join-Path $env:TEMP "winget_download_$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        # Get version info if dynamic versioning is needed
        $version = ""
        if ($DynamicVersion) {
            $wingetInfo = winget show --id $ID --accept-source-agreements 2>&1 | Out-String
            if ($wingetInfo -match "Version:\s+(.+)") {
                $version = $matches[1].Trim()
                Write-Host "  Detected version: $version" -ForegroundColor Cyan
            }
        }

        Write-Host ""
        winget download --id $ID --download-directory $tempDir --accept-source-agreements --accept-package-agreements
        Write-Host ""

        # Find the downloaded installer
        $installers = Get-ChildItem -Path $tempDir -Recurse -Include *.exe, *.msi, *.msix, *.appx

        if ($installers.Count -gt 0) {
            $installer = $installers[0]
            $extension = $installer.Extension

            # Replace {version} placeholder if present
            $finalName = $Name
            if ($DynamicVersion -and $version -and $Name -match "\{version\}") {
                $finalName = $Name -replace "\{version\}", $version
            }

            $destPath = Join-Path $Destination "$finalName$extension"

            # Remove old file if exists
            if (Test-Path $destPath) {
                Remove-Item $destPath -Force
            }

            Move-Item $installer.FullName $destPath -Force

            # Cleanup temp directory
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

            # Clean up any YAML files in target directory
            Get-ChildItem -Path $Destination -Filter "*.yaml" | Remove-Item -Force -ErrorAction SilentlyContinue

            $fileSize = (Get-Item $destPath).Length / 1MB
            Write-Host "  ✓ Downloaded with Winget ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
            return $true
        }
        else {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }
    }
    catch {
        Write-Host "  ✗ Winget failed: $_" -ForegroundColor Red
        return $false
    }
}

# Function to download with Chocolatey
function Download-WithChocolatey {
    param($ID, $Name, $Destination, $DynamicVersion)

    Write-Host "  Trying Chocolatey..." -ForegroundColor Yellow
    try {
        # Get version info if dynamic versioning is needed
        $version = ""
        if ($DynamicVersion) {
            $chocoInfo = choco info $ID 2>&1 | Out-String
            if ($chocoInfo -match "(\d+\.\d+\.\d+)") {
                $version = $matches[1]
                Write-Host "  Detected version: $version" -ForegroundColor Cyan
            }
        }

        Write-Host ""
        choco download $ID --output-directory="$Destination" -y
        Write-Host ""

        # Find the downloaded installer
        $installers = Get-ChildItem -Path $Destination -Filter "*$ID*" -Include *.exe, *.msi, *.nupkg |
                      Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-2) } |
                      Sort-Object LastWriteTime -Descending |
                      Select-Object -First 1

        if ($installers) {
            $extension = $installers.Extension

            # Replace {version} placeholder if present
            $finalName = $Name
            if ($DynamicVersion -and $version -and $Name -match "\{version\}") {
                $finalName = $Name -replace "\{version\}", $version
            }

            $destPath = Join-Path $Destination "$finalName$extension"

            # Remove old file if exists with same name
            if (Test-Path $destPath) {
                Remove-Item $destPath -Force
            }

            # Rename if different
            if ($installers.FullName -ne $destPath) {
                Move-Item $installers.FullName $destPath -Force
            }

            # Clean up any YAML files in target directory
            Get-ChildItem -Path $Destination -Filter "*.yaml" | Remove-Item -Force -ErrorAction SilentlyContinue

            $fileSize = (Get-Item $destPath).Length / 1MB
            Write-Host "  ✓ Downloaded with Chocolatey ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Host "  ✗ Chocolatey failed: $_" -ForegroundColor Red
        return $false
    }
}

# Function to get latest version from Winget
function Get-WingetVersion {
    param($ID)
    try {
        $wingetInfo = winget show --id $ID --accept-source-agreements 2>&1 | Out-String
        if ($wingetInfo -match "Version:\s+(.+)") {
            return $matches[1].Trim()
        }
    }
    catch {
        return $null
    }
    return $null
}

# Function to get latest version from Chocolatey
function Get-ChocoVersion {
    param($ID)
    try {
        $chocoInfo = choco info $ID 2>&1 | Out-String
        if ($chocoInfo -match "(\d+\.\d+[\.\d+]*)") {
            return $matches[1]
        }
    }
    catch {
        return $null
    }
    return $null
}

# Function to check if file exists with version
function Get-InstalledVersion {
    param($Destination, $NamePattern)

    # Convert name pattern to regex (replace {version} with version capture group)
    $pattern = $NamePattern -replace "\{version\}", "(\d+\.\d+[\.\d+]*)"

    # Find files matching the pattern
    $files = Get-ChildItem -Path $Destination -File | Where-Object {
        $_.BaseName -match $pattern
    }

    if ($files) {
        # Extract version from filename
        $file = $files | Select-Object -First 1
        if ($file.BaseName -match "(\d+\.\d+[\.\d+]*)") {
            return @{
                Version = $matches[1]
                File = $file
            }
        }
    }

    return $null
}

# Download each software
foreach ($app in $Software) {
    Write-Host "Checking $($app.Name)..." -ForegroundColor Cyan

    # Get latest available version
    $latestVersion = $null
    if ($wingetInstalled) {
        $latestVersion = Get-WingetVersion -ID $app.WingetID
    }
    if (!$latestVersion -and $chocoInstalled) {
        $latestVersion = Get-ChocoVersion -ID $app.ChocoID
    }

    if ($latestVersion) {
        Write-Host "  Latest version available: $latestVersion" -ForegroundColor Cyan
    }

    # Check if we already have this version installed
    $needsUpdate = $true
    if ($app.DynamicVersion -and $latestVersion) {
        $installed = Get-InstalledVersion -Destination $TargetFolder -NamePattern $app.DesiredName

        if ($installed) {
            Write-Host "  Installed version: $($installed.Version)" -ForegroundColor Yellow

            if ($installed.Version -eq $latestVersion) {
                Write-Host "  ✓ Already up to date! Skipping download." -ForegroundColor Green
                $needsUpdate = $false
            }
            else {
                Write-Host "  ⚠ Update available! Downloading..." -ForegroundColor Yellow
                # Delete old version
                Remove-Item $installed.File.FullName -Force
            }
        }
        else {
            Write-Host "  No existing version found. Downloading..." -ForegroundColor Yellow
        }
    }
    else {
        # For non-dynamic versions, check if any file exists with the desired name
        $existingFiles = Get-ChildItem -Path $TargetFolder -File | Where-Object {
            $_.BaseName -like "$($app.DesiredName)*"
        }

        if ($existingFiles) {
            Write-Host "  File exists: $($existingFiles[0].Name)" -ForegroundColor Yellow
            Write-Host "  Re-downloading to ensure latest version..." -ForegroundColor Yellow
            Remove-Item $existingFiles[0].FullName -Force
        }
        else {
            Write-Host "  No existing file found. Downloading..." -ForegroundColor Yellow
        }
    }

    # Download if needed
    if ($needsUpdate) {
        $success = $false

        # Try Winget first
        if ($wingetInstalled) {
            $success = Download-WithWinget -ID $app.WingetID -Name $app.DesiredName -Destination $TargetFolder -DynamicVersion $app.DynamicVersion
        }

        # Fall back to Chocolatey if Winget failed
        if (!$success -and $chocoInstalled) {
            $success = Download-WithChocolatey -ID $app.ChocoID -Name $app.DesiredName -Destination $TargetFolder -DynamicVersion $app.DynamicVersion
        }

        if (!$success) {
            Write-Host "  ✗ Failed to download $($app.Name)" -ForegroundColor Red
        }
    }

    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Download process completed!" -ForegroundColor Green
Write-Host "Files saved to: $TargetFolder" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
