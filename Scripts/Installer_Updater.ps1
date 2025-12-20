#Requires -Version 5.1
<#
.SYNOPSIS
    Advanced Winget Installer Manager with intelligent version tracking
.DESCRIPTION
    Downloads and updates software installers with version comparison and caching
.PARAMETER TargetFolder
    Destination folder for installers
.PARAMETER Force
    Force download regardless of current version
.PARAMETER SkipVersionCheck
    Skip online version verification
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetFolder = "D:\ISO FIles\Custom ISO Files\SetupFiles\Software",

    [switch]$Force,
    [switch]$SkipVersionCheck
)

#region Configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$Config = @{
    DownloadTimeout  = 300
    RetryAttempts    = 2
    FileAgeThreshold = 7
}

$Apps = @(
    @{
        Name    = "WinRAR"
        ID      = "RARLab.WinRAR"
        WantMSI = $false
    }
    @{
        Name      = "VLC Media Player"
        ID        = "VideoLAN.VLC"
        DirectUrl = "https://download.videolan.org/pub/videolan/vlc/{version}/win64/vlc-{version}-win64.msi"
        WantMSI   = $true
    }
    @{
        Name      = "QuickLook"
        ID        = "QL-Win.QuickLook"
        DirectUrl = "https://github.com/QL-Win/QuickLook/releases/download/{version}/QuickLook-{version}.msi"
        WantMSI   = $true
    }
    @{
        Name    = "K-Lite Codec Pack Mega"
        ID      = "CodecGuide.K-LiteCodecPack.Mega"
        WantMSI = $false
    }
    @{
        Name      = "Google Chrome"
        ID        = "Google.Chrome"
        DirectUrl = "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
        WantMSI   = $true
    }
    @{
        Name      = "File Converter"
        ID        = "AdrienAllard.FileConverter"
        DirectUrl = "https://github.com/Tichau/FileConverter/releases/download/v{version}/FileConverter-{version}-x64-setup.msi"
        WantMSI   = $true
    }
    @{
        Name      = "Everything"
        ID        = "voidtools.Everything"
        DirectUrl = "https://www.voidtools.com/Everything-{version}.x64-Setup.exe"
        WantMSI   = $false
    }
    @{
        Name    = "AnyDesk"
        ID      = "AnyDesk.AnyDesk"
        WantMSI = $false
    }
    @{
        Name      = "7-Zip"
        ID        = "7zip.7zip"
        DirectUrl = "https://github.com/ip7z/7zip/releases/download/{version}/7z{versionshort}-x64.msi"
        WantMSI   = $true
    }
)
#endregion

#region UI Helpers
function Write-Banner {
    param([string]$Text, [ConsoleColor]$Color = 'Cyan')

    $width = 80
    $padding = [math]::Max(0, ($width - $Text.Length - 2) / 2)
    $line = '═' * $width

    Write-Host $line -ForegroundColor $Color
    Write-Host ("║" + (' ' * [math]::Floor($padding)) + $Text + (' ' * [math]::Ceiling($padding)) + "║") -ForegroundColor $Color
    Write-Host $line -ForegroundColor $Color
}

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Progress')]
        [string]$Type = 'Info',
        [int]$Indent = 2
    )

    $prefix = ' ' * $Indent
    $icons = @{
        Info     = @{ Symbol = '●'; Color = 'Cyan' }
        Success  = @{ Symbol = '✓'; Color = 'Green' }
        Warning  = @{ Symbol = '⚠'; Color = 'Yellow' }
        Error    = @{ Symbol = '✗'; Color = 'Red' }
        Progress = @{ Symbol = '⟳'; Color = 'Magenta' }
    }

    $icon = $icons[$Type]
    Write-Host "$prefix$($icon.Symbol) " -NoNewline -ForegroundColor $icon.Color
    Write-Host $Message -ForegroundColor $icon.Color
}

function Write-AppHeader {
    param([string]$Name, [int]$Current, [int]$Total)
    Write-Host "`n┌─ " -NoNewline -ForegroundColor DarkGray
    Write-Host "[$Current/$Total] $Name" -NoNewline -ForegroundColor White
    Write-Host " ─┐" -ForegroundColor DarkGray
}
#endregion

#region Core Functions
function Test-WingetAvailable {
    try {
        $null = winget --version
        return $true
    }
    catch {
        return $false
    }
}

function Get-LatestVersion {
    param([string]$PackageID)

    try {
        $output = winget show --id $PackageID --accept-source-agreements 2>&1 | Out-String
        if ($output -match "Version:\s+(\d+\.[\d\.]+)") {
            return $matches[1].Trim()
        }
    }
    catch {}
    return $null
}

function Get-InstallerVersion {
    param([string]$FilePath)

    try {
        $fileInfo = Get-Item -LiteralPath $FilePath -ErrorAction Stop

        # First try to extract from filename (most reliable for renamed files)
        if ($fileInfo.Name -match "[\s_\-v](\d+(?:\.\d+)+)") {
            $version = $matches[1].TrimEnd('.')
            return $version
        }

        # Then try file metadata
        $versionInfo = $fileInfo.VersionInfo
        $version = $versionInfo.ProductVersion ?? $versionInfo.FileVersion
        if ($version -and $version -match "(\d+(?:\.\d+)+)") {
            return $matches[1].TrimEnd('.')
        }
    }
    catch {}
    return $null
}

function Compare-SoftwareVersions {
    param([string]$Version1, [string]$Version2)

    if ([string]::IsNullOrWhiteSpace($Version1) -or [string]::IsNullOrWhiteSpace($Version2)) {
        return $null
    }

    try {
        $v1 = [version]$Version1.Trim()
        $v2 = [version]$Version2.Trim()
        return [Math]::Sign($v1.CompareTo($v2))
    }
    catch {
        return $null
    }
}

function Find-ExistingInstaller {
    param([hashtable]$App, [string]$Folder)

    $searchTerms = @($App.Name) + ($App.ID -split '\.' | Select-Object -Last 1)
    $files = Get-ChildItem -Path $Folder -File -ErrorAction SilentlyContinue

    foreach ($term in $searchTerms) {
        $foundFiles = $files | Where-Object { $_.Name -like "*$term*" }
        if ($foundFiles) {
            # Return the most recently modified file
            return $foundFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        }
    }
    return $null
}

function Invoke-DirectDownload {
    param([string]$Url, [string]$Destination)

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $Destination)
        $webClient.Dispose()
        return $true
    }
    catch {
        return $false
    }
}

function Get-WingetInstaller {
    param([hashtable]$App, [string]$TempDir, [string]$LatestVersion)

    # Try direct download first (supports both MSI and EXE)
    if ($App.DirectUrl -and $LatestVersion) {
        $downloadUrl = $App.DirectUrl -replace '\{version\}', $LatestVersion
        if ($downloadUrl -match '\{versionshort\}') {
            $downloadUrl = $downloadUrl -replace '\{versionshort\}', ($LatestVersion -replace '\.', '')
        }

        $fileName = [System.IO.Path]::GetFileName($downloadUrl)
        $destination = Join-Path $TempDir $fileName

        Write-Status "Attempting direct download: $fileName" -Type Progress
        if (Invoke-DirectDownload -Url $downloadUrl -Destination $destination) {
            return Get-Item -LiteralPath $destination
        }
        Write-Status "Direct download failed, using Winget..." -Type Warning
    }

    # Fallback to Winget
    Write-Status "Downloading via Winget..." -Type Progress
    $downloadOutput = winget download --id $App.ID --download-directory $TempDir `
        --accept-source-agreements --accept-package-agreements 2>&1 | Out-String

    # Check for download errors
    if ($LASTEXITCODE -ne 0 -and $downloadOutput -match 'error|failed') {
        Write-Status "Winget download encountered issues" -Type Warning
    }

    # Find installer based on preference
    $patterns = if ($App.WantMSI) { @('*.msi', '*.exe') } else { @('*.exe', '*.msi') }
    foreach ($pattern in $patterns) {
        $installer = Get-ChildItem -Path $TempDir -Recurse -Filter $pattern -ErrorAction SilentlyContinue |
        Select-Object -First 1
        if ($installer) { return $installer }
    }

    return $null
}

function Update-Installer {
    param([hashtable]$App, [string]$TargetFolder, [hashtable]$Options)

    $result = @{ Status = 'Unknown'; Message = ''; SizeMB = 0 }

    try {
        # Check for latest version (use user-provided version for Chrome if available)
        $latestVersion = if ($App.UserProvidedVersion) {
            $App.UserProvidedVersion
        }
        elseif (-not $Options.SkipVersionCheck) {
            Get-LatestVersion -PackageID $App.ID
        }
        else { $null }

        if ($latestVersion) {
            Write-Status "Latest: v$latestVersion" -Type Info
        }

        # Check existing file
        $existingFile = Find-ExistingInstaller -App $App -Folder $TargetFolder

        if ($existingFile -and -not $Options.Force) {
            $existingVersion = Get-InstallerVersion -FilePath $existingFile.FullName

            if ($existingVersion -and $latestVersion) {
                Write-Status "Downloaded: v$existingVersion ($($existingFile.Name))" -Type Info

                $comparison = Compare-SoftwareVersions -Version1 $existingVersion -Version2 $latestVersion
                if ($comparison -eq 0) {
                    $result.Status = 'UpToDate'
                    $result.Message = "Already current (v$existingVersion)"
                    return $result
                }
                elseif ($comparison -gt 0) {
                    $result.Status = 'Newer'
                    $result.Message = "Installed version is newer"
                    return $result
                }
            }
            elseif (-not $latestVersion) {
                $fileAge = (Get-Date) - $existingFile.LastWriteTime
                if ($fileAge.TotalDays -lt $Config.FileAgeThreshold) {
                    $result.Status = 'Recent'
                    $result.Message = "File is recent ($([math]::Round($fileAge.TotalDays, 1))d old)"
                    return $result
                }
            }
        }

        # Download new installer
        $tempDir = Join-Path $env:TEMP "winget_$([guid]::NewGuid().ToString('N'))"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        try {
            Write-Status "Downloading..." -Type Progress

            $installer = Get-WingetInstaller -App $App -TempDir $tempDir -LatestVersion $latestVersion

            if (-not $installer) {
                throw "No installer found after download"
            }

            # Create filename with version for better tracking
            $extension = $installer.Extension
            if ($latestVersion) {
                $finalName = "$($App.Name) v$latestVersion$extension"
            }
            else {
                $finalName = "$($App.Name)$extension"
            }

            $destination = Join-Path $TargetFolder $finalName

            # Remove old version if exists
            if ($existingFile -and (Test-Path $existingFile.FullName)) {
                Remove-Item $existingFile.FullName -Force -ErrorAction SilentlyContinue
            }

            Copy-Item -LiteralPath $installer.FullName -Destination $destination -Force

            $result.Status = 'Updated'
            $result.SizeMB = [math]::Round((Get-Item $destination).Length / 1MB, 2)
            $result.Message = "Downloaded: $finalName ($($result.SizeMB) MB)"

        }
        finally {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        $result.Status = 'Failed'
        $result.Message = $_.Exception.Message
    }

    return $result
}
#endregion

#region Main Execution
Clear-Host

# Banner
Write-Banner "WINGET INSTALLER MANAGER" -Color Cyan
Write-Host "`n  📁 Target: " -NoNewline -ForegroundColor DarkGray
Write-Host $TargetFolder -ForegroundColor Yellow

if ($Force) {
    Write-Host "  ⚡ Mode: " -NoNewline -ForegroundColor DarkGray
    Write-Host "FORCE UPDATE" -ForegroundColor Magenta
}

# Prompt for Chrome version
Write-Host ""
Write-Host "  🌐 Google Chrome Version Input" -ForegroundColor Cyan
Write-Host "     (Check latest at: https://chromereleases.googleblog.com/)" -ForegroundColor DarkGray
$chromeVersion = Read-Host "     Enter Chrome version (e.g., 131.0.6778.86) or press Enter to skip"

if (-not [string]::IsNullOrWhiteSpace($chromeVersion)) {
    # Update Chrome app configuration with user-provided version
    $chromeApp = $Apps | Where-Object { $_.Name -eq "Google Chrome" }
    if ($chromeApp) {
        $chromeApp.UserProvidedVersion = $chromeVersion.Trim()
        Write-Status "Chrome version set to: v$($chromeApp.UserProvidedVersion)" -Type Success
    }
}

# Verify prerequisites
Write-Host ""
if (-not (Test-WingetAvailable)) {
    Write-Status "Winget is not installed!" -Type Error
    Write-Host "  Install from: https://aka.ms/getwinget" -ForegroundColor Yellow
    exit 1
}
Write-Status "Winget verified" -Type Success

# Create target folder
if (-not (Test-Path $TargetFolder)) {
    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
    Write-Status "Created target folder" -Type Success
}

# Process applications
$stats = @{ Updated = 0; UpToDate = 0; Failed = 0 }
$total = $Apps.Count
$current = 0

foreach ($app in $Apps) {
    $current++
    Write-AppHeader -Name $app.Name -Current $current -Total $total

    $options = @{
        Force            = $Force.IsPresent
        SkipVersionCheck = $SkipVersionCheck.IsPresent
    }

    $result = Update-Installer -App $app -TargetFolder $TargetFolder -Options $options

    switch ($result.Status) {
        'Updated' {
            Write-Status $result.Message -Type Success
            $stats.Updated++
        }
        { $_ -in 'UpToDate', 'Recent', 'Newer' } {
            Write-Status $result.Message -Type Info
            $stats.UpToDate++
        }
        'Failed' {
            Write-Status "Error: $($result.Message)" -Type Error
            $stats.Failed++
        }
    }
}

# Summary
Write-Host "`n"
Write-Banner "OPERATION COMPLETE" -Color Green
Write-Host "`n  📊 Results:" -ForegroundColor Cyan
Write-Host "     Updated:           " -NoNewline -ForegroundColor DarkGray
Write-Host $stats.Updated -ForegroundColor Green
Write-Host "     Already current:   " -NoNewline -ForegroundColor DarkGray
Write-Host $stats.UpToDate -ForegroundColor Yellow
Write-Host "     Failed:            " -NoNewline -ForegroundColor DarkGray
Write-Host $stats.Failed -ForegroundColor Red

Write-Host "`n  📂 Location: " -NoNewline -ForegroundColor DarkGray
Write-Host $TargetFolder -ForegroundColor Yellow
Write-Host ""

$ProgressPreference = 'Continue'
#endregion
