#Requires -Version 5.0

# ============================================================================
#  QuickLook Plugin Installer
#  Extracts plugin archives to QuickLook installation directory
# ============================================================================

# Configuration
$zipFolder = "$env:WINDIR\Setup\Files\QuickLook.Plugin"
$destinationRoot = "C:\Program Files (x86)\QuickLook\QuickLook.Plugin"
$7zipPath = "C:\Program Files\7-Zip\7z.exe"

# Spinner animation frames
$spinnerFrames = @('|', '/', '-', '\')

# Function to get color based on percentage
function Get-ProgressColor {
    param([int]$Percent)

    if ($Percent -lt 33) { return 'Red' }
    elseif ($Percent -lt 66) { return 'Yellow' }
    else { return 'Green' }
}

# Function to format file size
function Format-FileSize {
    param([long]$Size)

    if ($Size -gt 1MB) {
        return "{0:N2} MB" -f ($Size / 1MB)
    }
    elseif ($Size -gt 1KB) {
        return "{0:N2} KB" -f ($Size / 1KB)
    }
    else {
        return "$Size bytes"
    }
}

# Clear screen for clean output
Clear-Host

# Start timing
$startTime = Get-Date

# Display banner
Write-Host ""
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "                                                                 " -ForegroundColor Cyan
Write-Host "           QuickLook Plugin Installation Script                 " -ForegroundColor Cyan
Write-Host "                                                                 " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host ""

# Check if 7-Zip exists
Write-Host "[*] Checking prerequisites..." -ForegroundColor Yellow
Start-Sleep -Milliseconds 300

if (!(Test-Path $7zipPath)) {
    Write-Host ""
    Write-Host "[X] ERROR: 7-Zip not found at $7zipPath" -ForegroundColor Red
    Write-Host ""
    exit 1
}
Write-Host "  [+] 7-Zip found" -ForegroundColor Green
Start-Sleep -Milliseconds 200

# Check if source folder exists
if (!(Test-Path $zipFolder)) {
    Write-Host "  [X] Source folder not found: $zipFolder" -ForegroundColor Red
    Write-Host ""
    exit 1
}
Write-Host "  [+] Source folder found" -ForegroundColor Green
Start-Sleep -Milliseconds 200

# Create destination root if it doesn't exist
if (!(Test-Path $destinationRoot)) {
    Write-Host "  [~] Creating destination directory..." -ForegroundColor Gray
    New-Item -ItemType Directory -Path $destinationRoot | Out-Null
    Write-Host "  [+] Destination directory created" -ForegroundColor Green
}
else {
    Write-Host "  [+] Destination directory exists" -ForegroundColor Green
}

Start-Sleep -Milliseconds 200
Write-Host ""

# Get all zip files
$zipFiles = Get-ChildItem -Path $zipFolder -Filter "*.zip"
$totalFiles = $zipFiles.Count
$totalSize = ($zipFiles | Measure-Object -Property Length -Sum).Sum

if ($totalFiles -eq 0) {
    Write-Host "[!] No zip files found in $zipFolder" -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

Write-Host "[*] Found $totalFiles plugin(s) to install" -ForegroundColor Cyan
Write-Host "    Total size: $(Format-FileSize $totalSize)" -ForegroundColor Gray
Write-Host ""
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

# Extract each zip file
$currentFile = 0
$processedSize = 0
$successCount = 0
$failCount = 0

foreach ($zip in $zipFiles) {
    $currentFile++
    $fileSize = $zip.Length

    # Get the zip file name without extension
    $folderName = [System.IO.Path]::GetFileNameWithoutExtension($zip.Name)

    # Create destination path
    $extractPath = Join-Path -Path $destinationRoot -ChildPath $folderName

    # Display file header
    Write-Host "[$currentFile/$totalFiles] " -NoNewline -ForegroundColor Cyan
    Write-Host "$folderName " -NoNewline -ForegroundColor Yellow
    Write-Host "($(Format-FileSize $fileSize))" -ForegroundColor Gray

    # Create the folder if it doesn't exist
    if (!(Test-Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }

    # Show extraction with spinner
    $extractJob = {
        param($zipPath, $7zip, $dest)
        & $7zip x $zipPath -o"$dest" -y 2>&1 | Out-Null
        return $LASTEXITCODE
    }

    $frameIndex = 0
    $job = Start-Job -ScriptBlock $extractJob -ArgumentList $zip.FullName, $7zipPath, $extractPath

    while ($job.State -eq 'Running') {
        $frame = $spinnerFrames[$frameIndex % $spinnerFrames.Length]
        Write-Host "`r  $frame Extracting..." -NoNewline -ForegroundColor Cyan
        Start-Sleep -Milliseconds 100
        $frameIndex++
    }

    $exitCode = Receive-Job -Job $job -Wait
    Remove-Job -Job $job

    # Clear spinner line
    Write-Host "`r                          `r" -NoNewline

    # Update processed size
    if ($exitCode -eq 0) {
        $processedSize += $fileSize
        $successCount++
        Write-Host "  [+] Successfully installed" -ForegroundColor Green

        # Show individual file progress (100% when complete)
        $fileBar = "#" * 40
        Write-Host "  Progress: [" -NoNewline -ForegroundColor Gray
        Write-Host $fileBar -NoNewline -ForegroundColor Green
        Write-Host "] 100%" -ForegroundColor Gray
    }
    else {
        $failCount++
        Write-Host "  [X] Installation failed" -ForegroundColor Red

        # Show failed file progress
        $fileBar = "-" * 40
        Write-Host "  Progress: [" -NoNewline -ForegroundColor Gray
        Write-Host $fileBar -NoNewline -ForegroundColor Red
        Write-Host "] 0%" -ForegroundColor Gray
    }

    Write-Host ""

    # Small pause between files
    if ($currentFile -lt $totalFiles) {
        Start-Sleep -Milliseconds 300
    }
}

# Calculate total time
$totalTime = (Get-Date) - $startTime
$timeStr = if ($totalTime.TotalMinutes -ge 1) {
    "{0:0}m {1:0}s" -f [math]::Floor($totalTime.TotalMinutes), $totalTime.Seconds
}
else {
    "{0:0}s" -f $totalTime.TotalSeconds
}

# Final summary
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "[+] Installation Complete!" -ForegroundColor Green
Write-Host ""
Write-Host "  Installed:    " -NoNewline -ForegroundColor Gray
Write-Host "$successCount plugin(s)" -ForegroundColor Cyan
if ($failCount -gt 0) {
    Write-Host "  Failed:       " -NoNewline -ForegroundColor Gray
    Write-Host "$failCount plugin(s)" -ForegroundColor Red
}
Write-Host "  Total Size:   " -NoNewline -ForegroundColor Gray
Write-Host "$(Format-FileSize $processedSize)" -ForegroundColor Cyan
Write-Host "  Time Taken:   " -NoNewline -ForegroundColor Gray
Write-Host "$timeStr" -ForegroundColor Cyan
Write-Host "  Location:     " -NoNewline -ForegroundColor Gray
Write-Host "$destinationRoot" -ForegroundColor Cyan
Write-Host ""

# Show completion animation
$celebrationFrames = @('[*]', '[+]', '[*]', '[+]')
for ($i = 0; $i -lt 8; $i++) {
    $frame = $celebrationFrames[$i % $celebrationFrames.Length]
    $color = if ($i % 2 -eq 0) { 'Green' } else { 'Yellow' }
    Write-Host "`r  $frame All plugins ready to use!" -NoNewline -ForegroundColor $color
    Start-Sleep -Milliseconds 150
}
Write-Host "`r  [+] All plugins ready to use!" -ForegroundColor Green
Write-Host ""
