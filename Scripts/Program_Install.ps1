#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Automated Program Installation Suite
.DESCRIPTION
    Installs multiple programs silently with progress tracking
.PARAMETER InstallersPath
    Path to the directory containing installation files. Defaults to "$env:WINDIR\Setup\Files"
.EXAMPLE
    .\Program_Install.ps1
    Uses default path: <ScriptDirectory>\Softwares
.EXAMPLE
    .\Program_Install.ps1 -InstallersPath "C:\MyInstallers"
    Uses custom path: C:\MyInstallers
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$InstallersPath = "$env:WINDIR\Setup\Files"
)

# Set location to script directory
Set-Location -Path $PSScriptRoot

# Validate installers path
if (-not (Test-Path -Path $InstallersPath -PathType Container)) {
    Write-Host "ERROR: Installers path not found: $InstallersPath" -ForegroundColor Red
    Write-Host "Please ensure the directory exists and contains the installation files." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Configure console for better output
$Host.UI.RawUI.WindowTitle = "Program Installation Suite"

# Color scheme
$colors = @{
    Header    = 'Cyan'
    Success   = 'Green'
    Info      = 'Yellow'
    Progress  = 'Magenta'
    Separator = 'DarkGray'
    White     = 'White'
    Warning   = 'Yellow'
}

# ASCII-safe icons
$icons = @{
    Arrow   = ">"
    Success = "[OK]"
    Failure = "[FAIL]"
    Block   = "#"
    Empty   = "-"
    Check   = "[+]"
    Cross   = "[X]"
}

# Function to resolve installer path with wildcard support
function Resolve-InstallerPath {
    param(
        [string]$Pattern,
        [string]$BasePath
    )

    # If pattern contains wildcard, search for matching file
    if ($Pattern -like "*`**") {
        $searchPattern = Split-Path -Leaf $Pattern
        $matchingFiles = Get-ChildItem -Path $BasePath -Filter $searchPattern -File -ErrorAction SilentlyContinue

        if ($matchingFiles) {
            # Return the first match (or most recent if multiple)
            $resolved = ($matchingFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
            return $resolved
        }
        else {
            return $Pattern  # Return original pattern if no match found
        }
    }
    else {
        return $Pattern  # Return as-is if no wildcard
    }
}

# Function to display a fancy header
function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor $colors.Header
    Write-Host "          AUTOMATED PROGRAM INSTALLATION SUITE" -ForegroundColor $colors.White
    Write-Host "=================================================================" -ForegroundColor $colors.Header
    Write-Host ""
    Write-Host "  Installers Path: " -NoNewline -ForegroundColor $colors.Info
    Write-Host $InstallersPath -ForegroundColor $colors.White
    Write-Host ""
}

# Function to verify installer paths
function Show-InstallerPaths {
    param(
        [array]$Programs
    )

    Write-Host "=================================================================" -ForegroundColor $colors.Header
    Write-Host "              VERIFYING INSTALLER FILES" -ForegroundColor $colors.White
    Write-Host "=================================================================" -ForegroundColor $colors.Header
    Write-Host ""

    $missingFiles = @()

    foreach ($program in $Programs) {
        # Skip msiexec.exe and other system executables
        if ($program.FilePath -eq "msiexec.exe") {
            # Extract the actual installer path from arguments
            $installerArg = $program.Arguments | Where-Object { $_ -like "*$InstallersPath*" -or $_ -like "*.msi*" }
            if ($installerArg) {
                # Remove quotes and /i flag
                $installerPath = $installerArg -replace '^/i,?\s*', '' -replace '"', '' -replace '`"', ''

                # Resolve wildcard if present
                $resolvedPath = Resolve-InstallerPath -Pattern $installerPath -BasePath $InstallersPath

                Write-Host "  $($program.Name)" -ForegroundColor $colors.White
                if ($installerPath -ne $resolvedPath) {
                    Write-Host "    Pattern: " -NoNewline -ForegroundColor $colors.Separator
                    Write-Host $installerPath -ForegroundColor DarkGray
                    Write-Host "    Resolved: " -NoNewline -ForegroundColor $colors.Separator
                    Write-Host $resolvedPath -ForegroundColor Gray
                }
                else {
                    Write-Host "    Path: " -NoNewline -ForegroundColor $colors.Separator
                    Write-Host $resolvedPath -ForegroundColor Gray
                }

                if (Test-Path -Path $resolvedPath -PathType Leaf) {
                    Write-Host "    Status: " -NoNewline -ForegroundColor $colors.Separator
                    Write-Host "$($icons.Check) Found" -ForegroundColor $colors.Success
                    # Update the program's arguments with resolved path
                    $program.ResolvedPath = $resolvedPath
                }
                else {
                    Write-Host "    Status: " -NoNewline -ForegroundColor $colors.Separator
                    Write-Host "$($icons.Cross) NOT FOUND" -ForegroundColor Red
                    $missingFiles += $installerPath
                }
                Write-Host ""
            }
        }
        else {
            # Direct executable path - resolve wildcards
            $resolvedPath = Resolve-InstallerPath -Pattern $program.FilePath -BasePath $InstallersPath

            Write-Host "  $($program.Name)" -ForegroundColor $colors.White
            if ($program.FilePath -ne $resolvedPath) {
                Write-Host "    Pattern: " -NoNewline -ForegroundColor $colors.Separator
                Write-Host $program.FilePath -ForegroundColor DarkGray
                Write-Host "    Resolved: " -NoNewline -ForegroundColor $colors.Separator
                Write-Host $resolvedPath -ForegroundColor Gray
            }
            else {
                Write-Host "    Path: " -NoNewline -ForegroundColor $colors.Separator
                Write-Host $resolvedPath -ForegroundColor Gray
            }

            if (Test-Path -Path $resolvedPath -PathType Leaf) {
                Write-Host "    Status: " -NoNewline -ForegroundColor $colors.Separator
                Write-Host "$($icons.Check) Found" -ForegroundColor $colors.Success
                # Update the program's filepath with resolved path
                $program.ResolvedPath = $resolvedPath
            }
            else {
                Write-Host "    Status: " -NoNewline -ForegroundColor $colors.Separator
                Write-Host "$($icons.Cross) NOT FOUND" -ForegroundColor Red
                $missingFiles += $program.FilePath
            }
            Write-Host ""
        }
    }

    Write-Host "=================================================================" -ForegroundColor $colors.Separator
    Write-Host ""

    if ($missingFiles.Count -gt 0) {
        Write-Host "WARNING: $($missingFiles.Count) installer file(s) not found!" -ForegroundColor Red
        Write-Host "Missing files:" -ForegroundColor $colors.Warning
        foreach ($file in $missingFiles) {
            Write-Host "  - $file" -ForegroundColor Red
        }
        Write-Host ""
        Write-Host "Do you want to continue anyway? (Y/N): " -NoNewline -ForegroundColor $colors.Warning
        $response = Read-Host
        if ($response -ne 'Y' -and $response -ne 'y') {
            Write-Host ""
            Write-Host "Installation cancelled." -ForegroundColor $colors.Info
            Write-Host "Press any key to exit..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 0
        }
    }
    else {
        Write-Host ""
        for ($i = 10; $i -gt 0; $i--) {
            Write-Host "`rStarting installation in $i seconds... " -NoNewline -ForegroundColor $colors.Info
            Start-Sleep -Seconds 1
        }
        Write-Host ""
    }

    Clear-Host
}

# Function to display installation status
function Install-Program {
    param(
        [string]$Name,
        [string]$FilePath,
        [string]$ResolvedPath,
        [string[]]$Arguments,
        [int]$Current,
        [int]$Total
    )

    Write-Host ""
    Write-Host "[$Current/$Total] " -NoNewline -ForegroundColor $colors.Progress
    Write-Host $Name -ForegroundColor $colors.White
    Write-Host "  ---------------------------------------------------------------" -ForegroundColor $colors.Separator

    # Show 0% progress bar at start
    $progressBarEmpty = ($icons.Empty * 20)
    Write-Host "  $($icons.Arrow) Installing... " -NoNewline -ForegroundColor $colors.Info
    Write-Host "[$progressBarEmpty] 0%" -ForegroundColor $colors.Info

    # Use resolved path if available, otherwise use original
    $executablePath = if ($ResolvedPath) { $ResolvedPath } else { $FilePath }

    # If this is an msiexec call, update the arguments with resolved path
    if ($FilePath -eq "msiexec.exe" -and $ResolvedPath) {
        # Find and replace the installer path in arguments
        $updatedArgs = @()
        foreach ($arg in $Arguments) {
            if ($arg -like "*$InstallersPath*" -or $arg -like "*.msi*") {
                $updatedArgs += "`"$ResolvedPath`""
            }
            else {
                $updatedArgs += $arg
            }
        }
        $Arguments = $updatedArgs
        $executablePath = $FilePath  # Keep msiexec.exe as the executable
    }

    try {
        $process = Start-Process -FilePath $executablePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop

        # Move cursor up to overwrite the progress line
        $cursorPos = $Host.UI.RawUI.CursorPosition
        $cursorPos.Y = $cursorPos.Y - 1
        $Host.UI.RawUI.CursorPosition = $cursorPos

        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            # Show 100% progress bar on success
            $progressBarFull = ($icons.Block * 20)
            Write-Host "  $($icons.Success) Complete!    " -NoNewline -ForegroundColor $colors.Success
            Write-Host "[$progressBarFull] 100%" -ForegroundColor $colors.Success
        }
        else {
            # Show failure with incomplete progress
            $progressBarPartial = ($icons.Block * 10) + ($icons.Empty * 10)
            Write-Host "  $($icons.Failure) Failed!      " -NoNewline -ForegroundColor Red
            Write-Host "[$progressBarPartial] 50%  " -ForegroundColor Red
            Write-Host "  Exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    }
    catch {
        # Move cursor up to overwrite the progress line
        $cursorPos = $Host.UI.RawUI.CursorPosition
        $cursorPos.Y = $cursorPos.Y - 1
        $Host.UI.RawUI.CursorPosition = $cursorPos

        # Show failure
        Write-Host "  $($icons.Failure) Failed!      " -NoNewline -ForegroundColor Red
        Write-Host "[$progressBarEmpty] 0%  " -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }

    Start-Sleep -Seconds 1
}

# Display header
Show-Header

# Define installation queue with wildcard support
$programs = @(
    @{
        Name      = "K-Lite Codec Pack Mega"
        FilePath  = "$InstallersPath\K-Lite Codec Pack Mega*.exe"
        Arguments = @(
            "/VERYSILENT",
            "/NORESTART",
            "/SUPPRESSMSGBOXES",
            "/LOADINF=`"$InstallersPath\klcp_mega_unattended.ini`""
        )
    },
    @{
        Name      = "Google Chrome"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$InstallersPath\Google Chrome*.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "QuickLook"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$InstallersPath\QuickLook*.msi`"",
            "INSTALLFOLDER=`"C:\Program Files (x86)\QuickLook`"",
            "ALLUSERS=1",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "File Converter"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$InstallersPath\File Converter*.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "7-Zip"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$InstallersPath\7-Zip*.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "AnyDesk"
        FilePath  = "$InstallersPath\AnyDesk*.exe"
        Arguments = @(
            "--install",
            "`"C:\Program Files (x86)\AnyDesk`"",
            "--silent",
            "--create-shortcuts",
            "--create-desktop-icon"
        )
    },
    @{
        Name      = "Everything"
        FilePath  = "$InstallersPath\Everything*.exe"
        Arguments = @(
            "/S",
            "-install-options",
            "`"-app-data -disable-run-as-admin -install-all-users-desktop-shortcut -install-efu-association install-quick-launch-shortcut -install-all-users-start-menu-shortcuts -install-folder-context-menu -install-run-on-system-startup`"",
            "/D=`"C:\Program Files\Everything`""
        )
    },
    @{
        Name      = "VLC Media Player"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$InstallersPath\VLC Media Player*.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "WinRAR"
        FilePath  = "$InstallersPath\WinRAR*.exe"
        Arguments = @("/S")
    }
)

# Show all installer paths and verify they exist
Show-InstallerPaths -Programs $programs

# Display installation header
Show-Header

$totalPrograms = $programs.Count
$currentProgram = 0

# Install each program
foreach ($program in $programs) {
    $currentProgram++
    Install-Program -Name $program.Name -FilePath $program.FilePath -ResolvedPath $program.ResolvedPath -Arguments $program.Arguments -Current $currentProgram -Total $totalPrograms
}

# Completion message
Write-Host ""
Write-Host "=================================================================" -ForegroundColor $colors.Success
Write-Host "                    INSTALLATION COMPLETE!" -ForegroundColor $colors.White
Write-Host "=================================================================" -ForegroundColor $colors.Success
Write-Host ""
Write-Host "  All programs have been processed." -ForegroundColor $colors.Success
Write-Host ""
for ($i = 10; $i -gt 0; $i--) {
    Write-Host "`rSelf-destruct in $i seconds... " -NoNewline -ForegroundColor $colors.Info
    Start-Sleep -Seconds 1
}
Write-Host ""
