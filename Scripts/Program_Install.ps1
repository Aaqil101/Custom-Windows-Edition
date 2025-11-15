#Requires -Version 5.0
#Requires -RunAsAdministrator

# Set location to script directory
Set-Location -Path $PSScriptRoot

# Configure console for better output
$Host.UI.RawUI.WindowTitle = "Program Installation Suite"
$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"

# Color scheme
$colors = @{
    Header    = 'Cyan'
    Success   = 'Green'
    Info      = 'Yellow'
    Progress  = 'Magenta'
    Separator = 'DarkGray'
    White     = 'White'
}

# ASCII-safe icons
$icons = @{
    Arrow   = ">"
    Success = "[OK]"
    Failure = "[FAIL]"
    Block   = "#"
    Empty   = "-"
}

# Function to display a fancy header
function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor $colors.Header
    Write-Host "          AUTOMATED PROGRAM INSTALLATION SUITE" -ForegroundColor $colors.White
    Write-Host "=================================================================" -ForegroundColor $colors.Header
    Write-Host ""
}

# Function to display installation status
function Install-Program {
    param(
        [string]$Name,
        [string]$FilePath,
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

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop

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

# Define installation queue
$programs = @(
    @{
        Name      = "K-Lite Mega Codec Pack"
        FilePath  = "$env:WINDIR\Setup\Files\K-Lite_Codec_Pack_Mega.exe"
        Arguments = @(
            "/VERYSILENT",
            "/NORESTART",
            "/SUPPRESSMSGBOXES",
            "/LOADINF=`"$env:WINDIR\Setup\Files\klcp_mega_unattended.ini`""
        )
    },
    @{
        Name      = "Google Chrome"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$env:WINDIR\Setup\Files\googlechromestandaloneenterprise64.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "QuickLook"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$env:WINDIR\Setup\Files\QuickLook.msi`"",
            "INSTALLFOLDER=`"C:\Program Files (x86)\QuickLook`"",
            "ALLUSERS=1",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "FileConverter"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$env:WINDIR\Setup\Files\FileConverter.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "7-Zip"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$env:WINDIR\Setup\Files\7-zip.msi`"",
            "/qb",
            "/norestart"
        )
    },
    @{
        Name      = "AnyDesk"
        FilePath  = "$env:WINDIR\Setup\Files\AnyDesk.exe"
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
        FilePath  = "$env:WINDIR\Setup\Files\Everything.exe"
        Arguments = @(
            "/S",
            "-install-options",
            "`"-app-data -disable-run-as-admin -install-all-users-desktop-shortcut -install-efu-association install-quick-launch-shortcut -install-all-users-start-menu-shortcuts -install-folder-context-menu -install-run-on-system-startup`"",
            "/D=`"C:\Program Files\Everything`""
        )
    },
    @{
        Name      = "VLC Media Player"
        FilePath  = "$env:WINDIR\Setup\Files\vlc-3.0.21-win64.exe"
        Arguments = @(
            "/L=1033",
            "/S"
        )
    },
    @{
        Name      = "WinRAR"
        FilePath  = "$env:WINDIR\Setup\Files\winrar.exe"
        Arguments = @("/S")
    }
)

$totalPrograms = $programs.Count
$currentProgram = 0

# Install each program
foreach ($program in $programs) {
    $currentProgram++
    Install-Program -Name $program.Name -FilePath $program.FilePath -Arguments $program.Arguments -Current $currentProgram -Total $totalPrograms
}

# Completion message
Write-Host ""
Write-Host "=================================================================" -ForegroundColor $colors.Success
Write-Host "                    INSTALLATION COMPLETE!" -ForegroundColor $colors.White
Write-Host "=================================================================" -ForegroundColor $colors.Success
Write-Host ""
Write-Host "  All programs have been processed." -ForegroundColor $colors.Success
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor $colors.Info
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
