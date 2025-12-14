#Requires -Version 5.0
#Requires -RunAsAdministrator

param(
    [Parameter(Mandatory = $false)]
    [string]$SetupPath = "$env:WINDIR\Setup\Files",

    [Parameter(Mandatory = $false)]
    [switch]$SkipValidation,

    [Parameter(Mandatory = $false)]
    [string[]]$ProgramsToInstall,

    [Parameter(Mandatory = $false)]
    [switch]$Verbose
)

# Performance optimizations
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Set-Location -Path $PSScriptRoot

# Configure console
$Host.UI.RawUI.WindowTitle = "⚡ Program Installation Suite"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Enhanced color palette
$c = @{
    Cyan     = [ConsoleColor]::Cyan
    Green    = [ConsoleColor]::Green
    Yellow   = [ConsoleColor]::Yellow
    Magenta  = [ConsoleColor]::Magenta
    DarkGray = [ConsoleColor]::DarkGray
    White    = [ConsoleColor]::White
    Red      = [ConsoleColor]::Red
    Blue     = [ConsoleColor]::Blue
    DarkCyan = [ConsoleColor]::DarkCyan
}

# Unicode box drawing characters
$box = @{
    TopLeft     = '╔'
    TopRight    = '╗'
    BottomLeft  = '╚'
    BottomRight = '╝'
    Horizontal  = '═'
    Vertical    = '║'
    Block       = '█'
    Shade       = '░'
    Arrow       = '→'
    Check       = '✓'
    Cross       = '✗'
    Dot         = '●'
}

function Write-BoxedHeader {
    param([string]$Text, [int]$Width = 65)

    $padding = $Width - $Text.Length - 2
    $leftPad = [math]::Floor($padding / 2)
    $rightPad = [math]::Ceiling($padding / 2)

    Write-Host "$($box.TopLeft)$($box.Horizontal * $Width)$($box.TopRight)" -ForegroundColor $c.Cyan
    Write-Host "$($box.Vertical)$(' ' * $leftPad)$Text$(' ' * $rightPad)$($box.Vertical)" -ForegroundColor $c.Cyan
    Write-Host "$($box.BottomLeft)$($box.Horizontal * $Width)$($box.BottomRight)" -ForegroundColor $c.Cyan
}

function Show-AnimatedProgress {
    param([int]$Percent, [int]$Width = 30)

    $filled = [math]::Floor($Width * $Percent / 100)
    $empty = $Width - $filled

    $bar = ($box.Block * $filled) + ($box.Shade * $empty)

    $color = switch ($Percent) {
        { $_ -lt 50 } { $c.Yellow }
        { $_ -lt 100 } { $c.Blue }
        default { $c.Green }
    }

    Write-Host "  [$bar] " -NoNewline -ForegroundColor $color
    Write-Host "$Percent%" -ForegroundColor $color
}

function Install-Program {
    param(
        [string]$Name,
        [string]$FilePath,
        [string[]]$Arguments,
        [int]$Current,
        [int]$Total
    )

    Write-Host ""
    Write-Host " $($box.Dot) " -NoNewline -ForegroundColor $c.Magenta
    Write-Host "[$Current/$Total] " -NoNewline -ForegroundColor $c.DarkCyan
    Write-Host $Name -ForegroundColor $c.White
    Write-Host "  $($box.Horizontal * 63)" -ForegroundColor $c.DarkGray

    # Initial progress
    Write-Host "  $($box.Arrow) Installing... " -NoNewline -ForegroundColor $c.Yellow
    Show-AnimatedProgress -Percent 0

    $startTime = Get-Date

    try {
        # Validate file exists
        if (-not (Test-Path $FilePath)) {
            throw "Installer not found: $FilePath"
        }

        # Start installation with optimized settings
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $FilePath
        $processInfo.Arguments = $Arguments -join ' '
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        [void]$process.Start()

        # Simulate progress animation while waiting
        $progressSteps = @(25, 50, 75)
        $stepDelay = 500

        foreach ($step in $progressSteps) {
            if (-not $process.HasExited) {
                Start-Sleep -Milliseconds $stepDelay
                $cursorPos = $Host.UI.RawUI.CursorPosition
                $cursorPos.Y--
                $Host.UI.RawUI.CursorPosition = $cursorPos
                Write-Host "  $($box.Arrow) Installing... " -NoNewline -ForegroundColor $c.Yellow
                Show-AnimatedProgress -Percent $step
            }
        }

        $process.WaitForExit()
        $exitCode = $process.ExitCode

        # Move cursor up to final line
        $cursorPos = $Host.UI.RawUI.CursorPosition
        $cursorPos.Y--
        $Host.UI.RawUI.CursorPosition = $cursorPos

        $duration = ((Get-Date) - $startTime).TotalSeconds

        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
            Write-Host "  $($box.Check) Complete!    " -NoNewline -ForegroundColor $c.Green
            Show-AnimatedProgress -Percent 100
            Write-Host "  Time: " -NoNewline -ForegroundColor $c.DarkGray
            Write-Host "$([math]::Round($duration, 1))s" -ForegroundColor $c.White
        }
        else {
            Write-Host "  $($box.Cross) Failed!      " -NoNewline -ForegroundColor $c.Red
            Show-AnimatedProgress -Percent 50
            Write-Host "  Exit Code: $exitCode" -ForegroundColor $c.Red
        }
    }
    catch {
        $cursorPos = $Host.UI.RawUI.CursorPosition
        $cursorPos.Y--
        $Host.UI.RawUI.CursorPosition = $cursorPos

        Write-Host "  $($box.Cross) Error!       " -NoNewline -ForegroundColor $c.Red
        Show-AnimatedProgress -Percent 0
        Write-Host "  $($_.Exception.Message)" -ForegroundColor $c.Red
    }

    Start-Sleep -Milliseconds 500
}

# Display animated header
Clear-Host
Write-Host ""
Write-BoxedHeader -Text "⚡ AUTOMATED PROGRAM INSTALLATION SUITE ⚡"
Write-Host ""

# Installation queue with optimized paths
$programs = @(
    @{
        Name      = "K-Lite Codec Pack Mega"
        FilePath  = "$SetupPath\K-Lite Codec Pack Mega.exe"
        Arguments = @(
            "/VERYSILENT",
            "/NORESTART",
            "/SUPPRESSMSGBOXES",
            "/LOADINF=`"$SetupPath\klcp_mega_unattended.ini`""
        )
    }
    @{
        Name      = "Google Chrome"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$SetupPath\Google Chrome.msi`"",
            "/qb",
            "/norestart"
        )
    }
    @{
        Name      = "QuickLook"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$SetupPath\QuickLook.msi`"",
            "INSTALLFOLDER=`"C:\Program Files (x86)\QuickLook`"",
            "ALLUSERS=1",
            "/qb",
            "/norestart"
        )
    }
    @{
        Name      = "File Converter"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$SetupPath\File Converter.msi`"",
            "/qb",
            "/norestart"
        )
    }
    @{
        Name      = "7-Zip"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$SetupPath\7-Zip.msi`"",
            "/qb",
            "/norestart"
        )
    }
    @{
        Name      = "AnyDesk"
        FilePath  = "$SetupPath\AnyDesk.exe"
        Arguments = @(
            "--install",
            "`"C:\Program Files (x86)\AnyDesk`"",
            "--silent",
            "--create-shortcuts",
            "--create-desktop-icon"
        )
    }
    @{
        Name      = "Everything"
        FilePath  = "$SetupPath\Everything.exe"
        Arguments = @(
            "/S",
            "-install-options",
            "`"-app-data -disable-run-as-admin -install-all-users-desktop-shortcut -install-efu-association install-quick-launch-shortcut -install-all-users-start-menu-shortcuts -install-folder-context-menu -install-run-on-system-startup`"",
            "/D=`"C:\Program Files\Everything`""
        )
    }
    @{
        Name      = "VLC Media Player"
        FilePath  = "msiexec.exe"
        Arguments = @(
            "/i",
            "`"$SetupPath\VLC media player.msi`"",
            "/qb",
            "/norestart"
        )
    }
    @{
        Name      = "WinRAR"
        FilePath  = "$SetupPath\WinRAR.exe"
        Arguments = @("/S")
    }
)

$total = $programs.Count
$startTime = Get-Date

# Execute installations
for ($i = 0; $i -lt $total; $i++) {
    Install-Program -Name $programs[$i].Name -FilePath $programs[$i].FilePath -Arguments $programs[$i].Arguments -Current ($i + 1) -Total $total
}

$totalDuration = ((Get-Date) - $startTime).TotalSeconds

# Completion banner
Write-Host ""
Write-Host " $($box.TopLeft)$($box.Horizontal * 63)$($box.TopRight)" -ForegroundColor $c.Green
Write-Host " $($box.Vertical)" -NoNewline -ForegroundColor $c.Green
Write-Host "           $($box.Check) INSTALLATION COMPLETE! $($box.Check)             " -NoNewline -ForegroundColor $c.White
Write-Host "$($box.Vertical)" -ForegroundColor $c.Green
Write-Host " $($box.BottomLeft)$($box.Horizontal * 63)$($box.BottomRight)" -ForegroundColor $c.Green
Write-Host ""
Write-Host "  $($box.Dot) Programs Processed: " -NoNewline -ForegroundColor $c.Cyan
Write-Host $total -ForegroundColor $c.White
Write-Host "  $($box.Dot) Total Time: " -NoNewline -ForegroundColor $c.Cyan
Write-Host "$([math]::Round($totalDuration, 1))s" -ForegroundColor $c.White
Write-Host ""
Write-Host " Press any key to exit..." -ForegroundColor $c.Yellow
[void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
