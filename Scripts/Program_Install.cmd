@echo off
REM Changing working directory to location of installer:
cd /D "%~dp0"

echo Installing: K-Lite Mega Codec Pack
REM K-Lite_Codec_Pack_1930_Mega.exe
"%WINDIR%\Setup\Files\K-Lite_Codec_Pack_Mega.exe" /VERYSILENT /NORESTART /SUPPRESSMSGBOXES /LOADINF="%WINDIR%\Setup\Files\klcp_mega_unattended.ini"
echo K-Lite Mega Codec Pack Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: Google Chrome
msiexec /i "%WINDIR%\Setup\Files\googlechromestandaloneenterprise64.msi" /qb /norestart
echo Google Chrome Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: QuickLook
REM QuickLook-4.2.2.msi
msiexec /i "QuickLook.msi" INSTALLFOLDER="C:\Program Files (x86)\QuickLook" ALLUSERS=1 /qb /norestart
echo QuickLook Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: FileConverter
REM FileConverter-2.1-x64-setup
msiexec /i "FileConverter.msi" /qb /norestart
echo FileConverter Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: 7zip
REM 7z2501-x64.msi
msiexec /i "%WINDIR%\Setup\Files\7-zip.msi" /qb /norestart
echo 7zip Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: AnyDesk
"%WINDIR%\Setup\Files\AnyDesk.exe" --install "C:\Program Files (x86)\AnyDesk" --silent --create-shortcuts --create-desktop-icon
echo AnyDesk Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: Everything
REM Everything-1.4.1.1030.x64-Setup
"%WINDIR%\Setup\Files\Everything.exe" /S -install-options "-app-data -disable-run-as-admin -install-all-users-desktop-shortcut -install-efu-association install-quick-launch-shortcut -install-all-users-start-menu-shortcuts -install-folder-context-menu -install-run-on-system-startup" /D="C:\Program Files\Everything"
echo Everything Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: VLC Media Player
REM vlc-3.0.21-win64.exe
"%WINDIR%\Setup\Files\vlc-3.0.21-win64.exe" /L=1033 /S
echo VLC Media Player Installed...!

TIMEOUT /T 5 /nobreak
echo.

echo Installing: WinRAR
REM winrar-x64-713.exe
"%WINDIR%\Setup\Files\winrar.exe" /S
echo WinRAR Installed...!
