powershell -ExecutionPolicy Bypass -NoProfile -Command ^
"$Startup = [Environment]::GetFolderPath('Startup'); ^
$s = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $Startup 'QuickLook.lnk')); ^
$s.TargetPath = 'C:\Program Files\QuickLook\QuickLook.exe'; ^
$s.Arguments = '/autorun'; ^
$s.Save()"