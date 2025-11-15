@echo off
echo Activating: Windows and Office
fltmc >nul || exit /b
call "%~dp0MAS_AIO.cmd" /HWID /Ohook /Z-ESU
cd \
echo Done!