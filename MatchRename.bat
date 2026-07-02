@echo off
setlocal EnableDelayedExpansion
set "APM_CMDLINE=!CMDCMDLINE!"
set "APM_BAT=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0MatchRename.ps1"
echo.
pause
exit
