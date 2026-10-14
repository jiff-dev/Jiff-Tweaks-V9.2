@echo off
rem Jiff Tweaks launcher. Jiff-Tweaks.ps1 must be in the same folder.
rem The script asks for administrator rights itself.
cd /d "%~dp0"
if not exist "Jiff-Tweaks.ps1" (
  echo Jiff-Tweaks.ps1 was not found next to this file.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Jiff-Tweaks.ps1"
