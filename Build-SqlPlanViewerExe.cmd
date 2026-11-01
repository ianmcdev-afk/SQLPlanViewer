@echo off
cd /d "%~dp0"
:: Runs the build in Windows PowerShell 5.1 for this one run only, and keeps the window open.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-SqlPlanViewerExe.ps1"
echo.
pause
