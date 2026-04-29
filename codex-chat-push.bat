@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\codex-chat-sync.ps1" -Action push
echo.
pause
