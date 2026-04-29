@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tools\codex-chat-sync.ps1" -Action repair-paths
uv run --no-project python ".\tools\codex-sqlite-path-repair.py"
echo.
pause
