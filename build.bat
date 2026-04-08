@echo off
setlocal

cd /d "%~dp0"

echo Building FillDoc...

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name FillDoc ^
  --paths src ^
  --collect-all PySide6 ^
  --hidden-import fitz ^
  main.py

echo.
echo Build complete:
echo   dist\FillDoc\FillDoc.exe

endlocal
