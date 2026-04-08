$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

function Step($message) {
    Write-Host ""
    Write-Host "==> $message" -ForegroundColor Cyan
}

function Fail($message) {
    Write-Host ""
    Write-Host "ERROR: $message" -ForegroundColor Red
    exit 1
}

Step "Checking Python"
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Fail "Python is not available in PATH. Install Python 3.11+ and reopen PowerShell."
}

$pythonVersion = python --version 2>&1
Write-Host $pythonVersion

Step "Checking required project files"
if (-not (Test-Path "main.py")) {
    Fail "main.py not found. Run this script from the project root."
}
if (-not (Test-Path "src")) {
    Fail "src folder not found. Run this script from the project root."
}

Step "Checking PyInstaller"
$pyInstallerInstalled = $false

python -m pip show pyinstaller *> $null
if ($LASTEXITCODE -eq 0) {
    $pyInstallerInstalled = $true
}

if (-not $pyInstallerInstalled) {
    Write-Host "PyInstaller is not installed. Installing it automatically..." -ForegroundColor Yellow
    python -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Fail "Automatic PyInstaller installation failed. Run: python -m pip install pyinstaller"
    }
    Write-Host "PyInstaller installed successfully." -ForegroundColor Green
}

Step "Building FillDoc"
python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --name FillDoc `
  --paths src `
  --collect-all PySide6 `
  --hidden-import fitz `
  main.py

if ($LASTEXITCODE -ne 0) {
    Fail "Build failed. Read the messages above to see the cause."
}

$exePath = Join-Path $PSScriptRoot "dist/FillDoc/FillDoc.exe"

Step "Checking build result"
if (-not (Test-Path $exePath)) {
    Fail "Build finished, but FillDoc.exe was not found in dist/FillDoc."
}

Write-Host "Build complete." -ForegroundColor Green
Write-Host "EXE file: $exePath"
