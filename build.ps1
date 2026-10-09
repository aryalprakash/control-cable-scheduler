# Local Windows build helper
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path .venv)) {
    python -m venv .venv
}
& .\.venv\Scripts\Activate.ps1
pip install -q -r requirements-build.txt
if (Test-Path build) { Remove-Item -Recurse -Force build }
if (Test-Path dist) { Remove-Item -Recurse -Force dist }
pyinstaller --noconfirm control_cable_scheduler.spec

Write-Host ""
Write-Host "Build complete."
Write-Host "  Dir: dist\Control Cable Scheduler\"
Write-Host "  Exe: dist\Control Cable Scheduler\Control Cable Scheduler.exe"
