$ErrorActionPreference = 'Stop'
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir
& '.\.venv\Scripts\python.exe' -m pip install pyinstaller
& '.\.venv\Scripts\python.exe' -m PyInstaller --noconfirm --clean --onefile --windowed --name FlowBot-v7 --distpath . --workpath build --specpath build launcher.py
Write-Host 'EXE selesai dibuat: FlowBot-v7.exe (di samping config.json dan profiles.json)'
