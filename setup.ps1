$ErrorActionPreference = 'Stop'
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir

if (-not (Test-Path -LiteralPath '.venv')) {
    py -m venv .venv
}
& '.\.venv\Scripts\python.exe' -m pip install -r requirements.txt
if (-not (Test-Path -LiteralPath 'config.json')) {
    Copy-Item -LiteralPath 'config.example.json' -Destination 'config.json'
}
Write-Host 'Selesai. Bot akan memakai Google Chrome yang sudah terpasang. Edit config.json, lalu jalankan run-login.ps1.'
