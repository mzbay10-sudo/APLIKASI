$ErrorActionPreference = 'Stop'
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProfileDir = Join-Path $ProjectDir 'runtime\browser-profile'
$FlowUrl = 'https://labs.google/fx/tools/flow'

$ChromeCandidates = @(
    (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
    (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
    (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
)
$ChromeExe = $ChromeCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $ChromeExe) {
    throw 'Google Chrome tidak ditemukan. Instal Chrome biasa terlebih dahulu.'
}

New-Item -ItemType Directory -Path $ProfileDir -Force | Out-Null
Write-Host 'Chrome normal akan dibuka tanpa Playwright.'
Write-Host 'Login ke Google dan buka Flow, lalu TUTUP seluruh jendela Chrome profil bot sebelum menjalankan bot.'
Start-Process -FilePath $ChromeExe -ArgumentList @("--user-data-dir=$ProfileDir", '--no-first-run', $FlowUrl)
