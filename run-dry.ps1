$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir
& '.\.venv\Scripts\python.exe' bot.py --dry-run --input data/input.csv

