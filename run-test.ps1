$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir
& '.\.venv\Scripts\python.exe' bot.py --max-rows 1
