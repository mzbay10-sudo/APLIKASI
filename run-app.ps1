$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir
& '.\.venv\Scripts\pythonw.exe' launcher.py
