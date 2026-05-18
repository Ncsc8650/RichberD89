$ErrorActionPreference = "Stop"
$python = "C:\Users\ditsa\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
Set-Location $PSScriptRoot
& $python app.py
