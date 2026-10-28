# Dot-source this file in PowerShell 7: . ./scripts/env.ps1
$officeWorkspace = Split-Path $PSScriptRoot -Parent
$officeBin = Join-Path $officeWorkspace '.tools/officecli'
if (($env:PATH -split ';') -notcontains $officeBin) { $env:PATH = "$officeBin;$env:PATH" }
$env:TEMP = $env:TMP = Join-Path $officeWorkspace '.tmp'
New-Item -ItemType Directory -Force $env:TEMP | Out-Null
$env:OFFICECLI_SKIP_UPDATE = '1'
$env:OFFICECLI_NO_AUTO_RESIDENT = '1'
$env:OFFICECLI_RESIDENT_FLUSH = 'each'
$env:OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT = '1'
$env:PIP_CACHE_DIR = Join-Path $officeWorkspace '.tmp/pip-cache'
$env:MPLCONFIGDIR = Join-Path $officeWorkspace '.tmp/matplotlib'
$env:PYTHONDONTWRITEBYTECODE = '1'
