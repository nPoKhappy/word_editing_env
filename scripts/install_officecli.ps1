#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$lock = Get-Content (Join-Path $root 'officecli.lock.json') -Raw | ConvertFrom-Json
$arch = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
$asset = $lock.assets.$arch
if (-not $asset) { throw "Unsupported architecture: $arch" }
$dir = Join-Path $root '.tools/officecli'
New-Item -ItemType Directory -Force $dir | Out-Null
$binary = Join-Path $dir 'officecli.exe'
if ((Test-Path $binary) -and (Get-FileHash $binary -Algorithm SHA256).Hash -eq $asset.sha256) {
    Write-Host "OfficeCLI $($lock.version) is installed and checksum verified."
    return
}
$url = "https://github.com/$($lock.repository)/releases/download/v$($lock.version)/$($asset.name)"
$download = Join-Path $dir ('download-' + [guid]::NewGuid() + '.exe')
try {
    Invoke-WebRequest -Uri $url -OutFile $download -TimeoutSec 120
    if ((Get-FileHash $download -Algorithm SHA256).Hash -ne $asset.sha256) {
        throw 'OfficeCLI SHA256 mismatch; installation aborted.'
    }
    Move-Item -LiteralPath $download -Destination $binary -Force
    Write-Host "Installed OfficeCLI $($lock.version): $binary"
} finally {
    if (Test-Path $download) { Remove-Item -LiteralPath $download }
}
