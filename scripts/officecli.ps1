#requires -Version 7.0
param([Parameter(ValueFromRemainingArguments=$true)][string[]]$CommandArgs)
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
if (-not $CommandArgs) { $CommandArgs = @('--help') }
$result = Invoke-OfficeCli -Arguments $CommandArgs -AllowFailure
[Console]::Out.Write($result.Stdout)
[Console]::Error.Write($result.Stderr)
exit $result.ExitCode
