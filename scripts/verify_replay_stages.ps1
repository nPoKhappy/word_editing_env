#requires -Version 7.0
param([int[]]$Versions=@(7,8,9,10,11,12,13,14,15))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'env.ps1')
$run=Join-Path $root 'outputs/chat-replay-20261028'
$python=Join-Path $root '.venv/Scripts/python.exe'
foreach($v in $Versions){
    $file=Join-Path $run $(if($v -eq 15){'stages/v15-officecli-03.docx'}elseif($v -eq 12){'stages/v12-officecli-02.docx'}else{"stages/v$v-officecli.docx"})
    $native=Join-Path $run "logs/v$v-native.json"
    $nativeMatches=(Test-Path -LiteralPath $native) -and ((Get-Content -LiteralPath $native -Raw|ConvertFrom-Json).document -eq $file)
    if(-not $nativeMatches){
        & $python -X utf8 (Join-Path $PSScriptRoot 'replay_native_fields.py') $file --spec (Join-Path $run "specs/v$v.json") --report $native
        if($LASTEXITCODE -ne 0){throw "Native stage v$v failed"}
    }
    & $python -X utf8 (Join-Path $PSScriptRoot 'finalize_word.py') $file --caption-spec (Join-Path $run "specs/v$v.json")
    if($LASTEXITCODE -ne 0){throw "Finalization v$v failed"}
    & $python -X utf8 (Join-Path $PSScriptRoot 'verify_replay_table_fonts.py') $file --report (Join-Path $run "logs/v$v-table-fonts.json")
    if($LASTEXITCODE -ne 0){throw "Directory fonts v$v failed"}
    $reference=(Get-ChildItem (Join-Path $run 'inputs') -Filter "*_v${v}_*.docx").FullName
    & $python -X utf8 (Join-Path $PSScriptRoot 'audit_chat_stage.py') $file --reference $reference --report (Join-Path $run "logs/v$v-comparison.json")
    if($LASTEXITCODE -ne 0){throw "Semantic comparison v$v failed"}
    Write-Output "VERIFIED stage v$v"
}
