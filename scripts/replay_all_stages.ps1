#requires -Version 7.0
param([int[]]$Versions=@(7,8,9,10,11,12,13,14,15))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$run=Join-Path $root 'outputs/chat-replay-20261028'
$source=(Get-ChildItem (Join-Path $run 'inputs') -Filter '*_v6_*.docx').FullName
foreach($v in $Versions){
    $ref=@(Get-ChildItem (Join-Path $run 'inputs') -Filter "*_v${v}_*.docx")
    if($ref.Count -ne 1){throw "Version $v ambiguous"}
    $spec=Join-Path $run "specs/v$v.json"
    $output=Join-Path $run $(if($v -eq 12){'stages/v12-officecli-02.docx'}elseif($v -eq 15){'stages/v15-officecli-03.docx'}else{"stages/v$v-officecli.docx"})
    if((Test-Path -LiteralPath $output) -and (Test-Path -LiteralPath ($output+'.batch-result.json'))){
        $result=Get-Content ($output+'.batch-result.json') -Raw|ConvertFrom-Json
        if($result.success){Write-Output "Stage v$v already authored";continue}
    }
    & (Join-Path $PSScriptRoot 'export_chat_stage.ps1') -Reference $ref[0].FullName -SpecPath $spec
    & (Join-Path $PSScriptRoot 'rebuild_chat_stage.ps1') -Source $source -SpecPath $spec -OutputPath $output
}
