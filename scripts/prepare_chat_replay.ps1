#requires -Version 7.0
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$run = Join-Path $root 'outputs/chat-replay-20261028'
$source = 'D:\word_editing_env'
foreach ($dir in 'inputs','evidence','stages','render','plots','specs','logs') {
    New-Item -ItemType Directory -Force (Join-Path $run $dir) | Out-Null
}
$manifest = [Collections.Generic.List[object]]::new()
function Get-SharedHash([string]$Path) {
    $stream = [IO.File]::Open($Path,'Open','Read',[IO.FileShare]::ReadWrite)
    try { return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) }
    finally { $stream.Dispose() }
}
function Copy-Evidence([string]$From,[string]$To) {
    $target = Join-Path $run $To
    New-Item -ItemType Directory -Force (Split-Path $target -Parent) | Out-Null
    if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $From -Destination $target }
    $before = Get-SharedHash $From
    $after = Get-SharedHash $target
    if ($before -ne $after) { throw "Source copy differs: $From" }
    $manifest.Add(@{source=$From;copy=$target;sha256=$before})
}
$paper = Join-Path $source 'thesis_outline/jtice_paper'
Get-ChildItem -LiteralPath $paper -File -Filter '*.docx' | Where-Object Name -Match '_v(6|7|8|9|10|11|12|13|14|15)_' | ForEach-Object {
    Copy-Evidence $_.FullName ('inputs/' + $_.Name)
}
Get-ChildItem -LiteralPath (Join-Path $source '.tmp') -File | Where-Object { $_.Name -match '^(jtice_.*\.(py|json)|chapter4_content.json|rmse_tables_cases.json|create_chapter4_content.py)$' } | ForEach-Object {
    Copy-Evidence $_.FullName ('evidence/original-scripts/' + $_.Name)
}
Get-ChildItem -LiteralPath (Join-Path $paper 'assets') -Recurse -File | ForEach-Object {
    Copy-Evidence $_.FullName ('inputs/assets/' + [IO.Path]::GetRelativePath((Join-Path $paper 'assets'),$_.FullName))
}
Copy-Evidence 'C:\Users\Administrator\Desktop\meeting ppts\20261006.pptx' 'inputs/20261006.pptx'
Copy-Evidence (Join-Path $paper 'JTICE_驗證案例原始RMSE_補充資料.csv') 'inputs/JTICE_驗證案例原始RMSE_補充資料.csv'
$dataRoot = 'C:\Users\Administrator\Desktop\digital_twin_revised\results\visualization\gradnorm_ratio0p005_0p01_0p015_0p02_seed42'
Get-ChildItem -LiteralPath $dataRoot -File -Filter '*.csv' | ForEach-Object { Copy-Evidence $_.FullName ('inputs/data/' + $_.Name) }
$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $run 'evidence/source-hashes.json') -Encoding utf8
Write-Output "Copied and hashed $($manifest.Count) inputs into $run"
