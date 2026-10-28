#requires -Version 7.0
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
$root=Split-Path $PSScriptRoot -Parent
$run=Join-Path $root 'outputs/chat-replay-20261028'
$source=Join-Path $run 'stages/v15-officecli-03.docx'
$delivery=Join-Path $run 'delivery'
New-Item -ItemType Directory -Force $delivery|Out-Null
$out=Join-Path $delivery 'JTICE_OfficeCLI_聊天室復現完成版.docx'
if(Test-Path -LiteralPath $out){throw 'Delivery exists; inspect before regenerating'}
$x=Read-OfficeXml $source
$ns=[Xml.XmlNamespaceManager]::new($x.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$ns.AddNamespace('wp','http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
$paras=@($x.SelectNodes('/w:document/w:body/w:p',$ns))
function Text-Of($n){return (($n.SelectNodes('.//w:t',$ns)|ForEach-Object InnerText)-join '')}
$jobs=@(
    @{caption='八組近外插情境之操作設定值';files=@('five-model/nearseries.png')},
    @{caption='穩態 ANN 與 Aspen 預測值比較';files=@('five-model/ann_parity.png')},
    @{caption='不同 ρ 之內部驗證 RMSE';files=@('five-model/rmse_internal_validation.png')},
    @{caption='不同 ρ 之外部驗證 RMSE';files=@('five-model/rmse_external_validation.png')},
    @{caption='各 gain pair 的 all-18 KCI';files=@('four-model/four_model_all18_kci.png')},
    @{caption='依參考增益正負方向分組的 KCI';files=@('four-model/four_model_conditional_kci_by_teacher_sign.png')},
    @{caption='依參考增益大小分組的 KCI';files=@('four-model/four_model_conditional_kci_by_gain_magnitude.png')},
    @{caption='核心區域測試 RMSE';files=@('four-model/four_model_rmse_internal_test.png')},
    @{caption='近外插測試 RMSE';files=@('four-model/four_model_rmse_external_test_a_d.png','four-model/four_model_rmse_external_test_e_h.png')}
)
$ops=[Collections.Generic.List[object]]::new();$audit=[Collections.Generic.List[object]]::new()
Add-Type -AssemblyName System.Drawing
foreach($job in $jobs){
    $caps=@(0..($paras.Count-1)|Where-Object {(Text-Of $paras[$_]).Contains($job.caption) -and $paras[$_].SelectNodes('.//w:instrText[contains(text(),"SEQ")] | .//w:fldSimple[contains(@w:instr,"SEQ")]',$ns).Count})
    if($caps.Count -ne 1){throw "Caption ambiguous: $($job.caption)"}
    $pictures=@(0..($caps[0]-1)|Where-Object {$paras[$_].SelectNodes('.//w:drawing',$ns).Count})
    $selected=@($pictures|Select-Object -Last $job.files.Count)
    for($j=0;$j -lt $selected.Count;$j++){
        $pn=$selected[$j];$runs=@($paras[$pn].SelectNodes('w:r',$ns));$rn=@(0..($runs.Count-1)|Where-Object {$runs[$_].SelectNodes('.//w:drawing',$ns).Count})
        if($rn.Count -ne 1){throw 'Ambiguous drawing run'}
        $extent=$runs[$rn[0]].SelectSingleNode('.//wp:extent',$ns)
        $width=[double]$extent.GetAttribute('cx')/12700
        $file=Join-Path $run ('plots/'+$job.files[$j]);$im=[Drawing.Image]::FromFile($file)
        try{$height=$width*$im.Height/$im.Width}finally{$im.Dispose()}
        $path="/body/p[$($pn+1)]/r[$($rn[0]+1)]"
        $ops.Add(@{command='set';path=$path;props=@{src=$file;width=($width.ToString([cultureinfo]::InvariantCulture)+'pt');height=($height.ToString([cultureinfo]::InvariantCulture)+'pt');alt=$job.caption}})
        $audit.Add(@{caption=$job.caption;picture=$file;sha256=(Get-FileHash -LiteralPath $file).Hash;officecli_path=$path})
    }
}
$batch=Join-Path $delivery 'regenerated-figure-batch.json'
$ops.ToArray()|ConvertTo-Json -Depth 8 -AsArray|Set-Content $batch -Encoding utf8
Copy-Item -LiteralPath $source -Destination $out
$result=Invoke-OfficeCli -Arguments @('batch',$out,'--input',$batch,'--json')
$result.Stdout|Set-Content (Join-Path $delivery 'figure-batch-result.json') -Encoding utf8
$null=Invoke-OfficeCli -Arguments @('close',$out,'--json')
$audit.ToArray()|ConvertTo-Json -Depth 6|Set-Content (Join-Path $delivery 'regenerated-figure-provenance.json') -Encoding utf8
Copy-Item -LiteralPath (Join-Path $run 'inputs/JTICE_驗證案例原始RMSE_補充資料.csv') -Destination $delivery -ErrorAction SilentlyContinue
Write-Output "OfficeCLI inserted $($ops.Count) newly generated figures into $out"
