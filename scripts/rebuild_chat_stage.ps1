#requires -Version 7.0
param([Parameter(Mandatory)][string]$Source,[Parameter(Mandatory)][string]$SpecPath,[Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
if(Test-Path -LiteralPath $OutputPath){throw 'Choose a new output filename'}
$spec=Get-Content -LiteralPath $SpecPath -Raw|ConvertFrom-Json -AsHashtable
$x=Read-OfficeXml $Source
$ns=[Xml.XmlNamespaceManager]::new($x.NameTable);$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
function Text-Of($n){return (($n.SelectNodes('.//w:t',$ns)|ForEach-Object InnerText)-join '')}
$body=$x.SelectSingleNode('/w:document/w:body',$ns);$nodes=@($body.ChildNodes)
$a=@(0..($nodes.Count-1)|Where-Object {(Text-Of $nodes[$_]) -eq '4 結果與討論'})
$b=@(0..($nodes.Count-1)|Where-Object {(Text-Of $nodes[$_]) -eq '5 結論'})
if($a.Count -ne 1 -or $b.Count -ne 1){throw 'Ambiguous chapter boundaries'}
foreach($i in ($a[0]+1)..($b[0]-1)){if((Text-Of $nodes[$i]).Trim() -or $nodes[$i].LocalName -ne 'p'){throw 'Base source chapter is not empty'}}
$ops=[Collections.Generic.List[object]]::new()
$pcount=@($nodes[0..$a[0]]|Where-Object LocalName -eq 'p').Count
$tcount=@($nodes[0..$a[0]]|Where-Object LocalName -eq 'tbl').Count
$blankCount=$b[0]-$a[0]-1
for($i=$blankCount;$i -ge 1;$i--){$ops.Add(@{command='remove';path="/body/p[$($pcount+$i)]"})}
$script:insert=$a[0]+1;$script:pn=$pcount;$script:tn=$tcount
function Add-Runs([string]$Path,$Runs){
    foreach($r in $Runs){
        $props=@{'font.latin'='Times New Roman';'font.ea'='標楷體';size='12';'size.cs'='12';text=$r.text}
        foreach($k in $r.props.Keys){$props[$k]=$r.props[$k]}
        if($props.ContainsKey('size')){$props['size.cs']=$props.size}
        # Run text does not interpret paragraph-level \v escapes. Use a typed
        # inline break rather than leaking two literal characters into Word.
        $parts=$r.text.Split([char]11)
        for($i=0;$i -lt $parts.Count;$i++){
            if($parts[$i]){$piece=$props.Clone();$piece.text=$parts[$i];$ops.Add(@{command='add';parent=$Path;type='run';props=$piece})}
            if($i -lt $parts.Count-1){$ops.Add(@{command='add';parent=$Path;type='break';props=@{type='line'}})}
        }
    }
}
function Add-Paragraph($block){
    $props=$block.props.Clone()
    $script:pn++;$path="/body/p[$script:pn]"
    if($block.type -eq 'caption'){
        $props.text='@@CAP'+$block.caption.id+'@@'
        # Native adapter creates a real caption in this paragraph.
    }
    $ops.Add(@{command='add';parent='/body';type='paragraph';index=$script:insert;props=$props});$script:insert++
    if($block.type -ne 'caption'){
        Add-Runs $path $block.runs
        foreach($pic in $block.pictures){$ops.Add(@{command='add';parent=$path;type='picture';props=$pic})}
    }
}
foreach($block in $spec.blocks){
    if($block.type -in @('paragraph','caption')){Add-Paragraph $block;continue}
    if($block.type -ne 'table'){throw "Unsupported block $($block.type)"}
    $script:tn++;$table="/body/tbl[$script:tn]"
    $props=@{rows=[string]$block.rows.Count;cols=[string]$block.rows[0].cells.Count;layout='fixed';align='center';width=([int]($block.widths|Measure-Object -Sum).Sum).ToString();colWidths=($block.widths -join ',');'border.all'='none';'border.top'='single;6;000000';'border.bottom'='single;6;000000'}
    if($block.caption){$props.caption=$block.caption}
    $ops.Add(@{command='add';parent='/body';type='table';index=$script:insert;props=$props});$script:insert++
    for($i=0;$i -lt $block.rows.Count;$i++){
        $row=$block.rows[$i];$rowpath="$table/tr[$($i+1)]"
        $ops.Add(@{command='set';path=$rowpath;props=@{header=$row.header.ToString().ToLowerInvariant();cantSplit='true'}})
        for($j=0;$j -lt $row.cells.Count;$j++){
            $cell=$row.cells[$j];$cellpath="$rowpath/tc[$($j+1)]"
            if($cell.fill -and $cell.fill -ne 'auto'){$ops.Add(@{command='set';path=$cellpath;props=@{fill=$cell.fill}})}
            if($i -eq 0){$ops.Add(@{command='set';path=$cellpath;props=@{'border.bottom'='single;6;000000'}})}
            for($k=0;$k -lt $cell.paragraphs.Count;$k++){
                $p=$cell.paragraphs[$k];$ppath="$cellpath/p[$($k+1)]"
                if($k -eq 0){$ops.Add(@{command='set';path=$ppath;props=$p.props})}
                else{$ops.Add(@{command='add';parent=$cellpath;type='paragraph';props=$p.props})}
                Add-Runs $ppath $p.runs
            }
        }
    }
}
# Earlier case-study figures (including Fig. 8) also change in the shared chat.
# The v6 body positions before chapter 4 remain stable; replace image bytes only.
$prePn=0
foreach($n in $nodes[0..$a[0]]){
    if($n.LocalName -ne 'p'){continue};$prePn++
    $idx=[array]::IndexOf($nodes,$n)
    $replacement=@($spec.outside_pictures|Where-Object body_index -eq $idx)
    if(-not $replacement.Count){continue}
    for($j=0;$j -lt $replacement[0].paragraph.pictures.Count;$j++){
        $pic=$replacement[0].paragraph.pictures[$j]
        $runs=@($n.SelectNodes('w:r',$ns));$drawingRuns=@(0..($runs.Count-1)|Where-Object {$runs[$_].SelectSingleNode('.//w:drawing',$ns)})
        $ops.Add(@{command='set';path="/body/p[$prePn]/r[$($drawingRuns[$j]+1)]";props=@{src=$pic.src;width=$pic.width;height=$pic.height;alt=$pic.alt}})
    }
}
$full=[IO.Path]::GetFullPath($OutputPath);New-Item -ItemType Directory -Force (Split-Path $full -Parent)|Out-Null
$batch=$full+'.batch.json';$ops.ToArray()|ConvertTo-Json -Depth 20 -AsArray|Set-Content -LiteralPath $batch -Encoding utf8
Copy-Item -LiteralPath $Source -Destination $full
$r=Invoke-OfficeCli -Arguments @('batch',$full,'--input',$batch,'--json') -TimeoutSeconds 180 -AllowFailure
$r.Stdout|Set-Content ($full+'.batch-result.json') -Encoding utf8
$r.Stderr|Set-Content ($full+'.stderr.txt') -Encoding utf8
if($r.ExitCode -ne 0){throw "OfficeCLI stage failed ($($r.ExitCode)); inspect $full.batch-result.json"}
$response=$r.Stdout|ConvertFrom-Json
if(-not $response.success){throw 'Batch reported failure'}
$null=Invoke-OfficeCli -Arguments @('close',$full,'--json')
Write-Output "OfficeCLI authored $($ops.Count) operations: $full"
