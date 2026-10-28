#requires -Version 7.0
param([Parameter(Mandatory)][string]$Reference,[Parameter(Mandatory)][string]$SpecPath)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
$x=Read-OfficeXml $Reference
$ns=[Xml.XmlNamespaceManager]::new($x.NameTable)
$w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$ns.AddNamespace('w',$w)
$ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
$ns.AddNamespace('wp','http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
$ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
function Text-Of($n) { return (($n.SelectNodes('.//w:t',$ns)|ForEach-Object InnerText)-join '') }
function Attr($n,$key) { if($null -eq $n){return ''};return $n.GetAttribute($key,$w) }
function Bool-Of($n) { return ($null -ne $n -and (Attr $n 'val') -notin @('0','false','off')) }
$body=$x.SelectSingleNode('/w:document/w:body',$ns)
$nodes=@($body.ChildNodes)
$start=@(0..($nodes.Count-1)|Where-Object { (Text-Of $nodes[$_]) -eq '4 結果與討論' })
$end=@(0..($nodes.Count-1)|Where-Object { (Text-Of $nodes[$_]) -eq '5 結論' })
if($start.Count -ne 1 -or $end.Count -ne 1){throw 'Chapter boundaries ambiguous'}
$captionMap=@{};$captions=[Collections.Generic.List[object]]::new()
foreach($p in $nodes){
    if($p.LocalName -ne 'p'){continue}
    $seq=@($p.SelectNodes('.//w:instrText',$ns)|Where-Object InnerText -Match '^\s*SEQ\s+(Fig\.|Table)')
    $simpleSeq=@($p.SelectNodes('.//w:fldSimple',$ns)|Where-Object { (Attr $_ 'instr') -match '^\s*SEQ\s+(Fig\.|Table)' })
    if($simpleSeq.Count){$seq+=@($simpleSeq|ForEach-Object { [pscustomobject]@{InnerText=(Attr $_ 'instr')} })}
    if(-not $seq.Count){continue}
    $text=Text-Of $p
    if($text -notmatch '^(Fig\.|Table)\s+(\d+)([ab]?)\s*(.*)$'){throw "Caption parse: $text"}
    $c=@{id=('c'+$captions.Count);label=$Matches[1];number=[int]$Matches[2];suffix=$Matches[3];title=$Matches[4];repeat=($seq[0].InnerText -match '\\c\b');text=$text;in_chapter=([array]::IndexOf($nodes,$p) -gt $start[0])}
    $captions.Add($c);$captionMap[$p.GetHashCode()]=$c
}
$bookmarkMap=@{}
foreach($b in $x.SelectNodes('//w:bookmarkStart',$ns)){
    $p=$b.ParentNode
    while($null -ne $p -and $p.LocalName -ne 'p'){$p=$p.ParentNode}
    if($null -ne $p -and $captionMap.ContainsKey($p.GetHashCode())){$bookmarkMap[(Attr $b 'name')]=$captionMap[$p.GetHashCode()].id}
}
$refs=[Collections.Generic.List[object]]::new();$pictures=[Collections.Generic.List[object]]::new()
$assetDir=Join-Path (Split-Path $SpecPath -Parent) ([IO.Path]::GetFileNameWithoutExtension($SpecPath)+'-assets')
New-Item -ItemType Directory -Force $assetDir | Out-Null
$rels=Read-OfficeXml $Reference 'word/_rels/document.xml.rels'
$relmap=@{};foreach($r in $rels.DocumentElement.ChildNodes){$relmap[$r.GetAttribute('Id')]=$r.GetAttribute('Target')}
$zip=[IO.Compression.ZipFile]::OpenRead($Reference)
function Runs-Of($p) {
    $result=[Collections.Generic.List[object]]::new();$depth=0;$code=''
    foreach($r in $p.SelectNodes('.//w:r[not(ancestor::w:fldSimple)]|.//w:fldSimple',$ns)){
        if($r.LocalName -eq 'fldSimple'){
            $simpleCode=Attr $r 'instr'
            if($simpleCode -notmatch '^\s*REF\s+(\S+)'){throw "Unsupported simple field: $simpleCode"}
            $bookmark=$Matches[1]
            if(-not $bookmarkMap.ContainsKey($bookmark)){throw "Unmapped simple reference: $bookmark"}
            $token='@@REF'+$refs.Count+'@@'
            $refs.Add(@{token=$token;caption=$bookmarkMap[$bookmark];original=$simpleCode})
            $result.Add(@{text=$token;props=@{}});continue
        }
        $fc=$r.SelectSingleNode('w:fldChar',$ns)
        if($null -ne $fc){
            switch(Attr $fc 'fldCharType'){
                'begin' {$depth++;if($depth -eq 1){$code=''}}
                'end' {
                    $depth--
                    if($depth -eq 0){
                        if($code -notmatch '^\s*REF\s+(\S+)'){throw "Unsupported chapter field: $code"}
                        $bookmark=$Matches[1]
                        if(-not $bookmarkMap.ContainsKey($bookmark)){throw "Unmapped chapter reference: $bookmark"}
                        $token='@@REF'+$refs.Count+'@@'
                        $refs.Add(@{token=$token;caption=$bookmarkMap[$bookmark];original=$code})
                        $result.Add(@{text=$token;props=@{}})
                    }
                }
            };continue
        }
        $instr=$r.SelectSingleNode('w:instrText',$ns)
        if($null -ne $instr){$code+=$instr.InnerText;continue}
        if($depth -gt 0){continue}
        $t=''
        foreach($child in $r.ChildNodes){
            if($child.LocalName -eq 't'){$t+=$child.InnerText}
            elseif($child.LocalName -eq 'br'){$t+=[char]11}
            elseif($child.LocalName -eq 'tab'){$t+="`t"}
        }
        if(-not $t){continue}
        $props=@{}
        foreach($pair in @(@('b','bold'),@('i','italic'))){$n=$r.SelectSingleNode('w:rPr/w:'+$pair[0],$ns);if($null -ne $n){$props[$pair[1]]=(Bool-Of $n).ToString().ToLowerInvariant()}}
        $va=Attr ($r.SelectSingleNode('w:rPr/w:vertAlign',$ns)) 'val'
        if($va){$props['vertAlign']=$va}
        $sz=Attr ($r.SelectSingleNode('w:rPr/w:sz',$ns)) 'val';if($sz){$props.size=([double]$sz/2).ToString([cultureinfo]::InvariantCulture)}
        $color=Attr ($r.SelectSingleNode('w:rPr/w:color',$ns)) 'val';if($color -and $color -ne 'auto'){$props.color=$color}
        $result.Add(@{text=$t;props=$props})
    }
    return ,$result.ToArray()
}
function Para-Of($p) {
    $props=@{style='Normal';'font.latin'='Times New Roman';'font.ea'='標楷體';size='12';'size.cs'='12'}
    $style=Attr ($p.SelectSingleNode('w:pPr/w:pStyle',$ns)) 'val';if($style){$props.style=$style}
    $jc=Attr ($p.SelectSingleNode('w:pPr/w:jc',$ns)) 'val';if($jc){$props.align=$jc}
    foreach($key in 'keepNext','keepLines','pageBreakBefore'){
        $n=$p.SelectSingleNode('w:pPr/w:'+$key,$ns)
        if($null -ne $n){$props[$key]=(Bool-Of $n).ToString().ToLowerInvariant()}
    }
    $sp=$p.SelectSingleNode('w:pPr/w:spacing',$ns)
    foreach($key in 'before','after'){ $v=Attr $sp $key;if($v){$props[@{before='spaceBefore';after='spaceAfter'}[$key]]=([double]$v/20).ToString([cultureinfo]::InvariantCulture)+'pt'} }
    $line=Attr $sp 'line';$rule=Attr $sp 'lineRule'
    if($line){$props.lineSpacing=if($rule -in @('exact','atLeast')){([double]$line/20).ToString([cultureinfo]::InvariantCulture)+'pt'}else{([double]$line/240).ToString([cultureinfo]::InvariantCulture)+'x'};if($rule){$props.lineRule=$rule}}
    $ind=$p.SelectSingleNode('w:pPr/w:ind',$ns)
    foreach($key in 'left','right','firstLine'){ $v=Attr $ind $key;if($v){$props[@{left='indent';right='rightIndent';firstLine='firstLineIndent'}[$key]]=$v} }
    $pics=[Collections.Generic.List[object]]::new()
    foreach($draw in $p.SelectNodes('.//w:drawing',$ns)){
        $blip=$draw.SelectSingleNode('.//a:blip',$ns);$rel=$blip.GetAttribute('embed',$ns.LookupNamespace('r'))
        $part='word/'+$relmap[$rel];$part=$part.Replace('word/../','')
        $entry=$zip.GetEntry($part);if($null -eq $entry){throw "Missing image: $part"}
        $file=Join-Path $assetDir ('image-'+$pictures.Count+[IO.Path]::GetExtension($part))
        $stream=$entry.Open();$dest=[IO.File]::Create($file);try{$stream.CopyTo($dest)}finally{$stream.Dispose();$dest.Dispose()}
        $ext=$draw.SelectSingleNode('.//wp:extent',$ns)
        $pic=@{src=$file;width=([double]$ext.GetAttribute('cx')/12700).ToString([cultureinfo]::InvariantCulture)+'pt';height=([double]$ext.GetAttribute('cy')/12700).ToString([cultureinfo]::InvariantCulture)+'pt';alt='Research result figure'}
        $pics.Add($pic);$pictures.Add($pic)
    }
    $c=$captionMap[$p.GetHashCode()]
    if($null -ne $c){return @{type='caption';caption=$c;props=$props}}
    return @{type='paragraph';props=$props;runs=(Runs-Of $p);pictures=$pics.ToArray();text=(Text-Of $p)}
}
$blocks=[Collections.Generic.List[object]]::new()
try{
    for($i=$start[0]+1;$i -lt $end[0];$i++){
        $n=$nodes[$i]
        if($n.LocalName -eq 'p'){$blocks.Add((Para-Of $n));continue}
        if($n.LocalName -ne 'tbl'){throw "Unsupported chapter block: $($n.LocalName)"}
        if($n.SelectNodes('.//w:gridSpan|.//w:vMerge',$ns).Count){throw 'Merged cells require an explicit replay recipe'}
        $rows=[Collections.Generic.List[object]]::new()
        foreach($row in $n.SelectNodes('w:tr',$ns)){
            $cells=[Collections.Generic.List[object]]::new()
            foreach($cell in $row.SelectNodes('w:tc',$ns)){
                $paras=@($cell.SelectNodes('w:p',$ns)|ForEach-Object { Para-Of $_ })
                $cells.Add(@{paragraphs=$paras;fill=(Attr ($cell.SelectSingleNode('w:tcPr/w:shd',$ns)) 'fill')})
            }
            $rows.Add(@{cells=$cells.ToArray();header=($null -ne $row.SelectSingleNode('w:trPr/w:tblHeader',$ns))})
        }
        $widths=@($n.SelectNodes('w:tblGrid/w:gridCol',$ns)|ForEach-Object {Attr $_ 'w'})
        $blocks.Add(@{type='table';rows=$rows.ToArray();widths=$widths;caption=(Attr ($n.SelectSingleNode('w:tblPr/w:tblCaption',$ns)) 'val')})
    }
    $outside=@()
    for($i=0;$i -le $start[0];$i++){
        if($nodes[$i].LocalName -eq 'p' -and $nodes[$i].SelectNodes('.//w:drawing',$ns).Count){$outside+=@{body_index=$i;paragraph=(Para-Of $nodes[$i])}}
    }
    $spec=@{reference=[IO.Path]::GetFullPath($Reference);blocks=$blocks.ToArray();captions=$captions.ToArray();references=$refs.ToArray();outside_pictures=$outside;source_chapter_index=$start[0];source_end_index=$end[0]}
    $spec|ConvertTo-Json -Depth 35|Set-Content -LiteralPath $SpecPath -Encoding utf8
    Write-Output "Exported $($blocks.Count) chapter blocks, $($refs.Count) native reference mappings, $($pictures.Count) images: $SpecPath"
}finally{$zip.Dispose()}
