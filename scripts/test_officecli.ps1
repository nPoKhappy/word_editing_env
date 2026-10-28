#requires -Version 7.0
[CmdletBinding()]
param([switch]$SkipNativeWord)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
$root = Split-Path $PSScriptRoot -Parent
$dir = Join-Path $root ('.tmp/officecli-test-' + [guid]::NewGuid())
New-Item -ItemType Directory $dir | Out-Null
$checks = [Collections.Generic.List[string]]::new()
function Check([bool]$Condition,[string]$Name) {
    if (-not $Condition) { throw "FAIL: $Name" }
    $checks.Add($Name)
    Write-Host "PASS: $Name"
}
function Run-TestCli([string[]]$A) { return (Invoke-OfficeCli -Arguments $A).Stdout }
$word = Join-Path $dir '中文 測試.docx'
$null = Run-TestCli @('create',$word,'--json')
$null = Run-TestCli @('add',$word,'/body','--type','paragraph','--prop','text=中文初稿 $5 "quoted"','--prop','font.latin=Times New Roman','--prop','font.ea=標楷體','--prop','size=12','--prop','size.cs=12','--json')
$null = Run-TestCli @('add',$word,'/body/p[1]','--type','run','--prop','text=2','--prop','subscript=true','--prop','size=8','--prop','size.cs=8','--json')
$null = Run-TestCli @('close',$word)
$xml = Read-OfficeXml $word
$ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
Check ($xml.SelectSingleNode('//w:t',$ns).InnerText -ceq '中文初稿 $5 "quoted"') 'Unicode, spaces, literal currency and quotes survive arguments'
Check ($xml.SelectSingleNode('//w:r/w:rPr/w:szCs',$ns).GetAttribute('val',$ns.LookupNamespace('w')) -eq '24') 'Explicit complex-script 12 pt'
Check ($xml.SelectSingleNode('//w:r/w:rPr/w:rFonts',$ns).GetAttribute('eastAsia',$ns.LookupNamespace('w')) -eq '標楷體') 'Chinese font slot'
$hash = (Get-FileHash $word).Hash
$badPlan = Join-Path $dir 'bad.json'
@(@{command='set';path='/body/p[1]';props=@{text='should roll back'}},@{command='set';path='/body/p[999]';props=@{text='invalid'}}) | ConvertTo-Json -AsArray -Depth 6 | Set-Content $badPlan -Encoding utf8
$bad = Invoke-OfficeCli -Arguments @('batch',$word,'--input',$badPlan,'--json') -AllowFailure
Check ($bad.ExitCode -ne 0 -and (Get-FileHash $word).Hash -eq $hash) 'Atomic batch failure leaves disk byte-identical'
$changes = Join-Path $dir 'changes.json'
@(@{old='中文初稿';new='中文修訂';count=1}) | ConvertTo-Json -AsArray | Set-Content $changes -Encoding utf8
$out = Join-Path $dir 'revised.docx'
if (-not $SkipNativeWord) {
    & (Join-Path $PSScriptRoot 'edit_office.ps1') -InputPath $word -Changes $changes -OutputPath $out
    Check ((Get-FileHash $word).Hash -eq $hash) 'Guarded editor preserves source'
    $text = Run-TestCli @('view',$out,'text','--json')
    Check ($text.Contains('中文修訂')) 'Guarded OfficeCLI replacement with native QA'
    $editedXml=Read-OfficeXml $out
    $sub=$editedXml.SelectSingleNode('//w:r[w:rPr/w:vertAlign[@w:val="subscript"]]',$ns)
    Check ($null -ne $sub -and $sub.SelectSingleNode('w:rPr/w:sz',$ns).GetAttribute('val',$ns.LookupNamespace('w')) -eq '16') 'Inline subscript and its smaller size survive editing'
}
@(@{old='missing';new='text';count=1}) | ConvertTo-Json -AsArray | Set-Content $changes -Encoding utf8
$rejected = $false
try { & (Join-Path $PSScriptRoot 'edit_office.ps1') -InputPath $word -Changes $changes -OutputPath (Join-Path $dir 'must-not-exist.docx') } catch { $rejected = $_.Exception.Message.Contains('Expected 1 editable') }
Check ($rejected -and -not (Test-Path (Join-Path $dir 'must-not-exist.docx'))) 'Count mismatch does not publish output'
$lock=Join-Path $dir ('~$' + [IO.Path]::GetFileName($word))
Set-Content -LiteralPath $lock -Value 'test lock'
$rejected=$false
try { & (Join-Path $PSScriptRoot 'edit_office.ps1') -InputPath $word -Changes $changes -OutputPath (Join-Path $dir 'locked.docx') } catch { $rejected=$_.Exception.Message.Contains('open in Office') }
finally { Remove-Item -LiteralPath $lock }
Check ($rejected -and -not (Test-Path (Join-Path $dir 'locked.docx'))) 'Open Office lock refuses saved-package edits'
$fieldDoc=Join-Path $dir 'field.docx'
$null=Run-TestCli @('create',$fieldDoc,'--json')
$null=Run-TestCli @('add',$fieldDoc,'/body','--type','paragraph','--prop','text=Protected field paragraph','--json')
$null=Run-TestCli @('add',$fieldDoc,'/body/p[1]','--type','field','--prop','fieldType=date','--json')
@(@{old='Protected';new='Changed';count=1}) | ConvertTo-Json -AsArray | Set-Content $changes -Encoding utf8
$rejected=$false
try { & (Join-Path $PSScriptRoot 'edit_office.ps1') -InputPath $fieldDoc -Changes $changes -OutputPath (Join-Path $dir 'field-damaged.docx') } catch { $rejected=$_.Exception.Message.Contains('contains a field') }
Check ($rejected -and -not (Test-Path (Join-Path $dir 'field-damaged.docx'))) 'Field-bearing target paragraph is protected'
$null = Run-TestCli @('validate',$word,'--json')
Check $true 'DOCX OpenXML schema validation'
$sheet = Join-Path $dir 'numbers.xlsx'
$null = Run-TestCli @('create',$sheet,'--json')
$null = Run-TestCli @('set',$sheet,'/Sheet1/A1','--prop','value=10','--json')
$null = Run-TestCli @('set',$sheet,'/Sheet1/A2','--prop','value=20','--json')
$null = Run-TestCli @('set',$sheet,'/Sheet1/A3','--prop','formula=SUM(A1:A2)','--json')
$cell = (Run-TestCli @('get',$sheet,'/Sheet1/A3','--json')) | ConvertFrom-Json
$cell | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $dir 'formula-readback.json')
Check ($cell.data.results[0].format.computedValue -eq '30' -and $cell.data.results[0].format.evaluated) 'Excel formula evaluates to 30'
$null = Run-TestCli @('validate',$sheet,'--json')
Check $true 'XLSX OpenXML schema validation'
$deck = Join-Path $dir 'slides.pptx'
$null = Run-TestCli @('create',$deck,'--json')
$null = Run-TestCli @('add',$deck,'/','--type','slide','--prop','title=OfficeCLI test','--json')
Check ((Run-TestCli @('view',$deck,'text','--json')).Contains('OfficeCLI test')) 'PowerPoint slide text readback'
$null = Run-TestCli @('validate',$deck,'--json')
Check $true 'PPTX OpenXML schema validation'
foreach ($file in $word,$sheet,$deck) {
    $null = Run-TestCli @('view',$file,'html','-o',($file + '.html'))
    Check ((Get-Item ($file + '.html')).Length -gt 100) ('HTML preview: ' + [IO.Path]::GetExtension($file))
}
# Real MCP initialize/list/call over stdio, independent of Codex hot-reloading.
$psi = [Diagnostics.ProcessStartInfo]::new((Join-Path $root '.tools/officecli/officecli.exe'))
$psi.ArgumentList.Add('mcp')
$psi.UseShellExecute=$false; $psi.CreateNoWindow=$true
$psi.RedirectStandardInput=$psi.RedirectStandardOutput=$psi.RedirectStandardError=$true
$psi.StandardInputEncoding=$psi.StandardOutputEncoding=[Text.UTF8Encoding]::new($false)
$psi.Environment['TEMP']=$psi.Environment['TMP']=(Join-Path $root '.tmp')
$psi.Environment['OFFICECLI_SKIP_UPDATE']='1'
$mcp=[Diagnostics.Process]::new(); $mcp.StartInfo=$psi
try {
    $null=$mcp.Start()
    $err=$mcp.StandardError.ReadToEndAsync()
    foreach ($req in @(
        @{jsonrpc='2.0';id=1;method='initialize';params=@{protocolVersion='2024-11-05';capabilities=@{};clientInfo=@{name='officecli-integration-test';version='1'}}},
        @{jsonrpc='2.0';id=2;method='tools/list'},
        @{jsonrpc='2.0';id=3;method='tools/call';params=@{name='officecli';arguments=@{command='help docx paragraph --json'}}}
    )) {
        $mcp.StandardInput.WriteLine(($req | ConvertTo-Json -Compress -Depth 10)); $mcp.StandardInput.Flush()
        $read=$mcp.StandardOutput.ReadLineAsync()
        if (-not $read.Wait(15000)) { throw 'MCP response timed out.' }
        $response=$read.Result | ConvertFrom-Json
        Check ($response.id -eq $req.id -and -not $response.error) ('MCP ' + $req.method)
        if ($req.id -eq 2) { Check ($response.result.tools[0].name -eq 'officecli') 'MCP official officecli tool discovered' }
        if ($req.id -eq 3) { Check (-not $response.result.isError) 'MCP actual tool call succeeds' }
    }
} finally { if (-not $mcp.HasExited) { $mcp.Kill() }; $mcp.Dispose() }
@{status='PASS';nativeWordTested=(-not $SkipNativeWord);checks=$checks.ToArray();evidence=$dir} | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $dir 'report.json') -Encoding utf8
Write-Host "All $($checks.Count) checks passed. Evidence: $dir"
