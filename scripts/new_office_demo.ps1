#requires -Version 7.0
[CmdletBinding()]
param([string]$OutputDirectory = ('outputs/officecli-demo/' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
$dir = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path $dir) { throw 'Choose a new demo directory.' }
New-Item -ItemType Directory -Path $dir -Force | Out-Null
function Run-Cli([string[]]$CliArgs) { $null = Invoke-OfficeCli -Arguments $CliArgs }
function Run-Batch([string]$File,[object[]]$Items) {
    $batchPath = $File + '.batch.json'
    $Items | ConvertTo-Json -AsArray -Depth 12 | Set-Content $batchPath -Encoding utf8
    Run-Cli @('batch',$File,'--input',$batchPath,'--json')
    Run-Cli @('close',$File,'--json')
}
$word = Join-Path $dir 'word-source.docx'
Run-Cli @('create',$word,'--json')
function WordProps([hashtable]$Extra) {
    $props = @{'font.latin'='Times New Roman';'font.ea'='標楷體';size='12';'size.cs'='12'}
    foreach ($key in $Extra.Keys) { $props[$key] = $Extra[$key] }
    return $props
}
$ops = [Collections.Generic.List[object]]::new()
$ops.Add(@{command='set';path='/styles/Normal';props=@{font='Times New Roman';'font.ea'='標楷體';size='12'}})
$ops.Add(@{command='raw-set';part='/styles';xpath='//w:style[@w:styleId="Normal"]/w:rPr';action='append';xml='<w:szCs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="24" />'})
$ops.Add(@{command='add';parent='/styles';type='style';props=@{id='Heading1';name='heading 1';basedOn='Normal';size='18';bold='true'}})
$ops.Add(@{command='add';parent='/styles';type='style';props=@{id='TOC1';name='toc 1';basedOn='Normal';size='12'}})
$ops.Add(@{command='raw-set';part='/styles';xpath='//w:style[@w:styleId="TOC1"]/w:rPr';action='append';xml='<w:szCs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="24" />'})
$ops.Add(@{command='add';parent='/body';type='paragraph';props=(WordProps @{text='OfficeCLI 整合試用';size='18';'size.cs'='18';bold='true';align='center'})})
$ops.Add(@{command='add';parent='/body';type='toc';props=@{levels='1-1';hyperlinks='true'}})
$ops.Add(@{command='add';parent='/body';type='paragraph';props=(WordProps @{text='一般 Office 操作';style='Heading1';size='18';'size.cs'='18';spaceBefore='12pt'})})
$ops.Add(@{command='add';parent='/body';type='paragraph';props=(WordProps @{text='這段文字將由 OfficeCLI 取代。English text uses Times New Roman.';spaceAfter='8pt'})})
$ops.Add(@{command='add';parent='/body';type='paragraph';props=(WordProps @{text='來源文件保留，編輯結果另存副本。中文使用標楷體，內文 12 pt。';spaceAfter='8pt'})})
$ops.Add(@{command='add';parent='/body';type='paragraph';props=(WordProps @{text='原生 Word 驗證';style='Heading1';size='18';'size.cs'='18';spaceBefore='12pt'})})
$ops.Add(@{command='add';parent='/body';type='paragraph';props=(WordProps @{text='Word 更新自動目錄與頁碼；EndNote 引用繼續使用既有原生工具。'})})
Run-Batch $word $ops.ToArray()
$changes = Join-Path $dir 'changes.json'
@(@{old='這段文字將由 OfficeCLI 取代。';new='這段文字已由 OfficeCLI 完成替換。';count=1}) | ConvertTo-Json -AsArray | Set-Content $changes -Encoding utf8
$revised = Join-Path $dir 'word-revised.docx'
& (Join-Path $PSScriptRoot 'edit_office.ps1') -InputPath $word -Changes $changes -OutputPath $revised

$sheet = Join-Path $dir 'excel-demo.xlsx'
Run-Cli @('create',$sheet,'--json')
Run-Batch $sheet @(
    @{command='set';path='/Sheet1/col[A]';props=@{width='22'}},
    @{command='set';path='/Sheet1/col[B]';props=@{width='14'}},
    @{command='set';path='/Sheet1/A1';props=@{value='項目';bold='true'}},
    @{command='set';path='/Sheet1/B1';props=@{value='數值';bold='true'}},
    @{command='set';path='/Sheet1/A2';props=@{value='Sample A'}},
    @{command='set';path='/Sheet1/B2';props=@{value='10'}},
    @{command='set';path='/Sheet1/A3';props=@{value='Sample B'}},
    @{command='set';path='/Sheet1/B3';props=@{value='20'}},
    @{command='set';path='/Sheet1/A4';props=@{value='總和';bold='true'}},
    @{command='set';path='/Sheet1/B4';props=@{formula='SUM(B2:B3)';bold='true'}}
)
$deck = Join-Path $dir 'powerpoint-demo.pptx'
Run-Cli @('create',$deck,'--json')
Run-Batch $deck @(
    @{command='add';parent='/';type='slide';props=@{background='102A43'}},
    @{command='add';parent='/slide[1]';type='shape';props=@{text='OfficeCLI';x='1.5cm';y='1.8cm';width='27cm';height='2cm';font='Times New Roman';size='40';color='FFFFFF'}},
    @{command='add';parent='/slide[1]';type='shape';props=@{text='Word · Excel · PowerPoint';x='1.5cm';y='4.5cm';width='27cm';height='1.5cm';font='Times New Roman';size='26';color='9ED8F5'}},
    @{command='add';parent='/slide[1]';type='shape';props=@{text='讀取、批次編輯與預覽';x='1.5cm';y='7.2cm';width='27cm';height='1.5cm';font='標楷體';size='26';color='FFFFFF'}}
)
foreach ($file in $revised,$sheet,$deck) {
    Run-Cli @('validate',$file,'--json')
    Run-Cli @('view',$file,'html','-o',($file + '.html'))
}
Write-Host "Demo artifacts: $dir"
Write-Host "Word native rendering: .venv/Scripts/python.exe -X utf8 scripts/render_docx.py `"$revised`" --out-dir `"$dir/word-render`""
