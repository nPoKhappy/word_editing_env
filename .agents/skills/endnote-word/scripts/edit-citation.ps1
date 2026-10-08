[CmdletBinding()]
param([Parameter(Mandatory)][string]$ConfigPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$doc = $null
$answer = $null

function Get-Cites {
    return @($doc.Fields | Where-Object { $_.Code.Text -match '^\s*ADDIN EN\.CITE\s' })
}
function Get-CitationXml {
    param($Field)
    $code = [string]$Field.Code.Text
    $start = $code.IndexOf('<EndNote>')
    if ($start -lt 0) { throw 'Citation contains no traveling-library XML.' }
    $xml = New-Object System.Xml.XmlDocument
    $xml.XmlResolver = $null
    $xml.LoadXml($code.Substring($start).Trim())
    return $xml
}
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class EndNoteWordPid {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
'@
    $beforePids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $bootstrap = $word.Documents.Add()
    [uint32]$ownPid = 0
    $null = [EndNoteWordPid]::GetWindowThreadProcessId([IntPtr]$word.ActiveWindow.Hwnd,[ref]$ownPid)
    if ($ownPid -eq 0 -or $beforePids -contains $ownPid) { throw 'Word did not create an isolated owned process.' }
    [Console]::Error.WriteLine("WORD_PID=$ownPid")
    $bootstrap.Close(0)
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrap)
    $doc = $word.Documents.Open($config.scratch_path,$false,$false)
    $doc.Activate()
    if ($doc.Revisions.Count -gt 0) { throw 'Review tracked changes before citation editing.' }
    $paragraphs = @($doc.Paragraphs | Where-Object { $_.Range.Text.Contains([string]$config.paragraph_text) })
    if ($paragraphs.Count -ne 1) { throw "Expected one matching paragraph; found $($paragraphs.Count)." }
    $paragraph = $paragraphs[0].Range
    if ($paragraph.Information(12)) { throw 'Target paragraph is in a table; main-body paragraph editing is supported.' }
    $initial = @(Get-Cites).Count
    if ($config.action -eq 'insert') {
        foreach ($field in @(Get-Cites)) {
            if ($field.Code.Start -ge $paragraph.Start -and $field.Code.Start -lt $paragraph.End) {
                $xml = Get-CitationXml $field
                foreach ($cite in @($xml.EndNote.Cite)) {
                    if ([int]$cite.RecNum -eq [int]$config.rec_number) { throw 'This record is already cited in the target paragraph.' }
                }
            }
        }
        $point = $paragraph.End - 1
        $temp = '{' + $config.record.first_author_surname + ', ' + $config.record.year + ' #' + $config.rec_number + '}'
        $range = $doc.Range($point,$point)
        $range.Text = ' ' + $temp
        $range.Font.Name = 'Times New Roman'
        $doc.Save()
        $null = $word.Run('ENFormatBibliography')
        if (@(Get-Cites).Count -ne $initial + 1) { throw 'CWYW did not add exactly one native citation.' }
        $matching = @()
        foreach ($field in @(Get-Cites)) {
            if ($field.Code.Start -ge $paragraph.Start -and $field.Code.Start -lt $paragraph.End) {
                $xml = Get-CitationXml $field
                foreach ($cite in @($xml.EndNote.Cite)) {
                    if ([int]$cite.RecNum -eq [int]$config.rec_number -and
                        [string]$cite.SelectSingleNode('record/titles/title').InnerText -eq [string]$config.expected_title) {
                        $matching += $field
                    }
                }
            }
        }
        if ($matching.Count -ne 1) { throw 'Inserted field is not the requested record at the requested paragraph.' }
    }
    elseif ($config.action -eq 'remove') {
        $matching = @()
        foreach ($field in @(Get-Cites)) {
            if ($field.Code.Start -lt $paragraph.Start -or $field.Code.Start -ge $paragraph.End) { continue }
            $xml = Get-CitationXml $field
            $cites = @($xml.EndNote.Cite)
            foreach ($cite in $cites) {
                if ([int]$cite.RecNum -eq [int]$config.rec_number -and
                    [string]$cite.SelectSingleNode('record/titles/title').InnerText -eq [string]$config.expected_title) {
                    if ($cites.Count -ne 1) { throw 'Merged citation fields require a dedicated edit; refusing to delete other references.' }
                    $matching += $field
                }
            }
        }
        if ($matching.Count -ne 1) { throw "Expected one standalone citation match; found $($matching.Count)." }
        $position = $matching[0].Code.Start - 1
        $space = $doc.Range($position - 1,$position)
        $hasSpace = ($space.Text -eq ' ')
        $matching[0].Delete()
        if ($config.remove_preceding_space -and $position -gt $paragraph.Start) {
            if ($hasSpace) { $space.Text = '' }
        }
        if (@(Get-Cites).Count -eq 0) {
            foreach ($field in @($doc.Fields | Where-Object { $_.Code.Text -match '^\s*ADDIN EN\.REFLIST' })) {
                $field.Delete()
            }
            foreach ($variable in @($doc.Variables)) {
                if ($variable.Name -eq 'EN.Libraries') { $variable.Delete() }
            }
        }
        else { $null = $word.Run('ENFormatBibliography') }
        if (@(Get-Cites).Count -ne $initial - 1) { throw 'Removal did not delete exactly one citation field.' }
    }
    else { throw 'Unknown edit action.' }
    $doc.Save()
    $answer = [ordered]@{ok=$true;action=$config.action;owned_word_pid=$ownPid;citation_count=@(Get-Cites).Count}
}
catch {
    $answer = [ordered]@{ok=$false;action=$config.action;error=$_.Exception.Message;location=$_.ScriptStackTrace}
}
finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch { [Console]::Error.WriteLine($_.Exception.Message) } }
    if ($null -ne $word) { try { $word.Quit() } catch { [Console]::Error.WriteLine($_.Exception.Message) } }
    if ($null -ne $doc) { $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) }
    if ($null -ne $word) { $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$answer | ConvertTo-Json -Depth 5 -Compress
if (-not $answer.ok) { exit 1 }
