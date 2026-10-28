#requires -Version 7.0
<#!
Apply literal main-body DOCX replacements with OfficeCLI. The source is preserved;
native Word updates and font checks must pass before the output is published.
Changes: [{"old":"before","new":"after","count":1}].
Optional officecli_path scopes a change to /body/p[N].
!#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$InputPath,
      [Parameter(Mandatory)][string]$Changes,
      [Parameter(Mandatory)][string]$OutputPath,
      [string[]]$TableText = @())
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'OfficeCli.psm1') -Force
$root = Split-Path $PSScriptRoot -Parent
$source = (Resolve-Path -LiteralPath $InputPath).Path
$output = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetExtension($source) -ne '.docx' -or [IO.Path]::GetExtension($output) -ne '.docx') { throw 'This editor supports .docx only.' }
if ($source -eq $output -or (Test-Path -LiteralPath $output)) { throw 'Choose a new output path; source and existing outputs are never overwritten.' }
Assert-OfficeFileClosed $source
$xml = Read-OfficeXml $source
$ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
if ($xml.SelectNodes('//w:ins|//w:del|//w:moveFrom|//w:moveTo|//*[contains(local-name(),"PrChange")]', $ns).Count) { throw 'Active tracked changes require the native Word workflow.' }
$settings = Read-OfficeXml $source 'word/settings.xml'
if ($settings.SelectNodes('//*[local-name()="documentProtection" or local-name()="trackRevisions"]').Count) { throw 'Protected or revision-tracking documents require the native Word workflow.' }
$paragraphs = @($xml.SelectNodes('/w:document/w:body/w:p',$ns))
$texts = @($paragraphs | ForEach-Object { ($_.SelectNodes('.//w:t',$ns) | ForEach-Object InnerText) -join '' })
$changesData = @(Get-Content -LiteralPath $Changes -Raw | ConvertFrom-Json)
if (-not $changesData.Count) { throw 'Changes must contain at least one replacement.' }
$commands = [Collections.Generic.List[object]]::new()
$editedIndices = [Collections.Generic.HashSet[int]]::new()
foreach ($change in $changesData) {
    $old = [string]$change.old
    $new = [string]$change.new
    if (-not $old -or ($old + $new) -match '[\x00-\x1f]' -or ($old + $new).Contains('\n') -or ($old + $new).Contains('\v')) { throw 'Use nonempty literal text without control characters or OfficeCLI line-break escapes.' }
    if ($change.PSObject.Properties.Name -contains 'paragraph_index' -or $change.PSObject.Properties.Name -contains 'segment_index') { throw 'Word indices differ from OfficeCLI paths. Use officecli_path or the native editor.' }
    $expected = 1
    if ($change.PSObject.Properties.Name -contains 'count') {
        if ($change.count -isnot [long] -and $change.count -isnot [int]) { throw 'count must be a positive integer.' }
        $expected = [int]$change.count
    }
    if ($expected -lt 1) { throw 'count must be a positive integer.' }
    $total = 0
    for ($i=0; $i -lt $paragraphs.Count; $i++) {
        $path = '/body/p[' + ($i+1) + ']'
        if ($change.officecli_path -and $path -ne $change.officecli_path) { continue }
        $hits = [regex]::Matches($texts[$i], [regex]::Escape($old)).Count
        if (-not $hits) { continue }
        if ($paragraphs[$i].SelectNodes('.//w:fldChar|.//w:fldSimple|.//w:instrText|.//w:sdt|.//w:drawing|.//w:object|.//w:txbxContent', $ns).Count) {
            throw "Target $path contains a field, content control, or embedded object; use the native Word workflow."
        }
        $commands.Add(@{command='set';path=$path;props=@{find=$old;replace=$new;regex='false'}})
        $null = $editedIndices.Add($i)
        $texts[$i] = $texts[$i].Replace($old,$new)
        $total += $hits
    }
    if ($total -ne $expected) { throw "Expected $expected editable occurrences of '$old', found $total. No document was modified." }
}
$stageDir = Join-Path $root ('.tmp/office-edit-' + [guid]::NewGuid())
New-Item -ItemType Directory $stageDir | Out-Null
$stage = Join-Path $stageDir 'staging.docx'
$plan = Join-Path $stageDir 'batch.json'
$commands.ToArray() | ConvertTo-Json -AsArray -Depth 8 | Set-Content $plan -Encoding utf8
$sourceHash = (Get-FileHash -LiteralPath $source).Hash
Copy-Item -LiteralPath $source -Destination $stage
$null = Invoke-OfficeCli -Arguments @('batch',$stage,'--input',$plan,'--json')
$null = Invoke-OfficeCli -Arguments @('close',$stage,'--json')
$actual = Read-OfficeXml $stage
$actualNs = [Xml.XmlNamespaceManager]::new($actual.NameTable)
$actualNs.AddNamespace('w',$ns.LookupNamespace('w'))
$actualTexts = @($actual.SelectNodes('/w:document/w:body/w:p',$actualNs) | ForEach-Object { ($_.SelectNodes('.//w:t',$actualNs) | ForEach-Object InnerText) -join '' })
if (($texts | ConvertTo-Json -Compress) -cne ($actualTexts | ConvertTo-Json -Compress)) { throw 'Text readback differs from the planned replacements.' }
$originalRemainder = $xml.CloneNode($true)
$actualRemainder = $actual.CloneNode($true)
$originalPs = @($originalRemainder.SelectNodes('/w:document/w:body/w:p',$ns))
$actualPs = @($actualRemainder.SelectNodes('/w:document/w:body/w:p',$actualNs))
foreach ($i in $editedIndices) {
    $null = $originalPs[$i].ParentNode.RemoveChild($originalPs[$i])
    $null = $actualPs[$i].ParentNode.RemoveChild($actualPs[$i])
}
if ((Get-OfficeXmlHash $originalRemainder) -ne (Get-OfficeXmlHash $actualRemainder)) { throw 'Unexpected change outside the targeted paragraphs.' }
# Binary parts must remain byte-identical. XML is canonicalized so attribute
# ordering and indentation do not trigger false alarms. Only the CLI's own
# version/time properties and their package registration may differ.
$beforeZip = [IO.Compression.ZipFile]::OpenRead($source)
$afterZip = [IO.Compression.ZipFile]::OpenRead($stage)
try {
    foreach ($entry in $beforeZip.Entries | Where-Object FullName -ne 'word/document.xml') {
        $other = $afterZip.GetEntry($entry.FullName)
        if (-not $other) { throw "Lost package part: $($entry.FullName)" }
        $a=$entry.Open(); $b=$other.Open()
        try {
            $ha=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($a))
            $hb=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($b))
            if ($ha -ne $hb) {
                if ($entry.FullName -eq 'docProps/custom.xml') {
                    # Upstream stamps its own version/time on every mutation.
                    # All user properties must remain exactly as before.
                    $oldProps = Read-OfficeXml $source 'docProps/custom.xml'
                    $newProps = Read-OfficeXml $stage 'docProps/custom.xml'
                    foreach ($props in $oldProps,$newProps) {
                        foreach ($node in @($props.SelectNodes('/*/*[@name="OfficeCLI.Version" or @name="OfficeCLI.LastModified"]'))) { $null=$node.ParentNode.RemoveChild($node) }
                    }
                    if ((Get-OfficeXmlHash $oldProps) -ne (Get-OfficeXmlHash $newProps)) { throw 'Unexpected custom document property change.' }
                } elseif ($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')) {
                    $oldPart = Read-OfficeXml $source $entry.FullName
                    $newPart = Read-OfficeXml $stage $entry.FullName
                    foreach ($partXml in $oldPart,$newPart) {
                        if ($entry.FullName -eq '[Content_Types].xml') {
                            foreach ($node in @($partXml.SelectNodes('/*/*[local-name()="Override" and @PartName="/docProps/custom.xml"]'))) { $null=$node.ParentNode.RemoveChild($node) }
                        } elseif ($entry.FullName -eq '_rels/.rels') {
                            foreach ($node in @($partXml.SelectNodes('/*/*[local-name()="Relationship" and @Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" and (@Target="docProps/custom.xml" or @Target="/docProps/custom.xml")]'))) { $null=$node.ParentNode.RemoveChild($node) }
                        }
                    }
                    if ((Get-OfficeXmlHash $oldPart) -ne (Get-OfficeXmlHash $newPart)) { throw "Unexpected XML change to package part: $($entry.FullName)" }
                } else { throw "Unexpected binary change to package part: $($entry.FullName)" }
            }
        } finally { $a.Dispose(); $b.Dispose() }
    }
    foreach ($entry in $afterZip.Entries) {
        if (-not $beforeZip.GetEntry($entry.FullName)) {
            if ($entry.FullName -ne 'docProps/custom.xml') { throw "Unexpected added package part: $($entry.FullName)" }
            $newProps = Read-OfficeXml $stage 'docProps/custom.xml'
            if ($newProps.SelectNodes('/*/*[not(@name="OfficeCLI.Version" or @name="OfficeCLI.LastModified")]').Count) { throw 'Unexpected added custom properties.' }
        }
    }
} finally { $beforeZip.Dispose(); $afterZip.Dispose() }
$python = Join-Path $root '.venv/Scripts/python.exe'
if (-not (Test-Path $python)) { throw 'Native QA requires the local .venv; see README.md.' }
$env:TEMP = $env:TMP = Join-Path $root '.tmp'
& $python -X utf8 (Join-Path $PSScriptRoot 'finalize_word.py') $stage
if ($LASTEXITCODE -ne 0) { throw "Native directory update failed; staging retained at $stage" }
$fontArgs = @('-X','utf8',(Join-Path $PSScriptRoot 'verify_word_font_sizes.py'),$stage,'--directories','--allow-no-directories','--report',(Join-Path $stageDir 'font-report.json'))
foreach ($needle in $TableText) { $fontArgs += @('--table-text',$needle) }
& $python @fontArgs
if ($LASTEXITCODE -ne 0) { throw "Font verification failed; staging retained at $stage" }
if ((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash) { throw 'Source changed during editing; output was not published.' }
[IO.File]::Copy($stage,$output,$false)
Write-Host "OfficeCLI edited copy: $output"
Write-Host "QA reports: $stageDir"
Write-Host 'Render and inspect the final output before delivery: scripts/render_docx.py'
