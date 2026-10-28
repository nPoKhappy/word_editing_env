#requires -Version 7.0
Set-StrictMode -Version Latest
$script:Workspace = Split-Path $PSScriptRoot -Parent

function Invoke-OfficeCli {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Arguments, [int]$TimeoutSeconds = 120,
          [switch]$AllowFailure)
    $binary = Join-Path $script:Workspace '.tools/officecli/officecli.exe'
    if (-not (Test-Path $binary)) { throw 'Run pwsh -File scripts/install_officecli.ps1 first.' }
    $info = [Diagnostics.ProcessStartInfo]::new($binary)
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = $info.StandardErrorEncoding = [Text.Encoding]::UTF8
    $temp = Join-Path $script:Workspace '.tmp'
    New-Item -ItemType Directory -Force $temp | Out-Null
    foreach ($key in 'TEMP','TMP') { $info.Environment[$key] = $temp }
    $info.Environment['OFFICECLI_SKIP_UPDATE'] = '1'
    $info.Environment['OFFICECLI_NO_AUTO_RESIDENT'] = '1'
    $info.Environment['OFFICECLI_RESIDENT_FLUSH'] = 'each'
    $info.Environment['OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT'] = '1'
    foreach ($arg in $Arguments) { $info.ArgumentList.Add($arg) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    try {
        $null = $process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            # Terminate only this CLI, never a pre-existing Word/browser process.
            $process.Kill()
            throw "OfficeCLI timed out after ${TimeoutSeconds}s; inspect staging files before retrying."
        }
        $result = [pscustomobject]@{ ExitCode=$process.ExitCode; Stdout=$stdout.GetAwaiter().GetResult(); Stderr=$stderr.GetAwaiter().GetResult() }
        if (-not $AllowFailure -and $result.ExitCode -ne 0) {
            throw "OfficeCLI exited $($result.ExitCode): $($result.Stderr) $($result.Stdout)"
        }
        return $result
    } finally { $process.Dispose() }
}

function Read-OfficeXml {
    param([string]$Path, [string]$Part = 'word/document.xml')
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry($Part)
        if (-not $entry) { throw "Missing package part: $Part" }
        $reader = [IO.StreamReader]::new($entry.Open())
        try {
            $xml = [xml]::new()
            $xml.PreserveWhitespace = $true
            $xml.LoadXml($reader.ReadToEnd())
            return ,$xml
        } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
}

function Get-OfficeXmlHash {
    param([xml]$Xml)
    $normalized = [xml]::new()
    $normalized.PreserveWhitespace = $false
    $normalized.LoadXml($Xml.OuterXml)
    # OpenXML SDK may change namespace prefixes as well as attribute order.
    # Compare expanded names + values, preserving child order and text content.
    $builder = [Text.StringBuilder]::new()
    function Add-Token([string]$Value) { $null=$builder.Append($Value.Length).Append(':').Append($Value) }
    function Visit-Node([Xml.XmlNode]$Node) {
        Add-Token $Node.NodeType.ToString()
        Add-Token $Node.NamespaceURI
        Add-Token $Node.LocalName
        if ($Node.NodeType -eq [Xml.XmlNodeType]::Element) {
            $attrs = @($Node.Attributes | Where-Object NamespaceURI -ne 'http://www.w3.org/2000/xmlns/' | Sort-Object NamespaceURI,LocalName)
            Add-Token ([string]$attrs.Count)
            foreach ($attr in $attrs) { Add-Token $attr.NamespaceURI; Add-Token $attr.LocalName; Add-Token $attr.Value }
        } else { Add-Token $Node.Value }
        Add-Token ([string]$Node.ChildNodes.Count)
        foreach ($child in $Node.ChildNodes) { Visit-Node $child }
    }
    Visit-Node $normalized.DocumentElement
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($builder.ToString())))
}

function Assert-OfficeFileClosed {
    param([string]$Path)
    $name = [IO.Path]::GetFileName($Path)
    $dir = [IO.Path]::GetDirectoryName($Path)
    foreach ($lockName in @(('~$' + $name), ('~$' + $name.Substring([Math]::Min(2,$name.Length))))) {
        if (Test-Path -LiteralPath (Join-Path $dir $lockName)) {
            throw 'Document is open in Office. Use edit_word.py for live Word edits, or save and close this target before OfficeCLI editing.'
        }
    }
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'None')
    $stream.Dispose()
}

Export-ModuleMember -Function Invoke-OfficeCli, Read-OfficeXml, Get-OfficeXmlHash, Assert-OfficeFileClosed
