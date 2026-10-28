#requires -Version 7.0
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$files=@(
    'AGENTS.md', '.agents/skills/word-editing/SKILL.md',
    '.agents/skills/word-native-captions/SKILL.md', '.agents/skills/endnote-word/SKILL.md',
    '.agents/skills/endnote-word/references/extensions.md',
    '.agents/skills/endnote-word/references/manual.zh-TW.md',
    '.agents/skills/jtice-format/SKILL.md', '.codex/config.toml'
)
foreach($relative in $files){
    $file=Join-Path $root $relative
    $content=Get-Content -LiteralPath $file -Raw
    $revised=[regex]::Replace($content,'D:\\word_editing_env(?!_officeCLI)','D:\word_editing_env_officeCLI')
    $revised=$revised.Replace('D:/word_editing_env/','D:/word_editing_env_officeCLI/')
    if($revised -cne $content){Set-Content -LiteralPath $file -Value $revised -Encoding utf8 -NoNewline;Write-Output "Relocated $relative"}
}
