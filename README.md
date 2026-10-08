# word_editing_env

Windows workspace for editing Word documents with Python and visually checking them with Microsoft Word.

## Setup

Use Python 3.12 and an installed desktop copy of Microsoft Word for Word automation and rendering.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.lock.txt
New-Item -ItemType Directory -Force .tmp | Out-Null
$env:TEMP = Join-Path $PWD '.tmp'
$env:TMP = $env:TEMP
```

Use `.venv\Scripts\python.exe` for Python commands. Set `TEMP` and `TMP` before importing `win32com.client` or running Word automation.

## Render a document

```powershell
.\.venv\Scripts\python.exe scripts\render_docx.py input.docx --out-dir .tmp\rendered
```

The renderer exports a PDF through Microsoft Word and writes a PNG for each page. Inspect the page images after editing.

If the exact document is already open, the renderer exports its current Word
content (including unsaved changes) without closing, hiding, or saving it.
Otherwise it uses a private background instance. `--timeout 120` limits the
whole operation, including blocked COM calls. Increase it for large documents.

## Edit with Word open or closed

Use a UTF-8 JSON file containing exact main-body text replacements:

```json
[{"old": "原本的文字", "new": "修改後的文字", "count": 1}]
```

Preserve the source and create a working copy:

```powershell
.\.venv\Scripts\python.exe scripts\edit_word.py input.docx --changes changes.json --output revised.docx --visible
```

For an explicitly authorized in-place edit:

```powershell
.\.venv\Scripts\python.exe scripts\edit_word.py revised.docx --changes changes.json --in-place
```

An open target is edited in its existing Word window with screen refresh after
each replacement. A closed target runs in the background; `--visible` shows it
and leaves it open after success. Move that window to the second monitor to
watch edits. This does not provide simultaneous typing: avoid changing the
document while the command is applying edits. Existing Word windows are never
closed or terminated by the tool. The default deadline is 120 seconds.

The editor checks all occurrence counts before modifying text, refuses targets
overlapping fields and active tracked changes, updates existing native contents
and figure/equation/table lists, and backs up the disk file before in-place
edits to `.tmp/word-backups`. A source with unsaved Word changes must be saved
before making an output copy. On failure or timeout, an attached document may
contain unsaved partial edits; inspect before retrying. EndNote citation changes
continue to use the separate native CWYW tools.

## Workspace guidance

See `AGENTS.md` and `.agents/skills/` for editing workflows. New Chinese text uses 標楷體; English and Latin text uses Times New Roman.

The `thesis_outline/` folder, document-specific thesis scripts, generated files in
`output/` and `outputs/`, virtual environment, caches, and temporary files are
excluded from version control. Local thesis documents are kept on this computer.

## Native captions and font-size checks

See `.agents/skills/word-native-captions/SKILL.md` for native captions,
cross-references, and automatic contents and figure/table/equation lists. After
editing, update the native lists and verify the final document's directory and
new table text sizes with `scripts/verify_word_font_sizes.py`. Body text, table
text, and directory entries default to 12 pt in both ordinary and complex-script
font sizes; preserve heading and mathematical subscript/superscript sizes.

```powershell
.\.venv\Scripts\python.exe scripts\verify_word_font_sizes.py revised.docx --directories --table-text 'Unique table text' --report .tmp\font-size-report.json
```

## EndNote integration

The local EndNote skill and its MIT-licensed upstream tools are included under
`.agents/skills/endnote-word/`. See its `references/manual.zh-TW.md` for setup,
manual group lookup, and native CWYW citation insertion/removal. Microsoft Word
and desktop EndNote with the CWYW add-in are external prerequisites.

`.codex/config.toml` registers the local MCP server for this workspace at
`D:\word_editing_env`. Adjust its absolute paths if cloning to another location.
EndNote libraries and manuscript files remain local.
