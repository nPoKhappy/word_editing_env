# word_editing_env_officeCLI

OfficeCLI is the preferred engine for ordinary Word, Excel and PowerPoint
creation, inspection, editing and previews in this checkout. Native Word and
Python remain available for live editing, selective directory updates, EndNote
and final visual QA. The original `D:\word_editing_env` is unchanged.

## Try OfficeCLI

```powershell
pwsh -File scripts/install_officecli.ps1
. ./scripts/env.ps1
officecli --version
officecli help docx paragraph --json
officecli view input.docx outline
officecli view input.docx html -o outputs/preview.html
```

The official executable is pinned to 1.0.159, verified against its SHA256 and
installed only under `.tools/officecli`. The environment script changes the
current PowerShell process; it does not modify system PATH. Without dot-sourcing,
use `pwsh -File scripts/officecli.ps1 <command> ...`.

```powershell
# Create Word / Excel / PowerPoint examples and native-check the revised Word:
pwsh -File scripts/new_office_demo.ps1
# Integration tests (Word must be available for the default full run):
pwsh -File scripts/test_officecli.ps1
# CLI/MCP-only tests without Word:
pwsh -File scripts/test_officecli.ps1 -SkipNativeWord
```

See [the integration guide](docs/officecli-integration.md) for the capability
map, batch examples, native QA, version-specific findings and MCP configuration.
The project `.codex/config.toml` includes the official OfficeCLI MCP server;
reload the connection before expecting it in an existing chat. Its stdio
handshake and actual tool execution are covered by the integration test.

## Edit a closed Word document with OfficeCLI

Use a UTF-8 JSON file, for example
`[{"old":"原本的文字","new":"修改後的文字","count":1}]`:

```powershell
pwsh -File scripts/edit_office.ps1 -InputPath input.docx -Changes changes.json -OutputPath revised.docx
```

This path preserves the source and refuses existing destinations. It checks
counts, revisions, locks and protected target paragraphs, stages an atomic
OfficeCLI batch, verifies text and package preservation, then selectively updates
native directories and audits their actual font sizes. Targets containing fields
or objects, table-cell edits and live Word buffers use the native workflow below.
Final visual QA is still required. HTML preview pagination can differ from Word.

## Native Word QA setup

Use Python 3.12 and installed desktop Word for the retained native adapters.
Ordinary OfficeCLI commands do not need Python.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install pywin32==312 pypdfium2==5.13.0 pillow==12.3.0 python-docx==1.2.0 pyyaml==6.0.3
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

## Native fallback for live Word and protected scopes

For live Word buffers or scopes unsupported by the OfficeCLI adapter, use a
UTF-8 JSON file containing exact main-body text replacements:

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

Project EndNote and OfficeCLI MCP entries now use this checkout's local runtimes,
scripts, packages and search snapshot. Reload existing MCP connections after
changing the project configuration.
Adjust absolute paths when cloning to another location.
EndNote libraries and manuscript files remain local.

## Full manuscript chat replay

The [chat replay report](docs/officecli-chat-replay-results.md) records nine
OfficeCLI-authored manuscript stages, 27,135 successful operations, raw CSV
checks, native EndNote preservation and the final Word PDF review. Generated
manuscripts and evidence stay in ignored `outputs/chat-replay-20261028/`.
It also records the remaining native Word/EndNote and scientific plotting
boundaries; this is not a claim that every Python dependency can be removed.
