# Word editing workspace

## OfficeCLI-first integration

The user requested replacing general Python Office manipulation with OfficeCLI
in this checkout (`D:\word_editing_env_officeCLI`). Read
`.agents/skills/officecli/SKILL.md` for Office tasks. Use the pinned binary via
`scripts/officecli.ps1`, or dot-source `scripts/env.ps1` for `officecli` on PATH.
Install/verify with `scripts/install_officecli.ps1`. The native `officecli` MCP
server is registered in `.codex/config.toml`; reload the MCP connection to load it.

Prefer OfficeCLI view/get/query/add/set/batch/validate/HTML preview for ordinary
DOCX/XLSX/PPTX operations. For closed DOCX body replacements, use
`scripts/edit_office.ps1 -InputPath ... -Changes ... -OutputPath ...`; it keeps
the source, stages edits, checks readback, and runs native directory/font QA.
The direct CLI/MCP does not enforce these guards. Flush/close residents before
external readers. Never edit the saved package while the same file is open in
Word; use the live native editor described below for that case.

This routing takes precedence over older skills' default Python/XML authoring
recipes. Preserve the existing formatting, citation, caption, directory and
rendered QA requirements. EndNote/CWYW and live Word remain native. For final
directory updates use `scripts/finalize_word.py`, not OfficeCLI's whole-field
`refresh`, which can touch ADDIN fields. OfficeCLI HTML/watch previews do not
establish final Word pagination. See `docs/officecli-integration.md` and the
verified replay boundaries in `docs/officecli-chat-replay-results.md`.

For DOCX work in this folder, read `.agents/skills/word-editing/SKILL.md`.
For targeted text edits in existing DOCX files, also read
`.agents/skills/docx-text-ops/SKILL.md`. For visual style changes, read
`.agents/skills/docx-design-ops/SKILL.md`.
Use `.venv/Scripts/python.exe` for Python commands. Set `TEMP` and `TMP` to
`D:\word_editing_env_officeCLI\.tmp` before importing `win32com.client` or running Word
automation, because this environment cannot write to the default temp folder.
For Word text added in this workspace, use 標楷體 for Chinese and Times New Roman
for English/Latin text, as described in the local word-editing skill.

## Font size verification

Unless the user or an explicit document template specifies a different size,
new body text, new table cell text, and main/figure/table/equation directory
entries use 12 pt. Preserve the established sizes of chapter headings and
mathematical subscripts/superscripts; do not flatten the whole document to 12 pt.
Never reduce table text to 10 pt just to make it fit. Adjust column widths,
wrapping, spacing, or pagination while keeping the required font size.

Set both ordinary and complex-script sizes (`Font.Size` and `Font.SizeBi` in
Word, or `w:sz` and `w:szCs` in OOXML). A style-level size alone does not prove
that visible text has the same size: direct formatting can override it.

After each DOCX edit, complete the native directory updates required below,
then run `scripts/verify_word_font_sizes.py` against the exact final document.
Use `--directories` and one `--table-text` argument with a unique text fragment
for each added or reformatted table. Specify `--expected-size` only when a
different size was explicitly requested or required by the document template.
The check must cover both directory styles and the actual generated entries;
checking cached entries alone can miss a style that reverts on the next update.
Resolve failures before reporting the edited document as complete. This check
is read-only and does not replace directory updates or rendered visual QA.

Example (PowerShell):

```powershell
& .\.venv\Scripts\python.exe -X utf8 scripts/verify_word_font_sizes.py `
  'thesis_outline\研究計劃書 周莘宸 第三版.docx' --directories `
  --table-text '分組增益平衡' --report '.tmp\font-size-report.json'
```

For EndNote/CWYW citations and group lookup, read `.agents/skills/endnote-word/SKILL.md`. Use its native citation tools and preserve source documents.

For inserting, editing, moving, or deleting Word figures, tables, or equations,
read `.agents/skills/word-native-captions/SKILL.md`. Use native Word InsertCaption,
the existing figure/table/equation paragraph styles, and at least one meaningful
in-text native cross-reference for every generated numbered item. On completion
of every DOCX edit, automatically update and verify the full main contents and
all figure/table/equation lists; new figures and tables must appear in the
corresponding native automatic lists. Captions and directory entries do not
count as in-text mentions.

## Open or closed Word documents

The user prefers watching edits in Word on a second monitor. For main-body
literal text edits in an open Word document, use `scripts/edit_word.py` with `--changes` and an explicit
`--output` or user-authorized `--in-place`. It attaches to the exact open
document, refreshes after each edit, and leaves existing Word windows open.
Closed targets run in a background Word instance; `--visible` opens a review
window and leaves it open after a successful edit. Preserve the source unless
the user authorized an in-place edit. The command has a 120-second watchdog.

For other native Word edits, reuse `scripts/word_session.py:WordSession` and
run COM work in a child process through `run_bounded`; direct COM calls alone
cannot enforce a deadline. Do not use Selection for ordinary text edits or
hide/close/quit an attached Word instance. A timeout may leave unsaved partial
edits in an attached document; inspect before retrying. Use the updated
`scripts/render_docx.py` to render an already-open document, including its
unsaved content, without opening another copy or saving/closing the document.

EndNote/CWYW insertion and removal still use the native EndNote tools and their
transactional publication rules. This general text-editing path must not edit
or replace native citation fields, directory fields, or active tracked changes.
