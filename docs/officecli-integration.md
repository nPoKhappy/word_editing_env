# OfficeCLI integration

This workspace prefers OfficeCLI for ordinary Office work. The original
`D:\word_editing_env` is untouched. OfficeCLI 1.0.159 is installed in
`.tools/officecli`; the executable and generated previews are ignored by Git.

## Routing

| Task | Preferred path | Verification |
| --- | --- | --- |
| Read DOCX/XLSX/PPTX structure and text | OfficeCLI view/get/query | Read structured JSON |
| Create paragraphs, runs, cells, slides, shapes | OfficeCLI add/set/batch | Readback + validate + preview |
| Closed DOCX body literal replacements | scripts/edit_office.ps1 | Counts, staged copy, native directories and font audit |
| Existing Word document open on second monitor | scripts/edit_word.py | Exact live attachment and watchdog |
| Native numbered captions/cross-references | Existing Word native caption workflow | SEQ/REF and automatic list membership |
| EndNote/CWYW | Existing native EndNote tools | Genuine field/library identity preservation |
| Word directory pagination | scripts/finalize_word.py | Selective update, stability, native font verifier |
| Quick preview | OfficeCLI HTML/watch/screenshot --render html | Preview only; check all relevant pages |
| Final DOCX layout | scripts/render_docx.py | Microsoft Word PDF and page images |

OfficeCLI handles OOXML editing without Microsoft Office. Native Word remains
necessary for this workspace's final Word pagination, live editing and CWYW.
Python is retained for those adapters and QA; it is no longer the default
document authoring engine. Old scripts remain available for unsupported cases.

## Commands

Use PowerShell 7. No global installer, PATH changes or changes to other agents
are required.

```powershell
pwsh -File scripts/install_officecli.ps1
. ./scripts/env.ps1
officecli --version
officecli help docx paragraph --json
officecli view report.docx outline
officecli get report.docx / --depth 2 --json
officecli batch report.docx --input operations.json --json
officecli close report.docx
officecli validate report.docx --json
officecli view report.docx html -o outputs/report.html
officecli watch report.docx --port 26315
# After reviewing:
officecli unwatch report.docx
```

Without dot-sourcing the environment, use `pwsh -File scripts/officecli.ps1 ...`.
It configures writable temp directories, disables automatic binary upgrades and
auto residents, and bounds each CLI process to 120 seconds. It preserves the
CLI exit code. The module helper treats exit 2 as a failed clean-pass gate.

```powershell
pwsh -File scripts/new_office_demo.ps1
pwsh -File scripts/test_officecli.ps1
```

For a new existing-DOCX edit, put literal replacements in a UTF-8 JSON file:

```json
[{"old":"原本的文字","new":"修改後的文字","count":1}]
```

```powershell
pwsh -File scripts/edit_office.ps1 -InputPath original.docx -Changes changes.json -OutputPath revised.docx
# For any tables reformatted by a separate operation, pass -TableText 'unique text'.
.venv/Scripts/python.exe -X utf8 scripts/render_docx.py revised.docx --out-dir outputs/revised-render
```

This guarded editor supports top-level body paragraphs. It intentionally fails
for Word-specific `paragraph_index`/`segment_index`, fields in the target,
revisions and open files; use the native editor for those cases. It preserves
the source and refuses to overwrite an existing destination. It retains staging
and QA evidence in `.tmp/office-edit-*`. Successful native size checks alone do
not establish visual correctness; inspect the final Word-rendered pages.

## MCP

The project `.codex/config.toml` registers the official binary with `args=['mcp']`
and writable TEMP/TMP. Restart/reload the project's MCP connection before expecting
a new `officecli` tool to appear in an existing chat. The server exposes one
tool named `officecli`, whose `command` is a CLI string. CLI use is available
immediately. EndNote MCP paths now use this checkout's Python, scripts,
dependencies and search snapshot. The original environment is a read-only
migration source.

The integration test communicates directly with that binary over MCP stdio.
This proves the server works; it does not prove an already-running chat has
reloaded its tool catalog. The OfficeCLI tool is now available in this chat.
Project MCP path changes require reloading existing connections. The chat replay
tested fresh instances of all four project servers with actual tool calls.

## Native QA dependencies

```powershell
python -m venv .venv
.venv/Scripts/python.exe -m pip install pywin32==312 pypdfium2==5.13.0 pillow==12.3.0 python-docx==1.2.0 pyyaml==6.0.3
```

Before native Word commands set TEMP/TMP as required by AGENTS.md. The native
session helper uses this checkout's `.tmp` internally. The full legacy Python
dependency set remains recorded in `requirements.lock.txt`; install it only if
a legacy workflow actually needs it. OfficeCLI's normal read/create/edit/HTML
commands do not require Python.

## Version-specific findings

- Official x64 binary SHA256 is pinned in `officecli.lock.json`; installation
  aborts on mismatch. Updating requires a deliberate lock-file revision and tests.
- OfficeCLI needs a writable temp directory even for package reads on this host.
- `size` on a DOCX style does not set `w:szCs` in 1.0.159. New run-level content
  supports `size.cs`; directory styles also need an explicit complex-script size.
- The Windows refresh backend invokes the whole document's `Fields.Update`.
  The selective finalizer updates only SEQ, REF, PAGEREF and TOC/list fields.
- HTML refresh can fall back to HTML pagination; it can disagree with Word.
- A resident session can hold edits in memory. Close/save before external reads.
- Raw CLI/MCP is a full editing interface, not a policy enforcement layer. Follow
  AGENTS.md and the local OfficeCLI skill when choosing safe scopes and QA.
- On save, the OpenXML SDK can reorder attributes and rename namespace prefixes.
  The guarded editor compares XML by expanded names and values, binary parts by
  SHA256, and permits only OfficeCLI's own version/time metadata additions.

## Verified on 2026-10-28

The full integration run passed 23 checks: Unicode and shell argument fidelity,
font slots and complex-script size, atomic rollback, source preservation, native
Word finalization, inline subscript preservation, count mismatch and open-file
lock rejection, field-bearing target protection, three-format schema validation,
Excel computed formula readback, PowerPoint text, three-format HTML previews,
and MCP initialize/discovery/actual tool invocation.

An additional migration check successfully edited a genuine python-docx-created
fixture with OfficeCLI and completed native Word QA. The Word demo has two native
contents entries at 12 pt (including complex-script size), a one-page Word PDF,
and an inspected page PNG. PowerPoint and Excel HTML screenshots were inspected;
the Excel demo column widths were adjusted to avoid truncated labels.

Artifacts: `outputs/officecli-demo/ready/`. The full test JSON is copied there as
`integration-test-report.json`. These generated files remain local and ignored.
This does not validate every OfficeCLI feature or genuine CWYW round-tripping
through OfficeCLI. Real manuscripts with citations continue using the existing
EndNote tools and native Word finalization. No research manuscript was edited.

Sources inspected for this integration:
[upstream repository](https://github.com/iOfficeAI/OfficeCLI),
[pinned release](https://github.com/iOfficeAI/OfficeCLI/releases/tag/v1.0.159),
[Word refresh backend](https://github.com/iOfficeAI/OfficeCLI/blob/v1.0.159/src/officecli/Core/WordPdfBackend.cs),
[Codex MCP configuration](https://learn.chatgpt.com/docs/extend/mcp?surface=cli).
