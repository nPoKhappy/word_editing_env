---
name: officecli
description: Use the workspace OfficeCLI binary to create, inspect, edit, validate and preview local DOCX, XLSX and PPTX files. Preferred over Python document authoring in this workspace. Preserve native Word/EndNote workflows for live documents, citations, captions and final pagination.
---

# OfficeCLI in this workspace

Use OfficeCLI for ordinary Office authoring and inspection, as requested by the
user. Entrypoint: `scripts/officecli.ps1` (PowerShell 7), or the `officecli` MCP
tool configured in `.codex/config.toml`. Read [the integration guide](../../../docs/officecli-integration.md)
for setup, tested boundaries and native QA. The binary is local and version-pinned
by `officecli.lock.json`; `scripts/install_officecli.ps1` verifies its SHA256.

## Discover, edit, read back

Use `help docx paragraph --json` (or xlsx/pptx and the actual element) before
guessing property names. Inspect with `view file outline`, `get file / --depth 2
--json`, and `query`. Use returned stable paths where available. Prefer semantic
`set`/`add`/`remove` to `raw-set`. Use JSON files and `batch --input` for Unicode,
quotes, currency symbols and multiple operations. A batch item is an object:

```json
{"command":"add","parent":"/body","type":"paragraph","props":{"text":"中文 text","font.latin":"Times New Roman","font.ea":"標楷體","size":"12","size.cs":"12"}}
```

MCP exposes one tool named `officecli` with one string argument `command`, e.g.
`{"command":"help docx paragraph --json"}`. Batch JSON must be passed through
`--input` or `--commands`; MCP stdin is reserved for the protocol.

Check exit status **and** JSON `success`/`warnings`. Exit 2 is a caveat, not a
clean pass. A failed atomic batch must not be described as successful. Flush with
`save` or `close` before Word, QA, or any external reader, even though the local
defaults disable auto residents and request immediate flushes.

## Existing Word files

Use `scripts/edit_office.ps1 -InputPath source.docx -Changes changes.json
-OutputPath revised.docx` for closed documents and ordinary top-level body
literal replacements. It checks exact counts, stages changes, reads back text,
checks other package parts (semantic XML and binary hashes), selectively updates native directories, and runs
the native font verifier before publishing a new output. It refuses existing
outputs, locks, revisions and target paragraphs containing fields or objects.
Use `officecli_path` to scope a change; Word's paragraph indices are different.

For unsupported scopes, inspect first and use a copied file with a focused
OfficeCLI batch, or the existing native workflow where fields/Word ranges are
needed. Direct CLI/MCP edits do not inherit the guarded editor's protections.
Preserve the original unless the user authorized in-place edits.

If the exact file is open in Word or has unsaved changes, use `edit_word.py` and
`WordSession` with the existing watchdog. OfficeCLI edits the saved OOXML package;
its watch preview is a browser view and does not synchronize a live Word buffer.

EndNote insertion/removal uses the native EndNote tools. Do not construct ADDIN
fields or typed citation substitutes. Use `word-native-captions` for new numbered
figures/tables/display equations, including in-text cross-references.

## Typography and delivery

New Word body/table text uses `font.latin=Times New Roman`, `font.ea=標楷體`,
`size=12`, **and** `size.cs=12`. Preserve heading and index sizes. In 1.0.159,
style `size` writes only `w:sz`; use a scoped `raw-set` for `w:szCs` on a new
style when necessary, then verify XML and actual generated entries. Never apply
12 pt to all runs indiscriminately.

Use `finalize_word.py` for selective native directory updates, followed by
`verify_word_font_sizes.py --directories` (add `--allow-no-directories` only when
absence is appropriate). Audit added/reformatted tables with `--table-text`.
`officecli refresh` calls Word's whole Fields.Update in this version, so use the
selective finalizer for manuscripts with citations. Browser/HTML pagination is
not proof of final Word page numbers.

Preview quickly with `view file html -o preview.html` or `watch file --port N`.
The wrapper is time-bounded: for a persistent watch use the direct binary after
dot-sourcing `scripts/env.ps1`. Stop only your own preview with `unwatch file`.
Final DOCX QA uses `render_docx.py` and visual inspection of every page.

## Verified manuscript replay

See [the replay report](../../../docs/officecli-chat-replay-results.md) for the
nine-stage real-manuscript test and evidence locations.

- Use the picture path returned by query/get (usually a drawing run `/r[N]`).
  `/body/p[N]/picture[1]` is not a supported substitute for that returned path.
- Run `text` does not decode paragraph-level `\v` escapes. For a soft line break,
  add `type=break` under the paragraph with `props.type=line`, then the next run.
- Native caption placeholders must be deleted by their exact paragraph after
  InsertCaption. Do not use visible text length to compute native field ranges.
- Audit effective formatting, including inherited styles. A missing direct
  `w:b` does not establish that a number is not bold.
- Preserve native chemical indices in generated captions as well as body text.
  `finalize_word.py --caption-spec` supports the replay's native caption ranges.
- The native render helper now isolates each PDF page to avoid the observed
  late-page PNG glyph-loss issue. A correct PDF text layer alone does not prove
  that a cached PNG displays all text.