# Validation record

2026-10-23, server 1.5.0 (current extension validation):

- Local suite: 58 tests passed. Vendored upstream suite: 74 tests, 72 passed / two optional integration tests skipped. New coverage includes exact-path native readiness, title-search constraints, ambiguity/DOI mismatch, unknown import outcomes, duplicate journal terms with equal abbreviations, stored metadata versus style omissions, batch failure publication, backup restore and complete snapshot freshness. An actual same-size SQLite edit exposed timestamp-only freshness unreliability; content hashes now detect it.
- Fresh stdio MCP initialization/catalog/calls passed: local 44 tools, search 12 tools, citecheck six raw tools. Configuration disables citecheck's direct rewrite tool, leaving 61 configured tools total. Configuration TOML parsed, preserved unrelated servers and `codex mcp list` outside the restricted sandbox confirmed all three EndNote servers enabled. Existing app catalog reload remains required.
- Live native import used only an isolated copy of the shipped EndNote sample library. Fixed public DOI `10.1038/nature14539` was imported via native Direct Export as record 2470 (190 → 191 live references), with exact target path confirmed. Repeating resolution by title/author/year returned `already_present` record 2470 in 2.086 seconds without another native import. Registry title matching uses the title query field and supplied author/year constraints.
- Native import readiness initially reproduced two 90-second false waits: EndNote's window omits `.enl`, and its MRU registry entry can remain on another library. The compatibility patch accepts both observed title forms and requires native process open-file evidence for the exact requested library. The subsequent public sample import completed in about 3.4 seconds. `ENDNOTE_READY_TIMEOUT_MS=20000` is configured; this is not a general latency guarantee.
- Native batch fixtures: two inserted references, one owned Word session, one CWYW format, 3.289 seconds. A second move/remove batch retained one citation, one owned session and one format, 3.104 seconds. Source body and native identities were checked. These scratch DOCX fixtures were not rendered and are not manuscript deliverables.
- Existing research library read/index only: 57 records, 17 PDFs, 233 indexed PDF pages, 57 local CPU embeddings. All 17 nested-path PDFs extracted; real FTS and semantic results passed through the search MCP. Local index checks now include database/WAL content signatures. Refresh is required after attachment changes; PDF-only changes are not automatically fingerprinted.
- Public sample citecheck DOCX online analysis passed with a Semantic Scholar 429/degraded-verification warning preserved. Provider rate limits are not interpreted as nonexistent papers. Citation matching logic was reused from the installed package; only Windows ZIP extraction was patched.
- Read-only preflight of the existing English JTICE v2 found 24 native citations / 20 distinct records, five tables, four media objects, 22 editable equations and 31 unique journal terms. Preflight does not establish source support, inherited Word sizes, figure resolution or rendered quality.

Current reports: `.tools/endnote/validation/mcp-smoke.json`, `mcp-smoke-public-online.json`, `import-native/title-dedup-result.json`, `jtice-preflight.json`; native import evidence under `.tools/endnote/imports/20261023T095942Z-9f7aa072/`. The user's manuscripts and research records were not edited by these tests. See the workspace report `outputs/endnote_jtice_workflow_audit_20261023.md` for investigation evidence and limits.

2026-10-23, server 1.2.0: added `insert_endnote_citation_at_text` and
`move_endnote_citation_to_text`. Both resolve a unique literal anchor within
a unique main-body paragraph; they preserve native CWYW fields and refuse
merged-field movement. The legacy paragraph-end tool retains its original
semantics. An already correctly placed citation is a validated no-op.

- Local suite: 21 tests passed, including the fresh MCP catalog and anchor
  parameters. A live manuscript run completed 12 native field moves and 12
  native insertions, yielding 24 citation fields and 20 distinct records.
- Content preservation, citation identities and native bibliography counts
  were checked after each operation before publication.
- Bibliographic record fields and journal-term abbreviations still require
  verified native EndNote editing; these MCP tools do not write the library
  database. Web research is available through scholarly/web tools rather than
  a bundled automatic metadata repair operation.
- Existing app sessions need MCP reload to expose the new schema. The one-shot
  CLI is immediately usable. These checks do not replace final rendered QA.

Date: 2026-10-07. Local environment: Windows, Word 16.0, EndNote 2025 22.3.1 build 19926, Windows PowerShell 5.1, workspace Python 3.12.

- Upstream commit: `69a5bb2be8d12e4129654df5fb30da3886139fb0`.
- Upstream offline suite: 74 tests, 72 passed and 2 optional integration tests skipped. The checkout's RIS fixture had LF endings although its tests require CRLF; normalized that fixture without changing production logic, then reran successfully.
- Local suite: 11 tests passed, including manual membership, trash exclusion, duplicate-group ambiguity, smart-group refusal, malformed member encoding, MCP stdio handshake/catalog/tool call, split OOXML instructions and owned-PID cleanup rules.
- Live upstream probe: Word COM, EndNote 2025 discovery, CWYW template, add-in and actual `ENFormatBibliography` call all succeeded.
- Real library read: `D:\My_Research_References.enl`, manual group `Claus` ID 9, 8 members. Requested paper found as record 55, DOI empty. No library write was used.
- Real thesis source copy: inserted record 55 after the requested paragraph; stored output had one `ADDIN EN.CITE` and one nonempty `ADDIN EN.REFLIST`.
- Last-citation removal: removed record 55, including the insertion separator when explicitly requested; output had zero citation and bibliography fields. Nonempty paragraph text matched the source; table XML matched canonically.
- Multiple-citation document: inserted record 54 in another paragraph, retained record 55, then removed only record 55. Output retained record 54 and its rebuilt bibliography. Other body text and citation identities were verified by the tool.
- Duplicate insertion: refused, and no final output was published.

Limits: smart/combined groups and removal from a merged multi-paper field are intentionally unsupported. The group codec was observed in this EndNote 2025 library; other encodings fail explicitly. This validation establishes native fields and content preservation, not PDF layout or reliability for every EndNote/Office version. No page-render QA was performed for the scratch integration documents. No live timeout was forced; the owned-PID cleanup rule has unit coverage.

Registration repair, 2026-10-07: added endnote-word-local to the user-level Codex config, preserving unrelated entries and keeping a local pre-change backup. CLI lists it enabled both inside and outside D:/word_editing_env. A fresh stdio process from the user config passed initialization, listed 18 tools, and returned record 55 from Claus. App tool availability still requires MCP reload/application restart and was not claimed as verified in the existing chat.

In-place publication, 2026-10-07 (server 1.1.0): both targeted citation tools now accept explicit `overwrite=true`, with `output_path` omitted or equal to `input_path`. Validation runs on a disposable copy; successful publication saves an exact backup beside the input in `.endnote-backups` and replaces the original filename. Other existing destinations are not authorized by this flag.

- Local suite: 20 tests passed. New coverage includes exact backups, explicit same-path and omitted-output modes, legacy separate-output behavior, invalid flags, another destination, bridge failure, failed body validation, concurrent edits, and replacement failure on a locked destination. The fresh stdio catalog confirms the new optional parameters.
- Native Word/CWYW check: isolated thesis copy `D:/word_editing_env/.tmp/endnote-inplace-check-kll9o6qy/test.docx`. Removed record 51 and then inserted the actual Kazempour 2017 library record, using the same input filename and no output path. Stored native citation/bibliography counts were 0/0 after removal and 1/1 after insertion. Both backups were byte-identical to their respective pre-operation inputs; other body text was preserved.
- A duplicate insertion was refused with the working document unchanged and no additional backup. The test directory retained only one working DOCX. The user's original thesis hash was unchanged by the integration test. Detailed results: `D:/word_editing_env/.tmp/endnote-inplace-check-kll9o6qy/report.json`.
- Existing app sessions still have the previously loaded tool schema until MCP reload/application restart. The updated one-shot CLI is usable immediately. These integration checks establish publication and native citation fields; they do not assert a new page-layout review for the test copy.
