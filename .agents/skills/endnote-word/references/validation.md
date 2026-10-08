# Validation record

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
