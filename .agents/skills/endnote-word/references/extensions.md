# EndNote extensions, 1.5.0

Validated on 2026-10-23, Windows / EndNote 2025 / Word 16. Native existing-record updates and native journal-term installation still require EndNote's own interface. No extension writes the EndNote SQLite database.

## Fast path

For “把這篇文章匯入 EndNote”, resolve the actual target library (the known workspace library is `D:\My_Research_References.enl`) and call:

```json
{"library_path":"D:\\My_Research_References.enl","title":"ARTICLE TITLE","author":"OPTIONAL AUTHOR","year":2020}
```

Use `import_verified_endnote_reference`; omit unknown optional fields. Alternatives are `doi`, `url` containing a DOI, or `pdf_path`. `preview_only:true` prepares evidence and RIS without native import. A readable PDF requires a unique DOI on page one and matching registry title. This imports verified bibliographic data; it does not download/attach the PDF or automatically add the record to a group. Generic links without DOI require a title or DOI. Scanned PDFs and ambiguous titles return the missing identity detail instead of importing a guess.

An existing exact record returns `already_present` and its real record number. Conflicting or multiple existing matches return review candidates and are not overwritten. New records use the existing native RIS/Direct Export bridge and are re-read after import; an unconfirmed result must not be retried blindly. Record/term/style diagnosis and actual native Word output complete JTICE verification when a document is available.

One-shot use works before an app reload. Save UTF-8 JSON arguments to a file, then:

```powershell
$env:TEMP = $env:TMP = 'D:\word_editing_env\.tmp'
& 'D:\word_editing_env\.venv\Scripts\python.exe' -X utf8 `
  'D:\word_editing_env\.agents\skills\endnote-word\scripts\endnote_mcp.py' `
  --call import_verified_endnote_reference --arguments 'D:\word_editing_env\.tmp\import-arguments.json'
```

## Installed capabilities

| Server / tools | Purpose and boundary |
| --- | --- |
| `endnote-word-local`: `import_verified_endnote_reference` | Title/DOI/link/PDF identity → deduplication → native import → assigned ID confirmation. |
| `search_online_endnote_metadata`, `lookup_endnote_doi_metadata` | Crossref title/bibliographic search with optional author/year; Crossref then DataCite DOI lookup. Candidates and registry evidence, not automatic existing-record edits. |
| `verify_endnote_reference_online`, `verify_endnote_references_online` | Installed citecheck multi-index review; up to three selected records per bounded batch, real record IDs preserved. Bibliographic fields are sent to public indexes. Rate limiting is reported as degraded verification. |
| `audit_endnote_library`, `find_endnote_duplicate_candidates` | Complete active library or exact manual-group missing fields and duplicate candidates; no automatic merge/delete. |
| `list_endnote_journal_terms`, `prepare_endnote_journal_terms`, `diagnose_endnote_journal_abbreviations` | Read terms, detect duplicate names even when abbreviations agree, merge/export four-column preview, distinguish record/term/style problems. Conflicting nonempty values block export. Installation remains native. |
| `get_endnote_reference_full`, `list_endnote_tags`, `list_endnote_group_tree` | Complete metadata/abstracts, actual native reference types, tags and manual group hierarchy through corrected reader adapters. |
| `export_endnote_records`, `export_endnote_catalog` | Selected JSON/RIS and JSON/CSV exports; never overwrite source library or claim typed bibliography output is CWYW. |
| `audit_endnote_attachments` | Missing/unsafe local attachment paths; no attachment write or automatic full-text retrieval. |
| `backup_endnote_library`, `restore_endnote_library_backup` | Closed-library full `.enl` / `.Data` backup with hashes and empty directories; verified restore into a new directory, never over a live library. |
| `audit_endnote_document_library`, `inspect_endnote_document` | Compare actual native citation identities, bibliography presence and library records without starting Word. |
| `preflight_jtice_manuscript` | Read-only DOCX inventory, optional source-reference omissions and journal diagnostics. Stages: draft, translation, submission. Explicit font findings do not prove inherited visible sizes. |
| `batch_edit_endnote_citations` | 1–50 insert/move/remove operations, one owned Word session, one final CWYW format, atomic publication, automatic backup for authorized overwrite. Requires unique sentence anchors; merged fields refused. Final native directories/fonts/render QA still required. |
| `endnote-library-search` | Existing endnote-mcp: local references, PDF page full-text search, CPU semantic search, PDF sections and BibTeX. Search-only cached XML/SQLite; no live library writes. |
| `get_endnote_search_index_status`, `refresh_endnote_search_index` | Detect library changes with content signatures, refresh complete metadata/PDF snapshot and optional embeddings. Current source PDF cap is 100 pages per PDF. Refresh after native library/attachment changes; PDF content changes alone are not tracked by the library signature. |
| `endnote-reference-check` | Existing citecheck MCP workspace scan, reference analysis, repair/rewrite previews and version. Its direct file rewrite tool is disabled to preserve the native CWYW manuscript. |

The original native CWYW rendering/style/format/import tools remain available. Use genuine fields, not the search server's formatted plain-text citation or BibTeX output, when editing Word.

## Runtime and source reuse

Dependencies are isolated in `D:\word_editing_env\.tools\endnote`, excluded from Git. The main workspace venv was not populated with the large search stack. Versions, 75 Python packages and the installed Windows patch hash are recorded in `.tools/endnote/installed-manifest.json`.

- [endnote-for-word](https://github.com/felimet/endnote-for-word), vendored MIT commit `69a5bb2be8d12e4129654df5fb30da3886139fb0`: native RIS and CWYW. Local readiness compatibility patch is disclosed in `upstream.json`.
- [endnote-mcp](https://github.com/gokmengokhan/endnote-mcp) 1.4.13: metadata/PDF/semantic search, installed with semantic dependencies, Torch 2.9.1+cpu and the cached all-MiniLM-L6-v2 model. Runtime adapter preserves nested attachment paths; normal serving is offline with telemetry disabled.
- [citecheck](https://github.com/jhlee0619/citecheck), npm `@jhlee0619/citecheck@0.1.0` (its server reports 0.1.12): scholarly matching is reused unchanged. Windows DOCX extraction uses workspace Python ZipFile instead of Unix unzip; `repair_installed_extensions.py` reapplies this narrow patch idempotently.
- [endnote-cli](https://github.com/TCMzhoutong/endnote-cli), source commit `7faca0f95c9aa89b683f61707148a2967117d608`: read/export only, snapshot-backed connection. The adapter corrects reference-type mapping and native group decoding; writer components are not used. This source is not substituted with an unverified PyPI package.
- Public registry search follows [Crossref's REST API](https://www.crossref.org/documentation/retrieve-metadata/rest-api/); its metadata may still omit publication fields. Missing fields are reported rather than invented.

Workspace/user MCP configurations register three servers. Reload MCP or restart the app to refresh an existing chat's catalog. Before changing global registration, preserve unrelated TOML settings and validate the resulting file. To reinstall, use the pinned sources/package versions above in the same isolated directories, run the Windows repair script and refresh the index; do not import search-only XML into EndNote.

## Verification and remaining limits

The live integration records are under `.tools/endnote/validation/`. The isolated shipped sample library received public DOI `10.1038/nature14539` as actual record 2470; subsequent title+author+year resolution returned that existing ID without another import. The two-insert native Word fixture used one session/one format in about 3.3 seconds; a move/remove batch also passed. These are measured examples, not latency guarantees.

The user's research library and manuscript were only read during these integration tests. Its local search snapshot contains 57 records, 17 PDFs and 233 pages, with 57 embeddings. Public sample multi-index checking passed with an explicitly reported Semantic Scholar rate-limit warning. Scratch fixture layout was not rendered or claimed validated; manuscript delivery still requires the workspace's native final QA.

No verified automatic existing-record native field editor, journal-term installer, native merge/delete, smart-group evaluator, sync/sharing API or authenticated full-text downloader is installed. EndNote's own Find Reference Updates can help eligible records through its UI. Do not promise these as MCP writeback capabilities or use direct database writes as a workaround.
