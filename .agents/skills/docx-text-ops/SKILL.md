---
name: docx-text-ops
description: "Use this skill when the user wants to inspect or edit text in an existing Word (.docx) file without redesigning the document. Triggers include: replacing or rewriting body text, proofreading or translating a document, normalizing product names or terminology, cleaning up placeholder text (Lorem ipsum, XXXX, template prompts), editing headers, footers, footnotes, endnotes, or comments, or auditing document content for duplicates or missing translations. Use this skill when the existing layout, styles, images, tables, and relationships must be preserved. Do NOT use for creating new documents from scratch or redesigning styles — use python-docx scripting for those. Do NOT use for documents with active tracked changes — accept or reject all changes in Word before editing."
license: Original work, no third-party license constraints
---

# DOCX Text Operations Skill

[日本語](SKILL.ja.md)

Text-only editing of existing `.docx` files. Preserves styles, images, tables, numbering, and relationships.

## Quick Reference

| Task | Approach |
|------|----------|
| Extract all document text | `python3 docx-text-ops/scripts/extract.py document.docx` |
| Inspect raw XML runs | Unpack → read `word/document.xml` |
| Edit text | Unpack → edit XML → pack |
| Validate after edit | `python3 docx-text-ops/scripts/extract.py edited.docx` |
| Visual spot-check | Open in Word or LibreOffice |

---

## Non-Goals

- Creating new documents from scratch → use python-docx scripting
- Redesigning styles, themes, or page layouts → use docx-design-ops
- Documents with active tracked changes → accept/reject all changes in Word first
- Editing embedded object content (Excel tables, OLE objects)
- Live Word automation → use a COM/MCP server

---

## Core Workflow

**Always follow this order. Never skip the copy and inspection steps.**

1. **Copy** the source file before any edits:
   ```bash
   cp original.docx working_copy.docx
   ```

2. **Extract and read** all text before changing anything:
   ```bash
   python3 docx-text-ops/scripts/extract.py working_copy.docx
   ```
   If the output contains a `WARNING: document contains tracked changes` line, stop and ask the user to accept or reject all changes in Word before proceeding.

3. **Unpack** for XML-level inspection or editing:
   ```bash
   python3 docx-text-ops/scripts/unpack.py working_copy.docx unpacked/
   ```

4. **Identify edit surfaces** — determine which surfaces contain the target text:
   - Body paragraphs and table cells (`word/document.xml`)
   - Headers / footers (`word/header*.xml`, `word/footer*.xml`)
   - Footnotes (`word/footnotes.xml`)
   - Endnotes (`word/endnotes.xml`)
   - Comments (`word/comments.xml`)
   - Text boxes (`<w:txbxContent>` elements in `word/document.xml`)

5. **Confirm scope** when the same text appears in multiple surfaces. Ask before editing headers, footers, or footnotes if only body text was requested.

6. **Edit** XML with the Edit tool. Apply the smallest possible text-only change. See [text-edit-recipes.md](references/text-edit-recipes.md).

7. **Pack**:
   ```bash
   python3 docx-text-ops/scripts/pack.py unpacked/ edited.docx
   ```

8. **Validate** — re-extract and compare:
   ```bash
   python3 docx-text-ops/scripts/extract.py edited.docx
   ```
   Check: old text gone, new text present, no duplicates, no leftover placeholders.

---

## Edit Safety Rules

- Preserve all non-text XML (styles, numbering, relationships, drawing anchors).
- Do not delete or renumber relationship IDs (`r:id`, `rId*`).
- Do not edit `.rels` files for text-only changes.
- Use namespace-aware XML parsing for programmatic edits — not regex string replacement.
- Use the Edit tool for targeted changes; do not write throwaway Python scripts.
- **Never edit `<w:ins>` or `<w:del>` runs** — these are tracked-change marks. Editing them corrupts revision state.
- Keep replacements close in length when layout preservation matters (narrow table cells, text boxes).
- Add `xml:space="preserve"` to `<w:t>` elements with leading or trailing spaces.

Read [docx-text-structure.md](references/docx-text-structure.md) for details on the XML structure and the split-run problem.

---

## Placeholder Detection

After editing, always run:

```bash
python3 docx-text-ops/scripts/extract.py edited.docx | grep -iE "xxxx|lorem|ipsum|tbd|placeholder"
```

If grep returns results, fix them before declaring success.

---

## Validation Checklist

- [ ] Old text is gone
- [ ] Replacement text is present and correct
- [ ] No duplicated passages
- [ ] No leftover placeholder text
- [ ] Headers, footers, footnotes, and comments were not unintentionally changed
- [ ] File opens without errors (`python3 -c "import docx; docx.Document('edited.docx'); print('OK')"`)
- [ ] (If layout risk) Visual spot-check shows no overflow in table cells or text boxes

---

## Dependencies

- `pip install python-docx` — text extraction (`extract.py`)
- `docx-text-ops/scripts/unpack.py` — unpack DOCX to XML
- `docx-text-ops/scripts/pack.py` — repack with OPC-correct ordering
