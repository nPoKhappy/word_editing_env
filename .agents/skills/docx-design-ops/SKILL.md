---
name: docx-design-ops
description: "Use this skill when the user wants to improve the visual design of an
  existing Word (.docx) file — fonts, colors, heading styles, paragraph spacing,
  line height, or the document theme. Triggers: change the body font, update heading
  colors, fix line spacing, apply a consistent style, update the color palette, make
  the document look more professional, the design is poor, text looks cramped.
  Use this skill when the text content is already correct and the visual presentation
  needs improvement. Do NOT use for editing text content — use docx-text-ops for that.
  Do NOT use for creating a new document from scratch — use python-docx scripting for that.
  Do NOT use for documents with active tracked changes — accept or reject all changes
  in Word before editing."
license: Original work, no third-party license constraints
---

# DOCX Design Operations Skill

[日本語](SKILL.ja.md)

Visual design editing for existing `.docx` files. Changes fonts, colors, heading styles, paragraph spacing, and the document theme while preserving all text content, images, tables, and relationships.

## Quick Reference

| Task | Approach |
|------|----------|
| Audit current design | `python3 docx-design-ops/scripts/inspect_design.py document.docx` |
| Inspect raw XML | Unpack → read `word/styles.xml`, `word/theme/theme1.xml` |
| Apply design changes | Unpack → edit XML with Edit tool or run `apply_design.py` → pack |
| Batch apply multiple changes | `python3 docx-design-ops/scripts/apply_design.py unpacked/ spec.json` |
| Preview changes before applying | `apply_design.py unpacked/ spec.json --dry-run` |
| Validate after edit | `inspect_design.py edited.docx --json` and diff with before |

---

## Non-Goals

- Editing text content → use docx-text-ops
- Creating new documents from scratch → use python-docx scripting
- Documents with active tracked changes → accept/reject all changes in Word first
- Table cell shading, borders, and cell-level design → requires direct XML editing in `document.xml`
- Image and drawing manipulation
- Animations and transitions

---

## Core Workflow

**Always follow this order. Never skip the copy and inspection steps.**

1. **Copy** the source file before any edits:
   ```bash
   cp original.docx working_copy.docx
   ```

2. **Inspect** the current design state:
   ```bash
   python3 docx-design-ops/scripts/inspect_design.py working_copy.docx
   ```
   Review: current fonts, theme colors, style definitions, and inline override count.

3. **Unpack** for XML-level editing:
   ```bash
   python3 docx-text-ops/scripts/unpack.py working_copy.docx unpacked/
   ```

4. **Audit inline overrides** — if the override count is high, style-level changes will have limited effect:
   ```bash
   grep -n "w:rFonts\|w:color\|<w:sz " unpacked/word/document.xml | head -20
   ```
   Decide whether to remove the overrides (Recipe 5) before applying new styles.

5. **Plan changes** — decide the approach:
   - **Theme font change** → edit `word/theme/theme1.xml` (affects all theme-referenced styles automatically)
   - **Theme color change** → edit `word/theme/theme1.xml` (affects all theme-color-referenced elements)
   - **Style-level change** → edit `word/styles.xml` (affects styled paragraphs)
   - **Inline override removal** → edit `word/document.xml` (removes direct formatting on specific paragraphs/runs)

6. **Edit** XML with the Edit tool or run `apply_design.py`. See [design-recipes.md](references/design-recipes.md).

7. **Pack**:
   ```bash
   python3 docx-text-ops/scripts/pack.py unpacked/ edited.docx
   ```

8. **Validate** — compare before and after:
   ```bash
   python3 docx-design-ops/scripts/inspect_design.py original.docx --json > before.json
   python3 docx-design-ops/scripts/inspect_design.py edited.docx   --json > after.json
   diff before.json after.json
   ```
   Then open in Word or LibreOffice for visual confirmation.

---

## Design Safety Rules

- Never modify `w:basedOn` or `w:next` attributes — these define style inheritance and must remain intact.
- Never edit `.rels` files.
- Never change element order within `w:style` definitions — only modify `w:rPr` and `w:pPr` children.
- Only edit `word/document.xml` when explicitly removing inline overrides (Recipe 5). Do not set design properties there.
- Always use `--dry-run` before running `apply_design.py` on an unfamiliar document.
- When setting a direct hex color with `w:color w:val`, remove `w:themeColor`, `w:themeTint`, and `w:themeShade` from the same element.
- When changing `w:sz`, always sync `w:szCs` to the same value.

---

## Validation Checklist

- [ ] Design state matches intended changes (`inspect_design.py` output)
- [ ] Text content is unchanged (`extract.py` line count matches original)
- [ ] File opens without errors (`python3 -c "import docx; docx.Document('edited.docx'); print('OK')"`)
- [ ] Visual spot-check in Word or LibreOffice confirms the design improvements
- [ ] JSON diff shows no unintended changes outside the target styles/colors

---

## Dependencies

- `pip install python-docx` — required for text integrity check only
- `docx-text-ops/scripts/unpack.py` — unpack DOCX to XML (shared with docx-text-ops)
- `docx-text-ops/scripts/pack.py` — repack with OPC-correct ordering (shared with docx-text-ops)
- `docx-design-ops/scripts/inspect_design.py` — design audit (standard library only)
- `docx-design-ops/scripts/apply_design.py` — batch apply (standard library only)

Read [docx-design-structure.md](references/docx-design-structure.md) for XML reference and size unit tables.
