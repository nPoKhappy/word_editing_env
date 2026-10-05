# DOCX Design Structure Reference

[日本語](docx-design-structure.ja.md)

## Design Files

A `.docx` file is a ZIP package. Design changes usually target:

```text
word/styles.xml            paragraph, character, table, and list styles
word/theme/theme1.xml      theme fonts and color palette
word/document.xml          body content; only touch for explicit inline override cleanup
```

Do not edit `.rels` files for design-only changes.

## Two-Layer Design Model

Word resolves design through two layers:

```text
Layer 1: style and theme definitions
Layer 2: inline direct formatting in document.xml
```

Inline direct formatting wins over style definitions. If a run has a direct red color in `word/document.xml`, changing the heading style color will not affect that run.

## `word/styles.xml`

Important areas:

- `w:docDefaults`: document-wide default run and paragraph properties.
- `w:style`: style definitions for paragraphs, characters, tables, and lists.
- `w:rPr`: run properties such as font, size, color, bold, and italic.
- `w:pPr`: paragraph properties such as spacing, line height, alignment, and indentation.

Critical rule: preserve `w:basedOn` and `w:next`. These define style inheritance and follow-on paragraph behavior.

## `word/theme/theme1.xml`

Important areas:

- `a:fontScheme`: theme major/minor fonts.
- `a:clrScheme`: named theme color slots such as `accent1`, `accent2`, `dk1`, `lt1`, and hyperlink colors.

Theme-referenced styles update automatically when theme slots change. Direct colors and direct font names do not.

## Units

| Property | Unit | Example |
|----------|------|---------|
| `w:sz`, `w:szCs` | half-points | `24` = 12pt |
| `w:spacing w:before`, `w:after` | twentieths of a point | `240` = 12pt |
| `w:spacing w:line` with `auto` | twentieths of a point | `276` = 1.15 line spacing |
| `w:ind w:left`, `w:right` | twentieths of a point | `720` = 0.5 inch |

## Namespaces

| Prefix | Namespace |
|--------|-----------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
