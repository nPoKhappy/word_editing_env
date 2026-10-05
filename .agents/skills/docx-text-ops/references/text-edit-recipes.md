# DOCX Text Edit Recipes

[日本語](text-edit-recipes.ja.md)

## Recipe 1: Replace Plain Text In One Run

Use when the target text appears inside a single `<w:t>` element.

```xml
<w:t>Old product name</w:t>
```

Replace only the text node:

```xml
<w:t>New product name</w:t>
```

Keep surrounding `w:rPr`, paragraph properties, and relationship IDs unchanged.

## Recipe 2: Replace Text Split Across Runs

When a phrase is split across runs, preserve formatting unless the user explicitly approves simplifying it.

Before:

```xml
<w:r><w:t>Old </w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>Phrase</w:t></w:r>
```

After:

```xml
<w:r><w:t>New </w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>Phrase</w:t></w:r>
```

## Recipe 3: Header, Footer, Footnote, Or Comment Edits

Confirm scope before editing when the same phrase appears in multiple document surfaces. Body-only requests should not silently change headers, footers, notes, or comments.

Common files:

- `word/header*.xml`
- `word/footer*.xml`
- `word/footnotes.xml`
- `word/endnotes.xml`
- `word/comments.xml`

## Recipe 4: Placeholder Cleanup

Find placeholders first:

```bash
python3 docx-text-ops/scripts/extract.py working_copy.docx | grep -iE "xxxx|lorem|ipsum|tbd|placeholder"
```

Edit each occurrence deliberately and validate with the same grep after repacking.

## Recipe 5: Translation Pass

1. Extract all text and keep it as a reference.
2. Confirm whether headers, footers, notes, and comments are in scope.
3. Edit XML text nodes one surface at a time.
4. Keep replacements close in length when text boxes or narrow table cells are involved.
5. Re-extract and compare for missing or duplicated text.

## Validation Checklist

- Old text is gone.
- New text is present and correct.
- No placeholders remain.
- Headers, footers, notes, and comments were changed only when intended.
- The file opens in Word or LibreOffice.
