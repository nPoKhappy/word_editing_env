# DOCX Design Recipes

[日本語](design-recipes.ja.md)

Common patterns for visual design changes in existing `.docx` files. These recipes assume the file has been unpacked with `docx-text-ops/scripts/unpack.py`.

## Recipe 1: Change Body Font And Size

Target: `word/styles.xml`

1. Inspect the current state.
   ```bash
   python3 docx-design-ops/scripts/inspect_design.py working_copy.docx
   ```
2. Update the `Normal` style and, when needed, `w:docDefaults`.
3. Set both `w:sz` and `w:szCs` to the same half-point value.
4. Prefer style-level edits over direct formatting in `word/document.xml`.

Example style change:

```json
{
  "styles": {
    "Normal": {
      "rPr": { "rFonts": "Calibri", "sz": 22 },
      "pPr": { "spacingAfter": 160, "line": 276, "lineRule": "auto" }
    }
  }
}
```

Run:

```bash
python3 docx-design-ops/scripts/apply_design.py unpacked/ spec.json --dry-run
python3 docx-design-ops/scripts/apply_design.py unpacked/ spec.json
```

## Recipe 2: Modify Heading Styles

Target: `word/styles.xml`

1. Find heading style IDs by searching for `w:styleId="Heading"` or `w:name w:val="heading"`.
2. Edit only `w:rPr` and `w:pPr`.
3. Preserve `w:basedOn`, `w:next`, outline levels, and relationship IDs.

Example:

```json
{
  "styles": {
    "Heading1": {
      "rPr": { "rFonts": "Calibri Light", "sz": 36, "color": "003366", "bold": true },
      "pPr": { "spacingBefore": 360, "spacingAfter": 120, "line": 360, "lineRule": "auto" }
    }
  }
}
```

## Recipe 3: Update Theme Colors

Target: `word/theme/theme1.xml`

Changing theme color slots affects content that references those slots with `w:themeColor`. Direct hex colors in `w:color w:val` are not affected.

```json
{
  "theme": {
    "colors": {
      "accent1": "003366",
      "accent2": "CC6600",
      "hlink": "0563C1"
    }
  }
}
```

Validate with:

```bash
python3 docx-design-ops/scripts/inspect_design.py edited.docx --json
```

## Recipe 4: Update Theme Fonts

Target: `word/theme/theme1.xml`

Theme font changes affect styles that use theme references such as `majorHAnsi` or `minorHAnsi`.

```json
{
  "theme": {
    "majorFont": "Aptos Display",
    "minorFont": "Aptos"
  }
}
```

If styles use direct font names, update the style definitions instead.

## Recipe 5: Handle Inline Overrides

Target: usually `word/document.xml`

Inline formatting takes precedence over styles. Use `inspect_design.py` to count direct run and paragraph formatting before deciding whether style changes are enough.

Remove inline overrides only when the user explicitly wants style-level normalization. Do not remove structural properties such as paragraph style, numbering, or section properties.

## Validation Checklist

- `inspect_design.py` shows the intended theme or style changes.
- Text extraction before and after has the same content.
- The edited file opens in Word or LibreOffice.
- No `.rels` files were modified.
- Visual spot-check confirms headings, body text, and spacing look correct.
