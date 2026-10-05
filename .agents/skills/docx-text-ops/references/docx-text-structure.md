# DOCX Text Structure Reference

[日本語](docx-text-structure.ja.md)

## DOCX As A ZIP Package

A `.docx` file is a ZIP package. Unpack it before XML-level inspection:

```bash
python3 docx-text-ops/scripts/unpack.py document.docx unpacked/
```

Key files:

```text
word/document.xml          body paragraphs, tables, and text boxes
word/header*.xml           headers
word/footer*.xml           footers
word/footnotes.xml         footnotes
word/endnotes.xml          endnotes
word/comments.xml          comments
word/styles.xml            styles; do not edit for text-only changes
word/_rels/*.rels          relationships; do not edit for text-only changes
```

## Text Containers

Most visible text is stored in `w:t` elements inside runs:

```xml
<w:p>
  <w:r>
    <w:t>Body text</w:t>
  </w:r>
</w:p>
```

Tables, headers, footers, comments, footnotes, endnotes, and text boxes use the same paragraph/run/text pattern.

## Split-Run Problem

A single visible phrase can be split across multiple runs because of formatting, proofing, fields, or editing history:

```xml
<w:r><w:t>Quarterly </w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>Report</w:t></w:r>
```

Do not assume a visible phrase appears as one XML string. Edit the smallest safe run content, or preserve the run split while replacing text.

## Whitespace

Use `xml:space="preserve"` when a `w:t` value starts or ends with whitespace:

```xml
<w:t xml:space="preserve"> leading or trailing space </w:t>
```

Removing it can silently change rendered text.

## Tracked Changes

Tracked insertions and deletions use elements such as `w:ins` and `w:del`. Do not edit those runs. Ask the user to accept or reject tracked changes in Word before XML editing.

## Relationships

Do not edit `.rels` files for text-only changes. They connect the document to images, headers, footers, comments, and other parts.
