---
name: word-editing
description: Create, edit, and visually check local Word DOCX files. In this checkout, prefer OfficeCLI for ordinary authoring and retain Python/Word for native fields, live editing and rendered QA.
---

# Word editing in this workspace

## OfficeCLI routing in this checkout

The user's OfficeCLI migration supersedes this skill's older Python-first
authoring recipes. Read [officecli](../officecli/SKILL.md). Use OfficeCLI for
ordinary package reads, authoring and closed-file edits; keep native Word for
live documents, captions, selective directory updates, EndNote and final render.
Use this checkout's `.venv/Scripts/python.exe` only for those adapters and QA.
The full legacy dependency versions remain in `requirements.lock.txt`.

- Default Word typography: Chinese text uses 標楷體 and English or Latin text uses Times New Roman, including mixed-language paragraphs. With `python-docx`, set the run or style's `w:rFonts` `w:eastAsia` to `標楷體` and `w:ascii`/`w:hAnsi` to `Times New Roman`; setting only `font.name` does not control the East Asian font. Apply this to new content and to a whole document when the user requests font standardization. For a local edit of an existing document, preserve unrelated formatting and honor any explicit template or document-specific font requirement.
- Before Python commands that use `pywin32`, set `TEMP` and `TMP` to `D:\word_editing_env_officeCLI\.tmp`. The default temp location is not writable in this environment.
- Prefer OfficeCLI `get/view/query`, semantic `add/set/batch`, and `merge` for ordinary DOCX work. Python libraries remain fallbacks for demonstrated unsupported cases, rather than the default authoring path.
- For existing files, the installed GitHub skills `../docx-text-ops/SKILL.md` and `../docx-design-ops/SKILL.md` provide focused XML editing procedures. Use their scripts through this workspace's Python virtual environment, and adapt their shell examples to PowerShell.
- Preserve a supplied source document unless the user asks for an in-place edit. Inspect its structure and styles before changing it, then verify requested text, tables, fields, and page settings after the edit.
- Render every final DOCX and inspect every page image. Run `scripts/render_docx.py input.docx --out-dir output-folder` to export with installed Microsoft Word and create page PNGs. If Word automation is unavailable, report that visual verification did not complete; XML or text checks alone do not prove the layout.
- For PDF inspection, use `pypdf`, `pdfplumber`, or `pymupdf`; use `pypdfium2` for page images. `Pillow`, `reportlab`, and `matplotlib` support images, PDFs, and figures when the document needs them.
- Do not assume Word COM succeeds just because `pywin32` imports. Test the actual export when delivering a DOCX. In this sandbox, a normal session returned COM error `0x80070520`; the same export succeeded in an approved elevated session.

## Inline mathematical symbols

- The user reported that displayed equations were correct but subscripts/superscripts in body prose were often left unformatted. Treat inline symbols as a separate QA surface during authorized Word edits, including affected table cells and captions.
- Match each symbol to its equation definition. Use native Word subscript/superscript run formatting for simple indices and chemical formulas; prefer editable inline OMath for combined indices, hats, bars, or other compound notation. Inline OMath does not receive a displayed-equation number.
- Check examples such as H2S/SO2, L_MSE, theta_out, G_base, alpha_g, and K_ij^SS according to their intended notation. Literal underscores/carets or a mixture of Unicode subscript characters are not evidence of correct Word typography. Do not turn ordinary prose numbers into indices.
- Keep surrounding prose on the baseline and preserve its fonts and base size; do not flatten the mathematical indices to the body-text size. Recheck after literal replacements or font normalization, which can reset or spread run formatting.
- Verify both the native run/OMath properties and the rendered inline appearance. Correct standalone equations alone do not establish that body symbols are correct.
