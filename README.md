# word_editing_env

Windows workspace for editing Word documents with Python and visually checking them with Microsoft Word.

## Setup

Use Python 3.12 and an installed desktop copy of Microsoft Word for Word automation and rendering.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.lock.txt
New-Item -ItemType Directory -Force .tmp | Out-Null
$env:TEMP = Join-Path $PWD '.tmp'
$env:TMP = $env:TEMP
```

Use `.venv\Scripts\python.exe` for Python commands. Set `TEMP` and `TMP` before importing `win32com.client` or running Word automation.

## Render a document

```powershell
.\.venv\Scripts\python.exe scripts\render_docx.py input.docx --out-dir .tmp\rendered
```

The renderer exports a PDF through Microsoft Word and writes a PNG for each page. Inspect the page images after editing.

## Workspace guidance

See `AGENTS.md` and `.agents/skills/` for editing workflows. New Chinese text uses 標楷體; English and Latin text uses Times New Roman.

The `thesis_outline/` folder, virtual environment, caches, and temporary files are excluded from version control. Local thesis documents are kept on this computer.
