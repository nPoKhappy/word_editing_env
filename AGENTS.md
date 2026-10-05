# Word editing workspace

For DOCX work in this folder, read `.agents/skills/word-editing/SKILL.md`.
For targeted text edits in existing DOCX files, also read
`.agents/skills/docx-text-ops/SKILL.md`. For visual style changes, read
`.agents/skills/docx-design-ops/SKILL.md`.
Use `.venv/Scripts/python.exe` for Python commands. Set `TEMP` and `TMP` to
`D:\word_editing_env\.tmp` before importing `win32com.client` or running Word
automation, because this environment cannot write to the default temp folder.
For Word text added in this workspace, use 標楷體 for Chinese and Times New Roman
for English/Latin text, as described in the local word-editing skill.
