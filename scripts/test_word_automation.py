"""Live regression checks using only disposable Word documents/instances."""

import hashlib
from datetime import datetime
import json
from pathlib import Path
import subprocess
import sys
import unittest

from word_session import (LOCAL_TEMP, WordSession, find_open_document, pythoncom,
                          run_bounded, win32com, win32process)

from docx import Document
import pymupdf

SCRIPTS = Path(__file__).resolve().parent
QA = LOCAL_TEMP / f"word-automation-qa-{datetime.now():%Y%m%d-%H%M%S-%f}"


def cli(script, *arguments):
    result = subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, arguments), "--timeout", "40"],
                            text=True, encoding="utf-8", errors="replace",
                            capture_output=True, timeout=90)
    print(result.stdout, end="", flush=True)
    print(result.stderr, end="", flush=True)
    return result


class WordAutomationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        QA.mkdir(exist_ok=True)

    def fixture(self, name):
        path = QA / f"{name}.docx"
        document = Document()
        document.add_heading("Live editing check", level=1)
        document.add_paragraph("原本的文字 Original text.")
        document.save(path)
        return path

    def changes(self, name, old="原本的文字", new="修改後的文字", count=1):
        path = QA / f"{name}.json"
        path.write_text(json.dumps([{"old": old, "new": new, "count": count}], ensure_ascii=False), encoding="utf-8")
        return path

    def test_01_closed_edit_and_render(self):
        path = self.fixture("closed")
        changes = self.changes("closed")
        result = cli("edit_word.py", path, "--changes", changes, "--in-place")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("background", result.stdout)
        self.assertIn("修改後的文字", Document(path).paragraphs[1].text)
        result = cli("render_docx.py", path, "--out-dir", QA / "closed-render")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue((QA / "closed-render/page-1.png").is_file())

    def test_02_open_edit_unsaved_export_and_timeout(self):
        path = self.fixture("live")
        changes = self.changes("live")
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        document = None
        other = None
        try:
            word.Visible = True
            document = word.Documents.Open(str(path))
            document.Content.InsertAfter(" UNSAVED_VIEW_MARKER")
            other = word.Documents.Add()
            other.Content.Text = "UNRELATED_UNSAVED_DOCUMENT"
            # Leave the other document active: the editor must identify by path.
            original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
            result = cli("render_docx.py", path, "--out-dir", QA / "live-render")
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("Attached to open document", result.stdout)
            self.assertFalse(document.Saved)
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), original_hash)
            with pymupdf.open(QA / "live-render/live.pdf") as pdf:
                self.assertIn("UNSAVED_VIEW_MARKER", "".join(p.get_text() for p in pdf))
            self.assertTrue(word.Visible)

            # Invalid plans must fail before mutating any part of an open document.
            before = document.Content.Text
            invalid = self.changes("wrong-count", count=2)
            result = cli("edit_word.py", path, "--changes", invalid, "--in-place")
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(document.Content.Text, before)
            self.assertFalse(document.Saved)

            result = cli("edit_word.py", path, "--changes", changes, "--in-place")
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("Attached to open document", result.stdout)
            self.assertIn("修改後的文字", document.Content.Text)
            self.assertIn("UNSAVED_VIEW_MARKER", document.Content.Text)
            self.assertTrue(word.Visible)
            self.assertTrue(word.ScreenUpdating)
            self.assertEqual(word.Documents.Count, 2)
            self.assertIn("UNRELATED_UNSAVED_DOCUMENT", other.Content.Text)
            self.assertFalse(other.Saved)
            self.assertEqual(word.ActiveDocument.Name, other.Name)

            # Parent watchdog interrupts an indefinitely blocked attached worker.
            hanger = QA / "hang-attached.py"
            hanger.write_text(
                "import sys, time\n"
                f"sys.path.insert(0, {str(SCRIPTS)!r})\n"
                "from word_session import WordSession\n"
                f"with WordSession({str(path)!r}) as session:\n"
                "    time.sleep(60)\n", encoding="utf-8")
            pid = win32process.GetWindowThreadProcessId(document.ActiveWindow.Hwnd)[1]
            self.assertEqual(run_bounded(hanger, [], 5), 124)
            self.assertEqual(win32process.GetWindowThreadProcessId(document.ActiveWindow.Hwnd)[1], pid)
            self.assertTrue(word.Visible)
            self.assertIn("修改後的文字", document.Content.Text)
        finally:
            if other is not None:
                other.Close(False)
            if document is not None:
                document.Close(False)
            word.Quit()
            pythoncom.CoUninitialize()

    def test_03_copy_preserves_source_and_updates_native_contents(self):
        path = self.fixture("source")
        output = QA / "copy.docx"
        if output.exists():
            output.unlink()
        before = path.read_bytes()
        changes = self.changes("copy")
        result = cli("edit_word.py", path, "--changes", changes, "--output", output)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(path.read_bytes(), before)
        self.assertIn("修改後的文字", Document(output).paragraphs[1].text)
        # Native automatic contents must include a heading changed by the editor.
        with WordSession(output) as session:
            doc = session.document
            doc.Content.InsertBefore("\r")
            doc.TablesOfContents.Add(doc.Range(0, 0), UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=1)
            doc.Save()
        changes = self.changes("heading", old="Live editing check", new="Updated heading")
        result = cli("edit_word.py", output, "--changes", changes, "--in-place")
        self.assertEqual(result.returncode, 0, result.stderr)
        with WordSession(output, read_only=True) as session:
            doc = session.document
            self.assertEqual(doc.TablesOfContents.Count, 1)
            self.assertIn("Updated heading", doc.TablesOfContents.Item(1).Range.Text)
            self.assertNotIn("Live editing check", doc.TablesOfContents.Item(1).Range.Text)
            # A search which occurs only inside a native field cannot edit it.
        # Use a unique native field result, without paragraph/control characters.
        with WordSession(output) as session:
            doc = session.document
            area = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            field = doc.Fields.Add(area, 31)  # Native DATE field
            field.Update()
            date_text = field.Result.Text
            doc.Save()
        protected = self.changes("field-only", old=date_text, new="damaged")
        saved = output.read_bytes()
        result = cli("edit_word.py", output, "--changes", protected, "--in-place")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("protected", result.stderr)
        self.assertEqual(output.read_bytes(), saved)

    def test_04_visible_new_document_stays_open(self):
        path = self.fixture("visible")
        changes = self.changes("visible")
        result = cli("edit_word.py", path, "--changes", changes, "--in-place", "--visible")
        self.assertEqual(result.returncode, 0, result.stderr)
        pythoncom.CoInitialize()
        try:
            document = find_open_document(path)
            self.assertIsNotNone(document)
            word = document.Application
            self.assertTrue(word.Visible)
            self.assertIn("修改後的文字", document.Content.Text)
            # Only the visible instance created for this test is closed here.
            document.Close(False)
            word.Quit()
        finally:
            pythoncom.CoUninitialize()

    def test_05_timeout_reaps_only_owned_background_process(self):
        path = self.fixture("timeout-private")
        hanger = QA / "hang-private.py"
        identity = QA / "private-identity.json"
        identity.unlink(missing_ok=True)
        hanger.write_text(
            "import sys, time\n"
            f"sys.path.insert(0, {str(SCRIPTS)!r})\n"
            "from word_session import WordSession\n"
            f"with WordSession({str(path)!r}, state_path=sys.argv[-1]) as session:\n"
            f"    from pathlib import Path\n    Path({str(identity)!r}).write_text(Path(sys.argv[-1]).read_text())\n"
            "    time.sleep(60)\n", encoding="utf-8")
        self.assertEqual(run_bounded(hanger, [], 10), 124)
        self.assertTrue(identity.is_file(), "Background Word must start before testing its timeout cleanup")
        record = json.loads(identity.read_text())
        # Termination is asynchronous; wait briefly for the specific owned PID.
        import time
        for _ in range(20):
            if record["pid"] not in win32process.EnumProcesses():
                break
            time.sleep(0.1)
        self.assertNotIn(record["pid"], win32process.EnumProcesses())


if __name__ == "__main__":
    unittest.main(verbosity=2)
