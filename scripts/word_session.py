"""Attach to the exact open Word document, or own a background instance.

CLI callers run COM in a child process so even a blocked COM call has a deadline.
Never close or terminate a Word process that was already running.
"""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time

LOCAL_TEMP = Path(__file__).resolve().parents[1] / ".tmp"
LOCAL_TEMP.mkdir(exist_ok=True)
os.environ["TEMP"] = os.environ["TMP"] = str(LOCAL_TEMP)
tempfile.tempdir = str(LOCAL_TEMP)

import pythoncom
import win32api
import win32com.client
import win32process


def normalized(path):
    return os.path.normcase(os.path.abspath(str(path)))


def retry(call, seconds=10):
    """Retry only Word's temporary 'busy' responses, with a finite deadline."""
    deadline = time.monotonic() + seconds
    while True:
        try:
            return call()
        except pythoncom.com_error as exc:
            if exc.hresult not in (-2147418111, -2147417846) or time.monotonic() >= deadline:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)


def find_open_document(path):
    """ROT lookup covers multiple Word instances without opening a second copy."""
    target = normalized(path)
    rot = pythoncom.GetRunningObjectTable()
    context = pythoncom.CreateBindCtx(0)
    for moniker in rot.EnumRunning():
        name = moniker.GetDisplayName(context, None)
        if normalized(name) == target:
            document = win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
            if normalized(retry(lambda: document.FullName)) == target:
                return document
    try:
        word = retry(lambda: win32com.client.GetActiveObject("Word.Application"))
    except pythoncom.com_error as exc:
        if exc.hresult == -2147221021:  # MK_E_UNAVAILABLE: no registered instance
            return None
        raise
    for index in range(1, retry(lambda: word.Documents.Count) + 1):
        document = retry(lambda: word.Documents.Item(index))
        if normalized(retry(lambda: document.FullName)) == target:
            return document
    return None


def process_birth(pid):
    handle = win32api.OpenProcess(0x1000, False, pid)
    try:
        return str(win32process.GetProcessTimes(handle)["CreationTime"])
    finally:
        handle.Close()


class WordSession:
    def __init__(self, path, *, read_only=False, visible=False, state_path=None):
        self.path = Path(path).resolve()
        self.read_only = read_only
        self.visible = visible
        self.state_path = state_path
        self.word = self.document = None
        self.attached = False
        self.owns_process = False
        self.keep_open = False

    def __enter__(self):
        pythoncom.CoInitialize()
        try:
            if not self.path.is_file():
                raise FileNotFoundError(self.path)
            self.document = find_open_document(self.path)
            if self.document is not None:
                self.attached = True
                self.word = self.document.Application
                if not self.read_only and self.document.ReadOnly:
                    raise RuntimeError("The open document is read-only.")
                if not self.read_only:
                    self.word.ScreenUpdating = True
                print(f"Attached to open document: {self.path}", flush=True)
                return self
            # An inaccessible open document must not be reopened behind a lock dialog.
            owner_files = [self.path.with_name("~$" + self.path.name),
                           self.path.with_name("~$" + self.path.name[2:])]
            if any(owner_file.exists() for owner_file in owner_files):
                raise RuntimeError("Word's lock file exists, but the document could not be attached. "
                                   "Check Word dialogs/session permissions; no second copy was opened.")
            baseline = set(win32process.EnumProcesses())
            self.word = win32com.client.DispatchEx("Word.Application")
            # Word exposes Hwnd on Window, not Application. Use a disposable
            # blank document to identify the private process before opening files.
            self.document = self.word.Documents.Add()
            pid = win32process.GetWindowThreadProcessId(self.document.ActiveWindow.Hwnd)[1]
            self.owns_process = pid not in baseline
            if self.state_path and self.owns_process:
                Path(self.state_path).write_text(json.dumps({"pid": pid, "birth": process_birth(pid)}), encoding="utf-8")
            if not self.owns_process:
                raise RuntimeError("Word returned an existing process; refusing to change its application settings.")
            self.document.Close(False)
            self.document = None
            self.word.Visible = self.visible
            self.word.ScreenUpdating = True
            self.word.DisplayAlerts = 0
            self.document = self.word.Documents.Open(
                str(self.path), ConfirmConversions=False, ReadOnly=self.read_only,
                AddToRecentFiles=False, Visible=self.visible,
            )
            if not self.read_only and self.document.ReadOnly:
                raise RuntimeError("Word opened the document read-only; no edits were made.")
            print(f"Opened {'visible' if self.visible else 'background'} document: {self.path}", flush=True)
            return self
        except BaseException:
            self.__exit__(*sys.exc_info())
            raise

    def refresh(self):
        if self.attached or self.visible:
            self.word.ScreenRefresh()
            pythoncom.PumpWaitingMessages()

    def __exit__(self, exc_type, exc, traceback):
        try:
            if not self.attached:
                if self.keep_open and exc_type is None:
                    self.word.DisplayAlerts = -1
                    if self.state_path:
                        Path(self.state_path).unlink(missing_ok=True)
                else:
                    if self.document is not None:
                        self.document.Close(False)
                    if self.word is not None and self.owns_process:
                        self.word.Quit()
        finally:
            self.document = self.word = None
            pythoncom.CoUninitialize()


def run_bounded(script, arguments, timeout):
    """A watchdog outside COM bounds Open, edits, export, and cleanup together."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    with tempfile.TemporaryDirectory(prefix="word-worker-", dir=LOCAL_TEMP) as directory:
        state = Path(directory) / "owned-word.json"
        command = [sys.executable, "-u", str(script), *arguments, "--worker-state", str(state)]
        worker = subprocess.Popen(command)
        try:
            return worker.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            worker.kill()
            worker.wait(timeout=10)
            # PID and creation time must both match the instance created by this worker.
            if state.exists():
                record = json.loads(state.read_text(encoding="utf-8"))
                try:
                    handle = win32api.OpenProcess(0x1001, False, record["pid"])
                    try:
                        birth = str(win32process.GetProcessTimes(handle)["CreationTime"])
                        if birth == record["birth"]:
                            win32api.TerminateProcess(handle, 1)
                    finally:
                        handle.Close()
                except (OSError, win32api.error):
                    pass  # Already exited, or cleanup not permitted.
            print(f"Word automation timed out after {timeout}s. Existing Word windows were preserved. "
                  "Open documents may contain partial, unsaved changes; inspect them before retrying.", file=sys.stderr)
            return 124


def update_directories(document):
    """Update native captions/references/TOCs without touching EndNote ADDINs."""
    def selected_fields(types):
        return [document.Fields.Item(i) for i in range(1, document.Fields.Count + 1)
                if document.Fields.Item(i).Type in types]

    for field in selected_fields({12, 3, 37}):  # SEQ, REF, PAGEREF
        field.Update()
    previous = None
    for _ in range(3):
        document.Repaginate()
        covered = set()
        for collection in (document.TablesOfContents, document.TablesOfFigures):
            for index in range(1, collection.Count + 1):
                item = collection.Item(index)
                covered.add(item.Range.Start)
                item.Update()
        for field in selected_fields({13}):  # TOC, including custom TOC-based lists
            if field.Result.Start not in covered:
                field.Update()
        document.Repaginate()
        for collection in (document.TablesOfContents, document.TablesOfFigures):
            for index in range(1, collection.Count + 1):
                collection.Item(index).UpdatePageNumbers()
        for field in selected_fields({37}):
            field.Update()
        signature = (document.ComputeStatistics(2), tuple(f.Result.Text for f in selected_fields({13})))
        if signature == previous:
            return
        previous = signature
    raise RuntimeError("Directory pagination did not stabilize within three rounds; inspect the document.")
