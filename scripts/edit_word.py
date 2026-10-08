"""Apply exact main-body text edits, visibly if the target is open in Word."""

import argparse
from datetime import datetime
import json
from pathlib import Path
import shutil
import sys

from word_session import (LOCAL_TEMP, WordSession, find_open_document,
                          pythoncom, run_bounded, update_directories)


def plan_replacements(document, changes):
    """Resolve and validate every target before the first mutation."""
    if document.TrackRevisions or document.Revisions.Count:
        raise ValueError("Active tracked changes are unsupported; no edits were made.")
    if document.ProtectionType != -1:
        raise ValueError("The document is protected; no edits were made.")
    protected = []
    for index in range(1, document.Fields.Count + 1):
        field = document.Fields.Item(index)
        protected.append((field.Code.Start - 1, field.Result.End + 1))
    targets = []
    for change in changes:
        old, new = change["old"], change["new"]
        if not old or any(ord(c) < 32 for c in old + new):
            raise ValueError("Use nonempty literal text without paragraph/control characters.")
        found = []
        field_matches = 0
        position = document.Content.Start
        while position < document.Content.End:
            area = document.Range(position, document.Content.End)
            search = area.Find
            search.ClearFormatting()
            matched = search.Execute(FindText=old.replace("^", "^^"), MatchCase=True,
                                     MatchWholeWord=False, MatchWildcards=False,
                                     MatchSoundsLike=False, MatchAllWordForms=False,
                                     Forward=True, Wrap=0, Format=False)
            if not matched:
                break
            if area.Start < position or area.End <= position:
                raise RuntimeError("Word search returned a range outside its search boundary; no edits were made.")
            if area.Text != old:
                raise ValueError("Word search did not return an exact literal match.")
            overlapping = [end for start, end in protected if area.Start < end and area.End > start]
            if overlapping:
                # Generated TOC/citation text is not a main-body editing target.
                # Searching again from inside a TOC can return the same hit forever.
                field_matches += 1
                position = max(overlapping)
            else:
                found.append((area.Start, area.End, old, new))
                position = area.End
        expected = change.get("count", 1)
        if not isinstance(expected, int) or isinstance(expected, bool) or expected <= 0:
            raise ValueError("Each replacement count must be a positive integer.")
        if len(found) != expected:
            raise ValueError(f"Expected {expected} editable occurrence(s) of {old!r}, found {len(found)} "
                             f"({field_matches} field/directory match(es) protected); no edits were made.")
        targets.extend(found)
    targets.sort()
    for left, right in zip(targets, targets[1:]):
        if left[1] > right[0]:
            raise ValueError("Replacement targets overlap; no edits were made.")
    return list(reversed(targets))


def edit(args):
    source = args.docx.resolve()
    changes = json.loads(args.changes.read_text(encoding="utf-8-sig"))
    if not isinstance(changes, list) or not changes:
        raise ValueError("changes must be a nonempty JSON array of old/new/count objects")
    for change in changes:
        if not isinstance(change, dict) or not all(isinstance(change.get(key), str) for key in ("old", "new")):
            raise ValueError("Each change must have string old and new values.")
    target = source
    visible = args.visible
    if args.output:
        target = args.output.resolve()
        if target == source:
            raise ValueError("Use --in-place explicitly to edit the original.")
        pythoncom.CoInitialize()
        try:
            opened = find_open_document(source)
            if opened is not None:
                if not opened.Saved:
                    raise ValueError("The source has unsaved Word edits; save it before making an output copy.")
                visible = True
        finally:
            pythoncom.CoUninitialize()
        target.parent.mkdir(parents=True, exist_ok=True)
        with source.open("rb") as original, target.open("xb") as output:
            shutil.copyfileobj(original, output)
        print(f"Working copy: {target}", flush=True)
    with WordSession(target, visible=visible, state_path=args.worker_state) as session:
        document = session.document
        targets = plan_replacements(document, changes)
        if args.in_place:
            backup_dir = LOCAL_TEMP / "word-backups"
            backup_dir.mkdir(exist_ok=True)
            backup = backup_dir / f"{source.stem}-{datetime.now():%Y%m%d-%H%M%S-%f}.docx"
            shutil.copy2(source, backup)
            print(f"Disk-file backup: {backup}", flush=True)
        # Range editing does not move the user's selection or scroll position.
        session.word.UndoRecord.StartCustomRecord("Codex text edits")
        try:
            for start, end, old, new in targets:
                area = document.Range(start, end)
                if area.Text != old:
                    raise RuntimeError("Document content changed during editing; remaining edits were stopped.")
                area.Text = new
                area.Font.NameFarEast = "標楷體"
                area.Font.NameAscii = "Times New Roman"
                area.Font.NameOther = "Times New Roman"
                session.refresh()
                print(f"Updated: {old!r} -> {new!r}", flush=True)
            update_directories(document)
            document.Save()
            session.refresh()
            # A newly visible working copy remains open for review after success.
            session.keep_open = visible
        finally:
            session.word.UndoRecord.EndCustomRecord()
        print(f"Saved: {target}", flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--changes", type=Path, required=True)
    destination = parser.add_mutually_exclusive_group(required=True)
    destination.add_argument("--in-place", action="store_true")
    destination.add_argument("--output", type=Path)
    parser.add_argument("--visible", action="store_true", help="Show a new Word window and leave it open after success")
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--worker-state", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker_state:
        try:
            edit(args)
        except Exception as exc:
            print(f"Editing stopped: {exc}. An attached document may contain unsaved partial changes.", file=sys.stderr)
            sys.exit(1)
    else:
        sys.exit(run_bounded(Path(__file__), sys.argv[1:], args.timeout))
