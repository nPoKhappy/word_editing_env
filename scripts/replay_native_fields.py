"""Native-only caption/reference adapter for OfficeCLI-authored replay stages.

This module does not author chapter prose, images, or tables. It creates genuine
Word captions and references at OfficeCLI placeholders and validates native fields.
"""
import argparse
import json
from pathlib import Path
import re
import sys

from word_session import WordSession, run_bounded, update_directories
from finalize_word import caption_chemistry
import pythoncom


def find(doc, text):
    r = doc.Content.Duplicate
    r.Find.ClearFormatting()
    if not r.Find.Execute(FindText=text, Forward=True, Wrap=0, MatchWildcards=False):
        raise RuntimeError(f"Missing native placeholder: {text}")
    return r


def font(r):
    r.Font.NameFarEast = "標楷體"
    r.Font.NameAscii = r.Font.NameOther = "Times New Roman"
    r.Font.Size = r.Font.SizeBi = 12


def addins(doc):
    return [(f.Code.Text, f.Result.Text) for f in doc.Fields if f.Type == 81]


def worker(args):
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8-sig"))
    with WordSession(args.document, state_path=args.worker_state) as session:
        if session.attached:
            raise RuntimeError("Replay output unexpectedly open; inspect before continuing")
        doc = session.document
        assert not doc.TrackRevisions and doc.Revisions.Count == 0
        before = addins(doc)
        math_before = doc.OMaths.Count
        outside_caption_updates = 0
        # Inspect native SEQ fields once. Scanning every table-cell paragraph for
        # every caption caused thousands of unnecessary out-of-process COM calls.
        existing_caps = [f.Result.Paragraphs.Item(1).Range.Duplicate
                         for f in doc.Fields if f.Type == 12]
        for c in spec["captions"]:
            if c["in_chapter"]:
                continue
            prefix = f'{c["label"]} {c["number"]}{c["suffix"]}'
            matches = [p for p in existing_caps if p.Text.startswith(prefix + " ")]
            assert len(matches) == 1, (prefix, len(matches))
            cap = matches[0]
            if cap.Text.strip() != c["text"]:
                seq = [f for f in cap.Fields if f.Type == 12]
                assert len(seq) == 1
                doc.Range(seq[0].Result.End + 1 + len(c["suffix"]), cap.End - 1).Text = " " + c["title"]
                outside_caption_updates += 1
        labels = []
        captions = {}
        try:
            for label in ("Fig.", "Table"):
                item = session.word.CaptionLabels.Item(label)
                labels.append((item, bool(item.IncludeChapterNumber)))
                item.IncludeChapterNumber = False
            for block in spec["blocks"]:
                if block["type"] != "caption":
                    continue
                c = block["caption"]
                token = "@@CAP" + c["id"] + "@@"
                r = find(doc, token)
                start = r.Start
                doc.Range(start, start).Paragraphs.Item(1).Range.InsertCaption(
                    c["label"], " " + c["title"], pythoncom.Missing, 0, False
                )
                find(doc, token).Paragraphs.Item(1).Range.Delete()
                cap = find(doc, c["title"]).Paragraphs.Item(1).Range
                start = cap.Start
                cap.Style = doc.Styles("圖" if c["label"] == "Fig." else "表")
                font(cap)
                cap.Font.Bold = False
                cap.Font.Color = 0
                caption_chemistry(doc, cap)
                pf = cap.ParagraphFormat
                pf.Alignment = 1
                pf.LineSpacingRule = 2
                pf.KeepTogether = True
                pf.KeepWithNext = c["label"] == "Table"
                pf.SpaceBefore = pf.SpaceAfter = 4
                pf.FirstLineIndent = pf.LeftIndent = pf.RightIndent = 0
                seq = [f for f in cap.Fields if f.Type == 12]
                assert len(seq) == 1, c
                seq = seq[0]
                if c["repeat"]:
                    seq.Code.Text = f' SEQ {c["label"]} \\c \\* ARABIC '
                seq.Update()
                if c["suffix"]:
                    doc.Range(seq.Result.End + 1, seq.Result.End + 1).InsertBefore(c["suffix"])
                cap = doc.Range(start, start + 1).Paragraphs.Item(1).Range
                # Custom bookmarks target only a native-created label and number.
                mark = "Replay_" + c["id"]
                doc.Bookmarks.Add(mark, doc.Range(start, seq.Result.End + 1 + len(c["suffix"])))
                captions[c["id"]] = {**c, "bookmark": mark}
            for ref in spec["references"]:
                c = captions.get(ref["caption"])
                if c is None:
                    raise RuntimeError(f"Reference outside regenerated chapter: {ref}")
                r = find(doc, ref["token"])
                r.Text = ""
                if c["suffix"]:
                    doc.Fields.Add(r, 3, c["bookmark"] + " \\h", True)
                else:
                    entries = list(doc.GetCrossReferenceItems(c["label"]))
                    matches = [i + 1 for i, item in enumerate(entries) if c["title"] in item]
                    assert len(matches) == 1, (c, matches)
                    r.InsertCrossReference(c["label"], 3, str(matches[0]), True)
            update_directories(doc)
            assert before == addins(doc), "Native ADDIN fields changed"
            assert math_before == doc.OMaths.Count
            assert "@@REF" not in doc.Content.Text and "@@CAP" not in doc.Content.Text
            for c in captions.values():
                actual = doc.Bookmarks.Item(c["bookmark"]).Range.Text
                assert actual == f'{c["label"]} {c["number"]}{c["suffix"]}', (actual, c)
            broken = []
            for field in doc.Fields:
                if field.Type != 3:
                    continue
                m = re.search(r"\bREF\s+(\S+)", field.Code.Text)
                if m and not doc.Bookmarks.Exists(m[1]):
                    broken.append(m[1])
            assert all(x == "EqJ_322" for x in broken), broken
            doc.Save()
            report = {
                "document": str(Path(args.document).resolve()),
                "native_InsertCaption_calls": len(captions),
                "native_references_created": len(spec["references"]),
                "outside_caption_titles_updated": outside_caption_updates,
                "addin_fields_preserved": len(before),
                "equations_preserved": math_before,
                "pictures": doc.InlineShapes.Count,
                "tables": doc.Tables.Count,
                "pages": doc.ComputeStatistics(2),
                "existing_broken_references": broken,
                "contents": doc.TablesOfContents.Count,
                "figure_table_lists": doc.TablesOfFigures.Count,
            }
            Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            print(json.dumps(report, ensure_ascii=False), flush=True)
        finally:
            for item, original in labels:
                item.IncludeChapterNumber = original


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document")
    parser.add_argument("--spec", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--timeout", type=int, default=240)
    parser.add_argument("--worker-state")
    args = parser.parse_args()
    if args.worker_state:
        worker(args)
    else:
        raise SystemExit(run_bounded(__file__, sys.argv[1:], args.timeout))
