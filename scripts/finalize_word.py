"""Selective native directory refresh after OfficeCLI; never updates ADDIN fields."""
import argparse
import json
import sys
from pathlib import Path

from word_session import WordSession, run_bounded, update_directories


def caption_chemistry(doc, paragraph):
    """Apply native inline chemical indices without rewriting caption fields."""
    count = 0
    for formula, offset in (("H2S", 1), ("SO2", 2), ("CO2", 2), ("H2O", 1)):
        cursor = paragraph.Start
        while cursor < paragraph.End - 1:
            r = doc.Range(cursor, paragraph.End - 1)
            r.Find.ClearFormatting()
            if not r.Find.Execute(FindText=formula, Forward=True, Wrap=0, MatchWildcards=False):
                break
            digit = doc.Range(r.Start + offset, r.Start + offset + 1)
            digit.Font.Subscript = True
            assert digit.Font.Subscript
            count += 1
            cursor = r.End
    return count


def worker(args):
    with WordSession(args.document, state_path=args.worker_state) as session:
        doc = session.document
        if session.attached:
            raise RuntimeError("OfficeCLI finalization requires a closed staging file.")
        before = [(doc.Fields.Item(i).Code.Text, doc.Fields.Item(i).Result.Text)
                  for i in range(1, doc.Fields.Count + 1)
                  if doc.Fields.Item(i).Type == 81]
        caption_indices = 0
        if args.caption_spec:
            spec = json.loads(Path(args.caption_spec).read_text(encoding='utf-8-sig'))
            for cap in spec['captions']:
                if cap['in_chapter']:
                    mark = 'Replay_' + cap['id']
                    assert doc.Bookmarks.Exists(mark), mark
                    caption_indices += caption_chemistry(doc, doc.Bookmarks.Item(mark).Range.Paragraphs.Item(1).Range)
        update_directories(doc)
        after = [(doc.Fields.Item(i).Code.Text, doc.Fields.Item(i).Result.Text)
                 for i in range(1, doc.Fields.Count + 1)
                 if doc.Fields.Item(i).Type == 81]
        if before != after:
            raise RuntimeError("ADDIN field codes changed during directory update.")
        doc.Save()
        report = {"document": str(Path(args.document).resolve()),
                  "contents": doc.TablesOfContents.Count,
                  "figure_table_equation_lists": doc.TablesOfFigures.Count,
                  "addin_codes_preserved": True,
                  "native_caption_subscripts_checked": caption_indices}
        print(json.dumps(report, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document")
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--caption-spec", help="Optional replay recipe: format chemical indices only in its native-created captions")
    parser.add_argument("--worker-state", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker_state:
        worker(args)
        return 0
    return run_bounded(__file__, sys.argv[1:], args.timeout)


if __name__ == "__main__":
    raise SystemExit(main())
