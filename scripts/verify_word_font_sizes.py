"""Read-only font-size gate for generated Word directories and selected tables.

Audit the exact open document when available, including unsaved content. Call
after native directory updates. This tool neither changes nor saves the DOCX.
Exit 0 means sizes match; 1 means a formatting failure; 2 means an audit error.
"""

import argparse
import json
import os
from pathlib import Path
import sys

from word_session import WordSession, run_bounded


def check_font(report, location, font, text=''):
    size, size_bi = float(font.Size), float(font.SizeBi)
    expected = report['expected_size_pt']
    if abs(size - expected) > 0.01 or abs(size_bi - expected) > 0.01:
        report['issues'].append({
            'location': location,
            'size_pt': 'mixed' if size == 9999999 else size,
            'complex_script_size_pt': 'mixed' if size_bi == 9999999 else size_bi,
            'text': text[:180],
        })


def visible_range(rng):
    result = rng.Duplicate
    text = result.Text
    trailing = len(text) - len(text.rstrip('\r\x07'))
    if trailing:
        result.End -= trailing
    return result


def audit_directory(report, label, item):
    rng = visible_range(item.Range)
    report['checked_directory_entries'] += rng.Paragraphs.Count
    # A uniformly formatted directory needs only two native font reads. Inspect
    # paragraphs individually when Word reports a mix, to locate the failure.
    if (abs(float(rng.Font.Size) - report['expected_size_pt']) <= 0.01
            and abs(float(rng.Font.SizeBi) - report['expected_size_pt']) <= 0.01):
        return
    for index, paragraph in enumerate(rng.Paragraphs, 1):
        pr = visible_range(paragraph.Range)
        text = pr.Text.strip()
        if text:
            check_font(report, f'{label}: entry {index}', pr.Font, text)


def worker(args):
    report = {
        'document': str(Path(args.document).resolve()),
        'expected_size_pt': args.expected_size,
        'checked_directory_entries': 0,
        'checked_directory_styles': 0,
        'checked_tables': 0,
        'issues': [],
    }
    with WordSession(args.document, read_only=True,
                     state_path=args.worker_state) as session:
        doc = session.document
        report['source'] = 'open Word document' if session.attached else 'saved DOCX'
        report['saved_at_audit_start'] = bool(doc.Saved)
        if args.directories:
            total = doc.TablesOfContents.Count + doc.TablesOfFigures.Count
            if not total and not args.allow_no_directories:
                report['issues'].append({'location': 'directories',
                                         'error': 'No native directories found.'})
            if not total and args.allow_no_directories:
                report['directories'] = 'not applicable: no native directories present'
            style_ids = set()
            for index in range(1, doc.TablesOfContents.Count + 1):
                item = doc.TablesOfContents.Item(index)
                style_ids.update(-19 - level for level in
                                 range(item.UpperHeadingLevel, item.LowerHeadingLevel + 1))
                audit_directory(report, f'main contents {index}', item)
            for index in range(1, doc.TablesOfFigures.Count + 1):
                style_ids.add(-36)  # Built-in Table of Figures style, any UI language.
                audit_directory(report, f'figure/table/equation list {index}',
                                doc.TablesOfFigures.Item(index))
            for style_id in sorted(style_ids, reverse=True):
                style = doc.Styles.Item(style_id)
                report['checked_directory_styles'] += 1
                check_font(report, f'directory style: {style.NameLocal}', style.Font)

        if args.table_text:
            tables = [(index, table, table.Range.Text)
                      for index, table in enumerate(doc.Tables, 1)]
            for needle in args.table_text:
                matches = [(index, table, text) for index, table, text in tables
                           if needle in text]
                if len(matches) != 1:
                    report['issues'].append({
                        'location': f'table containing {needle!r}',
                        'error': f'Expected exactly one table; found {len(matches)}.',
                    })
                    continue
                index, table, text = matches[0]
                report['checked_tables'] += 1
                font = table.Range.Font
                if float(font.Size) == 9999999 or float(font.SizeBi) == 9999999:
                    # Word can report wdUndefined for an entire table even when
                    # every cell and row marker has the same size. Inspect the
                    # actual cells and row markers rather than accepting or
                    # rejecting that aggregate value as evidence.
                    for cell_index, cell in enumerate(table.Range.Cells, 1):
                        check_font(report, f'table {index}: cell {cell_index}',
                                   cell.Range.Font, cell.Range.Text)
                    for row in table.Rows:
                        marker = doc.Range(row.Range.End - 1, row.Range.End)
                        check_font(report, f'table {index}: row {row.Index} marker',
                                   marker.Font)
                else:
                    check_font(report, f'table {index} containing {needle!r}',
                               font, text.replace('\r\x07', ' | '))
        report['status'] = 'FAIL' if report['issues'] else 'PASS'
        if args.report:
            output = Path(args.report)
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(json.dumps(report, ensure_ascii=False, indent=2),
                              encoding='utf-8')
        print(json.dumps({key: value for key, value in report.items() if key != 'issues'},
                         ensure_ascii=False, indent=2), flush=True)
        print(f"Formatting failures: {len(report['issues'])}", flush=True)
        for issue in report['issues'][:8]:
            print(json.dumps(issue, ensure_ascii=False), flush=True)
        if len(report['issues']) > 8:
            print('Additional failures are in the JSON report.', flush=True)
        return 1 if report['issues'] else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('document')
    parser.add_argument('--directories', action='store_true')
    parser.add_argument('--allow-no-directories', action='store_true',
                        help='Report absent directories as not applicable; still audit all existing entries and styles.')
    parser.add_argument('--table-text', action='append', default=[])
    parser.add_argument('--expected-size', type=float, default=12)
    parser.add_argument('--report')
    parser.add_argument('--timeout', type=int, default=120)
    parser.add_argument('--worker-state')
    args = parser.parse_args()
    if not args.directories and not args.table_text:
        parser.error('Select --directories or at least one --table-text target.')
    if not 0 < args.expected_size < 1000:
        parser.error('--expected-size must be a positive point size below 1000.')
    if args.worker_state:
        try:
            return worker(args)
        except Exception as exc:
            print(f'Font-size audit failed: {exc}', file=sys.stderr, flush=True)
            return 2
    os.environ['PYTHONUTF8'] = '1'
    return run_bounded(Path(__file__), sys.argv[1:], args.timeout)


if __name__ == '__main__':
    sys.exit(main())
