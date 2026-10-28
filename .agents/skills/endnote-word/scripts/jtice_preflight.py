"""Read-only manuscript inventory before prose, layout, citation or COM changes."""
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile
import local_endnote as local
import journal_diagnostics

W = local.W
NS = local.NS


def inspect(args):
    path = Path(args['document_path']).resolve()
    stage = args.get('stage', 'draft')
    with ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        style_names = {}
        if 'word/styles.xml' in z.namelist():
            styles = ET.fromstring(z.read('word/styles.xml'))
            for s in styles.findall('w:style', NS):
                name = s.find('w:name', NS)
                if name is not None:
                    style_names[s.get(W + 'styleId')] = name.get(W + 'val')
        images = len([n for n in z.namelist() if n.startswith('word/media/')])
    issues, tables, captions = [], [], []
    revisions = len(root.findall('.//w:ins', NS)) + len(root.findall('.//w:del', NS))
    if revisions:
        issues.append({'kind': 'tracked_changes', 'count': revisions, 'action': 'Inspect before editing; preserve the source.'})
    fields = [''.join(e.text or '' for e in p.findall('.//w:instrText', NS)) for p in root.findall('.//w:p', NS)]
    for index, table in enumerate(root.findall('.//w:tbl', NS), 1):
        texts = [''.join(t.itertext()) for t in table.findall('.//w:t', NS)]
        small = []
        for run in table.findall('.//w:r', NS):
            text = ''.join(t.text or '' for t in run.findall('w:t', NS))
            if not text.strip():
                continue
            props = run.find('w:rPr', NS)
            if props is not None:
                sizes = {tag: props.find('w:' + tag, NS) for tag in ('sz', 'szCs')}
                if any(e is not None and e.get(W + 'val') != '24' for e in sizes.values()):
                    small.append(text[:60])
        tables.append({'table_index': index, 'unique_text_candidate': ''.join(texts)[:70],
                       'explicit_non_12pt_runs': small,
                       'inherited_and_actual_word_sizes': 'Run verify_word_font_sizes.py on the final file.'})
        if small:
            issues.append({'kind': 'table_direct_font_size', 'table_index': index, 'sample': small[:3]})
    for i, paragraph in enumerate(root.findall('.//w:body/w:p', NS), 1):
        text = ''.join(t.text or '' for t in paragraph.findall('.//w:t', NS)).strip()
        if re.match(r'^(Fig(?:ure)?\.?|Table|圖|表)\s*\d+', text, re.I):
            code = ''.join(t.text or '' for t in paragraph.findall('.//w:instrText', NS))
            style = paragraph.find('w:pPr/w:pStyle', NS)
            captions.append({'paragraph_index': i, 'text': text[:140], 'native_seq': bool(re.search(r'\bSEQ\b', code)),
                             'style': style_names.get(style.get(W + 'val'), style.get(W + 'val')) if style is not None else None})
    native = None
    try:
        native = local.inspect_document(path)
    except ValueError as exc:
        issues.append({'kind': 'native_field_inspection_blocked', 'detail': str(exc)})
    cited = sorted({r['rec_number'] for c in (native or {}).get('citations', []) for r in c['records']})
    source_inventory = None
    if args.get('source_document_path'):
        source = local.inspect_document(args['source_document_path'])
        source_records = sorted({r['rec_number'] for c in source['citations'] for r in c['records']})
        source_inventory = {'source_distinct_native_records': len(source_records),
                            'target_distinct_native_records': len(cited),
                            'source_only_record_numbers': sorted(set(source_records) - set(cited)),
                            'note': 'Require a retained/omitted claim-to-source map; do not automatically insert every source reference.'}
    abbreviation = None
    if args.get('library_path'):
        abbreviation = journal_diagnostics.diagnose({'library_path': args['library_path'],
                                                     'document_path': str(path)})
    return {'ok': True, 'document_path': str(path), 'stage': stage, 'document_modified': False,
            'inventory': {'native_citations': (native or {}).get('citation_count'),
                          'distinct_native_records': len(cited), 'tables': len(tables), 'media_files': images,
                          'editable_math_objects': len(root.findall('.//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath')),
                          'seq_paragraphs': sum(bool(re.search(r'\bSEQ\b', f)) for f in fields),
                          'ref_paragraphs': sum(bool(re.search(r'\b(?:REF|PAGEREF)\b', f)) for f in fields)},
            'issues': issues, 'tables': tables, 'caption_candidates': captions,
            'source_reference_inventory': source_inventory, 'abbreviation_diagnostics': abbreviation,
            'execution_policy': {'draft': 'Complete authorized content; do not impose submission length/display-item budgets.',
                                 'translation': 'Translate only authorized text; preserve scientific claims, figures, equations and fields.',
                                 'submission': 'Use the current official guide and validate actual fonts, directories, captions, bibliography and render.'}[stage],
            'checks_not_established': ['Source support for scientific claims', 'Current official guide verification',
                                       'Rendered layout', 'Actual inherited Word font sizes', 'Figure resolution at final size'],
            'next_steps': ['Resolve related metadata/term/style issues before citation formatting.',
                           'Use batch_edit_endnote_citations for multiple citation operations.',
                           'After all edits, perform one final directory update, actual font verification and rendered QA.']}
