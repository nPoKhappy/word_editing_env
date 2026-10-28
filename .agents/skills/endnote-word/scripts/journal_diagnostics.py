"""Separate library metadata, duplicate journal terms, and actual CWYW output."""
from collections import defaultdict
import re
import unicodedata
import local_endnote as local
import endnote_metadata as metadata


def key(text):
    return ' '.join(unicodedata.normalize('NFKC', text).split()).casefold()


def diagnose(args):
    records = metadata._records(args)
    terms = metadata._terms(args['library_path'])
    by_name = defaultdict(list)
    for term in terms:
        by_name[key(term['full_title'])].append(term)
    duplicates = []
    for name, rows in by_name.items():
        if len(rows) > 1:
            conflicts = {f: sorted({r[f].strip() for r in rows if r[f].strip()})
                         for f in ('abbreviation_1', 'abbreviation_2', 'abbreviation_3')}
            duplicates.append({'full_title': rows[0]['full_title'], 'count': len(rows), 'entries': rows,
                               'conflicting_columns': {k: v for k, v in conflicts.items() if len(v) > 1},
                               'blank_abbreviation_1_count': sum(not r['abbreviation_1'].strip() for r in rows)})
    entries, cited = {}, {}
    if args.get('document_path'):
        document = local.inspect_document(args['document_path'])
        for citation in document['citations']:
            numbers = re.findall(r'\[(\d+)\]', citation['result'])
            if len(numbers) == 1 and len(citation['records']) == 1:
                record = citation['records'][0]
                cited[record['rec_number']] = {'number': int(numbers[0]), 'title': record['title']}
        text = '\n'.join(b['result'] for b in document['bibliographies'])
        matches = list(re.finditer(r'\[(\d+)\]', text))
        entries = {int(m.group(1)): text[m.end():matches[i + 1].start() if i + 1 < len(matches) else len(text)]
                   for i, m in enumerate(matches)}
    findings = []
    for record in records:
        if args.get('document_path') and record['rec_number'] not in cited:
            continue
        if record['reference_type'] not in (0, 28):
            continue
        journal = record['journal']
        rows = by_name.get(key(journal), [])
        states = []
        if not journal:
            states.append('record_journal_missing')
        elif not rows:
            states.append('journal_term_missing')
        elif len(rows) > 1:
            states.append('duplicate_full_title_terms')
        if rows and not any(r['abbreviation_1'].strip() for r in rows):
            states.append('abbreviation_1_missing')
        choices = sorted({r['abbreviation_1'].strip() for r in rows if r['abbreviation_1'].strip()})
        if len(choices) > 1:
            states.append('abbreviation_1_conflict')
        rendered = None
        if record['rec_number'] in cited:
            identity = cited[record['rec_number']]
            if metadata.norm(identity['title']) != metadata.norm(record['title']):
                states.append('traveling_library_identity_mismatch')
            else:
                rendered = entries.get(identity['number'])
                if rendered is None:
                    states.append('bibliography_entry_not_mapped')
                elif args.get('expect_abbreviation', True) and len(choices) == 1:
                    if metadata.norm(journal) != metadata.norm(choices[0]) and metadata.norm(journal) in metadata.norm(rendered):
                        states.append('rendered_full_journal_title_review_style_or_terms')
                    elif metadata.norm(choices[0]) not in metadata.norm(rendered):
                        states.append('expected_abbreviation_not_observed')
                if rendered and record['volume'] and not re.search(r'(?<!\d)' + re.escape(record['volume']) + r'(?!\d)', rendered):
                    states.append('volume_stored_but_not_observed_review_style_template')
        findings.append({'rec_number': record['rec_number'], 'journal': journal,
                         'record_volume': record['volume'], 'record_pages_or_article_number': record['pages'],
                         'term_count': len(rows), 'abbreviation_1_candidates': choices,
                         'findings': states, 'rendered_entry': rendered})
    return {'ok': True, 'term_count': len(terms), 'unique_full_titles': len(by_name),
            'duplicate_titles': duplicates, 'records': findings, 'library_modified': False,
            'style_binary_inspected': False,
            'interpretation': 'A stored field missing in CWYW output needs native style/template review. '
                              'Duplicate names remain a diagnosis even if all abbreviations agree. '
                              'No term import, merge, or record update has been performed.'}
