"""Resolve a supplied paper, deduplicate, and delegate import to native EndNote."""
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import sys
import uuid
from urllib.parse import unquote
import endnote_metadata as metadata
import endnote_integrations as integrations
import local_endnote as local


def _from_pdf(path):
    path = Path(path).resolve()
    if path.suffix.lower() != '.pdf' or not path.is_file():
        raise ValueError('pdf_path must identify an existing PDF.')
    sys.path.insert(0, str(integrations.TOOLS / 'python'))
    import pymupdf
    with pymupdf.open(path) as doc:
        text = doc[0].get_text() if len(doc) else ''
    matches = {metadata.lib.normalize_doi(m.rstrip('.,;)]}'))
               for m in re.findall(r'10\.\d{4,9}/[^\s<>"\x00-\x1f]+', text, re.I)}
    return sorted(matches), text


def resolve(args):
    doi = args.get('doi')
    title = args.get('title', '').strip()
    first_page = None
    if args.get('url'):
        url = unquote(args['url'])
        match = re.search(r'10\.\d{4,9}/[^?#\s]+', url, re.I)
        if not match and not doi and not title:
            raise ValueError('This link contains no DOI; also provide its article title or DOI.')
        if match:
            linked = metadata.lib.normalize_doi(match.group(0))
            if doi and metadata.lib.normalize_doi(doi).casefold() != linked.casefold():
                raise ValueError('DOI and URL identify different papers.')
            doi = linked
    if args.get('pdf_path') and not doi:
        candidates, first_page = _from_pdf(args['pdf_path'])
        if len(candidates) == 1:
            doi = candidates[0]
        elif not title:
            return {'ok': False, 'status': 'needs_disambiguation', 'doi_candidates': candidates,
                    'error': 'PDF first page has zero or multiple DOI candidates; provide its title or DOI.'}
    if doi:
        result = metadata.lookup_doi({'doi': doi})
        if not result['ok']:
            return result
        found = result['metadata']
        if title and metadata.norm(title) != metadata.norm(found['title']):
            return {'ok': False, 'status': 'needs_disambiguation', 'candidate': found,
                    'error': 'Supplied title and DOI metadata disagree; no import.'}
        if first_page is not None and metadata.norm(found['title']) not in metadata.norm(first_page):
            return {'ok': False, 'status': 'needs_disambiguation', 'candidate': found,
                    'error': 'Registry title was not observed on the PDF first page; no import.'}
    elif title:
        result = metadata.search_online({'query': title, 'limit': 10, 'match_mode': 'title',
                                         **{k: args[k] for k in ('author', 'year') if args.get(k)}})
        matches = [c for c in result['candidates'] if metadata.norm(c['title']) == metadata.norm(title)]
        if args.get('year'):
            matches = [c for c in matches if c['year'] == str(args['year'])]
        if args.get('author'):
            matches = [c for c in matches if any(metadata.norm(args['author']) in metadata.norm(a) for a in c['authors'])]
        unique = {c['doi'].casefold(): c for c in matches}
        if len(unique) != 1:
            return {'ok': False, 'status': 'needs_disambiguation', 'candidates': result['candidates'],
                    'error': 'No unique exact title match; add DOI, author, or year. No import.'}
        found = next(iter(unique.values()))
    else:
        raise ValueError('Provide title, doi, a DOI-containing URL, or pdf_path.')
    if not found['title'] or not found['authors'] or not re.fullmatch(r'\d{4}', found['year']):
        return {'ok': False, 'status': 'metadata_incomplete', 'candidate': found,
                'error': 'Title, author, and year are required; no import.'}
    if args.get('year') and found['year'] != str(args['year']):
        return {'ok': False, 'status': 'needs_disambiguation', 'candidate': found, 'error': 'Supplied year disagrees.'}
    if args.get('author') and not any(metadata.norm(args['author']) in metadata.norm(a) for a in found['authors']):
        return {'ok': False, 'status': 'needs_disambiguation', 'candidate': found, 'error': 'Supplied author disagrees.'}
    return {'ok': True, 'metadata': found}


def _existing(library_path, found):
    records = metadata._records({'library_path': library_path})
    matches = []
    for record in records:
        try:
            same_doi = bool(record['doi']) and metadata.lib.normalize_doi(record['doi']).casefold() == found['doi'].casefold()
        except ValueError:
            same_doi = False
        same_identity = (metadata.norm(record['title']) == metadata.norm(found['title']) and
                         record['year'] == found['year'] and found['authors'] and
                         metadata.norm(record['first_author_surname']) == metadata.norm(metadata.lib.author_surname(found['authors'][0])))
        if same_doi or same_identity:
            matches.append(record)
    return matches


def import_reference(args):
    resolved = resolve(args)
    if not resolved['ok']:
        return {**resolved, 'library_modified': False}
    found = resolved['metadata']
    library = Path(args['library_path']).resolve()
    existing = _existing(str(library), found)
    if existing:
        aligned = len(existing) == 1 and metadata._identity(existing[0], found)['verified']
        return {'ok': bool(aligned), 'status': 'already_present' if aligned else 'existing_record_needs_native_review' if len(existing) == 1 else 'duplicate_candidates',
                'references': existing, 'verified_metadata': found, 'library_modified': False,
                'note': 'Existing records were not reimported or overwritten; metadata differences need native review.'}
    types = {'journal-article': 'JOUR', 'book': 'BOOK', 'edited-book': 'EDBOOK',
             'book-chapter': 'CHAP', 'proceedings-article': 'CPAPER', 'proceedings': 'CONF',
             'posted-content': 'UNPB', 'report': 'RPRT', 'dissertation': 'THES',
             'JournalArticle': 'JOUR', 'Book': 'BOOK', 'BookChapter': 'CHAP', 'Dataset': 'DATA'}
    kind = types.get(found['publication_type'])
    if not kind:
        return {'ok': False, 'status': 'unsupported_reference_type', 'metadata': found, 'library_modified': False}
    missing = [field for field in ('journal', 'volume', 'pages') if kind == 'JOUR' and not found[field]]
    run = integrations.TOOLS / 'imports' / (datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ-') + uuid.uuid4().hex[:8])
    run.mkdir(parents=True)
    reference = {**found, 'type': kind}
    ris = run / 'verified-reference.ris'
    exported = local.library.export_ris(ris, [reference])
    if not exported['ok']:
        return exported
    (run / 'metadata-and-evidence.json').write_text(json.dumps(found, ensure_ascii=False, indent=2), encoding='utf-8')
    if args.get('preview_only', False):
        return {'ok': True, 'status': 'ready_for_native_import', 'metadata': found, 'ris_path': str(ris),
                'missing_publication_fields': missing, 'library_modified': False}
    # The existing bridge confirms the exact target library and uses EndNote Direct Export.
    bridge = local.upstream._invoke_bridge('import-ris', {'library_path': str(library), 'ris_path': str(ris)})
    after = _existing(str(library), found)
    success = len(after) == 1 and metadata.norm(after[0]['title']) == metadata.norm(found['title'])
    report = {'ok': success, 'status': 'imported' if success else 'import_not_confirmed',
              'references': after, 'verified_metadata': found, 'native_import': bridge,
              'ris_path': str(ris), 'missing_publication_fields': missing,
              'library_modified': True if success else 'unknown',
              'jtice_bibliography_validated': False,
              'next_step': 'Use diagnose_endnote_journal_abbreviations, then inspect actual CWYW output with the JTICE style. '
                           'Import success does not prove style or abbreviation correctness. Do not retry blindly after an unknown outcome.'}
    (run / 'result.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report
