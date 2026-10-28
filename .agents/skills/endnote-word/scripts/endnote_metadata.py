"""Evidence-based bibliographic checks and read-only EndNote maintenance.

No function writes an EndNote database or substitutes imports for existing-record edits.
Network results are data; reports explicitly distinguish candidates from verified identity.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timezone
from difflib import SequenceMatcher
import hashlib
import html
import json
import os
from pathlib import Path
import re
import tempfile
import time
import unicodedata
from urllib.error import HTTPError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import local_endnote as local

lib = local.library
MAX_RESPONSE = 4 * 1024 * 1024
FIELDS = ('title', 'authors', 'year', 'journal', 'volume', 'issue', 'pages', 'doi', 'url')


def norm(value):
    return ''.join(c for c in unicodedata.normalize('NFKD', str(value)).casefold()
                   if c.isalnum() and not unicodedata.combining(c))


def clean(value):
    return html.unescape(re.sub(r'<[^>]+>', '', str(value or ''))).strip()


def _json_get(url):
    """Bounded public registry GET; errors never mean a paper does not exist."""
    request = Request(url, headers={'User-Agent': 'EndNoteWordLocal/1.3 (bibliographic verification)',
                                    'Accept': 'application/json'})
    for attempt in range(2):
        try:
            with urlopen(request, timeout=12) as response:
                if response.geturl().split(':', 1)[0] != 'https':
                    raise ValueError('Registry redirected to a non-HTTPS URL.')
                data = response.read(MAX_RESPONSE + 1)
                if len(data) > MAX_RESPONSE:
                    raise ValueError('Registry response exceeds 4 MiB.')
                return json.loads(data)
        except HTTPError as exc:
            if attempt == 0 and exc.code in (429, 503):
                time.sleep(1)
                continue
            raise


def _first(values):
    return values[0] if isinstance(values, list) and values else ''


def _crossref(item, source):
    date = next((item[k].get('date-parts', []) for k in
                 ('published-print', 'published', 'issued', 'published-online')
                 if item.get(k, {}).get('date-parts')), [])
    year = str(date[0][0]) if date and date[0] else ''
    authors = [', '.join(x for x in (a.get('family'), a.get('given')) if x)
               or a.get('name', '') for a in item.get('author', [])]
    return {'title': clean(_first(item.get('title'))), 'authors': authors, 'year': year,
            'journal': clean(_first(item.get('container-title'))),
            'volume': str(item.get('volume', '')), 'issue': str(item.get('issue', '')),
            'pages': str(item.get('page') or item.get('article-number') or ''),
            'doi': lib.normalize_doi(item.get('DOI', '')), 'url': item.get('URL', ''),
            'publication_type': item.get('type', ''), 'provider': 'Crossref',
            'source_url': source, 'publisher': item.get('publisher', ''),
            'journal_abbreviations': item.get('short-container-title', []),
            'issn': item.get('ISSN', []), 'updates': item.get('update-to', []),
            'retrieved_at': datetime.now(timezone.utc).isoformat()}


def _datacite(item, source):
    a = item['attributes']
    authors = [', '.join(x for x in (v.get('familyName'), v.get('givenName')) if x)
               or v.get('name', '') for v in a.get('creators', [])]
    container = a.get('container') or {}
    return {'title': clean(next((t.get('title', '') for t in a.get('titles', [])
                                if not t.get('titleType')), '')),
            'authors': authors, 'year': str(a.get('publicationYear') or ''),
            'journal': container.get('title') or '', 'volume': container.get('volume') or '',
            'issue': container.get('issue') or '',
            'pages': '-'.join(str(container[k]) for k in ('firstPage', 'lastPage')
                              if container.get(k)),
            'doi': lib.normalize_doi(a.get('doi', '')), 'url': a.get('url') or '',
            'publication_type': a.get('types', {}).get('resourceTypeGeneral', ''),
            'provider': 'DataCite', 'source_url': source,
            'retrieved_at': datetime.now(timezone.utc).isoformat()}


def lookup_doi(args):
    doi = lib.normalize_doi(args['doi'])
    if not doi:
        raise ValueError('A nonempty valid DOI is required.')
    errors = []
    for provider, url, convert in (
        ('Crossref', 'https://api.crossref.org/works/' + quote(doi, safe=''), _crossref),
        ('DataCite', 'https://api.datacite.org/dois/' + quote(doi, safe=''), _datacite),
    ):
        try:
            payload = _json_get(url)
            metadata = convert(payload['message'] if provider == 'Crossref' else payload['data'], url)
            if metadata['doi'].casefold() != doi.casefold():
                raise ValueError('Registry returned a different DOI.')
            return {'ok': True, 'metadata': metadata, 'provider_errors': errors,
                    'identity_verified_against_library': False}
        except Exception as exc:
            errors.append({'provider': provider, 'error': f'{type(exc).__name__}: {exc}'})
    return {'ok': False, 'doi': doi, 'provider_errors': errors,
            'error': 'No registry metadata retrieved; this does not establish nonexistence.'}


def search_online(args):
    query = args['query'].strip()
    if not query or len(query) > 1000:
        raise ValueError('query must contain 1-1000 characters.')
    limit = args.get('limit', 5)
    field = 'query.title' if args.get('match_mode') == 'title' else 'query.bibliographic'
    params = {field: query, 'rows': limit}
    if args.get('author'):
        params['query.author'] = args['author']
    if args.get('year'):
        params['filter'] = f'from-pub-date:{args["year"]}-01-01,until-pub-date:{args["year"]}-12-31'
    url = 'https://api.crossref.org/works?' + urlencode(params)
    payload = _json_get(url)
    records = [_crossref(item, 'https://api.crossref.org/works/' + quote(item['DOI'], safe=''))
               for item in payload['message']['items']]
    return {'ok': True, 'query': query, 'source_url': url, 'candidates': records,
            'count': len(records), 'identity_verified_against_library': False}


def _record(args):
    result = lib.get_reference(args['library_path'], args['rec_number'])
    if not result['ok']:
        raise ValueError(result['error'])
    rec = result['reference']
    if args.get('expected_title') is not None and norm(rec['title']) != norm(args['expected_title']):
        raise ValueError('Record/title identity changed; re-query before proceeding.')
    return rec


def _identity(record, candidate):
    title_score = SequenceMatcher(None, norm(record['title']), norm(candidate['title'])).ratio()
    authors = candidate.get('authors', [])
    surname = lib.author_surname(authors[0]) if authors else ''
    author_match = bool(surname and norm(record['first_author_surname']) == norm(surname))
    doi_match = (bool(record['doi']) and
                 lib.normalize_doi(record['doi']).casefold() == candidate['doi'].casefold())
    year_match = record['year'] == candidate['year']
    # DOI alone is insufficient: a wrong DOI must not overwrite an unrelated reference.
    verified = title_score >= 0.90 and author_match and (doi_match or year_match)
    return {'verified': verified, 'title_similarity': round(title_score, 4),
            'first_author_match': author_match, 'year_match': year_match,
            'doi_match': doi_match}


def verify_record(args):
    record = _record(args)
    try:
        if record['doi']:
            result = lookup_doi({'doi': record['doi']})
            candidates = [result['metadata']] if result['ok'] else []
        else:
            result = search_online({'query': ' '.join((record['title'],
                                    record['first_author_surname'], record['year'])), 'limit': 5})
            candidates = result['candidates']
    except Exception as exc:
        result = {'ok': False, 'error': f'{type(exc).__name__}: {exc}'}
        candidates = []
    matches = [{'metadata': c, 'identity': _identity(record, c)} for c in candidates]
    verified = [m for m in matches if m['identity']['verified']]
    unique = {m['metadata']['doi'].casefold(): m for m in verified}
    selected = next(iter(unique.values())) if len(unique) == 1 else None
    changes = []
    if selected:
        c = selected['metadata']
        for field in FIELDS:
            before, after = record.get(field, ''), c.get(field, '')
            if after and before != after:
                changes.append({'field': field, 'before': before, 'after': after,
                                'source_url': c['source_url'],
                                'kind': 'fill_empty' if not before else 'replace_existing',
                                'review_required': bool(before)})
    report = {'ok': bool(result['ok']), 'library_path': str(Path(args['library_path']).resolve()),
              'rec_number': record['rec_number'], 'record': record,
              'status': 'verified' if selected else 'ambiguous' if len(unique) > 1 else
                        'identity_unverified' if candidates else 'lookup_failed',
              'candidates': matches, 'changes': changes, 'library_modified': False,
              'native_writeback_available': False,
              'writeback_limitation': 'This installation has no verified RSServices RDK or native record-edit adapter. '
                                     'This report is a preview; no record has been written.',
              'retrieved_at': datetime.now(timezone.utc).isoformat()}
    if not result['ok']:
        report['lookup_error'] = result
    if args.get('report_path'):
        _write_json(args['report_path'], report)
    return report


def _records(args):
    path = Path(args['library_path']).resolve()
    if path.suffix.lower() == '.xml':
        if args.get('group') or args.get('group_id'):
            raise ValueError('XML exports do not retain manual groups.')
        return [lib._public_record(r) for r in lib._read_xml(path) if r['trash_state'] == 0]
    if args.get('group') or args.get('group_id'):
        # Resolve the exact group and all members, without the search tool's 200-item cap.
        with lib._sqlite_reader(lib.resolve_sqlite_path(path)) as db:
            groups = local._groups(db)
            matches = [g for g in groups if g['group_id'] == args['group_id']] if 'group_id' in args else [
                g for g in groups if g['name'].casefold() == args['group'].casefold()]
            if len(matches) != 1 or not matches[0]['manual']:
                raise ValueError('Select one unambiguous manual group.')
            return [lib._public_record(r) for n in matches[0]['members']
                    if (r := lib._fetch_sqlite(db, n, False))]
    with lib._sqlite_reader(lib.resolve_sqlite_path(path)) as db:
        return [lib._public_record(lib._row_to_raw(row)) for row in
                db.execute('SELECT ' + ','.join(lib.SELECT_COLUMNS) +
                           ' FROM refs WHERE trash_state=0 ORDER BY id')]


def audit_library(args):
    records = _records(args)
    issues = []
    for r in records:
        missing = [f for f in ('title', 'authors', 'year') if not r[f]]
        if missing:
            issues.append({'rec_number': r['rec_number'], 'kind': 'missing_identity', 'fields': missing})
        if r['reference_type'] == 0 and not r['journal']:
            issues.append({'rec_number': r['rec_number'], 'kind': 'missing_journal'})
        if r['reference_type'] == 0 and not r['doi']:
            issues.append({'rec_number': r['rec_number'], 'kind': 'doi_not_recorded',
                           'note': 'May legitimately have no DOI; online verification is needed.'})
        if r['doi']:
            try:
                lib.normalize_doi(r['doi'])
            except ValueError:
                issues.append({'rec_number': r['rec_number'], 'kind': 'invalid_doi_syntax'})
        if r['year'] and not re.fullmatch(r'\d{4}[a-z]?', r['year'], re.I):
            issues.append({'rec_number': r['rec_number'], 'kind': 'review_year', 'value': r['year']})
    result = {'ok': True, 'record_count': len(records), 'issue_count': len(issues),
              'issues': issues, 'online_verified': False, 'library_modified': False}
    if args.get('report_path'):
        _write_json(args['report_path'], result)
    return result


def duplicates(args):
    records = _records(args)
    by_doi, by_title = defaultdict(list), defaultdict(list)
    for r in records:
        if r['doi']:
            try:
                by_doi[lib.normalize_doi(r['doi']).casefold()].append(r['rec_number'])
            except ValueError:
                pass
        if norm(r['title']):
            by_title[(norm(r['title']), norm(r['first_author_surname']), r['year'])].append(r['rec_number'])
    groups = [{'basis': basis, 'key': str(key), 'rec_numbers': numbers}
              for basis, mapping in (('doi', by_doi), ('title_author_year', by_title))
              for key, numbers in mapping.items() if len(numbers) > 1]
    return {'ok': True, 'record_count': len(records), 'duplicate_candidates': groups,
            'library_modified': False, 'note': 'Candidates only; no merge or deletion was performed.'}


def _terms(path):
    with lib._sqlite_reader(lib.resolve_sqlite_path(path)) as db:
        columns = {r[1] for r in db.execute('PRAGMA table_info(jterms)')}
        if not {'term', 'abbr1', 'abbr2', 'abbr3', 'list_id'} <= columns:
            raise ValueError('Unsupported journal-term schema; refusing to guess.')
        return [{'list_id': row[0], 'full_title': row[1], 'abbreviation_1': row[2],
                 'abbreviation_2': row[3], 'abbreviation_3': row[4]}
                for row in db.execute('SELECT list_id,term,abbr1,abbr2,abbr3 FROM jterms ORDER BY term')]


def journal_terms(args):
    terms = _terms(args['library_path'])
    if args.get('query'):
        terms = [t for t in terms if args['query'].casefold() in json.dumps(t, ensure_ascii=False).casefold()]
    return {'ok': True, 'terms': terms, 'count': len(terms), 'library_modified': False}


def prepare_terms(args):
    terms = _terms(args['library_path'])
    grouped = defaultdict(list)
    for t in terms:
        grouped[(t['list_id'], t['full_title'].strip().casefold())].append(t)
    merged, conflicts = [], []
    for key, values in grouped.items():
        row = {'list_id': key[0], 'full_title': values[0]['full_title']}
        for f in ('abbreviation_1', 'abbreviation_2', 'abbreviation_3'):
            options = sorted({v[f].strip() for v in values if v[f].strip()})
            row[f] = options[0] if len(options) == 1 else ''
            if len(options) > 1:
                conflicts.append({'full_title': row['full_title'], 'field': f, 'values': options})
        merged.append(row)
    if len({r['list_id'] for r in terms}) > 1:
        raise ValueError('Multiple journal lists require separate exports; refusing to combine them.')
    output = None
    if args.get('output_path') and not conflicts:
        output = str(Path(args['output_path']).resolve())
        lines = ['\t'.join(r[k] for k in ('full_title', 'abbreviation_1', 'abbreviation_2', 'abbreviation_3'))
                 for r in merged]
        if any('\n' in v or '\r' in v or '\t' in v for r in merged for k, v in r.items() if k != 'list_id'):
            raise ValueError('Term contains a separator; no export published.')
        _write_new(output, ('\r\n'.join(lines) + '\r\n').encode('utf-8'))
    return {'ok': True, 'original_count': len(terms), 'unique_count': len(merged),
            'merged_terms': merged, 'conflicts': conflicts, 'output_path': output,
            'ready_for_native_import': not conflicts, 'library_modified': False,
            'note': 'Prepared a unique-name term file; native EndNote import and CWYW output still require verification.'}


def attachments(args):
    path = Path(args['library_path']).resolve()
    data = lib.resolve_data_directory(path).resolve()
    with lib._sqlite_reader(lib.resolve_sqlite_path(path)) as db:
        cols = {r[1] for r in db.execute('PRAGMA table_info(file_res)')}
        if not {'refs_id', 'file_path', 'file_type'} <= cols:
            raise ValueError('Unsupported attachment schema.')
        rows = list(db.execute('SELECT f.refs_id,f.file_path,f.file_type FROM file_res f '
                               'JOIN refs r ON r.id=f.refs_id WHERE r.trash_state=0'))
    result = []
    for rec, stored, kind in rows:
        stored = stored or ''
        if stored.startswith('internal-pdf://'):
            stored = stored[len('internal-pdf://'):]
        p = Path(stored)
        relative = not p.is_absolute()
        candidate = (data / 'PDF' / p).resolve() if relative else p.resolve()
        unsafe = relative and not candidate.is_relative_to(data / 'PDF')
        result.append({'rec_number': rec, 'stored_path': stored, 'file_type': kind,
                       'resolved_path': str(candidate), 'exists': False if unsafe or not stored else candidate.is_file(),
                       'unsafe_relative_path': unsafe})
    return {'ok': True, 'attachment_count': len(result), 'missing_count': sum(not r['exists'] for r in result),
            'attachments': result, 'library_modified': False}


def _write_new(path, payload):
    path = Path(path).resolve()
    # Exclusive creation prevents accidental overwrites of reports and source files.
    with path.open('xb') as f:
        f.write(payload)


def _write_json(path, data):
    _write_new(path, json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8'))


def export_records(args):
    records = _records(args)
    wanted = args.get('rec_numbers')
    if wanted is not None:
        present = {r['rec_number'] for r in records}
        if set(wanted) - present:
            raise ValueError('Requested record is missing, trashed, or outside the selected group.')
        records = [r for r in records if r['rec_number'] in wanted]
    if args.get('format', 'json') == 'ris':
        if Path(args['output_path']).exists():
            raise ValueError('Export destination already exists.')
        # EndNote numeric type IDs must not silently become journal articles on export.
        if Path(args['library_path']).suffix.lower() == '.xml':
            raise ValueError('RIS export requires .enl: XML export type numbers differ from native SQLite type IDs.')
        # IDs checked against the installed EndNote RefTypeTableEN9.xml.
        codes = {0: 'JOUR', 1: 'BOOK', 2: 'THES', 3: 'CONF', 7: 'CHAP', 9: 'EDBOOK',
                 10: 'RPRT', 16: 'ELEC', 28: 'EJOUR', 29: 'BOOK', 31: 'GEN', 33: 'CPAPER'}
        if any(r['reference_type'] not in codes for r in records):
            raise ValueError('A reference type lacks a verified RIS mapping; use JSON to preserve its type ID.')
        payload = [{**r, 'type': codes[r['reference_type']]} for r in records]
        return lib.export_ris(args['output_path'], payload)
    _write_json(args['output_path'], {'records': records})
    return {'ok': True, 'output_path': str(Path(args['output_path']).resolve()),
            'count': len(records), 'library_modified': False}


def capabilities(args):
    return {'ok': True, 'version': '1.5.0',
            'available': ['Crossref title search', 'Crossref/DataCite DOI metadata',
                          'identity-checked correction previews', 'library/group completeness audit',
                          'duplicate candidates', 'journal terms read/merge/export preview',
                          'attachment path audit', 'selected record JSON/RIS export',
                          'closed-library full backup with hashes', 'document/library citation identity audit',
                          'existing native CWYW tools', 'citecheck multi-index correction previews',
                          'endnote-mcp PDF full-text and CPU semantic search',
                          'endnote-cli full metadata, tags, hierarchy and CSV exports',
                          'verified title/DOI/link/PDF DOI native import with duplicate prevention',
                          'transactional batch native citation editing with one CWYW format',
                          'JTICE read-only preflight and journal-term/style diagnostics',
                          'hash-verified backup restoration into a new directory'],
            'third_party_packages': {'endnote-mcp': '1.4.13', '@jhlee0619/citecheck': '0.1.0',
                                     'endnote-cli': '7faca0f95c9aa89b683f61707148a2967117d608'},
            'unavailable': {'existing_record_writeback': 'No verified native editor/RSServices RDK installed.',
                            'journal_term_writeback': 'No verified native term-list editor installed.',
                            'smart_group_evaluation': 'Rules are not decoded.',
                            'record_merge_delete': 'Not exposed; references and traveling-library identities must be preserved.',
                            'sync_sharing_fulltext': 'No authenticated native API implemented.'},
            'native_api_documentation': 'https://docs.endnote.com/docs/endnote/2025/v1/windows/en/content/appendices/apis_and_plug-ins.htm',
            'reload_required_for_new_mcp_tools': True}


def verify_batch(args):
    """Keep a batch below the MCP deadline and report failures for each record."""
    records = _records(args)
    wanted = args.get('rec_numbers')
    if wanted is not None:
        if set(wanted) - {r['rec_number'] for r in records}:
            raise ValueError('A requested record is missing, trashed, or outside the group.')
        records = [r for r in records if r['rec_number'] in wanted]
    limit = args.get('limit', 3)
    results = []
    for record in records[:limit]:
        try:
            report = verify_record({'library_path': args['library_path'], 'rec_number': record['rec_number'],
                                    'expected_title': record['title']})
        except Exception as exc:
            report = {'ok': False, 'rec_number': record['rec_number'], 'status': 'lookup_failed',
                      'error': f'{type(exc).__name__}: {exc}'}
        results.append(report)
    report = {'ok': all(r['ok'] for r in results), 'reports': results, 'processed_count': len(results),
              'remaining_rec_numbers': [r['rec_number'] for r in records[limit:]],
              'library_modified': False}
    if args.get('report_path'):
        _write_json(args['report_path'], report)
    return report


def audit_document(args):
    stored = local.inspect_document(args['document_path'])
    records = {r['rec_number']: r for r in _records(args)}
    issues = []
    for ci, citation in enumerate(stored['citations'], 1):
        for cited in citation['records']:
            actual = records.get(cited['rec_number'])
            reason = 'missing_record' if actual is None else 'identity_mismatch' if (
                norm(actual['title']) != norm(cited['title']) or
                norm(actual['first_author_surname']) != norm(cited['author']) or
                actual['year'] != cited['year']) else None
            if reason:
                issues.append({'citation_index': ci, 'rec_number': cited['rec_number'],
                               'kind': reason, 'traveling_record': cited, 'library_record': actual})
    if stored['citation_count'] and not any(b['result'].strip() for b in stored['bibliographies']):
        issues.append({'kind': 'missing_or_empty_bibliography'})
    return {'ok': True, 'audit_passed': not issues, 'citation_count': stored['citation_count'],
            'issues': issues, 'document_modified': False,
            'note': 'Checks stored record identities; database IDs, full traveling metadata and layout need separate review.'}


def backup_library(args):
    """Copy the full library only while a Windows deny-write handle can be held."""
    import win32file
    import win32con
    source = Path(args['library_path']).resolve()
    if source.suffix.lower() != '.enl' or not source.is_file():
        raise ValueError('An existing .enl library is required.')
    data = lib.resolve_data_directory(source).resolve()
    output = Path(args['output_directory']).resolve()
    if output.exists() or output.is_relative_to(data) or output == source.parent:
        raise ValueError('Choose a new backup directory outside the library data directory.')
    # Existing writer means sharing violation before any backup is published.
    handle = win32file.CreateFile(str(source), win32con.GENERIC_READ, win32con.FILE_SHARE_READ,
                                 None, win32con.OPEN_EXISTING, 0, None)
    try:
        with tempfile.TemporaryDirectory(prefix='endnote-backup-', dir=output.parent) as work:
            scratch = Path(work) / 'backup'
            scratch.mkdir()
            directories, files = [], []
            for folder, dirs, names in os.walk(data, followlinks=False):
                here = Path(folder)
                if any((here / n).is_symlink() or (here / n).is_junction() for n in dirs + names):
                    raise ValueError('Linked library paths are unsupported by backup.')
                directories.append(here)
                files.extend(here / n for n in names)
            for directory in directories:
                (scratch / data.name / directory.relative_to(data)).mkdir(parents=True, exist_ok=True)
            paths = [source] + sorted(files)
            hashes = {}
            for p in paths:
                if p.is_symlink() or p.is_junction():
                    raise ValueError('Linked library files are unsupported by backup.')
                rel = Path(source.name) if p == source else Path(data.name) / p.relative_to(data)
                payload = p.read_bytes()
                hashes[str(rel)] = hashlib.sha256(payload).hexdigest()
                target = scratch / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(payload)
            current = [source] + sorted(p for p in data.rglob('*') if p.is_file())
            if paths != current or any(hashlib.sha256(p.read_bytes()).hexdigest() != hashes[
                    str(Path(source.name) if p == source else Path(data.name) / p.relative_to(data))] for p in paths):
                raise ValueError('Library changed during backup; no backup published.')
            _write_json(scratch / 'backup-manifest.json', {'source': str(source), 'sha256': hashes,
                        'created_at': datetime.now(timezone.utc).isoformat()})
            scratch.rename(output)
    finally:
        handle.Close()
    return {'ok': True, 'output_directory': str(output), 'file_count': len(hashes),
            'manifest_path': str(output / 'backup-manifest.json'), 'library_modified': False}


def restore_backup(args):
    """Restore a hash-verified backup into a new directory, never over a live library."""
    backup = Path(args['backup_directory']).resolve()
    output = Path(args['output_directory']).resolve()
    if output.exists() or output.is_relative_to(backup):
        raise ValueError('Restore requires a new directory outside the backup.')
    manifest = json.loads((backup / 'backup-manifest.json').read_text(encoding='utf-8'))
    files = manifest['sha256']
    if not isinstance(files, dict) or not files:
        raise ValueError('Backup manifest must list files.')
    sources = {}
    for relative, digest in files.items():
        name = Path(relative)
        source = (backup / name).resolve()
        if name.is_absolute() or '..' in name.parts or not source.is_relative_to(backup) or not source.is_file():
            raise ValueError('Invalid backup member path.')
        if hashlib.sha256(source.read_bytes()).hexdigest() != digest:
            raise ValueError('Backup hash mismatch: ' + relative)
        sources[name] = source
    libraries = [p for p in sources if len(p.parts) == 1 and p.suffix.lower() == '.enl']
    if len(libraries) != 1:
        raise ValueError('Backup must contain exactly one root .enl file.')
    library = libraries[0]
    data_name = library.stem + '.Data'
    if Path(data_name) / 'sdb/sdb.eni' not in sources:
        raise ValueError('Backup lacks its paired data store.')
    if any(p != library and p.parts[0] != data_name for p in sources):
        raise ValueError('Backup contains files outside the paired library.')
    with tempfile.TemporaryDirectory(prefix='endnote-restore-', dir=output.parent) as work:
        scratch = Path(work) / 'restored'
        scratch.mkdir()
        for relative, source in sources.items():
            target = scratch / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            payload = source.read_bytes()
            if hashlib.sha256(payload).hexdigest() != files[str(relative)]:
                raise ValueError('Backup changed during restoration.')
            target.write_bytes(payload)
        # Retain legitimate empty data folders, without following links.
        for folder, dirs, names in os.walk(backup / data_name, followlinks=False):
            here = Path(folder)
            if any((here / n).is_symlink() or (here / n).is_junction() for n in dirs):
                raise ValueError('Linked backup folders are unsupported.')
            (scratch / data_name / here.relative_to(backup / data_name)).mkdir(parents=True, exist_ok=True)
        scratch.rename(output)
    return {'ok': True, 'library_path': str(output / library), 'file_count': len(sources),
            'original_library_modified': False, 'opened_in_endnote': False}
