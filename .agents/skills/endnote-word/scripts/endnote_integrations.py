"""Adapters for the installed EndNote CLI, citecheck and search MCP; no library writes."""
from __future__ import annotations

from contextlib import contextmanager
from dataclasses import asdict
from datetime import datetime, timezone
import json
import hashlib
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

import local_endnote as local
import endnote_metadata as metadata

WORKSPACE = Path(__file__).resolve().parents[4]
TOOLS = WORKSPACE / '.tools/endnote'
STATE = TOOLS / 'state'
SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS / 'endnote-cli-src/src'))


@contextmanager
def reader(path):
    """Reuse upstream readers on our existing read-only snapshot, never the live .enl."""
    from endnote_cli.core.reader import EndnoteLibrary
    p = Path(path).resolve()
    with local.library._sqlite_reader(local.library.resolve_sqlite_path(p)) as db:
        db.row_factory = sqlite3.Row
        upstream = EndnoteLibrary(p)
        upstream._conn = db
        try:
            yield upstream
        finally:
            upstream._conn = None


def _type_names():
    path = Path('C:/Program Files (x86)/EndNote 2025/XML Support/RefTypeTableEN9.xml')
    if path.is_file():
        return {int(e.get('id')): e.get('name') for e in ET.parse(path).getroot()}
    return {0: 'Journal Article', 1: 'Book', 2: 'Thesis', 7: 'Book Section', 33: 'Conference Paper'}


def _select(lib, args):
    refs = lib.list_refs()
    selected = {r['rec_number'] for r in metadata._records(args)} if (
        args.get('group') or args.get('group_id')) else {r.id for r in refs}
    if args.get('rec_numbers') is not None:
        wanted = set(args['rec_numbers'])
        if wanted - selected:
            raise ValueError('A requested record is missing, trashed, or outside the group.')
        selected &= wanted
    return [r for r in refs if r.id in selected]


def list_tags(args):
    with reader(args['library_path']) as lib:
        live = {r.id for r in lib.list_refs()}
        tags = [{**asdict(t), 'rec_numbers': sorted(set(lib.get_refs_by_tag(t.group_id)) & live)}
                for t in lib.list_tags()]
    return {'ok': True, 'tags': tags, 'library_modified': False}


def group_tree(args):
    with reader(args['library_path']) as lib:
        sets = [asdict(s) for s in lib.list_group_sets()]
        mapping = lib._get_group_set_mapping()
        groups = local._groups(lib.conn)
    for g in groups:
        g.pop('members', None)
        g['group_set_id'] = mapping.get(g['group_id'])
    return {'ok': True, 'group_sets': sets, 'groups': groups, 'library_modified': False,
            'provider': 'endnote-cli (hierarchy only); local validated manual membership codec'}


def get_full_record(args):
    with reader(args['library_path']) as lib:
        ref = lib.get_ref(args['rec_number'])
        if ref is None or ref.trash_state:
            raise ValueError('The active record does not exist.')
        result = asdict(ref)
        result['reference_type_name'] = _type_names().get(ref.reference_type, 'Unknown')
    return {'ok': True, 'reference': result, 'library_modified': False}


def export_catalog(args):
    """Reuse CLI exporters; correct the observed upstream type-name map in an adapter."""
    from endnote_cli.core import export, models
    names = _type_names()

    class TypedReference(models.Reference):
        @property
        def ref_type_name(self):
            return names.get(self.reference_type, 'Unknown')

    with reader(args['library_path']) as lib:
        refs = _select(lib, args)
        refs = [TypedReference(**{**asdict(r), 'attachments': r.attachments}) for r in refs]
    payload = export.refs_to_csv(refs) if args['format'] == 'csv' else export.refs_to_json(refs)
    metadata._write_new(args['output_path'], payload.encode('utf-8-sig' if args['format'] == 'csv' else 'utf-8'))
    return {'ok': True, 'output_path': str(Path(args['output_path']).resolve()), 'count': len(refs),
            'library_modified': False, 'provider': 'endnote-cli'}


def _bib_value(value):
    # Literal metadata only; never shell interpolation or template execution.
    return str(value or '').replace('\\', '\\\\').replace('{', r'\{').replace('}', r'\}').replace('\n', ' ').replace('\r', ' ')


def verify_online(args):
    if args.get('rec_number'):
        rec = metadata._record(args)
        records = [rec]
    else:
        records = metadata._records(args)
        if args.get('rec_numbers') is not None:
            wanted = set(args['rec_numbers'])
            if wanted - {r['rec_number'] for r in records}:
                raise ValueError('A requested record is missing, trashed, or outside the group.')
            records = [r for r in records if r['rec_number'] in wanted]
    limit = args.get('limit', 3)
    if not records:
        return {'ok': True, 'processed_count': 0, 'remaining_rec_numbers': [], 'library_modified': False}
    with tempfile.TemporaryDirectory(prefix='endnote-citecheck-', dir=WORKSPACE / '.tmp') as work:
        work = Path(work)
        lines = []
        for rec in records[:limit]:
            fields = {'title': rec['title'], 'author': ' and '.join(rec['authors']), 'year': rec['year'],
                      'journal': rec['journal'], 'volume': rec['volume'], 'number': rec['issue'],
                      'pages': rec['pages'], 'doi': rec['doi']}
            kind = {0: 'article', 28: 'article', 1: 'book', 2: 'phdthesis',
                    7: 'incollection', 33: 'inproceedings'}.get(rec['reference_type'], 'misc')
            lines.append('@' + kind + '{endnote' + str(rec['rec_number']) + ',\n' + '\n'.join(
                '  ' + k + ' = {' + _bib_value(v) + '},' for k, v in fields.items() if v) + '\n}')
        (work / 'references.bib').write_text('\n\n'.join(lines), encoding='utf-8')
        (work / 'arguments.json').write_text(json.dumps({'input_path': str(work / 'references.bib')}), encoding='utf-8')
        command = ['C:/Program Files/nodejs/node.exe', '--import',
                   (TOOLS / 'citecheck/node_modules/tsx/dist/loader.mjs').as_uri(),
                   str(SCRIPTS / 'citecheck_bridge.mjs'), str(work / 'arguments.json')]
        completed = subprocess.run(command, capture_output=True, encoding='utf-8', timeout=125,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if not completed.stdout.strip():
            raise RuntimeError('citecheck failed: ' + completed.stderr[-1500:])
        result = json.loads(completed.stdout)
        if completed.returncode:
            result['ok'] = False
        result.update(processed_count=min(len(records), limit),
                      rec_numbers=[r['rec_number'] for r in records[:limit]],
                      remaining_rec_numbers=[r['rec_number'] for r in records[limit:]],
                      library_modified=False,
                      note='Review only. correctedEnw/replacementBibtex are proposals; no EndNote record or CWYW field was edited.')
        if args.get('report_path'):
            metadata._write_json(args['report_path'], result)
        return result


def _signature(library_path):
    path = local.library.resolve_sqlite_path(Path(library_path))
    result = {}
    for p in [path, Path(str(path) + '-wal'), Path(str(path) + '-journal')]:
        if p.is_file():
            with p.open('rb') as stream:
                digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            result[str(p)] = {'size': p.stat().st_size, 'sha256': digest}
    return result


def create_search_snapshot(library_path):
    """Create a search-only XML snapshot, preserving full abstracts and nested PDF paths."""
    STATE.mkdir(parents=True, exist_ok=True)
    path = Path(library_path).resolve()
    before = _signature(path)
    root = ET.Element('xml')
    records = ET.SubElement(root, 'records')
    names = _type_names()
    with reader(path) as lib:
        refs = lib.list_refs()
        for ref in refs:
            record = ET.SubElement(records, 'record')
            ET.SubElement(record, 'rec-number').text = str(ref.id)
            # The search parser uses the explicit type name, never XML numeric type IDs.
            ET.SubElement(record, 'ref-type', name=names.get(ref.reference_type, 'Generic'))
            contributors = ET.SubElement(ET.SubElement(record, 'contributors'), 'authors')
            for author in ref.authors_list:
                ET.SubElement(contributors, 'author').text = author
            titles = ET.SubElement(record, 'titles')
            ET.SubElement(titles, 'title').text = ref.title
            ET.SubElement(titles, 'secondary-title').text = ref.journal
            ET.SubElement(ET.SubElement(record, 'dates'), 'year').text = ref.year
            keywords = ET.SubElement(record, 'keywords')
            for word in ref.keywords_list:
                ET.SubElement(keywords, 'keyword').text = word
            for tag, value in {'volume': ref.volume, 'number': ref.number, 'pages': ref.pages,
                               'abstract': ref.abstract, 'electronic-resource-num': ref.doi,
                               'publisher': ref.publisher, 'pub-location': ref.place_published,
                               'isbn': ref.isbn, 'notes': ref.notes, 'research-notes': ref.research_notes}.items():
                ET.SubElement(record, tag).text = value
            urls = ET.SubElement(record, 'urls')
            ET.SubElement(ET.SubElement(urls, 'related-urls'), 'url').text = ref.url
            attachment = lib.get_main_pdf(ref.id)
            if attachment:
                ET.SubElement(ET.SubElement(urls, 'pdf-urls'), 'url').text = 'internal-pdf://' + attachment.file_path.replace('\\', '/')
    if _signature(path) != before:
        raise ValueError('Library changed while snapshotting; retry after current edits finish.')
    xml_path = STATE / 'library-search-only.xml'
    temp = STATE / 'library-search-only.xml.tmp'
    ET.ElementTree(root).write(temp, encoding='utf-8', xml_declaration=True)
    os.replace(temp, xml_path)
    config = STATE / 'library.yaml'
    config.write_text(json.dumps({'endnote_xml': str(xml_path),
                      'pdf_dir': str(local.library.resolve_data_directory(path) / 'PDF'),
                      'db_path': str(STATE / 'search.db'), 'max_pdf_pages': 100}), encoding='utf-8')
    info = {'library_path': str(path), 'signature': before, 'record_count': len(refs),
            'created_at': datetime.now(timezone.utc).isoformat(), 'search_only': True,
            'indexing_complete': False}
    (STATE / 'snapshot.json').write_text(json.dumps(info, indent=2), encoding='utf-8')
    return info


def index_status(args):
    path = STATE / 'snapshot.json'
    if not path.exists():
        return {'ok': True, 'initialized': False}
    info = json.loads(path.read_text(encoding='utf-8'))
    info['stale'] = info['signature'] != _signature(info['library_path'])
    return {'ok': True, 'initialized': True, **info}


def refresh_index(args):
    command = [sys.executable, '-X', 'utf8', str(SCRIPTS / 'endnote_library_runner.py'),
               'refresh', '--library-path', args['library_path']]
    if args.get('semantic', True):
        command.append('--semantic')
    completed = subprocess.run(command, capture_output=True, encoding='utf-8', errors='replace', timeout=135,
                               creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    return {'ok': completed.returncode == 0, 'output': completed.stdout[-4000:],
            'stderr_tail': completed.stderr[-2000:], 'library_modified': False,
            'index': index_status({})}
