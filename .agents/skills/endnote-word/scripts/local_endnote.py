"""Read-only manual group access and surgical Word operations using the vendored MCP."""
from __future__ import annotations
import hashlib
from collections import Counter
import json
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
VENDOR = ROOT / 'vendor/endnote-for-word/scripts'
sys.path.insert(0, str(VENDOR))
import endnote_library as library
import mcp_server as upstream

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'
_owned_word_pid = None
_parse_pid = upstream._extract_word_pid


def remember_pid(stderr):
    global _owned_word_pid
    _owned_word_pid = _parse_pid(stderr)
    return _owned_word_pid


def safe_reap(before):
    """A time difference does not prove ownership. Kill only an emitted Word PID."""
    observed = {name: sorted(upstream._process_pids(name) - before.get(name, set()))
                for name in upstream.WATCHED_PROCESS_NAMES}
    observed['terminated'] = []
    if _owned_word_pid in observed.get('WINWORD.EXE', []):
        if upstream._terminate_process_pid(_owned_word_pid):
            observed['terminated'].append(_owned_word_pid)
    return observed


upstream._extract_word_pid = remember_pid
upstream._reap_orphans = safe_reap


def decode_members(data):
    """Observed EndNote 2025 manual-group format: version BE, count and IDs LE."""
    if not isinstance(data, bytes) or len(data) < 8 or data[:4] != b'\0\0\0\2':
        raise ValueError('Unsupported EndNote group-member encoding; refusing to guess.')
    count = struct.unpack_from('<I', data, 4)[0]
    if count > 100000 or len(data) != 8 + 4 * count:
        raise ValueError('Invalid group-member count/length.')
    ids = list(struct.unpack_from('<' + 'I' * count, data, 8)) if count else []
    if any(n <= 0 for n in ids) or len(ids) != len(set(ids)):
        raise ValueError('Invalid or duplicate group record numbers.')
    return ids


def _groups(db):
    result = []
    for gid, spec, members in db.execute('SELECT group_id,spec,members FROM groups ORDER BY group_id'):
        if not isinstance(spec, (bytes, str)):
            raise ValueError('Unsupported group specification.')
        if b'<!DOCTYPE' in (spec.encode() if isinstance(spec, str) else spec).upper():
            raise ValueError('Group specification contains a document type declaration.')
        root = ET.fromstring(spec)
        rules = [e.text or '' for e in root.findall('./rules/rule')]
        manual = 'TYPE;3' in rules
        item = {'group_id': gid, 'name': root.findtext('./ids/name', ''),
                'manual': manual, 'search_supported': manual}
        if manual:
            item['members'] = decode_members(members)
            item['member_count'] = len(item['members'])
        result.append(item)
    return result


def group_operation(args, search=False):
    path = Path(args['library_path']).resolve()
    if path.suffix.lower() != '.enl':
        raise ValueError('Group search requires an .enl library; XML exports do not retain groups.')
    with library._sqlite_reader(library.resolve_sqlite_path(path)) as db:
        groups = _groups(db)
        if not search:
            return {'ok': True, 'library_path': str(path), 'groups': [
                {k: v for k, v in g.items() if k != 'members'} for g in groups]}
        matches = [g for g in groups if
                   (g['group_id'] == args['group_id'] if 'group_id' in args else
                    g['name'].casefold() == args['group'].casefold())]
        if len(matches) != 1:
            raise ValueError(f'Expected one exact group match; found {len(matches)}. Use list_endnote_groups.')
        group = matches[0]
        if not group['search_supported']:
            raise ValueError('Only manual groups are supported; smart/combined/online groups are not evaluated.')
        limit = int(args.get('limit', 25))
        if not 1 <= limit <= 200:
            raise ValueError('limit must be between 1 and 200.')
        filters = {k: str(args[k]).strip() for k in ('query','title','author','year','doi') if args.get(k) is not None}
        results = []
        for rec in sorted(group['members']):
            raw = library._fetch_sqlite(db, rec, False)
            if raw and library._matches(raw, filters):
                results.append(library._public_record(raw))
                if len(results) > limit:
                    break
        return {'ok': True, 'library_path': str(path), 'group_id': group['group_id'],
                'group': group['name'], 'member_count': group['member_count'],
                'references': results[:limit], 'count': min(len(results), limit),
                'truncated': len(results) > limit}


def inspect_document(path):
    """Read stored fields without Word, reformatting, or external connections."""
    fields, part_names = [], []
    with ZipFile(path) as z:
        for name in z.namelist():
            if not re.fullmatch(r'word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml', name):
                continue
            root = ET.fromstring(z.read(name))
            part_names.append(name)
            stack = []
            for elem in root.iter():
                if elem.tag == W+'fldChar':
                    kind = elem.get(W+'fldCharType')
                    if kind == 'begin':
                        stack.append({'code': '', 'result': '', 'part': name, '_result': False})
                    elif kind == 'separate' and stack:
                        stack[-1]['_result'] = True
                    elif kind == 'end' and stack:
                        item = stack.pop()
                        item.pop('_result')
                        fields.append(item)
                elif elem.tag == W+'instrText' and stack:
                    stack[-1]['code'] += elem.text or ''
                elif elem.tag == W+'t' and stack and stack[-1]['_result']:
                    stack[-1]['result'] += elem.text or ''
                elif elem.tag == W+'fldSimple':
                    fields.append({'code': elem.get(W+'instr', ''), 'part': name,
                                   'result': ''.join(t.text or '' for t in elem.iter(W+'t'))})
            if stack:
                raise ValueError(f'Unbalanced Word fields in {name}.')
        root = ET.fromstring(z.read('word/document.xml'))
        if root.findall('.//w:ins', NS) or root.findall('.//w:del', NS):
            raise ValueError('Active tracked changes require review before citation edits.')
    citations, bibliographies = [], []
    for field in fields:
        code = field['code'].strip()
        if re.match(r'^ADDIN EN\.CITE\s', code):
            fragment = code[code.index('<EndNote>'):].strip()
            data = ET.fromstring(fragment)
            records = []
            for cite in data.findall('./Cite'):
                records.append({'rec_number': int(cite.findtext('RecNum') or cite.findtext('record/rec-number')),
                                'title': ''.join(cite.find('record/titles/title').itertext())
                                         if cite.find('record/titles/title') is not None else '',
                                'author': cite.findtext('Author', ''), 'year': cite.findtext('Year', ''),
                                'database_ids': [e.get('db-id') for e in cite.findall('record/foreign-keys/key')]})
            citations.append({'part': field['part'], 'result': field['result'], 'records': records})
        elif re.match(r'^ADDIN EN\.REFLIST(?:\s|$)', code):
            bibliographies.append({'part': field['part'], 'result': field['result']})
    return {'ok': True, 'document_path': str(Path(path).resolve()),
            'citation_count': len(citations), 'bibliography_count': len(bibliographies),
            'citations': citations, 'bibliographies': bibliographies}


def _identity(record, expected):
    return (record['rec_number'] == expected['rec_number'] and
            ' '.join(record['title'].split()) == ' '.join(expected['title'].split()) and
            record['author'] == expected['first_author_surname'] and record['year'] == expected['year'])


def body_content(path):
    """Compare body text with generated EndNote field results excluded."""
    with ZipFile(path) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    stack=[]
    paragraphs=[]
    for paragraph in root.findall('.//w:body//w:p',NS):
        texts=[]
        for elem in paragraph.iter():
            if elem.tag == W+'fldChar':
                kind=elem.get(W+'fldCharType')
                if kind=='begin': stack.append({'code':'','endnote':False})
                elif kind=='separate' and stack:
                    stack[-1]['endnote']=bool(re.match(r'^\s*ADDIN EN\.(CITE|REFLIST)(?:\s|$)',stack[-1]['code']))
                elif kind=='end' and stack: stack.pop()
            elif elem.tag==W+'instrText' and stack:
                stack[-1]['code']+=elem.text or ''
            elif elem.tag==W+'t' and not any(f['endnote'] for f in stack):
                texts.append(elem.text or '')
        text=''.join(texts).rstrip(' ')
        if text: paragraphs.append(text)
    return paragraphs


def _invoke_local(config, work):
    config_path = Path(work)/'operation.json'
    config_path.write_text(json.dumps(config, ensure_ascii=False), encoding='utf-8')
    executable = upstream._powershell_executable()
    command = [executable, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
               '-File', str(ROOT/'scripts/edit-citation.ps1'), '-ConfigPath', str(config_path)]
    before = upstream._snapshot_watched_pids()
    timeout = int(os.environ.get('ENDNOTE_WORD_TIMEOUT_SECONDS', '120'))
    try:
        completed = subprocess.run(command, capture_output=True, text=True, encoding='utf-8',
                                   errors='replace', timeout=timeout,
                                   creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    except subprocess.TimeoutExpired as exc:
        stderr = exc.stderr or b''
        if isinstance(stderr, bytes):
            stderr = stderr.decode('utf-8','replace')
        remember_pid(stderr)
        cleaned = safe_reap(before)
        return {'ok': False, 'error': f'Word/CWYW timed out after {timeout}s; no output published.',
                'cleanup': cleaned, 'stderr_tail': stderr[-4000:]}
    if completed.stderr:
        print(completed.stderr, file=sys.stderr)
    lines = completed.stdout.strip().splitlines()
    if not lines:
        raise RuntimeError(f'No bridge result; exit {completed.returncode}. {completed.stderr[-1000:]}')
    result = json.loads(lines[-1])
    if completed.returncode:
        result['ok'] = False
    return result


def edit_citation(args, action):
    source = Path(args['input_path']).resolve()
    overwrite = args.get('overwrite', False)
    if not isinstance(overwrite, bool):
        raise ValueError('overwrite must be a boolean.')
    if not args.get('output_path') and not overwrite:
        raise ValueError('Provide a separate output_path or overwrite=true to update the input.')
    output = Path(args.get('output_path') or source).resolve()
    in_place = source == output
    if source.suffix.lower() != '.docx' or output.suffix.lower() != '.docx':
        raise ValueError('Input and output must be .docx paths.')
    if in_place and not overwrite:
        raise ValueError('Use separate input/output paths or overwrite=true to update the input.')
    if overwrite and not in_place:
        raise ValueError('overwrite=true only supports updating input_path.')
    if not source.is_file() or not output.parent.is_dir():
        raise ValueError('Input and output parent directory must exist.')
    if output.exists() and not in_place:
        raise ValueError('Output exists; choose a new path.')
    if not str(args.get('paragraph_text','')).strip():
        raise ValueError('paragraph_text is required to bound the change.')
    if 'anchor_text' in args and not str(args['anchor_text']).strip():
        raise ValueError('anchor_text must be nonempty literal body text.')
    if args.get('position','after') not in ('before','after'):
        raise ValueError('position must be before or after.')
    if action == 'move' and not args.get('anchor_text'):
        raise ValueError('Moving a citation requires anchor_text.')
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    before = inspect_document(source)
    if any(c['part'] != 'word/document.xml' for c in before['citations']):
        raise ValueError('Targeted editing currently supports citations in the main document only.')
    record = None
    if action == 'insert':
        got = library.get_reference(args['library_path'], args['rec_number'])
        if not got.get('ok') or not got.get('found'):
            raise ValueError('Record not found in the requested library.')
        record = got['reference']
        if record['title'] != args['expected_title']:
            raise ValueError('Record number/title identity mismatch.')
        if not record['first_author_surname'] or not re.fullmatch(r'\d{4}', record['year']):
            raise ValueError('The existing record needs a first author and four-digit year.')
        if any(x in record['first_author_surname'] for x in '{};#\r\n'):
            raise ValueError('Author name contains temporary-citation syntax characters.')
        if args.get('group'):
            grouped = group_operation({'library_path':args['library_path'], 'group':args['group'],
                                       'title':record['title'], 'limit':200}, search=True)
            if not any(r['rec_number'] == args['rec_number'] for r in grouped['references']):
                raise ValueError('Record is not in the specified manual group.')
    with tempfile.TemporaryDirectory(prefix='endnote-edit-',dir=output.parent) as work:
        scratch = Path(work)/'working.docx'
        shutil.copy2(source,scratch)
        config = {'remove_preceding_space':False, **args, 'action':action, 'scratch_path':str(scratch), 'record':record}
        result = _invoke_local(config,work)
        if not result.get('ok'):
            return result
        after = inspect_document(scratch)
        expected_count = before['citation_count'] + (1 if action=='insert' else -1 if action=='remove' else 0)
        if after['citation_count'] != expected_count:
            raise ValueError('Stored citation-field count did not match the requested operation.')
        if action == 'insert':
            if not any(_identity(r,record) for c in after['citations'] for r in c['records']):
                raise ValueError('CWYW produced a different record identity.')
        def identities(doc):
            return Counter((r['rec_number'], r['title'], r['author'], r['year'])
                           for c in doc['citations'] for r in c['records'])
        expected = identities(before)
        if action == 'insert':
            key = (record['rec_number'],record['title'],record['first_author_surname'],record['year'])
            expected[key] += 1
        elif action == 'remove':
            removed = [key for key in expected if key[0] == args['rec_number'] and key[1] == args['expected_title']]
            if len(removed) != 1:
                raise ValueError('Removal identity is ambiguous in stored citation data.')
            expected[removed[0]] -= 1
            expected = +expected
        if identities(after) != expected:
            raise ValueError('An unrelated citation identity changed; no output published.')
        if body_content(source) != body_content(scratch):
            raise ValueError('Unrelated body text changed; no output published.')
        if after['citation_count'] and not any(b['result'].strip() for b in after['bibliographies']):
            raise ValueError('Remaining citations require a nonempty bibliography.')
        if not after['citation_count'] and after['bibliography_count']:
            raise ValueError('Last-citation removal left an orphan bibliography field.')
        if hashlib.sha256(source.read_bytes()).hexdigest() != original_hash:
            raise ValueError('Source changed during the operation; refusing publication.')
        backup = None
        if in_place:
            backup_dir = source.parent / '.endnote-backups'
            backup_dir.mkdir(exist_ok=True)
            fd, backup_name = tempfile.mkstemp(prefix=source.stem + '-', suffix='.docx', dir=backup_dir)
            os.close(fd)
            backup = Path(backup_name)
            try:
                shutil.copy2(source, backup)
                if hashlib.sha256(backup.read_bytes()).hexdigest() != original_hash:
                    raise ValueError('Source changed during backup; refusing publication.')
            except Exception:
                backup.unlink()
                raise
            if hashlib.sha256(source.read_bytes()).hexdigest() != original_hash:
                raise ValueError('Source changed during the operation; refusing publication.')
            os.replace(scratch, output)
        elif output.exists():
            raise ValueError('Output appeared during editing; refusing to overwrite.')
        else:
            # On Windows rename is atomic and refuses an existing destination.
            os.rename(scratch,output)
        return {'ok': True, 'action':action, 'output_path':str(output), 'source_unchanged':not in_place,
                'in_place':in_place, 'backup_path':str(backup) if backup else None,
                'citation_count_before':before['citation_count'], 'citation_count_after':after['citation_count'],
                'bibliography_count':after['bibliography_count'], 'other_body_text_preserved':True, 'bridge':result,
                'doi_required':False if action=='insert' else None}
