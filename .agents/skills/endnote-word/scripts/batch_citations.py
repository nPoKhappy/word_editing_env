"""Transactional native CWYW batch editing, one Word session and one final format."""
from collections import Counter
import hashlib
import os
from pathlib import Path
import re
import shutil
import tempfile
import time
import local_endnote as local


def identities(document):
    return Counter((r['rec_number'], r['title'], r['author'], r['year'])
                   for c in document['citations'] for r in c['records'])


def edit_batch(args):
    started = time.perf_counter()
    source = Path(args['input_path']).resolve()
    output = Path(args.get('output_path') or source).resolve()
    overwrite = args.get('overwrite', False)
    if type(overwrite) is not bool or (source == output) != overwrite:
        raise ValueError('Provide a new output_path, or overwrite=true for the same input path.')
    if source.suffix.lower() != '.docx' or output.suffix.lower() != '.docx' or not source.is_file():
        raise ValueError('Existing DOCX input and DOCX output are required.')
    if output != source and output.exists():
        raise ValueError('Output already exists.')
    operations = args['operations']
    if not isinstance(operations, list) or not 1 <= len(operations) <= 50:
        raise ValueError('A batch must contain 1-50 operations.')
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    before = local.inspect_document(source)
    if any(c['part'] != 'word/document.xml' for c in before['citations']):
        raise ValueError('Batch editing supports main-body citations only.')
    expected = identities(before)
    expected_count = before['citation_count']
    prepared, seen = [], set()
    cached_records = {}
    for operation in operations:
        op = {'remove_preceding_space': False, **operation}
        action = op['action']
        if action not in ('insert', 'move', 'remove'):
            raise ValueError('Unknown batch operation.')
        if not op['paragraph_text'].strip() or not op['expected_title'].strip():
            raise ValueError('Paragraph and expected title cannot be empty.')
        if action in ('insert', 'move') and not op.get('anchor_text', '').strip():
            raise ValueError('Batch insert/move requires a unique literal anchor_text.')
        key = (action, op['rec_number'], op['paragraph_text'], op.get('anchor_text'))
        if key in seen:
            raise ValueError('Duplicate batch operation.')
        seen.add(key)
        if action == 'insert':
            if not args.get('library_path'):
                raise ValueError('Insertion requires library_path.')
            if op['rec_number'] not in cached_records:
                got = local.library.get_reference(args['library_path'], op['rec_number'])
                if not got.get('ok'):
                    raise ValueError(got.get('error', 'Missing record.'))
                cached_records[op['rec_number']] = got['reference']
            record = cached_records[op['rec_number']]
            if record['title'] != op['expected_title'] or not record['first_author_surname'] or not re.fullmatch(r'\d{4}', record['year']):
                raise ValueError('Insertion record/title/author/year identity mismatch.')
            if any(c in record['first_author_surname'] for c in '{};#\r\n'):
                raise ValueError('Author contains temporary citation syntax.')
            op['record'] = record
            expected[(record['rec_number'], record['title'], record['first_author_surname'], record['year'])] += 1
            expected_count += 1
        elif action == 'remove':
            keys = [k for k in expected if k[0] == op['rec_number'] and k[1] == op['expected_title'] and expected[k] > 0]
            if len(keys) != 1:
                raise ValueError('Removal record identity is absent or ambiguous.')
            expected[keys[0]] -= 1
            expected_count -= 1
        prepared.append(op)
    with tempfile.TemporaryDirectory(prefix='endnote-batch-', dir=output.parent) as work:
        scratch = Path(work) / 'working.docx'
        shutil.copy2(source, scratch)
        bridge = local._invoke_local({'action': 'batch', 'scratch_path': str(scratch), 'operations': prepared}, work)
        if not bridge.get('ok'):
            return {'ok': False, 'error': 'Native batch failed; no output published.', 'bridge': bridge}
        after = local.inspect_document(scratch)
        if after['citation_count'] != expected_count or identities(after) != +expected:
            raise ValueError('Final native citation count or identity differs from the planned batch.')
        if local.body_content(source) != local.body_content(scratch):
            raise ValueError('Unrelated body text changed; no output published.')
        if expected_count and not any(b['result'].strip() for b in after['bibliographies']):
            raise ValueError('Native bibliography is missing or empty.')
        if not expected_count and after['bibliography_count']:
            raise ValueError('Last citation removal left a bibliography.')
        if hashlib.sha256(source.read_bytes()).hexdigest() != original_hash:
            raise ValueError('Source changed during batch; no output published.')
        backup = None
        if overwrite:
            folder = source.parent / '.endnote-backups'
            folder.mkdir(exist_ok=True)
            fd, name = tempfile.mkstemp(prefix=source.stem + '-', suffix='.docx', dir=folder)
            os.close(fd)
            backup = Path(name)
            shutil.copy2(source, backup)
            if hashlib.sha256(backup.read_bytes()).hexdigest() != original_hash or hashlib.sha256(source.read_bytes()).hexdigest() != original_hash:
                raise ValueError('Source changed before publication.')
            os.replace(scratch, source)
        else:
            os.rename(scratch, output)
    return {'ok': True, 'output_path': str(output), 'backup_path': str(backup) if backup else None,
            'source_unchanged': not overwrite, 'operation_count': len(prepared),
            'citation_count_before': before['citation_count'], 'citation_count_after': after['citation_count'],
            'other_body_text_preserved': True, 'bridge': bridge,
            'elapsed_seconds': round(time.perf_counter() - started, 3),
            'final_directory_and_render_qa_required': True}
