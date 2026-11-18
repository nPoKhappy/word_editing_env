"""Verify actual CWYW citation identities and positions, without generating fields."""
import base64
from collections import Counter, defaultdict
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'
SENTINEL = '\ufff0'


def is_cite(code):
    return bool(re.match(r'^\s*ADDIN EN\.CITE(?:\s|$)', code))


def parse_payload(field):
    code = field['code'].strip()
    payload = (code[code.index('<EndNote>'):] if '<EndNote>' in code
               else field['data']).strip()
    if not payload:
        raise ValueError('CWYW citation has no traveling-library payload')
    xml = ET.fromstring(payload)
    records = []
    for cite in xml.findall('./Cite'):
        number = cite.findtext('RecNum') or cite.findtext('record/rec-number')
        title = cite.find('record/titles/title')
        if not number or title is None:
            raise ValueError('CWYW citation record is missing identity')
        records.append((int(number), ''.join(title.itertext())))
    if not records:
        raise ValueError('CWYW citation has no records')
    return records


def citation_tokens(paragraph):
    """Return literal text and outer EN.CITE positions, ignoring citation results."""
    stack, tokens, visible = [], [], []
    for elem in paragraph.iter():
        if elem.tag == W + 'fldSimple':
            code = elem.get(W + 'instr', '')
            if is_cite(code) or re.match(r'^\s*ADDIN EN\.CITE\.DATA(?:\s|$)', code):
                raise ValueError('Simple EndNote fields are unsupported')
        elif elem.tag == W + 'fldChar':
            kind = elem.get(W + 'fldCharType')
            if kind == 'begin':
                payload = elem.find(W + 'fldData')
                encoded = payload.text if payload is not None and payload.text else ''
                data = (base64.b64decode(''.join(encoded.split()), validate=True)
                        .decode('utf-8').rstrip('\x00')) if encoded else ''
                stack.append({'code': '', 'result': '', 'data': data, 'in_result': False})
            elif kind == 'separate':
                if not stack:
                    raise ValueError('Unbalanced field separator')
                stack[-1]['in_result'] = True
            elif kind == 'end':
                if not stack:
                    raise ValueError('Unbalanced field end')
                item = stack.pop()
                if stack:
                    if stack[-1]['in_result']:
                        stack[-1]['result'] += item['result']
                    elif re.match(r'^\s*ADDIN EN\.CITE\.DATA(?:\s|$)', item['code']):
                        stack[-1]['data'] = stack[-1]['data'] or item['data'] or item['result']
                if is_cite(item['code']):
                    if stack:
                        raise ValueError('Unexpected nested outer citation')
                    tokens.append(parse_payload(item))
                    visible.append(SENTINEL)
        elif elem.tag == W + 'instrText' and stack:
            stack[-1]['code'] += elem.text or ''
        elif elem.tag == W + 't':
            value = elem.text or ''
            if SENTINEL in value:
                raise ValueError('Unsupported sentinel in document')
            if not stack:
                visible.append(value)
            elif stack[-1]['in_result']:
                stack[-1]['result'] += value
                if not any(is_cite(entry['code']) for entry in stack):
                    visible.append(value)
    if stack:
        raise ValueError('Field extends beyond target paragraph')
    marked = ''.join(visible)
    positions = defaultdict(Counter)
    fields = defaultdict(list)
    offset = token = 0
    for ch in marked:
        if ch == SENTINEL:
            records = tokens[token]
            positions[offset].update(records)
            fields[offset].append(records)
            token += 1
        else:
            offset += 1
    return marked.replace(SENTINEL, ''), positions, fields


def _paragraphs(path):
    with ZipFile(path) as package:
        root = ET.fromstring(package.read('word/document.xml'))
    return root.findall('./w:body/w:p', NS)


def _target(paragraphs, needle):
    matched = []
    for i, paragraph in enumerate(paragraphs):
        try:
            parsed = citation_tokens(paragraph)
        except ValueError:
            # Long bibliography fields can span paragraphs. They need not be
            # parsed unless they contain the requested literal paragraph text.
            if needle in ''.join(t.text or '' for t in paragraph.iter(W + 't')):
                raise
            continue
        if needle in parsed[0]:
            matched.append((i, parsed))
    if len(matched) != 1:
        raise ValueError('Target paragraph is not unique')
    return matched[0]


def _anchor(text, anchor, position):
    if not anchor or text.count(anchor) != 1:
        raise ValueError('Citation anchor is not unique')
    start = text.index(anchor)
    return start if position == 'before' else start + len(anchor)


def verify_anchors(source, output, operations):
    """Simulate insert/move/remove in each target paragraph; require exact positions."""
    original, actual = _paragraphs(source), _paragraphs(output)
    grouped = defaultdict(list)
    for operation in operations:
        needle = operation['paragraph_text']
        old_index, _ = _target(original, needle)
        new_index, _ = _target(actual, needle)
        if old_index != new_index:
            raise ValueError('Target paragraph moved during citation editing')
        grouped[old_index].append(operation)
    for index, ops in grouped.items():
        before_text, before_positions, before_fields = citation_tokens(original[index])
        after_text, after_positions, _ = citation_tokens(actual[index])
        if before_text != after_text:
            raise ValueError('Target body text changed during citation editing')
        expected = defaultdict(Counter, {k: Counter(v) for k, v in before_positions.items()})
        for op in ops:
            identity = (int(op['rec_number']), op['expected_title'])
            action = op['action']
            if action in ('remove', 'move'):
                found = [(pos, field) for pos, fields in before_fields.items() for field in fields
                         if identity in field and expected[pos][identity] > 0]
                if len(found) != 1 or len(found[0][1]) != 1:
                    raise ValueError('Move/remove requires one standalone source citation')
                old_pos = found[0][0]
                expected[old_pos][identity] -= 1
                expected[old_pos] += Counter()
            if action in ('insert', 'move'):
                pos = _anchor(before_text, op['anchor_text'], op.get('position', 'after'))
                expected[pos][identity] += 1
        expected = {p: c for p, c in expected.items() if c}
        observed = {p: c for p, c in after_positions.items() if c}
        if expected != observed:
            raise ValueError(f'Citation identities or positions differ in paragraph {index}')
    return {'validated_paragraphs': len(grouped), 'validated_operations': len(operations)}
