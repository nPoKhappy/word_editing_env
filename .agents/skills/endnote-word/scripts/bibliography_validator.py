"""Fail closed unless numbered CWYW citations match the generated bibliography.

This validator reads only EndNote-produced fields. It does not generate citations
or bibliography entries. Styles without bracketed numeric labels or without full
article titles cannot be verified by this method and are rejected explicitly.
"""
import re
import unicodedata


def _normal(text):
    return ''.join(c for c in unicodedata.normalize('NFKC', text).casefold() if c.isalnum())


def _numbers(display):
    match = re.fullmatch(r'\s*\[([\d\s,;\-–—]+)\]\s*', display)
    if not match:
        raise ValueError('Cannot verify non-bracketed numeric or author-date citation style; no output published.')
    numbers = []
    for item in re.split(r'[,;]', match.group(1)):
        item = item.strip()
        range_match = re.fullmatch(r'(\d+)\s*[-–—]\s*(\d+)', item)
        if range_match:
            first, last = map(int, range_match.groups())
            if first < 1 or last < first or last - first > 999:
                raise ValueError('Invalid numeric citation range; no output published.')
            numbers.extend(range(first, last + 1))
        elif re.fullmatch(r'\d+', item) and int(item) > 0:
            numbers.append(int(item))
        else:
            raise ValueError('Cannot verify numeric citation label; no output published.')
    if len(numbers) != len(set(numbers)):
        raise ValueError('Duplicate numeric labels in citation; no output published.')
    return numbers


def verify_bibliography(document):
    citations = document['citations']
    if not citations:
        if document['bibliography_count']:
            raise ValueError('Bibliography remains after last citation; no output published.')
        return {'unique_references': 0}
    if document['bibliography_count'] != 1 or len(document['bibliographies']) != 1:
        raise ValueError('Expected exactly one native bibliography field; no output published.')
    bib = document['bibliographies'][0]['result']
    matches = list(re.finditer(r'\[(\d+)\]', bib))
    if not matches:
        raise ValueError('Cannot verify non-numbered bibliography style; no output published.')
    entries = {}
    for i, match in enumerate(matches):
        number = int(match.group(1))
        if number in entries or number < 1:
            raise ValueError('Duplicate or invalid bibliography number; no output published.')
        end = matches[i+1].start() if i+1 < len(matches) else len(bib)
        entries[number] = _normal(bib[match.end():end])
    references = {}
    for citation in citations:
        for record in citation['records']:
            title = _normal(record['title'])
            if not title:
                raise ValueError('Cannot verify citation without a title; no output published.')
            identity = (record.get('rec_number'), record.get('author'), record.get('year'),
                        tuple(record.get('database_ids', [])))
            if title in references and references[title] != identity:
                raise ValueError('Distinct cited records have the same title; cannot verify bibliography safely.')
            references[title] = identity
    if len(entries) != len(references):
        raise ValueError('Bibliography entry count differs from unique cited references; no output published.')
    assignments = {}
    for number, entry in entries.items():
        found = [title for title in references if title in entry]
        if len(found) != 1 or found[0] in assignments.values():
            raise ValueError('Bibliography title missing, ambiguous, or duplicated; no output published.')
        assignments[number] = found[0]
    if set(entries) != set(range(1, len(entries)+1)):
        raise ValueError('Bibliography numbering is incomplete; no output published.')
    for citation in citations:
        labels = _numbers(citation['result'])
        expected_titles = [_normal(r['title']) for r in citation['records']]
        if len(labels) != len(expected_titles) or any(n not in assignments for n in labels):
            raise ValueError('Citation labels do not match bibliography entries; no output published.')
        if sorted(assignments[n] for n in labels) != sorted(expected_titles):
            raise ValueError('Citation label maps to a different bibliography title; no output published.')
    return {'unique_references': len(references), 'verified_style': 'bracketed numeric'}
