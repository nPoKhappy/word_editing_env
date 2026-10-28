"""Local EndNote MCP: upstream tools plus manual groups and targeted citation edits."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
import local_endnote as local
import endnote_metadata as metadata
import endnote_integrations as integrations
import batch_citations
import journal_diagnostics
import jtice_preflight
import import_verified_reference

base = local.upstream
STRING = {'type':'string'}
INTEGER = {'type':'integer','minimum':1}


def tool(name, description, properties, required, readonly):
    return {'name':name, 'description':description,
            'inputSchema':{'type':'object','properties':properties,'required':required,'additionalProperties':False},
            'annotations':{'readOnlyHint':readonly,'destructiveHint':False,'openWorldHint':False}}


COMMON = {'input_path':STRING,'output_path':STRING,'paragraph_text':STRING,
          'rec_number':INTEGER,'expected_title':STRING,
          'overwrite':{'type':'boolean','default':False,
                       'description':'Set true to update input_path with an automatic backup. Omit output_path or set it to input_path.'}}
EDIT_REQUIRED = ['input_path','paragraph_text','rec_number','expected_title']
EXTRA = [
    tool('insert_endnote_citation_at_text','Insert a genuine CWYW citation immediately before or after one unique literal anchor inside a unique main-body paragraph. Use this for sentence-level citations; paragraph-end insertion is not a substitute.',
         {**COMMON,'library_path':STRING,'group':STRING,'anchor_text':STRING,'position':{'type':'string','enum':['before','after'],'default':'after'}},EDIT_REQUIRED+['library_path','anchor_text'],False),
    tool('move_endnote_citation_to_text','Move one existing standalone native CWYW field to a unique literal anchor in the same main-body paragraph, preserving record identity and all other content. Does not fabricate fields. Merged fields are refused.',
         {**COMMON,'anchor_text':STRING,'position':{'type':'string','enum':['before','after'],'default':'after'},'remove_preceding_space':{'type':'boolean','default':True}},EDIT_REQUIRED+['anchor_text'],False),
    tool('list_endnote_groups','List actual EndNote groups read-only. Manual groups can be searched; smart/online groups are reported as unsupported.',
         {'library_path':STRING},['library_path'],True),
    tool('search_endnote_group','Search one exact manual group name (case-insensitive), or group_id. Membership is read from a library snapshot; no library writes.',
         {'library_path':STRING,'group':STRING,'group_id':INTEGER,'query':STRING,'title':STRING,
          'author':STRING,'year':INTEGER,'doi':STRING,'limit':{'type':'integer','minimum':1,'maximum':200}},['library_path'],True),
    tool('inspect_endnote_document','Inspect stored native citation records and bibliography fields without opening Word. Does not check page layout.',
         {'document_path':STRING},['document_path'],True),
    tool('insert_endnote_citation_after_paragraph','Insert one existing EndNote record at the end of the unique main-body paragraph containing paragraph_text. Uses native CWYW and exact record/title identity. Provide a new output_path to preserve the source, or overwrite=true to update input_path with an automatic backup after validation. DOI may be empty. Optional group verifies membership.',
         {**COMMON,'library_path':STRING,'group':STRING},EDIT_REQUIRED+['library_path'],False),
    tool('remove_endnote_citation','Delete exactly one standalone citation field matching record/title in the unique paragraph. Reformat remaining references; removing the last citation clears the bibliography. Does not delete the library record. Merged citations are refused. Provide a new output_path, or overwrite=true to update input_path with an automatic backup after validation.',
         {**COMMON,'remove_preceding_space':{'type':'boolean','default':False}},EDIT_REQUIRED,False),
]
LIBRARY_SCOPE = {'library_path': STRING, 'group': STRING, 'group_id': INTEGER}
REC_NUMBERS = {'type': 'array', 'items': INTEGER, 'minItems': 1, 'uniqueItems': True}
METADATA_ROUTES = {
    'get_endnote_capabilities': metadata.capabilities,
    'search_online_endnote_metadata': metadata.search_online,
    'lookup_endnote_doi_metadata': metadata.lookup_doi,
    'verify_endnote_reference_online': integrations.verify_online,
    'verify_endnote_references_online': integrations.verify_online,
    'audit_endnote_library': metadata.audit_library,
    'find_endnote_duplicate_candidates': metadata.duplicates,
    'list_endnote_journal_terms': metadata.journal_terms,
    'prepare_endnote_journal_terms': metadata.prepare_terms,
    'audit_endnote_attachments': metadata.attachments,
    'export_endnote_records': metadata.export_records,
    'backup_endnote_library': metadata.backup_library,
    'restore_endnote_library_backup': metadata.restore_backup,
    'audit_endnote_document_library': metadata.audit_document,
    'get_endnote_reference_full': integrations.get_full_record,
    'list_endnote_tags': integrations.list_tags,
    'list_endnote_group_tree': integrations.group_tree,
    'export_endnote_catalog': integrations.export_catalog,
    'get_endnote_search_index_status': integrations.index_status,
    'refresh_endnote_search_index': integrations.refresh_index,
    'batch_edit_endnote_citations': batch_citations.edit_batch,
    'diagnose_endnote_journal_abbreviations': journal_diagnostics.diagnose,
    'preflight_jtice_manuscript': jtice_preflight.inspect,
    'import_verified_endnote_reference': import_verified_reference.import_reference,
}
EXTRA.extend([
    tool('get_endnote_capabilities', 'Report implemented capabilities and explicit native writeback limits; does not claim unimplemented record editing.', {}, [], True),
    tool('search_online_endnote_metadata', 'Search public Crossref bibliographic metadata. Results are candidates, never automatic identity matches.',
         {'query': STRING, 'year': INTEGER, 'author': STRING,
          'match_mode': {'type': 'string', 'enum': ['title', 'bibliographic'], 'default': 'bibliographic'},
          'limit': {'type': 'integer', 'minimum': 1, 'maximum': 10}}, ['query'], True),
    tool('lookup_endnote_doi_metadata', 'Retrieve a DOI from Crossref, then DataCite as fallback, with registry source URL and retrieval time. Does not write a library.',
         {'doi': STRING}, ['doi'], True),
    tool('verify_endnote_reference_online', 'Use installed citecheck to verify an actual library record against scholarly indexes and return field differences and replacement-safety diagnostics. Sends bibliographic fields to public indexes. Review only; no native library writeback.',
         {'library_path': STRING, 'rec_number': INTEGER, 'expected_title': STRING, 'report_path': STRING}, ['library_path', 'rec_number'], False),
    tool('verify_endnote_references_online', 'Use installed citecheck for up to three real records per online batch, with manual group/record selection and remaining IDs. Sends bibliographic fields to public indexes. Review only; no native writeback.',
         {**LIBRARY_SCOPE, 'rec_numbers': REC_NUMBERS, 'limit': {'type': 'integer', 'minimum': 1, 'maximum': 3}, 'report_path': STRING}, ['library_path'], False),
    tool('audit_endnote_library', 'Check all active records or an exact manual group for missing identity/journal fields and malformed DOI/year. Offline; a missing DOI is not proof no DOI exists.',
         {**LIBRARY_SCOPE, 'report_path': STRING}, ['library_path'], False),
    tool('find_endnote_duplicate_candidates', 'Find same DOI or normalized title/first-author/year candidates across all active records or an exact manual group. Does not merge or delete.',
         LIBRARY_SCOPE, ['library_path'], True),
    tool('list_endnote_journal_terms', 'Read actual journal full names and all three abbreviations from a snapshot; optionally filter by query.',
         {'library_path': STRING, 'query': STRING}, ['library_path'], True),
    tool('prepare_endnote_journal_terms', 'Preview unique-name journal terms, fill unambiguous blank abbreviations and flag conflicts. Optionally export a four-column UTF-8 term file; no native import or library change.',
         {'library_path': STRING, 'output_path': STRING}, ['library_path'], False),
    tool('audit_endnote_attachments', 'Resolve stored attachment paths and report missing files without altering, opening, or downloading them.',
         {'library_path': STRING}, ['library_path'], True),
    tool('export_endnote_records', 'Export selected actual records as JSON or RIS to a new file, preserving supported native types. JSON/RIS is a bibliographic export, not a full library backup.',
         {**LIBRARY_SCOPE, 'rec_numbers': REC_NUMBERS, 'output_path': STRING, 'format': {'type': 'string', 'enum': ['json', 'ris'], 'default': 'json'}}, ['library_path', 'output_path'], False),
    tool('backup_endnote_library', 'Back up a closed library .enl plus its entire .Data folder with SHA-256 manifest and source-change checks. Refuses an open writer; never closes EndNote.',
         {'library_path': STRING, 'output_directory': STRING}, ['library_path', 'output_directory'], False),
    tool('restore_endnote_library_backup', 'Verify SHA-256 backup contents and restore the .enl/.Data pair into a new directory. Refuses existing destinations and never overwrites or opens a live library.',
         {'backup_directory': STRING, 'output_directory': STRING}, ['backup_directory', 'output_directory'], False),
    tool('audit_endnote_document_library', 'Compare stored native CWYW record/title/author/year identities against the specified library and check for a bibliography. Does not edit or render Word.',
         {'library_path': STRING, 'document_path': STRING}, ['library_path', 'document_path'], True),
    tool('get_endnote_reference_full', 'Read complete stored fields, full abstract, notes, attachments and tag IDs from one actual active record through the installed CLI reader on a read-only snapshot.',
         {'library_path': STRING, 'rec_number': INTEGER}, ['library_path', 'rec_number'], True),
    tool('list_endnote_tags', 'Read color tags and active record membership through the installed CLI reader on a read-only snapshot.',
         {'library_path': STRING}, ['library_path'], True),
    tool('list_endnote_group_tree', 'Read group-set hierarchy using the installed CLI reader and validated manual membership decoding. Does not evaluate smart groups.',
         {'library_path': STRING}, ['library_path'], True),
    tool('export_endnote_catalog', 'Reuse installed CLI exporters for selected active records as JSON or CSV in a new file, with actual installed native reference-type names.',
         {**LIBRARY_SCOPE, 'rec_numbers': REC_NUMBERS, 'output_path': STRING, 'format': {'type': 'string', 'enum': ['json', 'csv']}}, ['library_path', 'output_path', 'format'], False),
    tool('get_endnote_search_index_status', 'Report the local full-text/semantic search snapshot source, age and whether the library changed since refresh.', {}, [], True),
    tool('refresh_endnote_search_index', 'Refresh the installed EndNote search MCP from a read-only library snapshot; rebuild local PDF text and optional CPU embeddings. Writes search caches only. Model download may be required on first use.',
         {'library_path': STRING, 'semantic': {'type': 'boolean', 'default': True}}, ['library_path'], False),
    tool('diagnose_endnote_journal_abbreviations', 'Separate missing journal metadata, duplicate full-title terms (even with equal abbreviations), conflicting abbreviations and actual CWYW output. Optional document_path maps standalone numbered citations. Reports native style/template review needs; no library writes.',
         {**LIBRARY_SCOPE, 'document_path': STRING, 'expect_abbreviation': {'type': 'boolean', 'default': True}}, ['library_path'], True),
    tool('preflight_jtice_manuscript', 'Inventory an existing DOCX before drafting/translation/formatting: native citations, source omissions, tables, explicit font sizes, captions, math and journal-term diagnostics. Read-only and offline; does not replace final Word font checks, directory updates or rendered QA.',
         {'document_path': STRING, 'source_document_path': STRING, 'library_path': STRING,
          'stage': {'type': 'string', 'enum': ['draft', 'translation', 'submission'], 'default': 'draft'}}, ['document_path'], True),
    tool('batch_edit_endnote_citations', 'Apply up to 50 native insert/move/remove operations on one temporary DOCX in one owned Word session, then format CWYW once. Validates native identities and unrelated body before atomic publication. Use for multiple edits; avoids restarting Word per citation. Insert/move require unique literal anchors; merged fields are refused.',
         {'input_path': STRING, 'output_path': STRING, 'overwrite': {'type': 'boolean', 'default': False}, 'library_path': STRING,
          'operations': {'type': 'array', 'minItems': 1, 'maxItems': 50, 'items': {'type': 'object', 'additionalProperties': False,
             'properties': {'action': {'type': 'string', 'enum': ['insert', 'move', 'remove']}, 'paragraph_text': STRING,
                            'rec_number': INTEGER, 'expected_title': STRING, 'anchor_text': STRING,
                            'position': {'type': 'string', 'enum': ['before', 'after']}, 'remove_preceding_space': {'type': 'boolean'}},
             'required': ['action', 'paragraph_text', 'rec_number', 'expected_title']}}}, ['input_path', 'operations'], False),
    tool('import_verified_endnote_reference', 'One workflow: resolve title/DOI/DOI-containing URL or first-page PDF DOI, require unambiguous registry identity, avoid duplicates, generate RIS, import via native EndNote and re-query actual assigned record numbers. May modify the specified library. No existing-record overwrite, no silent retry, no automatic JTICE style compliance claim. preview_only prepares evidence/RIS without import.',
         {'library_path': STRING, 'title': STRING, 'doi': STRING, 'url': STRING, 'pdf_path': STRING, 'author': STRING,
          'year': INTEGER, 'preview_only': {'type': 'boolean', 'default': False}}, ['library_path'], False),
])
for entry in EXTRA:
    if entry['name'] in {'search_online_endnote_metadata', 'lookup_endnote_doi_metadata',
                          'verify_endnote_reference_online', 'verify_endnote_references_online', 'import_verified_endnote_reference'}:
        entry['annotations']['openWorldHint'] = True
EXTRA_BY_NAME = {t['name']:t for t in EXTRA}
base.TOOLS.extend(EXTRA)
original_call = base._call_tool
original_handle = base._handle
base.SERVER_NAME = 'endnote-word-local'
base.SERVER_VERSION = '1.5.0'


def validate(value, schema, path='arguments'):
    kind = schema.get('type')
    if kind == 'object':
        if not isinstance(value, dict):
            raise ValueError(f'{path} must be an object.')
        missing = set(schema.get('required', [])) - set(value)
        unknown = set(value) - set(schema.get('properties', {}))
        if missing or unknown:
            raise ValueError(f'Invalid arguments; missing={sorted(missing)}, unknown={sorted(unknown)}')
        for key, item in value.items():
            validate(item, schema['properties'][key], path + '.' + key)
    elif kind == 'string' and (not isinstance(value, str) or not value.strip()):
        raise ValueError(f'{path} must be a nonempty string.')
    elif kind == 'integer':
        if type(value) is not int or value < schema.get('minimum', -float('inf')) or value > schema.get('maximum', float('inf')):
            raise ValueError(f'{path} must be an integer within the allowed range.')
    elif kind == 'boolean' and type(value) is not bool:
        raise ValueError(f'{path} must be a boolean.')
    elif kind == 'array':
        if not isinstance(value, list) or len(value) < schema.get('minItems', 0) or len(value) > schema.get('maxItems', float('inf')):
            raise ValueError(f'{path} must be a nonempty array.')
        for item in value:
            validate(item, schema['items'], path)
        if schema.get('uniqueItems') and len(value) != len(set(value)):
            raise ValueError(f'{path} must not contain duplicates.')
    if 'enum' in schema and value not in schema['enum']:
        raise ValueError(f'{path} must be one of {schema["enum"]}.')


def call(name, args):
    if name not in EXTRA_BY_NAME:
        return original_call(name,args)
    try:
        schema = EXTRA_BY_NAME[name]['inputSchema']
        missing = [k for k in schema['required'] if k not in args]
        if missing or set(args)-set(schema['properties']):
            raise ValueError(f'Invalid arguments; missing={missing}, unknown={sorted(set(args)-set(schema["properties"]))}')
        if name in METADATA_ROUTES:
            validate(args, schema)
            if 'group' in args and 'group_id' in args:
                raise ValueError('Provide at most one of group or group_id.')
            result = METADATA_ROUTES[name](args)
        elif name == 'list_endnote_groups':
            result = local.group_operation(args)
        elif name == 'search_endnote_group':
            if ('group' in args) == ('group_id' in args):
                raise ValueError('Provide exactly one of group or group_id.')
            result = local.group_operation(args,True)
        elif name == 'inspect_endnote_document':
            result = local.inspect_document(args['document_path'])
        else:
            result = local.edit_citation(args,'insert' if name.startswith('insert_') else 'move' if name.startswith('move_') else 'remove')
    except Exception as exc:
        result = {'ok':False,'error':f'{type(exc).__name__}: {exc}'}
    return {'content':[{'type':'text','text':json.dumps(result,ensure_ascii=False)}],
            'structuredContent':result,'isError':not result.get('ok',False)}


def handle(message):
    answer = original_handle(message)
    if message.get('method') == 'initialize' and answer:
        answer['result']['instructions'] = (
            'Resolve actual library records and manual-group membership before editing. '
            'Use targeted insert/remove tools with a unique paragraph and a new output_path, '
            'or overwrite=true to update input_path with an automatic backup after validation. '
            'Native CWYW creates citations; never fabricate citation fields. Existing records may lack DOI. '
            'Inspect stored fields afterwards; layout requires a separate Word/PDF check. '
            'Merged-field removal and smart-group evaluation are not supported. '
            'Upstream manifest rendering/precheck retains mandatory DOI verification and its configured default style. '
            'Do not configure defaults or import RIS unless requested. No library record deletion tool is provided.'
            ' Online metadata verification returns evidence and correction previews, not successful writeback. '
            'Use get_endnote_capabilities for native editor limits; no direct database writes are allowed.'
            ' For JTICE, run preflight_jtice_manuscript before changes. Use one batch_edit_endnote_citations for multiple edits. '
            'Diagnose record fields, journal terms and style output separately. Native import is available through import_verified_endnote_reference.'
        )
    return answer


base._call_tool = call
base._handle = handle


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--call',choices=[t['name'] for t in base.TOOLS])
    parser.add_argument('--arguments',type=Path,help='UTF-8 JSON tool arguments for a direct one-shot call.')
    args=parser.parse_args()
    if args.call:
        base._force_utf8_streams()
        arguments=json.loads(args.arguments.read_text(encoding='utf-8-sig')) if args.arguments else {}
        result=call(args.call,arguments)
        print(json.dumps(result.get('structuredContent',result),ensure_ascii=False,indent=2))
        return 1 if result.get('isError') else 0
    return base.main()


if __name__=='__main__':
    raise SystemExit(main())
