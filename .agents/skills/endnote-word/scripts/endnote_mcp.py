"""Local EndNote MCP: upstream tools plus manual groups and targeted citation edits."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
import local_endnote as local

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
EXTRA_BY_NAME = {t['name']:t for t in EXTRA}
base.TOOLS.extend(EXTRA)
original_call = base._call_tool
original_handle = base._handle
base.SERVER_NAME = 'endnote-word-local'
base.SERVER_VERSION = '1.1.0'


def call(name, args):
    if name not in EXTRA_BY_NAME:
        return original_call(name,args)
    try:
        schema = EXTRA_BY_NAME[name]['inputSchema']
        missing = [k for k in schema['required'] if k not in args]
        if missing or set(args)-set(schema['properties']):
            raise ValueError(f'Invalid arguments; missing={missing}, unknown={sorted(set(args)-set(schema["properties"]))}')
        if name == 'list_endnote_groups':
            result = local.group_operation(args)
        elif name == 'search_endnote_group':
            if ('group' in args) == ('group_id' in args):
                raise ValueError('Provide exactly one of group or group_id.')
            result = local.group_operation(args,True)
        elif name == 'inspect_endnote_document':
            result = local.inspect_document(args['document_path'])
        else:
            result = local.edit_citation(args,'insert' if name.startswith('insert_') else 'remove')
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
