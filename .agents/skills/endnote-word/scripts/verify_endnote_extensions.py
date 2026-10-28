"""Exercise installed MCP servers over real stdio; public samples only for online checks."""
import argparse
import json
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
from zipfile import ZipFile

WORKSPACE = Path(__file__).resolve().parents[4]
SCRIPTS = Path(__file__).resolve().parent
TOOLS = WORKSPACE / '.tools/endnote'


class Client:
    def __init__(self, command, env=None):
        self.queue = queue.Queue()
        self.logs = []
        self.id = 0
        self.process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                        stderr=subprocess.PIPE, encoding='utf-8',
                                        env={**os.environ, **(env or {})},
                                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        def collect(stream, callback):
            for line in stream:
                callback(line)
            if stream is self.process.stdout:
                self.queue.put(None)
        threading.Thread(target=collect, args=(self.process.stdout, self.queue.put), daemon=True).start()
        threading.Thread(target=collect, args=(self.process.stderr, self.logs.append), daemon=True).start()

    def call(self, method, params=None, timeout=120):
        self.id += 1
        self.process.stdin.write(json.dumps({'jsonrpc': '2.0', 'id': self.id,
                                            'method': method, 'params': params or {}}) + '\n')
        self.process.stdin.flush()
        while True:
            try:
                line = self.queue.get(timeout=timeout)
            except queue.Empty:
                raise RuntimeError('MCP response timed out: ' + ''.join(self.logs[-10:]))
            if line is None:
                self.process.wait(timeout=5)
                raise RuntimeError('MCP exited: ' + ''.join(self.logs[-10:]))
            message = json.loads(line)
            if message.get('id') == self.id:
                if 'error' in message:
                    raise RuntimeError(message['error'])
                return message['result']

    def initialize(self):
        info = self.call('initialize', {'protocolVersion': '2025-06-18', 'capabilities': {},
                                      'clientInfo': {'name': 'endnote-integration-check', 'version': '1'}})
        self.process.stdin.write('{"jsonrpc":"2.0","method":"notifications/initialized"}\n')
        self.process.stdin.flush()
        return info

    def close(self):
        self.process.stdin.close()
        try:
            self.process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.process.terminate()
            self.process.wait(timeout=5)


def run(online=False, semantic=False):
    output = TOOLS / 'validation'
    output.mkdir(parents=True, exist_ok=True)
    bib = output / 'public-example.bib'
    bib.write_text('@article{LeCun2015,\n title={Deep learning},\n author={LeCun, Yann and Bengio, Yoshua and Hinton, Geoffrey},\n'
                   ' year={2015},\n journal={Nature},\n doi={10.1038/nature14539}\n}', encoding='utf-8')
    doc = output / 'public-example.docx'
    with ZipFile(doc, 'w') as z:
        z.writestr('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   '<w:body><w:p><w:r><w:t>References</w:t></w:r></w:p>'
                   '<w:p><w:r><w:t>LeCun Y, Bengio Y, Hinton G. Deep learning. Nature. 2015;521:436-444. doi:10.1038/nature14539.</w:t>'
                   '</w:r></w:p></w:body></w:document>')
    node = 'C:/Program Files/nodejs/node.exe'
    cite_command = [node, '--import', (TOOLS / 'citecheck/node_modules/tsx/dist/loader.mjs').as_uri(),
                    str(TOOLS / 'citecheck/node_modules/@jhlee0619/citecheck/src/mcp.ts')]
    cite_env = {'CITECHECK_PYTHON': sys.executable, 'REFFORGE_ENABLE_SEMANTIC_SCHOLAR': 'true',
                'REFFORGE_HTTP_MAX_RETRIES': '0', 'REFFORGE_HTTP_TIMEOUT_MS': '7000',
                **{f'REFFORGE_{s}_RETRIES': '0' for s in ('PUBMED', 'CROSSREF', 'ARXIV', 'SEMANTIC_SCHOLAR')}}
    jobs = [
        ('local', [sys.executable, '-X', 'utf8', str(SCRIPTS / 'endnote_mcp.py')], {}, [
            ('get_endnote_capabilities', {}), ('get_endnote_search_index_status', {})]),
        ('search', [sys.executable, '-X', 'utf8', str(SCRIPTS / 'endnote_library_runner.py'), 'serve'], {}, [
            ('search_references', {'query': 'sulfur', 'limit': 2}),
            ('search_fulltext', {'query': 'model', 'limit': 2}),
            *([('search_semantic', {'query': 'sulfur recovery process control', 'limit': 2})] if semantic else [])]),
        ('citecheck', cite_command, cite_env, [('citecheck_version', {}),
            ('scan_workspace', {'target_path': str(bib)}),
            *([('analyze_references', {'target_path': str(doc), 'mode': 'review'})] if online else [])]),
    ]
    reports = {}
    for name, command, env, calls in jobs:
        client = Client(command, env)
        try:
            initialized = client.initialize()
            catalog = client.call('tools/list')
            result = {'server': initialized['serverInfo'], 'tools': [t['name'] for t in catalog['tools']], 'calls': []}
            for tool, args in calls:
                answer = client.call('tools/call', {'name': tool, 'arguments': args})
                if answer.get('isError'):
                    raise RuntimeError(answer)
                visible = ''.join(v.get('text', '') for v in answer.get('content', []))
                if tool == 'search_fulltext' and 'No fulltext' in visible:
                    raise RuntimeError('Expected an indexed PDF full-text match.')
                if tool == 'search_semantic' and 'by semantic similarity' not in visible:
                    raise RuntimeError('Expected real semantic matches, received: ' + visible)
                result['calls'].append({'tool': tool, 'result': answer})
            reports[name] = result
            print(name, 'OK:', len(result['tools']), 'tools,', len(result['calls']), 'calls', flush=True)
        finally:
            client.close()
    (output / 'mcp-smoke.json').write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding='utf-8')
    print('Report:', output / 'mcp-smoke.json')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--online', action='store_true')
    parser.add_argument('--semantic', action='store_true')
    args = parser.parse_args()
    run(args.online, args.semantic)
