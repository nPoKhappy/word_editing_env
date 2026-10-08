import importlib.util
from contextlib import closing
import json
import os
from pathlib import Path
import sqlite3
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import local_endnote as local


class LocalTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        self.enl=self.root/'Library.enl'
        self.enl.touch()
        store=self.root/'Library.Data/sdb'
        store.mkdir(parents=True)
        self.dbpath=store/'sdb.eni'
        with closing(sqlite3.connect(self.dbpath)) as db, db:
            columns=[f'"{c}" '+('INTEGER' if c in ('id','trash_state','reference_type') else 'TEXT')
                     for c in local.library.SELECT_COLUMNS]
            db.execute('CREATE TABLE refs('+','.join(columns)+')')
            db.execute('CREATE TABLE groups(group_id INTEGER,spec BLOB,members BLOB)')
            for rec,trash in ((55,0),(56,1),(57,0)):
                db.execute('INSERT INTO refs(id,trash_state,reference_type,author,year,title) VALUES(?,?,0,?,?,?)',
                           (rec,trash,'Liu, Jialin','2025','Target paper' if rec!=57 else 'Other paper'))
            for gid,kind in ((9,3),(10,4)):
                spec=f'<group><ids><name>Claus</name></ids><rules><rule>TYPE;{kind}</rule></rules></group>'.encode()
                blob=b'\0\0\0\2'+struct.pack('<III',2,55,56)
                db.execute('INSERT INTO groups VALUES(?,?,?)',(gid,spec,blob))

    def tearDown(self):
        self.temp.cleanup()

    def test_manual_members_search_excludes_trash_and_nonmembers_preserves_library(self):
        before=self.dbpath.read_bytes()
        result=local.group_operation({'library_path':str(self.enl),'group_id':9,'title':'Target'},True)
        self.assertEqual([r['rec_number'] for r in result['references']],[55])
        self.assertEqual(self.dbpath.read_bytes(),before)
        self.assertEqual(result['references'][0]['doi'],'')

    def test_duplicate_names_require_group_id(self):
        with self.assertRaisesRegex(ValueError,'found 2'):
            local.group_operation({'library_path':str(self.enl),'group':'claus'},True)

    def test_smart_group_is_not_silently_treated_as_manual(self):
        with self.assertRaisesRegex(ValueError,'Only manual'):
            local.group_operation({'library_path':str(self.enl),'group_id':10},True)

    def test_corrupt_members_fail_closed(self):
        for blob in (b'wrong',b'\0\0\0\2'+struct.pack('<I',10),
                     b'\0\0\0\2'+struct.pack('<III',2,55,55)):
            with self.assertRaises(ValueError): local.decode_members(blob)

    def test_missing_group_not_global_search(self):
        with self.assertRaisesRegex(ValueError,'found 0'):
            local.group_operation({'library_path':str(self.enl),'group':'absent'},True)

    def test_search_limit_is_honored_with_truncation(self):
        with closing(sqlite3.connect(self.dbpath)) as db, db:
            db.execute('UPDATE refs SET trash_state=0 WHERE id=56')
        result=local.group_operation({'library_path':str(self.enl),'group_id':9,'limit':1},True)
        self.assertEqual(result['count'],1)
        self.assertTrue(result['truncated'])

    def test_upstream_timeout_cleanup_does_not_kill_unowned_new_processes(self):
        local._owned_word_pid=101
        with patch.object(local.upstream,'_process_pids',side_effect=[{10,101,102},{20,202}]), \
             patch.object(local.upstream,'_terminate_process_pid',return_value=True) as kill:
            result=local.safe_reap({'WINWORD.EXE':{10},'EndNote.exe':{20}})
        kill.assert_called_once_with(101)
        self.assertEqual(result['terminated'],[101])

    def test_upstream_timeout_without_owned_pid_kills_nothing(self):
        local._owned_word_pid=None
        with patch.object(local.upstream,'_process_pids',return_value={101}), \
             patch.object(local.upstream,'_terminate_process_pid') as kill:
            local.safe_reap({'WINWORD.EXE':set(),'EndNote.exe':set()})
        kill.assert_not_called()

    def test_mcp_handshake_catalog_and_group_search_over_stdio(self):
        messages=[{'jsonrpc':'2.0','id':1,'method':'initialize','params':{'protocolVersion':'2025-06-18'}},
                  {'jsonrpc':'2.0','method':'notifications/initialized'},
                  {'jsonrpc':'2.0','id':2,'method':'tools/list'},
                  {'jsonrpc':'2.0','id':3,'method':'tools/call','params':{'name':'search_endnote_group',
                   'arguments':{'library_path':str(self.enl),'group_id':9}}}]
        run=subprocess.run([sys.executable,'-X','utf8',str(ROOT/'scripts/endnote_mcp.py')],
                           input='\n'.join(json.dumps(m) for m in messages)+'\n',
                           capture_output=True,text=True,encoding='utf-8',timeout=15)
        self.assertEqual(run.returncode,0,run.stderr)
        responses=[json.loads(line) for line in run.stdout.splitlines()]
        self.assertEqual(responses[0]['result']['serverInfo']['name'],'endnote-word-local')
        names={t['name'] for t in responses[1]['result']['tools']}
        self.assertIn('remove_endnote_citation',names)
        self.assertIn('update_endnote_word_fields',names)
        for name in ('insert_endnote_citation_after_paragraph', 'remove_endnote_citation'):
            schema = next(t['inputSchema'] for t in responses[1]['result']['tools'] if t['name'] == name)
            self.assertIn('overwrite', schema['properties'])
            self.assertFalse(schema['properties']['overwrite']['default'])
            self.assertNotIn('output_path', schema['required'])
        self.assertEqual(responses[2]['result']['structuredContent']['references'][0]['rec_number'],55)

    def test_source_or_existing_output_rejected_before_word(self):
        with patch.object(local,'_invoke_local') as word:
            with self.assertRaisesRegex(ValueError,'separate'):
                local.edit_citation({'input_path':'same.docx','output_path':'same.docx'},'remove')
        word.assert_not_called()

    def test_stored_split_field_instructions_are_parsed(self):
        # Minimal synthetic OOXML exercises split instruction runs, not a deliverable citation.
        path=self.root/'fields.docx'
        code=' ADDIN EN.CITE <EndNote><Cite><RecNum>55</RecNum><Author>Liu</Author><Year>2025</Year><record><titles><title>Target paper</title></titles></record></Cite></EndNote>'
        from xml.sax.saxutils import escape
        xml=f'<w:document xmlns:w="{local.NS["w"]}"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>{escape(code[:20])}</w:instrText></w:r><w:r><w:instrText>{escape(code[20:])}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>[1]</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>'
        with ZipFile(path,'w') as z: z.writestr('word/document.xml',xml)
        report=local.inspect_document(path)
        self.assertEqual(report['citations'][0]['records'][0]['title'],'Target paper')
        self.assertEqual(report['citations'][0]['result'],'[1]')


class CitationPublicationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / 'source.docx'
        self.write_document(self.source, citation=True)
        self.original = self.source.read_bytes()
        self.args = {'input_path':str(self.source), 'paragraph_text':'Target paragraph',
                     'rec_number':55, 'expected_title':'Target paper', 'overwrite':True}

    def tearDown(self):
        self.temp.cleanup()

    def write_document(self, path, citation=False, text='Target paragraph'):
        # Synthetic fields test publication guards only; native CWYW is tested separately.
        from xml.sax.saxutils import escape
        def field(code, result):
            return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                    f'<w:r><w:instrText>{escape(code)}</w:instrText></w:r>'
                    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                    f'<w:r><w:t>{escape(result)}</w:t></w:r>'
                    '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        cite = bib = ''
        if citation:
            code = ' ADDIN EN.CITE <EndNote><Cite><RecNum>55</RecNum><Author>Liu</Author><Year>2025</Year><record><titles><title>Target paper</title></titles></record></Cite></EndNote>'
            cite = field(code, '[1]')
            bib = '<w:p>' + field(' ADDIN EN.REFLIST ', 'Target paper') + '</w:p>'
        xml = f'<w:document xmlns:w="{local.NS["w"]}"><w:body><w:p><w:r><w:t>{escape(text)}</w:t></w:r>{cite}</w:p>{bib}</w:body></w:document>'
        with ZipFile(path, 'w') as z:
            z.writestr('word/document.xml', xml)

    def bridge(self, config, work):
        self.write_document(Path(config['scratch_path']))
        return {'ok':True}

    def test_in_place_publishes_validated_file_and_exact_backup(self):
        with patch.object(local, '_invoke_local', side_effect=self.bridge):
            result = local.edit_citation(self.args, 'remove')
        self.assertTrue(result['in_place'])
        self.assertFalse(result['source_unchanged'])
        self.assertEqual(result['output_path'], str(self.source))
        self.assertEqual(Path(result['backup_path']).read_bytes(), self.original)
        self.assertEqual(local.inspect_document(self.source)['citation_count'], 0)
        self.assertEqual(list(self.root.glob('*.docx')), [self.source])

    def test_explicit_same_output_is_supported(self):
        with patch.object(local, '_invoke_local', side_effect=self.bridge):
            result = local.edit_citation({**self.args, 'output_path':str(self.source)}, 'remove')
        self.assertTrue(result['in_place'])

    def test_separate_output_still_preserves_source(self):
        output = self.root / 'edited.docx'
        args = {**self.args, 'overwrite':False, 'output_path':str(output)}
        with patch.object(local, '_invoke_local', side_effect=self.bridge):
            result = local.edit_citation(args, 'remove')
        self.assertTrue(result['source_unchanged'])
        self.assertIsNone(result['backup_path'])
        self.assertEqual(self.source.read_bytes(), self.original)
        self.assertTrue(output.exists())

    def test_implicit_or_nonboolean_overwrite_is_refused_before_word(self):
        with patch.object(local, '_invoke_local') as word:
            for overwrite in (False, 'true', 1):
                with self.subTest(overwrite=overwrite), self.assertRaises(ValueError):
                    local.edit_citation({**self.args, 'overwrite':overwrite}, 'remove')
        word.assert_not_called()

    def test_overwrite_does_not_authorize_another_destination(self):
        other = self.root / 'other.docx'
        other.write_bytes(b'keep this')
        with patch.object(local, '_invoke_local') as word:
            with self.assertRaisesRegex(ValueError, 'only supports'):
                local.edit_citation({**self.args, 'output_path':str(other)}, 'remove')
        word.assert_not_called()
        self.assertEqual(other.read_bytes(), b'keep this')

    def test_bridge_failure_preserves_original_without_backup(self):
        with patch.object(local, '_invoke_local', return_value={'ok':False, 'error':'CWYW failed'}):
            self.assertFalse(local.edit_citation(self.args, 'remove')['ok'])
        self.assertEqual(self.source.read_bytes(), self.original)
        self.assertFalse((self.root / '.endnote-backups').exists())

    def test_failed_validation_preserves_original(self):
        def changed_body(config, work):
            self.write_document(Path(config['scratch_path']), text='Unexpected change')
            return {'ok':True}
        with patch.object(local, '_invoke_local', side_effect=changed_body):
            with self.assertRaisesRegex(ValueError, 'body text changed'):
                local.edit_citation(self.args, 'remove')
        self.assertEqual(self.source.read_bytes(), self.original)
        self.assertFalse((self.root / '.endnote-backups').exists())

    def test_concurrent_edit_is_not_overwritten(self):
        def concurrent_edit(config, work):
            self.bridge(config, work)
            self.write_document(self.source, citation=True, text='Target paragraph changed externally')
            return {'ok':True}
        with patch.object(local, '_invoke_local', side_effect=concurrent_edit):
            with self.assertRaises(ValueError):
                local.edit_citation(self.args, 'remove')
        self.assertEqual(local.body_content(self.source), ['Target paragraph changed externally'])

    def test_locked_destination_preserves_original_and_backup(self):
        with patch.object(local, '_invoke_local', side_effect=self.bridge), \
             patch.object(local.os, 'replace', side_effect=PermissionError('Document is open')):
            with self.assertRaises(PermissionError):
                local.edit_citation(self.args, 'remove')
        self.assertEqual(self.source.read_bytes(), self.original)
        backups = list((self.root / '.endnote-backups').glob('*.docx'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), self.original)


if __name__=='__main__': unittest.main()
