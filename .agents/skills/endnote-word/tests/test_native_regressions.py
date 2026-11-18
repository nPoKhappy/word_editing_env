"""Regression coverage for native RIS, child runtime and large CWYW payload failures."""
import base64
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'scripts'))
import local_endnote as local
import import_verified_reference as importer
import endnote_metadata as metadata

class NativeRegressionTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
        self.found={'title':'Verified preprint','authors':['Example, Alice'],'year':'2023',
                    'journal':'','volume':'','issue':'','pages':'','doi':'10.1234/test',
                    'url':'https://doi.org/10.1234/test','publication_type':'Preprint'}
    def tearDown(self):self.tmp.cleanup()
    def test_missing_library_preview_supports_preprint_and_preserves_ris_end(self):
        with patch.object(importer,'resolve',return_value={'ok':True,'metadata':self.found}), \
             patch.object(importer.integrations,'TOOLS',self.root), \
             patch.object(local.upstream,'_invoke_bridge') as native:
            r=importer.import_reference({'library_path':str(self.root/'new.enl'),'preview_only':True})
        self.assertTrue(r['ok']);native.assert_not_called()
        data=Path(r['ris_path']).read_bytes()
        self.assertTrue(data.startswith(b'TY  - UNPB\r\n'))
        self.assertTrue(data.endswith(b'ER  - \r\n'))
    def test_missing_library_without_template_fails_before_native_import(self):
        with patch.object(importer,'resolve',return_value={'ok':True,'metadata':self.found}), \
             patch.object(importer.integrations,'TOOLS',self.root), \
             patch.object(local.upstream,'_invoke_bridge') as native:
            r=importer.import_reference({'library_path':str(self.root/'new.enl')})
        self.assertFalse(r['ok']);self.assertEqual(r['status'],'needs_empty_library_template')
        self.assertFalse(r['library_modified']);native.assert_not_called()
    def test_native_nested_flddata_with_base64_whitespace_and_null_terminator(self):
        payload=b'<EndNote><Cite><RecNum>8</RecNum><Author>Example</Author><Year>2023</Year><record><titles><title>Verified preprint</title></titles></record></Cite></EndNote>\0'
        encoded=base64.b64encode(payload).decode();encoded=encoded[:20]+'\n'+encoded[20:]
        def begin(data=''):return '<w:r><w:fldChar w:fldCharType="begin">'+(f'<w:fldData>{data}</w:fldData>' if data else '')+'</w:fldChar></w:r>'
        def code(t):return f'<w:r><w:instrText>{t}</w:instrText></w:r>'
        end='<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        sep='<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        xml=f'<w:document xmlns:w="{local.NS["w"]}"><w:body><w:p>'+begin()+code(' ADDIN EN.CITE ')+begin(encoded)+code(' ADDIN EN.CITE.DATA ')+sep+end+sep+'<w:r><w:t>[1]</w:t></w:r>'+end+'</w:p></w:body></w:document>'
        p=self.root/'fixture.docx'
        with ZipFile(p,'w') as z:z.writestr('word/document.xml',xml)
        r=local.inspect_document(p)
        self.assertEqual(r['citation_count'],1)
        self.assertEqual(r['citations'][0]['records'][0]['rec_number'],8)
        self.assertEqual(r['citations'][0]['result'],'[1]')
    def test_bridge_uses_helper_and_saves_owned_copy_without_saveas(self):
        s=(ROOT/'vendor/endnote-for-word/scripts/endnote-word.ps1').read_text(encoding='utf-8-sig')
        save=s.split('function Save-DocumentAtomically {',1)[1].split('\nfunction ',1)[0]
        self.assertIn('$script:Document.Save()',save)
        self.assertNotIn('$script:Document.SaveAs2(',save)
        self.assertIn('$script:WorkingCopyPath',save)
        self.assertIn('Risxtd\\Risweb32.exe',s)
        self.assertIn('Start-Process -FilePath $directExportHelper',s)
        self.assertIn('$risEndCount -ne $risRecordCount',s)
    def test_child_python_defaults_to_current_interpreter(self):
        s=(ROOT/'scripts/endnote_mcp.py').read_text(encoding='utf-8-sig')
        self.assertIn("os.environ.setdefault('ENDNOTE_WORD_PYTHON', sys.executable)",s)

if __name__=='__main__':unittest.main()
