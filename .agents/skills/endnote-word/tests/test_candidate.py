"""Offline regression tests; synthetic fields are not deliverable citations."""
import base64
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'scripts'
sys.path.insert(0, str(SOURCE))
import local_endnote as local
sys.path.insert(0, str(ROOT / 'scripts'))
import anchor_validator as validator
import bibliography_validator as bibliography
spec = importlib.util.spec_from_file_location('candidate_batch', ROOT / 'scripts' / 'batch_citations.py')
batch = importlib.util.module_from_spec(spec)
spec.loader.exec_module(batch)
W = validator.NS['w']


def payload(numbers):
    return '<EndNote>' + ''.join(
        f'<Cite><RecNum>{n}</RecNum><Author>A{n}</Author><Year>2021</Year>'
        f'<record><titles><title>Title {n}</title></titles></record></Cite>'
        for n in numbers) + '</EndNote>'


def field(numbers, kind='inline'):
    data = payload(numbers)
    encoded = base64.b64encode(data.encode()).decode()
    begin = '<w:r><w:fldChar w:fldCharType="begin">'
    if kind == 'fldData':
        begin += f'<w:fldData>{encoded}</w:fldData>'
    begin += '</w:fldChar></w:r>'
    code = ' ADDIN EN.CITE ' + (data if kind == 'inline' else '')
    nested = ''
    if kind == 'nested':
        nested = ('<w:r><w:fldChar w:fldCharType="begin">'
                  f'<w:fldData>{encoded}</w:fldData></w:fldChar></w:r>'
                  '<w:r><w:instrText> ADDIN EN.CITE.DATA </w:instrText></w:r>'
                  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                  '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    return (begin + f'<w:r><w:instrText>{escape(code)}</w:instrText></w:r>' + nested +
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>[' + ', '.join(map(str, numbers)) + ']</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def paragraph(*parts):
    return '<w:p>' + ''.join(p if p.startswith('<w:') else f'<w:r><w:t>{escape(p)}</w:t></w:r>'
                             for p in parts) + '</w:p>'


def doc(path, body, bibliography=False):
    bib = ''
    if bibliography:
        bib = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
               '<w:r><w:instrText> ADDIN EN.REFLIST </w:instrText></w:r>'
               '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
               f'<w:r><w:t>{escape(bibliography if isinstance(bibliography, str) else "Reference list")}</w:t></w:r>'
               '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    xml = f'<w:document xmlns:w="{W}"><w:body>{body}{bib}</w:body></w:document>'
    with ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', xml)


def op(action, number, anchor=None, position='after'):
    result = {'action': action, 'paragraph_text':'First Second', 'rec_number':number,
              'expected_title':f'Title {number}'}
    if anchor:
        result.update(anchor_text=anchor, position=position)
    return result


class AnchorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.a, self.b = Path(self.tmp.name)/'a.docx', Path(self.tmp.name)/'b.docx'

    def tearDown(self):
        self.tmp.cleanup()

    def test_five_records_two_merged_fields_with_nested_payloads(self):
        doc(self.a, paragraph('First Second'))
        doc(self.b, paragraph('First', field([1,2,3], 'fldData'), ' Second',
                              field([4,5], 'nested')), True)
        ops = [op('insert', n, 'First' if n <= 3 else 'First Second') for n in range(1,6)]
        self.assertEqual(validator.verify_anchors(self.a,self.b,ops)['validated_operations'],5)
        self.assertEqual(local.inspect_document(self.b)['citation_count'],2)

    def test_wrong_position_rejected_even_when_global_identities_match(self):
        doc(self.a, paragraph('First Second'))
        doc(self.b, paragraph('First', field([1,2]), ' Second'), True)
        ops = [op('insert', 1, 'First'), op('insert', 2, 'First Second')]
        with self.assertRaisesRegex(ValueError, 'positions differ'):
            validator.verify_anchors(self.a,self.b,ops)

    def test_before_move_remove_and_untouched_citation(self):
        doc(self.a, paragraph('First', field([1]), ' Second', field([3])), True)
        doc(self.b, paragraph(field([2]), 'First Second', field([1])), True)
        ops = [op('move',1,'First Second'), op('remove',3),
               op('insert',2,'First',position='before')]
        self.assertEqual(validator.verify_anchors(self.a,self.b,ops)['validated_operations'],3)

    def test_move_merged_refused(self):
        doc(self.a, paragraph('First', field([1,2]), ' Second'), True)
        doc(self.b, paragraph('First', field([2]), ' Second', field([1])), True)
        with self.assertRaisesRegex(ValueError, 'standalone'):
            validator.verify_anchors(self.a,self.b,[op('move',1,'First Second')])

    def test_changed_existing_location_refused(self):
        doc(self.a, paragraph('First', field([1]), ' Second'), True)
        doc(self.b, paragraph('First Second', field([1,2])), True)
        with self.assertRaisesRegex(ValueError, 'positions differ'):
            validator.verify_anchors(self.a,self.b,[op('insert',2,'First Second')])

    def test_distinct_needles_for_same_paragraph_are_combined(self):
        doc(self.a, paragraph('First Second'))
        doc(self.b, paragraph('First',field([1]),' Second',field([2])),True)
        one = {**op('insert',1,'First'),'paragraph_text':'First'}
        two = {**op('insert',2,'First Second'),'paragraph_text':'Second'}
        self.assertEqual(validator.verify_anchors(self.a,self.b,[one,two])['validated_paragraphs'],1)


class BatchTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root/'source.docx'
        self.out = self.root/'out.docx'
        doc(self.source,paragraph('First Second'))
        self.original = self.source.read_bytes()
        self.args = {'input_path':str(self.source), 'output_path':str(self.out),
                     'library_path':'test.enl', 'operations':[op('insert',1,'First')]}

    def tearDown(self):
        self.tmp.cleanup()

    def test_timeout_cleanup_error_preserves_failure_and_source(self):
        class FailingTemp:
            def __init__(inner, prefix, dir):
                inner.name = str(Path(dir)/'held')
                Path(inner.name).mkdir()
            def cleanup(inner):
                raise PermissionError('WinError 32')
        record = {'rec_number':1,'title':'Title 1','first_author_surname':'A1','year':'2021'}
        with patch.object(batch.tempfile,'TemporaryDirectory',FailingTemp), \
             patch.object(local.library,'get_reference',return_value={'ok':True,'reference':record}), \
             patch.object(local,'_invoke_local',return_value={'ok':False,'error':'timed out after 120s'}):
            result = batch.edit_batch(self.args)
        self.assertIn('timed out',result['bridge']['error'])
        self.assertIn('WinError 32',result['cleanup_warning']['error'])
        self.assertFalse(self.out.exists())
        self.assertEqual(self.source.read_bytes(),self.original)

    def test_timeout_validation_before_word(self):
        with patch.object(local,'_invoke_local') as bridge:
            for timeout in (0,1801,True,1.5,'120'):
                with self.subTest(timeout=timeout), self.assertRaises(ValueError):
                    batch.edit_batch({**self.args,'timeout_seconds':timeout})
        bridge.assert_not_called()

    def test_cleanup_does_not_mask_validation_exception(self):
        class FailingTemp:
            def __init__(inner, prefix, dir):
                inner.name = str(Path(dir)/'held')
                Path(inner.name).mkdir()
            def cleanup(inner):
                raise PermissionError('WinError 32')
        record = {'rec_number':1,'title':'Title 1','first_author_surname':'A1','year':'2021'}
        with patch.object(batch.tempfile,'TemporaryDirectory',FailingTemp), \
             patch.object(local.library,'get_reference',return_value={'ok':True,'reference':record}), \
             patch.object(local,'_invoke_local',return_value={'ok':True}):
            with self.assertRaisesRegex(ValueError,'identities differ.*WinError 32'):
                batch.edit_batch(self.args)
        self.assertFalse(self.out.exists())
        self.assertEqual(self.source.read_bytes(),self.original)

    def test_success_merged_native_fields_keeps_source(self):
        ops = [op('insert',n,'First') for n in (1,2,3)]
        def record(path, number):
            return {'ok':True,'reference':{'rec_number':number,'title':f'Title {number}',
                     'first_author_surname':f'A{number}','year':'2021'}}
        def bridge(config, work):
            self.assertEqual(config['timeout_seconds'],300)
            doc(Path(config['scratch_path']),paragraph('First',field([1,2,3],'fldData'),' Second'),
                '[1] A1 Title 1. [2] A2 Title 2. [3] A3 Title 3.')
            return {'ok':True}
        with patch.object(local.library,'get_reference',side_effect=record), \
             patch.object(local,'_invoke_local',side_effect=bridge):
            result = batch.edit_batch({**self.args,'operations':ops,'timeout_seconds':300})
        self.assertTrue(result['ok'])
        self.assertEqual(result['citation_count_after'],1)
        self.assertEqual(self.source.read_bytes(),self.original)
        self.assertTrue(self.out.exists())


class BibliographyTests(unittest.TestCase):
    def sample(self, result='[1, 2]', text='[1] Smith Title 1. [2] Jones Title 2.'):
        return {'citation_count':1, 'bibliography_count':1,
                'citations':[{'result':result, 'records':[
                    {'rec_number':1, 'title':'Title 1', 'author':'Smith', 'year':'2021'},
                    {'rec_number':2, 'title':'Title 2', 'author':'Jones', 'year':'2020'}]}],
                'bibliographies':[{'result':text}]}

    def test_two_citations_and_range_verified(self):
        self.assertEqual(bibliography.verify_bibliography(self.sample())['unique_references'],2)
        self.assertEqual(bibliography.verify_bibliography(self.sample(result='[1–2]'))['unique_references'],2)

    def test_stale_nonempty_bibliography_refused(self):
        with self.assertRaisesRegex(ValueError, 'entry count'):
            bibliography.verify_bibliography(self.sample(text='[1] Smith Title 1.'))

    def test_wrong_title_for_number_refused(self):
        source = self.sample(result='[1]')
        source['citations'] = [
            {'result':'[1]', 'records':[source['citations'][0]['records'][1]]},
            {'result':'[2]', 'records':[source['citations'][0]['records'][0]]}]
        with self.assertRaisesRegex(ValueError, 'different bibliography title'):
            bibliography.verify_bibliography(source)

    def test_author_date_style_explicitly_refused(self):
        with self.assertRaisesRegex(ValueError, 'author-date'):
            bibliography.verify_bibliography(self.sample(result='(Smith, 2021; Jones, 2020)'))

    def test_non_numbered_bibliography_explicitly_refused(self):
        with self.assertRaisesRegex(ValueError, 'non-numbered'):
            bibliography.verify_bibliography(self.sample(text='Smith Title 1. Jones Title 2.'))


if __name__ == '__main__':
    unittest.main()
