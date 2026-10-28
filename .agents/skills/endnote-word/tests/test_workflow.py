"""Regression coverage for the reported JTICE failures and one-step import."""
import json
from pathlib import Path
import sqlite3
from contextlib import closing
import unittest
from unittest.mock import patch
import test_extensions as fixtures
import test_local as publication_fixtures
import batch_citations
import journal_diagnostics
import jtice_preflight
import import_verified_reference as importer
import local_endnote as local
import endnote_native_state as native_state


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.ExtensionTests(methodName='runTest')
        self.fixture.setUp()
        self.root, self.enl, self.db = self.fixture.root, self.fixture.enl, self.fixture.db
        self.found = {'title': 'A different public paper', 'authors': ['Example, Alice'], 'year': '2020',
                      'journal': 'Example Journal', 'volume': '3', 'issue': '2', 'pages': '12-18',
                      'doi': '10.1234/example', 'url': 'https://doi.org/10.1234/example',
                      'source_url': 'https://api.crossref.org/works/10.1234%2Fexample',
                      'publication_type': 'journal-article', 'provider': 'Crossref'}

    def tearDown(self):
        self.fixture.tearDown()

    def args(self, **kwargs):
        return {'library_path': str(self.enl), **kwargs}

    def test_same_abbreviation_does_not_hide_duplicate_terms(self):
        with closing(sqlite3.connect(self.db)) as db, db:
            db.execute('UPDATE jterms SET abbr1="Nature"')
        r = journal_diagnostics.diagnose(self.args())
        self.assertEqual(r['duplicate_titles'][0]['count'], 2)
        self.assertFalse(r['duplicate_titles'][0]['conflicting_columns'])
        self.assertIn('duplicate_full_title_terms', r['records'][0]['findings'])

    def test_stored_volume_and_rendered_style_omission_are_separated(self):
        with closing(sqlite3.connect(self.db)) as db, db:
            db.execute('UPDATE refs SET volume="521" WHERE id=1')
            db.execute('UPDATE jterms SET abbr1="Nat."')
        doc = {'citations': [{'result': '[1]', 'records': [{'rec_number': 1, 'title': 'Deep learning'}]}],
               'bibliographies': [{'result': '[1] Authors. Deep learning. Nature. 2015:436-444.'}]}
        with patch.object(local, 'inspect_document', return_value=doc):
            r = journal_diagnostics.diagnose(self.args(document_path='fixture.docx'))
        findings = r['records'][0]['findings']
        self.assertIn('rendered_full_journal_title_review_style_or_terms', findings)
        self.assertIn('volume_stored_but_not_observed_review_style_template', findings)

    def test_import_preview_creates_evidence_without_native_write(self):
        with patch.object(importer, 'resolve', return_value={'ok': True, 'metadata': self.found}), \
             patch.object(importer.integrations, 'TOOLS', self.root), \
             patch.object(local.upstream, '_invoke_bridge') as native:
            r = importer.import_reference(self.args(doi=self.found['doi'], preview_only=True))
        self.assertTrue(r['ok'])
        self.assertTrue(Path(r['ris_path']).exists())
        self.assertFalse(r['library_modified'])
        native.assert_not_called()

    def test_import_deduplicates_even_when_existing_doi_is_blank(self):
        found = {**self.found, 'title': 'Deep learning', 'authors': ['LeCun, Yann'], 'year': '2015'}
        with closing(sqlite3.connect(self.db)) as db, db:
            db.execute('UPDATE refs SET trash_state=1 WHERE id=2')
        with patch.object(importer, 'resolve', return_value={'ok': True, 'metadata': found}), \
             patch.object(local.upstream, '_invoke_bridge') as native:
            r = importer.import_reference(self.args(doi=found['doi']))
        self.assertEqual(r['status'], 'already_present')
        self.assertEqual(r['references'][0]['rec_number'], 1)
        native.assert_not_called()

    def test_wrong_doi_on_an_existing_unrelated_record_is_not_reused(self):
        with closing(sqlite3.connect(self.db)) as db, db:
            db.execute('UPDATE refs SET electronic_resource_number=? WHERE id=1', (self.found['doi'],))
        with patch.object(importer, 'resolve', return_value={'ok': True, 'metadata': self.found}):
            r = importer.import_reference(self.args(doi=self.found['doi']))
        self.assertFalse(r['ok'])
        self.assertEqual(r['status'], 'existing_record_needs_native_review')

    def test_title_ambiguity_prevents_import(self):
        candidates = [self.found, {**self.found, 'doi': '10.1234/second'}]
        with patch.object(importer.metadata, 'search_online', return_value={'ok': True, 'candidates': candidates}):
            r = importer.resolve({'title': self.found['title']})
        self.assertFalse(r['ok'])
        self.assertEqual(r['status'], 'needs_disambiguation')

    def test_title_and_doi_disagreement_is_refused(self):
        with patch.object(importer.metadata, 'lookup_doi', return_value={'ok': True, 'metadata': self.found}):
            self.assertFalse(importer.resolve({'doi': self.found['doi'], 'title': 'Unrelated title'})['ok'])

    def test_multiple_pdf_dois_require_disambiguation(self):
        with patch.object(importer, '_from_pdf', return_value=(['10.1234/one', '10.1234/two'], 'text')):
            r = importer.resolve({'pdf_path': 'fixture.pdf'})
        self.assertEqual(r['status'], 'needs_disambiguation')

    def test_unknown_native_import_is_not_retried(self):
        with patch.object(importer, 'resolve', return_value={'ok': True, 'metadata': self.found}), \
             patch.object(importer.integrations, 'TOOLS', self.root), \
             patch.object(local.upstream, '_invoke_bridge', return_value={'ok': False, 'error': 'timeout'}) as native:
            r = importer.import_reference(self.args(doi=self.found['doi']))
        self.assertFalse(r['ok'])
        self.assertEqual(r['status'], 'import_not_confirmed')
        self.assertEqual(native.call_count, 1)

    def test_native_import_is_confirmed_by_assigned_record_identity(self):
        def native(action, args):
            self.assertEqual(action, 'import-ris')
            with closing(sqlite3.connect(self.db)) as db, db:
                db.execute('INSERT INTO refs(id,trash_state,reference_type,author,year,title,secondary_title,electronic_resource_number) '
                           'VALUES(99,0,0,?,?,?,?,?)', (self.found['authors'][0], self.found['year'], self.found['title'],
                                                     self.found['journal'], self.found['doi']))
            return {'ok': True}
        with patch.object(importer, 'resolve', return_value={'ok': True, 'metadata': self.found}), \
             patch.object(importer.integrations, 'TOOLS', self.root), \
             patch.object(local.upstream, '_invoke_bridge', side_effect=native):
            r = importer.import_reference(self.args(doi=self.found['doi']))
        self.assertTrue(r['ok'])
        self.assertEqual(r['references'][0]['rec_number'], 99)
        self.assertFalse(r['jtice_bibliography_validated'])

    def _document(self):
        path = self.root / 'source.docx'
        publication_fixtures.CitationPublicationTests.write_document(self, path, citation=True)
        return path

    def _batch(self, source, **kwargs):
        return {'input_path': str(source), 'output_path': str(self.root / 'out.docx'),
                'operations': [{'action': 'move', 'paragraph_text': 'Target paragraph', 'rec_number': 55,
                                'expected_title': 'Target paper', 'anchor_text': 'Target paragraph', 'position': 'before'}], **kwargs}

    def test_batch_failure_does_not_publish_or_modify_source(self):
        source = self._document()
        before = source.read_bytes()
        with patch.object(local, '_invoke_local', return_value={'ok': False, 'error': 'second operation failed'}):
            r = batch_citations.edit_batch(self._batch(source))
        self.assertFalse(r['ok'])
        self.assertEqual(source.read_bytes(), before)
        self.assertFalse((self.root / 'out.docx').exists())

    def test_batch_duplicate_operations_are_refused_before_word(self):
        args = self._batch(self._document())
        args['operations'] *= 2
        with patch.object(local, '_invoke_local') as word, self.assertRaises(ValueError):
            batch_citations.edit_batch(args)
        word.assert_not_called()

    def test_batch_prevents_unrelated_text_changes(self):
        source = self._document()
        def native(config, work):
            publication_fixtures.CitationPublicationTests.write_document(self, Path(config['scratch_path']),
                                                                         citation=True, text='Changed unrelated text')
            return {'ok': True}
        with patch.object(local, '_invoke_local', side_effect=native), self.assertRaises(ValueError):
            batch_citations.edit_batch(self._batch(source))
        self.assertFalse((self.root / 'out.docx').exists())

    def test_preflight_translation_is_read_only_and_preserves_scope(self):
        source = self._document()
        before = source.read_bytes()
        result = jtice_preflight.inspect({'document_path': str(source), 'stage': 'translation'})
        self.assertFalse(result['document_modified'])
        self.assertIn('Translate only', result['execution_policy'])
        self.assertEqual(source.read_bytes(), before)

    def test_import_title_search_uses_title_author_and_year(self):
        with patch.object(importer.metadata, 'search_online', return_value={'candidates': [self.found]}) as search:
            result = importer.resolve({'title': self.found['title'], 'author': 'Example', 'year': 2020})
        self.assertTrue(result['ok'])
        search.assert_called_once_with({'query': self.found['title'], 'limit': 10,
                                       'match_mode': 'title', 'author': 'Example', 'year': 2020})

    def test_native_ready_checks_exact_open_path_and_same_name_ambiguity(self):
        wrong = self.root / 'elsewhere' / self.enl.name
        self.assertTrue(native_state.confirm(1, self.enl, [str(self.enl)])['ok'])
        self.assertFalse(native_state.confirm(1, self.enl, [str(wrong)])['ok'])
        self.assertFalse(native_state.confirm(1, self.enl, [str(self.enl), str(wrong)])['ok'])
        self.assertFalse(native_state.confirm(1, self.enl, [])['ok'])


if __name__ == '__main__':
    unittest.main()
