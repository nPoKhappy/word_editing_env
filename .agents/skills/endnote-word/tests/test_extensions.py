"""Regression checks for adapters; all fixtures are synthetic and remain local."""
from contextlib import closing
import json
from pathlib import Path
import sqlite3
import struct
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import endnote_mcp as server
import endnote_integrations as adapter
import endnote_metadata as metadata
from endnote_cli.core.models import REFS_SELECT_FIELDS


class ExtensionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.enl = self.root / 'Library.enl'
        self.enl.write_bytes(b'fixture library stub')
        self.data = self.root / 'Library.Data'
        (self.data / 'sdb').mkdir(parents=True)
        (self.data / 'PDF/empty').mkdir(parents=True)
        self.db = self.data / 'sdb/sdb.eni'
        with closing(sqlite3.connect(self.db)) as db, db:
            fields = list(dict.fromkeys([*REFS_SELECT_FIELDS, *metadata.lib.SELECT_COLUMNS]))
            db.execute('CREATE TABLE refs (' + ','.join('"' + f + '" ' + (
                'INTEGER' if f in ('id', 'reference_type', 'trash_state', 'added_to_library', 'record_last_updated')
                else 'TEXT') for f in fields) + ')')
            for n, trash in ((1, 0), (2, 0), (3, 1)):
                db.execute('INSERT INTO refs(id,reference_type,trash_state,author,year,title,secondary_title,abstract) '
                           'VALUES(?,0,?,"LeCun, Yann", "2015", "Deep learning", "Nature", ?)',
                           (n, trash, 'long abstract ' * 40))
            db.executescript('CREATE TABLE groups(group_id INTEGER,spec BLOB,members BLOB);'
                             'CREATE TABLE misc(code INTEGER,subcode INTEGER,value TEXT);'
                             'CREATE TABLE tag_groups(group_id INTEGER,spec BLOB);'
                             'CREATE TABLE tag_members(tag_ids TEXT);'
                             'CREATE TABLE file_res(refs_id INTEGER,file_path TEXT,file_type INTEGER,file_pos INTEGER);'
                             'CREATE TABLE jterms(list_id INTEGER,term TEXT,abbr1 TEXT,abbr2 TEXT,abbr3 TEXT);')
            spec = b'<group><ids><name>Review</name><id>g9</id></ids><rules><rule>TYPE;3</rule></rules></group>'
            db.execute('INSERT INTO groups VALUES(9,?,?)', (spec, b'\0\0\0\2' + struct.pack('<III', 2, 1, 3)))
            db.execute('INSERT INTO misc VALUES(17,1,?)',
                       ('<group><ids><name>My Set</name><id>s1</id></ids><member>g9</member></group>',))
            db.execute('INSERT INTO tag_groups VALUES(1,?)',
                       (b'<tag><name>Read</name><rule>COLOR;ff0000</rule></tag>',))
            db.executemany('INSERT INTO tag_members(rowid,tag_ids) VALUES(?,?)', [(1, '1'), (3, '1')])
            db.executemany('INSERT INTO jterms VALUES(2,?,?,?,?)',
                           [('Nature', '', '', ''), ('Nature', 'Nature', '', '')])

    def tearDown(self):
        self.temp.cleanup()

    def args(self, **kw):
        return {'library_path': str(self.enl), **kw}

    def test_catalog_exposes_extensions_once_and_labels_network(self):
        names = [t['name'] for t in server.base.TOOLS]
        self.assertEqual(len(names), len(set(names)))
        for name in server.METADATA_ROUTES:
            self.assertIn(name, names)
        online = server.EXTRA_BY_NAME['verify_endnote_reference_online']
        self.assertTrue(online['annotations']['openWorldHint'])

    def test_invalid_args_are_refused_before_network(self):
        with patch.object(adapter, 'verify_online') as network:
            for values in ({'rec_number': True}, {'rec_number': 0}, {'rec_number': 1, 'unexpected': 1}):
                result = server.call('verify_endnote_reference_online', self.args(**values))
                self.assertTrue(result['isError'])
        network.assert_not_called()

    def test_duplicate_candidates_exclude_trash(self):
        result = metadata.duplicates(self.args())
        self.assertEqual(result['duplicate_candidates'][0]['rec_numbers'], [1, 2])

    def test_complete_abstract_and_type_are_preserved(self):
        result = adapter.get_full_record(self.args(rec_number=1))
        self.assertGreater(len(result['reference']['abstract']), 300)
        self.assertEqual(result['reference']['reference_type_name'], 'Journal Article')

    def test_full_record_rejects_trash(self):
        with self.assertRaises(ValueError):
            adapter.get_full_record(self.args(rec_number=3))

    def test_tag_members_exclude_trash(self):
        self.assertEqual(adapter.list_tags(self.args())['tags'][0]['rec_numbers'], [1])

    def test_group_hierarchy_uses_validated_membership_codec(self):
        result = adapter.group_tree(self.args())
        self.assertEqual(result['groups'][0]['member_count'], 2)
        self.assertEqual(result['groups'][0]['group_set_id'], 1)
        self.assertNotIn('members', result['groups'][0])

    def test_term_merge_fills_unique_blank_without_modifying_library(self):
        before = self.db.read_bytes()
        output = self.root / 'terms.txt'
        result = metadata.prepare_terms(self.args(output_path=str(output)))
        self.assertEqual(result['unique_count'], 1)
        self.assertEqual(output.read_bytes(), b'Nature\tNature\t\t\r\n')
        self.assertEqual(self.db.read_bytes(), before)

    def test_term_conflict_does_not_publish_incomplete_import(self):
        with closing(sqlite3.connect(self.db)) as db, db:
            db.execute('INSERT INTO jterms VALUES(2,"Nature","Nat.","","")')
        output = self.root / 'terms.txt'
        result = metadata.prepare_terms(self.args(output_path=str(output)))
        self.assertFalse(result['ready_for_native_import'])
        self.assertFalse(output.exists())

    def test_export_selected_group_excludes_trash(self):
        output = self.root / 'catalog.json'
        result = adapter.export_catalog(self.args(group='Review', format='json', output_path=str(output)))
        self.assertEqual(result['count'], 1)
        self.assertIn('Deep learning', output.read_text(encoding='utf-8'))

    def test_export_cannot_overwrite_source(self):
        before = self.enl.read_bytes()
        with self.assertRaises(FileExistsError):
            adapter.export_catalog(self.args(format='json', output_path=str(self.enl)))
        self.assertEqual(self.enl.read_bytes(), before)

    def test_missing_export_record_is_not_silently_ignored(self):
        with self.assertRaises(ValueError):
            adapter.export_catalog(self.args(rec_numbers=[99], format='csv', output_path=str(self.root / 'out.csv')))

    def test_attachment_path_traversal_is_refused(self):
        with closing(sqlite3.connect(self.db)) as db, db:
            db.execute('INSERT INTO file_res VALUES(1,"../../outside.pdf",1,0)')
        result = metadata.attachments(self.args())
        self.assertTrue(result['attachments'][0]['unsafe_relative_path'])
        self.assertFalse(result['attachments'][0]['exists'])

    def test_backup_preserves_files_empty_folders_and_hashes(self):
        out = self.root / 'backup'
        result = metadata.backup_library(self.args(output_directory=str(out)))
        self.assertTrue(result['ok'])
        self.assertEqual((out / self.enl.name).read_bytes(), self.enl.read_bytes())
        self.assertEqual((out / 'Library.Data/sdb/sdb.eni').read_bytes(), self.db.read_bytes())
        self.assertTrue((out / 'Library.Data/PDF/empty').is_dir())
        self.assertEqual(len(json.loads((out / 'backup-manifest.json').read_text())['sha256']), 2)

    def test_backup_will_not_copy_into_source_data(self):
        with self.assertRaises(ValueError):
            metadata.backup_library(self.args(output_directory=str(self.data / 'backup')))

    def test_verified_restore_preserves_library_pair_and_empty_folders(self):
        backup = self.root / 'backup'
        metadata.backup_library(self.args(output_directory=str(backup)))
        restored = self.root / 'restored'
        result = metadata.restore_backup({'backup_directory': str(backup), 'output_directory': str(restored)})
        self.assertTrue(result['ok'])
        self.assertEqual((restored / 'Library.enl').read_bytes(), self.enl.read_bytes())
        self.assertTrue((restored / 'Library.Data/PDF/empty').is_dir())

    def test_corrupt_backup_is_not_restored(self):
        backup = self.root / 'backup'
        metadata.backup_library(self.args(output_directory=str(backup)))
        (backup / 'Library.enl').write_bytes(b'corrupted')
        restored = self.root / 'restored'
        with self.assertRaises(ValueError):
            metadata.restore_backup({'backup_directory': str(backup), 'output_directory': str(restored)})
        self.assertFalse(restored.exists())

    def test_search_snapshot_is_local_and_full_abstract_is_preserved(self):
        state = self.root / 'state'
        before = self.db.read_bytes()
        with patch.object(adapter, 'STATE', state):
            info = adapter.create_search_snapshot(str(self.enl))
            self.assertEqual(info['record_count'], 2)
            self.assertFalse(adapter.index_status({})['stale'])
            self.assertIn('long abstract ' * 40, (state / 'library-search-only.xml').read_text())
            self.assertEqual(self.db.read_bytes(), before)

    def test_search_status_detects_changed_library(self):
        with patch.object(adapter, 'STATE', self.root / 'state'):
            adapter.create_search_snapshot(str(self.enl))
            with closing(sqlite3.connect(self.db)) as db, db:
                db.execute('UPDATE refs SET title="Changed" WHERE id=1')
            self.assertTrue(adapter.index_status({})['stale'])

    def test_citecheck_adapter_preserves_real_record_identity_and_source(self):
        before = self.db.read_bytes()
        def external(command, **kwargs):
            payload = json.loads(Path(command[-1]).read_text())
            bib = Path(payload['input_path']).read_text()
            self.assertIn('@article{endnote1,', bib)
            self.assertIn('LeCun, Yann', bib)
            return type('Result', (), {'returncode': 0, 'stdout': '{"ok":true,"report":{"entries":[]}}', 'stderr': ''})()
        with patch.object(adapter.subprocess, 'run', side_effect=external):
            result = adapter.verify_online(self.args(rec_number=1, expected_title='Deep learning'))
        self.assertEqual(result['rec_numbers'], [1])
        self.assertFalse(result['library_modified'])
        self.assertEqual(self.db.read_bytes(), before)

    def test_citecheck_title_mismatch_is_refused_before_network(self):
        with patch.object(adapter.subprocess, 'run') as network, self.assertRaises(ValueError):
            adapter.verify_online(self.args(rec_number=1, expected_title='Unrelated work'))
        network.assert_not_called()


if __name__ == '__main__':
    unittest.main()
