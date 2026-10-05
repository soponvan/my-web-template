import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import app as farm_app


class DatabaseBackupTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_manager = farm_app.manager
        data_dir = Path(self.temp_dir.name)
        paths = {
            'DATA_DIR': data_dir,
            'SHEEP_FILE': data_dir / 'sheep_records.json',
            'FEEDING_FILE': data_dir / 'feeding_records.json',
            'HEALTH_FILE': data_dir / 'health_records.json',
            'BREEDING_FILE': data_dir / 'breeding_records.json',
            'FEED_STOCK_FILE': data_dir / 'feed_stock.json',
            'USERS_FILE': data_dir / 'users.json',
            'DATABASE_URL': None,
            'FARM_DATA_FILES': {
                'sheep_records': data_dir / 'sheep_records.json',
                'feeding_records': data_dir / 'feeding_records.json',
                'health_records': data_dir / 'health_records.json',
                'breeding_records': data_dir / 'breeding_records.json',
                'feed_stock': data_dir / 'feed_stock.json',
            },
        }
        self.patches = [patch.object(farm_app, name, value) for name, value in paths.items()]
        for patcher in self.patches:
            patcher.start()
        self.manager = farm_app.SheepFarmManager()
        self.original_data = {
            'sheep_records': {'S001': {'id': 'S001'}},
            'feeding_records': {},
            'health_records': {},
            'breeding_records': {},
            'feed_stock': {},
        }
        self.manager.replace_all_data(self.original_data)
        farm_app.manager = self.manager
        self.client = farm_app.app.test_client()

    def tearDown(self):
        farm_app.manager = self.original_manager
        for patcher in reversed(self.patches):
            patcher.stop()
        self.temp_dir.cleanup()

    def set_admin_session(self, role='admin'):
        with self.client.session_transaction() as session:
            session['user_id'] = 'test-user'
            session['role'] = role
            session['display_name'] = 'Test User'

    def make_backup(self, data=None):
        return {
            'format': farm_app.BACKUP_FORMAT,
            'version': 1,
            'exported_at': '2026-10-05T00:00:00+00:00',
            'data': self.original_data if data is None else data,
        }

    def test_admin_can_export_versioned_backup(self):
        self.set_admin_session()

        response = self.client.get('/api/admin/database/export')

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, 'application/json')
        self.assertEqual(response.headers['Cache-Control'], 'private, no-store')
        backup = json.loads(response.get_data(as_text=True))
        self.assertEqual(backup['format'], farm_app.BACKUP_FORMAT)
        self.assertEqual(backup['version'], 1)
        self.assertEqual(backup['data'], self.original_data)
        self.assertNotIn('users', backup['data'])

    def test_admin_page_contains_backup_and_restore_controls(self):
        self.set_admin_session()

        response = self.client.get('/')

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn('menu-database', page)
        self.assertIn('database-backup-file', page)
        self.assertIn('exportDatabaseBackup()', page)
        self.assertIn('importDatabaseBackup()', page)

    def test_import_replaces_records_and_persists_them(self):
        self.set_admin_session()
        restored_data = {
            'sheep_records': {'S002': {'id': 'S002'}},
            'feeding_records': {'F001': {'quantity': 2}},
            'health_records': {'S002': [{'type': 'ตรวจสุขภาพ'}]},
            'breeding_records': {},
            'feed_stock': {'หญ้า': {'quantity': 30}},
        }
        payload = json.dumps(self.make_backup(restored_data), ensure_ascii=False).encode('utf-8')

        response = self.client.post(
            '/api/admin/database/import',
            data={'backup_file': (io.BytesIO(payload), 'backup.json')},
            content_type='multipart/form-data',
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.manager.export_data(), restored_data)
        reloaded = farm_app.SheepFarmManager()
        self.assertEqual(reloaded.export_data(), restored_data)

    def test_invalid_backup_does_not_change_existing_data(self):
        self.set_admin_session()
        invalid = self.make_backup({'sheep_records': {'bad': []}})
        payload = json.dumps(invalid).encode('utf-8')

        response = self.client.post(
            '/api/admin/database/import',
            data={'backup_file': (io.BytesIO(payload), 'bad.json')},
            content_type='multipart/form-data',
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(self.manager.export_data(), self.original_data)

    def test_file_import_rolls_back_when_replacing_a_file_fails(self):
        restored_data = {
            'sheep_records': {'S004': {'id': 'S004'}},
            'feeding_records': {},
            'health_records': {},
            'breeding_records': {},
            'feed_stock': {},
        }
        original_replace = os.replace
        replace_count = 0

        def fail_second_replace(source, destination):
            nonlocal replace_count
            replace_count += 1
            if replace_count == 2:
                raise OSError('file replacement failed')
            return original_replace(source, destination)

        with patch.object(farm_app.os, 'replace', side_effect=fail_second_replace):
            with self.assertRaisesRegex(OSError, 'file replacement failed'):
                self.manager.replace_all_data(restored_data)

        self.assertEqual(self.manager.export_data(), self.original_data)
        reloaded = farm_app.SheepFarmManager()
        self.assertEqual(reloaded.export_data(), self.original_data)

    def test_unsupported_version_does_not_change_existing_data(self):
        self.set_admin_session()
        backup = self.make_backup()
        backup['version'] = 99
        payload = json.dumps(backup).encode('utf-8')

        response = self.client.post(
            '/api/admin/database/import',
            data={'backup_file': (io.BytesIO(payload), 'future.json')},
            content_type='multipart/form-data',
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(self.manager.export_data(), self.original_data)

    def test_postgres_import_writes_all_categories_in_one_transaction(self):
        restored_data = {
            'sheep_records': {'S003': {'id': 'S003'}},
            'feeding_records': {},
            'health_records': {},
            'breeding_records': {},
            'feed_stock': {},
        }
        connection = MagicMock()
        connection.__enter__.return_value = connection
        cursor = MagicMock()
        connection.cursor.return_value.__enter__.return_value = cursor

        with patch.object(farm_app, 'DATABASE_URL', 'postgresql://database'), patch.object(
            self.manager, 'get_database_connection', return_value=connection
        ):
            self.manager.replace_all_data(restored_data)

        self.assertEqual(cursor.execute.call_count, 5)
        self.assertEqual(self.manager.export_data(), restored_data)
        self.assertTrue(connection.__exit__.called)

    def test_postgres_failure_keeps_in_memory_data_unchanged(self):
        restored_data = {
            'sheep_records': {'S003': {'id': 'S003'}},
            'feeding_records': {},
            'health_records': {},
            'breeding_records': {},
            'feed_stock': {},
        }
        connection = MagicMock()
        connection.__enter__.return_value = connection
        cursor = MagicMock()
        cursor.execute.side_effect = [None, RuntimeError('database write failed')]
        connection.cursor.return_value.__enter__.return_value = cursor

        with patch.object(farm_app, 'DATABASE_URL', 'postgresql://database'), patch.object(
            self.manager, 'get_database_connection', return_value=connection
        ):
            with self.assertRaisesRegex(RuntimeError, 'database write failed'):
                self.manager.replace_all_data(restored_data)

        self.assertEqual(self.manager.export_data(), self.original_data)
        self.assertTrue(connection.__exit__.called)

    def test_import_and_export_require_admin(self):
        self.set_admin_session(role='user')

        export_response = self.client.get('/api/admin/database/export')
        import_response = self.client.post('/api/admin/database/import')

        self.assertEqual(export_response.status_code, 403)
        self.assertEqual(import_response.status_code, 403)

    def test_unauthenticated_user_cannot_export_or_import(self):
        export_response = self.client.get('/api/admin/database/export')
        import_response = self.client.post('/api/admin/database/import')

        self.assertEqual(export_response.status_code, 401)
        self.assertEqual(import_response.status_code, 401)


if __name__ == '__main__':
    unittest.main()
