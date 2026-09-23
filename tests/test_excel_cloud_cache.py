import os
import tempfile
import unittest
from unittest.mock import patch

from autopptx2.cloud import _ts_equal, CloudClient, _write_json


class TestExcelCloudCache(unittest.TestCase):
    def test_ts_equal_timezone(self):
        self.assertTrue(_ts_equal(
            '2026-09-07T04:01:05.511699',
            '2026-09-07T04:01:05.511699+00:00'))
        self.assertTrue(_ts_equal(
            '2026-09-07T04:01:05.511699Z',
            '2026-09-07T04:01:05.511699+00:00'))
        self.assertFalse(_ts_equal('2026-09-07T04:01:05', '2026-09-07T04:01:06'))

    def test_display_name_from_meta(self):
        uid = 'test-user-123'
        with tempfile.TemporaryDirectory() as td:
            cache = os.path.join(td, f'{uid}.xlsx')
            with open(cache, 'wb') as f:
                f.write(b'x')
            meta_path = os.path.join(td, 'sync_meta.json')
            _write_json(meta_path, {uid: {'file_name': 'Master_GOP.xlsx'}})
            cloud = CloudClient()
            cloud.user_id = uid
            cloud.excel_meta_path = lambda: meta_path
            with patch('autopptx2.cloud.excel_cache_path', lambda u: cache if u == uid else ''):
                self.assertEqual(cloud.excel_local_display_name(cache), 'Master_GOP.xlsx')
            self.assertEqual(cloud.excel_local_display_name('/tmp/other.xlsx'), 'other.xlsx')

    def test_api_error_hides_42501(self):
        raw = ("{'message': 'permission denied for table excel_data_files', "
               "'code': '42501'}")
        msg = CloudClient._api_error(raw)
        self.assertIn('Đăng xuất', msg)
        self.assertNotIn('42501', msg)

    def test_broadcast_empty_profiles_is_failure(self):
        cloud = CloudClient()
        r = cloud.upload_excel_broadcast('x.xlsx', [])
        self.assertFalse(r.get('ok'))
        self.assertIn('tài khoản', r.get('msg', ''))


if __name__ == '__main__':
    unittest.main()
