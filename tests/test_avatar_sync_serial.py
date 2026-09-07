import unittest
from unittest.mock import MagicMock, patch

from autopptx2.cloud import CloudClient


class AvatarSyncSerialTests(unittest.TestCase):
    def test_sync_avatars_downloads_serially_with_lock(self):
        cloud = CloudClient()
        cloud.user_id = 'u1'
        sb = MagicMock()
        cloud.client = lambda: sb
        sb.table.return_value.select.return_value.range.return_value.execute.return_value = MagicMock(
            data=[{'storage_path': 'shared/a.jpg', 'file_name': 'a.jpg', 'updated_at': '1'}]
        )
        sb.storage.from_.return_value.download.return_value = b'img'

        lock_calls = []

        real_lock = cloud._api_lock

        class TrackingLock:
            def __enter__(self):
                lock_calls.append('enter')
                return real_lock.__enter__()

            def __exit__(self, *a):
                return real_lock.__exit__(*a)

        cloud._api_lock = TrackingLock()

        with patch('autopptx2.cloud._read_json', return_value={'files': {}}), \
                patch('autopptx2.cloud._write_json'), \
                patch('autopptx2.cloud.avatar_cache_dir', return_value='/tmp/av'), \
                patch('autopptx2.cloud.os.path.exists', return_value=False), \
                patch('autopptx2.cloud.os.makedirs'), \
                patch('builtins.open', MagicMock()), \
                patch.object(CloudClient, '_count_local_avatars', return_value=1):
            r = cloud.sync_avatars()

        self.assertTrue(r.get('ok'))
        self.assertGreaterEqual(len(lock_calls), 1)


if __name__ == '__main__':
    unittest.main()
