import os
import tempfile
import time
import unittest

from PIL import Image

from autopptx2.thumbs import ThumbCache


def _png(folder, name='a.png', color=(20, 80, 160)):
    path = os.path.join(folder, name)
    Image.new('RGB', (64, 48), color).save(path)
    return path


class _Dummy:
    def winfo_exists(self):
        return True


def _wait_cache(cache, path, timeout=3):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if cache.get(path) is not None:
            return True
        time.sleep(0.02)
    return False


class TestThumbWaiters(unittest.TestCase):
    def test_inflight_keeps_later_waiters(self):
        """Prefetch (callback=None) đang decode — request sau vẫn được xếp hàng."""
        cache = ThumbCache()
        try:
            path = os.path.join(tempfile.gettempdir(), 'never-open-this.png')
            with cache._lock:
                cache._inflight.add(path)
            cache.request(path, _Dummy(), lambda p, im: None)
            with cache._lock:
                self.assertIn(path, cache._waiters)
                self.assertEqual(len(cache._waiters[path]), 1)
                cache._inflight.discard(path)
                cache._waiters.pop(path, None)
        finally:
            cache.shutdown()

    def test_two_widgets_same_path(self):
        cache = ThumbCache()
        a, b = [], []
        try:
            with tempfile.TemporaryDirectory() as td:
                path = _png(td, 'b.png', (200, 30, 30))
                cache.request(path, _Dummy(), lambda p, im: a.append(im))
                cache.request(path, _Dummy(), lambda p, im: b.append(im))
                deadline = time.time() + 3
                while time.time() < deadline and not (a and b):
                    for _w, cb, p, im in cache.drain_ready():
                        cb(p, im)
                    time.sleep(0.02)
                self.assertTrue(a and b)
        finally:
            cache.shutdown()

    def test_cache_hit_after_prefetch(self):
        cache = ThumbCache()
        try:
            with tempfile.TemporaryDirectory() as td:
                path = _png(td)
                self.assertIsNone(cache.request(path, _Dummy(), None))
                self.assertTrue(_wait_cache(cache, path))
                hit = cache.request(path, _Dummy(), lambda *_: None)
                self.assertIsNotNone(hit)
        finally:
            cache.shutdown()
