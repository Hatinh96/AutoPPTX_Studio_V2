import io
import os
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

from PIL import Image

from autopptx2 import effects as FX
from autopptx2.effects import (
    DEFAULT_FX, _preview_cache, _preview_epoch, _preview_inflight,
    _preview_key, _preview_lock, _preview_ready, _preview_waiters,
    invalidate_preview_paths, process_photo, request_preview, stamp_minimap,
)


class TestMinimapStamp(unittest.TestCase):
    def test_offline_stamp_paints_corner(self):
        im = Image.new('RGBA', (400, 300), (20, 20, 20, 255))
        out = stamp_minimap(im, 10.7769, 106.7009, allow_net=False)
        self.assertEqual(out.size, im.size)
        self.assertNotEqual(out.getpixel((16, out.height - 24))[:3], (20, 20, 20))

    @patch('autopptx2.effects.resolve_gps', return_value=(10.7769, 106.7009))
    def test_preview_pipeline_stamps_without_network(self, _gps):
        src = Image.new('RGBA', (240, 180), (20, 20, 20, 255))
        fx = dict(DEFAULT_FX)
        fx['minimap'] = True
        fx['use_timestamp'] = False
        im = process_photo('no-net.jpg', src, fx, box=(240, 180), preview=True)
        self.assertIsNotNone(im)
        self.assertNotEqual(im.getpixel((16, im.height - 24))[:3], (20, 20, 20))


class TestBlockedOsmFallback(unittest.TestCase):
    def setUp(self):
        FX._tile_fail.clear()
        FX._geo_fail.clear()

    def tearDown(self):
        FX._tile_fail.clear()
        FX._geo_fail.clear()

    def test_tile_uses_mirror_when_osm_blocked(self):
        buf = io.BytesIO()
        Image.new('RGB', (256, 256), (9, 99, 199)).save(buf, 'PNG')
        png = buf.getvalue()
        seen = []

        class _Resp(io.BytesIO):
            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

        def fake_open(req, timeout=0):
            seen.append(req.full_url)
            if 'tile.openstreetmap.org' in req.full_url:
                raise ConnectionResetError(10054, 'forcibly closed')
            return _Resp(png)

        with tempfile.TemporaryDirectory() as d, \
                patch.object(FX, 'MAP_CACHE_DIR', d), \
                patch.object(FX, 'V1_MAP_CACHE', d), \
                patch('urllib.request.urlopen', side_effect=fake_open):
            tile = FX._fetch_osm_tile(16, 52201, 30794)
            self.assertTrue(os.path.exists(os.path.join(d, '16_52201_30794.png')))
        self.assertEqual(tile.getpixel((0, 0)), (9, 99, 199))
        self.assertIn('openstreetmap.fr', seen[-1])

    def test_partial_map_is_not_cached(self):
        tile = Image.new('RGB', (256, 256), (9, 99, 199))
        calls = {'n': 0}

        def half(z, x, y, allow_net=True):
            calls['n'] += 1
            return tile if calls['n'] % 2 else None

        key = (round(10.1, 5), round(106.1, 5), 64, 48)
        with patch.object(FX, '_fetch_osm_tile', side_effect=half):
            FX.compose_minimap(10.1, 106.1, 64, 48, allow_net=False)
        with FX._minimap_lock:
            self.assertNotIn(key, FX._minimap_cache)
        with patch.object(FX, '_fetch_osm_tile', return_value=tile):
            FX.compose_minimap(10.1, 106.1, 64, 48, allow_net=False)
        with FX._minimap_lock:
            self.assertIn(key, FX._minimap_cache)
            FX._minimap_cache.pop(key, None)

    def test_reverse_falls_back_to_photon(self):
        photon = {'features': [{'properties': {
            'name': 'Pasta Fresca', 'street': 'Tạ Hiện',
            'locality': 'Khu dân cư Thạnh Mỹ Lợi', 'district': 'Cát Lái',
            'city': 'Thành phố Hồ Chí Minh', 'country': 'Việt Nam'}}]}

        def fake_json(source, _url):
            return None if source == 'nominatim' else photon

        with patch.object(FX, '_geo_json', side_effect=fake_json):
            name = FX._reverse_geocode(10.7757, 106.7505)
        self.assertEqual(
            name.split('\n'),
            ['Tạ Hiện', 'Khu dân cư Thạnh Mỹ Lợi', 'Cát Lái', 'Thành phố Hồ Chí Minh'])

    def test_search_falls_back_to_photon(self):
        photon = {'features': [{'geometry': {'coordinates': [106.76, 10.85]}}]}

        def fake_json(source, _url):
            return [] if source == 'nominatim' else photon

        with patch.object(FX, '_geo_json', side_effect=fake_json):
            self.assertEqual(FX._search_geocode('53 Võ Văn Ngân'), (10.85, 106.76))


class TestPreviewEpoch(unittest.TestCase):
    def setUp(self):
        self._saved = (
            list(_preview_cache.items()),
            set(_preview_inflight),
            dict(_preview_epoch),
            list(_preview_ready),
            {k: list(v) for k, v in _preview_waiters.items()},
        )
        with _preview_lock:
            _preview_cache.clear()
            _preview_inflight.clear()
            _preview_epoch.clear()
            _preview_ready.clear()
            _preview_waiters.clear()

    def tearDown(self):
        cache, inflight, epoch, ready, waiters = self._saved
        with _preview_lock:
            _preview_cache.clear()
            _preview_cache.update(cache)
            _preview_inflight.clear()
            _preview_inflight.update(inflight)
            _preview_epoch.clear()
            _preview_epoch.update(epoch)
            _preview_ready[:] = ready
            _preview_waiters.clear()
            _preview_waiters.update(waiters)

    def test_stale_preview_does_not_overwrite_fresh_map(self):
        path = 'store.jpg'
        fx = {'minimap': True}
        src = Image.new('RGB', (80, 60), (1, 2, 3))
        started = threading.Event()
        release = threading.Event()
        finished = threading.Event()

        def slow(*_a, **_k):
            started.set()
            release.wait(3)
            finished.set()
            return Image.new('RGBA', (8, 8), (255, 0, 0, 255))

        with patch('autopptx2.effects.process_photo', side_effect=slow):
            request_preview(path, src, fx, 'fill', 0, 1.0)
            self.assertTrue(started.wait(3))
            invalidate_preview_paths([path])
            key = _preview_key(path, fx, 'fill', 0, 1.0)
            fresh = Image.new('RGBA', (8, 8), (0, 255, 0, 255))
            with _preview_lock:
                _preview_cache[key] = fresh
            release.set()
            self.assertTrue(finished.wait(3))
            time.sleep(0.3)
            with _preview_lock:
                self.assertIs(_preview_cache.get(key), fresh)
