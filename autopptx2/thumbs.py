"""Cache thumbnail chạy nền — UI không bao giờ chờ decode ảnh.

Nguyên tắc: mọi decode (mở file, convert, thu nhỏ) chạy ở ThreadPool;
luồng UI chỉ resize từ thumbnail nhỏ đã có sẵn trong RAM (sub-ms).
"""
import os
import threading
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor

from PIL import Image

from . import geometry as G


class ThumbCache:
    def __init__(self, max_items=256, max_px=1024):
        self.max_items = max_items
        self.max_px = max_px
        self._cache = OrderedDict()          # path -> PIL.Image (RGB, đã nhỏ)
        self._lock = threading.Lock()
        self._inflight = set()
        workers = max(2, min(8, (os.cpu_count() or 4) // 2))
        self._pool = ThreadPoolExecutor(max_workers=workers,
                                        thread_name_prefix="thumb")

    def get(self, path):
        """Ảnh đã decode hoặc None (không chặn)."""
        with self._lock:
            im = self._cache.get(path)
            if im is not None:
                self._cache.move_to_end(path)
            return im

    def request(self, path, widget, callback):
        """Nếu có sẵn → trả ngay. Chưa có → decode nền rồi gọi
        callback(path, img) trên luồng Tk qua widget.after."""
        im = self.get(path)
        if im is not None:
            return im
        with self._lock:
            if path in self._inflight:
                return None
            self._inflight.add(path)
        self._pool.submit(self._decode, path, widget, callback)
        return None

    def _decode(self, path, widget, callback):
        im = None
        try:
            im = Image.open(path)
            try:
                im.draft('RGB', (self.max_px, self.max_px))
            except Exception:
                pass
            im = G.exif_upright(im)
            im = im.convert("RGB")
            try:
                im.info.pop('exif', None)
            except Exception:
                pass
            if max(im.size) > self.max_px:
                im.thumbnail((self.max_px, self.max_px),
                             Image.Resampling.LANCZOS)
        except Exception:
            im = None
        with self._lock:
            self._inflight.discard(path)
            if im is not None:
                self._cache[path] = im
                while len(self._cache) > self.max_items:
                    self._cache.popitem(last=False)
        if im is not None and callback is not None:
            try:
                widget.after(0, lambda: self._safe_cb(widget, callback, path, im))
            except Exception:
                pass

    @staticmethod
    def _safe_cb(widget, callback, path, im):
        try:
            if widget.winfo_exists():
                callback(path, im)
        except Exception:
            pass

    def clear(self):
        with self._lock:
            self._cache.clear()

    def shutdown(self):
        self._pool.shutdown(wait=False, cancel_futures=True)
