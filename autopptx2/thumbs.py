"""Cache thumbnail chạy nền — UI không bao giờ chờ decode ảnh.

Nguyên tắc: mọi decode (mở file, convert, thu nhỏ) chạy ở ThreadPool;
luồng UI chỉ resize từ thumbnail nhỏ đã có sẵn trong RAM (sub-ms).

Nhiều widget (canvas + timeline + prefetch) có thể request cùng 1 path.
Mọi waiter đều được báo khi decode xong — không để slide trắng vì
prefetch/timeline "giữ" decode trước.
Callback được bơm trên luồng Tk, không gọi Tcl từ worker.
"""
import os
import threading
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor

from PIL import Image

from . import geometry as G


class ThumbCache:
    def __init__(self, max_items=320, max_px=1024):
        self.max_items = max_items
        self.max_px = max_px
        self._cache = OrderedDict()          # path -> PIL.Image (RGB, đã nhỏ)
        self._lock = threading.Lock()
        self._inflight = set()
        self._waiters = {}                   # path -> [(widget, callback)]
        self._ready = []                     # [(widget, callback, path, im)]
        self._pump_widget = None
        self._pump_job = None
        workers = max(2, min(8, (os.cpu_count() or 4) // 2))
        self._pool = ThreadPoolExecutor(max_workers=workers,
                                        thread_name_prefix="thumb")

    def bind_ui(self, widget):
        """Widget Tk sống lâu — bơm callback an toàn trên luồng UI."""
        self._pump_widget = widget
        self._ensure_pump()

    def get(self, path):
        """Ảnh đã decode hoặc None (không chặn)."""
        if not path:
            return None
        with self._lock:
            im = self._cache.get(path)
            if im is not None:
                self._cache.move_to_end(path)
            return im

    def request(self, path, widget, callback):
        """Nếu có sẵn → trả ngay. Chưa có → decode nền rồi gọi
        callback(path, img) trên luồng Tk."""
        if not path:
            return None
        im = self.get(path)
        if im is not None:
            return im
        start = False
        with self._lock:
            im = self._cache.get(path)
            if im is not None:
                self._cache.move_to_end(path)
                return im
            if callback is not None and widget is not None:
                self._waiters.setdefault(path, []).append((widget, callback))
            if path not in self._inflight:
                self._inflight.add(path)
                start = True
        if start:
            self._pool.submit(self._decode, path)
        self._ensure_pump()
        return None

    def _decode(self, path):
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
            waiters = self._waiters.pop(path, [])
            if im is not None:
                self._cache[path] = im
                while len(self._cache) > self.max_items:
                    self._cache.popitem(last=False)
            for widget, callback in waiters:
                if callback is not None:
                    self._ready.append((widget, callback, path, im))

    def _ensure_pump(self):
        w = self._pump_widget
        if w is None or self._pump_job is not None:
            return
        try:
            if not w.winfo_exists():
                return
            self._pump_job = w.after(20, self._pump)
        except Exception:
            self._pump_job = None

    def _pump(self):
        self._pump_job = None
        batch = []
        more = False
        with self._lock:
            if self._ready:
                batch = self._ready
                self._ready = []
            more = bool(self._inflight or self._waiters or self._ready)
        for widget, callback, path, im in batch:
            if im is None:
                continue
            self._safe_cb(widget, callback, path, im)
        if more:
            self._ensure_pump()

    def drain_ready(self):
        """Dùng cho test — lấy callback đã sẵn, không cần Tk."""
        with self._lock:
            batch = self._ready
            self._ready = []
            return batch

    @staticmethod
    def _safe_cb(widget, callback, path, im):
        try:
            if widget is None or widget.winfo_exists():
                callback(path, im)
        except Exception:
            pass

    def clear(self):
        with self._lock:
            self._cache.clear()
            self._waiters.clear()
            self._ready.clear()

    def shutdown(self):
        try:
            if self._pump_job and self._pump_widget is not None:
                self._pump_widget.after_cancel(self._pump_job)
        except Exception:
            pass
        self._pump_job = None
        self._pool.shutdown(wait=False, cancel_futures=True)
