"""Đóng dấu ngày/giờ, GPS mini-map, tự chỉnh sáng — không phụ thuộc Tkinter.

Dùng chung preview và export (WYSIWYG). Worker thread chỉ nhận dict cấu hình.
"""
import os
import io
import json
import math
import time
import datetime
import hashlib
import threading
from collections import OrderedDict
from functools import lru_cache

from PIL import Image, ImageDraw, ImageFont, ImageColor, ImageEnhance, ImageOps, ImageStat, ImageFilter

from .constants import CONFIG_DIR
from . import geometry as G
from . import fonts as FN

OSM_ZOOM = 16
OSM_TILE_URL = "https://tile.openstreetmap.org/{z}/{x}/{y}.png"
MAP_CACHE_DIR = os.path.join(CONFIG_DIR, "map_cache")
V1_MAP_CACHE = os.path.join(os.path.expanduser("~"), ".autopptx_studio", "map_cache")
GEO_CACHE_PATH = os.path.join(CONFIG_DIR, "geocode_cache.json")
OSM_UA = "AutoPPTXStudio/2.0 (desktop; sales report)"

DEFAULT_FX = {
    'use_timestamp': True,
    'font_scale': 3.5,
    'text_color': '#FF9900',
    'position': 'bottom_right',
    'add_background': True,
    'bg_opacity': 100,
    'location_text': '',
    'location_mode': 'gps',   # gps | auto | excel | manual
    'display_mode': 'Ngày + Giờ',
    'date_format': 'DD/MM/YYYY',
    'time_format': '24h',
    'use_exif': True,
    'show_gps': False,
    'line_spacing': 1.2,
    'letter_spacing': 0,
    'fix_date': False,
    'fix_date_val': '',
    'fix_time': False,
    'fix_time_val': '',
    'logo_enable': False,
    'logo_path': '',
    'logo_size_pct': 8,
    'logo_opacity': 50,
    'logo_position': 'bottom_right',
    'stamp_margin_pct': 3.5,
    'minimap': False,
    'minimap_opacity': 85,
    'auto_enhance': False,
    'quality_check': False,
    'dashboard': True,
    'watermark': '',
    'shadow': False,
}

_gps_cache = {}
_gps_lock = threading.Lock()
_tile_fail_ts = 0.0
_geo_fail_ts = 0.0
_geo_last_net = 0.0
_geo_mem = {}
_geo_disk = None
_geo_lock = threading.Lock()
_minimap_cache = OrderedDict()
_MINIMAP_MAX = 40
_preview_cache = OrderedDict()
_preview_lock = threading.Lock()
_PREVIEW_MAX = 32


def needed(fx):
    """Có cần bake PNG (không chèn file gốc) không."""
    if not fx:
        return False
    return bool(fx.get('use_timestamp') or fx.get('minimap') or fx.get('auto_enhance')
                or fx.get('watermark') or fx.get('shadow') or fx.get('logo_enable'))


# ════════════════════════════════════════════════════════════
#  Font
# ════════════════════════════════════════════════════════════
@lru_cache(maxsize=48)
def _font(size):
    size = max(6, int(size))
    fp, idx = FN.fallback_font_file()
    if fp:
        return FN.pil_font(fp, size, idx)
    win = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')
    for fp in (os.path.join(win, 'arial.ttf'), os.path.join(win, 'segoeui.ttf'),
               os.path.join(win, 'tahoma.ttf'), 'arial.ttf', 'DejaVuSans.ttf'):
        try:
            return ImageFont.truetype(fp, size)
        except Exception:
            continue
    return ImageFont.load_default()


# ════════════════════════════════════════════════════════════
#  GPS
# ════════════════════════════════════════════════════════════
def extract_gps(path):
    """(lat, lon) float có dấu, hoặc None."""
    try:
        with Image.open(path) as im:
            exif = im.getexif()
        if not exif:
            return None
        gps_ifd = exif.get_ifd(34853) if hasattr(exif, 'get_ifd') else exif.get(34853)
        if not gps_ifd:
            return None

        def to_deg(v):
            return float(v[0]) + float(v[1]) / 60.0 + float(v[2]) / 3600.0

        lat, lat_ref = gps_ifd.get(2), gps_ifd.get(1)
        lon, lon_ref = gps_ifd.get(4), gps_ifd.get(3)
        if not (lat and lon and lat_ref and lon_ref):
            return None
        lat_d, lon_d = to_deg(lat), to_deg(lon)
        if str(lat_ref).upper().startswith('S'):
            lat_d = -lat_d
        if str(lon_ref).upper().startswith('W'):
            lon_d = -lon_d
        return (lat_d, lon_d)
    except Exception:
        return None


def gps_cached(path):
    with _gps_lock:
        if path in _gps_cache:
            return _gps_cache[path]
    g = extract_gps(path)
    with _gps_lock:
        if len(_gps_cache) > 500:
            _gps_cache.clear()
        _gps_cache[path] = g
    return g


def gps_text(path):
    g = gps_cached(path)
    if not g:
        return ''
    lat, lon = g
    return "GPS: {} {:.4f}°, {} {:.4f}°".format(
        'N' if lat >= 0 else 'S', abs(lat),
        'E' if lon >= 0 else 'W', abs(lon))


def _geo_key(lat, lon):
    return f"{round(float(lat), 4):.4f},{round(float(lon), 4):.4f}"


def _load_geo_disk():
    global _geo_disk
    if _geo_disk is not None:
        return _geo_disk
    try:
        with open(GEO_CACHE_PATH, "r", encoding="utf-8") as f:
            _geo_disk = json.load(f) or {}
    except Exception:
        _geo_disk = {}
    return _geo_disk


def _save_geo_disk():
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        data = _load_geo_disk()
        if len(data) > 2000:
            for old in list(data)[:400]:
                data.pop(old, None)
        with open(GEO_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass


def _fmt_nominatim(addr):
    """Địa chỉ theo cấp: đường / khu → phường → quận → tỉnh (mỗi cấp 1 dòng)."""
    if not isinstance(addr, dict):
        return ""
    parts = []
    for k in ('road', 'pedestrian', 'residential', 'neighbourhood', 'quarter',
              'suburb', 'village', 'town', 'city_district', 'district',
              'county', 'city', 'state'):
        v = str(addr.get(k) or '').strip()
        if v and v not in parts and v.lower() not in ('vietnam', 'việt nam'):
            parts.append(v)
        if len(parts) >= 4:
            break
    return '\n'.join(parts)


def _nominatim_reverse(lat, lon):
    global _geo_fail_ts, _geo_last_net
    if time.time() - _geo_fail_ts < 60:
        return ""
    wait = 1.05 - (time.time() - _geo_last_net)
    if wait > 0:
        time.sleep(min(wait, 1.2))
    try:
        import urllib.parse
        import urllib.request
        q = urllib.parse.urlencode({
            'lat': f'{lat:.6f}', 'lon': f'{lon:.6f}',
            'format': 'jsonv2', 'zoom': '18', 'addressdetails': '1',
            'accept-language': 'vi',
        })
        req = urllib.request.Request(
            'https://nominatim.openstreetmap.org/reverse?' + q,
            headers={'User-Agent': OSM_UA})
        _geo_last_net = time.time()
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode('utf-8', errors='replace') or '{}')
        name = _fmt_nominatim(data.get('address') or {})
        if not name:
            bits = str(data.get('name') or data.get('display_name') or '').split(',')
            name = '\n'.join(p.strip() for p in bits[:4] if p.strip())
        return name[:200]
    except Exception:
        _geo_fail_ts = time.time()
        return ""


def place_name_from_gps(path, allow_net=True):
    """Tên đường/khu vực từ EXIF GPS. Cache theo toạ độ ~11 m."""
    g = gps_cached(path)
    if not g:
        return ""
    key = _geo_key(g[0], g[1])
    with _geo_lock:
        if key in _geo_mem:
            return _geo_mem[key]
        disk = _load_geo_disk()
        hit = disk.get(key)
        if isinstance(hit, str) and hit:
            _geo_mem[key] = hit
            return hit
    if not allow_net:
        return ""
    name = _nominatim_reverse(g[0], g[1])
    if name:
        with _geo_lock:
            _geo_mem[key] = name
            _load_geo_disk()[key] = name
            _save_geo_disk()
    return name


def exif_coord_text(path):
    """Toạ độ EXIF gọn, hoặc '' nếu ảnh không có GPS."""
    g = gps_cached(path)
    if not g:
        return ''
    lat, lon = g
    return "{:.5f}°{}, {:.5f}°{}".format(
        abs(lat), 'N' if lat >= 0 else 'S',
        abs(lon), 'E' if lon >= 0 else 'W')


def location_from_exif(path, allow_net=True):
    """Tên đường từ toạ độ EXIF; chưa có cache / offline → in lat/lon."""
    return place_name_from_gps(path, allow_net=allow_net) or exif_coord_text(path)


def location_stamp_lines(text):
    """Tách địa điểm thành từng cấp, trên xuống — không dàn 1 hàng."""
    raw = str(text or '').strip()
    if not raw:
        return []
    seps = ('\n', ' · ', ' • ', ' | ', ' / ')
    bits = None
    for sep in seps:
        if sep in raw:
            bits = [p.strip(' ·•,;') for p in raw.split(sep) if p.strip(' ·•,;')]
            break
    if bits is None:
        if raw.count(',') >= 2:
            bits = [p.strip() for p in raw.split(',') if p.strip()]
        else:
            bits = [raw]
    skip = {'vietnam', 'việt nam', 'vn'}
    out, seen = [], set()
    for b in bits:
        key = b.lower()
        if not b or key in skip or key in seen:
            continue
        seen.add(key)
        out.append(b)
        if len(out) >= 6:
            break
    return out


def resolve_location(path, fx, allow_net=True):
    """Dòng địa điểm đóng dấu. Mặc định: toạ độ EXIF trên từng ảnh."""
    fx = fx or {}
    mode = str(fx.get('location_mode') or 'gps').strip().lower()
    excel = str(fx.get('location_excel') or '').strip()
    manual = str(fx.get('location_text') or '').strip()
    if mode == 'manual':
        return manual
    if mode == 'excel':
        return excel or location_from_exif(path, allow_net=allow_net) or manual
    exif_loc = location_from_exif(path, allow_net=allow_net)
    if mode == 'auto':
        return excel or exif_loc or manual
    return exif_loc or manual


def prefetch_geocode(paths):
    """Nominatim nền — 1 ảnh/giây, bỏ qua nếu đã có cache."""
    got = False
    for p in paths or []:
        if place_name_from_gps(p, allow_net=True):
            got = True
    return got


def _exif_datetime(path):
    try:
        with Image.open(path) as im:
            exif = im.getexif()
        if not exif:
            return None
        for tag in (36867, 306):
            if tag in exif:
                return datetime.datetime.strptime(str(exif.get(tag)), "%Y:%m:%d %H:%M:%S")
    except Exception:
        return None
    return None


# ════════════════════════════════════════════════════════════
#  Mini-map OSM
# ════════════════════════════════════════════════════════════
def _tile_disk(z, x, y):
    name = f"{z}_{x}_{y}.png"
    for d in (MAP_CACHE_DIR, V1_MAP_CACHE):
        fp = os.path.join(d, name)
        if os.path.exists(fp):
            return fp
    return os.path.join(MAP_CACHE_DIR, name)


def _fetch_osm_tile(z, x, y, allow_net=True):
    global _tile_fail_ts
    x = x % (2 ** z)
    if y < 0 or y >= 2 ** z:
        return None
    fp = _tile_disk(z, x, y)
    if os.path.exists(fp):
        try:
            return Image.open(fp).convert("RGB")
        except Exception:
            pass
    if not allow_net:
        return None
    if time.time() - _tile_fail_ts < 60:
        return None
    try:
        import urllib.request
        req = urllib.request.Request(
            OSM_TILE_URL.format(z=z, x=x, y=y),
            headers={"User-Agent": OSM_UA})
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = resp.read()
        try:
            os.makedirs(MAP_CACHE_DIR, exist_ok=True)
            with open(os.path.join(MAP_CACHE_DIR, f"{z}_{x}_{y}.png"), "wb") as f:
                f.write(data)
        except Exception:
            pass
        return Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        _tile_fail_ts = time.time()
        return None


def _render_real_map(lat, lon, W, H, zoom=OSM_ZOOM, allow_net=True):
    TS, n = 256, 2 ** zoom
    lat_r = math.radians(max(-85.05, min(85.05, lat)))
    xt = (lon + 180.0) / 360.0 * n
    yt = (1.0 - math.asinh(math.tan(lat_r)) / math.pi) / 2.0 * n
    cx_px, cy_px = xt * TS, yt * TS
    x0, y0 = int(cx_px - W / 2), int(cy_px - H / 2)
    canvas = Image.new("RGB", (W, H), (222, 228, 233))
    got = False
    for tx in range(x0 // TS, (x0 + W) // TS + 1):
        for ty in range(y0 // TS, (y0 + H) // TS + 1):
            tile = _fetch_osm_tile(zoom, tx, ty, allow_net=allow_net)
            if tile is None:
                continue
            got = True
            canvas.paste(tile, (tx * TS - x0, ty * TS - y0))
    return canvas if got else None


def _fallback_map(lat, lon, W, H):
    base = Image.new("RGB", (W, H), (0, 0, 0))
    d = ImageDraw.Draw(base)
    top_c, bot_c = (198, 226, 252), (56, 118, 180)
    for y in range(H):
        t = y / max(1, H - 1)
        d.line([(0, y), (W, y)],
               fill=tuple(int(top_c[i] + (bot_c[i] - top_c[i]) * t) for i in range(3)))
    step = max(12, W // 8)
    for gx in range(0, W, step):
        d.line([(gx, 0), (gx, H)], fill=(172, 206, 238), width=1)
    for gy in range(0, H, step):
        d.line([(0, gy), (W, gy)], fill=(172, 206, 238), width=1)
    cx, cy = W // 2, int(H * 0.40)
    r = max(5, H // 9)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(220, 38, 38),
              outline=(255, 255, 255), width=max(1, r // 4))
    return _finish_map(base.convert("RGBA"), lat, lon, W, H)


def _finish_map(base_rgba, lat, lon, W, H):
    d2 = ImageDraw.Draw(base_rgba)
    bar_h = max(12, int(H * 0.18))
    bar = Image.new("RGBA", (W, bar_h), (0, 0, 0, 150))
    base_rgba.alpha_composite(bar, (0, H - bar_h))
    font = _font(max(8, int(bar_h * 0.55)))
    txt = "{:.5f}°{}, {:.5f}°{}".format(abs(lat), 'N' if lat >= 0 else 'S',
                                        abs(lon), 'E' if lon >= 0 else 'W')
    tb = d2.textbbox((0, 0), txt, font=font)
    d2.text(((W - (tb[2] - tb[0])) // 2,
             H - bar_h + (bar_h - (tb[3] - tb[1])) // 2 - tb[1]),
            txt, font=font, fill=(255, 255, 255, 255))
    out = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    rad = max(6, H // 10)
    mask = Image.new("L", (W, H), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, W - 1, H - 1], radius=rad, fill=255)
    out.paste(base_rgba, (0, 0), mask)
    ImageDraw.Draw(out).rounded_rectangle([0, 0, W - 1, H - 1], radius=rad,
                                          outline=(255, 255, 255, 235),
                                          width=max(2, H // 50))
    return out


def compose_minimap(lat, lon, W, H, allow_net=True):
    key = (round(lat, 5), round(lon, 5), W, H)
    m = _minimap_cache.get(key)
    if m is not None:
        _minimap_cache.move_to_end(key)
        return m
    RW, RH = 512, 384
    base = _render_real_map(lat, lon, RW, RH, allow_net=allow_net)
    if base is None:
        return _fallback_map(lat, lon, W, H)
    d = ImageDraw.Draw(base)
    cx, cy = RW // 2, RH // 2
    r = max(10, RH // 11)
    d.ellipse([cx - r, cy - int(r * 2.2), cx + r, cy - int(r * 0.2)],
              fill=(220, 38, 38), outline=(255, 255, 255), width=3)
    d.polygon([(cx, cy), (cx - int(r * 0.7), cy - r // 2),
               (cx + int(r * 0.7), cy - r // 2)], fill=(220, 38, 38))
    scaled = base.convert("RGBA").resize((W, H), Image.Resampling.LANCZOS)
    m = _finish_map(scaled, lat, lon, W, H)
    _minimap_cache[key] = m
    while len(_minimap_cache) > _MINIMAP_MAX:
        _minimap_cache.popitem(last=False)
    return m


def stamp_minimap(img, lat, lon, ts_pos="bottom_right", opacity=85, allow_net=True):
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    w, h = img.size
    map_h = max(48, int(h * 0.20))
    map_w = int(map_h * 4 / 3)
    m = compose_minimap(lat, lon, map_w, map_h, allow_net=allow_net).copy()
    opacity = max(10, min(100, int(opacity)))
    if opacity < 100:
        a = m.split()[-1].point(lambda p: int(p * opacity / 100))
        m.putalpha(a)
    margin = int(h * 0.02)
    if "bottom" in ts_pos and "left" in ts_pos:
        x = w - map_w - margin
    else:
        x = margin
    y = h - map_h - margin
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    overlay.paste(m, (x, y), m)
    return Image.alpha_composite(img, overlay)


# ════════════════════════════════════════════════════════════
#  Timestamp
# ════════════════════════════════════════════════════════════
def _dt_from(path, fx):
    try:
        dt = datetime.datetime.fromtimestamp(os.path.getmtime(path))
    except Exception:
        dt = datetime.datetime.now()
    if fx.get('use_exif'):
        ex = _exif_datetime(path)
        if ex:
            dt = ex
    day, month, year = dt.day, dt.month, dt.year
    hour, minute, second = dt.hour, dt.minute, dt.second
    if fx.get('fix_date') and fx.get('fix_date_val'):
        try:
            d = datetime.datetime.strptime(fx['fix_date_val'], "%d/%m/%Y")
            day, month, year = d.day, d.month, d.year
        except Exception:
            pass
    if fx.get('fix_time') and fx.get('fix_time_val'):
        try:
            t = datetime.datetime.strptime(fx['fix_time_val'], "%H:%M")
            hour, minute, second = t.hour, t.minute, 0
        except Exception:
            pass
    try:
        return datetime.datetime(year, month, day, hour, minute, second)
    except Exception:
        return dt


def render_timestamp(img, path, fx, allow_net=True):
    orig = img.copy().convert("RGBA")
    w, h = orig.size
    if fx.get('use_timestamp', True):
        dt = _dt_from(path, fx)
        date_map = {"DD/MM/YYYY": "%d/%m/%Y", "MM/DD/YYYY": "%m/%d/%Y",
                    "YYYY-MM-DD": "%Y-%m-%d", "YYYY/MM/DD": "%Y/%m/%d"}
        date_fmt = date_map.get(fx.get('date_format', 'DD/MM/YYYY'), "%d/%m/%Y")
        tf = fx.get('time_format', '24h')
        if 'giây' in tf or tf.endswith('_sec') or '24h+s' in tf:
            time_fmt = "%H:%M:%S"
        elif '12h' in tf.lower() and 'giây' in tf:
            time_fmt = "%I:%M:%S %p"
        elif '12h' in tf.lower() or 'AM' in tf:
            time_fmt = "%I:%M %p"
        else:
            time_fmt = "%H:%M"
        mode = fx.get('display_mode', 'Ngày + Giờ')
        if mode in ('Chỉ ngày', 'Ngày'):
            txt_date = dt.strftime(date_fmt)
        elif mode in ('Chỉ giờ', 'Giờ'):
            txt_date = dt.strftime(time_fmt)
        else:
            txt_date = dt.strftime(f"{date_fmt} {time_fmt}")

        gps_line = gps_text(path) if fx.get('show_gps') else ''
        loc = resolve_location(path, fx, allow_net=allow_net)
        base_font_h = max(1, int(h * (float(fx.get('font_scale', 3.5)) / 100)))
        min_font_h = max(6, int(h * 0.015))
        margin = max(2, int(min(w, h) * float(fx.get('stamp_margin_pct', 3.5)) / 100.0))
        max_text_w = max(1, w - margin * 2)
        max_text_h = max(1, h - margin * 2)
        spacing_x = int(max(0, fx.get('letter_spacing', 0)))
        line_spacing = max(0.0, float(fx.get('line_spacing', 1.2)) - 1.0)

        def _measure(s, fnt):
            if not s:
                bb = fnt.getbbox("Ag")
                return 0, bb[3] - bb[1]
            if spacing_x <= 0:
                bb = fnt.getbbox(s)
                return bb[2] - bb[0], bb[3] - bb[1]
            lw, lh = 0, 0
            for i, ch in enumerate(s):
                bb = fnt.getbbox(ch)
                lw += bb[2] - bb[0]
                if i < len(s) - 1:
                    lw += spacing_x
                lh = max(lh, bb[3] - bb[1])
            return lw, lh

        def _wrap(text, fnt):
            out = []
            for para in str(text).split('\n'):
                cur = ""
                for word in para.split(' '):
                    trial = word if not cur else cur + ' ' + word
                    if not cur or _measure(trial, fnt)[0] <= max_text_w:
                        cur = trial
                    else:
                        out.append(cur)
                        cur = word
                out.append(cur)
            return out

        font_h = base_font_h
        lines, rights, tw, th, line_h, spacing_y = [], [], 0, 0, 0, 0
        font = _font(font_h)
        for _ in range(24):
            font = _font(max(1, font_h))
            loc_lines = []
            for piece in location_stamp_lines(loc):
                loc_lines.extend(_wrap(piece, font) or [piece])
            parts = loc_lines + ([gps_line] if gps_line else []) + [txt_date]
            final_txt = "\n".join(p for p in parts if p).strip()
            lines = final_txt.splitlines() or ['']
            spacing_y = int(line_spacing * font_h)
            heights, rights = [], []
            for line in lines:
                lw, lh = _measure(line, font)
                heights.append(lh)
                if spacing_x <= 0:
                    rights.append(font.getbbox(line)[2] if line else 0)
                else:
                    acc = 0
                    for i, ch in enumerate(line):
                        bb = font.getbbox(ch)
                        acc += bb[2] - bb[0]
                        if i < len(line) - 1:
                            acc += spacing_x
                    rights.append(acc)
            tw = max(rights) if rights else 0
            line_h = max(heights) if heights else font_h
            th = line_h * len(lines) + spacing_y * max(0, len(lines) - 1)
            out_pad = max(1, int(font_h / 15)) * 2
            if (tw + out_pad <= max_text_w and th + out_pad <= max_text_h) or font_h <= min_font_h:
                break
            font_h = max(min_font_h, int(font_h * 0.92))

        pos = fx.get('position', 'bottom_right')
        x, y = margin, margin
        if "right" in pos:
            x = w - tw - margin
        if "bottom" in pos:
            y = h - th - int(margin * 1.5)
        x = max(margin, min(x, max(margin, w - tw - margin)))
        y = max(margin, min(y, max(margin, h - th - margin)))

        if fx.get('add_background'):
            pad = int(font_h * 0.4)
            overlay = Image.new('RGBA', orig.size, (0, 0, 0, 0))
            ImageDraw.Draw(overlay).rectangle(
                [x - pad, y - pad, x + tw + pad, y + th + pad],
                fill=(0, 0, 0, int(fx.get('bg_opacity', 100))))
            orig = Image.alpha_composite(orig, overlay)

        draw = ImageDraw.Draw(orig)
        try:
            tcolor = ImageColor.getrgb(fx.get('text_color') or '#FF9900')
        except Exception:
            tcolor = (255, 153, 0)
        outline = max(1, int(font_h / 15))
        align = "right" if "right" in pos else "left"

        def draw_lines(color, dx=0, dy=0):
            yy = y + dy
            for i, line in enumerate(lines):
                rt = rights[i] if i < len(rights) else 0
                xx = (x + (tw - rt) + dx) if align == "right" else (x + dx)
                if spacing_x <= 0:
                    draw.text((xx, yy), line, font=font, fill=color)
                else:
                    cx = xx
                    for ch in line:
                        draw.text((cx, yy), ch, font=font, fill=color)
                        cx += (font.getbbox(ch)[2] - font.getbbox(ch)[0]) + spacing_x
                yy += line_h + spacing_y

        for ddx in range(-outline, outline + 1):
            for ddy in range(-outline, outline + 1):
                if ddx or ddy:
                    draw_lines("black", ddx, ddy)
        draw_lines(tcolor)

    if fx.get('logo_enable') and fx.get('logo_path') and os.path.exists(fx['logo_path']):
        try:
            logo = Image.open(fx['logo_path']).convert("RGBA")
            twl = max(1, int(w * (int(fx.get('logo_size_pct', 8)) / 100.0)))
            logo = logo.resize((twl, max(1, int(logo.height * twl / logo.width))),
                               Image.Resampling.LANCZOS)
            lop = int(fx.get('logo_opacity', 50))
            if lop < 100:
                a = logo.split()[-1].point(lambda p: int(p * lop / 100.0))
                logo.putalpha(a)
            lm = int(h * 0.02)
            lx, ly = lm, lm
            lpos = fx.get('logo_position', 'bottom_right')
            if "right" in lpos:
                lx = w - logo.width - lm
            if "bottom" in lpos:
                ly = h - logo.height - lm
            ov = Image.new("RGBA", orig.size, (0, 0, 0, 0))
            ov.paste(logo, (lx, ly), logo)
            orig = Image.alpha_composite(orig, ov)
        except Exception:
            pass
    return orig


# ════════════════════════════════════════════════════════════
#  Auto enhance + quality
# ════════════════════════════════════════════════════════════
def detect_blur(gray, threshold=60):
    try:
        g = gray
        if g.width > 512 or g.height > 512:
            g = g.copy()
            g.thumbnail((512, 512))
        lap = g.filter(ImageFilter.Kernel((3, 3), [0, 1, 0, 1, -4, 1, 0, 1, 0],
                                          scale=1, offset=128))
        return ImageStat.Stat(lap).var[0] < threshold
    except Exception:
        return False


def auto_enhance(img):
    try:
        has_alpha = img.mode == "RGBA"
        alpha = img.split()[-1] if has_alpha else None
        rgb = img.convert("RGB")
        gray = rgb.convert("L")
        mean, std = ImageStat.Stat(gray).mean[0], ImageStat.Stat(gray).stddev[0]
        if mean < 90:
            rgb = ImageEnhance.Brightness(rgb).enhance(min(1.8, 105.0 / max(1.0, mean)))
        if std < 45:
            rgb = ImageOps.autocontrast(rgb, cutoff=1)
        r_m, g_m, b_m = ImageStat.Stat(rgb).mean
        avg = (r_m + g_m + b_m) / 3.0
        if avg > 0 and (max(r_m, g_m, b_m) - min(r_m, g_m, b_m)) / avg > 0.12:
            bands = rgb.split()
            scaled = []
            for band, m in zip(bands, (r_m, g_m, b_m)):
                k = max(0.85, min(1.15, avg / max(1.0, m)))
                scaled.append(band.point(lambda p, k=k: min(255, int(p * k))))
            rgb = Image.merge("RGB", scaled)
        if detect_blur(gray, threshold=120):
            rgb = ImageEnhance.Sharpness(rgb).enhance(1.35)
        if has_alpha:
            rgb = rgb.convert("RGBA")
            rgb.putalpha(alpha)
        return rgb
    except Exception:
        return img


def check_quality(paths):
    results = {'blurry': [], 'dark': [], 'duplicates': []}
    hashes = {}
    for p in paths:
        try:
            with Image.open(p) as im:
                im = G.exif_upright(im)
                im.thumbnail((512, 512))
                small = im.convert("RGB")
            gray = small.convert("L")
            if detect_blur(gray):
                results['blurry'].append(p)
            if ImageStat.Stat(gray).mean[0] < 40:
                results['dark'].append(p)
            g = small.convert("L").resize((8, 8), Image.Resampling.LANCZOS)
            px = list(g.getdata())
            avg = sum(px) / 64.0
            h = "".join('1' if v > avg else '0' for v in px)
            if h in hashes:
                results['duplicates'].append((hashes[h], p))
            else:
                hashes[h] = p
        except Exception:
            continue
    return results


# ════════════════════════════════════════════════════════════
#  Pipeline ảnh hoàn chỉnh
# ════════════════════════════════════════════════════════════
def process_photo(path, src=None, fx=None, box=None, fit='fill', radius=0,
                  preview=False):
    """src: PIL đã decode (preview). box=(w,h) px đích. Trả RGBA."""
    fx = fx or {}
    try:
        if src is not None:
            # Thumbnail đã exif_upright — không transpose lần nữa.
            im = src.copy()
        else:
            im = Image.open(path)
            try:
                im.draft('RGB', (2048, 2048))
            except Exception:
                pass
            im = G.exif_upright(im)
        im = im.convert("RGBA")
    except Exception:
        return None

    if fx.get('auto_enhance'):
        im = auto_enhance(im)

    tw = th = None
    if box:
        tw, th = max(4, int(box[0])), max(4, int(box[1]))
        fit = G.photo_fit_mode(fit, im.width, im.height)
        if fit == 'fill':
            im = G.center_crop_to_ar(im, tw / max(1, th))

    REF_W = 800 if preview else 1600
    if im.width > REF_W:
        im = im.resize((REF_W, max(1, int(im.height * REF_W / im.width))),
                       Image.Resampling.LANCZOS)

    w, h = im.size
    rad_pct = int(radius or 0)
    if rad_pct > 0:
        r_px = int((rad_pct / 100) * (min(w, h) / 2))
        mask = Image.new("L", (w, h), 0)
        ImageDraw.Draw(mask).rounded_rectangle([0, 0, w, h], radius=r_px, fill=255)
        im.putalpha(mask)

    if fx.get('shadow'):
        s_rad, s_off = 12, (8, 8)
        shadow = Image.new('RGBA', (w + s_rad * 2 + s_off[0], h + s_rad * 2 + s_off[1]),
                           (255, 255, 255, 0))
        shadow.paste(Image.new('RGBA', (w, h), (0, 0, 0, 100)),
                     (s_rad + s_off[0], s_rad + s_off[1]), im)
        shadow = shadow.filter(ImageFilter.GaussianBlur(s_rad))
        shadow.paste(im, (s_rad, s_rad), im)
        im = shadow
        w, h = im.size

    if fx.get('use_timestamp') or fx.get('logo_enable'):
        try:
            im = render_timestamp(im, path, fx, allow_net=not preview)
            w, h = im.size
        except Exception as e:
            print('Timestamp err:', e)

    wm = (fx.get('watermark') or '').strip()
    if wm:
        try:
            d = ImageDraw.Draw(im)
            d.text((w - int(w * 0.05), h - int(h * 0.05)), wm,
                   fill=(255, 255, 255, 128),
                   font=_font(max(8, int(h * 0.05))), anchor="rd")
        except Exception:
            pass

    if fx.get('minimap'):
        gps = gps_cached(path)
        if gps:
            try:
                im = stamp_minimap(im, gps[0], gps[1],
                                   fx.get('position', 'bottom_right'),
                                   fx.get('minimap_opacity', 85),
                                   allow_net=not preview)
            except Exception as e:
                print('Minimap err:', e)

    if tw and th:
        im = G.resize_into_box(im, tw, th, fit)
    return im


def _fx_sig(fx):
    return json.dumps(fx or {}, sort_keys=True, default=str)


def cached_preview(path, src, fx, fit, radius, box_ar):
    """Ảnh đã đóng dấu ở kích thước ổn định — resize ra ô preview (kéo không lag)."""
    try:
        mt = os.path.getmtime(path)
    except Exception:
        mt = 0
    key = (path, mt, _fx_sig(fx), fit, int(radius or 0), round(float(box_ar or 1), 3), 'locstack')
    with _preview_lock:
        hit = _preview_cache.get(key)
        if hit is not None:
            _preview_cache.move_to_end(key)
            return hit.copy()
    tw = 720
    th = max(4, int(tw / max(0.2, box_ar)))
    im = process_photo(path, src, fx, box=(tw, th), fit=fit, radius=radius,
                       preview=True)
    if im is None:
        return None
    with _preview_lock:
        _preview_cache[key] = im
        while len(_preview_cache) > _PREVIEW_MAX:
            _preview_cache.popitem(last=False)
    return im.copy()


def clear_preview_cache():
    with _preview_lock:
        _preview_cache.clear()


def prefetch_maps(paths):
    """Tải tile OSM nền — gọi từ worker. Trả True nếu có tile mới."""
    got = False
    for p in paths or []:
        gps = gps_cached(p)
        if not gps:
            continue
        try:
            compose_minimap(gps[0], gps[1], 160, 120, allow_net=True)
            got = True
        except Exception:
            continue
    return got


def bake_export_png(path, fx, box_ar, fit, radius, temp_dir, opacity=100):
    """File PNG tạm cho export khi có đóng dấu / map / enhance."""
    payload = json.dumps({
        'p': path, 'm': os.path.getmtime(path), 'ar': round(box_ar, 4),
        'fit': fit, 'r': radius, 'op': opacity, 'fx': fx,
    }, sort_keys=True, default=str)
    key = hashlib.md5(payload.encode(errors='replace')).hexdigest()
    out = os.path.join(temp_dir, f"fx_{key}.png")
    if os.path.exists(out):
        return out
    tw = 1600
    th = max(1, int(tw / max(0.2, box_ar)))
    im = process_photo(path, None, fx, box=(tw, th), fit=fit, radius=radius,
                       preview=False)
    if im is None:
        return None
    op = int(opacity if opacity is not None else 100)
    if op < 99:
        if im.mode != 'RGBA':
            im = im.convert('RGBA')
        a = im.split()[-1].point(lambda p: int(p * op / 100))
        im.putalpha(a)
    im.save(out, "PNG")
    return out
