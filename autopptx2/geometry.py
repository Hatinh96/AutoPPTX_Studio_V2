"""Toán bố cục dùng CHUNG cho preview và export.

Mọi phép tính vị trí/cỡ chữ đi qua đây đúng MỘT lần — preview vẽ sao thì
file PPTX xuất ra y như vậy (WYSIWYG). Đơn vị mặc định là inch trên hệ toạ
độ slide 13.33 x 7.5.
"""
import math
from functools import lru_cache

from PIL import Image, ImageFont, ImageOps

from .constants import SLIDE_W_IN
from . import fonts as FN

EXIF_ORIENTATION = 0x0112


def exif_orientation_value(im):
    """Tag Orientation trên ảnh đã mở: 1 = đứng sẵn, 2–8 = cần xoay/lật."""
    try:
        v = im.getexif().get(EXIF_ORIENTATION, 1)
        n = int(v or 1)
        return 1 if n <= 1 else n
    except Exception:
        return 1


def exif_orientation(path):
    """Tag Orientation của file. Lỗi đọc → 1 (không xoay)."""
    try:
        with Image.open(path) as im:
            return exif_orientation_value(im)
    except Exception:
        return 1


def exif_upright(im):
    """Áp EXIF Orientation đúng một lần. Ảnh đã đứng / không tag thì giữ nguyên.

    Gỡ blob EXIF sau khi xoay để bước sau (convert, save, PowerPoint) không
    xoay thêm 90°.
    """
    if im is None:
        return im
    if exif_orientation_value(im) == 1:
        return im
    try:
        out = ImageOps.exif_transpose(im)
        if out is not None:
            im = out
    except Exception:
        pass
    try:
        im.info.pop('exif', None)
    except Exception:
        pass
    try:
        exif = im.getexif()
        if exif and EXIF_ORIENTATION in exif:
            del exif[EXIF_ORIENTATION]
    except Exception:
        pass
    return im


def open_upright(path):
    """Mở ảnh với pixel đúng chiều đứng (cùng nguồn sự thật cho preview + export)."""
    return exif_upright(Image.open(path))


def image_size_upright(path):
    """(w, h) sau EXIF — header only, không decode pixel."""
    with Image.open(path) as im:
        w, h = im.size
        if exif_orientation_value(im) in (5, 6, 7, 8):
            return h, w
        return w, h


# ════════════════════════════════════════════════════════════
#  Lưới ảnh
# ════════════════════════════════════════════════════════════
# 0 / "Tự" = căn ô theo số ảnh thật, tối đa 4 (trang sau nếu nhiều hơn)
AUTO_PER_SLIDE = 0
AUTO_PER_SLIDE_CAP = 4


def per_slide_max(n_per_slide):
    """Số ảnh tối đa một slide. Tự (0) → 4."""
    if n_per_slide in (None, '', 'auto', 'Tự', 'tự', 'tu'):
        return AUTO_PER_SLIDE_CAP
    try:
        n = int(n_per_slide)
    except (TypeError, ValueError):
        return AUTO_PER_SLIDE_CAP
    if n <= 0:
        return AUTO_PER_SLIDE_CAP
    return max(1, min(9, n))


def parse_n_per_slide(v, default=AUTO_PER_SLIDE):
    """Config → int; Tự / auto / 0 = tự căn."""
    if v in (None, '', 'auto', 'Tự', 'tự', 'tu'):
        return AUTO_PER_SLIDE
    try:
        n = int(v)
    except (TypeError, ValueError):
        return default
    return 0 if n <= 0 else max(1, min(9, n))


def n_per_slide_label(n_per_slide):
    n = parse_n_per_slide(n_per_slide)
    return 'Tự' if n <= 0 else str(n)


def slide_images(paths, n_per_slide):
    """Ảnh trang đầu (preview): cắt theo tối đa, không pad ô trống."""
    cap = per_slide_max(n_per_slide)
    return [p for p in (paths or []) if p][:cap]


def grid_dims(n):
    """(cột, hàng) cho n ảnh thật trên slide — không chừa ô trống."""
    try:
        n = int(n or 1)
    except (TypeError, ValueError):
        n = 1
    n = max(1, n)
    if n <= 1:
        return 1, 1
    if n == 2:
        return 2, 1
    if n <= 4:
        return 2, 2
    if n <= 6:
        return 3, 2
    return 3, 3


def fit_cell(area_w, area_h, cols, rows, ar):
    """Kích thước 1 ô ảnh LỚN NHẤT đúng tỉ lệ `ar` mà lưới vẫn vừa vùng.
    ar=None (Tự do) → ô lấp đầy đúng vùng."""
    cw = max(0.0, area_w) / max(1, cols)
    ch = max(0.0, area_h) / max(1, rows)
    if not ar or ch <= 0:
        return cw, ch
    if cw / ch > ar:
        return ch * ar, ch          # giới hạn bởi chiều cao
    return cw, cw / ar              # giới hạn bởi chiều rộng


def grid_cells(area, n, ar, gap):
    """Toạ độ n ô ảnh (đã trừ khe hở, căn giữa từng hàng + cả khối).

    area: (x, y, w, h) inch của vùng ảnh.
    Trả list [(x, y, w, h)] inch cho từng ô.
    """
    ax, ay, aw, ah = area
    cols, rows = grid_dims(n)
    cw, ch = fit_cell(aw, ah, cols, rows, ar)
    rows_used = max(1, math.ceil(n / cols))
    oy = ay + max(0.0, ah - ch * rows_used) / 2.0
    cells = []
    for i in range(n):
        r, c = divmod(i, cols)
        in_row = min(cols, n - r * cols)
        ox = ax + max(0.0, aw - cw * in_row) / 2.0
        cx, cy = ox + c * cw, oy + r * ch
        bw, bh = max(0.05, cw - gap), max(0.05, ch - gap)
        cells.append((cx + (cw - bw) / 2.0, cy + (ch - bh) / 2.0, bw, bh))
    return cells


def crop_fractions(img_ar, box_ar):
    """(left, right, top, bottom) — tỉ lệ crop để ảnh lấp đầy ô không méo."""
    l = r = t = b = 0.0
    if img_ar > box_ar:            # ảnh rộng hơn ô → cắt 2 bên
        l = r = max(0.0, (1.0 - box_ar / img_ar) / 2.0)
    elif img_ar < box_ar:          # ảnh cao hơn ô → cắt trên/dưới
        t = b = max(0.0, (1.0 - img_ar / box_ar) / 2.0)
    return l, r, t, b


def center_crop_to_ar(img, target_ar):
    """Crop giữa ảnh PIL về đúng tỉ lệ target_ar (w/h)."""
    w, h = img.size
    if h == 0 or target_ar <= 0:
        return img
    ar = w / h
    if abs(ar - target_ar) < 1e-3:
        return img
    if ar > target_ar:
        nw = int(round(h * target_ar))
        x0 = (w - nw) // 2
        return img.crop((x0, 0, x0 + nw, h))
    nh = int(round(w / target_ar))
    y0 = (h - nh) // 2
    return img.crop((0, y0, w, y0 + nh))


def is_portrait_size(w, h):
    """Sau EXIF-upright: cao hơn rộng → ảnh dọc. Không xoay 90°."""
    try:
        return int(h) > int(w) > 0
    except (TypeError, ValueError):
        return False


def photo_fit_mode(requested, w, h):
    """Ảnh dọc luôn contain (giữ tỉ lệ gốc). Ảnh ngang theo chế độ user.

    'fill' trên ảnh dọc từng crop thành 4:3 ngang — mất đỉnh/chân. Ảnh ngang
    vẫn lấp đầy ô (V1). Chế độ 'fit'/'Vừa khung' giữ nguyên cho mọi ảnh.
    """
    if is_portrait_size(w, h):
        return 'fit'
    return requested or 'fill'


def contain_dims(iw, ih, bw, bh):
    """(w, h) vừa khít ô, giữ tỉ lệ ảnh, không crop / không méo."""
    if iw <= 0 or ih <= 0 or bw <= 0 or bh <= 0:
        return max(0.05, float(bw or 0.05)), max(0.05, float(bh or 0.05))
    iar = iw / ih
    box_ar = bw / bh
    if box_ar > iar:
        return bh * iar, bh
    return bw, bw / iar


def resize_into_box(img, tw, th, fit='fill'):
    """Đưa ảnh vào (tw, th) px: fill = crop giữa; fit = contain không crop."""
    tw, th = max(1, int(tw)), max(1, int(th))
    if img is None:
        return img
    if (fit or 'fill') == 'fill':
        im = center_crop_to_ar(img, tw / max(1, th))
        if im.size != (tw, th):
            return im.resize((tw, th), Image.Resampling.LANCZOS)
        return im
    sw, sh = img.size
    if sw <= 0 or sh <= 0:
        return img
    sc = min(tw / sw, th / sh)
    nw = max(1, int(round(sw * sc)))
    nh = max(1, int(round(sh * sc)))
    return img.resize((nw, nh), Image.Resampling.LANCZOS)


# ════════════════════════════════════════════════════════════
#  Đo & tự co chữ (đo bằng font thật → khớp PowerPoint)
# ════════════════════════════════════════════════════════════
@lru_cache(maxsize=1)
def _ref_font():
    fp, idx = FN.fallback_font_file()
    if fp:
        try:
            return ImageFont.truetype(fp, 100, index=idx)
        except Exception:
            pass
    for name in ("arial.ttf", "segoeui.ttf", "tahoma.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, 100)
        except Exception:
            continue
    return None


def text_width_pt(s, size):
    """Bề rộng (pt) của chuỗi s ở cỡ chữ size."""
    f = _ref_font()
    if f is not None:
        try:
            return f.getbbox(s)[2] * size / 100.0
        except Exception:
            pass
    return len(s) * size * 0.52


def wrap_line_count(text, width_pt, size):
    """(số dòng, bề rộng dòng rộng nhất) khi ngắt `text` trong width_pt."""
    text = str(text or "")
    if not text or width_pt <= 0 or size <= 0:
        return 1, 0.0
    lines, cur, widest = 1, "", 0.0
    for word in text.split(' '):
        trial = word if not cur else cur + ' ' + word
        if not cur or text_width_pt(trial, size) <= width_pt:
            cur = trial
        else:
            widest = max(widest, text_width_pt(cur, size))
            lines += 1
            cur = word
    widest = max(widest, text_width_pt(cur, size))
    return lines, widest


def fit_text_size(text, width_in, height_in, default_size,
                  max_lines=2, min_size=9, line_factor=1.25):
    """Cỡ chữ (pt) lớn nhất để `text` vừa TRỌN khung, cho phép wrap max_lines.

    Trừ hao 6% bề rộng vì chữ thường in đậm (rộng hơn font đo ~5%).
    """
    text = str(text or "")
    if not text or width_in <= 0 or height_in <= 0:
        return default_size
    width_pt = width_in * 72.0 * 0.94
    height_pt = height_in * 72.0
    best = min_size
    for size in range(int(default_size), int(min_size) - 1, -1):
        lines_needed, widest = wrap_line_count(text, width_pt, size)
        if lines_needed > max_lines or widest > width_pt:
            continue
        if lines_needed <= 1:      # 1 dòng → tôn trọng cỡ người dùng đặt
            best = size
            break
        if lines_needed * size * line_factor <= height_pt:
            best = size
            break
    return max(min_size, best)


def fit_block_size(text, width_in, height_in, default_size,
                   max_lines=2, min_size=9):
    """Cỡ chữ cho khối nhiều dòng (tách sẵn bằng \\n)."""
    lines = str(text or "").split("\n")
    per_line = height_in / max(1, len(lines))
    size = int(default_size)
    for ln in lines:
        size = min(size, fit_text_size(ln, width_in, per_line, default_size,
                                       max_lines=max_lines, min_size=min_size))
    return size


# ════════════════════════════════════════════════════════════
#  Khung phần tử
# ════════════════════════════════════════════════════════════
def title_box(tcfg):
    """(x, y, w, h) thực dùng của khung tiêu đề.

    full_width_on_slide + align center/right → trải hết bề ngang slide
    (chừa lề 0.35") để căn giữa/căn phải theo SLIDE chứ không theo khung."""
    al = tcfg.get('align', 'left')
    if tcfg.get('full_width_on_slide', True) and al in ('center', 'right'):
        m = 0.35
        return m, tcfg['y'], SLIDE_W_IN - m * 2, tcfg['h']
    return tcfg['x'], tcfg['y'], tcfg['w'], tcfg['h']


def elem_box(layout, name):
    """(x, y, w, h) inch của phần tử — nguồn sự thật duy nhất."""
    c = layout[name]
    if name == 'title':
        return title_box(c)
    if name == 'avatar':
        ar = c.get('ar') or (4 / 3)
        return c['x'], c['y'], c['w'], c['w'] / ar
    return c['x'], c['y'], c['w'], c.get('h', 0.8)


def title_string(tcfg, name_val, k, K):
    """Chuỗi tiêu đề cuối cùng (hậu tố trang + viết HOA) — dùng chung."""
    if tcfg.get('show_counter', False):
        s = "{} ({}/{})".format(name_val, k, K)
    else:
        s = str(name_val)
    if tcfg.get('upper', True):
        s = s.upper()
    return s
