"""Toán bố cục dùng CHUNG cho preview và export.

Mọi phép tính vị trí/cỡ chữ đi qua đây đúng MỘT lần — preview vẽ sao thì
file PPTX xuất ra y như vậy (WYSIWYG). Đơn vị mặc định là inch trên hệ toạ
độ slide 13.33 x 7.5.
"""
import math
from functools import lru_cache

from PIL import ImageFont

from .constants import SLIDE_W_IN


# ════════════════════════════════════════════════════════════
#  Lưới ảnh
# ════════════════════════════════════════════════════════════
def grid_dims(n):
    """(cột, hàng) cho n ảnh/slide."""
    if n <= 1:
        return 1, 1
    if n == 2:
        return 2, 1
    return 2, 2


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


# ════════════════════════════════════════════════════════════
#  Đo & tự co chữ (đo bằng font thật → khớp PowerPoint)
# ════════════════════════════════════════════════════════════
@lru_cache(maxsize=1)
def _ref_font():
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
