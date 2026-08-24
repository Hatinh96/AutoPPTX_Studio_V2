"""Vẽ chữ có độ mờ / giãn chữ — dùng chung preview.

Luôn nạp đúng file .ttf qua fonts.resolve_font_file (cùng nguồn với PPTX).
"""
import os

from PIL import Image, ImageDraw, ImageFont

from . import fonts as FN


def hex_rgb(hexv, fallback=(0, 0, 0)):
    try:
        h = str(hexv).lstrip('#')
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except Exception:
        return fallback


def load_font(name, size_px, bold=False, italic=False, path=None, index=0):
    """Nạp đúng file .ttf (cùng path mà picker / PPTX đang trỏ)."""
    size_px = max(7, int(size_px))
    if path and os.path.isfile(path):
        return FN.pil_font(path, size_px, int(index or 0))
    fp, idx = FN.resolve_font_file(name, bold, italic)
    if fp:
        return FN.pil_font(fp, size_px, idx)
    for fn in ('arial.ttf', 'segoeui.ttf', r'C:\Windows\Fonts\arial.ttf'):
        try:
            return ImageFont.truetype(fn, size_px)
        except Exception:
            continue
    return ImageFont.load_default()


def load_font_for_text(name, size_px, bold=False, italic=False,
                       path=None, index=0, text=''):
    """Nạp font; nếu thiếu glyph tiếng Việt thì chuyển Arial / fallback."""
    fam = FN.safe_family(name, text)
    use_path = path if fam.lower() == (name or '').lower() else None
    font = load_font(fam, size_px, bold, italic, use_path, index)
    if FN.font_covers(font, text):
        return font
    for alt in FN.VIET_FALLBACKS:
        if alt.lower() == fam.lower():
            continue
        fb = load_font(alt, size_px, bold, italic, None, 0)
        if FN.font_covers(fb, text):
            return fb
    return font


def render_textbox(text, wpx, hpx, *, font_name='Arial', size_px=16,
                   color='#000000', opacity=100, align='left',
                   bold=True, italic=False, tracking_px=0, valign='middle',
                   font_path=None, font_index=0):
    """Ảnh RGBA trong suốt — chữ có alpha + giãn chữ."""
    wpx, hpx = max(4, int(wpx)), max(4, int(hpx))
    img = Image.new('RGBA', (wpx, hpx), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    font = load_font_for_text(font_name, size_px, bold, italic,
                             font_path, font_index, text)
    r, g, b = hex_rgb(color)
    fill = (r, g, b, max(0, min(255, int(round(255 * float(opacity) / 100.0)))))
    pad = 2
    max_w = max(4, wpx - 2 * pad)
    lines = _wrap(str(text or ''), font, max_w, tracking_px)
    line_h = size_px + max(2, int(size_px * 0.25))
    total_h = line_h * len(lines)
    if valign == 'middle':
        y = max(0, (hpx - total_h) // 2)
    elif valign == 'bottom':
        y = max(0, hpx - total_h)
    else:
        y = 0
    for ln in lines:
        tw = _line_width(ln, font, tracking_px)
        if align == 'center':
            x = (wpx - tw) // 2
        elif align == 'right':
            x = wpx - pad - tw
        else:
            x = pad
        _draw_tracked(draw, x, y, ln, font, fill, tracking_px)
        y += line_h
        if y > hpx:
            break
    return img


def _line_width(s, font, tracking_px):
    if not s:
        return 0
    w = 0
    for i, ch in enumerate(s):
        bbox = font.getbbox(ch)
        w += max(1, bbox[2] - bbox[0])
        if i < len(s) - 1:
            w += tracking_px
    return w


def _wrap(text, font, max_w, tracking_px):
    out = []
    for para in text.split('\n'):
        words = para.split(' ')
        cur = ''
        for word in words:
            trial = word if not cur else cur + ' ' + word
            if _line_width(trial, font, tracking_px) <= max_w or not cur:
                cur = trial
            else:
                out.append(cur)
                cur = word
        out.append(cur)
    return out or ['']


def _draw_tracked(draw, x, y, s, font, fill, tracking_px):
    cx = x
    for i, ch in enumerate(s):
        draw.text((cx, y), ch, font=font, fill=fill)
        bbox = font.getbbox(ch)
        cx += max(1, bbox[2] - bbox[0])
        if i < len(s) - 1:
            cx += tracking_px
