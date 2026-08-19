"""EditorCanvas — khung xem trước slide kéo-thả MƯỢT.

Nguyên tắc chống lag (khác hẳn bản 1):
  1. KÉO DI CHUYỂN: chỉ canvas.move() các item có sẵn — không vẽ lại gì.
  2. KÉO CẠNH (resize): chỉ vẽ lại ĐÚNG phần tử đang kéo, giới hạn ~30fps;
     ảnh được resize từ thumbnail nhỏ trong RAM (sub-ms), không decode lại.
  3. Decode ảnh luôn ở luồng nền (ThumbCache) — canvas chỉ hiển thị khi sẵn.
  4. Không widget CTk nào bị dựng lại trong lúc kéo (timeline/sidebar đứng yên).

Tính năng: 8 tay nắm resize, Shift giữ tỉ lệ, snap + đường gióng (Alt tắt tạm),
badge toạ độ/kích thước, zoom tới con trỏ 25–400%, pan, undo/redo, khoá/ẩn
phần tử, lưới, phím mũi tên.
"""
import json
import math
import time
import tkinter as tk

from PIL import Image, ImageTk, ImageDraw

from . import geometry as G
from . import style as ST
from . import effects as FX
from .constants import (SLIDE_W_IN, SLIDE_H_IN, ELEMENT_COLORS, GUIDE_COLOR,
                        SNAP_PX, UNDO_MAX, ACCENT)

BASE_PPI = 96                      # px/inch ở zoom 100%
MIN_SIZE_IN = 0.3
HANDLE_PX = 8

_CURSORS = {"nw": "size_nw_se", "se": "size_nw_se",
            "ne": "size_ne_sw", "sw": "size_ne_sw",
            "n": "size_ns", "s": "size_ns",
            "e": "size_we", "w": "size_we"}

# thứ tự vẽ (dưới → trên); cũng là thứ tự khôi phục z-order sau redraw cục bộ
_Z_TAGS = ("bgimg", "grid", "el_image", "el_avatar", "el_title", "el_info",
           "el_channel", "guide", "sel", "badge")
_DRAW_ORDER = ("image", "avatar", "title", "info", "channel")


def _clamp(v, lo, hi):
    return max(lo, min(hi, v))


def _lighten(hexv, f=0.45):
    try:
        h = str(hexv).lstrip('#')
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return '#%02x%02x%02x' % tuple(int(c + (255 - c) * f) for c in (r, g, b))
    except Exception:
        return '#999999'


class EditorCanvas(tk.Canvas):
    def __init__(self, master, ctrl, **kw):
        kw.setdefault('highlightthickness', 0)
        kw.setdefault('bg', '#0a0a0d')
        super().__init__(master, **kw)
        self.ctrl = ctrl
        self.zoom = 100
        self.S = BASE_PPI
        self.CW = int(SLIDE_W_IN * self.S)
        self.CH = int(SLIDE_H_IN * self.S)
        self.snap_on = True
        self.grid_on = False
        self.selected = None
        self._ct = None                  # content cache của lần render hiện tại
        self._photos = {}                # name -> [PhotoImage] giữ tham chiếu
        self._drag = None
        self._panning = False
        self._sr = (0, 0, 0, 0)
        self._undo, self._redo = [], []
        self._last_guides = None
        self._rs_job = None
        self._rs_last = 0.0
        self._render_job = None

        self.bind('<Button-1>', self._on_press)
        self.bind('<B1-Motion>', self._on_motion)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.bind('<Double-Button-1>', self._on_double)
        self.bind('<Button-3>', self._on_right)
        self.bind('<Motion>', self._on_hover)
        self.bind('<Button-2>', self._on_mid_press)
        self.bind('<B2-Motion>', self._on_mid_motion)
        self.bind('<MouseWheel>', self._on_wheel)
        self.bind('<Shift-MouseWheel>', self._on_wheel)
        self.bind('<Control-MouseWheel>', self._on_wheel)
        self.bind('<Configure>', lambda e: self._update_scrollregion())

    # ══════════════════════════════ RENDER ══════════════════════════════
    def schedule_render(self, delay=30):
        """Gom nhiều thay đổi liên tiếp thành 1 lần vẽ."""
        if self._render_job:
            try:
                self.after_cancel(self._render_job)
            except Exception:
                pass
        self._render_job = self.after(delay, self.render)

    def render(self):
        self._render_job = None
        self.delete("all")
        self._photos = {}
        self._last_guides = None
        self._ct = self.ctrl.content()
        self._draw_bg()
        if self.grid_on:
            self._draw_grid()
        for name in _DRAW_ORDER:
            if self._visible(name):
                self._draw_element(name)
        self._draw_selection()
        self._restack()
        self._update_scrollregion()

    def _visible(self, name):
        return (self._ct or {}).get('visible', {}).get(name, True)

    def _locked(self, name):
        return (self._ct or {}).get('locked', {}).get(name, False)

    def _restack(self):
        for t in _Z_TAGS:
            try:
                self.tag_raise(t)
            except Exception:
                pass

    def _draw_bg(self):
        ct = self._ct
        self.create_rectangle(0, 0, self.CW, self.CH, fill='white',
                              outline='#3a3a40', tags=("bgimg",))
        bg = ct.get('bg')
        if bg:
            src = self.ctrl.thumbs.request(bg, self, self._thumb_ready)
            if src is not None:
                ph = ImageTk.PhotoImage(src.resize((self.CW, self.CH)))
                self._photos['bg'] = [ph]
                self.create_image(0, 0, anchor='nw', image=ph, tags=("bgimg",))

    def _draw_grid(self):
        step = 0.5 * self.S
        n = 1
        x = step
        while x < self.CW - 1:
            col = '#c9d4e0' if n % 2 == 0 else '#e3eaf2'
            self.create_line(x, 0, x, self.CH, fill=col, tags=("grid",))
            x += step
            n += 1
        n = 1
        y = step
        while y < self.CH - 1:
            col = '#c9d4e0' if n % 2 == 0 else '#e3eaf2'
            self.create_line(0, y, self.CW, y, fill=col, tags=("grid",))
            y += step
            n += 1

    # ─────────────────────── vẽ từng phần tử ───────────────────────
    def _redraw_element(self, name):
        self.delete(f"el_{name}")
        if self._visible(name):
            self._draw_element(name)
        self._restack()

    def _draw_element(self, name):
        L = self.ctrl.layout
        x, y, w, h = G.elem_box(L, name)
        S = self.S
        x0, y0 = x * S, y * S
        wpx, hpx = max(8, w * S), max(8, h * S)
        x1, y1 = x0 + wpx, y0 + hpx
        tag = ("el_" + name,)
        col = ELEMENT_COLORS[name]
        self._photos[name] = []

        if name == 'image':
            self._draw_image_grid(x, y, w, h, tag, col)
        elif name == 'avatar':
            self._draw_avatar(x0, y0, int(wpx), int(hpx), tag, col)
        elif name == 'info':
            self._draw_info(x0, y0, wpx, hpx, tag, col)
        else:
            self._draw_textblock(name, x0, y0, wpx, hpx, tag, col)

        if self._locked(name):
            self.create_text(x0 + 10, y0 + 10, text='🔒', anchor='nw',
                             font=('Segoe UI Emoji', -14), tags=tag)

    def _draw_image_grid(self, x, y, w, h, tag, col):
        ct = self._ct
        S = self.S
        n = max(1, ct.get('n', 4))
        paths = ct.get('images', [])[:n]
        cells = G.grid_cells((x, y, w, h), n, ct.get('ar'), ct.get('gap', 0.1))
        self.create_rectangle(x * S, y * S, (x + w) * S, (y + h) * S,
                              outline=col, width=1, dash=(4, 3), tags=tag)
        radius = int(ct.get('radius', 0) or 0)
        fit = ct.get('fit', 'fill')
        opacity = self.ctrl.layout['image'].get('opacity', 100)
        for i, (bx, by, bw, bh) in enumerate(cells):
            px0, py0 = bx * S, by * S
            pw, ph_ = max(4, int(bw * S)), max(4, int(bh * S))
            if i < len(paths):
                photo = self._cell_photo(paths[i], pw, ph_, radius, fit, opacity,
                                         apply_fx=True)
                if photo is not None:
                    self._photos['image'].append(photo)
                    # 'fit' có thể nhỏ hơn ô → căn giữa
                    ox = px0 + (pw - photo.width()) / 2
                    oy = py0 + (ph_ - photo.height()) / 2
                    self.create_image(ox, oy, anchor='nw', image=photo, tags=tag)
                else:
                    self.create_rectangle(px0, py0, px0 + pw, py0 + ph_,
                                          fill='#e9ecef', outline='', tags=tag)
            else:
                self.create_rectangle(px0, py0, px0 + pw, py0 + ph_,
                                      fill='#f1f3f5', outline=col, tags=tag)
                self.create_text(px0 + pw / 2, py0 + ph_ / 2, text='Ảnh',
                                 fill=col, tags=tag)

    def _cell_photo(self, path, w, h, radius_pct, fit, opacity=100, apply_fx=False):
        """PhotoImage cho 1 ô — resize từ thumbnail RAM, không decode file."""
        src = self.ctrl.thumbs.request(path, self, self._thumb_ready)
        if src is None:
            return None
        try:
            fx = {}
            try:
                fx = self.ctrl.fx_snapshot()
            except Exception:
                fx = {}
            use_fx = apply_fx and FX.needed(fx)
            dragging = (self._drag and self._drag.get('name') == 'image'
                        and self._drag.get('mode') == 'resize')
            if dragging:
                use_fx = False
            if use_fx:
                box_ar = w / max(1, h)
                im = FX.cached_preview(path, src, fx, fit, radius_pct, box_ar)
                if im is None:
                    return None
                if fit == 'fill' and im.size != (w, h):
                    im = im.resize((w, h))
                elif fit != 'fill':
                    sw, sh = im.size
                    sc = min(w / sw, h / sh)
                    im = im.resize((max(1, int(sw * sc)), max(1, int(sh * sc))))
                self._kick_minimap(path, fx)
            else:
                if fit == 'fill':
                    im = G.center_crop_to_ar(src, w / max(1, h)).resize((w, h))
                else:
                    sw, sh = src.size
                    sc = min(w / sw, h / sh)
                    im = src.resize((max(1, int(sw * sc)), max(1, int(sh * sc))))
                # Lúc kéo resize: chỉ ảnh cache, không bo góc / đóng dấu / map.
                if radius_pct > 0 and not dragging:
                    im = im.convert('RGBA')
                    rad = int((radius_pct / 100) * (min(im.size) / 2))
                    mask = Image.new('L', im.size, 0)
                    ImageDraw.Draw(mask).rounded_rectangle(
                        [0, 0, im.width, im.height], radius=rad, fill=255)
                    im.putalpha(mask)
            op = max(0, min(100, int(opacity if opacity is not None else 100)))
            if op < 99:
                if im.mode != 'RGBA':
                    im = im.convert('RGBA')
                a = im.split()[-1] if im.mode == 'RGBA' else Image.new('L', im.size, 255)
                im.putalpha(a.point(lambda p: int(p * op / 100)))
            return ImageTk.PhotoImage(im)
        except Exception:
            return None

    def _kick_minimap(self, path, fx):
        if not (fx.get('minimap') or
                str(fx.get('location_mode') or '') in ('gps', 'auto')):
            return
        try:
            self.ctrl.prefetch_minimap([path])
        except Exception:
            pass

    def clear_fx_cache(self):
        FX.clear_preview_cache()

    def _draw_avatar(self, x0, y0, w, h, tag, col):
        ct = self._ct
        path = ct.get('avatar')
        if path:
            photo = self._cell_photo(path, w, h, int(ct.get('avatar_radius', 0)),
                                     'fill', self.ctrl.layout['avatar'].get('opacity', 100))
            if photo is not None:
                self._photos['avatar'].append(photo)
                self.create_image(x0, y0, anchor='nw', image=photo, tags=tag)
                return
            self.create_rectangle(x0, y0, x0 + w, y0 + h, fill='#e9ecef',
                                  outline='', tags=tag)
            return
        self.create_rectangle(x0, y0, x0 + w, y0 + h, fill='#f1f3f5',
                              outline=col, width=2, tags=tag)
        self.create_text(x0 + w / 2, y0 + h / 2, text='Avatar', fill=col,
                         tags=tag)

    def _draw_info(self, x0, y0, wpx, hpx, tag, col):
        """Khối Thông tin — mô phỏng đúng bảng 2 cột của file xuất."""
        ct = self._ct
        rows = ct.get('info_rows') or []
        self.create_rectangle(x0, y0, x0 + wpx, y0 + hpx, outline=col,
                              width=1, dash=(3, 2), tags=tag)
        if not rows:
            return
        L = self.ctrl.layout
        fcfg = ct.get('font', {})
        fname = L['info'].get('font') or fcfg.get('name', 'Arial')
        fpath = L['info'].get('font_path') or fcfg.get('path')
        fidx = int(L['info'].get('font_index', 0) or fcfg.get('index', 0) or 0)
        base_pt = float(L['info'].get('size', 12))
        body_col = fcfg.get('color') or '#0f172a'
        label_col = _lighten(body_col, 0.45)
        accent_col = L['info'].get('accent_color') or '#2563eb'
        op = int(L['info'].get('opacity', 100) or 100)
        S = self.S
        rh = hpx / len(rows)
        pad = 3
        tmp = Image.new('RGBA', (max(4, int(wpx)), max(4, int(hpx))), (0, 0, 0, 0))
        draw = ImageDraw.Draw(tmp)
        for i, (label, value, accent) in enumerate(rows):
            ym = int(i * rh + rh / 2)
            lab_px = max(7, int(round(base_pt * 0.85 * S / 72)))
            val_pt = base_pt * (1.5 if accent else 1.0)
            val_px = max(7, int(round(val_pt * S / 72)))
            a = int(255 * op / 100)
            lr, lg, lb = ST.hex_rgb(label_col)
            vr, vg, vb = ST.hex_rgb(accent_col if accent else body_col)
            lf = ST.load_font(fname, lab_px, False, False, fpath, fidx)
            vf = ST.load_font(fname, val_px, True, False, fpath, fidx)
            draw.text((pad, ym), label, font=lf, fill=(lr, lg, lb, a), anchor='lm')
            draw.text((int(wpx) - pad, ym), str(value), font=vf,
                      fill=(vr, vg, vb, a), anchor='rm')
        ph = ImageTk.PhotoImage(tmp)
        self._photos['info'].append(ph)
        self.create_image(x0, y0, anchor='nw', image=ph, tags=tag)

    def _draw_textblock(self, name, x0, y0, wpx, hpx, tag, col):
        ct = self._ct
        L = self.ctrl.layout
        c = L[name]
        fcfg = ct.get('font', {})
        txt = ct.get('title' if name == 'title' else 'channel') or ''
        S = self.S
        w_in, h_in = wpx / S, hpx / S
        fpt = G.fit_block_size(txt, w_in, h_in, c.get('size', 12),
                               max_lines=c.get('max_lines', 2),
                               min_size=c.get('min_size', 8))
        fpx = max(7, int(round(fpt * S / 72)))
        f_color = c.get('color') or fcfg.get('color') or '#000000'
        fname = c.get('font') or fcfg.get('name', 'Arial')
        fpath = c.get('font_path') or fcfg.get('path')
        fidx = int(c.get('font_index', 0) or fcfg.get('index', 0) or 0)
        fbold = fcfg.get('bold', True)
        fital = fcfg.get('italic', False)
        op = int(c.get('opacity', 100) or 100)
        tracking = float(c.get('tracking', 0) or 0)
        al = c.get('align', 'left')
        self.create_rectangle(x0, y0, x0 + wpx, y0 + hpx, outline=col,
                              width=1, dash=(3, 2), tags=tag)
        tr_px = int(round(tracking * S / 72))
        im = ST.render_textbox(
            txt, wpx, hpx, font_name=fname, size_px=fpx, color=f_color,
            opacity=op, align=al, bold=fbold, italic=fital,
            tracking_px=tr_px, valign='middle' if name == 'title' else 'top',
            font_path=fpath, font_index=fidx)
        ph = ImageTk.PhotoImage(im)
        self._photos[name].append(ph)
        self.create_image(x0, y0, anchor='nw', image=ph, tags=tag)

    def _thumb_ready(self, path, img):
        """Thumbnail decode xong ở luồng nền → vẽ lại đúng phần tử liên quan."""
        ct = self._ct or {}
        if path == ct.get('bg'):
            self.schedule_render(20)
            return
        if path == ct.get('avatar'):
            self.after_idle(lambda: (self._redraw_element('avatar'),
                                     self._draw_selection()))
        if path in (ct.get('images') or []):
            self.after_idle(lambda: (self._redraw_element('image'),
                                     self._draw_selection()))

    # ══════════════════════════ SELECTION ══════════════════════════
    def select(self, name):
        if name == self.selected:
            self._draw_selection()
            return
        self.selected = name
        self._draw_selection()
        try:
            self.ctrl.on_select(name)
        except Exception:
            pass

    def _handles_for(self, name):
        if name == 'avatar':
            return ("nw", "ne", "se", "sw")
        if name == 'title':
            t = self.ctrl.layout['title']
            if (t.get('full_width_on_slide', True)
                    and t.get('align') in ('center', 'right')):
                return ("n", "s")
        return ("nw", "n", "ne", "e", "se", "s", "sw", "w")

    def _draw_selection(self):
        self.delete("sel")
        name = self.selected
        if not name or not self._visible(name):
            return
        x, y, w, h = G.elem_box(self.ctrl.layout, name)
        S = self.S
        x0, y0, x1, y1 = x * S, y * S, (x + w) * S, (y + h) * S
        self.create_rectangle(x0, y0, x1, y1, outline=ACCENT, width=2,
                              tags=("sel",))
        if self._locked(name):
            self._restack()
            return
        hs = HANDLE_PX / 2
        pos = {'nw': (x0, y0), 'n': ((x0 + x1) / 2, y0), 'ne': (x1, y0),
               'e': (x1, (y0 + y1) / 2), 'se': (x1, y1),
               's': ((x0 + x1) / 2, y1), 'sw': (x0, y1),
               'w': (x0, (y0 + y1) / 2)}
        for hk in self._handles_for(name):
            cx, cy = pos[hk]
            self.create_rectangle(cx - hs, cy - hs, cx + hs, cy + hs,
                                  fill='white', outline=ACCENT, width=1,
                                  tags=("sel", "h_" + hk))
        self._restack()

    # ══════════════════════════ EVENTS ══════════════════════════
    def _hit(self, cx, cy):
        """(handle_key|None, elem|None) tại toạ độ canvas."""
        items = self.find_overlapping(cx - 2, cy - 2, cx + 2, cy + 2)
        handle, elem = None, None
        for it in reversed(items):          # trên cùng trước
            tags = self.gettags(it)
            if handle is None:
                hk = next((t[2:] for t in tags if t.startswith('h_')), None)
                if hk:
                    handle = hk
                    break
            if elem is None:
                en = next((t[3:] for t in tags if t.startswith('el_')), None)
                if en:
                    elem = en
        return handle, elem

    def _on_press(self, ev):
        self.focus_set()
        cx, cy = self.canvasx(ev.x), self.canvasy(ev.y)
        handle, elem = self._hit(cx, cy)
        if handle and self.selected and not self._locked(self.selected):
            name = self.selected
            bx, by, bw, bh = G.elem_box(self.ctrl.layout, name)
            self.push_undo()
            self._drag = dict(mode='resize', name=name, hk=handle,
                              box=(bx, by, bw, bh),
                              ar=bw / max(0.001, bh), mx=cx, my=cy)
            self._build_snap_targets(exclude=name)
            return
        if elem:
            self.select(elem)
            if self._locked(elem):
                self._drag = None
                return
            c = self.ctrl.layout[elem]
            self.push_undo()
            self._drag = dict(mode='move', name=elem, mx=cx, my=cy,
                              sx=c['x'], sy=c['y'], ax=c['x'], ay=c['y'])
            self._build_snap_targets(exclude=elem)
            return
        # vùng trống → pan + bỏ chọn
        self._drag = None
        self._panning = True
        self.scan_mark(ev.x, ev.y)
        self.select(None)

    def _on_motion(self, ev):
        if self._panning:
            self.scan_dragto(ev.x, ev.y, gain=1)
            return
        d = self._drag
        if not d:
            return
        cx, cy = self.canvasx(ev.x), self.canvasy(ev.y)
        alt = bool(ev.state & 0x20000)
        shift = bool(ev.state & 0x0001)
        if d['mode'] == 'move':
            self._motion_move(d, cx, cy, alt)
        else:
            self._motion_resize(d, cx, cy, alt, shift)

    def _motion_move(self, d, cx, cy, alt):
        name = d['name']
        c = self.ctrl.layout[name]
        S = self.S
        _, _, w, h = G.elem_box(self.ctrl.layout, name)
        nx = d['sx'] + (cx - d['mx']) / S
        ny = d['sy'] + (cy - d['my']) / S
        # tiêu đề full-width: chỉ di chuyển dọc
        lock_x = False
        if name == 'title':
            t = self.ctrl.layout['title']
            lock_x = (t.get('full_width_on_slide', True)
                      and t.get('align') in ('center', 'right'))
        if lock_x:
            nx = d['sx']
        nx = _clamp(nx, 0, max(0, SLIDE_W_IN - w))
        ny = _clamp(ny, 0, max(0, SLIDE_H_IN - h))
        guides = []
        if self.snap_on and not alt:
            nx, ny, guides = self._snap_move(nx, ny, w, h, lock_x)
        dxp, dyp = (nx - d['ax']) * S, (ny - d['ay']) * S
        if dxp or dyp:
            self.move("el_" + name, dxp, dyp)
            self.move("sel", dxp, dyp)
            d['ax'], d['ay'] = nx, ny
            c['x'], c['y'] = nx, ny
            self.ctrl.on_drag_live(name)
        self._show_guides(guides)
        self._show_badge(f"X {nx:.2f}″   Y {ny:.2f}″", cx, cy)

    def _motion_resize(self, d, cx, cy, alt, shift):
        name = d['name']
        c = self.ctrl.layout[name]
        S = self.S
        hk = d['hk']
        bx, by, bw, bh = d['box']
        x0, y0, x1, y1 = bx, by, bx + bw, by + bh
        dx, dy = (cx - d['mx']) / S, (cy - d['my']) / S

        if name == 'avatar':
            ar = c.get('ar') or (4 / 3)
            sign = 1 if 'e' in hk else -1
            w = _clamp(bw + sign * dx, MIN_SIZE_IN, SLIDE_W_IN)
            h = w / ar
            if 'w' in hk:
                x0 = x1 - w
            else:
                x1 = x0 + w
            if 'n' in hk:
                y0 = y1 - h
            else:
                y1 = y0 + h
        else:
            if 'w' in hk:
                x0 = bx + dx
            if 'e' in hk:
                x1 = bx + bw + dx
            if 'n' in hk:
                y0 = by + dy
            if 's' in hk:
                y1 = by + bh + dy
            if shift and hk in ('nw', 'ne', 'se', 'sw'):
                w, h = x1 - x0, y1 - y0
                if h > 0 and w / h > d['ar']:
                    w = h * d['ar']
                else:
                    h = w / d['ar'] if d['ar'] else h
                if 'w' in hk:
                    x0 = x1 - w
                else:
                    x1 = x0 + w
                if 'n' in hk:
                    y0 = y1 - h
                else:
                    y1 = y0 + h

        guides = []
        if self.snap_on and not alt and not shift:
            x0, y0, x1, y1, guides = self._snap_resize(hk, x0, y0, x1, y1)

        # chặn nhỏ nhất + nằm trong slide
        x0 = _clamp(x0, 0, SLIDE_W_IN)
        x1 = _clamp(x1, 0, SLIDE_W_IN)
        y0 = _clamp(y0, 0, SLIDE_H_IN)
        y1 = _clamp(y1, 0, SLIDE_H_IN)
        if x1 - x0 < MIN_SIZE_IN:
            if 'w' in hk:
                x0 = x1 - MIN_SIZE_IN
            else:
                x1 = x0 + MIN_SIZE_IN
        if y1 - y0 < MIN_SIZE_IN and name != 'avatar':
            if 'n' in hk:
                y0 = y1 - MIN_SIZE_IN
            else:
                y1 = y0 + MIN_SIZE_IN

        if name == 'avatar':
            ar = c.get('ar') or (4 / 3)
            w = x1 - x0
            c['x'], c['y'], c['w'] = x0, y0, w
        elif hk in ('n', 's') and name == 'title' and len(self._handles_for('title')) == 2:
            c['y'], c['h'] = y0, y1 - y0
        else:
            c['x'], c['y'] = x0, y0
            c['w'] = x1 - x0
            c['h'] = y1 - y0

        self._throttled_redraw(name)
        self._show_guides(guides)
        self._show_badge(f"W {x1 - x0:.2f}″ × H {y1 - y0:.2f}″", cx, cy)
        self.ctrl.on_drag_live(name)

    def _throttled_redraw(self, name):
        now = time.perf_counter()
        if now - self._rs_last >= 0.03:
            self._rs_last = now
            self._redraw_element(name)
            self._draw_selection()
        elif not self._rs_job:
            def flush():
                self._rs_job = None
                self._rs_last = time.perf_counter()
                self._redraw_element(name)
                self._draw_selection()
            self._rs_job = self.after(35, flush)

    def _on_release(self, ev):
        self._panning = False
        if self._drag:
            name = self._drag['name']
            self._drag = None
            if self._rs_job:
                try:
                    self.after_cancel(self._rs_job)
                except Exception:
                    pass
                self._rs_job = None
            self.delete("guide")
            self.delete("badge")
            self._last_guides = None
            self._redraw_element(name)
            self._draw_selection()
            self.ctrl.on_layout_change()

    def _on_double(self, ev):
        cx, cy = self.canvasx(ev.x), self.canvasy(ev.y)
        _, elem = self._hit(cx, cy)
        if elem:
            self.select(elem)
            self.ctrl.open_style_panel(elem)

    def _on_right(self, ev):
        cx, cy = self.canvasx(ev.x), self.canvasy(ev.y)
        _, elem = self._hit(cx, cy)
        if elem:
            self.select(elem)
            self.ctrl.open_style_panel(elem)

    def _on_hover(self, ev):
        if self._drag or self._panning:
            return
        cx, cy = self.canvasx(ev.x), self.canvasy(ev.y)
        handle, elem = self._hit(cx, cy)
        cur = ''
        if handle:
            cur = _CURSORS.get(handle, '')
        elif elem and not self._locked(elem):
            cur = 'fleur'
        try:
            self.configure(cursor=cur)
        except Exception:
            self.configure(cursor='')

    def _on_mid_press(self, ev):
        self.scan_mark(ev.x, ev.y)

    def _on_mid_motion(self, ev):
        self.scan_dragto(ev.x, ev.y, gain=1)

    def _on_wheel(self, ev):
        if ev.state & 0x0004:                       # Ctrl → zoom tới con trỏ
            step = 10 if ev.delta > 0 else -10
            self.set_zoom(self.zoom + step, anchor=(ev.x, ev.y))
        elif ev.state & 0x0001:                     # Shift → cuộn ngang
            self.xview_scroll(-1 if ev.delta > 0 else 1, 'units')
        else:
            self.yview_scroll(-1 if ev.delta > 0 else 1, 'units')

    # ══════════════════════════ SNAP ══════════════════════════
    def _build_snap_targets(self, exclude=None):
        xs = [0.0, SLIDE_W_IN / 2, SLIDE_W_IN]
        ys = [0.0, SLIDE_H_IN / 2, SLIDE_H_IN]
        for other in _DRAW_ORDER:
            if other == exclude or not self._visible(other):
                continue
            ox, oy, ow, oh = G.elem_box(self.ctrl.layout, other)
            xs += [ox, ox + ow / 2, ox + ow]
            ys += [oy, oy + oh / 2, oy + oh]
        self._snap_xs, self._snap_ys = xs, ys

    def _nearest(self, targets, v):
        thr = SNAP_PX / self.S
        best, bd = None, thr
        for t in targets:
            dvi = abs(v - t)
            if dvi < bd:
                best, bd = t, dvi
        return best

    def _snap_move(self, nx, ny, w, h, lock_x=False):
        guides = []
        if not lock_x:
            for off in (0.0, w / 2, w):
                t = self._nearest(self._snap_xs, nx + off)
                if t is not None:
                    nx = t - off
                    guides.append(('v', t))
                    break
        for off in (0.0, h / 2, h):
            t = self._nearest(self._snap_ys, ny + off)
            if t is not None:
                ny = t - off
                guides.append(('h', t))
                break
        return nx, ny, guides

    def _snap_resize(self, hk, x0, y0, x1, y1):
        guides = []
        if 'w' in hk:
            t = self._nearest(self._snap_xs, x0)
            if t is not None:
                x0 = t
                guides.append(('v', t))
        if 'e' in hk:
            t = self._nearest(self._snap_xs, x1)
            if t is not None:
                x1 = t
                guides.append(('v', t))
        if 'n' in hk:
            t = self._nearest(self._snap_ys, y0)
            if t is not None:
                y0 = t
                guides.append(('h', t))
        if 's' in hk:
            t = self._nearest(self._snap_ys, y1)
            if t is not None:
                y1 = t
                guides.append(('h', t))
        return x0, y0, x1, y1, guides

    def _show_guides(self, guides):
        key = tuple(guides)
        if key == self._last_guides:
            return
        self._last_guides = key
        self.delete("guide")
        S = self.S
        for kind, v in guides:
            if kind == 'v':
                self.create_line(v * S, 0, v * S, self.CH, fill=GUIDE_COLOR,
                                 width=1, dash=(4, 3), tags=("guide",))
            else:
                self.create_line(0, v * S, self.CW, v * S, fill=GUIDE_COLOR,
                                 width=1, dash=(4, 3), tags=("guide",))
        self._restack()

    def _show_badge(self, text, cx, cy):
        self.delete("badge")
        x, y = cx + 16, cy + 20
        t = self.create_text(x, y, anchor='nw', text=text, fill='white',
                             font=('Segoe UI', -12, 'bold'), tags=("badge",))
        bx = self.bbox(t)
        if bx:
            r = self.create_rectangle(bx[0] - 6, bx[1] - 3, bx[2] + 6,
                                      bx[3] + 3, fill='#1f2937',
                                      outline='', tags=("badge",))
            self.tag_lower(r, t)
        self._restack()

    # ══════════════════════════ ZOOM / PAN ══════════════════════════
    def set_zoom(self, pct, anchor=None):
        pct = _clamp(int(pct), 25, 400)
        if pct == self.zoom:
            return
        if anchor is None:
            anchor = (self.winfo_width() // 2, self.winfo_height() // 2)
        ax_in = self.canvasx(anchor[0]) / self.S
        ay_in = self.canvasy(anchor[1]) / self.S
        self.zoom = pct
        self._apply_scale()
        self.render()
        # giữ điểm dưới con trỏ đứng yên
        x0, y0, x1, y1 = self._sr
        tw, th = max(1, x1 - x0), max(1, y1 - y0)
        self.xview_moveto((ax_in * self.S - anchor[0] - x0) / tw)
        self.yview_moveto((ay_in * self.S - anchor[1] - y0) / th)
        try:
            self.ctrl.on_zoom(self.zoom)
        except Exception:
            pass

    def _apply_scale(self):
        self.S = BASE_PPI * self.zoom / 100.0
        self.CW = int(SLIDE_W_IN * self.S)
        self.CH = int(SLIDE_H_IN * self.S)

    def zoom_fit(self):
        vw = max(100, self.winfo_width())
        vh = max(100, self.winfo_height())
        pct = min((vw - 50) / (SLIDE_W_IN * BASE_PPI),
                  (vh - 50) / (SLIDE_H_IN * BASE_PPI)) * 100
        self.zoom = _clamp(int(pct), 25, 400)
        self._apply_scale()
        self.render()
        self.xview_moveto(0)
        self.yview_moveto(0)
        try:
            self.ctrl.on_zoom(self.zoom)
        except Exception:
            pass

    def _update_scrollregion(self):
        vw = max(1, self.winfo_width())
        vh = max(1, self.winfo_height())
        padx = max(40, (vw - self.CW) // 2)
        pady = max(40, (vh - self.CH) // 2)
        self._sr = (-padx, -pady, self.CW + padx, self.CH + pady)
        self.config(scrollregion=self._sr)
        if vw - self.CW > 80:
            self.xview_moveto(0)
        if vh - self.CH > 80:
            self.yview_moveto(0)

    # ══════════════════════════ UNDO / NUDGE / ALIGN ══════════════════════════
    def push_undo(self):
        try:
            self._undo.append(json.dumps(self.ctrl.layout))
        except Exception:
            return
        if len(self._undo) > UNDO_MAX:
            self._undo.pop(0)
        self._redo.clear()

    def undo(self):
        if not self._undo:
            return
        self._redo.append(json.dumps(self.ctrl.layout))
        self._apply_state(self._undo.pop())

    def redo(self):
        if not self._redo:
            return
        self._undo.append(json.dumps(self.ctrl.layout))
        self._apply_state(self._redo.pop())

    def _apply_state(self, s):
        L = self.ctrl.layout
        L.clear()
        L.update(json.loads(s))
        self.render()
        self.ctrl.on_layout_change()

    def nudge(self, dx_in, dy_in):
        name = self.selected
        if not name or self._locked(name):
            return
        self.push_undo()
        c = self.ctrl.layout[name]
        _, _, w, h = G.elem_box(self.ctrl.layout, name)
        c['x'] = _clamp(c['x'] + dx_in, 0, max(0, SLIDE_W_IN - w))
        c['y'] = _clamp(c['y'] + dy_in, 0, max(0, SLIDE_H_IN - h))
        self._redraw_element(name)
        self._draw_selection()
        self.ctrl.on_layout_change()

    def align_selected(self, mode):
        """Căn phần tử đang chọn theo slide: left|cx|right|top|cy|bottom."""
        name = self.selected
        if not name or self._locked(name):
            return
        self.push_undo()
        c = self.ctrl.layout[name]
        _, _, w, h = G.elem_box(self.ctrl.layout, name)
        if mode == 'left':
            c['x'] = 0.35
        elif mode == 'cx':
            c['x'] = (SLIDE_W_IN - w) / 2
        elif mode == 'right':
            c['x'] = SLIDE_W_IN - w - 0.35
        elif mode == 'top':
            c['y'] = 0.35
        elif mode == 'cy':
            c['y'] = (SLIDE_H_IN - h) / 2
        elif mode == 'bottom':
            c['y'] = SLIDE_H_IN - h - 0.35
        self._redraw_element(name)
        self._draw_selection()
        self.ctrl.on_layout_change()

    def set_grid(self, on):
        self.grid_on = bool(on)
        self.render()
