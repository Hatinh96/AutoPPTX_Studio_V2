"""Cửa sổ chính AutoPPTX Studio V2.

Quy tắc hiệu năng:
  • Sidebar/timeline KHÔNG bao giờ bị dựng lại trong lúc kéo phần tử.
  • Mọi việc nặng (đọc Excel, decode ảnh, xuất PPTX) chạy ở thread nền;
    worker không đụng vào biến Tk — dữ liệu được chụp sẵn trên main thread.
"""
import os
import sys
import json
import math
import copy
import threading
import tkinter as tk
from collections import OrderedDict
from tkinter import filedialog, messagebox, colorchooser

import customtkinter as ctk

from .constants import (CONFIG_DIR, CONFIG_PATH, PRESET_DIR, DEFAULT_LAYOUT,
                        ELEMENT_NAMES, ELEMENT_LABELS, AR_CHOICES, ALL_CHANNELS,
                        ALL_CITIES, ALL_DISTRICTS, STAMP_PRESETS,
                        ACCENT, ACCENT_HOVER, BG, CARD, TEXT, MUTED, INPUT,
                        BORDER, WORKSPACE, SUPABASE_URL, SUPABASE_ANON_KEY,
                        V1_CREDS_PATH, COLOR_PALETTE, find_asset,
                        WINDOW_ICONS, STAMP_LOGOS, BUILTIN_PRESETS,
                        builtin_pack, DEPT_UI, dept_key, dept_label,
                        pack_for_dept)
from . import geometry as G
from .datasource import (ExcelSource, ImageLibrary, AvatarIndex, match_row,
                         build_merged_groups, build_info_rows, build_info_table,
                         channel_text, clean, build_overview, screen_qty,
                         place_label, inspect_excel, build_photo_rename_plan,
                         compare_list_with_master,
                         sort_codes_by_city, order_groups_by_city, order_groups,
                         filter_groups_by_geo, geo_field_values)
from .thumbs import ThumbCache
from .editor import EditorCanvas
from . import exporter
from . import effects as FX
from .effects import DEFAULT_FX
from . import fonts as FN
from . import qa as QA
from . import cloud as Cloud
from .cloud import CloudClient, avatar_cache_dir
from .login import build_login

try:
    import windnd                    # kéo-thả file vào cửa sổ (chỉ Windows)
except Exception:
    windnd = None

_TL_W, _TL_H = 118, 66               # kích thước 1 ô timeline


def _load_fx(cfg):
    """Cấu hình đóng dấu: V2 `opts.fx`, hoặc mang sang từ bản 1 nếu chưa có."""
    fx = dict(DEFAULT_FX)
    o = cfg.get('opts') or {}
    saved = o.get('fx') if isinstance(o.get('fx'), dict) else None
    if saved:
        fx.update(saved)
        if not saved.get('_overview_v2'):
            fx['dashboard'] = True
            fx['_overview_v2'] = True
        if not saved.get('_loc_exif_v1'):
            fx['location_mode'] = 'gps'
            fx['_loc_exif_v1'] = True
        return fx
    v1 = {}
    try:
        with open(V1_CREDS_PATH, 'r', encoding='utf-8') as f:
            v1 = json.load(f) or {}
    except Exception:
        v1 = {}
    ts = v1.get('timestamp') or {}
    if ts or v1.get('enable_minimap') or v1.get('watermark_text'):
        fx.update({
            'use_timestamp': ts.get('use_timestamp', True),
            'font_scale': ts.get('font_scale', 3.5),
            'text_color': ts.get('text_color', '#FF9900'),
            'position': ts.get('position', 'bottom_right'),
            'add_background': ts.get('add_background', True),
            'bg_opacity': ts.get('bg_opacity', 100),
            'location_text': ts.get('location_text', ''),
            'display_mode': ts.get('display_mode', 'Ngày + Giờ'),
            'date_format': ts.get('date_format', 'DD/MM/YYYY'),
            'time_format': ts.get('time_format', '24h'),
            'use_exif': ts.get('use_exif', True),
            'show_gps': ts.get('show_gps_coords', False),
            'line_spacing': ts.get('line_spacing', 1.2),
            'letter_spacing': ts.get('letter_spacing', 0),
            'fix_date': ts.get('fix_date', False),
            'fix_date_val': ts.get('fix_date_val', ''),
            'fix_time': ts.get('fix_time', False),
            'fix_time_val': ts.get('fix_time_val', ''),
            'logo_enable': ts.get('logo_enable', False),
            'logo_path': ts.get('logo_path', ''),
            'logo_size_pct': ts.get('logo_size_pct', 8),
            'logo_opacity': ts.get('logo_opacity', 50),
            'logo_position': ts.get('logo_position', 'bottom_right'),
            'minimap': bool(v1.get('enable_minimap', False)),
            'minimap_opacity': int(v1.get('minimap_opacity', 85) or 85),
            'auto_enhance': bool(v1.get('enable_auto_enhance', False)),
            'quality_check': bool(v1.get('enable_quality_check', False)),
            'dashboard': bool(v1.get('enable_dashboard', False)),
            'watermark': v1.get('watermark_text', '') or '',
            'shadow': bool(v1.get('use_shadow', False)),
        })
        fx['location_mode'] = 'gps'
        fx['_loc_exif_v1'] = True
    import datetime as _dt
    if not fx.get('fix_date_val'):
        fx['fix_date_val'] = _dt.datetime.now().strftime('%d/%m/%Y')
    if not fx.get('fix_time_val'):
        fx['fix_time_val'] = _dt.datetime.now().strftime('%H:%M')
    if not fx.get('logo_path') or not os.path.exists(fx.get('logo_path')):
        p = find_asset(*STAMP_LOGOS)
        if p:
            fx['logo_path'] = p
    return fx


def _deep_merge(base, extra):
    out = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def _use_arial_layout(layout):
    """Montserrat thiếu nhiều dấu Việt — chuyển về Arial."""
    if not isinstance(layout, dict):
        return
    f = layout.get('font')
    if isinstance(f, dict) and str(f.get('name') or '').lower() == 'montserrat':
        f['name'] = 'Arial'
        f['path'] = ''
    for k in ('title', 'info', 'channel'):
        c = layout.get(k)
        if isinstance(c, dict) and str(c.get('font') or '').lower() == 'montserrat':
            c['font'] = 'Arial'
            c['font_path'] = ''


class GeoScrollPicker:
    """Dropdown cuộn chuột + tìm kiếm — thay CTkOptionMenu cho list dài."""

    def __init__(self, parent, values, width, command, initial, pack_side='left'):
        self._command = command
        self._values = list(values or [])
        self._value = initial
        self._popup = None
        self.btn = ctk.CTkButton(
            parent, text=self._label(initial), width=width, height=28,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            anchor='w', command=self.open)
        self.btn.pack(side=pack_side, padx=6)

    @staticmethod
    def _label(text, n=20):
        s = str(text or '')
        return s if len(s) <= n else s[: n - 1] + '…'

    def get(self):
        return self._value

    def set(self, value):
        self._value = value
        self.btn.configure(text=self._label(value))

    def configure(self, *, values=None, **kwargs):
        if values is not None:
            self._values = list(values)
        if kwargs:
            self.btn.configure(**kwargs)

    def open(self):
        if self._popup is not None:
            try:
                if self._popup.winfo_exists():
                    self._popup.focus()
                    return
            except Exception:
                pass
        win = ctk.CTkToplevel(self.btn)
        win.title('Chọn')
        w = max(240, int(self.btn.winfo_width()) + 48)
        win.geometry(f'{w}x340')
        try:
            x = self.btn.winfo_rootx()
            y = self.btn.winfo_rooty() + self.btn.winfo_height() + 2
            win.geometry(f'+{x}+{y}')
        except Exception:
            pass
        win.attributes('-topmost', True)
        try:
            win.transient(self.btn.winfo_toplevel())
        except Exception:
            pass
        self._popup = win
        win.protocol('WM_DELETE_WINDOW', lambda: self._close_popup())

        ent = ctk.CTkEntry(win, placeholder_text='Gõ để lọc…', height=30,
                           fg_color=INPUT)
        ent.pack(fill='x', padx=10, pady=(10, 6))
        sc = ctk.CTkScrollableFrame(win, fg_color=INPUT, corner_radius=8,
                                    height=260)
        sc.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        def pick(val):
            self.set(val)
            self._close_popup()
            if self._command:
                self._command(val)

        def render(q=''):
            for child in sc.winfo_children():
                child.destroy()
            needle = str(q or '').strip().casefold()
            shown = 0
            for val in self._values:
                if needle and needle not in str(val).casefold():
                    continue
                shown += 1
                ctk.CTkButton(
                    sc, text=str(val), height=32, anchor='w',
                    fg_color='transparent', hover_color=BORDER,
                    text_color=TEXT,
                    font=ctk.CTkFont(size=12),
                    command=lambda v=val: pick(v),
                ).pack(fill='x', pady=1)
            if not shown:
                ctk.CTkLabel(sc, text='Không có kết quả', text_color=MUTED,
                             font=ctk.CTkFont(size=11)).pack(pady=12)

        ent.bind('<KeyRelease>', lambda _e: render(ent.get()))
        render()
        ent.focus_set()
        win.bind('<Escape>', lambda _e: self._close_popup())

    def _close_popup(self):
        if self._popup is not None:
            try:
                self._popup.destroy()
            except Exception:
                pass
        self._popup = None


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        cfg = self._load_config()
        ctk.set_appearance_mode(cfg.get('appearance', 'Light'))
        ctk.set_default_color_theme('blue')
        self.title("AutoPPTX Studio V2")
        self.geometry(cfg.get('geometry', '1560x900'))
        self.minsize(1280, 720)

        # ── Model ──
        self.layout = _deep_merge(DEFAULT_LAYOUT, cfg.get('layout'))
        _use_arial_layout(self.layout)
        o = cfg.get('opts', {})
        folders = [f for f in (o.get('image_folders') or []) if f]
        if not folders and o.get('image_folder'):
            folders = [o['image_folder']]
        style = o.get('slide_style')
        if style not in ('report', 'saleskit'):
            try:
                style = 'saleskit' if float(self.layout.get('info', {}).get('w') or 0) >= 8 else 'report'
            except Exception:
                style = 'report'
        self.opts = {
            'excel_path': o.get('excel_path', ''),
            'compare_list_path': o.get('compare_list_path', ''),
            'image_folder': folders[0] if folders else '',
            'image_folders': folders,
            'scan_subfolders': bool(o.get('scan_subfolders', True)),
            'avatar_folder': o.get('avatar_folder', ''),
            'bg_mode': o.get('bg_mode', 'white'),          # white | image
            'bg_path': o.get('bg_path', ''),
            'n_per_slide': G.parse_n_per_slide(o.get('n_per_slide', 0)),
            'slides_per_file': G.parse_slides_per_file(o.get('slides_per_file', 0)),
            'ar_label': o.get('ar_label', '4:3'),
            'channel_filter': o.get('channel_filter', ALL_CHANNELS),
            'city_filter': o.get('city_filter', ALL_CITIES),
            'district_filter': o.get('district_filter', ALL_DISTRICTS),
            'sort_mode': o.get('sort_mode', 'city'),
            'pad_blank_slides': bool(o.get('pad_blank_slides', False)),
            'split_export_by': o.get('split_export_by', 'none'),
            'export_pdf': bool(o.get('export_pdf', False)),
            'stamp_preset': o.get('stamp_preset', 'Tùy chỉnh'),
            'show_no_exif_badge': bool(o.get('show_no_exif_badge', True)),
            'channel_enabled': bool(o.get('channel_enabled', True)),
            'channel_template': o.get('channel_template', 'Report {channel} 2026'),
            'open_after_export': bool(o.get('open_after_export', True)),
            'write_checklist': bool(o.get('write_checklist', True)),
            'qa_gate': bool(o.get('qa_gate', True)),
            # Mỗi lần mở app: Đám mây. Có thể đổi sang Máy này trong phiên.
            'app_mode': 'cloud',
            'visible': {n: True for n in ELEMENT_NAMES} | o.get('visible', {}),
            'locked': {n: False for n in ELEMENT_NAMES} | o.get('locked', {}),
            'fx': _load_fx(cfg),
            'recent_colors': list(o.get('recent_colors') or []),
            'slide_style': style,
            'dept': o.get('dept') if o.get('dept') in ('sales', 'bd') else (
                'bd' if style == 'saleskit' else 'sales'),
            'dept_state': dict(o.get('dept_state') or {}),
        }
        for rec in (self.opts.get('dept_state') or {}).values():
            if isinstance(rec, dict):
                _use_arial_layout(rec.get('layout'))
        # Pack BD cũ ẩn avatar. Bật một lần; không đụng bố cục Sales.
        if not o.get('bd_show_avatar'):
            self.opts['bd_show_avatar'] = True
            bd_av = copy.deepcopy(pack_for_dept('bd')['layout']['avatar'])
            if self.opts.get('dept') == 'bd':
                self.opts['visible']['avatar'] = True
                if not (o.get('visible') or {}).get('avatar', False):
                    self.layout['avatar'] = _deep_merge(
                        self.layout.get('avatar') or {}, bd_av)
            bd_st = (self.opts.get('dept_state') or {}).get('bd')
            if isinstance(bd_st, dict):
                vis = bd_st.setdefault('visible', {})
                was_on = bool(vis.get('avatar'))
                vis['avatar'] = True
                if not was_on and isinstance(bd_st.get('layout'), dict):
                    bd_st['layout']['avatar'] = _deep_merge(
                        bd_st['layout'].get('avatar') or {}, bd_av)
        # Pack BD cũ ẩn kênh quảng cáo. Bật một lần cho cả preview và xuất.
        if not o.get('bd_show_channel'):
            self.opts['bd_show_channel'] = True
            self.opts['channel_enabled'] = True
            self.opts['visible']['channel'] = True
            bd_st = (self.opts.get('dept_state') or {}).get('bd')
            if isinstance(bd_st, dict):
                bd_st['channel_enabled'] = True
                vis = bd_st.setdefault('visible', {})
                vis['channel'] = True
        # Lưới cũ luôn vẽ đủ 2–4 ô (khung "Ảnh" trống). Chuyển mặc định sang Tự.
        if not o.get('auto_grid_v1'):
            self.opts['n_per_slide'] = 0
            self.opts['auto_grid_v1'] = True
        if int(cfg.get('layout_version') or 0) == 2 and not o.get('dept') and not o.get('slide_style'):
            pack = pack_for_dept('sales')
            self.layout = copy.deepcopy(pack['layout'])
            self.opts['slide_style'] = 'report'
            self.opts['dept'] = 'sales'
            self.opts['visible'].update(pack['visible'])
            self.opts['channel_enabled'] = True
        self.title('AutoPPTX Studio V2  ·  ' + dept_label(self.opts.get('dept')))
        self.cloud = CloudClient(
            cfg.get('supabase_url') or SUPABASE_URL,
            cfg.get('supabase_key') or SUPABASE_ANON_KEY)
        self._bg_manifest = []
        self.excel = ExcelSource()
        self.imglib = ImageLibrary()
        self.avatars = AvatarIndex()
        self.thumbs = ThumbCache()
        self.group_idx = 0
        self._by_code = {}
        self._merged = {}
        self._content_cache = None
        self._save_job = None
        self._insp_updating = False
        self._tl_pending = {}
        self._tl_photos = {}
        self._tl_items = []
        self._search = ''
        self._ui_built = False
        self._bg_photos = []
        self._bg_thumb_btns = {}
        self._mm_busy = False
        self._mm_done = set()
        self._collapse = {}
        self._side_mode = 'nguon'
        self._side_open = True
        self._insp_visible = False
        self._rail_btns = {}
        self._export_win = None
        self._loc_mode_labels = {
            'gps': 'Theo toạ độ EXIF',
            'auto': 'Mã Excel → GPS',
            'excel': 'Theo mã Excel',
            'manual': 'Gõ tay',
        }
        self._loc_mode_vals = {v: k for k, v in self._loc_mode_labels.items()}

        self.shell = ctk.CTkFrame(self, fg_color='transparent')
        self.shell.pack(fill='both', expand=True)
        self.protocol('WM_DELETE_WINDOW', self._on_close)
        self._set_app_window_icon()
        self.show_login()

        if windnd is not None:
            try:
                windnd.hook_dropfiles(self, func=self._on_drop)
            except Exception:
                pass

    def _set_app_window_icon(self):
        def _apply():
            try:
                if sys.platform == 'darwin':
                    icns = find_asset('app_icon.icns')
                    if icns:
                        self.iconbitmap(icns)
                        return
                ico = find_asset('app_icon.ico')
                png = find_asset('app_icon_1024.png')
                if not png:
                    pngs = tuple(p for p in WINDOW_ICONS if p.lower().endswith('.png'))
                    png = find_asset(*pngs)
                if ico:
                    try:
                        self.iconbitmap(ico)
                    except Exception:
                        pass
                if getattr(sys, 'frozen', False) and sys.platform == 'win32':
                    try:
                        self.iconbitmap(sys.executable)
                    except Exception:
                        pass
                if png:
                    from PIL import Image, ImageTk
                    im = Image.open(png).convert('RGBA')
                    im = im.resize((256, 256), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(im)
                    self.iconphoto(True, photo)
                    self._app_icon_photo = photo
            except Exception:
                pass

        _apply()
        self.after(50, _apply)
        self.after(200, _apply)

    # ════════════════════════ CONFIG ════════════════════════
    @staticmethod
    def _load_config():
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_config(self):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        data = {
            'layout': self.layout,
            'opts': self.opts,
            'geometry': self.geometry(),
            'appearance': ctk.get_appearance_mode(),
            'supabase_url': self.cloud.url,
            'supabase_key': self.cloud.key,
        }
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=1)
        except Exception as e:
            print('Lỗi lưu config:', e)

    def _schedule_save(self, delay=900):
        if self._save_job:
            try:
                self.after_cancel(self._save_job)
            except Exception:
                pass
        self._save_job = self.after(delay, lambda: (self._save_config(),
                                                    setattr(self, '_save_job', None)))

    def _on_close(self):
        self._save_config()
        try:
            self.thumbs.shutdown()
        except Exception:
            pass
        self.destroy()

    def _clear_shell(self):
        for w in self.shell.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass

    def show_login(self):
        self._ui_built = False
        self._clear_shell()
        self.configure(fg_color=BG)
        self.shell.configure(fg_color=BG)
        build_login(self.shell, self)

    def _login_set_busy(self, widgets, busy, err=''):
        try:
            widgets['btn'].configure(
                text='Đang đăng nhập…' if busy else 'Đăng nhập',
                state='disabled' if busy else 'normal')
            for k in ('email', 'pw', 'eye'):
                widgets[k].configure(state='disabled' if busy else 'normal')
            widgets['err'].configure(text=err)
            if err:
                widgets['err'].pack(fill='x', pady=(10, 0), before=widgets['btn'])
            else:
                widgets['err'].pack_forget()
        except Exception:
            pass

    def _try_login(self, email, pw, remember, widgets):
        if not email or not pw:
            self._login_set_busy(widgets, False, 'Nhập email và mật khẩu.')
            return
        self._login_set_busy(widgets, True)

        def work():
            ok, msg = self.cloud.login(email, pw)
            self.after(0, lambda: self._login_done(ok, msg, email, pw, remember, widgets))
        threading.Thread(target=work, daemon=True).start()

    def _login_done(self, ok, msg, email, pw, remember, widgets):
        if not ok:
            self._login_set_busy(widgets, False, msg)
            return
        if remember:
            Cloud.save_creds(email, pw)
        else:
            Cloud.clear_creds()
        self.show_main()

    def logout(self):
        try:
            self._save_config()
        except Exception:
            pass
        self.cloud.logout()
        self.show_login()

    def show_main(self):
        self._clear_shell()
        self._build_ui()
        self._ui_built = True
        self._bind_keys()
        self._restore_sources()
        self.after(250, self._first_show)
        if self.opts.get('app_mode') == 'cloud' and self.cloud.logged_in():
            self.after(700, lambda: self._sync_all_cloud(silent=True))

        def _fonts_ready():
            try:
                self._sync_inspector()
                self._invalidate()
            except Exception:
                pass

        def _scan_fonts():
            try:
                FN.catalog()
            except Exception:
                pass
            try:
                self.after(0, _fonts_ready)
            except Exception:
                pass
        threading.Thread(target=_scan_fonts, daemon=True).start()

    # ════════════════════════ UI ════════════════════════
    def _build_ui(self):
        host = self.shell
        host.configure(fg_color=BG)
        # ── Toolbar ──
        bar = ctk.CTkFrame(host, height=46, fg_color=CARD, corner_radius=0)
        bar.pack(fill='x', side='top')
        bar.pack_propagate(False)

        def tbtn(text, cmd, w=34, tip=None):
            return ctk.CTkButton(bar, text=text, width=w, height=30,
                                 corner_radius=6, fg_color='transparent',
                                 hover_color=INPUT, text_color=TEXT,
                                 font=ctk.CTkFont(size=14), command=cmd)

        tbtn('↩', lambda: self.editor.undo()).pack(side='left', padx=(10, 2), pady=8)
        tbtn('↪', lambda: self.editor.redo()).pack(side='left', padx=2)
        ctk.CTkFrame(bar, width=1, height=24, fg_color=BORDER).pack(side='left', padx=8)
        tbtn('−', lambda: self.editor.set_zoom(self.editor.zoom - 10)).pack(side='left', padx=2)
        self.lbl_zoom = ctk.CTkLabel(bar, text='100%', width=46, text_color=TEXT)
        self.lbl_zoom.pack(side='left')
        tbtn('+', lambda: self.editor.set_zoom(self.editor.zoom + 10)).pack(side='left', padx=2)
        tbtn('Vừa', lambda: self.editor.zoom_fit(), w=44).pack(side='left', padx=2)
        tbtn('100%', lambda: self.editor.set_zoom(100), w=48).pack(side='left', padx=2)
        ctk.CTkFrame(bar, width=1, height=24, fg_color=BORDER).pack(side='left', padx=8)
        self.sw_snap = ctk.CTkSwitch(bar, text='Hít', width=60,
                                     command=self._toggle_snap,
                                     progress_color=ACCENT)
        self.sw_snap.select()
        self.sw_snap.pack(side='left', padx=4)
        self.sw_grid = ctk.CTkSwitch(bar, text='Lưới', width=66,
                                     command=self._toggle_grid,
                                     progress_color=ACCENT)
        self.sw_grid.pack(side='left', padx=4)

        ctk.CTkLabel(bar, text='Phòng:', text_color=MUTED).pack(side='left', padx=(16, 4))
        self._dept_ready = False
        self.seg_dept = self._seg_btn(
            bar, DEPT_UI, dept_label(self.opts.get('dept')),
            command=self._on_dept,
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER,
            width=160, height=30)
        self.seg_dept.pack(side='left', padx=2)

        ctk.CTkLabel(bar, text='Preset:', text_color=MUTED).pack(side='left', padx=(12, 4))
        self.om_preset = ctk.CTkOptionMenu(bar, width=150, height=30,
                                           values=['— Preset —'],
                                           fg_color=INPUT, button_color=INPUT,
                                           button_hover_color=BORDER,
                                           text_color=TEXT,
                                           command=self._apply_preset)
        self.om_preset.pack(side='left', padx=2)
        tbtn('💾', self._save_preset, w=34).pack(side='left', padx=2)
        tbtn('🗑', self._delete_preset, w=34).pack(side='left', padx=2)
        tbtn('↺', self._reset_layout, w=34).pack(side='left', padx=6)

        self.btn_theme = tbtn('☀', self._toggle_appearance, w=34)
        self.btn_theme.pack(side='right', padx=(2, 6))
        who = self.cloud.name or self.cloud.email or 'Tài khoản'
        self.btn_account = ctk.CTkButton(
            bar, text='👤', width=40, height=30, corner_radius=6,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(size=14), command=self._show_account_menu)
        self.btn_account.pack(side='right', padx=(2, 10))
        ctk.CTkLabel(bar, text=who, text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(side='right', padx=4)

        act = ctk.CTkFrame(bar, fg_color='transparent')
        act.pack(side='right', padx=(8, 10))
        ctk.CTkButton(
            act, text='Xuất PPTX', width=108, height=30,
            font=ctk.CTkFont(size=13, weight='bold'),
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.export
        ).pack(side='left')
        ctk.CTkButton(
            act, text='Xuất ảnh', width=88, height=30,
            font=ctk.CTkFont(size=12, weight='bold'),
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=self.export_stamped
        ).pack(side='left', padx=(6, 0))
        ctk.CTkButton(
            act, text='⚙', width=32, height=30,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=self._open_export_opts
        ).pack(side='left', padx=(6, 0))
        self.btn_insp = ctk.CTkButton(
            act, text='Khung', width=58, height=30,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=self._toggle_inspector)
        self.btn_insp.pack(side='left', padx=(6, 0))

        # ── Thân: rail | panel | canvas | inspector ──
        body = ctk.CTkFrame(host, fg_color='transparent')
        body.pack(fill='both', expand=True)

        self.rail = ctk.CTkFrame(body, width=72, fg_color=CARD, corner_radius=10)
        self.rail.pack(side='left', fill='y', padx=(8, 0), pady=8)
        self.rail.pack_propagate(False)
        self._rail_btns = {}
        for key, title in (
                ('nguon', 'Nguồn'), ('mau', 'Mẫu'),
                ('gps', 'GPS'), ('stamp', 'Dấu')):
            b = ctk.CTkButton(
                self.rail, text=title, width=60, height=52,
                corner_radius=8, fg_color='transparent', hover_color=INPUT,
                text_color=TEXT, font=ctk.CTkFont(size=12, weight='bold'),
                command=lambda k=key: self._show_side(k))
            b.pack(padx=6, pady=(8, 0))
            self._rail_btns[key] = b
        self.btn_hide_side = ctk.CTkButton(
            self.rail, text='◂ Ẩn', width=60, height=36,
            corner_radius=8, fg_color=INPUT, hover_color=BORDER,
            text_color=TEXT, font=ctk.CTkFont(size=12, weight='bold'),
            command=self._hide_side)
        self.btn_hide_side.pack(side='bottom', padx=6, pady=10)

        self.insp_host = ctk.CTkFrame(body, width=252, fg_color=CARD, corner_radius=10)
        self.insp_host.pack_propagate(False)

        self.side_host = ctk.CTkFrame(body, width=348, fg_color=CARD, corner_radius=10)
        self.side_host.pack(side='left', fill='y', padx=(8, 0), pady=8)
        self.side_host.pack_propagate(False)
        self._side_w_default = 300
        self._side_w_mau = 336
        self.panel_nguon = ctk.CTkFrame(self.side_host, fg_color='transparent')
        self.panel_mau = ctk.CTkFrame(self.side_host, fg_color='transparent')
        self.panel_gps = ctk.CTkFrame(self.side_host, fg_color='transparent')
        self.panel_fx = ctk.CTkFrame(self.side_host, fg_color='transparent')
        self._build_tab_nguon(self.panel_nguon)
        self._build_tab_mau(self.panel_mau)
        self._build_tab_gps(self.panel_gps)
        self._build_tab_stamp(self.panel_fx)
        self._build_tab_slide(self.insp_host)
        self.panel_nguon.pack(fill='both', expand=True)
        self._side_open = True
        self._side_mode = 'nguon'
        self._paint_rail()

        center = ctk.CTkFrame(body, fg_color='transparent')
        center.pack(side='left', fill='both', expand=True, padx=8, pady=8)

        self._build_format_bar(center)

        self.editor = EditorCanvas(center, self)
        self.editor.pack(fill='both', expand=True)

        # ── Timeline nhóm ảnh ──
        tl_row = ctk.CTkFrame(center, fg_color=CARD, corner_radius=10)
        tl_row.pack(fill='x', pady=(8, 0))
        nav = ctk.CTkFrame(tl_row, fg_color='transparent')
        nav.pack(side='left', padx=6, pady=6)
        ctk.CTkButton(nav, text='◀', width=30, height=30, fg_color=INPUT,
                      hover_color=BORDER, text_color=TEXT,
                      command=lambda: self._step_group(-1)).pack(pady=1)
        ctk.CTkButton(nav, text='▶', width=30, height=30, fg_color=INPUT,
                      hover_color=BORDER, text_color=TEXT,
                      command=lambda: self._step_group(1)).pack(pady=1)
        right = ctk.CTkFrame(tl_row, fg_color='transparent')
        right.pack(side='left', fill='x', expand=True, padx=(0, 8), pady=6)
        srow = ctk.CTkFrame(right, fg_color='transparent')
        srow.pack(fill='x')
        self.ent_search = ctk.CTkEntry(srow, placeholder_text='Tìm mã điểm…',
                                       width=180, height=26, fg_color=INPUT)
        self.ent_search.pack(side='left')
        self.ent_search.bind('<KeyRelease>', self._on_search)
        self.lbl_group = ctk.CTkLabel(srow, text='', text_color=MUTED)
        self.lbl_group.pack(side='left', padx=10)
        wrap = tk.Frame(right)
        wrap.pack(fill='x', pady=(4, 0))
        self.tl = tk.Canvas(wrap, height=_TL_H + 32, highlightthickness=0,
                            bg=self._tl_chrome()['bg'])
        hbar = tk.Scrollbar(wrap, orient='horizontal', command=self.tl.xview)
        self.tl.configure(xscrollcommand=hbar.set)
        self.tl.pack(fill='x')
        hbar.pack(fill='x')
        self.tl.bind('<Button-1>', self._tl_click)
        self.tl.bind('<MouseWheel>',
                     lambda e: self.tl.xview_scroll(-1 if e.delta > 0 else 1, 'units'))

        # ── Status bar ──
        st = ctk.CTkFrame(host, height=28, fg_color=CARD, corner_radius=0)
        st.pack(fill='x', side='bottom')
        st.pack_propagate(False)
        self.lbl_status = ctk.CTkLabel(st, text='Sẵn sàng.', text_color=MUTED,
                                       font=ctk.CTkFont(size=11))
        self.lbl_status.pack(side='left', padx=10)
        self.lbl_est = ctk.CTkLabel(st, text='', text_color=MUTED,
                                    font=ctk.CTkFont(size=11))
        self.lbl_est.pack(side='left', padx=8)
        self.lbl_coverage = ctk.CTkLabel(st, text='', text_color=MUTED,
                                         font=ctk.CTkFont(size=11))
        self.lbl_coverage.pack(side='right', padx=10)
        try:
            self.lbl_coverage.configure(cursor='hand2')
        except Exception:
            pass
        self.lbl_coverage.bind('<Button-1>', lambda _e: self._show_coverage())

        self._preset_ready = False
        self._refresh_preset_menu()
        self._preset_ready = True
        self._dept_ready = True

        class _Nav:
            def set(_self, name):
                self._goto_panel(name)
        self.sidebar = _Nav()

    def _tl_chrome(self):
        light = ctk.get_appearance_mode() == 'Light'
        if light:
            return {'bg': '#ececef', 'card': '#ffffff', 'slot': '#e4e4e7',
                    'outline': '#d4d4d8', 'text': '#3f3f46'}
        return {'bg': '#131316', 'card': '#1b1b1f', 'slot': '#26262b',
                'outline': '#2a2a30', 'text': '#b9b9c3'}

    def _build_format_bar(self, parent):
        """Thanh soạn: phần tử + font/GPS. Xuất/Khung nằm trên thanh cửa sổ."""
        bar = ctk.CTkFrame(parent, height=48, fg_color=CARD, corner_radius=10)
        bar.pack(fill='x', pady=(0, 8))
        bar.pack_propagate(False)
        self.fmt_bar = bar

        self.lbl_fmt_sel = ctk.CTkLabel(
            bar, text='Vùng ảnh', width=80,
            font=ctk.CTkFont(size=13, weight='bold'), text_color=TEXT)
        self.lbl_fmt_sel.pack(side='left', padx=(12, 4))
        ctk.CTkFrame(bar, width=1, height=24, fg_color=BORDER).pack(side='left', padx=6)

        self.frm_fmt_align = ctk.CTkFrame(bar, fg_color='transparent')
        self._fmt_align_btns = {}
        for lab in ('Trái', 'Giữa', 'Phải'):
            b = ctk.CTkButton(
                self.frm_fmt_align, text=lab, width=56, height=32,
                fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
                font=ctk.CTkFont(size=12),
                command=lambda v=lab: self._on_elem_align(v))
            b.pack(side='left', padx=1)
            self._fmt_align_btns[lab] = b

        self.frm_fmt_text = ctk.CTkFrame(bar, fg_color='transparent')
        self.btn_fmt_font = ctk.CTkButton(
            self.frm_fmt_text, text='Arial', width=140, height=32,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT, anchor='w',
            command=self._open_font_dialog)
        self.btn_fmt_font.pack(side='left', padx=4)
        ctk.CTkButton(self.frm_fmt_text, text='−', width=32, height=32,
                      fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
                      command=lambda: self._fmt_nudge_size(-1)).pack(side='left')
        self.lbl_fmt_size = ctk.CTkLabel(
            self.frm_fmt_text, text='16', width=32,
            font=ctk.CTkFont(size=14, weight='bold'), text_color=TEXT)
        self.lbl_fmt_size.pack(side='left', padx=2)
        ctk.CTkButton(self.frm_fmt_text, text='+', width=32, height=32,
                      fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
                      command=lambda: self._fmt_nudge_size(1)).pack(side='left')
        self.chip_fmt = self._make_color_chip(
            self.frm_fmt_text, '#000000', 32, command=self._on_elem_color)
        self.chip_fmt.pack(side='left', padx=(8, 4))
        self.btn_fmt_bold = ctk.CTkButton(
            self.frm_fmt_text, text='B', width=32, height=32,
            font=ctk.CTkFont(size=15, weight='bold'),
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=self._fmt_toggle_bold)
        self.btn_fmt_bold.pack(side='left', padx=2)
        self.btn_fmt_italic = ctk.CTkButton(
            self.frm_fmt_text, text='I', width=32, height=32,
            font=ctk.CTkFont(size=15, weight='bold'),
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=self._fmt_toggle_italic)
        self.btn_fmt_italic.pack(side='left', padx=2)

        self.frm_fmt_image = ctk.CTkFrame(bar, fg_color='transparent')
        self.btn_fmt_date = ctk.CTkButton(
            self.frm_fmt_image, text='Ngày', width=64, height=32,
            command=lambda: self._toggle_fx_chip('use_timestamp'))
        self.btn_fmt_date.pack(side='left', padx=2)
        self.btn_fmt_loc = ctk.CTkButton(
            self.frm_fmt_image, text='Địa điểm', width=78, height=32,
            command=lambda: self._toggle_fx_chip('stamp_location'))
        self.btn_fmt_loc.pack(side='left', padx=2)
        self.btn_fmt_gps = ctk.CTkButton(
            self.frm_fmt_image, text='GPS', width=56, height=32,
            command=lambda: self._toggle_fx_chip('show_gps'))
        self.btn_fmt_gps.pack(side='left', padx=2)
        self.btn_fmt_map = ctk.CTkButton(
            self.frm_fmt_image, text='Map', width=56, height=32,
            command=lambda: self._toggle_fx_chip('minimap'))
        self.btn_fmt_map.pack(side='left', padx=2)
        self.chip_fmt_stamp = self._make_color_chip(
            self.frm_fmt_image, self.opts.get('fx', {}).get('text_color', '#FF9900'),
            32, command=self._pick_ts_color)
        self.chip_fmt_stamp.pack(side='left', padx=(8, 4))
        ctk.CTkButton(
            self.frm_fmt_image, text='Chi tiết', width=72, height=32,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=lambda: self._show_side('gps')
        ).pack(side='left', padx=4)

        self.frm_fmt_image.pack(side='left')
        self._fmt_mode = 'image'

    # ─────────── các thẻ sidebar ───────────
    @staticmethod
    def _seg_btn(parent, values, value, command=None, **kw):
        """Tạo segmented button: .set() trước, command sau — tránh CTk bắn callback lúc khởi tạo."""
        w = ctk.CTkSegmentedButton(parent, values=list(values), **kw)
        try:
            w.set(value)
        except Exception:
            pass
        if command is not None:
            w.configure(command=command)
        return w

    @staticmethod
    def _card(parent, title):
        f = ctk.CTkFrame(parent, fg_color=INPUT, corner_radius=10)
        f.pack(fill='x', padx=8, pady=(8, 0))
        ctk.CTkLabel(
            f, text=title, font=ctk.CTkFont(size=12, weight='bold'),
            text_color=TEXT, anchor='w', justify='left', wraplength=220
        ).pack(fill='x', padx=10, pady=(8, 2))
        return f

    def _path_label(self, parent):
        lbl = ctk.CTkLabel(parent, text='(chưa chọn)', text_color=MUTED,
                           font=ctk.CTkFont(size=11), anchor='w', wraplength=280)
        lbl.pack(fill='x', padx=10, pady=(0, 8))
        return lbl

    def _step_card(self, parent, num, title, hint=''):
        f = ctk.CTkFrame(parent, fg_color=INPUT, corner_radius=10)
        f.pack(fill='x', padx=8, pady=(8, 0))
        head = ctk.CTkFrame(f, fg_color='transparent')
        head.pack(fill='x', padx=10, pady=(10, 2))
        badge = ctk.CTkFrame(head, width=26, height=26, fg_color=ACCENT,
                             corner_radius=13)
        badge.pack(side='left')
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text=str(num), text_color='#ffffff',
                     font=ctk.CTkFont(size=12, weight='bold')).pack(expand=True)
        ctk.CTkLabel(head, text=title, font=ctk.CTkFont(size=13, weight='bold'),
                     text_color=TEXT).pack(side='left', padx=8)
        if hint:
            ctk.CTkLabel(f, text=hint, text_color=MUTED,
                         font=ctk.CTkFont(size=11), wraplength=270,
                         justify='left', anchor='w').pack(
                             fill='x', padx=10, pady=(0, 6))
        return f

    def _collapsible(self, parent, key, title, padx=8, pady=(6, 0), start_open=False):
        wrap = ctk.CTkFrame(parent, fg_color='transparent')
        wrap.pack(fill='x', padx=padx, pady=pady)
        head = ctk.CTkButton(
            wrap, text=('▾  ' if start_open else '▸  ') + title, height=34,
            anchor='w', fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(size=12),
            command=lambda k=key: self._toggle_collapsible(k))
        head.pack(fill='x')
        body = ctk.CTkFrame(wrap, fg_color=INPUT, corner_radius=10)
        self._collapse[key] = {
            'head': head, 'body': body, 'title': title, 'open': bool(start_open)}
        if start_open:
            body.pack(fill='x', pady=(4, 0))
        return body

    def _toggle_collapsible(self, key):
        rec = (self._collapse or {}).get(key)
        if not rec:
            return
        rec['open'] = not rec['open']
        rec['head'].configure(
            text=('▾  ' if rec['open'] else '▸  ') + rec['title'])
        if rec['open']:
            rec['body'].pack(fill='x', pady=(4, 0))
        else:
            rec['body'].pack_forget()

    def _hide_side(self):
        host = getattr(self, 'side_host', None)
        if host is not None:
            try:
                host.pack_forget()
            except Exception:
                pass
        self._side_open = False
        self._side_mode = None
        self._paint_rail()

    def _hide_inspector(self):
        host = getattr(self, 'insp_host', None)
        if host is not None:
            try:
                host.pack_forget()
            except Exception:
                pass
        self._insp_visible = False
        self._paint_insp_btn()

    def _show_side(self, key):
        aliases = {
            'fx': 'stamp', 'hieuung': 'stamp', 'dau': 'stamp', 'stamp': 'stamp',
            'gps': 'gps', 'map': 'gps',
            'mau': 'mau', 'mẫu': 'mau', 'template': 'mau', 'bg': 'mau', 'nen': 'mau',
            'nguon': 'nguon',
        }
        key = aliases.get((key or '').strip().lower(), 'nguon')
        if self._side_open and self._side_mode == key:
            self._hide_side()
            return
        self._hide_inspector()
        host = getattr(self, 'side_host', None)
        if host is not None and not self._side_open:
            try:
                host.pack(side='left', fill='y', padx=(8, 0), pady=8,
                          after=self.rail)
            except Exception:
                try:
                    host.pack(side='left', fill='y', padx=(8, 0), pady=8)
                except Exception:
                    pass
        self._side_open = True
        self._side_mode = key
        panels = (
            ('nguon', getattr(self, 'panel_nguon', None)),
            ('mau', getattr(self, 'panel_mau', None)),
            ('gps', getattr(self, 'panel_gps', None)),
            ('stamp', getattr(self, 'panel_fx', None)),
        )
        for name, frm in panels:
            if frm is None:
                continue
            try:
                frm.pack_forget()
            except Exception:
                pass
            if name == key:
                try:
                    frm.pack(fill='both', expand=True)
                except Exception:
                    pass
        self._paint_rail()
        self._fit_side_width(key)
        if key == 'mau':
            self.after(80, self._fetch_bg_list)

    def _toggle_inspector(self):
        if self._insp_visible:
            self._hide_inspector()
            return
        self._hide_side()
        host = getattr(self, 'insp_host', None)
        if host is None:
            return
        try:
            host.pack(side='right', fill='y', padx=(0, 8), pady=8)
        except Exception:
            pass
        self._insp_visible = True
        self._paint_insp_btn()

    def _fit_side_width(self, key):
        host = getattr(self, 'side_host', None)
        if host is None:
            return
        w = self._side_w_mau if key == 'mau' else self._side_w_default
        try:
            host.configure(width=w)
        except Exception:
            pass

    def _paint_rail(self):
        cur = self._side_mode if getattr(self, '_side_open', False) else None
        for k, b in (self._rail_btns or {}).items():
            on = k == cur
            try:
                b.configure(
                    fg_color=ACCENT if on else 'transparent',
                    hover_color=ACCENT_HOVER if on else INPUT,
                    text_color='#ffffff' if on else TEXT)
            except Exception:
                pass

    def _paint_insp_btn(self):
        btn = getattr(self, 'btn_insp', None)
        if btn is None:
            return
        on = bool(self._insp_visible)
        try:
            btn.configure(
                fg_color=ACCENT if on else INPUT,
                hover_color=ACCENT_HOVER if on else BORDER,
                text_color='#ffffff' if on else TEXT)
        except Exception:
            pass

    def _goto_panel(self, name):
        n = (name or '').strip()
        if n in ('Đóng dấu', 'Dấu', 'Hiệu ứng', 'fx', 'stamp'):
            self._show_side('stamp')
        elif n in ('GPS', 'gps', 'Map', 'Địa điểm'):
            self._show_side('gps')
        elif n in ('Mẫu', 'Nền', 'Template', 'mau'):
            self._show_side('mau')
        elif n in ('Xuất', 'Export'):
            self._open_export_opts()
        elif n in ('Slide',):
            return
        else:
            self._show_side('nguon')

    def _paint_toggle(self, btn, on):
        if btn is None:
            return
        try:
            btn.configure(
                fg_color=ACCENT if on else INPUT,
                hover_color=ACCENT_HOVER if on else BORDER,
                text_color='#ffffff' if on else TEXT)
        except Exception:
            pass

    def _toggle_fx_chip(self, key):
        fx = self.opts.setdefault('fx', {})
        self._set_fx(key, not bool(fx.get(key)))

    def _open_export_opts(self):
        old = getattr(self, '_export_win', None)
        if old is not None:
            try:
                if old.winfo_exists():
                    old.lift()
                    old.focus()
                    self._update_estimate()
                    return
            except Exception:
                pass
        win = ctk.CTkToplevel(self)
        win.title('Tùy chọn xuất')
        win.geometry('400x700')
        win.attributes('-topmost', True)
        try:
            win.transient(self)
        except Exception:
            pass
        self._export_win = win

        def _close():
            self._export_win = None
            try:
                win.destroy()
            except Exception:
                pass
        win.protocol('WM_DELETE_WINDOW', _close)
        self._build_tab_export(win)
        self._update_estimate()

    def _build_tab_nguon(self, tab):
        sc = ctk.CTkScrollableFrame(tab, fg_color='transparent')
        sc.pack(fill='both', expand=True)
        head = ctk.CTkFrame(sc, fg_color='transparent')
        head.pack(fill='x', padx=10, pady=(8, 0))
        ctk.CTkLabel(
            head, text='Nguồn → ảnh → xuất (thanh trên). FILE TỔNG: ⚙ Cài đặt.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=300, justify='left', anchor='w'
        ).pack(side='left', fill='x', expand=True)
        ctk.CTkButton(
            head, text='Hướng dẫn', width=88, height=26,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            command=self._show_nguon_guide
        ).pack(side='right')
        self._build_step_source_filter(sc)
        self._build_step_photos(sc, step=2)
        extra = self._collapsible(
            sc, 'nguon_extra', 'So sánh CODE · xuất nhanh · tuỳ chọn',
            padx=8, pady=(8, 0))
        self._build_nguon_extras(extra)
        av = self._collapsible(sc, 'avatar', 'Avatar (bấm nếu cần đổi thư mục)')
        self._build_extra_avatar(av)
        if self.cloud.is_admin():
            adm = self._collapsible(
                sc, 'admin', 'Quản trị — đẩy thư mục Avatar Cloud (bấm để mở)')
            self._build_extra_admin(adm)
        self._apply_mode_ui()
        self.after(400, self._fetch_bg_list)

    def _show_nguon_guide(self):
        messagebox.showinfo(
            'Thứ tự thao tác',
            '1. Nguồn & Lọc — đồng bộ Cloud (nếu dùng), chọn kênh / tỉnh / quận.\n\n'
            '2. Thư mục ảnh — chọn folder báo cáo (kéo-thả được).\n\n'
            '3. Xuất — nút trên thanh toolbar (PPTX / ảnh đóng dấu).\n\n'
            'Mở rộng (panel ▸ So sánh CODE…):\n'
            '   · FILE SO SÁNH CODE, chia file, tùy chọn xuất\n\n'
            '⚙ Cài đặt: chọn FILE TỔNG trên máy · đẩy list Cloud\n\n'
            'GPS: tab GPS · Dấu: tab Dấu · Mẫu nền: tab Mẫu.')

    def _build_tab_data(self, tab):
        pass

    def _build_tab_cloud(self, tab):
        pass

    def _build_settings_source(self, sc):
        """Chọn FILE TỔNG trên máy & đẩy Cloud — trong Cài đặt."""
        ctk.CTkLabel(
            sc, text='FILE TỔNG / MASTER',
            font=ctk.CTkFont(size=14, weight='bold')
        ).pack(anchor='w', pady=(0, 4))
        ctk.CTkLabel(
            sc,
            text='File chuẩn có cột Code_RP. Chế độ Đám mây: dùng nút Đồng bộ ở panel Nguồn.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=520, justify='left', anchor='w'
        ).pack(anchor='w', pady=(0, 8))
        ctk.CTkLabel(
            sc, text='Trạng thái: ' + self.cloud.excel_status_text(),
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=520, justify='left', anchor='w'
        ).pack(anchor='w', pady=(0, 8))

        self.btn_excel = ctk.CTkButton(
            sc, text='Chọn FILE TỔNG trên máy…', height=36,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            command=lambda: self.pick_master_excel(purpose='local'))
        self.btn_excel.pack(fill='x', pady=(0, 4))
        self.lbl_excel = ctk.CTkLabel(
            sc, text=self._excel_path_caption(), text_color=MUTED,
            font=ctk.CTkFont(size=11), anchor='w', wraplength=520, justify='left')
        self.lbl_excel.pack(fill='x', pady=(0, 10))
        self._apply_mode_ui()

        ctk.CTkLabel(
            sc, text='Đẩy list lên Cloud',
            font=ctk.CTkFont(size=13, weight='bold')
        ).pack(anchor='w', pady=(8, 4))
        ctk.CTkButton(
            sc, text='Đẩy FILE TỔNG lên Cloud…', height=36,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=lambda: self.pick_master_excel(purpose='push')
        ).pack(fill='x', pady=(0, 6))
        if not self.cloud.logged_in():
            ctk.CTkLabel(
                sc, text='Cần đăng nhập để đồng bộ / đẩy Cloud.',
                text_color='#ca8a04', font=ctk.CTkFont(size=11),
                wraplength=520, justify='left', anchor='w'
            ).pack(anchor='w', pady=(0, 8))
        elif self.cloud.is_admin():
            ctk.CTkLabel(
                sc, text='Quản trị: có thể đẩy cho tất cả tài khoản.',
                text_color=MUTED, font=ctk.CTkFont(size=11),
                wraplength=520, justify='left', anchor='w'
            ).pack(anchor='w', pady=(0, 8))

    def _excel_path_caption(self):
        p = self.opts.get('excel_path') or ''
        if p and os.path.isfile(p):
            n = len(getattr(self.excel, 'rows', None) or [])
            ch = self.excel.channels() if n else []
            extra = f' — {n} dòng, {len(ch)} kênh' if n else ''
            return os.path.basename(p) + extra + '\n' + p
        return '(chưa chọn — dùng Cloud hoặc bấm chọn ở trên)'

    def _data_brief_text(self):
        parts = []
        n = len(getattr(self.excel, 'rows', None) or [])
        if n:
            parts.append(f'{n} dòng')
            ch = self.excel.channels()
            if ch:
                parts.append(f'{len(ch)} kênh')
        p = self.opts.get('excel_path') or ''
        if p and os.path.isfile(p):
            bn = os.path.basename(p)
            parts.append(bn[:22] + ('…' if len(bn) > 22 else ''))
        chf = self.opts.get('channel_filter', ALL_CHANNELS)
        if chf and chf != ALL_CHANNELS:
            parts.append(chf)
        city = self.opts.get('city_filter', ALL_CITIES)
        if city and city != ALL_CITIES:
            parts.append(city)
        dist = self.opts.get('district_filter', ALL_DISTRICTS)
        if dist and dist != ALL_DISTRICTS:
            parts.append(dist)
        return ' · '.join(parts) if parts else 'Chưa có FILE TỔNG — mở ⚙ Cài đặt'

    def _update_data_brief(self):
        try:
            if hasattr(self, 'lbl_data_brief'):
                self.lbl_data_brief.configure(text=self._data_brief_text())
        except Exception:
            pass
        self._refresh_filter_chips()

    def _refresh_filter_chips(self):
        frm = getattr(self, 'frm_filter_chips', None)
        if frm is None:
            return
        for w in frm.winfo_children():
            w.destroy()
        chips = []
        ch = self.opts.get('channel_filter', ALL_CHANNELS)
        if ch != ALL_CHANNELS:
            chips.append(('Kênh', ch, 'channel'))
        city = self.opts.get('city_filter', ALL_CITIES)
        if city != ALL_CITIES:
            chips.append(('Tỉnh', city, 'city'))
        dist = self.opts.get('district_filter', ALL_DISTRICTS)
        if dist != ALL_DISTRICTS:
            chips.append(('Quận', dist, 'district'))
        if not chips:
            frm.pack_forget()
            return
        frm.pack(fill='x', padx=10, pady=(2, 4))
        for label, val, kind in chips:
            short = val if len(val) <= 16 else val[:15] + '…'
            ctk.CTkButton(
                frm, text=f'{label}: {short}  ✕', height=24,
                fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                font=ctk.CTkFont(size=10),
                command=lambda k=kind: self._clear_filter(k)
            ).pack(side='left', padx=(0, 4))

    def _clear_filter(self, kind):
        if kind == 'channel':
            self.opts['channel_filter'] = ALL_CHANNELS
            try:
                self.om_channel.set(ALL_CHANNELS)
            except Exception:
                pass
            self._on_channel(ALL_CHANNELS)
        elif kind == 'city':
            self.opts['city_filter'] = ALL_CITIES
            self.opts['district_filter'] = ALL_DISTRICTS
            try:
                self.pick_city.set(ALL_CITIES)
            except Exception:
                pass
            self._on_city_filter(ALL_CITIES)
        elif kind == 'district':
            self.opts['district_filter'] = ALL_DISTRICTS
            try:
                self.pick_district.set(ALL_DISTRICTS)
            except Exception:
                pass
            self._on_district_filter(ALL_DISTRICTS)

    def _scope_estimate_text(self):
        groups = self._ordered_groups()
        if not groups:
            n = len(getattr(self.excel, 'rows', None) or [])
            if n:
                return f'List {n} dòng · chưa chọn thư mục ảnh'
            return 'Chưa có dữ liệu — đồng bộ / chọn FILE TỔNG'
        ov, n_ov = None, 0
        if self.opts.get('fx', {}).get('dashboard', True):
            ov = build_overview(groups, self._by_code, self._merged)
            n_ov = len(ov['rows'])
        ng, ns, np_ = exporter.estimate(
            groups, self.opts['n_per_slide'],
            overview_rows=n_ov,
            slides_per_file=self.opts.get('slides_per_file', 0),
            by_code=self._by_code, merged=self._merged,
            pad_blank=self.opts.get('pad_blank_slides', False))
        qa = QA.check(groups, self._by_code, self._merged)
        cap = G.file_part_limit(self.opts.get('slides_per_file', 0))
        if np_ <= 1:
            files_txt = '1 file'
        else:
            files_txt = f'{np_} file'
        txt = f'≈ {ng} nhóm · ~{ns} slide · {files_txt}'
        if qa.missing:
            txt += f' · thiếu {len(qa.missing)} ảnh'
        split = self.opts.get('split_export_by', 'none')
        if split and split != 'none':
            txt += f' · tách {split}'
        return txt

    def _show_sync_progress(self, show):
        frm = getattr(self, 'frm_sync_prog', None)
        if frm is None:
            return
        if show:
            if not frm.winfo_ismapped():
                frm.pack(fill='x', padx=10, pady=(0, 8))
        else:
            frm.pack_forget()

    def _build_step_source_filter(self, tab, step=1):
        f = self._step_card(tab, step, 'Nguồn & Lọc', '')
        self.lbl_data_brief = ctk.CTkLabel(
            f, text=self._data_brief_text(), text_color=MUTED,
            font=ctk.CTkFont(size=10), anchor='w', wraplength=268)
        self.lbl_data_brief.pack(fill='x', padx=10, pady=(0, 4))

        mode_lbl = ('Đám mây công ty' if self.opts.get('app_mode') == 'cloud'
                    else 'Máy này')
        self.seg_mode = self._seg_btn(
            f, ['Đám mây công ty', 'Máy này'], mode_lbl,
            command=self._on_app_mode,
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
        self.seg_mode.pack(fill='x', padx=10, pady=(0, 6))

        self.frm_cloud_block = ctk.CTkFrame(f, fg_color='transparent')
        self.btn_sync_cloud = ctk.CTkButton(
            self.frm_cloud_block, text='Tải FILE TỔNG + Avatar từ Cloud',
            height=34, font=ctk.CTkFont(size=12, weight='bold'),
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=lambda: self._sync_all_cloud(silent=False))
        self.btn_sync_cloud.pack(fill='x', padx=10, pady=(0, 4))
        self.lbl_excel_sync = ctk.CTkLabel(
            self.frm_cloud_block, text='📊 ' + self.cloud.excel_status_text(),
            text_color=MUTED, font=ctk.CTkFont(size=10),
            anchor='w', wraplength=268)
        self.lbl_excel_sync.pack(fill='x', padx=10)
        self.lbl_avatar_sync = ctk.CTkLabel(
            self.frm_cloud_block, text='👤 ' + self.cloud.avatar_status_text(),
            text_color=MUTED, font=ctk.CTkFont(size=10),
            anchor='w', wraplength=268)
        self.lbl_avatar_sync.pack(fill='x', padx=10, pady=(0, 4))
        self.frm_sync_prog = ctk.CTkFrame(self.frm_cloud_block, fg_color='transparent')
        self.excel_progress = ctk.CTkProgressBar(
            self.frm_sync_prog, height=6, corner_radius=4, progress_color=ACCENT)
        self.excel_progress.set(0)
        self.excel_progress.pack(fill='x', pady=(0, 2))
        self.lbl_excel_prog = ctk.CTkLabel(
            self.frm_sync_prog, text='', text_color=ACCENT,
            font=ctk.CTkFont(size=10), anchor='w')
        self.lbl_excel_prog.pack(fill='x')
        self.avatar_progress = ctk.CTkProgressBar(
            self.frm_sync_prog, height=6, corner_radius=4, progress_color=ACCENT)
        self.avatar_progress.set(0)
        self.avatar_progress.pack(fill='x', pady=(3, 2))
        self.lbl_avatar_prog = ctk.CTkLabel(
            self.frm_sync_prog, text='', text_color=ACCENT,
            font=ctk.CTkFont(size=10), anchor='w')
        self.lbl_avatar_prog.pack(fill='x', pady=(0, 6))

        self.frm_local_block = ctk.CTkFrame(f, fg_color='transparent')
        ctk.CTkButton(
            self.frm_local_block, text='⚙  Chọn FILE TỔNG trên máy…', height=30,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(size=11),
            command=self._show_settings
        ).pack(fill='x', padx=10, pady=(0, 6))

        ctk.CTkFrame(f, height=1, fg_color=BORDER).pack(fill='x', padx=10, pady=(2, 6))

        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Kênh', text_color=MUTED, width=44).pack(side='left')
        self.om_channel = ctk.CTkOptionMenu(
            row, values=[ALL_CHANNELS], width=200, height=28,
            fg_color=CARD, button_color=CARD, button_hover_color=BORDER,
            text_color=TEXT, command=self._on_channel)
        self.om_channel.pack(side='left', padx=(4, 0), fill='x', expand=True)
        row2 = ctk.CTkFrame(f, fg_color='transparent')
        row2.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row2, text='Tỉnh', text_color=MUTED, width=44).pack(side='left')
        self.pick_city = GeoScrollPicker(
            row2, [ALL_CITIES], width=168,
            command=self._on_city_filter,
            initial=self.opts.get('city_filter', ALL_CITIES))
        row3 = ctk.CTkFrame(f, fg_color='transparent')
        row3.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row3, text='Quận', text_color=MUTED, width=44).pack(side='left')
        self.pick_district = GeoScrollPicker(
            row3, [ALL_DISTRICTS], width=168,
            command=self._on_district_filter,
            initial=self.opts.get('district_filter', ALL_DISTRICTS))
        if getattr(self.excel, 'rows', None):
            self._refresh_geo_filters()

        self.frm_filter_chips = ctk.CTkFrame(f, fg_color='transparent')
        self._refresh_filter_chips()

        self.lbl_scope_est = ctk.CTkLabel(
            f, text=self._scope_estimate_text(), text_color=ACCENT,
            font=ctk.CTkFont(size=11, weight='bold'), anchor='w', wraplength=268)
        self.lbl_scope_est.pack(fill='x', padx=10, pady=(4, 2))
        ctk.CTkButton(
            f, text='⚙  Cài đặt FILE TỔNG & Cloud…', height=24,
            fg_color='transparent', hover_color=BORDER, text_color=MUTED,
            font=ctk.CTkFont(size=10),
            command=self._show_settings
        ).pack(fill='x', padx=10, pady=(0, 8))
        self._apply_mode_ui()

    def _build_step_photos(self, tab, step=2):
        f = self._step_card(tab, step, 'Thư mục ảnh', '')
        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkButton(row, text='Thêm thư mục ảnh…', height=40, fg_color=ACCENT,
                      hover_color=ACCENT_HOVER,
                      command=self.add_image_folder).pack(
                          side='left', expand=True, fill='x')
        ctk.CTkButton(row, text='Xoá hết', width=70, height=40, fg_color=CARD,
                      hover_color=BORDER, text_color=TEXT,
                      command=self.clear_image_folders).pack(side='left', padx=(6, 0))
        self.sw_subfolders = ctk.CTkSwitch(
            f, text='Gồm cả thư mục con', progress_color=ACCENT,
            command=self._on_scan_subfolders)
        if self.opts.get('scan_subfolders', True):
            self.sw_subfolders.select()
        self.sw_subfolders.pack(anchor='w', padx=10, pady=(6, 2))
        self.frm_folders = ctk.CTkFrame(f, fg_color=CARD, corner_radius=8)
        self.frm_folders.pack(fill='x', padx=10, pady=(0, 4))
        self.lbl_images = self._path_label(f)
        self._refresh_folder_list()
        cov = self._collapsible(
            f, 'coverage', 'Kiểm tra thiếu ảnh · sửa mã',
            padx=10, pady=(0, 8))
        ctk.CTkButton(
            cov, text='So sánh ẢNH với FILE TỔNG…', height=32,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            command=self._show_coverage
        ).pack(fill='x', padx=10, pady=(8, 4))
        ctk.CTkButton(
            cov, text='Sửa mã ảnh / đổi tên file…', height=32,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self._show_fix_codes_dialog
        ).pack(fill='x', padx=10, pady=(0, 4))
        self.lbl_compare_list = ctk.CTkLabel(
            cov, text=self._current_list_caption(),
            text_color=MUTED, font=ctk.CTkFont(size=10),
            wraplength=260, justify='left', anchor='w')
        self.lbl_compare_list.pack(fill='x', padx=10, pady=(0, 10))

    def _build_nguon_extras(self, body):
        ctk.CTkLabel(
            body, text='FILE SO SÁNH CODE (tuỳ chọn)',
            font=ctk.CTkFont(size=12, weight='bold'), anchor='w'
        ).pack(fill='x', padx=10, pady=(10, 4))
        ctk.CTkButton(
            body, text='Chọn FILE SO SÁNH CODE…', height=34,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            command=self._choose_list_for_master_compare
        ).pack(fill='x', padx=10, pady=(0, 4))
        self.lbl_list_master_compare = ctk.CTkLabel(
            body, text=self._list_master_compare_caption(),
            text_color=MUTED, font=ctk.CTkFont(size=10),
            wraplength=260, justify='left', anchor='w')
        self.lbl_list_master_compare.pack(fill='x', padx=10, pady=(0, 10))

        ctk.CTkLabel(
            body, text='Xuất nhanh',
            font=ctk.CTkFont(size=12, weight='bold'), anchor='w'
        ).pack(fill='x', padx=10, pady=(0, 4))
        row = ctk.CTkFrame(body, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(0, 4))
        ctk.CTkButton(
            row, text='Xuất PPTX', height=32,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.export
        ).pack(side='left', expand=True, fill='x')
        ctk.CTkButton(
            row, text='Xuất ảnh', height=32,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            command=self.export_stamped
        ).pack(side='left', padx=(6, 0), expand=True, fill='x')
        row2 = ctk.CTkFrame(body, fg_color='transparent')
        row2.pack(fill='x', padx=10, pady=(0, 4))
        ctk.CTkLabel(row2, text='Chia file', text_color=MUTED).pack(side='left')
        self.om_spf_step = ctk.CTkOptionMenu(
            row2,
            values=G.slides_per_file_menu_values(self.opts.get('slides_per_file', 0)),
            width=128, height=26, fg_color=CARD,
            button_color=CARD, button_hover_color=BORDER,
            text_color=TEXT, command=self._on_spf)
        self.om_spf_step.set(G.slides_per_file_label(self.opts.get('slides_per_file', 0)))
        self.om_spf_step.pack(side='left', padx=8)
        ctk.CTkButton(
            body, text='Tùy chọn xuất đầy đủ…', height=30,
            fg_color='transparent', hover_color=BORDER, text_color=MUTED,
            font=ctk.CTkFont(size=11),
            command=self._open_export_opts
        ).pack(fill='x', padx=10, pady=(0, 10))

    def _build_extra_avatar(self, extra):
        ctk.CTkLabel(
            extra,
            text='Đám mây đã lấy avatar lúc đồng bộ. Chỉ chọn thư mục khi làm trên máy này.',
            text_color=MUTED, font=ctk.CTkFont(size=11), wraplength=260,
            justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(10, 4))
        self.btn_avatars = ctk.CTkButton(
            extra, text='Chọn thư mục avatar…', height=30,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self.select_avatars)
        self.btn_avatars.pack(fill='x', padx=10, pady=2)
        self.lbl_avatars = self._path_label(extra)

    def _build_tab_mau(self, tab):
        """Tab Mẫu: header gọn, gallery chiếm hết chiều cao."""
        self.seg_bg = self._seg_btn(
            tab, ['Trắng', 'Ảnh nền'],
            'Trắng' if self.opts['bg_mode'] == 'white' else 'Ảnh nền',
            command=self._on_bg_mode,
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
        self.seg_bg.pack(fill='x', padx=10, pady=(10, 4))

        row = ctk.CTkFrame(tab, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(0, 4))
        ctk.CTkButton(
            row, text='Tải lên Cloud', height=32,
            font=ctk.CTkFont(size=12, weight='bold'),
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=self._upload_my_bg
        ).pack(side='left', expand=True, fill='x')
        ctk.CTkButton(
            row, text='Làm mới', width=78, height=32, fg_color=CARD,
            hover_color=BORDER, text_color=TEXT,
            command=self._fetch_bg_list
        ).pack(side='left', padx=(6, 0))
        ctk.CTkButton(
            row, text='PNG máy', width=84, height=32,
            fg_color=CARD, hover_color=BORDER, text_color=TEXT,
            command=self.select_bg
        ).pack(side='left', padx=(6, 0))

        if self.cloud.is_admin():
            ctk.CTkButton(
                tab, text='Đẩy thư mục mẫu chung…', height=28,
                fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                command=lambda: self._upload_bg_dialog(shared=True)
            ).pack(fill='x', padx=10, pady=(0, 4))

        self.lbl_bg = ctk.CTkLabel(
            tab, text='', text_color=MUTED, font=ctk.CTkFont(size=11),
            anchor='w', wraplength=320, justify='left')
        self.lbl_bg.pack(fill='x', padx=10, pady=(0, 2))

        meta = ctk.CTkFrame(tab, fg_color='transparent')
        meta.pack(fill='x', padx=10, pady=(2, 4))
        ctk.CTkLabel(
            meta, text='Kho mẫu',
            font=ctk.CTkFont(size=12, weight='bold'),
            text_color=TEXT
        ).pack(side='left')
        self.lbl_bg_sync = ctk.CTkLabel(
            meta, text='', text_color=MUTED, font=ctk.CTkFont(size=11),
            anchor='e', wraplength=220)
        self.lbl_bg_sync.pack(side='right', fill='x', expand=True, padx=(8, 0))

        self.bg_gallery = ctk.CTkScrollableFrame(tab, fg_color=CARD, corner_radius=10)
        self.bg_gallery.pack(fill='both', expand=True, padx=8, pady=(0, 10))
        self._bg_photos = []

    def _build_extra_admin(self, extra):
        pad = dict(fill='x', padx=10, pady=2)
        ctk.CTkLabel(
            extra, text='Chỉ quản trị: đẩy kho avatar / nền chung cho cả công ty.',
            text_color=MUTED, font=ctk.CTkFont(size=11), wraplength=260,
            justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(10, 4))
        ctk.CTkButton(extra, text='Đẩy thư mục Avatar…', height=30,
                      fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                      command=self._upload_avatars_dialog).pack(**pad)
        ctk.CTkLabel(extra, text='Đẩy mẫu nền chung: tab Mẫu.',
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     wraplength=260, justify='left', anchor='w'
                     ).pack(fill='x', padx=10, pady=(2, 10))

    def _build_tab_slide(self, tab):
        bar = ctk.CTkFrame(tab, fg_color='transparent')
        bar.pack(fill='x', padx=8, pady=(8, 0))
        ctk.CTkLabel(
            bar, text='Khung',
            font=ctk.CTkFont(size=13, weight='bold'), text_color=TEXT
        ).pack(side='left')
        ctk.CTkButton(
            bar, text='Đóng', width=58, height=26,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=self._hide_inspector
        ).pack(side='right')
        sc = ctk.CTkScrollableFrame(tab, fg_color='transparent')
        sc.pack(fill='both', expand=True)
        self._slide_scroll = sc
        ctk.CTkLabel(
            sc, text='Kéo trên slide hoặc sửa số X/Y/W/H. Bấm Khung / Đóng để ẩn.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=220, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(8, 0))
        self._build_tab_elems(sc)
        self._sync_ctx_panels('image')

    def _build_tab_elems(self, tab):
        f = self._card(tab, 'Phần tử')
        self.om_elem = ctk.CTkOptionMenu(
            f, values=[ELEMENT_LABELS[n] for n in ELEMENT_NAMES],
            height=28, fg_color=CARD, button_color=CARD,
            button_hover_color=BORDER, text_color=TEXT,
            command=self._on_pick_elem)
        self.om_elem.pack(fill='x', padx=10, pady=2)

        grid = ctk.CTkFrame(f, fg_color='transparent')
        grid.pack(fill='x', padx=10, pady=6)
        self._insp = {}
        for i, key in enumerate(('x', 'y', 'w', 'h')):
            ctk.CTkLabel(grid, text=key.upper(), width=18,
                         text_color=MUTED).grid(row=i // 2, column=(i % 2) * 2,
                                                padx=(0, 4), pady=2)
            e = ctk.CTkEntry(grid, width=92, height=26, fg_color=CARD)
            e.grid(row=i // 2, column=(i % 2) * 2 + 1, padx=(0, 10), pady=2)
            e.bind('<Return>', self._apply_inspector)
            e.bind('<FocusOut>', self._apply_inspector)
            self._insp[key] = e

        al = ctk.CTkFrame(f, fg_color='transparent')
        al.pack(fill='x', padx=10, pady=(0, 6))
        for txt, mode in (('⇤', 'left'), ('↔', 'cx'), ('⇥', 'right'),
                          ('⤒', 'top'), ('↕', 'cy'), ('⤓', 'bottom')):
            ctk.CTkButton(al, text=txt, width=34, height=26, fg_color=CARD,
                          hover_color=BORDER, text_color=TEXT,
                          command=lambda m=mode: self.editor.align_selected(m)
                          ).pack(side='left', padx=2)

        sw = ctk.CTkFrame(f, fg_color='transparent')
        sw.pack(fill='x', padx=10, pady=(0, 8))
        self.sw_lock = ctk.CTkSwitch(sw, text='Khoá', width=70,
                                     command=self._toggle_lock,
                                     progress_color=ACCENT)
        self.sw_lock.pack(side='left')
        self.sw_hide = ctk.CTkSwitch(sw, text='Ẩn', width=64,
                                     command=self._toggle_hide,
                                     progress_color=ACCENT)
        self.sw_hide.pack(side='left', padx=10)

        self.frm_type_text = ctk.CTkFrame(tab, fg_color='transparent')
        self._build_live_style(self.frm_type_text)

        self.frm_title_extra = ctk.CTkFrame(tab, fg_color='transparent')
        self._build_title_extra(self.frm_title_extra)
        self.frm_channel_extra = ctk.CTkFrame(tab, fg_color='transparent')
        self._build_channel_extra(self.frm_channel_extra)

        self.frm_type_image = ctk.CTkFrame(tab, fg_color='transparent')
        f = self._card(self.frm_type_image, 'Vùng ảnh')
        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Ảnh/slide:', text_color=MUTED).pack(side='left')
        self.seg_n = self._seg_btn(
            row, ['Tự', '1', '2', '3', '4'],
            G.n_per_slide_label(self.opts['n_per_slide']),
            command=self._on_n, width=190,
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
        self.seg_n.pack(side='left', padx=8)
        ctk.CTkLabel(
            f, text='Tự: số ô = số ảnh (tối đa 4). 1 ảnh lấp cả vùng, không để khung trống.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=240, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(0, 2))
        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Tỉ lệ ô:', text_color=MUTED).pack(side='left')
        self.om_ar = ctk.CTkOptionMenu(row, values=list(AR_CHOICES),
                                       width=110, height=26, fg_color=CARD,
                                       button_color=CARD,
                                       button_hover_color=BORDER,
                                       text_color=TEXT, command=self._on_ar)
        self.om_ar.set(self.opts['ar_label'])
        self.om_ar.pack(side='left', padx=8)
        self.seg_fit = self._seg_btn(
            f, ['Lấp đầy', 'Vừa khung'],
            'Lấp đầy' if self.layout['image'].get('fit_mode', 'fill') == 'fill'
            else 'Vừa khung',
            command=self._on_fit,
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER)
        self.seg_fit.pack(fill='x', padx=10, pady=4)
        ctk.CTkLabel(f, text='Ảnh dọc giữ tỉ lệ gốc (không cắt 4:3).',
                     text_color=MUTED).pack(fill='x', padx=10, pady=(0, 2))
        self.sl_gap = self._slider(f, 'Khe hở ảnh', 0, 0.4,
                                   self.layout['image'].get('gap', 0.1),
                                   self._on_gap, fmt='{:.2f}″')
        self.sl_radius = self._slider(f, 'Bo góc ảnh', 0, 50,
                                      self.layout['image'].get('radius', 0),
                                      self._on_radius, fmt='{:.0f}%')
        self.sl_img_opacity = self._slider(
            f, 'Độ mờ ảnh', 15, 100,
            self.layout['image'].get('opacity', 100),
            lambda v: self._set_elem_key('image', 'opacity', int(float(v))),
            fmt='{:.0f}%')

        self.frm_type_avatar = ctk.CTkFrame(tab, fg_color='transparent')
        f = self._card(self.frm_type_avatar, 'Avatar')
        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Tỉ lệ khung:', text_color=MUTED).pack(side='left')
        self.om_av_ar = ctk.CTkOptionMenu(row, values=['4:3', '1:1', '3:4'],
                                          width=90, height=26, fg_color=CARD,
                                          button_color=CARD,
                                          button_hover_color=BORDER,
                                          text_color=TEXT,
                                          command=self._on_avatar_ar)
        cur = self.layout['avatar'].get('ar', 4 / 3)
        self.om_av_ar.set('1:1' if abs(cur - 1) < .01
                          else ('3:4' if cur < 1 else '4:3'))
        self.om_av_ar.pack(side='left', padx=8)
        self.sl_av_radius = self._slider(f, 'Bo góc avatar', 0, 50,
                                         self.layout['avatar'].get('radius', 0),
                                         self._on_av_radius, fmt='{:.0f}%')
        self.sl_av_opacity = self._slider(
            f, 'Độ mờ avatar', 15, 100,
            self.layout['avatar'].get('opacity', 100),
            lambda v: self._set_elem_key('avatar', 'opacity', int(float(v))),
            fmt='{:.0f}%')

    def _slider(self, parent, label, lo, hi, val, cmd, fmt='{:.0f}'):
        row = ctk.CTkFrame(parent, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(2, 6))
        ctk.CTkLabel(row, text=label, text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor='w')
        inner = ctk.CTkFrame(row, fg_color='transparent')
        inner.pack(fill='x')
        vl = ctk.CTkLabel(inner, text=fmt.format(val), width=48,
                          text_color=TEXT, font=ctk.CTkFont(size=11))
        vl.pack(side='right')
        s = ctk.CTkSlider(inner, from_=lo, to=hi, height=14,
                          progress_color=ACCENT,
                          command=lambda v: (vl.configure(text=fmt.format(v)),
                                             cmd(v)))
        s.set(val)
        s.pack(side='left', fill='x', expand=True, padx=(0, 6))
        return s

    def _fx_sw(self, parent, key, text):
        fx = self.opts['fx']
        w = ctk.CTkSwitch(parent, text=text, progress_color=ACCENT)
        w.configure(command=lambda *_a, k=key, ww=w: self._set_fx(k, bool(ww.get())))
        if fx.get(key):
            w.select()
        w.pack(anchor='w', padx=10, pady=(6, 2))
        setattr(self, 'sw_fx_' + key, w)
        return w

    def _fx_om(self, parent, key, values, label):
        fx = self.opts['fx']
        ctk.CTkLabel(parent, text=label, text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor='w', padx=10, pady=(6, 0))
        m = ctk.CTkOptionMenu(
            parent, values=values, height=26, fg_color=CARD,
            button_color=CARD, button_hover_color=BORDER, text_color=TEXT,
            command=lambda v, k=key: self._set_fx(k, v))
        cur = str(fx.get(key, values[0]))
        m.set(cur if cur in values else values[0])
        m.pack(fill='x', padx=10, pady=2)
        return m

    def _build_tab_gps(self, tab):
        """Tab GPS: địa điểm trên ảnh + mini-map."""
        sc = ctk.CTkScrollableFrame(tab, fg_color='transparent')
        sc.pack(fill='both', expand=True)
        fx = self.opts['fx']
        ctk.CTkLabel(
            sc, text='Chip GPS / Map trên thanh canvas bật tắt nhanh. Chế độ chi tiết ở đây.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=300, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(8, 0))

        f = self._card(sc, 'Địa điểm trên ảnh')
        self._fx_sw(f, 'stamp_location', 'In địa điểm (đường / quận / tỉnh) lên ảnh')
        self._loc_mode_labels = {
            'gps': 'Theo toạ độ EXIF',
            'auto': 'Mã Excel → GPS',
            'excel': 'Theo mã Excel',
            'manual': 'Gõ tay',
        }
        self._loc_mode_vals = {v: k for k, v in self._loc_mode_labels.items()}
        cur_mode = fx.get('location_mode', 'gps')
        ctk.CTkLabel(
            f, text='Chế độ dòng địa điểm', text_color=MUTED,
            font=ctk.CTkFont(size=11)
        ).pack(anchor='w', padx=10, pady=(2, 0))
        self.om_loc_mode = ctk.CTkOptionMenu(
            f, values=list(self._loc_mode_labels.values()),
            command=self._on_loc_mode,
            fg_color=INPUT, button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=CARD, text_color=TEXT)
        self.om_loc_mode.set(self._loc_mode_labels.get(cur_mode, 'Theo toạ độ EXIF'))
        self.om_loc_mode.pack(fill='x', padx=10, pady=(0, 4))
        self.lbl_loc_hint = ctk.CTkLabel(
            f, text='', text_color=MUTED, font=ctk.CTkFont(size=10),
            wraplength=280, justify='left', anchor='w')
        self.lbl_loc_hint.pack(fill='x', padx=10, pady=(0, 4))
        self.ent_ts_loc = ctk.CTkEntry(f, height=28, fg_color=CARD,
                                       placeholder_text='VD: Q.1, TP.HCM')
        self.ent_ts_loc.insert(0, fx.get('location_text', ''))
        self.ent_ts_loc.pack(fill='x', padx=10, pady=(0, 4))
        self.ent_ts_loc.bind('<KeyRelease>', lambda e: self._set_fx(
            'location_text', self.ent_ts_loc.get()))
        self._fx_sw(f, 'show_gps', 'In thêm toạ độ GPS (nếu ảnh có)')
        self._sync_loc_hint()

        f = self._card(sc, 'Mini-map GPS')
        self._fx_sw(f, 'minimap', 'Dán bản đồ mini góc ảnh (OpenStreetMap)')
        ctk.CTkLabel(
            f, text='Cần EXIF GPS trên file ảnh. Lần đầu tải map cần mạng, lần sau dùng cache.',
            text_color=MUTED, font=ctk.CTkFont(size=10), wraplength=280,
            justify='left').pack(anchor='w', padx=10, pady=(0, 4))
        self.sl_mm_op = self._slider(
            f, 'Độ hòa trộn map', 30, 100,
            int(fx.get('minimap_opacity', 85)),
            lambda v: self._set_fx('minimap_opacity', int(float(v))), fmt='{:.0f}%')

    def _build_tab_stamp(self, tab):
        """Tab Dấu: ngày giờ, logo, watermark, AI."""
        sc = ctk.CTkScrollableFrame(tab, fg_color='transparent')
        sc.pack(fill='both', expand=True)
        fx = self.opts['fx']
        sw = self._fx_sw
        om = self._fx_om

        ctk.CTkLabel(
            sc, text='Chip Ngày trên thanh canvas. GPS / map nằm ở tab GPS.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=300, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(8, 0))

        f = self._card(sc, 'Đóng dấu ngày / giờ')
        sw(f, 'use_timestamp', 'In ngày giờ lên ảnh')
        om(f, 'display_mode', ['Ngày + Giờ', 'Chỉ ngày', 'Chỉ giờ'], 'Hiển thị')
        om(f, 'date_format', ['DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD', 'YYYY/MM/DD'],
           'Định dạng ngày')
        om(f, 'time_format', ['24h', '24h + giây', '12h AM/PM', '12h + giây'],
           'Định dạng giờ')
        om(f, 'position',
           ['bottom_right', 'bottom_left', 'top_right', 'top_left'],
           'Vị trí')

        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=4)
        self.lbl_ts_color = self._make_color_chip(
            row, fx.get('text_color', '#FF9900'), 36,
            command=self._pick_ts_color)
        self.lbl_ts_color.pack(side='left', padx=(0, 8))
        ctk.CTkLabel(row, text='Màu chữ đóng dấu', text_color=TEXT,
                     font=ctk.CTkFont(size=12)).pack(side='left')

        self.sl_ts_size = self._slider(
            f, 'Cỡ chữ (% chiều cao ảnh)', 1.0, 10.0,
            float(fx.get('font_scale', 3.5)),
            lambda v: self._set_fx('font_scale', float(v)), fmt='{:.1f}%')
        sw(f, 'add_background', 'Nền mờ sau chữ')
        self.sl_ts_bg = self._slider(
            f, 'Độ đậm nền (0–255)', 0, 255, int(fx.get('bg_opacity', 100)),
            lambda v: self._set_fx('bg_opacity', int(float(v))), fmt='{:.0f}')
        self.sl_ts_ls = self._slider(
            f, 'Giãn dòng', 1.0, 2.5, float(fx.get('line_spacing', 1.2)),
            lambda v: self._set_fx('line_spacing', float(v)), fmt='{:.1f}x')
        self.sl_ts_let = self._slider(
            f, 'Giãn chữ', 0, 10, int(fx.get('letter_spacing', 0)),
            lambda v: self._set_fx('letter_spacing', int(float(v))), fmt='{:.0f}px')
        sw(f, 'use_exif', 'Lấy ngày từ EXIF (ngày chụp)')

        extra_fix = self._collapsible(
            sc, 'stamp_fix', 'Ngày / giờ cố định', start_open=True)
        row = ctk.CTkFrame(extra_fix, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(8, 2))
        self.sw_fix_date = ctk.CTkSwitch(
            row, text='Ngày', width=70, progress_color=ACCENT,
            command=lambda: self._set_fx('fix_date', bool(self.sw_fix_date.get())))
        if fx.get('fix_date'):
            self.sw_fix_date.select()
        self.sw_fix_date.pack(side='left')
        self.ent_fix_date = ctk.CTkEntry(row, width=110, height=26, fg_color=CARD)
        self.ent_fix_date.insert(0, fx.get('fix_date_val', ''))
        self.ent_fix_date.pack(side='left', padx=6)
        self.ent_fix_date.bind('<KeyRelease>', lambda e: self._set_fx(
            'fix_date_val', self.ent_fix_date.get()))
        row = ctk.CTkFrame(extra_fix, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(2, 4))
        self.sw_fix_time = ctk.CTkSwitch(
            row, text='Giờ', width=70, progress_color=ACCENT,
            command=lambda: self._set_fx('fix_time', bool(self.sw_fix_time.get())))
        if fx.get('fix_time'):
            self.sw_fix_time.select()
        self.sw_fix_time.pack(side='left')
        self.ent_fix_time = ctk.CTkEntry(row, width=110, height=26, fg_color=CARD)
        self.ent_fix_time.insert(0, fx.get('fix_time_val', ''))
        self.ent_fix_time.pack(side='left', padx=6)
        self.ent_fix_time.bind('<KeyRelease>', lambda e: self._set_fx(
            'fix_time_val', self.ent_fix_time.get()))
        ctk.CTkLabel(extra_fix, text='Định dạng: 17/08/2026 và 14:30',
                     text_color=MUTED, font=ctk.CTkFont(size=10)
                     ).pack(anchor='w', padx=10, pady=(0, 8))

        extra_rand = self._collapsible(
            sc, 'stamp_random', 'Random khi không EXIF (Zalo, Messenger…)',
            start_open=bool(fx.get('random_no_exif')))
        sw(extra_rand, 'random_no_exif',
           'Random ngày/giờ nếu ảnh không có EXIF ngày chụp')
        ctk.CTkLabel(
            extra_rand,
            text='Ảnh gửi qua Zalo/Messenger thường mất EXIF. Mỗi ảnh nhận giờ khác nhau '
                 'trong khoảng bạn chọn (cùng file → cùng giờ mỗi lần xuất).',
            text_color=MUTED, font=ctk.CTkFont(size=10), wraplength=300,
            justify='left'
        ).pack(anchor='w', padx=10, pady=(0, 6))
        row = ctk.CTkFrame(extra_rand, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Giờ từ', text_color=MUTED, width=52).pack(side='left')
        self.ent_rand_t0 = ctk.CTkEntry(row, width=72, height=26, fg_color=CARD,
                                        placeholder_text='09:00')
        self.ent_rand_t0.insert(0, fx.get('random_time_from', '09:00'))
        self.ent_rand_t0.pack(side='left', padx=(0, 8))
        self.ent_rand_t0.bind('<KeyRelease>', lambda e: self._set_fx(
            'random_time_from', self.ent_rand_t0.get()))
        ctk.CTkLabel(row, text='đến', text_color=MUTED).pack(side='left')
        self.ent_rand_t1 = ctk.CTkEntry(row, width=72, height=26, fg_color=CARD,
                                        placeholder_text='17:00')
        self.ent_rand_t1.insert(0, fx.get('random_time_to', '17:00'))
        self.ent_rand_t1.pack(side='left', padx=(8, 0))
        self.ent_rand_t1.bind('<KeyRelease>', lambda e: self._set_fx(
            'random_time_to', self.ent_rand_t1.get()))
        sw(extra_rand, 'random_date', 'Random cả ngày trong khoảng')
        row = ctk.CTkFrame(extra_rand, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Ngày từ', text_color=MUTED, width=52).pack(side='left')
        self.ent_rand_d0 = ctk.CTkEntry(row, width=100, height=26, fg_color=CARD,
                                        placeholder_text='01/08/2026')
        self.ent_rand_d0.insert(0, fx.get('random_date_from', ''))
        self.ent_rand_d0.pack(side='left', padx=(0, 8))
        self.ent_rand_d0.bind('<KeyRelease>', lambda e: self._set_fx(
            'random_date_from', self.ent_rand_d0.get()))
        ctk.CTkLabel(row, text='đến', text_color=MUTED).pack(side='left')
        self.ent_rand_d1 = ctk.CTkEntry(row, width=100, height=26, fg_color=CARD,
                                        placeholder_text='31/08/2026')
        self.ent_rand_d1.insert(0, fx.get('random_date_to', ''))
        self.ent_rand_d1.pack(side='left', padx=(8, 0))
        self.ent_rand_d1.bind('<KeyRelease>', lambda e: self._set_fx(
            'random_date_to', self.ent_rand_d1.get()))
        ctk.CTkLabel(
            extra_rand,
            text='Tắt «Random cả ngày» → giữ ngày cố định (ở trên) hoặc ngày file.',
            text_color=MUTED, font=ctk.CTkFont(size=10), wraplength=300,
            justify='left'
        ).pack(anchor='w', padx=10, pady=(0, 8))

        extra_logo = self._collapsible(
            sc, 'stamp_logo', 'Logo đóng lên ảnh', start_open=True)
        sw(extra_logo, 'logo_enable', 'Dán logo lên ảnh')
        row = ctk.CTkFrame(extra_logo, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkButton(row, text='Chọn file…', width=90, height=26, fg_color=CARD,
                      hover_color=BORDER, text_color=TEXT,
                      command=self._choose_logo).pack(side='left')
        self.lbl_logo = ctk.CTkLabel(
            row, text=os.path.basename(fx.get('logo_path') or '') or '(chưa chọn)',
            text_color=MUTED, font=ctk.CTkFont(size=11), anchor='w')
        self.lbl_logo.pack(side='left', padx=8, fill='x', expand=True)
        om(extra_logo, 'logo_position',
           ['bottom_right', 'bottom_left', 'top_right', 'top_left'], 'Vị trí logo')
        self.sl_logo_sz = self._slider(
            extra_logo, 'Cỡ logo', 2, 35, int(fx.get('logo_size_pct', 8)),
            lambda v: self._set_fx('logo_size_pct', int(float(v))), fmt='{:.0f}%')
        self.sl_logo_op = self._slider(
            extra_logo, 'Độ mờ logo', 10, 100, int(fx.get('logo_opacity', 50)),
            lambda v: self._set_fx('logo_opacity', int(float(v))), fmt='{:.0f}%')

        extra_ai = self._collapsible(
            sc, 'stamp_ai', 'Thông minh — chỉnh sáng, quét ảnh, overview',
            start_open=True)
        ctk.CTkLabel(
            extra_ai, text='Mini-map và địa điểm nằm ở tab GPS bên trái.',
            text_color=MUTED, font=ctk.CTkFont(size=10), wraplength=280,
            justify='left').pack(anchor='w', padx=10, pady=(8, 4))
        sw(extra_ai, 'auto_enhance', 'Tự chỉnh sáng / nét')
        sw(extra_ai, 'quality_check', 'Quét chất lượng trước khi xuất')
        ctk.CTkButton(extra_ai, text='Quét thư mục ảnh ngay', height=26, fg_color=CARD,
                      hover_color=BORDER, text_color=TEXT,
                      command=self._scan_quality_now).pack(fill='x', padx=10, pady=(0, 6))
        sw(extra_ai, 'dashboard', 'Slide Overview đầu báo cáo')
        ctk.CTkLabel(
            extra_ai,
            text='Trang đầu: số ảnh đã up đối chiếu số màn hình (LCD+DP+DS+GP).',
            text_color=MUTED, font=ctk.CTkFont(size=10), wraplength=260,
            justify='left').pack(anchor='w', padx=10, pady=(0, 8))

        extra_wm = self._collapsible(
            sc, 'stamp_wm', 'Watermark & đổ bóng', start_open=True)
        sw(extra_wm, 'shadow', 'Đổ bóng dưới ảnh')
        ctk.CTkLabel(extra_wm, text='Chữ mờ góc ảnh', text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor='w', padx=10, pady=(6, 0))
        self.ent_wm = ctk.CTkEntry(extra_wm, height=28, fg_color=CARD,
                                   placeholder_text='VD: GOLDEN ASIA')
        self.ent_wm.insert(0, fx.get('watermark', ''))
        self.ent_wm.pack(fill='x', padx=10, pady=(0, 10))
        self.ent_wm.bind('<KeyRelease>', lambda e: self._set_fx(
            'watermark', self.ent_wm.get()))

    def _build_live_style(self, parent):
        """Chỉnh font / cỡ / màu / độ mờ của ĐÚNG phần tử đang chọn."""
        box = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=8)
        box.pack(fill='x', padx=8, pady=(4, 10))
        self.lbl_style_for = ctk.CTkLabel(
            box, text='Kiểu chữ',
            font=ctk.CTkFont(size=12, weight='bold'), text_color=TEXT,
            anchor='w', wraplength=220, justify='left')
        self.lbl_style_for.pack(fill='x', padx=10, pady=(8, 2))

        self.btn_font = ctk.CTkButton(
            box, text=self.layout.get('font', {}).get('name', 'Arial'),
            height=30, fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            anchor='w', command=self._open_font_dialog)
        self.btn_font.pack(fill='x', padx=10, pady=(4, 2))
        self.lbl_font_warn = ctk.CTkLabel(
            box, text='', text_color='#f59e0b', font=ctk.CTkFont(size=11),
            anchor='w', wraplength=220, justify='left')
        self.lbl_font_warn.pack(fill='x', padx=10)

        row = ctk.CTkFrame(box, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(2, 4))
        self.sw_bold = ctk.CTkSwitch(row, text='Đậm', width=70,
                                     command=self._on_style,
                                     progress_color=ACCENT)
        if self.layout['font'].get('bold', True):
            self.sw_bold.select()
        self.sw_bold.pack(side='left')
        self.sw_italic = ctk.CTkSwitch(row, text='Nghiêng', width=90,
                                       command=self._on_style,
                                       progress_color=ACCENT)
        if self.layout['font'].get('italic', False):
            self.sw_italic.select()
        self.sw_italic.pack(side='left', padx=8)

        self.sl_elem_size = self._slider(
            box, 'Cỡ chữ', 7, 48, 16, self._on_elem_size, fmt='{:.0f}pt')
        self.sl_elem_opacity = self._slider(
            box, 'Độ mờ', 15, 100, 100, self._on_elem_opacity,
            fmt='{:.0f}%')
        self.sl_elem_track = self._slider(
            box, 'Giãn chữ', 0, 8, 0, self._on_elem_track, fmt='{:.1f}pt')

        row = ctk.CTkFrame(box, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(0, 4))
        self.chip_live = self._make_color_chip(
            row, self._elem_color_hex('title'), 32,
            command=self._on_elem_color)
        self.chip_live.pack(side='left', padx=(0, 8))
        ctk.CTkLabel(row, text='Màu chữ', text_color=TEXT,
                     font=ctk.CTkFont(size=12)).pack(side='left')

        al0 = {'left': 'Trái', 'center': 'Giữa', 'right': 'Phải'}.get(
            self.layout.get('title', {}).get('align', 'left'), 'Trái')
        self.seg_elem_align = self._seg_btn(
            box, ['Trái', 'Giữa', 'Phải'], al0,
            command=self._on_elem_align, selected_color=ACCENT,
            selected_hover_color=ACCENT_HOVER)
        self.seg_elem_align.pack(fill='x', padx=10, pady=(2, 8))

        row = ctk.CTkFrame(box, fg_color='transparent')
        row.pack(fill='x', padx=8, pady=(0, 8))
        self._color_chips = {}
        for label, key in (('Tiêu đề', 'title'), ('Thân', 'body'),
                           ('Traffic', 'accent'), ('Kênh', 'channel')):
            cell = ctk.CTkFrame(row, fg_color='transparent')
            cell.pack(side='left', padx=4)
            chip = self._make_color_chip(
                cell, self._role_color_hex(key), 28)
            chip.configure(command=lambda k=key, w=chip: self._pick_color(k, near=w))
            chip.pack()
            self._color_chips[key] = chip
            ctk.CTkLabel(cell, text=label, text_color=MUTED,
                         font=ctk.CTkFont(size=10)).pack()
        self._color_btns = self._color_chips

    def _build_title_extra(self, tab):
        f = self._card(tab, 'Tuỳ chọn tiêu đề')
        row = ctk.CTkFrame(f, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=(0, 8))
        self.sw_upper = ctk.CTkSwitch(row, text='VIẾT HOA', width=100,
                                      command=self._on_title_flags,
                                      progress_color=ACCENT)
        if self.layout['title'].get('upper', True):
            self.sw_upper.select()
        self.sw_upper.pack(side='left')
        self.sw_counter = ctk.CTkSwitch(row, text='Số trang (1/3)', width=120,
                                        command=self._on_title_flags,
                                        progress_color=ACCENT)
        if self.layout['title'].get('show_counter', False):
            self.sw_counter.select()
        self.sw_counter.pack(side='left', padx=6)
        self.sw_fullw = ctk.CTkSwitch(f, text='Căn theo cả slide (trải ngang)',
                                      command=self._on_title_flags,
                                      progress_color=ACCENT)
        if self.layout['title'].get('full_width_on_slide', True):
            self.sw_fullw.select()
        self.sw_fullw.pack(anchor='w', padx=10, pady=(0, 8))

    def _build_channel_extra(self, tab):
        f = self._card(tab, 'Tuỳ chọn kênh')
        self.sw_channel = ctk.CTkSwitch(f, text='Hiện phần tử Kênh',
                                        command=self._on_channel_toggle,
                                        progress_color=ACCENT)
        if self.opts['channel_enabled']:
            self.sw_channel.select()
        self.sw_channel.pack(anchor='w', padx=10, pady=2)
        self.ent_ch_tpl = ctk.CTkEntry(f, height=26, fg_color=CARD)
        self.ent_ch_tpl.insert(0, self.opts['channel_template'])
        self.ent_ch_tpl.pack(fill='x', padx=10, pady=4)
        self.ent_ch_tpl.bind('<KeyRelease>', self._on_channel_tpl)

    def _build_tab_export(self, tab):
        ctk.CTkLabel(
            tab,
            text='Checklist và cảnh báo list — không ghi lên slide.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=360, justify='left', anchor='w'
        ).pack(fill='x', padx=12, pady=(10, 0))
        f = self._card(tab, 'Ước tính')
        self.lbl_est_full = ctk.CTkLabel(f, text='—', text_color=TEXT,
                                         justify='left', anchor='w')
        self.lbl_est_full.pack(fill='x', padx=10, pady=(0, 8))
        ctk.CTkButton(f, text='Tính lại', height=26, fg_color=CARD,
                      hover_color=BORDER, text_color=TEXT,
                      command=self._update_estimate).pack(fill='x', padx=10,
                                                          pady=(0, 8))
        f_region = self._card(tab, 'Vùng & thứ tự xuất')
        ctk.CTkLabel(f_region, text='Thứ tự slide', text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor='w', padx=10)
        self.om_sort = ctk.CTkOptionMenu(
            f_region,
            values=['Tỉnh Bắc → Nam', 'Theo thứ tự list Excel'],
            width=220, height=28, fg_color=CARD,
            button_color=CARD, button_hover_color=BORDER,
            text_color=TEXT, command=self._on_sort_mode)
        self.om_sort.set('Theo thứ tự list Excel' if self.opts.get('sort_mode') == 'list'
                         else 'Tỉnh Bắc → Nam')
        self.om_sort.pack(fill='x', padx=10, pady=(0, 6))
        ctk.CTkLabel(f_region, text='Tách file PPTX theo', text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor='w', padx=10)
        self.om_split_geo = ctk.CTkOptionMenu(
            f_region,
            values=['Không tách', 'Theo tỉnh', 'Theo quận'],
            width=220, height=28, fg_color=CARD,
            button_color=CARD, button_hover_color=BORDER,
            text_color=TEXT, command=self._on_split_geo)
        split_map = {'none': 'Không tách', 'city': 'Theo tỉnh', 'district': 'Theo quận'}
        self.om_split_geo.set(split_map.get(self.opts.get('split_export_by', 'none'),
                                             'Không tách'))
        self.om_split_geo.pack(fill='x', padx=10, pady=(0, 6))
        self.sw_pad_blank = ctk.CTkSwitch(
            f_region, text='Slide trắng / ô trống theo số màn (Excel)',
            command=self._on_pad_blank, progress_color=ACCENT)
        if self.opts.get('pad_blank_slides'):
            self.sw_pad_blank.select()
        self.sw_pad_blank.pack(anchor='w', padx=10, pady=(0, 4))
        ctk.CTkLabel(
            f_region,
            text='VD: 28 màn, 4 ảnh/slide → 7 slide; thiếu ảnh = ô xám trống.',
            text_color=MUTED, font=ctk.CTkFont(size=10),
            wraplength=340, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(0, 8))
        f_stamp_p = self._card(tab, 'Preset đóng dấu')
        self.om_stamp_preset = ctk.CTkOptionMenu(
            f_stamp_p, values=list(STAMP_PRESETS.keys()), width=220, height=28,
            fg_color=CARD, button_color=CARD, button_hover_color=BORDER,
            text_color=TEXT, command=self._on_stamp_preset)
        self.om_stamp_preset.set(self.opts.get('stamp_preset', 'Tùy chỉnh'))
        self.om_stamp_preset.pack(fill='x', padx=10, pady=(0, 4))
        self.sw_no_exif = ctk.CTkSwitch(
            f_stamp_p, text='Cảnh báo ⚠ trên timeline nếu ảnh mất EXIF',
            command=self._on_no_exif_badge,
            progress_color=ACCENT)
        if self.opts.get('show_no_exif_badge', True):
            self.sw_no_exif.select()
        self.sw_no_exif.pack(anchor='w', padx=10, pady=(0, 8))
        self.sw_export_pdf = ctk.CTkSwitch(
            tab, text='Xuất thêm PDF (cần PowerPoint trên máy)',
            command=self._on_export_pdf, progress_color=ACCENT)
        if self.opts.get('export_pdf'):
            self.sw_export_pdf.select()
        self.sw_export_pdf.pack(anchor='w', padx=16, pady=(0, 4))
        f_split = self._card(tab, 'Chia file PPTX')
        row = ctk.CTkFrame(f_split, fg_color='transparent')
        row.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row, text='Mỗi file:', text_color=MUTED).pack(side='left')
        self.om_spf = ctk.CTkOptionMenu(
            row,
            values=G.slides_per_file_menu_values(self.opts.get('slides_per_file', 0)),
            width=140, height=26, fg_color=CARD,
            button_color=CARD, button_hover_color=BORDER,
            text_color=TEXT, command=self._on_spf)
        self.om_spf.set(G.slides_per_file_label(self.opts.get('slides_per_file', 0)))
        self.om_spf.pack(side='left', padx=8)
        row2 = ctk.CTkFrame(f_split, fg_color='transparent')
        row2.pack(fill='x', padx=10, pady=2)
        ctk.CTkLabel(row2, text='Hoặc nhập:', text_color=MUTED).pack(side='left')
        self.ent_spf = ctk.CTkEntry(row2, width=80, height=26, fg_color=CARD)
        n_spf = G.parse_slides_per_file(self.opts.get('slides_per_file', 0))
        self.ent_spf.insert(0, '0' if n_spf <= 0 else str(n_spf))
        self.ent_spf.pack(side='left', padx=8)
        self.ent_spf.bind('<Return>', self._on_spf_custom)
        self.ent_spf.bind('<FocusOut>', self._on_spf_custom)
        ctk.CTkLabel(row2, text='(0 = một file)', text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(side='left')
        ctk.CTkLabel(
            f_split,
            text='Một file = gộp hết slide. Số khác = cắt file khi đủ số slide đó. '
                 'File rất lớn có thể chậm khi mở PowerPoint.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=360, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(0, 8))
        self.sw_open = ctk.CTkSwitch(tab, text='Mở thư mục sau khi xuất',
                                     command=self._on_open_after,
                                     progress_color=ACCENT)
        if self.opts['open_after_export']:
            self.sw_open.select()
        self.sw_open.pack(anchor='w', padx=16, pady=8)
        self.sw_qa_gate = ctk.CTkSwitch(
            tab, text='Cảnh báo đối chiếu list Excel lúc xuất',
            command=self._on_qa_gate, progress_color=ACCENT)
        if self.opts.get('qa_gate', True):
            self.sw_qa_gate.select()
        self.sw_qa_gate.pack(anchor='w', padx=16, pady=(0, 4))
        ctk.CTkLabel(
            tab,
            text='Tắt = xuất hết ảnh theo mã trên tên file, không hỏi điểm thiếu. Trùng Code_RP thì vẫn lấy tên/info từ Excel.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=360, justify='left'
        ).pack(anchor='w', padx=16, pady=(0, 8))
        self.sw_checklist = ctk.CTkSwitch(
            tab, text='Xuất checklist.xlsx (đối chiếu list sales)',
            command=self._on_write_checklist, progress_color=ACCENT)
        if self.opts.get('write_checklist', True):
            self.sw_checklist.select()
        self.sw_checklist.pack(anchor='w', padx=16, pady=(0, 8))
        ctk.CTkLabel(tab, text='File Excel nội bộ cạnh PPTX: điểm list chưa có ảnh, tên file lệch mã.',
                     text_color=MUTED, font=ctk.CTkFont(size=11),
                     wraplength=360, justify='left').pack(anchor='w', padx=16, pady=(0, 8))
        f_stamp = self._card(tab, 'Xuất ảnh đóng dấu')
        ctk.CTkLabel(
            f_stamp,
            text='Chỉ in ngày / vị trí lên ảnh (tab Dấu, GPS) — không tạo slide PPTX. '
                 'Ảnh lưu theo thư mục mã điểm.',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=360, justify='left', anchor='w'
        ).pack(fill='x', padx=10, pady=(0, 8))
        ctk.CTkButton(
            f_stamp, text='Xuất ảnh đóng dấu…', height=40,
            font=ctk.CTkFont(size=14, weight='bold'),
            fg_color='#0d9488', hover_color='#0f766e',
            command=self.export_stamped
        ).pack(fill='x', padx=10, pady=(0, 8))
        ctk.CTkButton(tab, text='Xuất PPTX', height=44,
                      font=ctk.CTkFont(size=15, weight='bold'),
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=self.export).pack(fill='x', padx=8, pady=8)
        self.lbl_export_split = ctk.CTkLabel(
            tab, text=self._spf_hint_text(),
            text_color=MUTED, font=ctk.CTkFont(size=11))
        self.lbl_export_split.pack(padx=8)

    # ════════════════════════ CONTROLLER (EditorCanvas gọi) ════════════════════════
    def _resolve_avatar(self, code, row=None, mcode=None, kind=None):
        """Avatar qua AvatarIndex.find: đúng mã, hậu tố, alias, rồi gần giống.

        Không có file và không tên gần giống → None (không lấy ảnh đầu thư mục).
        """
        if not code:
            return None
        if kind is None:
            row, mcode, kind = match_row(code, self._by_code, self._merged)
        excel = None
        if kind in ('exact', 'merged'):
            excel = mcode or str((row or {}).get('Code_RP') or '')
        return self.avatars.resolve(code, excel)

    def content(self):
        if self._content_cache is not None:
            return self._content_cache
        groups = self.imglib.groups()
        codes = self._ordered_codes()
        code, paths = None, []
        if codes:
            self.group_idx = max(0, min(self.group_idx, len(codes) - 1))
            code = codes[self.group_idx]
            paths = self.imglib.group_paths(code)
        row, mcode, kind = {}, None, 'none'
        if code:
            row, mcode, kind = match_row(code, self._by_code, self._merged)
        elif self._by_code:
            row = next(iter(self._by_code.values()))
        name_val = str(clean(row.get('Name'))).strip() if row else ''
        name_val = name_val or code or 'TÊN ĐỊA ĐIỂM'
        n = self.opts['n_per_slide']
        cap = G.per_slide_max(n)
        K = max(1, math.ceil(len(paths) / cap)) if paths else 1
        avatar = self._resolve_avatar(code, row=row, mcode=mcode, kind=kind)
        code_key = str((row or {}).get('Code_RP') or code or '').upper()
        sibs = self.excel.rows_of(code_key) if code_key else []
        info_table = build_info_table(row, sibs or ([row] if row else []))
        if not info_table.get('school'):
            info_table['school'] = name_val
        self._content_cache = {
            'title': G.title_string(self.layout['title'], name_val, 1, K),
            'info_rows': build_info_rows(row, len(paths) or n),
            'info_table': info_table,
            'slide_style': self.opts.get('slide_style') or 'report',
            'n_screens': screen_qty(row),
            'n_photos': len(paths),
            'channel': channel_text(row, self.opts['channel_template']),
            'images': paths,
            'avatar': avatar,
            'bg': (self.opts['bg_path']
                   if self.opts['bg_mode'] == 'image' and self.opts['bg_path']
                   else None),
            'n': n,
            'ar': AR_CHOICES.get(self.opts['ar_label'], 4 / 3),
            'fit': self.layout['image'].get('fit_mode', 'fill'),
            'gap': self.layout['image'].get('gap', 0.1),
            'radius': self.layout['image'].get('radius', 0),
            'avatar_radius': self.layout['avatar'].get('radius', 0),
            'font': self.layout.get('font', {}),
            'visible': dict(self.opts['visible']),
            'locked': dict(self.opts['locked']),
        }
        vis = self._content_cache['visible']
        vis['channel'] = vis.get('channel', True) and self.opts['channel_enabled']
        return self._content_cache

    def _invalidate(self, render=True):
        self._content_cache = None
        if render and hasattr(self, 'editor'):
            self.editor.schedule_render()
        self._schedule_save()

    def fx_snapshot(self):
        fx = dict(self.opts.get('fx') or {})
        fx['location_excel'] = self._excel_location_label()
        return fx

    def _excel_location_label(self):
        codes = self._ordered_codes()
        if not codes:
            return ''
        code = codes[max(0, min(self.group_idx, len(codes) - 1))]
        row, _, _ = match_row(code, self._by_code, self._merged)
        return place_label(row)

    def _on_loc_mode(self, label):
        mode = (getattr(self, '_loc_mode_vals', {}) or {}).get(label, 'gps')
        cur = str((self.opts.get('fx') or {}).get('location_mode') or 'gps')
        if mode == cur:
            self._sync_loc_hint()
            return
        self._set_fx('location_mode', mode)
        self._sync_loc_hint()
        lab = (self._loc_mode_labels or {}).get(mode, label)
        for w in (getattr(self, 'om_loc_mode', None), getattr(self, 'om_fmt_loc', None)):
            if w is None:
                continue
            try:
                if w.get() != lab:
                    w.set(lab)
            except Exception:
                pass

    def _sync_loc_hint(self):
        lbl = getattr(self, 'lbl_loc_hint', None)
        if lbl is None:
            return
        mode = str((self.opts.get('fx') or {}).get('location_mode') or 'gps')
        hints = {
            'gps': 'Mỗi ảnh: toạ độ EXIF → tên đường (OpenStreetMap). Chưa tải được thì in lat/lon. Ảnh không GPS thì trống.',
            'excel': 'Lấy Name · quận từ Excel theo mã trên tên file ảnh.',
            'auto': 'Ưu tiên tên Excel theo mã. Không khớp → toạ độ EXIF.',
            'manual': 'Một dòng gõ tay cho mọi ảnh.',
        }
        try:
            lbl.configure(text=hints.get(mode, hints['gps']))
        except Exception:
            pass

    def _set_fx(self, key, val):
        fx = self.opts.setdefault('fx', {})
        if fx.get(key) == val:
            return
        fx[key] = val
        if key in ('minimap', 'location_mode'):
            self._mm_done.clear()
        try:
            self.editor.clear_fx_cache()
        except Exception:
            pass
        self._invalidate()
        try:
            self._sync_format_bar()
        except Exception:
            pass

    def _pick_ts_color(self):
        cur = self.opts['fx'].get('text_color', '#FF9900')
        near = getattr(self, 'chip_fmt_stamp', None) or getattr(self, 'lbl_ts_color', None)

        def apply(hx):
            self._set_fx('text_color', hx)
            self._remember_color(hx)
            self._refresh_color_ui()

        self._open_palette(cur, apply, near=near)

    def _norm_hex(self, h, fallback='#000000'):
        s = str(h or '').strip()
        if not s:
            return fallback
        if not s.startswith('#'):
            s = '#' + s
        if len(s) == 4:
            s = '#' + s[1] * 2 + s[2] * 2 + s[3] * 2
        if len(s) != 7:
            return fallback
        try:
            int(s[1:], 16)
        except ValueError:
            return fallback
        return s.upper()

    def _chip_style(self, hexv):
        hx = self._norm_hex(hexv)
        r, g, b = int(hx[1:3], 16), int(hx[3:5], 16), int(hx[5:7], 16)
        luma = 0.299 * r + 0.587 * g + 0.114 * b
        border = '#111111' if luma > 186 else '#f4f4f5'
        return hx, border

    def _make_color_chip(self, parent, hexv, size=34, command=None):
        hx, bd = self._chip_style(hexv)
        return ctk.CTkButton(
            parent, text='', width=size, height=size, corner_radius=8,
            fg_color=hx, hover_color=hx, border_width=2, border_color=bd,
            command=command)

    def _paint_chip(self, btn, hexv):
        if btn is None:
            return
        hx, bd = self._chip_style(hexv)
        try:
            btn.configure(fg_color=hx, hover_color=hx, border_color=bd)
        except Exception:
            pass

    def _role_color_hex(self, key):
        if key == 'title':
            return (self.layout['title'].get('color')
                    or self.layout['font'].get('color') or '#000000')
        if key == 'body':
            return self.layout['font'].get('color') or '#000000'
        if key == 'accent':
            return self.layout['info'].get('accent_color') or '#2563EB'
        if key == 'channel':
            return (self.layout['channel'].get('color')
                    or self.layout['font'].get('color') or '#000000')
        return '#000000'

    def _elem_color_hex(self, name=None):
        name = name or (self.editor.selected if hasattr(self, 'editor') else None) or 'title'
        if name == 'image' or name == 'avatar':
            return self.opts.get('fx', {}).get('text_color', '#FF9900')
        key = 'title' if name == 'title' else ('channel' if name == 'channel' else 'body')
        return self._role_color_hex(key)

    def _remember_color(self, hx):
        hx = self._norm_hex(hx)
        rec = [c for c in (self.opts.get('recent_colors') or [])
               if str(c).upper() != hx]
        rec.insert(0, hx)
        self.opts['recent_colors'] = rec[:12]

    def _refresh_color_ui(self):
        name = 'image'
        try:
            name = self.editor.selected or self._sel_name()
        except Exception:
            pass
        self._paint_chip(getattr(self, 'chip_fmt', None), self._elem_color_hex(name))
        self._paint_chip(getattr(self, 'chip_live', None), self._elem_color_hex(name))
        for k, chip in (getattr(self, '_color_chips', {}) or {}).items():
            self._paint_chip(chip, self._role_color_hex(k))
        stamp = self.opts.get('fx', {}).get('text_color', '#FF9900')
        self._paint_chip(getattr(self, 'lbl_ts_color', None), stamp)
        self._paint_chip(getattr(self, 'chip_fmt_stamp', None), stamp)

    def _sync_format_bar(self):
        if not hasattr(self, 'fmt_bar'):
            return
        name = 'image'
        try:
            name = self.editor.selected or self._sel_name()
        except Exception:
            pass
        try:
            self.lbl_fmt_sel.configure(text=ELEMENT_LABELS.get(name, name))
        except Exception:
            pass
        is_text = name in ('title', 'info', 'channel')
        mode = 'text' if is_text else 'image'
        if getattr(self, '_fmt_mode', None) != mode:
            try:
                self.frm_fmt_text.pack_forget()
                self.frm_fmt_image.pack_forget()
                self.frm_fmt_align.pack_forget()
            except Exception:
                pass
            try:
                if is_text:
                    self.frm_fmt_align.pack(side='right', padx=(8, 10))
                    self.frm_fmt_text.pack(side='left', padx=4)
                else:
                    self.frm_fmt_image.pack(side='left', padx=4)
            except Exception:
                pass
            self._fmt_mode = mode
        if is_text:
            c = self.layout.get(name, {})
            fn = c.get('font') or self.layout.get('font', {}).get('name', 'Arial')
            try:
                self.btn_fmt_font.configure(text=fn)
                self.lbl_fmt_size.configure(text=str(int(c.get('size', 16) or 16)))
                al = c.get('align', 'left')
                self._paint_fmt_align(
                    {'left': 'Trái', 'center': 'Giữa', 'right': 'Phải'}.get(al, 'Trái'))
            except Exception:
                pass
            bold = bool(self.layout.get('font', {}).get('bold', True))
            italic = bool(self.layout.get('font', {}).get('italic', False))
            try:
                self.btn_fmt_bold.configure(
                    fg_color=ACCENT if bold else INPUT,
                    text_color='#ffffff' if bold else TEXT)
                self.btn_fmt_italic.configure(
                    fg_color=ACCENT if italic else INPUT,
                    text_color='#ffffff' if italic else TEXT)
            except Exception:
                pass
        else:
            fx = self.opts.get('fx') or {}
            self._paint_toggle(getattr(self, 'btn_fmt_date', None), fx.get('use_timestamp'))
            self._paint_toggle(getattr(self, 'btn_fmt_loc', None), fx.get('stamp_location'))
            self._paint_toggle(getattr(self, 'btn_fmt_gps', None), fx.get('show_gps'))
            self._paint_toggle(getattr(self, 'btn_fmt_map', None), fx.get('minimap'))
            lab = (self._loc_mode_labels or {}).get(
                str(fx.get('location_mode') or 'gps'), 'Theo toạ độ EXIF')
            om = getattr(self, 'om_fmt_loc', None)
            if om is not None:
                try:
                    if om.get() != lab:
                        om.set(lab)
                except Exception:
                    pass
            for key in ('use_timestamp', 'stamp_location', 'show_gps', 'minimap'):
                w = getattr(self, 'sw_fx_' + key, None)
                if w is None:
                    continue
                on = bool(fx.get(key))
                try:
                    if bool(w.get()) != on:
                        w.select() if on else w.deselect()
                except Exception:
                    pass
        self._refresh_color_ui()

    def _fmt_nudge_size(self, delta):
        name = 'title'
        try:
            name = self.editor.selected or self._sel_name()
        except Exception:
            pass
        if name not in ('title', 'info', 'channel'):
            return
        cur = int(self.layout[name].get('size', 16) or 16)
        self._set_elem_key(name, 'size', max(7, min(48, cur + int(delta))))
        try:
            self.sl_elem_size.set(float(self.layout[name]['size']))
        except Exception:
            pass
        self._sync_format_bar()

    def _fmt_toggle_bold(self):
        try:
            if bool(self.sw_bold.get()):
                self.sw_bold.deselect()
            else:
                self.sw_bold.select()
        except Exception:
            self.layout['font']['bold'] = not bool(self.layout['font'].get('bold', True))
        self._on_style()

    def _fmt_toggle_italic(self):
        try:
            if bool(self.sw_italic.get()):
                self.sw_italic.deselect()
            else:
                self.sw_italic.select()
        except Exception:
            self.layout['font']['italic'] = not bool(self.layout['font'].get('italic', False))
        self._on_style()

    def _open_palette(self, current, on_pick, near=None):
        old = getattr(self, '_pal_win', None)
        if old is not None:
            try:
                old.destroy()
            except Exception:
                pass
        win = ctk.CTkToplevel(self)
        self._pal_win = win
        win.title('Chọn màu')
        win.resizable(False, False)
        try:
            win.attributes('-topmost', True)
            win.transient(self)
        except Exception:
            pass
        try:
            if near is not None:
                near.update_idletasks()
                x = int(near.winfo_rootx())
                y = int(near.winfo_rooty() + near.winfo_height() + 8)
            else:
                x = int(self.winfo_rootx()) + 380
                y = int(self.winfo_rooty()) + 110
            win.geometry(f'300x470+{x}+{y}')
        except Exception:
            win.geometry('300x470')

        cur = self._norm_hex(current)
        head = ctk.CTkFrame(win, fg_color='transparent')
        head.pack(fill='x', padx=12, pady=(12, 4))
        preview = self._make_color_chip(head, cur, 44)
        preview.pack(side='left')
        hex_lbl = ctk.CTkLabel(head, text=cur,
                               font=ctk.CTkFont(size=16, weight='bold'),
                               text_color=TEXT)
        hex_lbl.pack(side='left', padx=12)
        ctk.CTkLabel(win, text='Bấm ô màu là tô ngay — giống Canva.',
                     text_color=MUTED, font=ctk.CTkFont(size=11)).pack(
                         anchor='w', padx=12, pady=(0, 6))

        def pick(hx):
            hx = self._norm_hex(hx)
            try:
                on_pick(hx)
            except Exception:
                pass
            try:
                win.destroy()
            except Exception:
                pass

        recents = self.opts.get('recent_colors') or []
        if recents:
            ctk.CTkLabel(win, text='Vừa dùng', text_color=MUTED,
                         font=ctk.CTkFont(size=11)).pack(anchor='w', padx=12)
            row = ctk.CTkFrame(win, fg_color='transparent')
            row.pack(fill='x', padx=10, pady=(0, 6))
            for hx in recents[:8]:
                self._make_color_chip(row, hx, 30,
                                      command=lambda h=hx: pick(h)).pack(
                                          side='left', padx=3, pady=2)

        ctk.CTkLabel(win, text='Bảng màu', text_color=MUTED,
                     font=ctk.CTkFont(size=11)).pack(anchor='w', padx=12)
        grid = ctk.CTkFrame(win, fg_color='transparent')
        grid.pack(padx=10, pady=4)
        for i, hx in enumerate(COLOR_PALETTE):
            r, c = divmod(i, 6)
            b = self._make_color_chip(grid, hx, 38, command=lambda h=hx: pick(h))
            b.grid(row=r, column=c, padx=3, pady=3)

        def custom():
            col = colorchooser.askcolor(color=cur, parent=self)
            if col and col[1]:
                pick(col[1])

        ctk.CTkButton(win, text='Màu khác…', height=32, fg_color=INPUT,
                      hover_color=BORDER, text_color=TEXT,
                      command=custom).pack(fill='x', padx=12, pady=(8, 12))
        win.bind('<Escape>', lambda e: win.destroy())

    def _pick_color(self, key, near=None):
        targets = {'title': ('title', 'color'), 'body': ('font', 'color'),
                   'accent': ('info', 'accent_color'),
                   'channel': ('channel', 'color')}
        sect, field = targets[key]
        cur = self.layout[sect].get(field) or '#000000'
        if near is None:
            near = (getattr(self, '_color_chips', {}) or {}).get(key)
            near = near or getattr(self, 'chip_fmt', None)

        def apply(hx):
            self.layout[sect][field] = hx
            self._remember_color(hx)
            self._invalidate()
            self._refresh_color_ui()

        self._open_palette(cur, apply, near=near)

    def _choose_logo(self):
        p = filedialog.askopenfilename(
            title='Chọn logo đóng lên ảnh',
            filetypes=[('Ảnh', '*.png;*.jpg;*.jpeg;*.webp')])
        if not p:
            return
        self._set_fx('logo_path', p)
        try:
            self.lbl_logo.configure(text=os.path.basename(p))
        except Exception:
            pass

    def prefetch_minimap(self, paths):
        fx = self.opts.get('fx') or {}
        want_map = bool(fx.get('minimap'))
        want_geo = str(fx.get('location_mode') or '') in ('gps', 'auto')
        if not want_map and not want_geo:
            return
        todo = [p for p in (paths or []) if p and p not in self._mm_done]
        if not todo or self._mm_busy:
            return
        self._mm_busy = True
        self._mm_done.update(todo)

        def work():
            try:
                if want_map:
                    FX.prefetch_maps(todo)
                if want_geo:
                    FX.prefetch_geocode(todo)
            except Exception:
                pass

            def done():
                self._mm_busy = False
                FX.clear_preview_cache()
                if hasattr(self, 'editor'):
                    self.editor.schedule_render()
            try:
                self.after(0, done)
            except Exception:
                self._mm_busy = False

        threading.Thread(target=work, daemon=True).start()

    def _scan_quality_now(self):
        groups = self.imglib.groups()
        paths = [p for files in groups.values() for p in files]
        if not paths:
            messagebox.showinfo('Chất lượng ảnh', 'Chưa chọn thư mục ảnh.')
            return
        self._run_quality_scan(paths, then_export=False)

    def _run_quality_scan(self, paths, then_export=False, export_kwargs=None):
        pop = ctk.CTkToplevel(self)
        pop.title('Đang quét ảnh…')
        pop.geometry('360x120')
        pop.attributes('-topmost', True)
        ctk.CTkLabel(pop, text=f'Đang kiểm tra {len(paths)} ảnh…',
                     font=ctk.CTkFont(size=13, weight='bold')).pack(pady=28)

        def work():
            res = FX.check_quality(paths)
            self.after(0, lambda: self._quality_done(pop, res, then_export, export_kwargs))

        threading.Thread(target=work, daemon=True).start()

    def _quality_done(self, pop, results, then_export, export_kwargs):
        try:
            pop.destroy()
        except Exception:
            pass
        n = (len(results.get('blurry') or []) + len(results.get('dark') or [])
             + len(results.get('duplicates') or []))
        if n == 0:
            self.log('Quét chất lượng: không thấy vấn đề.')
            if then_export:
                self._start_export(**(export_kwargs or {}))
            else:
                messagebox.showinfo('Chất lượng ảnh', 'Không phát hiện ảnh mờ / tối / trùng.')
            return
        win = ctk.CTkToplevel(self)
        win.title('Cảnh báo chất lượng ảnh')
        win.geometry('560x460')
        win.attributes('-topmost', True)
        ctk.CTkLabel(win, text=f'⚠️ Phát hiện {n} vấn đề',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(pady=(16, 8))
        lst = ctk.CTkScrollableFrame(win, fg_color=INPUT, corner_radius=10)
        lst.pack(fill='both', expand=True, padx=16, pady=6)

        def add_group(title, items, fmt=os.path.basename):
            if not items:
                return
            ctk.CTkLabel(lst, text=title, font=ctk.CTkFont(size=12, weight='bold'),
                         text_color=TEXT).pack(anchor='w', padx=8, pady=(8, 2))
            for it in items[:40]:
                txt = f'{fmt(it[0])}  ↔  {fmt(it[1])}' if isinstance(it, tuple) else fmt(it)
                ctk.CTkLabel(lst, text='• ' + txt, text_color=MUTED, anchor='w',
                             font=ctk.CTkFont(size=11)).pack(anchor='w', padx=16)

        add_group(f'Ảnh mờ ({len(results["blurry"])})', results['blurry'])
        add_group(f'Ảnh tối ({len(results["dark"])})', results['dark'])
        add_group(f'Ảnh trùng ({len(results["duplicates"])})', results['duplicates'])

        row = ctk.CTkFrame(win, fg_color='transparent')
        row.pack(fill='x', padx=16, pady=10)
        ctk.CTkButton(row, text='Đóng', width=90, fg_color=CARD, hover_color=BORDER,
                      text_color=TEXT, command=win.destroy).pack(side='right')
        if then_export:
            ctk.CTkButton(row, text='Vẫn xuất', width=110, fg_color=ACCENT,
                          hover_color=ACCENT_HOVER,
                          command=lambda: (win.destroy(),
                                           self._start_export(**(export_kwargs or {})))
                          ).pack(side='right', padx=8)

    def on_layout_change(self):
        self._sync_inspector()
        self._schedule_save()

    def on_drag_live(self, name):
        self._sync_inspector()

    def on_select(self, name):
        if name:
            self.om_elem.set(ELEMENT_LABELS[name])
        self._sync_inspector()

    def on_zoom(self, z):
        self.lbl_zoom.configure(text=f'{z}%')

    def quick_color(self, name):
        self.open_style_panel(name)

    def open_style_panel(self, name):
        """Bấm đúp → mở cột Khung, đúng phần tử."""
        if name:
            self.editor.select(name)
            try:
                self.om_elem.set(ELEMENT_LABELS.get(name, name))
            except Exception:
                pass
            self._sync_inspector()
            if not getattr(self, '_insp_visible', False):
                self._toggle_inspector()

    def _sync_live_style(self, name):
        if not hasattr(self, 'lbl_style_for'):
            return
        c = self.layout.get(name, {})
        font_cfg = self.layout.get('font', {})
        self.lbl_style_for.configure(text=ELEMENT_LABELS.get(name, name))
        is_text = name in ('title', 'info', 'channel')
        fn = c.get('font') or font_cfg.get('name', 'Arial')
        try:
            if hasattr(self, 'btn_font'):
                self.btn_font.configure(text=fn)
            info = FN.family_info(fn)
            warn = ''
            if not FN.ready():
                warn = ''
            elif info is not None and not info.get('viet'):
                warn = 'Font này thiếu dấu tiếng Việt — chữ ÁÀẢÃẠ có thể thành ô vuông.'
            elif info is None:
                warn = 'Chưa thấy file .ttf của font này trên máy.'
            self.lbl_font_warn.configure(text=warn)
        except Exception:
            pass
        try:
            self.sl_elem_size.set(float(c.get('size', 16) or 16) if is_text else 16)
            self.sl_elem_opacity.set(float(c.get('opacity', 100) or 100))
            self.sl_elem_track.set(float(c.get('tracking', 0) or 0) if is_text else 0)
            al = c.get('align', 'left')
            self.seg_elem_align.set({'left': 'Trái', 'center': 'Giữa', 'right': 'Phải'}.get(al, 'Trái'))
        except Exception:
            pass

    def _on_elem_size(self, v):
        if self._insp_updating:
            return
        name = self.editor.selected or self._sel_name()
        if name in ('title', 'info', 'channel'):
            self._set_elem_key(name, 'size', int(float(v)))
            try:
                self.lbl_fmt_size.configure(text=str(int(float(v))))
            except Exception:
                pass

    def _on_elem_opacity(self, v):
        if self._insp_updating:
            return
        self._set_elem_key(self.editor.selected or self._sel_name(), 'opacity', int(float(v)))

    def _on_elem_track(self, v):
        if self._insp_updating:
            return
        name = self.editor.selected or self._sel_name()
        if name in ('title', 'channel'):
            self._set_elem_key(name, 'tracking', round(float(v), 1))

    def _on_elem_color(self):
        name = self.editor.selected or self._sel_name()
        key = 'title' if name == 'title' else (
            'channel' if name == 'channel' else 'body')
        near = getattr(self, 'chip_fmt', None) or getattr(self, 'chip_live', None)
        self._pick_color(key, near=near)

    def _on_elem_align(self, v):
        if getattr(self, '_insp_updating', False):
            return
        name = self.editor.selected or self._sel_name()
        if name not in ('title', 'info', 'channel'):
            return
        self._set_elem_key(name, 'align', {'Trái': 'left', 'Giữa': 'center', 'Phải': 'right'}[v])
        self._paint_fmt_align(v)

    def _paint_fmt_align(self, v):
        for lab, btn in (getattr(self, '_fmt_align_btns', None) or {}).items():
            on = lab == v
            try:
                btn.configure(
                    fg_color=ACCENT if on else INPUT,
                    hover_color=ACCENT_HOVER if on else BORDER,
                    text_color='#ffffff' if on else TEXT)
            except Exception:
                pass

    def _set_elem_font(self, name, font_name, path=None, index=0):
        info = FN.family_info(font_name)
        path = path or (info or {}).get('path') or ''
        index = int(index or (info or {}).get('index', 0) or 0)
        self.layout['font']['name'] = font_name
        self.layout['font']['path'] = path
        self.layout['font']['index'] = index
        if name in ('title', 'info', 'channel'):
            self.layout[name]['font'] = font_name
            self.layout[name]['font_path'] = path
            self.layout[name]['font_index'] = index
        self._invalidate()
        try:
            self.btn_font.configure(text=font_name)
        except Exception:
            pass
        try:
            self.btn_fmt_font.configure(text=font_name)
        except Exception:
            pass
        try:
            self._sync_format_bar()
        except Exception:
            pass

    def _set_elem_key(self, name, key, value):
        if name not in self.layout:
            return
        self.layout[name][key] = value
        self._invalidate()

    def _color_for(self, name):
        if name == 'title':
            self._pick_color('title')
        elif name == 'channel':
            self._pick_color('channel')
        elif name == 'info':
            self._pick_color('body')
        else:
            self._pick_color('body')

    # ════════════════════════ INSPECTOR ════════════════════════
    def _sel_name(self):
        label = self.om_elem.get()
        return next((n for n, l in ELEMENT_LABELS.items() if l == label),
                    'image')

    def _on_pick_elem(self, _=None):
        name = self._sel_name()
        self.editor.select(name)
        self._sync_inspector()

    def _sync_inspector(self):
        if self._insp_updating:
            return
        self._insp_updating = True
        try:
            name = self.editor.selected or self._sel_name()
            x, y, w, h = G.elem_box(self.layout, name)
            c = self.layout[name]
            vals = {'x': c['x'], 'y': c['y'], 'w': c['w'],
                    'h': c.get('h', h)}
            for k, e in self._insp.items():
                cur = e.get()
                new = f'{vals[k]:.2f}'
                if cur != new and e.focus_get() is not e:
                    e.delete(0, 'end')
                    e.insert(0, new)
            self.sw_lock.select() if self.opts['locked'].get(name) else self.sw_lock.deselect()
            self.sw_hide.select() if not self.opts['visible'].get(name, True) else self.sw_hide.deselect()
            dragging = bool(getattr(self.editor, '_drag', None))
            if not dragging:
                self._sync_live_style(name)
                self._sync_ctx_panels(name)
                self._sync_format_bar()
        except Exception:
            pass
        finally:
            self._insp_updating = False

    def _sync_ctx_panels(self, name):
        """Hiện inspector đúng loại phần tử ở cột phải. GPS ở tab GPS; ngày giờ ở tab Dấu."""
        name = name or 'image'
        mapping = {
            'title': ('frm_type_text', 'frm_title_extra'),
            'info': ('frm_type_text',),
            'channel': ('frm_type_text', 'frm_channel_extra'),
            'image': ('frm_type_image',),
            'avatar': ('frm_type_avatar',),
        }
        want = mapping.get(name, ('frm_type_image',))
        for key in ('frm_type_text', 'frm_title_extra', 'frm_channel_extra',
                    'frm_type_image', 'frm_type_avatar'):
            frm = getattr(self, key, None)
            if frm is None:
                continue
            try:
                frm.pack_forget()
            except Exception:
                pass
        for key in want:
            frm = getattr(self, key, None)
            if frm is None:
                continue
            try:
                frm.pack(fill='x', pady=(0, 4))
            except Exception:
                pass

    def _apply_inspector(self, _=None):
        if self._insp_updating:
            return
        name = self.editor.selected or self._sel_name()
        c = self.layout[name]
        try:
            self.editor.push_undo()
            for k, e in self._insp.items():
                v = float(e.get().replace(',', '.'))
                if k == 'h' and 'h' not in c:
                    continue
                c[k] = max(0.0, v)
            self.editor.render()
            self._schedule_save()
        except ValueError:
            pass

    def _toggle_lock(self):
        name = self.editor.selected or self._sel_name()
        self.opts['locked'][name] = bool(self.sw_lock.get())
        self._invalidate()

    def _toggle_hide(self):
        name = self.editor.selected or self._sel_name()
        on = not bool(self.sw_hide.get())
        self.opts['visible'][name] = on
        if name == 'channel':
            self.opts['channel_enabled'] = on
            self._sync_channel_switch()
        self._invalidate()

    # ════════════════════════ TUỲ CHỌN ════════════════════════
    def _toggle_snap(self):
        self.editor.snap_on = bool(self.sw_snap.get())

    def _toggle_grid(self):
        self.editor.set_grid(bool(self.sw_grid.get()))

    def _toggle_appearance(self):
        mode = 'Light' if ctk.get_appearance_mode() == 'Dark' else 'Dark'
        ctk.set_appearance_mode(mode)
        try:
            self.tl.configure(bg=self._tl_chrome()['bg'])
            self._rebuild_timeline()
        except Exception:
            pass
        self._schedule_save()

    def _on_n(self, v):
        self.opts['n_per_slide'] = G.parse_n_per_slide(v)
        self._invalidate()
        self._update_estimate()

    def _spf_hint_text(self):
        n = G.parse_slides_per_file(self.opts.get('slides_per_file', 0))
        if n <= 0:
            return 'Gộp mọi slide vào một file. Có thể huỷ giữa chừng.'
        return f'Tự chia file mỗi {n} slide. Có thể huỷ giữa chừng.'

    def _refresh_spf_ui(self):
        n = G.parse_slides_per_file(self.opts.get('slides_per_file', 0))
        label = G.slides_per_file_label(n)
        values = G.slides_per_file_menu_values(n)
        for attr in ('om_spf', 'om_spf_step'):
            w = getattr(self, attr, None)
            if w is None:
                continue
            try:
                if not w.winfo_exists():
                    continue
                w.configure(values=values)
                if w.get() != label:
                    w.set(label)
            except Exception:
                pass
        ent = getattr(self, 'ent_spf', None)
        if ent is not None:
            try:
                if ent.winfo_exists():
                    want = '0' if n <= 0 else str(n)
                    if ent.get().strip() != want:
                        ent.delete(0, 'end')
                        ent.insert(0, want)
            except Exception:
                pass
        hint = getattr(self, 'lbl_export_split', None)
        if hint is not None:
            try:
                if hint.winfo_exists():
                    hint.configure(text=self._spf_hint_text())
            except Exception:
                pass

    def _on_spf(self, v):
        n = G.parse_slides_per_file(v)
        if G.parse_slides_per_file(self.opts.get('slides_per_file', 0)) == n:
            return
        self.opts['slides_per_file'] = n
        self._schedule_save()
        self._update_estimate()
        self._refresh_spf_ui()

    def _on_spf_custom(self, _e=None):
        ent = getattr(self, 'ent_spf', None)
        if ent is None:
            return
        try:
            raw = ent.get().strip()
        except Exception:
            return
        if raw == '':
            return
        n = G.parse_slides_per_file(
            raw, default=G.parse_slides_per_file(self.opts.get('slides_per_file', 0)))
        self.opts['slides_per_file'] = n
        self._schedule_save()
        self._update_estimate()
        self._refresh_spf_ui()

    def _on_ar(self, v):
        self.opts['ar_label'] = v
        self._invalidate()

    def _on_fit(self, v):
        self.layout['image']['fit_mode'] = 'fill' if v == 'Lấp đầy' else 'fit'
        try:
            self.editor.clear_fx_cache()
        except Exception:
            pass
        self._invalidate()

    def _on_gap(self, v):
        self.layout['image']['gap'] = round(float(v), 2)
        self._invalidate()

    def _on_radius(self, v):
        self.layout['image']['radius'] = int(float(v))
        try:
            self.editor.clear_fx_cache()
        except Exception:
            pass
        self._invalidate()

    def _on_avatar_ar(self, v):
        self.layout['avatar']['ar'] = {'4:3': 4 / 3, '1:1': 1.0,
                                       '3:4': 3 / 4}[v]
        self._invalidate()

    def _on_av_radius(self, v):
        self.layout['avatar']['radius'] = int(float(v))
        self._invalidate()

    def _on_font(self, v):
        self._set_elem_font(self.editor.selected or self._sel_name(), v)

    def _open_font_dialog(self):
        old = getattr(self, '_font_win', None)
        if old is not None:
            try:
                if old.winfo_exists():
                    old.lift()
                    old.focus_force()
                    return
            except Exception:
                pass
        win = ctk.CTkToplevel(self)
        self._font_win = win
        win.title('Chọn font')
        win.geometry('440x620')
        win.attributes('-topmost', True)
        try:
            win.transient(self)
        except Exception:
            pass
        ctk.CTkLabel(win, text='Font trên máy — mẫu tiếng Việt ÁÀẢÃẠ',
                     font=ctk.CTkFont(size=13, weight='bold')).pack(
                         anchor='w', padx=14, pady=(12, 4))
        ent = ctk.CTkEntry(win, height=30, placeholder_text='Gõ tên font…')
        ent.pack(fill='x', padx=14, pady=4)
        preview = ctk.CTkLabel(win, text='', width=392, height=54)
        preview.pack(fill='x', padx=14, pady=(4, 2))
        warn = ctk.CTkLabel(win, text='Đang tải danh sách font…',
                            text_color=MUTED, font=ctk.CTkFont(size=11),
                            anchor='w', wraplength=400, justify='left')
        warn.pack(fill='x', padx=14)
        lst = ctk.CTkScrollableFrame(win, fg_color=INPUT)
        lst.pack(fill='both', expand=True, padx=14, pady=8)
        ctk.CTkButton(win, text='Đóng', height=32, fg_color=CARD,
                      hover_color=BORDER, text_color=TEXT,
                      command=win.destroy).pack(fill='x', padx=14, pady=(0, 12))
        state = {'btns': [], 'sel': None, 'job': None}

        def show_sample(row):
            im = FN.preview_sample(row.get('path'), row.get('index', 0))
            img = ctk.CTkImage(light_image=im, dark_image=im, size=im.size)
            win._sample_img = img
            preview.configure(image=img, text='')
            if not row.get('viet'):
                warn.configure(
                    text='Font này thiếu dấu tiếng Việt — ÁÀẢÃẠ có thể thành ô vuông.')
            else:
                p = row.get('path') or ''
                warn.configure(
                    text='Preview và PPTX trỏ cùng file: ' + os.path.basename(p))

        def apply_row(row):
            name = self.editor.selected or self._sel_name()
            self._set_elem_font(name, row['family'], row.get('path'),
                                row.get('index', 0))
            show_sample(row)
            state['sel'] = row['family']
            for b in state['btns']:
                fam = getattr(b, '_family', '')
                try:
                    b.configure(fg_color=ACCENT if fam == row['family'] else CARD)
                except Exception:
                    pass

        def paint(query=''):
            for b in state['btns']:
                try:
                    b.destroy()
                except Exception:
                    pass
            state['btns'] = []
            rows = FN.search_families(query, limit=80)
            cur = (self.layout.get('font') or {}).get('name', '')
            for row in rows:
                tag = '' if row.get('viet') else '  · thiếu dấu Việt'
                b = ctk.CTkButton(
                    lst, text=row['family'] + tag, height=28, anchor='w',
                    fg_color=ACCENT if row['family'] == cur else CARD,
                    hover_color=BORDER, text_color=TEXT,
                    command=lambda r=row: apply_row(r))
                b._family = row['family']
                b.pack(fill='x', pady=1)
                state['btns'].append(b)
            if not query:
                warn.configure(text=f'{len(FN.catalog())} font trên máy. Bấm để áp dụng.')
            else:
                warn.configure(text=f'{len(rows)} font khớp «{query}»')
            if rows and not getattr(win, '_sample_img', None):
                show_sample(rows[0])

        def on_key(_=None):
            if state['job']:
                try:
                    win.after_cancel(state['job'])
                except Exception:
                    pass
            state['job'] = win.after(120, lambda: paint(ent.get()))

        def loaded():
            if not win.winfo_exists():
                return
            paint()
            ent.bind('<KeyRelease>', on_key)
            try:
                ent.focus_set()
            except Exception:
                pass

        def work():
            try:
                FN.catalog()
            except Exception:
                pass
            try:
                self.after(0, loaded)
            except Exception:
                pass
        threading.Thread(target=work, daemon=True).start()

    def _on_style(self):
        self.layout['font']['bold'] = bool(self.sw_bold.get())
        self.layout['font']['italic'] = bool(self.sw_italic.get())
        self._invalidate()
        self._sync_format_bar()

    def _set_size(self, name, v):
        self.layout[name]['size'] = int(float(v))
        self._invalidate()

    def _on_title_align(self, v):
        self.layout['title']['align'] = {'Trái': 'left', 'Giữa': 'center',
                                         'Phải': 'right'}[v]
        self._invalidate()

    def _on_title_flags(self):
        t = self.layout['title']
        t['upper'] = bool(self.sw_upper.get())
        t['show_counter'] = bool(self.sw_counter.get())
        t['full_width_on_slide'] = bool(self.sw_fullw.get())
        self._invalidate()

    def _sync_channel_switch(self):
        sw = getattr(self, 'sw_channel', None)
        if sw is None:
            return
        on = bool(self.opts.get('channel_enabled', True))
        try:
            sw.select() if on else sw.deselect()
        except Exception:
            pass

    def _on_channel_toggle(self):
        on = bool(self.sw_channel.get())
        self.opts['channel_enabled'] = on
        self.opts.setdefault('visible', {})['channel'] = on
        self._invalidate()

    def _on_channel_tpl(self, _=None):
        self.opts['channel_template'] = self.ent_ch_tpl.get()
        self._invalidate()

    def _on_open_after(self):
        try:
            self.opts['open_after_export'] = bool(self.sw_open.get())
        except Exception:
            pass
        self._schedule_save()

    def _on_write_checklist(self):
        try:
            self.opts['write_checklist'] = bool(self.sw_checklist.get())
        except Exception:
            pass
        self._schedule_save()

    def _on_qa_gate(self):
        try:
            self.opts['qa_gate'] = bool(self.sw_qa_gate.get())
        except Exception:
            pass
        self._schedule_save()

    # ════════════════════════ NGUỒN DỮ LIỆU ════════════════════════
    def _restore_sources(self):
        o = self.opts
        if o.get('app_mode') == 'cloud':
            cache = avatar_cache_dir()
            o['avatar_folder'] = cache
            self.avatars.set_folder(cache)
            if hasattr(self, 'lbl_avatars'):
                self.lbl_avatars.configure(text='☁ Cache Cloud: ' + cache)
        elif o['avatar_folder'] and os.path.isdir(o['avatar_folder']):
            self.avatars.set_folder(o['avatar_folder'])
            self.lbl_avatars.configure(text=o['avatar_folder'])
        if o.get('image_folders') or o.get('image_folder'):
            self._apply_image_folders()
        if o['bg_path'] and os.path.isfile(o['bg_path']):
            self.lbl_bg.configure(text=o['bg_path'])
        excel = o.get('excel_path') or ''
        if excel and os.path.isfile(excel):
            # Cloud: đợi sync (show_main) — tránh load cache cũ rồi race với bản mới.
            if not (o.get('app_mode') == 'cloud' and self.cloud.logged_in()):
                self._load_excel_async(excel)
        self._apply_mode_ui()

    def _first_show(self):
        self.editor.zoom_fit()
        self._rebuild_timeline()
        self._update_coverage_label()
        self._update_estimate()
        self._sync_format_bar()

    def select_excel(self):
        self.pick_master_excel(purpose='local')

    def pick_master_excel(self, purpose='local'):
        """Chọn FILE TỔNG: local = dùng máy; push = đẩy Cloud (trong Cài đặt)."""
        if purpose == 'push' and not self.cloud.logged_in():
            messagebox.showwarning('Cloud', 'Cần đăng nhập để đẩy FILE TỔNG lên Cloud.')
            return
        title = ('Chọn FILE TỔNG / MASTER để đẩy lên Cloud'
                 if purpose == 'push'
                 else 'Chọn FILE TỔNG / MASTER chuẩn (cột Code_RP)')
        p = filedialog.askopenfilename(
            title=title,
            filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not p:
            return
        pop = ctk.CTkToplevel(self)
        pop.title('Đang kiểm tra list…')
        pop.geometry('320x100')
        pop.attributes('-topmost', True)
        ctk.CTkLabel(pop, text='Đang đọc file Excel…',
                     font=ctk.CTkFont(size=13, weight='bold')).pack(pady=28)

        def work():
            info = inspect_excel(p)
            self.after(0, lambda: self._list_inspected(pop, p, info, purpose=purpose))

        threading.Thread(target=work, daemon=True).start()

    def up_sales_list(self):
        """Tương thích cũ — chọn file trên máy."""
        self.pick_master_excel(purpose='local')

    def _list_inspected(self, pop, path, info, purpose='local'):
        try:
            pop.destroy()
        except Exception:
            pass
        if not info.get('ok'):
            messagebox.showerror(
                'FILE TỔNG / MASTER',
                'Không nhận được file tổng chuẩn.\n\n'
                f"{info.get('error') or ''}\n\n"
                'Cần cột Code_RP (hoặc Code / Mã báo cáo) ở sheet đầu.')
            return
        win = ctk.CTkToplevel(self)
        win.title('Đẩy FILE TỔNG lên Cloud' if purpose == 'push'
                  else 'FILE TỔNG / MASTER')
        win.geometry('520x460')
        win.attributes('-topmost', True)
        ctk.CTkLabel(
            win, text=os.path.basename(path),
            font=ctk.CTkFont(size=15, weight='bold')
        ).pack(fill='x', padx=16, pady=(16, 4))
        n = int(info.get('n') or 0)
        chans = info.get('channels') or []
        found = info.get('found') or []
        missing = info.get('missing') or []
        sample = info.get('sample_codes') or []
        ctk.CTkLabel(
            win,
            text=f'{n} điểm  ·  {len(chans)} kênh  ·  cột: {", ".join(found)}',
            text_color=MUTED, wraplength=480, justify='left', anchor='w'
        ).pack(fill='x', padx=16)
        if missing:
            ctk.CTkLabel(
                win,
                text='Thiếu cột (vẫn dùng được): ' + ', '.join(missing),
                text_color='#ca8a04', wraplength=480, justify='left', anchor='w'
            ).pack(fill='x', padx=16, pady=(4, 0))
        if sample:
            ctk.CTkLabel(
                win, text='Mã mẫu: ' + ', '.join(sample),
                text_color=MUTED, wraplength=480, justify='left', anchor='w'
            ).pack(fill='x', padx=16, pady=(4, 8))
        tb = ctk.CTkTextbox(win, height=90, font=ctk.CTkFont(size=12))
        tb.pack(fill='x', padx=16, pady=(0, 8))
        tb.insert('1.0', (
            'Đẩy lên Cloud để máy khác đồng bộ cùng FILE TỔNG / MASTER.'
            if purpose == 'push' else
            'FILE TỔNG / MASTER dùng để tạo slide và gắn ảnh theo Code_RP trên máy này.'))
        tb.configure(state='disabled')

        def use_local():
            try:
                win.destroy()
            except Exception:
                pass
            self._load_excel_async(path)

        def push_me():
            if not self.cloud.logged_in():
                messagebox.showwarning('Cloud', 'Cần đăng nhập để đẩy list lên Cloud.')
                return
            try:
                win.destroy()
            except Exception:
                pass
            self._do_upload_excel(path, None, None, then_load=True)

        def push_all():
            if not self.cloud.is_admin():
                messagebox.showwarning('Cloud', 'Chỉ quản trị mới đẩy list cho mọi tài khoản.')
                return
            if not messagebox.askyesno(
                    'FILE TỔNG / MASTER',
                    f'Đẩy FILE TỔNG «{os.path.basename(path)}» ({n} điểm) cho TẤT CẢ tài khoản?'):
                return
            try:
                win.destroy()
            except Exception:
                pass
            self._do_upload_excel(path, 'all', None, then_load=True)

        bar = ctk.CTkFrame(win, fg_color='transparent')
        bar.pack(fill='x', padx=16, pady=(0, 16))
        if purpose == 'local':
            ctk.CTkButton(bar, text='Dùng trên máy này', height=34, fg_color=ACCENT,
                          hover_color=ACCENT_HOVER, command=use_local
                          ).pack(fill='x', pady=3)
        if purpose == 'push' and self.cloud.logged_in():
            ctk.CTkButton(bar, text='Đẩy Cloud (tài khoản tôi)',
                          height=34, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                          command=push_me).pack(fill='x', pady=3)
            if self.cloud.is_admin():
                ctk.CTkButton(bar, text='Đẩy cho TẤT CẢ tài khoản',
                              height=32, fg_color='#b45309', hover_color='#92400e',
                              command=push_all).pack(fill='x', pady=3)
        elif purpose == 'push':
            ctk.CTkLabel(bar, text='Cần đăng nhập để đẩy Cloud.',
                         text_color=MUTED).pack(pady=4)

    def _load_excel_async(self, path):
        if hasattr(self, 'lbl_excel'):
            self.lbl_excel.configure(text=f'Đang đọc… {os.path.basename(path)}')
        self.log(f'Đang đọc Excel: {os.path.basename(path)}')

        def work():
            ok = self.excel.load(path)
            self.after(0, lambda: self._excel_loaded(path, ok))
        threading.Thread(target=work, daemon=True).start()

    def _excel_loaded(self, path, ok):
        if not ok:
            try:
                self.lbl_excel.configure(text=f'Lỗi: {self.excel.error}')
            except Exception:
                pass
            messagebox.showerror('Excel', f'Không đọc được file:\n{self.excel.error}')
            return
        self.opts['excel_path'] = path
        chans = [ALL_CHANNELS] + self.excel.channels()
        self.om_channel.configure(values=chans)
        if self.opts['channel_filter'] not in chans:
            self.opts['channel_filter'] = ALL_CHANNELS
        self.om_channel.set(self.opts['channel_filter'])
        self._refresh_geo_filters()
        self._rebuild_code_map()
        display = self.cloud.excel_local_display_name(path)
        cap = (f'{display} — {len(self.excel.rows)} dòng, '
               f'{len(chans) - 1} kênh')
        try:
            self.lbl_excel.configure(text=cap)
        except Exception:
            pass
        self._update_data_brief()
        self.log(f'Đã đọc {len(self.excel.rows)} dòng Excel.')
        try:
            self.lbl_list_master_compare.configure(
                text=self._list_master_compare_caption())
        except Exception:
            pass
        self._invalidate()
        self._update_coverage_label()
        self._update_estimate()

    def _rebuild_code_map(self):
        ch = self.opts['channel_filter']
        self._by_code = self.excel.by_code(None if ch == ALL_CHANNELS else ch)
        self._merged = build_merged_groups(self._by_code)

    def _on_channel(self, v):
        self.opts['channel_filter'] = v
        self._rebuild_code_map()
        self._refresh_geo_filters()
        self._rebuild_timeline()
        self._update_estimate()
        self._update_data_brief()
        self._schedule_save()

    def _refresh_geo_filters(self):
        rows = self.excel.rows if getattr(self.excel, 'rows', None) else []
        ch = self.opts.get('channel_filter', ALL_CHANNELS)
        if ch != ALL_CHANNELS:
            rows = [r for r in rows if str(clean(r.get('Channel'))).strip() == ch]
        cities = [ALL_CITIES] + geo_field_values(rows, 'City')
        cur_city = self.opts.get('city_filter', ALL_CITIES)
        if cur_city not in cities:
            cur_city = ALL_CITIES
            self.opts['city_filter'] = cur_city
        try:
            self.pick_city.configure(values=cities)
            self.pick_city.set(cur_city)
        except Exception:
            pass
        dist_rows = rows
        if cur_city != ALL_CITIES:
            from .datasource import norm_city
            ck = norm_city(cur_city)
            dist_rows = [r for r in rows if norm_city(r.get('City')) == ck]
        dists = [ALL_DISTRICTS] + geo_field_values(dist_rows, 'District')
        cur_dist = self.opts.get('district_filter', ALL_DISTRICTS)
        if cur_dist not in dists:
            cur_dist = ALL_DISTRICTS
            self.opts['district_filter'] = cur_dist
        try:
            self.pick_district.configure(values=dists)
            self.pick_district.set(cur_dist)
        except Exception:
            pass

    def _on_city_filter(self, v):
        self.opts['city_filter'] = v
        if v != ALL_CITIES:
            self.opts['district_filter'] = ALL_DISTRICTS
        self._refresh_geo_filters()
        self._rebuild_timeline()
        self._update_estimate()
        self._update_data_brief()
        self._schedule_save()

    def _on_district_filter(self, v):
        self.opts['district_filter'] = v
        self._rebuild_timeline()
        self._update_estimate()
        self._update_data_brief()
        self._schedule_save()

    def _on_sort_mode(self, v):
        self.opts['sort_mode'] = 'list' if 'list' in (v or '').lower() else 'city'
        self._rebuild_timeline()
        self._update_estimate()
        self._schedule_save()

    def _on_split_geo(self, v):
        m = {'Không tách': 'none', 'Theo tỉnh': 'city', 'Theo quận': 'district'}
        self.opts['split_export_by'] = m.get(v, 'none')
        self._update_estimate()
        self._schedule_save()

    def _on_pad_blank(self):
        self.opts['pad_blank_slides'] = bool(self.sw_pad_blank.get())
        self._update_estimate()
        self._schedule_save()

    def _on_export_pdf(self):
        self.opts['export_pdf'] = bool(self.sw_export_pdf.get())
        self._schedule_save()

    def _on_no_exif_badge(self):
        self.opts['show_no_exif_badge'] = bool(self.sw_no_exif.get())
        self._rebuild_timeline()
        self._schedule_save()

    def _on_stamp_preset(self, v):
        self.opts['stamp_preset'] = v
        preset = STAMP_PRESETS.get(v) or {}
        if preset:
            self.opts.setdefault('fx', {}).update(preset)
            try:
                self.editor.clear_fx_cache()
            except Exception:
                pass
            self._invalidate()
            self._sync_format_bar()
        self._schedule_save()

    def _filtered_groups(self, groups=None):
        groups = OrderedDict(groups if groups is not None else self.imglib.groups())
        return filter_groups_by_geo(
            groups, self._by_code, self._merged,
            city=self.opts.get('city_filter', ALL_CITIES),
            district=self.opts.get('district_filter', ALL_DISTRICTS),
            all_cities=ALL_CITIES, all_districts=ALL_DISTRICTS)

    def _ordered_groups(self, groups=None):
        groups = self._filtered_groups(groups)
        return order_groups(
            groups, self._by_code, self._merged, self.excel.rows,
            self.opts.get('sort_mode', 'city'))

    def select_images(self):
        self.add_image_folder()

    def add_image_folder(self):
        p = filedialog.askdirectory(title='Thêm thư mục ảnh (có thể chọn nhiều thư mục con)')
        if not p:
            return
        children = []
        try:
            children = sorted(
                os.path.join(p, n) for n in os.listdir(p)
                if os.path.isdir(os.path.join(p, n)))
        except Exception:
            children = []
        if children:
            chosen = self._ask_which_folders(p, children)
            if not chosen:
                return
        else:
            chosen = [p]
        folders = list(self.opts.get('image_folders') or [])
        existing = {os.path.normcase(os.path.normpath(f)) for f in folders}
        for c in chosen:
            key = os.path.normcase(os.path.normpath(c))
            if key not in existing:
                folders.append(c)
                existing.add(key)
        self.opts['image_folders'] = folders
        self.opts['image_folder'] = folders[0] if folders else ''
        self._apply_image_folders()
        self.group_idx = 0
        self._invalidate()
        self._rebuild_timeline()
        self._update_coverage_label()
        self._update_estimate()

    def _ask_which_folders(self, parent_path, children):
        """Chọn nhiều thư mục con cùng lúc (Windows askdirectory chỉ 1 folder)."""
        result = {'paths': None}
        win = ctk.CTkToplevel(self)
        win.title('Chọn thư mục ảnh')
        win.geometry('440x520')
        win.attributes('-topmost', True)
        try:
            win.transient(self)
            win.grab_set()
        except Exception:
            pass
        ctk.CTkLabel(
            win, text='Chọn các thư mục con để quét ảnh',
            font=ctk.CTkFont(size=13, weight='bold')
        ).pack(anchor='w', padx=14, pady=(12, 4))
        ctk.CTkLabel(
            win, text=parent_path, text_color=MUTED,
            font=ctk.CTkFont(size=11), wraplength=400, justify='left'
        ).pack(anchor='w', padx=14)

        parent_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            win, text='Cả thư mục này (mọi cấp con)',
            variable=parent_var, text_color=TEXT
        ).pack(anchor='w', padx=14, pady=(8, 4))

        sc = ctk.CTkScrollableFrame(win, fg_color=INPUT)
        sc.pack(fill='both', expand=True, padx=14, pady=6)
        child_vars = []
        for ch in children:
            v = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(
                sc, text=os.path.basename(ch) or ch,
                variable=v, text_color=TEXT
            ).pack(anchor='w', padx=6, pady=2)
            child_vars.append((ch, v))

        def all_on():
            for _p, v in child_vars:
                v.set(True)

        def all_off():
            parent_var.set(False)
            for _p, v in child_vars:
                v.set(False)

        def ok():
            paths = []
            if parent_var.get():
                paths.append(parent_path)
            for pth, v in child_vars:
                if v.get():
                    paths.append(pth)
            result['paths'] = paths
            win.destroy()

        def cancel():
            result['paths'] = None
            win.destroy()

        row = ctk.CTkFrame(win, fg_color='transparent')
        row.pack(fill='x', padx=14, pady=(0, 12))
        ctk.CTkButton(row, text='Tất cả con', width=90, height=30,
                      fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                      command=all_on).pack(side='left')
        ctk.CTkButton(row, text='Bỏ chọn', width=80, height=30,
                      fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                      command=all_off).pack(side='left', padx=6)
        ctk.CTkButton(row, text='Huỷ', width=70, height=30,
                      fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                      command=cancel).pack(side='right')
        ctk.CTkButton(row, text='Thêm', width=80, height=30,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=ok).pack(side='right', padx=(0, 6))
        win.protocol('WM_DELETE_WINDOW', cancel)
        self.wait_window(win)
        return result['paths']

    def clear_image_folders(self):
        self.opts['image_folders'] = []
        self.opts['image_folder'] = ''
        self._apply_image_folders()
        self.group_idx = 0
        self._invalidate()
        self._rebuild_timeline()
        self._update_coverage_label()
        self._update_estimate()

    def _on_scan_subfolders(self):
        self.opts['scan_subfolders'] = bool(self.sw_subfolders.get())
        self._apply_image_folders()
        self._invalidate()
        self._rebuild_timeline()
        self._update_coverage_label()
        self._update_estimate()

    def _apply_image_folders(self):
        folders = [f for f in (self.opts.get('image_folders') or [])
                   if f and os.path.isdir(f)]
        self.opts['image_folders'] = folders
        self.opts['image_folder'] = folders[0] if folders else ''
        self.imglib.set_folders(folders, recursive=self.opts.get('scan_subfolders', True))
        self._refresh_folder_list()
        g = self.imglib.groups()
        n_img = sum(len(v) for v in g.values())
        if folders:
            self.lbl_images.configure(
                text=f'{len(folders)} thư mục · {n_img} ảnh · {len(g)} nhóm')
        else:
            self.lbl_images.configure(text='(chưa chọn)')

    def _refresh_folder_list(self):
        frm = getattr(self, 'frm_folders', None)
        if frm is None:
            return
        for w in frm.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass
        folders = self.opts.get('image_folders') or []
        if not folders:
            ctk.CTkLabel(frm, text='Chưa có thư mục. Bấm Thêm, hoặc kéo-thả nhiều folder.',
                         text_color=MUTED, font=ctk.CTkFont(size=11),
                         wraplength=280, justify='left').pack(fill='x', padx=8, pady=8)
            return
        for p in folders:
            row = ctk.CTkFrame(frm, fg_color='transparent')
            row.pack(fill='x', padx=6, pady=2)
            ctk.CTkLabel(row, text=os.path.basename(p) or p, anchor='w',
                         text_color=TEXT, font=ctk.CTkFont(size=11)
                         ).pack(side='left', fill='x', expand=True)
            ctk.CTkButton(row, text='×', width=28, height=22, fg_color=INPUT,
                          hover_color=BORDER, text_color=TEXT,
                          command=lambda f=p: self._remove_image_folder(f)
                          ).pack(side='right')

    def _remove_image_folder(self, folder):
        key = os.path.normcase(os.path.normpath(folder or ''))
        self.opts['image_folders'] = [
            f for f in (self.opts.get('image_folders') or [])
            if os.path.normcase(os.path.normpath(f)) != key]
        self._apply_image_folders()
        self.group_idx = 0
        self._invalidate()
        self._rebuild_timeline()
        self._update_coverage_label()
        self._update_estimate()

    def select_avatars(self):
        p = filedialog.askdirectory(title='Chọn thư mục avatar')
        if not p:
            return
        self.opts['avatar_folder'] = p
        self.avatars.set_folder(p)
        self.lbl_avatars.configure(text=p)
        self._invalidate()
        self._update_coverage_label()

    def select_bg(self):
        p = filedialog.askopenfilename(
            title='Chọn ảnh nền',
            filetypes=[('Ảnh', '*.png *.jpg *.jpeg')])
        if not p:
            return
        self.opts['bg_path'] = p
        self.opts['bg_mode'] = 'image'
        self.seg_bg.set('Ảnh nền')
        self.lbl_bg.configure(text=p)
        self._invalidate()

    def _on_bg_mode(self, v):
        self.opts['bg_mode'] = 'white' if v == 'Trắng' else 'image'
        self._invalidate()

    def _on_drop(self, files):
        if not getattr(self, '_ui_built', False):
            return
        paths = []
        for f in files:
            if isinstance(f, bytes):
                for enc in ('utf-8', 'mbcs'):
                    try:
                        f = f.decode(enc)
                        break
                    except Exception:
                        continue
            if isinstance(f, str):
                paths.append(f)
        added = False
        for p in paths:
            if os.path.isdir(p):
                folders = list(self.opts.get('image_folders') or [])
                key = os.path.normcase(os.path.normpath(p))
                if not any(os.path.normcase(os.path.normpath(f)) == key for f in folders):
                    folders.append(p)
                    self.opts['image_folders'] = folders
                    added = True
            elif p.lower().endswith(('.xlsx', '.xlsm')):
                self._load_excel_async(p)
        if added:
            self.opts['image_folder'] = self.opts['image_folders'][0]
            self._apply_image_folders()
            self.group_idx = 0
            self._invalidate()
            self._rebuild_timeline()
            self._update_coverage_label()
            self._update_estimate()

    # ════════════════════════ TIMELINE ════════════════════════
    def _ordered_codes(self):
        return list(self._ordered_groups().keys())

    def _filtered_codes(self):
        codes = self._ordered_codes()
        q = self._search.strip().upper()
        if q:
            codes = [c for c in codes if q in c.upper()]
        return codes

    def _on_search(self, _=None):
        self._search = self.ent_search.get()
        self._rebuild_timeline()

    def _rebuild_timeline(self):
        tl = self.tl
        tl.delete('all')
        self._tl_pending.clear()
        self._tl_photos.clear()
        self._tl_items = []
        ch = self._tl_chrome()
        try:
            tl.configure(bg=ch['bg'])
        except Exception:
            pass
        groups = self.imglib.groups()
        all_codes = self._ordered_codes()
        codes = self._filtered_codes()
        x = 8
        for code in codes:
            idx = all_codes.index(code)
            selected = (idx == self.group_idx)
            self._tl_items.append((code, idx, x))
            tl.create_rectangle(x - 2, 6, x + _TL_W + 2, 10 + _TL_H + 14,
                                outline=(ACCENT if selected else ch['outline']),
                                width=2 if selected else 1,
                                fill=ch['card'],
                                tags=(f'g_{idx}',))
            tl.create_rectangle(x, 8, x + _TL_W, 8 + _TL_H, fill=ch['slot'],
                                outline='', tags=(f'g_{idx}',))
            paths = self.imglib.group_paths(code)
            if paths:
                src = self.thumbs.request(paths[0], self, self._tl_thumb_ready)
                if src is not None:
                    self._tl_set_photo(code, idx, x, src)
                else:
                    self._tl_pending[paths[0]] = (code, idx, x)
            label = code if len(code) <= 16 else code[:15] + '…'
            nphotos = len(groups.get(code, []))
            exif_note = ''
            if (self.opts.get('show_no_exif_badge', True) and paths
                    and not FX.has_exif_datetime(paths[0])):
                exif_note = ' ⚠'
            tl.create_text(x + _TL_W / 2, 8 + _TL_H + 9,
                           text=f'{label} ({nphotos}){exif_note}',
                           fill=ch['text'], font=('Segoe UI', 8),
                           tags=(f'g_{idx}',))
            x += _TL_W + 12
        tl.configure(scrollregion=(0, 0, x, _TL_H + 30))
        n = len(all_codes)
        cur = (self.group_idx + 1) if n else 0
        hint = ''
        if n and self.opts.get('visible', {}).get('avatar', True) and self.opts.get('avatar_folder'):
            cur_code = all_codes[max(0, min(self.group_idx, n - 1))]
            if not self._resolve_avatar(cur_code):
                hint = ' · Không có avatar'
        self.lbl_group.configure(
            text=(f'Nhóm {cur}/{n}{hint}' if n else 'Chưa có ảnh'))

    def _tl_set_photo(self, code, idx, x, src):
        fit = G.photo_fit_mode('fill', src.width, src.height)
        im = G.resize_into_box(src, _TL_W, _TL_H, fit)
        from PIL import ImageTk
        ph = ImageTk.PhotoImage(im)
        self._tl_photos[idx] = ph
        ox = x + (_TL_W - im.size[0]) / 2
        oy = 8 + (_TL_H - im.size[1]) / 2
        self.tl.create_image(ox, oy, anchor='nw', image=ph, tags=(f'g_{idx}',))

    def _tl_thumb_ready(self, path, img):
        info = self._tl_pending.pop(path, None)
        if info:
            self._tl_set_photo(info[0], info[1], info[2], img)

    def _tl_click(self, ev):
        x = self.tl.canvasx(ev.x)
        for code, idx, ix in self._tl_items:
            if ix - 2 <= x <= ix + _TL_W + 2:
                self._select_group(idx)
                return

    def _select_group(self, idx):
        codes = self._ordered_codes()
        if not codes:
            return
        self.group_idx = max(0, min(idx, len(codes) - 1))
        self._invalidate()
        self._rebuild_timeline()
        try:
            paths = self.imglib.group_paths(codes[self.group_idx])
            self.prefetch_minimap(list(paths or [])[:8])
        except Exception:
            pass

    def _step_group(self, d):
        self._select_group(self.group_idx + d)

    # ════════════════════════ PHÍM TẮT ════════════════════════
    def _bind_keys(self):
        self.bind('<Control-z>', lambda e: self.editor.undo())
        self.bind('<Control-y>', lambda e: self.editor.redo())
        self.bind('<Prior>', lambda e: self._step_group(-1))
        self.bind('<Next>', lambda e: self._step_group(1))
        ed = self.editor
        for key, dx, dy in (('<Left>', -0.02, 0), ('<Right>', 0.02, 0),
                            ('<Up>', 0, -0.02), ('<Down>', 0, 0.02)):
            ed.bind(key, lambda e, a=dx, b=dy: ed.nudge(a, b))
            ed.bind(f'<Shift-{key[1:-1]}>',
                    lambda e, a=dx, b=dy: ed.nudge(a * 10, b * 10))
        ed.bind('<Escape>', lambda e: ed.select(None))

    # ════════════════════════ PRESET BỐ CỤC ════════════════════════
    def _preset_files(self):
        try:
            return sorted(f[:-5] for f in os.listdir(PRESET_DIR)
                          if f.endswith('.json'))
        except Exception:
            return []

    def _refresh_preset_menu(self):
        names = [n for n in self._preset_files() if n not in BUILTIN_PRESETS]
        self.om_preset.configure(values=['— Preset —'] + names)
        self.om_preset.set('— Preset —')

    def _snapshot_dept(self):
        key = self.opts.get('dept') or 'sales'
        stores = self.opts.setdefault('dept_state', {})
        stores[key] = {
            'layout': copy.deepcopy(self.layout),
            'visible': dict(self.opts.get('visible') or {}),
            'channel_enabled': bool(self.opts.get('channel_enabled', True)),
            'slide_style': self.opts.get('slide_style') or 'report',
        }

    def _sync_dept_ui(self):
        lab = dept_label(self.opts.get('dept'))
        try:
            self.title('AutoPPTX Studio V2  ·  ' + lab)
        except Exception:
            pass
        seg = getattr(self, 'seg_dept', None)
        if seg is None:
            return
        ready = getattr(self, '_dept_ready', False)
        self._dept_ready = False
        try:
            seg.set(lab)
        except Exception:
            pass
        self._dept_ready = ready

    def _on_dept(self, label):
        if not getattr(self, '_dept_ready', False):
            return
        dept = dept_key(label)
        if dept == self.opts.get('dept'):
            return
        self._snapshot_dept()
        self.opts['dept'] = dept
        saved = (self.opts.get('dept_state') or {}).get(dept)
        if saved and saved.get('layout'):
            pack = {
                'layout': saved['layout'],
                'visible': saved.get('visible') or {},
                'channel_enabled': saved.get('channel_enabled', True),
                'slide_style': saved.get('slide_style') or (
                    'saleskit' if dept == 'bd' else 'report'),
            }
        else:
            pack = pack_for_dept(dept)
        self._apply_pack(pack, log_name='phòng ' + dept_label(dept))

    def _apply_pack(self, pack, log_name=''):
        if hasattr(self, 'editor'):
            self.editor.push_undo()
        self.layout.clear()
        self.layout.update(copy.deepcopy(pack['layout']))
        style = pack.get('slide_style') or 'report'
        self.opts['slide_style'] = style
        self.opts['dept'] = 'bd' if style == 'saleskit' else 'sales'
        vis = self.opts.setdefault('visible', {})
        vis.update(pack.get('visible') or {})
        if 'channel_enabled' in pack:
            self.opts['channel_enabled'] = bool(pack['channel_enabled'])
        vis['channel'] = bool(self.opts.get('channel_enabled', True))
        self._sync_dept_ui()
        self._sync_channel_switch()
        self._invalidate()
        if hasattr(self, 'editor'):
            self.editor.render()
        if hasattr(self, 'tl'):
            self._rebuild_timeline()
        if log_name:
            self.log(f'Đã áp mẫu "{log_name}" · slide theo tỉnh/thành Bắc → Nam, rồi quận.')

    def _save_preset(self):
        dlg = ctk.CTkInputDialog(text='Tên preset:', title='Lưu bố cục')
        name = (dlg.get_input() or '').strip()
        if not name:
            return
        if name in BUILTIN_PRESETS:
            messagebox.showwarning('Preset',
                                   'Tên này là mẫu có sẵn. Hãy đặt tên khác.')
            return
        os.makedirs(PRESET_DIR, exist_ok=True)
        safe = ''.join(ch for ch in name if ch not in r'\/:*?"<>|')
        payload = {
            'layout': self.layout,
            'slide_style': self.opts.get('slide_style') or 'report',
            'visible': dict(self.opts.get('visible') or {}),
            'channel_enabled': bool(self.opts.get('channel_enabled', True)),
        }
        with open(os.path.join(PRESET_DIR, safe + '.json'), 'w',
                  encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
        self._refresh_preset_menu()
        self.om_preset.set(safe)
        self.log(f'Đã lưu preset "{safe}".')

    def _apply_preset(self, name):
        if not getattr(self, '_preset_ready', False):
            return
        if not name or name == '— Preset —':
            return
        if name in BUILTIN_PRESETS:
            self._apply_pack(builtin_pack(name), log_name=name)
            return
        try:
            with open(os.path.join(PRESET_DIR, name + '.json'), 'r',
                      encoding='utf-8') as f:
                loaded = json.load(f)
        except Exception as e:
            messagebox.showerror('Preset', f'Không đọc được preset:\n{e}')
            return
        if isinstance(loaded, dict) and 'layout' in loaded:
            pack = {
                'layout': _deep_merge(DEFAULT_LAYOUT, loaded.get('layout')),
                'slide_style': loaded.get('slide_style') or 'report',
                'visible': loaded.get('visible') or {},
                'channel_enabled': loaded.get('channel_enabled', True),
            }
        else:
            pack = {
                'layout': _deep_merge(DEFAULT_LAYOUT, loaded),
                'slide_style': 'saleskit' if float(
                    (loaded or {}).get('info', {}).get('w') or 0) >= 8 else 'report',
            }
        self._apply_pack(pack, log_name=name)

    def _delete_preset(self):
        name = self.om_preset.get()
        if name in BUILTIN_PRESETS or name == '— Preset —':
            messagebox.showinfo('Preset', 'Chọn preset đã lưu của bạn để xoá.')
            return
        if not name:
            return
        if not messagebox.askyesno('Xoá preset', f'Xoá preset "{name}"?'):
            return
        try:
            os.remove(os.path.join(PRESET_DIR, name + '.json'))
        except Exception:
            pass
        self._refresh_preset_menu()

    def _reset_layout(self):
        dept = self.opts.get('dept') or 'sales'
        lab = dept_label(dept)
        if not messagebox.askyesno('Bố cục', f'Khôi phục mẫu mặc định phòng {lab}?'):
            return
        self.opts.setdefault('dept_state', {}).pop(dept, None)
        self._apply_pack(pack_for_dept(dept), log_name=lab)
        try:
            self.om_preset.set('— Preset —')
        except Exception:
            pass

    # ════════════════════════ COVERAGE ════════════════════════
    def _list_is_loaded(self):
        return bool(getattr(self.excel, 'rows', None) or self._by_code)

    def _current_list_name(self):
        path = (getattr(self.excel, 'path', None)
                or (self.opts or {}).get('excel_path') or '')
        return os.path.basename(path) if path else ''

    def _current_list_caption(self):
        n = len(self._excel_code_list())
        name = self._current_list_name()
        if not self._list_is_loaded() or not n:
            return 'Chưa có FILE TỔNG / MASTER — đồng bộ hoặc chọn ở tab Nguồn.'
        if name:
            return f'FILE TỔNG / MASTER đang dùng: {name}  ·  {n} mã Code_RP'
        return f'FILE TỔNG / MASTER đang dùng  ·  {n} mã Code_RP'

    def _need_loaded_list(self, title='So sánh ảnh với FILE TỔNG / MASTER'):
        """True nếu đã có file tổng trong bộ nhớ. Không mở hộp thoại chọn file khác."""
        if self._list_is_loaded() and self._excel_code_list():
            return True
        messagebox.showinfo(
            title,
            'Chưa có FILE TỔNG / MASTER để so sánh với ảnh.\n\n'
            'Vào tab Nguồn:\n'
            '• Đám mây: bấm Tải FILE TỔNG + Avatar từ Cloud\n'
            '• Máy này: Chọn FILE TỔNG (MASTER) · đẩy Cloud: Cài đặt\n\n'
            'Đây là file tổng chính, không phải FILE LIST SO SÁNH.')
        return False

    def _suggest_from_list(self, img_code, codes=None):
        """Gợi ý Code_RP gần đúng trên file tổng đang mở (không đổi tên)."""
        import difflib
        codes = list(codes if codes is not None else self._excel_code_list())
        up_map = {c.upper(): c for c in codes}
        key = str(img_code or '').strip().upper()
        if not key or not up_map:
            return None
        if key in up_map:
            return up_map[key]
        hits = difflib.get_close_matches(key, list(up_map.keys()), n=1, cutoff=0.45)
        return up_map[hits[0]] if hits else None

    def _coverage(self):
        groups = self.imglib.groups()
        res = {'exact': [], 'merged': [], 'fuzzy': [], 'none': [],
               'no_avatar': []}
        by_code = self._by_code or {}
        for code in groups:
            _, mcode, kind = match_row(code, by_code, self._merged)
            res[kind].append((code, mcode))
            if self.opts['avatar_folder'] and not self._resolve_avatar(code):
                res['no_avatar'].append(code)
        return res

    def _update_coverage_label(self):
        try:
            if hasattr(self, 'lbl_compare_list'):
                self.lbl_compare_list.configure(text=self._current_list_caption())
        except Exception:
            pass
        groups = self.imglib.groups()
        if not groups:
            self.lbl_coverage.configure(text='')
            return
        if not self._list_is_loaded() or not self._by_code:
            extra = ' · chưa có list' if not self._list_is_loaded() else ''
            self.lbl_coverage.configure(text=f'{len(groups)} nhóm ảnh{extra}')
            return
        cov = self._coverage()
        ok = len(cov['exact']) + len(cov['merged'])
        qa = QA.check(groups, self._by_code, self._merged)
        txt = (f"{len(groups)} nhóm • {ok} khớp Excel"
               + (f" • {len(cov['fuzzy'])} gần đúng" if cov['fuzzy'] else '')
               + (f" • {len(qa.unmatched)} lệch mã" if qa.unmatched else '')
               + (f" • {len(qa.missing)} list chưa có ảnh" if qa.missing else '')
               + (f" • {len(cov['no_avatar'])} thiếu avatar"
                  if cov['no_avatar'] else ''))
        self.lbl_coverage.configure(text=txt)

    def _show_coverage(self):
        groups = self.imglib.groups()
        if not groups:
            messagebox.showinfo('So sánh ảnh với FILE TỔNG', 'Chưa chọn thư mục ảnh.')
            return
        if not self._need_loaded_list('So sánh ảnh với FILE TỔNG'):
            return
        cov = self._coverage()
        ov = build_overview(groups, self._by_code, self._merged)
        win = ctk.CTkToplevel(self)
        win.title('So sánh ẢNH với FILE TỔNG / MASTER')
        win.geometry('720x520')
        win.attributes('-topmost', True)
        tb = ctk.CTkTextbox(win, font=ctk.CTkFont(family='Consolas', size=12))
        tb.pack(fill='both', expand=True, padx=10, pady=(10, 4))
        cap = self._current_list_caption()
        n_codes = len(self._excel_code_list())
        lines = [
            cap,
            f'Số mã Code_RP trên list : {n_codes}',
            '',
            f'Tổng cửa hàng (ảnh) : {ov["n_stores"]}',
            f'Tổng ảnh            : {ov["n_photos"]}',
            f'Tổng màn hình list  : {ov["n_screens"]}',
            f'ĐỦ (ảnh ≥ màn)      : {ov["n_du"]}',
            f'THIẾU ảnh           : {ov["n_thieu"]}',
            f'Thừa ảnh            : {ov["n_thua"]}',
            f'Chưa khớp list      : {ov["n_chua_excel"]}',
            f'List chưa có ảnh    : {len(ov["excel_no_photo"])}',
            '',
            f'Khớp list chính xác: {len(cov["exact"])}  ·  gộp block: {len(cov["merged"])}'
            f'  ·  gần đúng: {len(cov["fuzzy"])}  ·  không khớp: {len(cov["none"])}',
            '',
            f'{"Mã":<16} {"Ảnh":>4}  {"Màn":>4}  {"Chênh":>6}  Kết luận',
            '-' * 64,
        ]
        for r in ov['rows']:
            dtxt = f'{r["delta"]:+d}' if r['screens'] else '—'
            lines.append(
                f'{r["code"]:<16} {r["photos"]:>4}  {r["screens"]:>4}  {dtxt:>6}  {r["label"]}')
        if ov['excel_no_photo']:
            lines.append('')
            lines.append(f'— List đang mở còn {len(ov["excel_no_photo"])} điểm chưa có ảnh '
                         f'(hiện {min(40, len(ov["excel_no_photo"]))} mã):')
            lines += [f'   {c}' for c in ov['excel_no_photo'][:40]]
        if cov['fuzzy']:
            lines.append('')
            lines.append('— Khớp gần đúng với FILE TỔNG / MASTER (kiểm tra tên ảnh):')
            lines += [f'   {a}  ≈  {b}' for a, b in cov['fuzzy']]
        if cov['none']:
            lines.append('')
            lines.append('— Tên ảnh không khớp mã trên FILE TỔNG / MASTER:')
            lines += [f'   {c}' for c, _ in cov['none'][:40]]
            if len(cov['none']) > 40:
                lines.append(f'   … còn {len(cov["none"]) - 40} nhóm')
        if cov['no_avatar']:
            lines.append('')
            lines.append('— Thiếu avatar:')
            lines += [f'   {c}' for c in cov['no_avatar']]
        tb.insert('1.0', '\n'.join(lines))
        tb.configure(state='disabled')
        bar = ctk.CTkFrame(win, fg_color='transparent')
        bar.pack(fill='x', padx=10, pady=(0, 12))
        ctk.CTkButton(
            bar, text='Đóng', width=90, height=32,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=win.destroy
        ).pack(side='right')
        if cov['fuzzy'] or cov['none']:
            ctk.CTkButton(
                bar, text='Sửa mã ảnh / đổi tên file…', height=32,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=lambda: (win.destroy(), self._show_fix_codes_dialog())
            ).pack(side='right', padx=(0, 8))

    def _list_master_compare_caption(self):
        path = str((self.opts or {}).get('compare_list_path') or '').strip()
        master = self._current_list_name()
        if path and os.path.isfile(path):
            return (f'FILE SO SÁNH CODE: {os.path.basename(path)}\n'
                    f'FILE MASTER CHẠY PPTX: {master or "chưa mở"}\n'
                    'Hai file độc lập; file so sánh không ghi đè Master.')
        return (f'FILE MASTER CHẠY PPTX: {master or "chưa mở"}.\n'
                'Chọn một file khác có cột Code_RP để đối chiếu; '
                'file này không thay đổi Master.')

    def _choose_list_for_master_compare(self):
        if not getattr(self.excel, 'rows', None):
            messagebox.showinfo(
                'FILE SO SÁNH CODE',
                'Chưa có FILE TỔNG / MASTER đang mở.\n\n'
                'Vào Bước 1 để đồng bộ Cloud hoặc Bước 2 để chọn FILE TỔNG.')
            return
        folders = list(getattr(self.imglib, 'folders', None) or [])
        path = filedialog.askopenfilename(
            title='Chọn FILE SO SÁNH CODE với FILE TỔNG / MASTER',
            filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not path:
            return

        busy = ctk.CTkToplevel(self)
        busy.title('Đang đối chiếu…')
        busy.geometry('360x120')
        busy.attributes('-topmost', True)
        busy_text = ('Đang đọc FILE SO SÁNH CODE và kiểm tra thư mục ảnh…'
                     if folders else
                     'Đang đọc FILE SO SÁNH CODE và đối chiếu Code_RP…')
        ctk.CTkLabel(
            busy, text=busy_text,
            font=ctk.CTkFont(size=13, weight='bold'),
            wraplength=320).pack(padx=18, pady=34)

        master_by_code = self.excel.by_code(None)
        merged = build_merged_groups(master_by_code)
        recursive = bool(getattr(self.imglib, 'recursive', True))
        master_name = self._current_list_name()

        def work():
            src = ExcelSource()
            if not src.load_compare_list(path):
                self.after(0, lambda: self._after_list_master_compare(
                    busy, path, master_name, None,
                    src.error or 'Không đọc được FILE SO SÁNH CODE.'))
                return
            groups = {}
            if folders:
                library = ImageLibrary()
                library.set_folders(folders, recursive=recursive)
                groups = library.groups()
            report = compare_list_with_master(
                src.rows, master_by_code, groups, merged,
                check_images=bool(folders))
            self.after(0, lambda: self._after_list_master_compare(
                busy, path, master_name, report, ''))

        threading.Thread(target=work, daemon=True).start()

    def _after_list_master_compare(self, busy, path, master_name, report, error):
        try:
            busy.destroy()
        except Exception:
            pass
        if error:
            messagebox.showerror(
                'FILE SO SÁNH CODE',
                f'Không đọc được FILE SO SÁNH CODE:\n{error}\n\n'
                'File cần có cột Code_RP (hoặc Code / Mã báo cáo).')
            return
        self.opts['compare_list_path'] = path
        self._schedule_save()
        try:
            self.lbl_list_master_compare.configure(
                text=self._list_master_compare_caption())
        except Exception:
            pass
        self._show_list_master_compare_report(path, master_name, report or {})

    def _show_list_master_compare_report(self, list_path, master_name, report):
        win = ctk.CTkToplevel(self)
        images_checked = bool(report.get('images_checked'))
        win.title('So sánh CODE với FILE MASTER')
        win.geometry('920x650')
        win.attributes('-topmost', True)
        tb = ctk.CTkTextbox(win, font=ctk.CTkFont(family='Consolas', size=12))
        tb.pack(fill='both', expand=True, padx=10, pady=(10, 4))
        lines = [
            'ĐỐI CHIẾU FILE SO SÁNH CODE VỚI FILE TỔNG / MASTER',
            f'FILE SO SÁNH CODE : {os.path.basename(list_path)}',
            f'FILE TỔNG / MASTER: {master_name or "—"}',
            '',
            f'Tổng mã cần kiểm tra : {report.get("n_list", 0)}',
            f'Khớp file tổng       : {report.get("n_master_matches", 0)}',
            f'Không có trong tổng  : {report.get("n_not_in_master", 0)}',
        ]
        if images_checked:
            lines += [
                f'Điểm đủ ảnh          : {report.get("n_ok_sites", 0)}',
                f'Điểm thiếu ảnh       : {report.get("n_missing_sites", 0)}',
                f'Tổng ảnh yêu cầu     : {report.get("n_required", 0)}',
                f'Tổng ảnh đã có       : {report.get("n_photos", 0)}',
                f'Tổng ảnh còn thiếu   : {report.get("n_missing_images", 0)}',
                '',
                'TỔNG HỢP THEO KÊNH',
                '-' * 92,
                f'{"Kênh":<30} {"Điểm":>6} {"Đủ":>6} {"Thiếu":>7} '
                f'{"Ngoài tổng":>11} {"Ảnh thiếu":>10}',
                '-' * 92,
            ]
        else:
            lines += [
                '',
                'Lưu ý: Chưa chọn thư mục ảnh — báo cáo này chỉ so sánh Code_RP.',
                'Muốn kiểm tra ảnh thiếu, chọn thư mục ảnh ở Bước 3 rồi chạy lại Bước 4.',
                '',
                'TỔNG HỢP THEO KÊNH',
                '-' * 76,
                f'{"Kênh":<38} {"Mã list":>9} {"Khớp Master":>12} {"Ngoài Master":>13}',
                '-' * 76,
            ]
        for item in report.get('channels') or []:
            if images_checked:
                channel = str(item.get('channel') or '')[:30]
                lines.append(
                    f'{channel:<30} {item.get("sites", 0):>6} '
                    f'{item.get("ok_sites", 0):>6} {item.get("missing_sites", 0):>7} '
                    f'{item.get("not_in_master", 0):>11} {item.get("missing", 0):>10}')
            else:
                channel = str(item.get('channel') or '')[:38]
                sites = item.get('sites', 0)
                outside = item.get('not_in_master', 0)
                lines.append(
                    f'{channel:<38} {sites:>9} {sites - outside:>12} {outside:>13}')

        if images_checked:
            detail = [r for r in (report.get('rows') or [])
                      if r.get('status') != 'ok']
        else:
            detail = [r for r in (report.get('rows') or [])
                      if r.get('status') == 'not_in_master']
        if detail:
            lines += ['', 'CHI TIẾT CẦN XỬ LÝ']
            current_channel = None
            for item in detail:
                channel = item.get('channel') or 'Không xác định'
                if channel != current_channel:
                    current_channel = channel
                    lines += [
                        '', f'[{channel}]',
                        f'{"Mã Code_RP":<20} {"Cần":>5} {"Có":>5} '
                        f'{"Thiếu":>6}  Kết luận / Địa điểm',
                        '-' * 92,
                    ]
                if item.get('status') == 'not_in_master':
                    conclusion = 'KHÔNG CÓ TRONG FILE TỔNG'
                else:
                    conclusion = 'THIẾU ẢNH'
                place = item.get('name') or item.get('district') or ''
                if place:
                    conclusion += ' · ' + str(place)
                lines.append(
                    f'{str(item.get("code") or "")[:20]:<20} '
                    f'{item.get("required", 0):>5} {item.get("photos", 0):>5} '
                    f'{item.get("missing", 0):>6}  {conclusion}')
        elif images_checked:
            lines += ['', '✓ Tất cả mã trong FILE SO SÁNH đã đủ ảnh.']
        else:
            lines += ['', '✓ Tất cả mã trong FILE SO SÁNH đều có trong FILE MASTER.']

        report_text = '\n'.join(lines)
        tb.insert('1.0', report_text)
        tb.configure(state='disabled')
        bar = ctk.CTkFrame(win, fg_color='transparent')
        bar.pack(fill='x', padx=10, pady=(0, 12))

        def copy_report():
            self.clipboard_clear()
            self.clipboard_append(report_text)
            messagebox.showinfo('FILE SO SÁNH CODE', 'Đã sao chép báo cáo.')

        ctk.CTkButton(
            bar, text='Sao chép báo cáo', height=32,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            command=copy_report).pack(side='right', padx=(8, 0))
        ctk.CTkButton(
            bar, text='Đóng', width=90, height=32,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=win.destroy).pack(side='right')

    # ════════════════════════ SỬA MÃ ẢNH / ĐỔI TÊN FILE ════════════════════════
    def _excel_code_list(self):
        """Code_RP trên FILE TỔNG / MASTER đang mở."""
        out, seen = [], set()

        def add(s):
            s = str(s or '').strip()
            if s and s.lower() != 'nan' and s.upper() not in seen:
                seen.add(s.upper())
                out.append(s)

        for r in (getattr(self.excel, 'rows', None) or []):
            add(r.get('Code_RP'))
        for k, row in (self._by_code or {}).items():
            add((row or {}).get('Code_RP') or k)
        out.sort(key=lambda x: x.upper())
        return out

    def _mismatch_photo_items(self):
        """Nhóm ảnh không khớp list + khớp gần đúng: [(mã ảnh, mã gợi ý|None)]."""
        cov = self._coverage()
        codes = self._excel_code_list()
        items = []
        for c, _ in cov['none']:
            items.append((c, self._suggest_from_list(c, codes)))
        items += [(a, b) for a, b in cov['fuzzy']]
        return items

    def _photo_rel(self, path):
        for folder in self.imglib.folders:
            try:
                rel = os.path.relpath(path, folder)
                if not str(rel).startswith('..'):
                    return rel
            except Exception:
                pass
        return os.path.basename(path)

    def _pick_excel_code(self, img_code, codes):
        """Chọn Code_RP trên FILE TỔNG / MASTER: gợi ý gần giống lên đầu."""
        import difflib
        win = ctk.CTkToplevel(self)
        win.title('Chọn mã trên FILE TỔNG / MASTER')
        win.geometry('460x520')
        win.attributes('-topmost', True)
        try:
            win.transient(self)
            win.grab_set()
        except Exception:
            pass
        result = {'code': None}
        up_map = {c.upper(): c for c in codes}
        sugg = [up_map[s] for s in difflib.get_close_matches(
            str(img_code).upper(), list(up_map.keys()), n=8, cutoff=0.3)]
        list_name = self._current_list_name() or 'FILE TỔNG / MASTER'

        ctk.CTkLabel(
            win, text=f'Ảnh: {img_code}',
            font=ctk.CTkFont(size=13, weight='bold'), text_color=TEXT
        ).pack(anchor='w', padx=14, pady=(12, 2))
        ctk.CTkLabel(
            win, text=f'Chọn Code_RP trên FILE TỔNG / MASTER: {list_name}',
            text_color=MUTED, font=ctk.CTkFont(size=11),
            wraplength=420, justify='left', anchor='w'
        ).pack(anchor='w', padx=14, pady=(0, 4))
        ent = ctk.CTkEntry(
            win, placeholder_text='Gõ để tìm mã trên FILE TỔNG / MASTER…', height=32)
        ent.pack(fill='x', padx=14, pady=(4, 6))
        lst = ctk.CTkScrollableFrame(win, fg_color=INPUT, corner_radius=8)
        lst.pack(fill='both', expand=True, padx=14, pady=(0, 4))
        info = ctk.CTkLabel(win, text='', text_color=MUTED, font=ctk.CTkFont(size=10))
        info.pack(anchor='w', padx=14)
        maxn = 300

        def choose(c):
            result['code'] = c
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

        def render(_e=None):
            q = ent.get().strip().upper()
            for w in lst.winfo_children():
                w.destroy()
            if not q and sugg:
                ctk.CTkLabel(
                    lst, text='GỢI Ý GẦN GIỐNG TRÊN LIST ĐANG MỞ',
                    font=ctk.CTkFont(size=10, weight='bold'), text_color=MUTED
                ).pack(anchor='w', padx=6, pady=(4, 2))
                for c in sugg:
                    ctk.CTkButton(
                        lst, text=c, height=28, corner_radius=6, anchor='w',
                        fg_color=ACCENT, hover_color=ACCENT_HOVER,
                        font=ctk.CTkFont(size=11),
                        command=lambda cc=c: choose(cc)
                    ).pack(fill='x', padx=6, pady=1)
                ctk.CTkLabel(
                    lst, text='TẤT CẢ MÃ TRÊN LIST ĐANG MỞ',
                    font=ctk.CTkFont(size=10, weight='bold'), text_color=MUTED
                ).pack(anchor='w', padx=6, pady=(8, 2))
            pool = [c for c in codes if q in c.upper()] if q else codes
            for c in pool[:maxn]:
                ctk.CTkButton(
                    lst, text=c, height=26, corner_radius=6, anchor='w',
                    fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                    font=ctk.CTkFont(size=11),
                    command=lambda cc=c: choose(cc)
                ).pack(fill='x', padx=6, pady=1)
            info.configure(
                text=(f'Hiện {maxn}/{len(pool)} mã — gõ thêm để lọc bớt'
                      if len(pool) > maxn else f'{len(pool)} mã trên FILE TỔNG / MASTER'))

        ent.bind('<KeyRelease>', render)
        render()
        ctk.CTkButton(
            win, text='Bỏ qua', height=30, corner_radius=8,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=win.destroy
        ).pack(pady=(2, 12))
        win.wait_window()
        return result['code']

    def _show_fix_codes_dialog(self):
        """Liệt kê ảnh lệch mã / khớp gần đúng trên FILE TỔNG, xác nhận rồi đổi tên."""
        groups = self.imglib.groups()
        if not groups:
            messagebox.showinfo('Sửa mã ảnh', 'Chưa chọn thư mục ảnh.')
            return
        if not self._need_loaded_list('Sửa mã ảnh'):
            return
        codes = self._excel_code_list()
        items = self._mismatch_photo_items()
        if not items:
            messagebox.showinfo(
                'Sửa mã ảnh',
                'Không có mã ảnh nào cần sửa — tất cả đã khớp chính xác với FILE TỔNG / MASTER.')
            return

        win = ctk.CTkToplevel(self)
        win.title('Sửa mã ảnh · FILE TỔNG / MASTER')
        win.geometry('640x580')
        win.attributes('-topmost', True)
        try:
            win.transient(self)
            win.grab_set()
        except Exception:
            pass
        n_none = sum(1 for _c, sug in items if not sug)
        n_fz = len(items) - n_none
        ctk.CTkLabel(
            win, text=f'Có {len(items)} nhóm ảnh cần kiểm: '
                      f'{n_none} lệch mã · {n_fz} gần đúng',
            font=ctk.CTkFont(size=13, weight='bold'), text_color=TEXT
        ).pack(anchor='w', padx=16, pady=(14, 2))
        ctk.CTkLabel(
            win, text=self._current_list_caption() + '\n'
                      'Chọn đúng Code_RP trên list này, rồi bấm Đổi tên file. '
                      'App không đổi tên cho đến khi bạn xác nhận.',
            text_color=MUTED, font=ctk.CTkFont(size=11), justify='left',
            wraplength=600
        ).pack(anchor='w', padx=16, pady=(0, 8))
        body = ctk.CTkScrollableFrame(win, fg_color=INPUT, corner_radius=8)
        body.pack(fill='both', expand=True, padx=16, pady=(0, 8))

        chosen = {}
        sug_rows = []
        lbl_bottom = ctk.CTkLabel(
            win, text='Chưa chọn mã nào.', text_color=MUTED,
            font=ctk.CTkFont(size=11))

        def update_bottom():
            lbl_bottom.configure(
                text=(f'Đã chọn {len(chosen)} nhóm sẽ đổi tên.'
                      if chosen else 'Chưa chọn mã nào.'))

        def use_all_suggestions():
            n = 0
            for c, sug, lb in sug_rows:
                if not sug:
                    continue
                chosen[c] = sug
                n += 1
                try:
                    lb.configure(text=f'→ sẽ đổi thành: {sug}', text_color=ACCENT)
                except Exception:
                    pass
            update_bottom()
            if n:
                self.log(f'Đã chọn {n} gợi ý FILE TỔNG / MASTER (chưa đổi tên).')

        for img_code, fuzzy_to in items:
            n_files = len(groups.get(img_code, []))
            row = ctk.CTkFrame(body, fg_color=CARD, corner_radius=8)
            row.pack(fill='x', padx=4, pady=3)
            left = ctk.CTkFrame(row, fg_color='transparent')
            left.pack(side='left', fill='x', expand=True, padx=10, pady=7)
            mark = '≈' if fuzzy_to else '?'
            ctk.CTkLabel(
                left, text=f'{mark}  {img_code}   ({n_files} ảnh)',
                font=ctk.CTkFont(size=12, weight='bold'), text_color=TEXT,
                anchor='w'
            ).pack(anchor='w')
            sub = (f"gợi ý từ FILE TỔNG / MASTER: {fuzzy_to}" if fuzzy_to
                   else 'không có trên FILE TỔNG / MASTER')
            lbl_new = ctk.CTkLabel(
                left, text=sub, text_color=MUTED,
                font=ctk.CTkFont(size=10), anchor='w')
            lbl_new.pack(anchor='w')

            def pick(c=img_code, lb=lbl_new):
                new = self._pick_excel_code(c, codes)
                if new:
                    chosen[c] = new
                    lb.configure(text=f'→ sẽ đổi thành: {new}', text_color=ACCENT)
                    update_bottom()

            def use_sug(c=img_code, sug=fuzzy_to, lb=lbl_new):
                if not sug:
                    return
                chosen[c] = sug
                lb.configure(text=f'→ sẽ đổi thành: {sug}', text_color=ACCENT)
                update_bottom()

            btns = ctk.CTkFrame(row, fg_color='transparent')
            btns.pack(side='right', padx=(4, 10), pady=6)
            ctk.CTkButton(
                btns, text='Chọn mã…', width=96, height=30, corner_radius=8,
                font=ctk.CTkFont(size=11, weight='bold'),
                fg_color=ACCENT, hover_color=ACCENT_HOVER, command=pick
            ).pack(side='right')
            if fuzzy_to:
                sug_rows.append((img_code, fuzzy_to, lbl_new))
                ctk.CTkButton(
                    btns, text='Dùng gợi ý', width=96, height=30, corner_radius=8,
                    font=ctk.CTkFont(size=11),
                    fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                    command=use_sug
                ).pack(side='right', padx=(0, 6))

        lbl_bottom.pack(anchor='w', padx=16)

        def do_rename():
            if not chosen:
                messagebox.showinfo('Sửa mã ảnh', 'Bạn chưa chọn mã mới cho nhóm nào.')
                return
            plan, conflicts = build_photo_rename_plan(self.imglib.groups(), chosen)
            if not plan and not conflicts:
                messagebox.showinfo('Sửa mã ảnh', 'Không có file nào cần đổi tên.')
                return
            prev = '\n'.join(
                f'  • {self._photo_rel(a)}  →  {os.path.basename(b)}'
                for a, b in plan[:15])
            if len(plan) > 15:
                prev += f'\n  … và {len(plan) - 15} file nữa'
            msg = f'Sẽ ĐỔI TÊN {len(plan)} file ảnh:\n\n{prev}'
            if conflicts:
                msg += (
                    f'\n\n⚠ {len(conflicts)} file BỊ BỎ QUA vì tên đích đã tồn tại:\n'
                    + '\n'.join(
                        f'  • {self._photo_rel(a)}  →  {os.path.basename(b)}'
                        for a, b in conflicts[:8]))
            msg += '\n\nThao tác này đổi tên file THẬT trên ổ đĩa. Tiếp tục?'
            if not messagebox.askyesno('Xác nhận đổi tên', msg):
                return
            ok = err = 0
            for a, b in plan:
                try:
                    os.rename(a, b)
                    ok += 1
                except Exception as e:
                    err += 1
                    print('Rename err:', a, '->', b, e)
            self.imglib.invalidate()
            self._apply_image_folders()
            self._invalidate()
            self._rebuild_timeline()
            self._update_coverage_label()
            self._update_estimate()
            res = f'Đã đổi tên {ok} file'
            if err:
                res += f', {err} lỗi'
            if conflicts:
                res += f', {len(conflicts)} bỏ qua (trùng tên)'
            self.log(res)
            messagebox.showinfo('Đổi tên file', res)
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

        bar = ctk.CTkFrame(win, fg_color='transparent')
        bar.pack(fill='x', padx=16, pady=(4, 12))
        if sug_rows:
            ctk.CTkButton(
                bar, text=f'Chọn tất cả dùng gợi ý ({len(sug_rows)})',
                height=34, corner_radius=8,
                fg_color=CARD, hover_color=BORDER, text_color=TEXT,
                font=ctk.CTkFont(size=12),
                command=use_all_suggestions
            ).pack(side='left')
        ctk.CTkButton(
            bar, text='Đổi tên file', height=34, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=12, weight='bold'),
            command=do_rename
        ).pack(side='right')
        ctk.CTkButton(
            bar, text='Đóng', height=34, width=90, corner_radius=8,
            fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
            command=win.destroy
        ).pack(side='right', padx=6)

    # ════════════════════════ XUẤT PPTX ════════════════════════
    def _update_estimate(self):
        groups = self._ordered_groups()
        if not groups:
            short, full = 'Chưa chọn thư mục ảnh.', 'Chưa chọn thư mục ảnh.'
        else:
            ov, n_ov = None, 0
            if self.opts.get('fx', {}).get('dashboard', True):
                ov = build_overview(groups, self._by_code, self._merged)
                n_ov = len(ov['rows'])
            ng, ns, np_ = exporter.estimate(
                groups, self.opts['n_per_slide'],
                overview_rows=n_ov,
                slides_per_file=self.opts.get('slides_per_file', 0),
                by_code=self._by_code, merged=self._merged,
                pad_blank=self.opts.get('pad_blank_slides', False))
            extra = ''
            qa = QA.check(groups, self._by_code, self._merged)
            if qa.missing:
                extra += f'\n• {len(qa.missing)} điểm list chưa có ảnh'
            if qa.unmatched:
                extra += f'\n• {len(qa.unmatched)} nhóm ảnh lệch mã list'
            if qa.fuzzy:
                extra += f'\n• {len(qa.fuzzy)} nhóm khớp gần đúng'
            short = f'{ng} nhóm · ~{ns} slide'
            if self.opts.get('pad_blank_slides'):
                short += ' · pad màn'
            if qa.missing:
                short += f' · thiếu {len(qa.missing)} ảnh'
            cap = G.file_part_limit(self.opts.get('slides_per_file', 0))
            if np_ <= 1 and cap <= 0:
                files_txt = '1 file PPTX (gộp hết)'
            elif np_ <= 1:
                files_txt = f'1 file PPTX (ngưỡng {cap} slide)'
            else:
                files_txt = f'{np_} file PPTX (mỗi {cap} slide)'
            full = f'• {ng} nhóm ảnh\n• ~{ns} slide\n• {files_txt}{extra}'
            if self.opts.get('pad_blank_slides'):
                full += '\n• Slide trắng/ô trống theo số màn Excel'
            split = self.opts.get('split_export_by', 'none')
            if split and split != 'none':
                full += f'\n• Tách file theo {split}'
        try:
            self.lbl_est.configure(text=short)
        except Exception:
            pass
        try:
            self.lbl_scope_est.configure(text=self._scope_estimate_text())
        except Exception:
            pass
        try:
            self.lbl_est_full.configure(text=full)
        except Exception:
            pass

    def export(self):
        groups_rel = self.imglib.groups()
        if not groups_rel:
            messagebox.showwarning('Xuất PPTX', 'Chưa chọn thư mục ảnh.')
            return
        if self.opts['bg_mode'] == 'image' and not self.opts['bg_path']:
            messagebox.showwarning('Xuất PPTX', 'Chưa chọn ảnh nền.')
            return
        qa = QA.check(groups_rel, self._by_code, self._merged)
        if self.opts.get('qa_gate', True) and qa.needs_gate():
            self._show_qa_gate(qa, lambda: self._export_after_qa(qa, confirm=False))
            return
        self._export_after_qa(qa)

    def _export_after_qa(self, qa, confirm=True):
        groups = self._ordered_groups()
        n_ov = 0
        if self.opts.get('fx', {}).get('dashboard', True):
            ov = build_overview(groups, self._by_code, self._merged)
            n_ov = len(ov['rows'])
        ng, ns, np_ = exporter.estimate(
            groups, self.opts['n_per_slide'],
            overview_rows=n_ov,
            slides_per_file=self.opts.get('slides_per_file', 0),
            by_code=self._by_code, merged=self._merged,
            pad_blank=self.opts.get('pad_blank_slides', False))
        cap = G.file_part_limit(self.opts.get('slides_per_file', 0))
        if np_ <= 1:
            split_txt = '1 file'
        else:
            split_txt = f'chia {np_} file, mỗi {cap} slide'
        if confirm and not messagebox.askyesno(
                'Xác nhận xuất',
                f'{ng} nhóm ảnh → khoảng {ns} slide ({split_txt}).\nTiếp tục?'):
            return
        folders = self.opts.get('image_folders') or []
        init_dir = folders[0] if folders else None
        out = filedialog.asksaveasfilename(
            title='Lưu file PPTX', defaultextension='.pptx',
            filetypes=[('PowerPoint', '*.pptx')],
            initialfile='Output_Report.pptx',
            initialdir=init_dir)
        if not out:
            return
        kwargs = {
            'out': out,
            'qa': qa,
            'write_checklist': bool(self.opts.get('write_checklist', True)),
        }
        if self.opts.get('fx', {}).get('quality_check'):
            paths = [p for files in groups.values() for p in files]
            self._run_quality_scan(paths, then_export=True, export_kwargs=kwargs)
            return
        self._start_export(**kwargs)

    def export_stamped(self):
        """Xuất ảnh đã đóng dấu ngày/vị trí — không tạo PPTX."""
        groups_rel = self.imglib.groups()
        if not groups_rel:
            messagebox.showwarning('Xuất ảnh', 'Chưa chọn thư mục ảnh.')
            return
        fx = self.opts.get('fx') or {}
        if (not fx.get('use_timestamp')
                and not (fx.get('watermark') or '').strip()
                and not fx.get('minimap')):
            if not messagebox.askyesno(
                    'Xuất ảnh',
                    'Chưa bật đóng dấu ngày/vị trí (tab Dấu / GPS).\n'
                    'Ảnh xuất ra sẽ giống bản gốc. Tiếp tục?'):
                return
        qa = QA.check(groups_rel, self._by_code, self._merged)
        if self.opts.get('qa_gate', True) and qa.needs_gate():
            self._show_qa_gate(qa, lambda: self._export_stamped_after_qa(qa, confirm=False))
            return
        self._export_stamped_after_qa(qa)

    def _export_stamped_after_qa(self, qa, confirm=True):
        groups_rel = self.imglib.groups()
        n_img = sum(len(v) for v in groups_rel.values())
        n_grp = len(groups_rel)
        if confirm and not messagebox.askyesno(
                'Xuất ảnh đóng dấu',
                f'{n_img} ảnh / {n_grp} mã → lưu vào thư mục bạn chọn.\n'
                'Không tạo file PPTX. Tiếp tục?'):
            return
        folders = self.opts.get('image_folders') or []
        init_dir = folders[0] if folders else None
        out_dir = filedialog.askdirectory(
            title='Chọn thư mục lưu ảnh đóng dấu',
            initialdir=init_dir)
        if not out_dir:
            return
        self._start_stamp_export(out_dir, qa=qa)

    def _start_stamp_export(self, out_dir, qa=None):
        groups = OrderedDict(
            (code, list(files))
            for code, files in self._ordered_groups().items())
        job = exporter.ExportJob(
            layout={},
            groups=groups,
            excel_by_code=dict(self._by_code),
            excel_rows=list(self.excel.rows),
            avatar_map={},
            out_path=out_dir,
            fx=copy.deepcopy(self.opts.get('fx') or {}),
            sort_mode=self.opts.get('sort_mode', 'city'),
            city_filter=self.opts.get('city_filter', ALL_CITIES),
            district_filter=self.opts.get('district_filter', ALL_DISTRICTS),
        )

        cancel = threading.Event()
        pop = ctk.CTkToplevel(self)
        pop.title('Đang đóng dấu ảnh…')
        pop.geometry('420x150')
        pop.attributes('-topmost', True)
        pop.resizable(False, False)
        pop.protocol('WM_DELETE_WINDOW', cancel.set)
        ctk.CTkLabel(pop, text='Đang in ngày / vị trí lên ảnh…',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(pady=(16, 6))
        pb = ctk.CTkProgressBar(pop, height=10, progress_color=ACCENT)
        pb.pack(fill='x', padx=28)
        pb.set(0)
        lbl = ctk.CTkLabel(pop, text='0%', text_color=MUTED)
        lbl.pack(pady=4)
        ctk.CTkButton(pop, text='Huỷ', width=90, fg_color='#ef4444',
                      hover_color='#b91c1c',
                      command=cancel.set).pack(pady=(2, 10))

        def on_prog(done, total):
            frac = done / max(1, total)
            self.after(0, lambda: (pb.set(frac),
                                   lbl.configure(text=f'{int(frac * 100)}%  '
                                                      f'({done}/{total} ảnh)')))

        def on_log(msg):
            self.after(0, lambda m=msg: self.log(m))

        def work():
            rep = exporter.export_stamped_images(job, out_dir, on_prog, on_log, cancel)
            self.after(0, lambda: self._export_done(
                pop, rep, kind='stamp', out_dir=out_dir))
        threading.Thread(target=work, daemon=True).start()

    def _show_qa_gate(self, qa, on_continue):
        win = ctk.CTkToplevel(self)
        win.title('Đối chiếu ảnh với FILE TỔNG / MASTER')
        win.geometry('720x560')
        win.attributes('-topmost', True)
        try:
            win.grab_set()
        except Exception:
            pass
        n_miss, n_un, n_fz = len(qa.missing), len(qa.unmatched), len(qa.fuzzy)
        title = []
        if n_miss:
            title.append(f'{n_miss} điểm list chưa có ảnh')
        if n_un:
            title.append(f'{n_un} nhóm ảnh lệch mã list')
        if n_fz:
            title.append(f'{n_fz} khớp gần đúng')
        ctk.CTkLabel(
            win, text='  ·  '.join(title) or 'Cần kiểm tra trước khi xuất',
            font=ctk.CTkFont(size=15, weight='bold'),
            wraplength=680, justify='left', anchor='w'
        ).pack(fill='x', padx=16, pady=(14, 4))
        ctk.CTkLabel(
            win, text=self._current_list_caption(),
            text_color=MUTED, font=ctk.CTkFont(size=12),
            wraplength=680, justify='left', anchor='w'
        ).pack(fill='x', padx=16, pady=(0, 2))
        ctk.CTkLabel(
            win, text=qa.scope_note + '  ·  PPTX khách hàng không ghi dòng này.',
            text_color=MUTED, font=ctk.CTkFont(size=12),
            wraplength=680, justify='left', anchor='w'
        ).pack(fill='x', padx=16, pady=(0, 8))
        tb = ctk.CTkTextbox(win, font=ctk.CTkFont(family='Consolas', size=12))
        tb.pack(fill='both', expand=True, padx=16, pady=(0, 8))
        lines = []
        if qa.missing:
            lines += [f'— LIST CHƯA CÓ ẢNH ({n_miss})']
            for it in qa.missing[:80]:
                lines.append(f'  {it.code:<16}  {it.name}')
            if n_miss > 80:
                lines.append(f'  … còn {n_miss - 80} điểm')
            lines.append('')
        if qa.unmatched:
            lines += [f'— TÊN FILE KHÔNG KHỚP MÃ TRÊN LIST ĐANG MỞ ({n_un})']
            for it in qa.unmatched[:80]:
                lines.append(f'  {it.code:<16}  {it.photos} ảnh')
            if n_un > 80:
                lines.append(f'  … còn {n_un - 80} nhóm')
            lines.append('')
        if qa.fuzzy:
            lines += [f'— KHỚP GẦN ĐÚNG VỚI LIST ĐANG MỞ — kiểm tra tên file ({n_fz})']
            for it in qa.fuzzy:
                lines.append(f'  {it.code:<16}  ≈  {it.excel_code}  {it.name}')
            lines.append('')
        tb.insert('1.0', '\n'.join(lines) or 'Không có mục nào.')
        tb.configure(state='disabled')

        bar = ctk.CTkFrame(win, fg_color='transparent')
        bar.pack(fill='x', padx=16, pady=(0, 14))
        sw = ctk.CTkSwitch(bar, text='Xuất checklist.xlsx cùng PPTX',
                           progress_color=ACCENT)

        def on_sw():
            self.opts['write_checklist'] = bool(sw.get())
            try:
                self.sw_checklist.select() if sw.get() else self.sw_checklist.deselect()
            except Exception:
                pass
            self._schedule_save()

        sw.configure(command=on_sw)
        if self.opts.get('write_checklist', True):
            sw.select()
        sw.pack(side='left')

        def go():
            try:
                win.grab_release()
                win.destroy()
            except Exception:
                pass
            on_continue()

        def stop():
            try:
                win.grab_release()
                win.destroy()
            except Exception:
                pass

        def fix_codes():
            stop()
            self._show_fix_codes_dialog()

        ctk.CTkButton(bar, text='Huỷ', width=90, fg_color=INPUT,
                      hover_color=BORDER, text_color=TEXT,
                      command=stop).pack(side='right', padx=(8, 0))
        ctk.CTkButton(
            bar, text='Xuất anyway', width=130,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, command=go
        ).pack(side='right')
        if qa.unmatched or qa.fuzzy:
            ctk.CTkButton(
                bar, text='Sửa mã ảnh / đổi tên file…', height=32,
                fg_color=INPUT, hover_color=BORDER, text_color=TEXT,
                command=fix_codes
            ).pack(side='left', padx=(12, 0))

    def _start_export(self, out, qa=None, write_checklist=None):
        groups = OrderedDict(
            (code, list(files))
            for code, files in self._ordered_groups().items())
        avatar_map = {}
        for code in groups:
            p = self._resolve_avatar(code)
            if p:
                avatar_map[code.upper()] = p
        vis = dict(self.opts['visible'])
        job = exporter.ExportJob(
            layout=copy.deepcopy(self.layout),
            groups=groups,
            excel_by_code=dict(self._by_code),
            excel_rows=list(self.excel.rows),
            avatar_map=avatar_map,
            out_path=out,
            n_per_slide=self.opts['n_per_slide'],
            slides_per_file=G.parse_slides_per_file(
                self.opts.get('slides_per_file', 0)),
            img_ar=AR_CHOICES.get(self.opts['ar_label'], 4 / 3),
            bg_image=(self.opts['bg_path']
                      if self.opts['bg_mode'] == 'image' else None),
            channel_enabled=self.opts['channel_enabled'],
            channel_template=self.opts['channel_template'],
            visible=vis,
            fx=copy.deepcopy(self.opts.get('fx') or {}),
            slide_style=self.opts.get('slide_style') or 'report',
            pad_blank_slides=bool(self.opts.get('pad_blank_slides', False)),
            sort_mode=self.opts.get('sort_mode', 'city'),
            city_filter=self.opts.get('city_filter', ALL_CITIES),
            district_filter=self.opts.get('district_filter', ALL_DISTRICTS),
            split_export_by=self.opts.get('split_export_by', 'none'),
            export_pdf=bool(self.opts.get('export_pdf', False)),
        )

        cancel = threading.Event()
        pop = ctk.CTkToplevel(self)
        pop.title('Đang xuất PPTX…')
        pop.geometry('420x150')
        pop.attributes('-topmost', True)
        pop.resizable(False, False)
        pop.protocol('WM_DELETE_WINDOW', cancel.set)
        ctk.CTkLabel(pop, text='Đang dựng slide…',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(pady=(16, 6))
        pb = ctk.CTkProgressBar(pop, height=10, progress_color=ACCENT)
        pb.pack(fill='x', padx=28)
        pb.set(0)
        lbl = ctk.CTkLabel(pop, text='0%', text_color=MUTED)
        lbl.pack(pady=4)
        ctk.CTkButton(pop, text='Huỷ', width=90, fg_color='#ef4444',
                      hover_color='#b91c1c',
                      command=cancel.set).pack(pady=(2, 10))

        def on_prog(done, total):
            frac = done / max(1, total)
            self.after(0, lambda: (pb.set(frac),
                                   lbl.configure(text=f'{int(frac * 100)}%  '
                                                      f'({done}/{total} slide)')))

        def on_log(msg):
            self.after(0, lambda m=msg: self.log(m))

        do_cl = bool(self.opts.get('write_checklist', True)
                     if write_checklist is None else write_checklist)
        do_cl = do_cl and qa is not None and qa.has_excel
        qa_snap = qa
        out_path = out

        def work():
            rep = exporter.export_pptx(job, on_prog, on_log, cancel)
            cl, cl_err = '', ''
            if do_cl and not getattr(rep, 'cancelled', False):
                try:
                    cl = QA.write_checklist(QA.checklist_path(out_path), qa_snap)
                    self.after(0, lambda p=cl: self.log(f'Checklist: {p}'))
                except Exception as e:
                    cl_err = str(e)
            self.after(0, lambda: self._export_done(pop, rep, checklist=cl,
                                                    cl_err=cl_err))
        threading.Thread(target=work, daemon=True).start()

    def _export_done(self, pop, rep, checklist='', cl_err='', kind='pptx',
                     out_dir=''):
        try:
            pop.destroy()
        except Exception:
            pass
        if rep.cancelled:
            self.log('Đã huỷ xuất.')
            title = 'Xuất ảnh' if kind == 'stamp' else 'Xuất PPTX'
            messagebox.showinfo(title, 'Đã huỷ — giữ phần đã lưu (nếu có).')
            return
        if kind == 'stamp':
            self.log(f'Xong: {rep.slides} ảnh trong {rep.seconds:.1f}s.')
        else:
            self.log(f'Xong: {rep.slides} slide / {len(rep.files)} file '
                     f'trong {rep.seconds:.1f}s.')
        if self.opts['open_after_export'] and rep.files:
            try:
                folder = out_dir or os.path.dirname(rep.files[0])
                os.startfile(folder)
            except Exception:
                pass
        win = ctk.CTkToplevel(self)
        win.title('Kết quả xuất ảnh' if kind == 'stamp' else 'Kết quả xuất')
        win.geometry('560x440')
        win.attributes('-topmost', True)
        tb = ctk.CTkTextbox(win, font=ctk.CTkFont(family='Consolas', size=12))
        tb.pack(fill='both', expand=True, padx=10, pady=10)
        if kind == 'stamp':
            folder = out_dir or (os.path.dirname(rep.files[0]) if rep.files else '')
            lines = [f'{rep.slides} ảnh — {rep.groups} mã — '
                     f'{rep.seconds:.1f} giây', '', 'Thư mục:', f'   {folder}']
            if rep.files:
                lines += ['', 'Ví dụ:']
                lines += [f'   {f}' for f in rep.files[:12]]
                if len(rep.files) > 12:
                    lines.append(f'   … còn {len(rep.files) - 12} ảnh')
        else:
            lines = [f'{rep.slides} slide — {rep.groups} nhóm — '
                     f'{rep.seconds:.1f} giây', '', 'File:']
            lines += [f'   {f}' for f in rep.files]
            if checklist:
                lines += ['', 'Checklist (nội bộ):', f'   {checklist}']
            if cl_err:
                lines += ['', f'Không ghi được checklist: {cl_err}']
        if rep.fuzzy:
            lines += ['', f'{len(rep.fuzzy)} mã khớp GẦN ĐÚNG:']
            lines += [f'   {a}  ≈  {b}' for a, b in rep.fuzzy]
        if rep.unmatched:
            lines += ['', f'{len(rep.unmatched)} mã KHÔNG có trong Excel:']
            lines += [f'   {c}' for c in rep.unmatched]
        if kind != 'stamp' and rep.missing_avatars:
            lines += ['', f'{len(rep.missing_avatars)} mã thiếu avatar:']
            lines += [f'   {c}' for c in rep.missing_avatars]
        tb.insert('1.0', '\n'.join(lines))
        tb.configure(state='disabled')

    # ════════════════════════ TIỆN ÍCH ════════════════════════
    def log(self, msg):
        try:
            self.lbl_status.configure(text=str(msg))
        except Exception:
            pass
        try:
            print(msg)
        except UnicodeEncodeError:
            try:
                print(str(msg).encode('ascii', 'replace').decode('ascii'))
            except Exception:
                pass

    def _on_app_mode(self, v):
        mode = 'cloud' if 'Đám mây' in (v or '') else 'local'
        self.opts['app_mode'] = mode
        if mode == 'cloud':
            cache = avatar_cache_dir()
            self.opts['avatar_folder'] = cache
            self.avatars.set_folder(cache)
            if hasattr(self, 'lbl_avatars'):
                self.lbl_avatars.configure(text='☁ Cache Cloud: ' + cache)
        self._apply_mode_ui()
        self._schedule_save()
        if mode == 'cloud' and self.cloud.logged_in():
            self._sync_all_cloud(silent=True)

    def _apply_mode_ui(self):
        cloud = self.opts.get('app_mode') == 'cloud'
        try:
            anchor = getattr(self, 'seg_mode', None)
            if hasattr(self, 'frm_cloud_block') and anchor is not None:
                if cloud:
                    self.frm_cloud_block.pack(fill='x', after=anchor)
                else:
                    self.frm_cloud_block.pack_forget()
            if hasattr(self, 'frm_local_block') and anchor is not None:
                if cloud:
                    self.frm_local_block.pack_forget()
                else:
                    self.frm_local_block.pack(fill='x', after=anchor)
        except Exception:
            pass
        try:
            if hasattr(self, 'btn_sync_cloud'):
                self.btn_sync_cloud.configure(
                    text='Tải FILE TỔNG + Avatar từ Cloud' if cloud
                    else 'Đồng bộ Cloud (tuỳ chọn)')
            if hasattr(self, 'btn_excel'):
                self.btn_excel.configure(
                    state='disabled' if cloud else 'normal',
                    text=('FILE TỔNG lấy từ Cloud — bấm Tải ở trên'
                          if cloud else 'Chọn FILE TỔNG trên máy…'))
            if hasattr(self, 'btn_avatars'):
                self.btn_avatars.configure(
                    state='disabled' if cloud else 'normal',
                    text='Avatar lấy từ Cloud (bước 1)' if cloud
                    else 'Chọn thư mục avatar…')
        except Exception:
            pass

    def _show_account_menu(self):
        menu = tk.Menu(self, tearoff=0, bg='#1e1e22', fg='#e8e8ec',
                       activebackground=ACCENT, activeforeground='#ffffff',
                       relief='flat', font=('Segoe UI', 10))
        who = self.cloud.email or ''
        if who:
            menu.add_command(label='✉  ' + who, state='disabled')
            role = {'admin': 'Quản trị', 'super_admin': 'Super Admin'}.get(
                self.cloud.role, 'Nhân viên')
            menu.add_command(label='   ' + role, state='disabled')
            menu.add_separator()
        if self.cloud.is_admin():
            menu.add_command(label='👥  Quản lý tài khoản',
                             command=self._show_account_manager)
        menu.add_command(label='⚙  Cài đặt',
                         command=self._show_settings)
        menu.add_separator()
        menu.add_command(label='⏻  Đăng xuất', command=self.logout)
        try:
            menu.tk_popup(self.btn_account.winfo_rootx(),
                          self.btn_account.winfo_rooty() + 32)
        finally:
            menu.grab_release()

    def _show_settings(self):
        win = ctk.CTkToplevel(self)
        win.title('Cài đặt')
        win.geometry('580x640')
        win.attributes('-topmost', True)
        try:
            win.transient(self)
        except Exception:
            pass
        sc = ctk.CTkScrollableFrame(win, fg_color='transparent')
        sc.pack(fill='both', expand=True, padx=12, pady=12)

        self._build_settings_source(sc)

        ctk.CTkLabel(
            sc, text='Cloud Supabase (nâng cao)',
            font=ctk.CTkFont(size=14, weight='bold')
        ).pack(anchor='w', pady=(12, 4))
        ctk.CTkLabel(sc, text='URL', text_color=MUTED).pack(anchor='w')
        ent_u = ctk.CTkEntry(sc, height=32)
        ent_u.pack(fill='x', pady=(0, 8))
        ent_u.insert(0, self.cloud.url)
        ctk.CTkLabel(sc, text='Anon key', text_color=MUTED).pack(anchor='w')
        ent_k = ctk.CTkEntry(sc, height=32)
        ent_k.pack(fill='x', pady=(0, 8))
        ent_k.insert(0, self.cloud.key)

        def _save():
            self.cloud.url = ent_u.get().strip()
            self.cloud.key = ent_k.get().strip()
            self.cloud._client = None
            self._save_config()
            win.destroy()
            messagebox.showinfo('Cài đặt', 'Đã lưu. Đăng nhập lại để áp dụng URL/Key mới.')
        bar = ctk.CTkFrame(win, fg_color='transparent')
        bar.pack(fill='x', padx=16, pady=(0, 12))
        ctk.CTkButton(bar, text='Lưu URL/Key', fg_color=ACCENT,
                      command=_save).pack(side='right')
        ctk.CTkButton(bar, text='Đóng', width=90, fg_color=CARD,
                      hover_color=BORDER, text_color=TEXT,
                      command=win.destroy).pack(side='right', padx=8)

    def _show_cloud_settings(self):
        self._show_settings()

    def _excel_prog(self, frac, text=''):
        try:
            if text:
                self._show_sync_progress(True)
            self.excel_progress.set(max(0.0, min(1.0, float(frac))))
            self.lbl_excel_prog.configure(text=text)
        except Exception:
            pass

    def _avatar_prog(self, cur, total, text=''):
        try:
            if text:
                self._show_sync_progress(True)
            self.avatar_progress.set(cur / max(1, total))
            self.lbl_avatar_prog.configure(text=text)
        except Exception:
            pass

    def _sync_all_cloud(self, silent=True):
        if not self.cloud.logged_in():
            if not silent:
                messagebox.showwarning('Cloud', 'Cần đăng nhập tài khoản công ty.')
            return
        self._show_sync_progress(True)

        def work():
            def ep(frac, text):
                self.after(0, lambda: self._excel_prog(frac, text))
            ex = self.cloud.sync_excel(progress=ep, force=not silent)

            def ap(cur, total, text):
                self.after(0, lambda: self._avatar_prog(cur, total, text))
            av = self.cloud.sync_avatars(progress=ap)
            self.after(0, lambda: self._after_cloud_sync(ex, av, silent))
        threading.Thread(target=work, daemon=True).start()

    def _after_cloud_sync(self, ex, av, silent):
        if ex.get('ok') and ex.get('path'):
            self.opts['excel_path'] = ex['path']
            self._load_excel_async(ex['path'])
            try:
                self.lbl_excel.configure(
                    text='☁ ' + (ex.get('file_name') or os.path.basename(ex['path'])))
            except Exception:
                pass
        self._update_data_brief()
        try:
            self.lbl_excel_sync.configure(text='📊 ' + self.cloud.excel_status_text())
        except Exception:
            pass
        if av.get('ok') and av.get('path'):
            self.opts['avatar_folder'] = av['path']
            self.avatars.set_folder(av['path'])
            try:
                self.lbl_avatars.configure(text='☁ Cache Cloud: ' + av['path'])
                self.lbl_avatar_sync.configure(text='👤 ' + self.cloud.avatar_status_text())
            except Exception:
                pass
            self._invalidate()
        msg = ' · '.join(x.get('msg', '') for x in (ex, av) if x)
        self.log('☁ ' + msg)
        if not silent:
            if ex.get('ok') and av.get('ok'):
                messagebox.showinfo('Đám mây', msg)
            else:
                messagebox.showwarning('Đám mây', msg or 'Đồng bộ chưa xong.')
        self.after(3500, lambda: (
            self._excel_prog(0, ''), self._avatar_prog(0, 1, ''),
            self._show_sync_progress(False)))

    def _fetch_bg_list(self):
        if not self.cloud.logged_in():
            try:
                self.lbl_bg_sync.configure(text='Đăng nhập để xem mẫu nền Cloud.')
            except Exception:
                pass
            return
        try:
            self.lbl_bg_sync.configure(text='⟳ Đang tải danh sách ảnh nền…')
        except Exception:
            pass

        def work():
            rows = self.cloud.list_backgrounds()
            rows.sort(key=lambda r: (r.get('file_name') or '').lower())
            self.after(0, lambda: self._fill_bg_gallery(rows))
        threading.Thread(target=work, daemon=True).start()

    def _fill_bg_gallery(self, rows):
        self._bg_manifest = rows
        mine = f'users/{self.cloud.user_id}/' if self.cloud.user_id else 'users/'
        n_mine = sum(1 for r in rows if (r.get('storage_path') or '').startswith(mine))
        try:
            self.lbl_bg_sync.configure(
                text=f'{len(rows)} mẫu · {n_mine} của bạn'
                if rows else 'Chưa có mẫu — bấm Tải lên Cloud')
        except Exception:
            pass
        gal = getattr(self, 'bg_gallery', None)
        if gal is None:
            return
        for w in gal.winfo_children():
            w.destroy()
        self._bg_photos = []
        self._bg_thumb_btns = {}
        grid = ctk.CTkFrame(gal, fg_color='transparent')
        grid.pack(fill='x')
        tw, th = 352, 148
        for r in rows:
            cell = ctk.CTkFrame(grid, fg_color=INPUT, corner_radius=10)
            cell.pack(fill='x', padx=4, pady=6)
            ph = ctk.CTkLabel(
                cell, text='…', width=tw, height=th, fg_color='#1a1a1e',
                text_color=MUTED, corner_radius=8)
            ph.pack(padx=6, pady=(6, 0), fill='x')
            sp = r.get('storage_path') or ''
            badge = 'của tôi' if sp.startswith(mine) else 'chung'
            cap = r.get('file_name') or sp
            if len(cap) > 42:
                cap = cap[:40] + '…'
            ctk.CTkLabel(
                cell, text=f'{cap}  ·  {badge}',
                font=ctk.CTkFont(size=11),
                text_color=TEXT if badge == 'của tôi' else MUTED,
                anchor='w'
            ).pack(fill='x', padx=8, pady=(4, 8))
            cell.bind('<Button-1>', lambda e, row=r: self._preview_bg_cloud(row))
            ph.bind('<Button-1>', lambda e, row=r: self._preview_bg_cloud(row))
            self._bg_thumb_btns[sp] = ph
        if rows:
            threading.Thread(target=self._load_bg_thumbs, args=(rows,), daemon=True).start()

    def _load_bg_thumbs(self, rows):
        from concurrent.futures import ThreadPoolExecutor
        from PIL import Image

        def one(r):
            return r, self.cloud.ensure_bg_thumb(r['storage_path'], r.get('updated_at'))

        with ThreadPoolExecutor(max_workers=4) as ex:
            for r, res in ex.map(one, rows):
                if not res.get('ok'):
                    continue
                try:
                    im = Image.open(res['path']).convert('RGB')
                except Exception:
                    continue
                self.after(0, lambda row=r, img=im: self._set_bg_thumb(row, img))

    def _set_bg_thumb(self, row, pil_img):
        sp = row.get('storage_path')
        lbl = (self._bg_thumb_btns or {}).get(sp)
        if lbl is None:
            return
        try:
            cimg = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(352, 148))
            self._bg_photos.append(cimg)
            lbl.configure(image=cimg, text='')
            lbl.bind('<Button-1>', lambda e, row=row: self._preview_bg_cloud(row))
        except Exception:
            pass

    def _preview_bg_cloud(self, row):
        """Cửa sổ xem trước lớn — mới Áp dụng khi người dùng xác nhận."""
        win = ctk.CTkToplevel(self)
        win.title(row.get('file_name') or 'Xem trước nền')
        win.geometry('780x500')
        win.attributes('-topmost', True)
        ctk.CTkLabel(win, text='Đang tải mẫu…', text_color=MUTED).pack(pady=20)
        status = ctk.CTkLabel(win, text='', text_color=MUTED)
        status.pack()

        def work():
            res = self.cloud.download_background(row['storage_path'], row.get('updated_at'))
            self.after(0, lambda: self._show_bg_preview(win, row, res))
        threading.Thread(target=work, daemon=True).start()

    def _show_bg_preview(self, win, row, res):
        for w in win.winfo_children():
            w.destroy()
        if not res.get('ok'):
            ctk.CTkLabel(win, text=res.get('msg') or 'Không tải được.').pack(pady=30)
            return
        from PIL import Image
        try:
            im = Image.open(res['path']).convert('RGB')
            im.thumbnail((740, 390), Image.Resampling.LANCZOS)
            photo = ctk.CTkImage(light_image=im, dark_image=im, size=im.size)
            self._bg_photos.append(photo)
            ctk.CTkLabel(win, image=photo, text='').pack(pady=8)
        except Exception as e:
            ctk.CTkLabel(win, text=str(e)).pack()
        ctk.CTkLabel(win, text=row.get('file_name') or '', text_color=MUTED).pack()
        rowb = ctk.CTkFrame(win, fg_color='transparent')
        rowb.pack(pady=10)
        ctk.CTkButton(rowb, text='Áp dụng nền này', width=160, height=36, fg_color=ACCENT,
                      command=lambda: (self._bg_cloud_done(res, row.get('file_name')),
                                       win.destroy())).pack(side='left', padx=6)
        ctk.CTkButton(rowb, text='Đóng', width=90, height=36, fg_color=INPUT,
                      hover_color=BORDER, text_color=TEXT,
                      command=win.destroy).pack(side='left', padx=6)

    def _apply_bg_cloud(self, row):
        self._preview_bg_cloud(row)

    def _upload_my_bg(self):
        if not self.cloud.logged_in():
            messagebox.showwarning('Cloud', 'Cần đăng nhập để tải nền lên.')
            return
        paths = filedialog.askopenfilenames(
            title='Chọn ảnh nền của bạn (PNG/JPG)',
            filetypes=[('Ảnh', '*.png *.jpg *.jpeg *.webp')])
        if not paths:
            return

        def work():
            res = self.cloud.upload_background_files(list(paths), shared=False)
            self.after(0, lambda: self._after_my_bg(res, list(paths)))
        threading.Thread(target=work, daemon=True).start()

    def _after_my_bg(self, res, paths):
        self.log('☁ ' + res.get('msg', ''))
        if res.get('ok'):
            messagebox.showinfo('Nền Cloud', res.get('msg', '') +
                                '\nMẫu của bạn sẽ hiện trong gallery (nhãn «của tôi»).')
            if paths:
                self.opts['bg_path'] = paths[0]
                self.opts['bg_mode'] = 'image'
                try:
                    self.seg_bg.set('Ảnh nền')
                    self.lbl_bg.configure(text=paths[0])
                except Exception:
                    pass
                self._invalidate()
            self._fetch_bg_list()
        else:
            messagebox.showerror('Nền Cloud', res.get('msg', 'Không đẩy được.'))

    def _bg_cloud_done(self, res, name):
        if not res.get('ok'):
            messagebox.showerror('Ảnh nền', res.get('msg') or 'Không tải được mẫu.')
            return
        self.opts['bg_path'] = res['path']
        self.opts['bg_mode'] = 'image'
        try:
            self.seg_bg.set('Ảnh nền')
            self.lbl_bg.configure(text=res['path'])
            self.lbl_bg_sync.configure(text=f'✓ Đã áp dụng {name}')
        except Exception:
            pass
        self._invalidate()
        self.log(f'☁ Áp dụng nền: {name}')

    def _upload_excel_dialog(self):
        self.pick_master_excel(purpose='push')

    def _do_upload_excel(self, path, target, win, then_load=False):
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass

        def work():
            if target == 'all':
                profiles = self.cloud.fetch_profiles()
                res = self.cloud.upload_excel_broadcast(path, profiles)
            else:
                res = self.cloud.upload_excel(path)

            def done():
                self.log('☁ ' + res.get('msg', ''))
                if res.get('ok'):
                    messagebox.showinfo('FILE TỔNG / MASTER',
                                        res.get('msg', 'Đã đẩy FILE TỔNG lên Cloud.'))
                    if then_load:
                        load_path = res.get('path') or path
                        self._load_excel_async(load_path)
                else:
                    messagebox.showerror('FILE TỔNG / MASTER',
                                         res.get('msg', 'Không đẩy được FILE TỔNG.'))
            self.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _upload_avatars_dialog(self):
        if not self.cloud.is_admin():
            messagebox.showwarning('Cloud', 'Chỉ quản trị mới được đẩy avatar.')
            return
        folder = filedialog.askdirectory(title='Chọn thư mục Avatar để đẩy lên Cloud')
        if not folder:
            return
        if not messagebox.askyesno('Cloud', 'Đẩy toàn bộ ảnh trong thư mục lên kho avatar?'):
            return

        def work():
            res = self.cloud.upload_avatars(folder)
            self.after(0, lambda: (
                self.log('☁ ' + res.get('msg', '')),
                messagebox.showinfo('Cloud', res.get('msg', ''))))
        threading.Thread(target=work, daemon=True).start()

    def _upload_bg_dialog(self, shared=False):
        folder = filedialog.askdirectory(title='Chọn thư mục ảnh nền để đẩy lên Cloud')
        if not folder:
            return

        def work():
            res = self.cloud.upload_backgrounds(folder, shared=shared)
            self.after(0, lambda: (
                self.log('☁ ' + res.get('msg', '')),
                messagebox.showinfo('Cloud', res.get('msg', '')),
                self._fetch_bg_list()))
        threading.Thread(target=work, daemon=True).start()

    def _show_account_manager(self):
        if not self.cloud.is_admin():
            messagebox.showwarning('Cloud', 'Chỉ quản trị mới mở được mục này.')
            return
        win = ctk.CTkToplevel(self)
        win.title('Quản lý tài khoản')
        win.geometry('560x520')
        win.attributes('-topmost', True)
        ctk.CTkLabel(win, text='👥  Quản lý tài khoản',
                     font=ctk.CTkFont(size=17, weight='bold')).pack(anchor='w', padx=18, pady=(14, 2))
        ctk.CTkLabel(win, text='Cấp / thu hồi quyền quản trị. Chỉ admin mới đẩy Excel/Avatar lên Cloud.',
                     text_color=MUTED, font=ctk.CTkFont(size=11), wraplength=500,
                     justify='left').pack(anchor='w', padx=18, pady=(0, 8))
        body = ctk.CTkScrollableFrame(win, fg_color=INPUT, corner_radius=8)
        body.pack(fill='both', expand=True, padx=18, pady=(0, 8))
        status = ctk.CTkLabel(win, text='', text_color=MUTED, font=ctk.CTkFont(size=11))
        status.pack(anchor='w', padx=18, pady=(0, 10))

        def _set(uid, name, make):
            res = self.cloud.set_admin(uid, name, make)
            if not res.get('ok'):
                messagebox.showerror('Cloud', res.get('msg', ''))
                return
            _render()

        def _render():
            for w in body.winfo_children():
                w.destroy()
            status.configure(text='⏳ Đang tải…')

            def work():
                rows = self.cloud.fetch_profiles()
                self.after(0, lambda: _paint(rows))
            threading.Thread(target=work, daemon=True).start()

        def _paint(rows):
            if not rows:
                status.configure(text='Không đọc được danh sách tài khoản.')
                return
            seen = set()
            for r in rows:
                uid, name, role = r.get('id'), r.get('name') or '(chưa đặt tên)', r.get('role')
                if uid in seen:
                    continue
                seen.add(uid)
                is_ad = role in ('admin', 'super_admin')
                card = ctk.CTkFrame(body, fg_color=CARD, corner_radius=8, height=52)
                card.pack(fill='x', padx=4, pady=3)
                card.pack_propagate(False)
                ctk.CTkLabel(card, text='🔑' if is_ad else '👤',
                             font=ctk.CTkFont(size=16)).pack(side='left', padx=(12, 8))
                info = ctk.CTkFrame(card, fg_color='transparent')
                info.pack(side='left', fill='x', expand=True)
                me = '   (bạn)' if uid == self.cloud.user_id else ''
                ctk.CTkLabel(info, text=name + me,
                             font=ctk.CTkFont(size=12, weight='bold'),
                             anchor='w').pack(anchor='w')
                ctk.CTkLabel(info, text={'super_admin': 'Super Admin', 'admin': 'Quản trị viên'}
                             .get(role, 'Nhân viên'), font=ctk.CTkFont(size=10),
                             text_color=ACCENT if is_ad else MUTED,
                             anchor='w').pack(anchor='w')
                if role == 'super_admin':
                    ctk.CTkLabel(card, text='không đổi được', text_color=MUTED,
                                 font=ctk.CTkFont(size=10)).pack(side='right', padx=12)
                elif is_ad:
                    ctk.CTkButton(card, text='Thu hồi quyền', width=118, height=30,
                                  fg_color=INPUT, text_color='#ef4444',
                                  command=lambda u=uid, n=name: _set(u, n, False)
                                  ).pack(side='right', padx=12)
                else:
                    ctk.CTkButton(card, text='Cấp quyền admin', width=130, height=30,
                                  fg_color=ACCENT,
                                  command=lambda u=uid, n=name: _set(u, n, True)
                                  ).pack(side='right', padx=12)
            n_ad = sum(1 for r in rows if r.get('role') in ('admin', 'super_admin'))
            status.configure(text=f'{len(seen)} tài khoản · {n_ad} quản trị viên')

        _render()


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
