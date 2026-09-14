"""Hằng số & bố cục mặc định dùng chung toàn app."""
import os
import re
import sys

# ── Slide 16:9 chuẩn ──
SLIDE_W_IN = 13.33
SLIDE_H_IN = 7.5
# Gợi ý cắt file (preset 200). Xuất thật dùng opts['slides_per_file'] (0 = một file).
MAX_SLIDES_PER_FILE = 200

IMG_EXTS = ('.png', '.jpg', '.jpeg', '.webp')
IMG_GAP_IN = 0.1

# "STORE_1.jpg", "STORE (2).png" → mã "STORE"
CODE_SUFFIX_RE = re.compile(r'(\s*\(\d+\)|_\d+)$')

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".autopptx_studio_v2")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
PRESET_DIR = os.path.join(CONFIG_DIR, "presets")
V1_CREDS_PATH = os.path.join(os.path.expanduser("~"), ".autopptx_studio.json")
KEYRING_SERVICE = "AutoPPTXStudio"


def app_root():
    """Thư mục cạnh .exe (ghi được). Không dùng temp PyInstaller."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def bundle_root():
    """Onefile: sys._MEIPASS (LOGO/ico được giải nén ở đây)."""
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', app_root())
    return app_root()


def find_asset(*rel_paths):
    """Tìm file đóng gói: _MEIPASS trước, rồi cạnh .exe / source."""
    roots = []
    for root in (bundle_root(), app_root()):
        n = os.path.normcase(os.path.abspath(root))
        if n not in roots:
            roots.append(n)
    for root in roots:
        for rel in rel_paths:
            p = os.path.normpath(os.path.join(root, rel))
            if os.path.isfile(p):
                return p
    return ''


LOGIN_LOGOS = (
    os.path.join('LOGO', 'Logo Golden Asia_Standard Horizontal_Name White.png'),
    os.path.join('LOGO', 'Logo Golden Asia_White Horizontal.png'),
    os.path.join('LOGO', 'Logo Golden_Symbol White.png'),
    os.path.join('LOGO', 'Logo Golden Asia_Symbol_Standard.png'),
)
WINDOW_ICONS = (
    'app_icon.icns',
    'app_icon.ico',
    'app_icon_1024.png',
    os.path.join('LOGO', 'Logo Golden Asia_Symbol_Standard.png'),
    os.path.join('LOGO', 'Logo Golden_Symbol Premium.png'),
    os.path.join('LOGO', 'Logo Golden_Symbol Multicolor.png'),
)
STAMP_LOGOS = (
    os.path.join('LOGO', 'Logo Golden Asia_Standard Horizontal.png'),
    os.path.join('LOGO', 'Logo Golden Asia_Symbol_Standard.png'),
    os.path.join('LOGO', 'Logo Golden_Symbol Multicolor.png'),
)

# Project Supabase của hệ thống RP SALES / AutoPPTX (anon key — giống bản 1)
SUPABASE_URL = "https://jrwsbpbtkopuqipojspl.supabase.co"
SUPABASE_ANON_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9."
    "eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Impyd3NicGJ0a29wdXFpcG9qc3BsIiwicm9sZSI6ImFub24i"
    "LCJpYXQiOjE3ODM0MTY4NjksImV4cCI6MjA5ODk5Mjg2OX0."
    "txHyGEWAS7C5wArp4i-qEtr7LUlo7ijVJk2IKSGuQSg"
)

GA_BLUE_DEEP = "#0e56c8"
GA_BLUE_GLOW = "#4aa3f5"

ALL_CHANNELS = "Tất cả kênh"
ALL_CITIES = "Tất cả tỉnh"
ALL_DISTRICTS = "Tất cả quận"

STAMP_PRESETS = {
    'Tùy chỉnh': {},
    'Chỉ ngày': {'use_timestamp': True, 'stamp_location': False,
                 'show_gps': False, 'minimap': False},
    'Ngày + địa điểm': {'use_timestamp': True, 'stamp_location': True,
                        'show_gps': False, 'minimap': False},
    'Đầy đủ GPS': {'use_timestamp': True, 'stamp_location': True,
                   'show_gps': True, 'minimap': True},
}

# Nhãn tỉ lệ ô ảnh → w/h (None = ô lấp đầy vùng đã kéo)
AR_CHOICES = {
    "4:3": 4 / 3,
    "16:9": 16 / 9,
    "3:2": 3 / 2,
    "1:1": 1.0,
    "3:4": 3 / 4,
    "Tự do": None,
}

ELEMENT_NAMES = ("image", "avatar", "title", "info", "channel")
ELEMENT_LABELS = {
    "image": "Vùng ảnh",
    "avatar": "Avatar",
    "title": "Tiêu đề",
    "info": "Thông tin",
    "channel": "Kênh",
}
ELEMENT_COLORS = {
    "image": "#e8590c",
    "avatar": "#e11d48",
    "title": "#1f6feb",
    "info": "#2fa572",
    "channel": "#8a2be2",
}

# Bố cục mặc định — mẫu báo cáo cũ (avatar + info trái, ảnh giữa)
DEFAULT_LAYOUT = {
    'title': {'x': 6.63, 'y': 0.27, 'w': 8.02, 'h': 0.43, 'size': 23,
              'align': 'right', 'max_lines': 2, 'min_size': 12, 'upper': True,
              'full_width_on_slide': True, 'show_counter': False,
              'color': '', 'opacity': 100, 'tracking': 0, 'font': ''},
    'info': {'x': 0.39, 'y': 4.09, 'w': 3.81, 'h': 3.15, 'size': 15,
             'align': 'left', 'max_lines': 2, 'min_size': 8,
             'accent_color': '', 'opacity': 100, 'font': ''},
    'image': {'x': 4.59, 'y': 0.69, 'w': 8.28, 'h': 5.55,
              'fit_mode': 'fill', 'gap': IMG_GAP_IN, 'radius': 0, 'opacity': 100},
    'avatar': {'x': 0.95, 'y': 1.16, 'w': 2.70, 'ar': 4 / 3, 'radius': 0,
               'opacity': 100},
    'channel': {'x': 10.16, 'y': 7.00, 'w': 4.27, 'h': 0.84, 'size': 10,
                'align': 'left', 'max_lines': 2, 'min_size': 8, 'color': '',
                'opacity': 100, 'tracking': 0, 'font': ''},
    'font': {'name': 'Arial', 'bold': True, 'italic': False,
             'color': '#000000', 'path': '', 'index': 0},
}

# Mẫu BD / SALESKIT — tên trường + khung bảng + avatar góc trái; nền user tự up
SALESKIT_LAYOUT = {
    'title': {'x': 4.20, 'y': 0.18, 'w': 8.80, 'h': 0.48, 'size': 22,
              'align': 'right', 'max_lines': 2, 'min_size': 12, 'upper': True,
              'full_width_on_slide': False, 'show_counter': False,
              'color': '#1B2430', 'opacity': 100, 'tracking': 0, 'font': 'Arial'},
    'info': {'x': 0.35, 'y': 5.05, 'w': 12.63, 'h': 2.22, 'size': 11,
             'align': 'left', 'max_lines': 2, 'min_size': 8,
             'accent_color': '#1E6EE8', 'opacity': 100, 'font': 'Arial'},
    'image': {'x': 0.35, 'y': 0.78, 'w': 12.63, 'h': 4.12,
              'fit_mode': 'fill', 'gap': IMG_GAP_IN, 'radius': 0, 'opacity': 100},
    # Cùng cỡ Sales (2.70″ · 4:3); góc trên-trái, bên trái tiêu đề, không đè bảng
    'avatar': {'x': 0.35, 'y': 0.14, 'w': 2.70, 'ar': 4 / 3, 'radius': 0,
               'opacity': 100},
    'channel': {'x': 10.16, 'y': 7.00, 'w': 4.27, 'h': 0.84, 'size': 10,
                'align': 'left', 'max_lines': 2, 'min_size': 8, 'color': '',
                'opacity': 100, 'tracking': 0, 'font': ''},
    'font': {'name': 'Arial', 'bold': True, 'italic': False,
             'color': '#1B2430', 'path': '', 'index': 0},
}

# Mẫu THE NELSON — nền user tự chèn; khung title/ảnh/bảng kéo được
NELSON_LAYOUT = {
    'title': {'x': 2.13, 'y': 0.07, 'w': 5.50, 'h': 0.66, 'size': 26,
              'align': 'left', 'max_lines': 2, 'min_size': 14, 'upper': False,
              'full_width_on_slide': False, 'show_counter': False,
              'color': '#E8B83C', 'opacity': 100, 'tracking': 0, 'font': 'Arial'},
    'info': {'x': 3.84, 'y': 5.25, 'w': 8.72, 'h': 2.37, 'size': 11,
             'align': 'left', 'max_lines': 2, 'min_size': 8,
             'accent_color': '#2C2419', 'opacity': 100, 'font': 'Arial'},
    'image': {'x': 1.51, 'y': 1.29, 'w': 10.23, 'h': 1.07,
              'fit_mode': 'fill', 'gap': IMG_GAP_IN, 'radius': 0, 'opacity': 100},
    'avatar': {'x': 0.35, 'y': 0.14, 'w': 2.70, 'ar': 4 / 3, 'radius': 0,
               'opacity': 100},
    'channel': {'x': 4.55, 'y': 0.74, 'w': 8.50, 'h': 0.27, 'size': 14,
                'align': 'right', 'max_lines': 2, 'min_size': 9, 'color': '#E1E1E1',
                'opacity': 100, 'tracking': 0, 'font': 'Arial'},
    'font': {'name': 'Arial', 'bold': True, 'italic': False,
             'color': '#E8B83C', 'path': '', 'index': 0},
}

NELSON_HEADER_BG = '#2C2419'
NELSON_HEADER_FG = '#FFFFFF'
NELSON_DATA_BG = '#F5F0E8'
NELSON_LINE = '#C4B89A'
NELSON_TEXT = '#1B2430'

BUILTIN_PRESETS = ('Báo cáo', 'SALESKIT', 'THE NELSON')
DEPT_UI = ('Sales', 'BD')


def element_label(name, slide_style='report'):
    """Nhãn phần tử trên UI — NELSON dùng khung channel làm dòng địa chỉ."""
    if slide_style == 'nelson' and name == 'channel':
        return 'Địa chỉ'
    return ELEMENT_LABELS.get(name, name)


def element_labels_for(slide_style='report'):
    return {n: element_label(n, slide_style) for n in ELEMENT_NAMES}


def dept_key(label):
    return 'bd' if str(label or '').strip() == 'BD' else 'sales'


def dept_label(key):
    return 'BD' if key == 'bd' else 'Sales'


def pack_for_dept(dept):
    return builtin_pack('SALESKIT' if dept == 'bd' else 'Báo cáo')


def builtin_pack(name):
    """Gói bố cục có sẵn: layout + kiểu bảng + ẩn/hiện phần tử."""
    if name == 'SALESKIT':
        vis = {n: True for n in ELEMENT_NAMES}
        return {
            'layout': SALESKIT_LAYOUT,
            'slide_style': 'saleskit',
            'visible': vis,
            'channel_enabled': True,
        }
    if name == 'THE NELSON':
        vis = {n: (n != 'avatar') for n in ELEMENT_NAMES}
        return {
            'layout': NELSON_LAYOUT,
            'slide_style': 'nelson',
            'visible': vis,
            'channel_enabled': True,
        }
    vis = {n: True for n in ELEMENT_NAMES}
    return {
        'layout': DEFAULT_LAYOUT,
        'slide_style': 'report',
        'visible': vis,
        'channel_enabled': True,
    }

TABLE_HEADER_BG = "#1E6EE8"
TABLE_HEADER_FG = "#FFFFFF"
TABLE_DATA_BG = "#EDEAF4"
TABLE_SUB_FG = "#1E6EE8"
TABLE_LINE = "#D0D4DE"
TABLE_TEXT = "#1B2430"

# ── Màu giao diện (Golden Asia blue) ──
ACCENT = "#1774d2"
ACCENT_HOVER = "#3595ee"
GUIDE_COLOR = "#2e9bff"
SNAP_PX = 7            # khoảng cách hít (pixel canvas)
UNDO_MAX = 100

BG = ("#f0f0f0", "#0c0c0f")
CARD = ("#ffffff", "#161619")
TEXT = ("#111111", "#e8e8ec")
MUTED = ("#888888", "#8b8b98")
INPUT = ("#f5f5f5", "#1e1e22")
BORDER = ("#e0e0e0", "#2a2a30")
WORKSPACE = ("#e4e4e6", "#0a0a0d")

# Bảng màu kiểu Canva — ô to, đủ tương phản, bấm là tô
COLOR_PALETTE = [
    "#000000", "#1f2937", "#4b5563", "#9ca3af", "#e5e7eb", "#ffffff",
    "#7f1d1d", "#dc2626", "#f87171", "#fecaca",
    "#9a3412", "#ea580c", "#fb923c", "#FF9900",
    "#854d0e", "#ca8a04", "#facc15", "#fef08a",
    "#14532d", "#16a34a", "#4ade80", "#bbf7d0",
    "#0e56c8", "#1774d2", "#3b82f6", "#93c5fd",
    "#6b21a8", "#8a2be2", "#c084fc", "#f0abfc",
    "#0f766e", "#14b8a6", "#5eead4", "#ccfbf1",
]
