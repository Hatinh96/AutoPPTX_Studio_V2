"""Hằng số & bố cục mặc định dùng chung toàn app."""
import os
import re

# ── Slide 16:9 chuẩn ──
SLIDE_W_IN = 13.33
SLIDE_H_IN = 7.5
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

# Bố cục mặc định (đã căn theo mẫu Golden Asia — giống bản 1)
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
