# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import customtkinter

ctk_path = os.path.dirname(customtkinter.__file__)

try:
    from make_app_icon import write_icons
    write_icons(os.path.abspath('.'))
except Exception as _e:
    print('[spec] make_app_icon:', _e)

_win_icon = 'app_icon.ico' if os.path.exists('app_icon.ico') else None
_mac_icon = 'app_icon.icns' if os.path.exists('app_icon.icns') else _win_icon

from PyInstaller.utils.hooks import copy_metadata

block_cipher = None

_rthooks = []
if sys.platform == 'darwin':
    _rthooks.append(os.path.join('hooks', 'rthook_darwin.py'))

# Chỉ đóng gói logo login/icon. KHÔNG nhét Data avatar (hàng trăm MB, tải lúc đồng bộ).
datas = [(ctk_path, 'customtkinter')]
_logo_names = (
    'Logo Golden Asia_Standard Horizontal_Name White.png',
    'Logo Golden Asia_White Horizontal.png',
    'Logo Golden_Symbol White.png',
    'Logo Golden Asia_Symbol_Standard.png',
    'Logo Golden Asia_Standard Horizontal.png',
    'Logo Golden_Symbol Multicolor.png',
)
for name in _logo_names:
    src = os.path.join('LOGO', name)
    if os.path.isfile(src):
        datas.append((src, 'LOGO'))
for extra in ('app_icon.ico', 'app_icon.icns', 'app_icon_1024.png'):
    if os.path.isfile(extra):
        datas.append((extra, '.'))

hiddenimports = [
    'customtkinter',
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',
    'openpyxl',
    'supabase',
    'keyring',
    'keyring.backends',
    'autopptx2',
    'autopptx2.app',
    'autopptx2.cloud',
    'autopptx2.constants',
    'autopptx2.datasource',
    'autopptx2.editor',
    'autopptx2.effects',
    'autopptx2.exporter',
    'autopptx2.fonts',
    'autopptx2.geometry',
    'autopptx2.login',
    'autopptx2.qa',
    'autopptx2.style',
    'autopptx2.thumbs',
]

if sys.platform == 'win32':
    hiddenimports.extend(['keyring.backends.Windows', 'windnd'])
elif sys.platform == 'darwin':
    # keyring >= 25 đã bỏ 'keyring.backends.OS_X'. Giữ tên cũ trong hiddenimports
    # làm PyInstaller báo lỗi và KHÔNG nhét backend mật khẩu vào bản Mac ->
    # tick "Nhớ mật khẩu" nạp hụt module, màn đăng nhập báo sai "Thiếu package supabase".
    hiddenimports.extend(['keyring.backends.macOS'])

# supabase & nhánh phụ thuộc đọc version qua importlib.metadata -> cần kèm
# thư mục *.dist-info, nếu không bản đóng gói báo "Thiếu package supabase".
# CHỈ copy_metadata. KHÔNG dùng collect_all cho nhánh này: đã thử và nó làm
# hỏng bootstrap của bản onefile (app không khởi động được).
# Chỉ in ASCII — console Windows dùng cp1252, in tiếng Việt ở đây sẽ ném
# UnicodeEncodeError ngay trong except và làm hỏng cả bản build.
_META_PKGS = (
    'supabase', 'supabase_auth', 'supabase_functions', 'gotrue', 'postgrest',
    'realtime', 'storage3', 'supafunc', 'httpx', 'httpcore',
    'websockets', 'pydantic', 'pydantic_core', 'keyring',
)
for _pkg in _META_PKGS:
    try:
        datas += copy_metadata(_pkg)
    except Exception as _e:
        print('[spec] copy_metadata skip %s: %s' % (_pkg, _e))

hiddenimports.extend([
    'supabase', 'supabase.client',
    'postgrest', 'storage3', 'realtime',
    'httpx', 'httpcore', 'h11', 'anyio', 'sniffio',
    'pydantic', 'pydantic_core',
])

excludes = [
    'numpy', 'numpy.tests',
    'pandas', 'matplotlib', 'scipy', 'cv2', 'torch',
    'IPython', 'notebook', 'pytest',
]

a = Analysis(
    ['AutoPPTX_Studio_V2.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=_rthooks,
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AutoPPTX_Studio_V2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=_win_icon,
)

if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='AutoPPTX_Studio_V2.app',
        icon=_mac_icon,
        bundle_identifier='com.autopptx.studio',
        info_plist={
            'NSHighResolutionCapable': 'True',
        },
    )
