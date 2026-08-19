# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import customtkinter

ctk_path = os.path.dirname(customtkinter.__file__)

block_cipher = None

# Safe data collection
datas = [(ctk_path, 'customtkinter')]
for folder in ['LOGO', 'Data background', 'Data avatar']:
    if os.path.exists(folder):
        datas.append((folder, folder))

# Base hidden imports
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
    hiddenimports.extend(['keyring.backends.macOS', 'keyring.backends.OS_X'])

a = Analysis(
    ['AutoPPTX_Studio_V2.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.ico'] if os.path.exists('app_icon.ico') else None,
)

if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='AutoPPTX_Studio_V2.app',
        icon='app_icon.ico' if os.path.exists('app_icon.ico') else None,
        bundle_identifier='com.autopptx.studio',
        info_plist={
            'NSHighResolutionCapable': 'True',
        },
    )

