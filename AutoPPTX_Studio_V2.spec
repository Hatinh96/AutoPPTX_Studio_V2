# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import customtkinter

ctk_path = os.path.dirname(customtkinter.__file__)

block_cipher = None

a = Analysis(
    ['AutoPPTX_Studio_V2.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        (ctk_path, 'customtkinter'),
        ('LOGO', 'LOGO'),
        ('Data background', 'Data background'),
        ('Data avatar', 'Data avatar'),
    ],
    hiddenimports=[
        'customtkinter',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'openpyxl',
        'supabase',
        'keyring',
        'keyring.backends',
        'keyring.backends.Windows',
        'windnd',
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
    ],
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
    icon=['app_icon.ico'],
)
