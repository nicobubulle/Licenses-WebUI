# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
import sys
import os

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('i18n', 'i18n'),
        ('static', 'static'),
    ],
    hiddenimports=['flask', 'pystray', 'PIL'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Get absolute path to icon (use current working directory)
icon_path = os.path.abspath(os.path.join(os.getcwd(), 'static', 'favicon.ico'))
if not os.path.exists(icon_path):
    icon_path = None
    print(f"Warning: Icon not found at {icon_path}, building without icon")
else:
    print(f"Using icon: {icon_path}")

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Licenses_WebUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=icon_path,  # Absolute path to .ico
)