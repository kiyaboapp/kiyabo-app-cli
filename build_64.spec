# -*- mode: python ; coding: utf-8 -*-

import os

a = Analysis(
    ['run_cli.py'],
    pathex=[os.path.abspath(".")],  # ← THIS LINE
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='kiyabo64',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon='icon.ico',
    disable_windowed_traceback=False,
)