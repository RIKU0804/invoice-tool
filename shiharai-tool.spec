# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[('template/集計用.xlsx', 'template')],
    hiddenimports=['pdfplumber', 'PIL', 'fitz'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='shiharai-tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # GUIアプリなのでコンソール非表示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,           # アイコンがあれば icon='assets/icon.ico' に変更
)
