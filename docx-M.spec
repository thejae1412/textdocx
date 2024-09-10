# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\\JOB\\python gui\\Docx-New\\docx-M.py'],
    pathex=[],
    binaries=[],
    datas=[('D:\\JOB\\python gui\\Docx-New\\img', 'img/')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='docx-M',
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
    icon=['D:\\JOB\\python gui\\Docx-New\\img\\program_ico.ico'],
)
