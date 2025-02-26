# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\main.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=['gui', 'label_creator'],  # Include only necessary hidden imports
    hookspath=[],
    runtime_hooks=[],
    excludes=['botocore', 'pyarrow', 'IPython', 'jedi', 'sqlite3', 'PIL', 'jinja2'],  # Add unwanted modules here
    noarchive=False,
    optimize=2,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='label_creator_pdf',
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
)
