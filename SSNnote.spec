# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files, collect_dynamic_libs

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[] + collect_dynamic_libs('winrt'),
    datas=[
        ('version.json', '.'),
        ('엑스아이콘.png', '.'),
        ('정리 아이콘.png', '.'),
        ('수정 아이콘.png', '.'),
        ('아래방향 아이콘.png', '.'),
        ('윗방향 아이콘.png', '.'),
    ] + collect_data_files('PyQt5') + collect_data_files('winrt'),
    hiddenimports=[
        'PyQt5.sip',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'ctypes',
        'ctypes.wintypes',
        'sqlite3',
        'socket',
        'threading',
        'urllib.request',
        'urllib.error',
        'json',
        'tempfile',
        'subprocess',
        'asyncio',
        'winreg',
        'PIL',
        'PIL.Image',
        'PIL.ImageFilter',
        'PIL.ImageOps',
    ] + collect_submodules('winrt') + collect_submodules('PyQt5'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'django',
        'flask',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SSNnote',
    icon='icon.ico',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='SSNnote'
)
