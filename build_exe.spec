# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['analytics_ui/excel_merger.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('analytics_ui/icon.png', '.'),
        ('analytics_ui/icon.ico', '.'),
        ('analytics_ui/Правила названия столбцов.xlsx', '.'),
    ],
    hiddenimports=[
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.skiplist',
        'openpyxl',
        'xlrd',
        'numpy',
        'xlsxwriter',
        'unicodedata',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'PySide2', 'PySide6'],
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
    name='AnalyticsUI_OGPZ',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='analytics_ui/icon.ico',
)
