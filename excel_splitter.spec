# excel_splitter.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Excell Spliter.py'],  # Replace with your actual script name
    pathex=[],
    binaries=[],
    datas=[('app.ico', '.')],
    hiddenimports=[
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.parsing',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.tslibs.conversion',
        'pandas._libs.tslibs.fields',
        'pandas._libs.tslibs.holiday',
        'pandas._libs.tslibs.frequencies',
        'pandas._libs.tslibs.timezone',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.strptime',
        'pandas._libs.skiplist',
        'pandas._libs.sparse',
        'pandas._libs.reduction',
        'pandas._libs.parsers',
        'pandas._libs.groupby',
        'pandas._libs.reshape',
        'pandas._libs.tslib',
        'pandas._libs.hashtable',
        'pandas._libs.indexing',
        'pandas._libs.interval',
        'pandas._libs.writers',
        'pandas._libs.join',
        'pandas._libs.window',
        'pandas._libs.aggregations',
        'pandas._libs.parsers',
        'pandas._libs.algos',
        'pandas._libs.properties',
        'pandas._libs.hashing',
        'openpyxl',
        'xlrd',
        'xlwt',
        'xlsxwriter',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'msoffcrypto',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'notebook',
        'jupyter',
        'IPython',
        'tornado',
        'jedi',
        'tkinter',
    ],
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
    name='ExcelSplitterPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False = no console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',  # Optional: for version info
    icon='app.ico',  # Your icon file
    uac_admin=False,  # Don't require admin rights
    uac_uiaccess=False,
)