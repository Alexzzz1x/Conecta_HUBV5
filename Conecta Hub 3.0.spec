# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = []
hiddenimports += collect_submodules('reportlab.graphics.barcode')
hiddenimports += collect_submodules('reportlab.lib')
hiddenimports += collect_submodules('reportlab.pdfgen')
hiddenimports += collect_submodules('reportlab.graphics')
hiddenimports += collect_submodules('tkcalendar')
hiddenimports += collect_submodules('babel')

hiddenimports += [
    'numpy',
    'pandas',
    'PIL',
    'openpyxl',
    'qrcode',
    'win32com',
    'win32com.client',
    'pythoncom',
    'tkcalendar',
    'babel.numbers',
    'reportlab.graphics.barcode.code128',
    'reportlab.lib.pagesizes',
    'reportlab.lib.colors',
    'reportlab.pdfgen.canvas',
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'tkinter.filedialog',
]

datas = [
    ('logo_cntc.png', '.'),
    ('eclipse.ico', '.'),
    ('icone_lua.ico', '.'),
    ('icone_lua.png', '.'),
    ('icone_taskbar.png', '.'),
    ('banco_de_dados.xlsx', '.'),
    ('SCRIPT DA IQ09.vbs', '.'),
    ('SCRIPT DA BANDEIRADA.vbs', '.'),
    ('Sciprt OSME2.vbs', '.'),
    ('SCRIPT DE RETE.vbs', '.'),
    ('OSME CL03.vbs', '.'),
    ('OSME CL04.vbs', '.'),
]

datas += collect_data_files('babel')
datas += collect_data_files('tkcalendar')

a = Analysis(
    ['conecta_hub2.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'matplotlib', 'scipy',
        'IPython', 'jupyter',
        'sqlalchemy', 'flask', 'django',
        'zmq', 'pyzmq',
        'setuptools', 'pip', 'wheel',
        'cffi', 'cryptography',
        'boto3', 'botocore', 'requests_aws',
        'cv2',
        'torch', 'tensorflow',
        'nltk', 'sklearn',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Conecta Hub 3.0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['eclipse.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Conecta Hub 3.0',
)
