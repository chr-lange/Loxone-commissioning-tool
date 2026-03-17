# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Loxone I/O Commissioning Tool
# Build with:  pyinstaller LoxoneCommissioning.spec

block_cipher = None

a = Analysis(
    ['loxone_checklist_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Bundle the core module alongside the GUI
        ('loxone_checklist.py', '.'),
        # Bundle the icon so the GUI can set it at runtime
        ('loxone_icon.ico', '.'),
    ],
    hiddenimports=[
        'reportlab',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.styles',
        'reportlab.lib.units',
        'reportlab.lib.colors',
        'reportlab.lib.enums',
        'reportlab.platypus',
        'reportlab.platypus.flowables',
        'reportlab.platypus.paragraph',
        'reportlab.platypus.tables',
        'reportlab.pdfgen',
        'reportlab.pdfgen.canvas',
        'requests',
        'urllib3',
        'urllib3.exceptions',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
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
    name='LoxoneCommissioning',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no black console window (GUI-only app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='loxone_icon.ico',
)
