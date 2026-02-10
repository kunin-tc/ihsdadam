# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('version.py', '.'), ('ihsdm_compiler_core.py', '.'), ('src/ihsdadam', 'ihsdadam')]
binaries = []
hiddenimports = ['PySide6', 'PySide6.QtWidgets', 'PySide6.QtCore', 'PySide6.QtGui', 'openpyxl', 'openpyxl.cell', 'openpyxl.styles', 'openpyxl.workbook', 'openpyxl.worksheet', 'PyPDF2']
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PyPDF2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pyxlsb')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'tkinter', 'matplotlib', 'IPython', 'sphinx', 'pandas', 'numpy'],
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
    name='IHSDadaM_Qt',
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
    icon='NONE',
)
