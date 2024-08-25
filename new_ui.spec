# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_submodules

# Define the base directory explicitly
base_dir = 'E:\\ONETHE416\\ADE\\ade-bms-python'

# Collect submodules for the specified packages to avoid missing imports
hiddenimports = collect_submodules('tkinter') + collect_submodules('ttkbootstrap') + collect_submodules('PIL') + collect_submodules('pcan_api') + collect_submodules('gui')

a = Analysis(
    ['E:\\ONETHE416\\ADE\\ade-bms-python\\new_ui.py'],
    pathex=[base_dir],
    binaries=[],
    datas=[
        (os.path.join(base_dir, 'myenv\\Lib\\site-packages\\customtkinter'), 'customtkinter'),
        (os.path.join(base_dir, 'myenv\\Lib\\site-packages\\ttkbootstrap'), 'ttkbootstrap'),
        ('C:\\Program Files\\Python312\\tcl\\tcl8.6', 'tcl/tcl8.6'),
        ('C:\\Program Files\\Python312\\tcl\\tk8.6', 'tcl/tk8.6'),
        (os.path.join(base_dir, 'assets'), 'assets')
    ],
    hiddenimports=hiddenimports,
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
    name='ADE_BMS',
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
