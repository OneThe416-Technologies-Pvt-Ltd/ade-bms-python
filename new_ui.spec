# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['new_ui.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/logo/drdo_icon.ico', 'assets/logo'),
    ('assets/images/splash_bg.png', 'assets/images'),
    ('assets/logo/ade_logo.png', 'assets/logo'),
    ('assets/images/disconnect_button.png', 'assets/images'),
    ('assets/images/load_icon.png', 'assets/images'),
    ('assets/images/settings.png', 'assets/images'),
    ('assets/images/info.png', 'assets/images'),
    ('assets/images/help.png', 'assets/images'),
    ('assets/images/menu.png', 'assets/images'),
    ('assets/images/download.png', 'assets/images'),
    ('assets/images/splash_bg.png', 'assets/images'),
    ('assets/images/refresh.png', 'assets/images'),
    ('assets/images/start.png', 'assets/images'),
    ('assets/images/stop.png', 'assets/images'),
    ('assets/images/folder.png', 'assets/images'),
    ('assets/images/reset.png', 'assets/images'),
    ('assets/images/activate.png', 'assets/images'),],
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
    [],
    exclude_binaries=True,
    name='ADE BMS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    icon='assets/logo/drdo_icon.ico',  # Path to your icon file
    upx=True,
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
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='new_ui',
)
