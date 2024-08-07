# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['E:\\ONETHE416\\ADE\\ade-bms-python'],
    binaries=[],
    datas=[
        ('E:\\ONETHE416\\ADE\\ade-bms-python\\assets\\logo\\drdo_icon.ico', 'assets/logo'),
        ('E:\\ONETHE416\\ADE\\ade-bms-python\\assets\\images\\bg_main1.png', 'assets/images'),
        ('E:\\ONETHE416\\ADE\\ade-bms-python\\assets\\images\\pcan_btn_img.png', 'assets/images'),
        ('E:\\ONETHE416\\ADE\\ade-bms-python\\assets\\images\\moxa_rs232_btn_img.png', 'assets/images'),
        ('E:\\ONETHE416\\ADE\\ade-bms-python\\assets\\images\\brenergy_battery_img.png', 'assets/images'),
        ('C:\\Python312\\tcl\\tcl8.6', 'tcl'),
        ('C:\\Python312\\tcl\\tk8.6', 'tk')
    ],
    hiddenimports=[
        'tkinter', 
        'PIL', 
        'ttkbootstrap', 
        'customtkinter', 
        'helpers.methods', 
        'PCAN_API.custom_pcan_methods', 
        'gui.can_connect'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
