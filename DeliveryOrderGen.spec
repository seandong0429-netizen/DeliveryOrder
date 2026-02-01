# -*- mode: python ; coding: utf-8 -*-
import sys


a = Analysis(
    ['src/main.py'],
    pathex=['src'],
    binaries=[],
    datas=[('wechat_qr.png', '.'), ('resources', 'resources')],
    hiddenimports=['openpyxl', 'reportlab', 'PIL'],
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
    name='DeliveryOrderGen',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='arm64' if sys.platform == 'darwin' else None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['resources/app_icon.icns'],
)
app = BUNDLE(
    exe,
    name='DeliveryOrderGen.app',
    icon='resources/app_icon.icns',
    bundle_identifier=None,
)
