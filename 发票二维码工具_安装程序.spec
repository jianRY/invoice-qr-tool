# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:/workbuddy/发票处理/installer_src/app_installer.py'],
    pathex=[],
    binaries=[],
    datas=[('D:/workbuddy/发票处理/app_icon.ico', '.'), ('D:/workbuddy/发票处理/dist/发票二维码工具.exe', 'app_payload'), ('D:/workbuddy/发票处理/dist/uninstaller.exe', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='发票二维码工具_安装程序',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=True,
    icon=['D:/workbuddy/发票处理/app_icon.ico'],
)
