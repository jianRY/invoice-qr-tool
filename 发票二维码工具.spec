# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['invoice_qr_tool.py'],
    # 主程序按功能拆出了 iqr_net / iqr_summary / iqr_update 三个同级模块，
    # 显式把 spec 所在目录加进搜索路径，确保它们一定被收进包里。
    pathex=[SPECPATH],
    binaries=[],
    datas=[('app_icon.ico', '.')],
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
    name='发票二维码工具',
    icon='app_icon.ico',
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
