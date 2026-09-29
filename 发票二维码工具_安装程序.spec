# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['installer_src/app_installer.py'],
    # spec 所在目录加入搜索路径；下方路径一律用**相对项目根**的写法，
    # 避免把本机绝对路径写进公开仓库（克隆到任何目录都能构建）。
    pathex=[SPECPATH],
    binaries=[],
    datas=[('app_icon.ico', '.'), ('dist/发票二维码工具.exe', 'app_payload'), ('dist/uninstaller.exe', '.')],
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
    icon=['app_icon.ico'],
)
