# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 指定图标文件的完整路径
icon_path = 'xdlovelife1.ico'

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[(icon_path, '.')],  # 确保图标文件被包含
    hiddenimports=[
        'pandas',
        'numpy',
        'selenium',
        'openpyxl',
        'xlrd',
        'win32com',
        'win32com.client',
        'psutil',
        'requests',
        'tkinter',
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
    name='AliDataOutput',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,  # 设置程序图标
    version='file_version_info.txt'  # 添加版本信息
)
