# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],  # 主程序文件名
    pathex=[],
    binaries=[],
    datas=[
        ('xdlovelife.ico', '.'),  # 图标文件
        ('app_config.json', '.'),  # 配置文件
    ],
    hiddenimports=[
        'pandas',
        'selenium',
        'win32com.client',
        'psutil',
        'xlwt',
        'openpyxl',
        'xlrd'
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
    name='阿里巴巴数据处理工具',  # 生成的exe名称
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为False以隐藏控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='xdlovelife.ico',  # 程序图标
    version='file_version_info.txt',  # 版本信息文件
) 