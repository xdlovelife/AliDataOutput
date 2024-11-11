# -*- mode: python ; coding: utf-8 -*-

import sys
import os

block_cipher = None

# 获取当前目录
CURR_DIR = os.getcwd()

# 添加资源文件
added_files = [
    ('app_config.json', '.'),
    ('xdlovelife1.ico', '.'),
]

a = Analysis(
    ['main.py'],
    pathex=[CURR_DIR],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        'win32com.client',
        'pandas',
        'selenium',
        'openpyxl',
        'xlrd',
        'xlwt',
        'psutil',
        'requests',
        'webdriver_manager',
        'tkinter',
        'json',
        'threading',
        'zipfile',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.common.by',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.common.keys',
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

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='阿里巴巴邮箱数据获取工具',
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
    icon='xdlovelife1.ico',
    uac_admin=True,
    version='file_version_info.txt'
) 