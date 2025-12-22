# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Smart Clause Comparison Tool v16
Windows .exe 打包配置

Usage (在 Windows 上执行):
    pyinstaller clause_diff_gui_win.spec

Author: Dachi Yijin
Date: 2025-12-22
"""

import os
import sys

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.abspath(SPEC))

# 主程序入口
MAIN_SCRIPT = os.path.join(PROJECT_ROOT, 'clause_diff_gui_ultimate_v16.py')

# 应用信息
APP_NAME = 'SmartClauseMatcher'
APP_VERSION = '16.0'

# 数据文件
datas = [
    (os.path.join(PROJECT_ROOT, 'clause_config_manager.py'), '.'),
    (os.path.join(PROJECT_ROOT, 'clause_config.json'), '.'),
    (os.path.join(PROJECT_ROOT, 'clause_config_client.json'), '.'),
]

# 隐式导入
hidden_imports = [
    'openpyxl',
    'openpyxl.cell',
    'openpyxl.styles',
    'openpyxl.utils',
    'pandas',
    'PyQt5',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
]

# 可选：翻译模块
try:
    import deep_translator
    hidden_imports.append('deep_translator')
except ImportError:
    pass

a = Analysis(
    [MAIN_SCRIPT],
    pathex=[PROJECT_ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'numpy.testing',
        'tkinter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# 单文件 EXE（推荐）
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 无控制台窗口（GUI程序）
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(PROJECT_ROOT, 'icon.ico'),  # Windows 图标
    version='file_version_info.txt',  # 可选：版本信息
)
