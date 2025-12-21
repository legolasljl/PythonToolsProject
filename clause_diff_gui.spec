# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Smart Clause Comparison Tool v16
macOS .app 打包配置

Usage:
    pyinstaller clause_diff_gui.spec

Author: Dachi Yijin
Date: 2025-12-21
"""

import os
import sys
from pathlib import Path

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.abspath(SPEC))

# 主程序入口
MAIN_SCRIPT = os.path.join(PROJECT_ROOT, 'clause_diff_gui_ultimate_v16.py')

# 应用信息
APP_NAME = 'SmartClauseMatcher'
APP_VERSION = '16.0'
BUNDLE_IDENTIFIER = 'com.clausetool.smartmatcher'

# 数据文件：(源路径, 目标目录)
# 在打包后，这些文件会被放到 app bundle 的 Resources 目录
datas = [
    # 配置管理模块
    (os.path.join(PROJECT_ROOT, 'clause_config_manager.py'), '.'),
    # 配置文件（作为默认配置）
    (os.path.join(PROJECT_ROOT, 'clause_config.json'), '.'),
]

# 隐式导入（PyInstaller 可能漏掉的模块）
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

# 可选：deep_translator（翻译功能）
try:
    import deep_translator
    hidden_imports.append('deep_translator')
except ImportError:
    pass

# Analysis 配置
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
        'matplotlib',  # 不需要的大型库
        'scipy',
        'numpy.testing',
        'PIL.ImageTk',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# 过滤不需要的二进制文件
a.binaries = [x for x in a.binaries if not x[0].startswith('libopenblas')]

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 无控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=True,  # macOS 支持拖放文件
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=APP_NAME,
)

# macOS .app Bundle
app = BUNDLE(
    coll,
    name=f'{APP_NAME}.app',
    icon=os.path.join(PROJECT_ROOT, 'icon.icns'),
    bundle_identifier=BUNDLE_IDENTIFIER,
    info_plist={
        'CFBundleName': APP_NAME,
        'CFBundleDisplayName': '智能条款比对工具',
        'CFBundleVersion': APP_VERSION,
        'CFBundleShortVersionString': APP_VERSION,
        'NSHighResolutionCapable': True,
        'NSRequiresAquaSystemAppearance': False,  # 支持深色模式
        'LSMinimumSystemVersion': '10.13.0',
        # 文件类型关联
        'CFBundleDocumentTypes': [
            {
                'CFBundleTypeName': 'Excel Workbook',
                'CFBundleTypeExtensions': ['xlsx', 'xls'],
                'CFBundleTypeRole': 'Viewer',
            },
        ],
    },
)
