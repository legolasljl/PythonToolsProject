#!/bin/bash
# =====================================================
# macOS 应用打包脚本
# Smart Clause Comparison Tool v16
# =====================================================

set -e  # 出错时停止

echo "=========================================="
echo "  智能条款比对工具 - macOS 打包脚本"
echo "=========================================="

# 检查 PyInstaller
if ! command -v pyinstaller &> /dev/null; then
    echo "❌ PyInstaller 未安装，正在安装..."
    pip install pyinstaller
fi

# 清理旧的构建文件
echo ""
echo "📦 清理旧的构建文件..."
rm -rf build/ dist/ *.spec.bak

# 检查必要文件
echo ""
echo "🔍 检查必要文件..."

MAIN_SCRIPT="clause_diff_gui_ultimate_v16.py"
CONFIG_MANAGER="clause_config_manager.py"
CONFIG_JSON="clause_config.json"
ICON_FILE="icon.icns"

for file in "$MAIN_SCRIPT" "$CONFIG_MANAGER" "$CONFIG_JSON" "$ICON_FILE"; do
    if [ ! -f "$file" ]; then
        echo "❌ 缺少必要文件: $file"
        exit 1
    fi
    echo "  ✓ $file"
done

# 开始打包
echo ""
echo "🔨 开始打包..."
pyinstaller clause_diff_gui.spec

# 检查结果
if [ -d "dist/SmartClauseMatcher.app" ]; then
    echo ""
    echo "=========================================="
    echo "✅ 打包成功！"
    echo "=========================================="
    echo ""
    echo "📍 应用位置: dist/SmartClauseMatcher.app"
    echo ""
    echo "📋 使用说明："
    echo "  1. 将 SmartClauseMatcher.app 拖入 /Applications 文件夹"
    echo "  2. 首次运行时右键选择「打开」以绕过 Gatekeeper"
    echo "  3. 自定义配置保存在 ~/.clause_diff/config.json"
    echo ""

    # 显示应用大小
    APP_SIZE=$(du -sh "dist/SmartClauseMatcher.app" | cut -f1)
    echo "📊 应用大小: $APP_SIZE"

    # 可选：打开 dist 目录
    # open dist/
else
    echo ""
    echo "❌ 打包失败，请检查错误信息"
    exit 1
fi
