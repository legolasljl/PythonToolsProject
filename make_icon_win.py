#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows ICO 图标生成器
将 JPG/PNG 图片转换为 Windows .ico 格式

Usage:
    python make_icon_win.py [input_image] [output_ico]

Examples:
    python make_icon_win.py                    # 默认: icon.jpg -> icon.ico
    python make_icon_win.py logo.png app.ico   # 自定义输入输出
"""

import sys
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("错误: 需要安装 Pillow 库")
    print("请运行: pip install Pillow")
    sys.exit(1)


def create_ico(input_path: str, output_path: str):
    """创建 Windows ICO 图标文件"""

    # ICO 文件需要的尺寸（Windows 标准）
    sizes = [16, 24, 32, 48, 64, 128, 256]

    print(f"正在读取: {input_path}")

    # 打开源图片
    img = Image.open(input_path)

    # 转换为 RGBA（支持透明）
    if img.mode != 'RGBA':
        img = img.convert('RGBA')

    # 确保是正方形
    width, height = img.size
    if width != height:
        # 取较小的边，居中裁剪
        size = min(width, height)
        left = (width - size) // 2
        top = (height - size) // 2
        img = img.crop((left, top, left + size, top + size))
        print(f"已裁剪为正方形: {size}x{size}")

    # 生成各尺寸图标
    icon_images = []
    for size in sizes:
        resized = img.resize((size, size), Image.Resampling.LANCZOS)
        icon_images.append(resized)
        print(f"  生成 {size}x{size} 图标")

    # 保存为 ICO
    icon_images[0].save(
        output_path,
        format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=icon_images[1:]
    )

    print(f"\n✓ ICO 图标已保存: {output_path}")
    print(f"  包含尺寸: {', '.join(f'{s}x{s}' for s in sizes)}")


def main():
    # 默认文件名
    input_file = "icon.jpg"
    output_file = "icon.ico"

    # 命令行参数
    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]

    # 检查输入文件
    if not Path(input_file).exists():
        # 尝试其他格式
        for ext in ['.png', '.jpeg', '.jpg', '.bmp']:
            alt_file = Path(input_file).stem + ext
            if Path(alt_file).exists():
                input_file = alt_file
                break
        else:
            print(f"错误: 未找到图片文件 '{input_file}'")
            print("用法: python make_icon_win.py [input_image] [output_ico]")
            sys.exit(1)

    create_ico(input_file, output_file)


if __name__ == "__main__":
    main()
