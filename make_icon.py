# -*- coding: utf-8 -*-
import os
import subprocess
import shutil
from PIL import Image

def create_icns_from_local(source_path, output_name="icon.icns"):
    # 1. 检查源文件是否存在
    if not os.path.exists(source_path):
        print(f"❌ 错误：找不到文件 {source_path}")
        print("请检查路径是否正确，或将图片拖入终端获取绝对路径。")
        return

    print(f"1. 正在读取本地图片: {source_path}...")
    try:
        # 读取图片并强制转换为 RGBA 模式 (解决 JPG 格式兼容性问题)
        img = Image.open(source_path)
        img = img.convert("RGBA")
        print("✅ 图片读取并解析成功")
    except Exception as e:
        print(f"❌ 图片处理出错: {e}")
        return

    # 2. 创建临时的 iconset 文件夹
    iconset_dir = "MyIcon.iconset"
    if os.path.exists(iconset_dir):
        shutil.rmtree(iconset_dir) # 如果存在旧的先清理
    os.makedirs(iconset_dir)

    # macOS 需要的尺寸列表 (标准 Retina 适配)
    sizes = [
        (16, "icon_16x16.png"),
        (32, "icon_16x16@2x.png"),
        (32, "icon_32x32.png"),
        (64, "icon_32x32@2x.png"),
        (128, "icon_128x128.png"),
        (256, "icon_128x128@2x.png"),
        (256, "icon_256x256.png"),
        (512, "icon_256x256@2x.png"),
        (512, "icon_512x512.png"),
        (1024, "icon_512x512@2x.png")
    ]

    print("2. 正在生成不同尺寸的图标...")
    for size, filename in sizes:
        # 使用高质量重采样算法缩放
        resized_img = img.resize((size, size), Image.Resampling.LANCZOS)
        save_path = os.path.join(iconset_dir, filename)
        resized_img.save(save_path, format="PNG")
    
    print("3. 正在打包为 .icns 文件...")
    try:
        # 调用 macOS 系统命令 iconutil
        subprocess.run(["iconutil", "-c", "icns", iconset_dir], check=True)
        
        # 重命名/移动生成的 MyIcon.icns 到目标文件名
        if os.path.exists("MyIcon.icns"):
            if os.path.exists(output_name):
                os.remove(output_name)
            os.rename("MyIcon.icns", output_name)
            print(f"🎉 成功！图标已生成: {os.path.abspath(output_name)}")
            
    except subprocess.CalledProcessError:
        print("❌ iconutil 命令执行失败，请确保您在 macOS 环境下运行。")
    finally:
        # 清理临时文件夹
        if os.path.exists(iconset_dir):
            shutil.rmtree(iconset_dir)

if __name__ == "__main__":
    # 这里填你刚才说的本地路径
    LOCAL_IMAGE_PATH = "/Users/jinlei/Documents/icon.jpg"
    
    create_icns_from_local(LOCAL_IMAGE_PATH, "icon.icns")