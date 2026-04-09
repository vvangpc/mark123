# -*- coding: utf-8 -*-
"""
make_ico.py — 将 app_icon.png 转换为多尺寸 ICO 文件
使用方法：python make_ico.py

【说明】
  PIL 在保存 256×256 的 ICO 帧时会自动改用 PNG 子格式编码，
  而 Windows 桌面/资源管理器对这种格式的透明通道处理存在 Bug，
  会将透明区域渲染为白色矩形背景框。
  解决方案：仅使用 ≤128px 的尺寸（均以 BMP+Alpha mask 编码），
  彻底避免白底框问题。Windows 会自动从较大尺寸缩放显示。
"""
from PIL import Image
import os
import sys

# 强制 stdout 使用 UTF-8
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

src = os.path.join(os.path.dirname(__file__), "app_icon.png")
dst = os.path.join(os.path.dirname(__file__), "app_icon.ico")

if not os.path.exists(src):
    print(f"❌ 找不到源文件: {src}")
    print("请先把图标图片（PNG）命名为 app_icon.png 放到项目根目录。")
    raise SystemExit(1)

img = Image.open(src).convert("RGBA")

# ✅ 去掉 256px：PIL 会对 >=256px 的帧自动使用 PNG 子格式，
#    Windows 某些渲染路径不正确处理其透明通道，导致白底框。
#    128px 及以下均以 BMP+Alpha mask 编码，透明度完全正常。
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128)]

# 直接使用最高分辨率原图的 save 方法，并传入 sizes 参数，
# PIL 会自动缩放并以 BMP 子格式封装不同尺寸进入 ICO 中。
img.save(
    dst,
    format="ICO",
    sizes=sizes
)

print(f"[OK] ico generated: {dst}")
print(f"     sizes: {sizes}")
print("     ✅ 已排除 256px 帧，透明背景正常，无白底框问题")
