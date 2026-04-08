# -*- coding: utf-8 -*-
"""
build.py — PyInstaller 打包脚本
将专利标记助手打包为独立的 Windows exe 文件。
"""
import subprocess
import sys
import os

def build():
    """执行打包"""
    base_dir = os.path.dirname(os.path.abspath(__file__))

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # 单文件模式
        "--windowed",                   # 窗口模式（不显示控制台）
        "--name", "专利标记助手",         # exe 文件名
        "--noconfirm",                  # 覆盖输出目录
        "--clean",                      # 清理缓存
        # ── 隐式导入 ──
        "--hidden-import", "PyQt6.sip",
        "--hidden-import", "PyQt6.QtCore",
        "--hidden-import", "PyQt6.QtGui",
        "--hidden-import", "PyQt6.QtWidgets",
        "--hidden-import", "docx",
        "--hidden-import", "docx.oxml.ns",
        "--hidden-import", "lxml",
        "--hidden-import", "lxml.etree",
        "--hidden-import", "lxml._elementpath",
        "--hidden-import", "lxml.html",
        # ── 收集整个包的数据文件 ──
        "--collect-all", "docx",
        "--collect-all", "lxml",
        # ── 入口文件 ──
        "main.py",
    ]

    print("=" * 60)
    print("开始打包「专利标记助手」...")
    print("=" * 60)

    result = subprocess.run(cmd, cwd=base_dir)

    if result.returncode == 0:
        print("\n" + "=" * 60)
        print("✅ 打包成功！")
        print(f"   输出文件: {os.path.join(base_dir, 'dist', '专利标记助手.exe')}")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print(f"❌ 打包失败，退出码: {result.returncode}")
        print("=" * 60)

    return result.returncode


if __name__ == "__main__":
    sys.exit(build())
