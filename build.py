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
    # PyInstaller 参数
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",           # 单文件模式
        "--windowed",          # 窗口模式（不显示控制台）
        "--name", "专利标记助手",  # exe文件名
        "--noconfirm",         # 覆盖输出目录
        "--clean",             # 清理缓存
        # 添加数据文件（如果有图标）
        # "--icon", "resources/icon.ico",
        # 隐藏导入
        "--hidden-import", "PyQt6.sip",
        # 入口文件
        "main.py"
    ]

    print("=" * 50)
    print("开始打包 专利标记助手...")
    print("=" * 50)
    print(f"命令: {' '.join(cmd)}")
    print()

    result = subprocess.run(cmd, cwd=os.path.dirname(__file__) or ".")

    if result.returncode == 0:
        print("\n" + "=" * 50)
        print("打包成功！")
        print(f"输出文件: dist/专利标记助手.exe")
        print("=" * 50)
    else:
        print("\n" + "=" * 50)
        print(f"打包失败，退出码: {result.returncode}")
        print("=" * 50)

    return result.returncode


if __name__ == "__main__":
    sys.exit(build())
