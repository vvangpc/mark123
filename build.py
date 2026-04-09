# -*- coding: utf-8 -*-
"""
build.py — PyInstaller 打包脚本
将专利标记助手打包为独立的 Windows exe 文件。
"""
import subprocess
import sys
import os
# 强制 stdout 使用 UTF-8，避免在 Windows GBK 控制台下 print emoji 报错
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

def build():
    """执行打包"""
    base_dir = os.path.dirname(os.path.abspath(__file__))

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # 单文件模式
        "--windowed",                   # 窗口模式（不显示控制台）
        "--icon", os.path.join(base_dir, "app_icon.ico"),  # 应用图标
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
        # ── 排除 NLP 相关重型依赖 ──
        # 离线 NLP 引擎（pycorrector + torch + transformers + 模型）只供
        # 「专业用户」在系统 Python 中手动安装，绝不打包进 exe，
        # 否则 exe 体积会从 ~80MB 飙到 2GB+ 且启动极慢。
        # cleaner.py 中 check_typos_pycorrector() 已用 try/except ImportError
        # 做了优雅降级，因此打包后软件仍能完整运行（仅 NLP 路径直接 return []）。
        "--exclude-module", "torch",
        "--exclude-module", "torchvision",
        "--exclude-module", "torchaudio",
        "--exclude-module", "transformers",
        "--exclude-module", "pycorrector",
        "--exclude-module", "datasets",
        "--exclude-module", "tokenizers",
        "--exclude-module", "huggingface_hub",
        "--exclude-module", "safetensors",
        "--exclude-module", "sentencepiece",
        "--exclude-module", "tensorflow",
        "--exclude-module", "tensorboard",
        "--exclude-module", "sklearn",
        "--exclude-module", "scipy",
        "--exclude-module", "pandas",
        "--exclude-module", "matplotlib",
        "--exclude-module", "PIL",
        "--exclude-module", "cv2",
        "--exclude-module", "jieba",
        "--exclude-module", "pypinyin",
        "--exclude-module", "kenlm",
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
