# -*- coding: utf-8 -*-
"""
build.py — PyInstaller 打包脚本

用法：
    python build.py            # 默认 onedir 模式（启动最快，推荐）
    python build.py onedir     # 同上
    python build.py onefile    # 单文件模式（启动慢 3-10 秒，但只有一个 exe）

两种模式对比：
    onedir   → dist/专利标记助手/ 文件夹 + 专利标记助手.exe
               · 启动 ~1-2 秒；分发时把整个文件夹打 zip
    onefile  → dist/专利标记助手.exe 单文件
               · 启动 ~5-15 秒（每次都要解压到 %TEMP%\\_MEIxxx）
"""
import subprocess
import sys
import os

# 强制 stdout 使用 UTF-8，避免在 Windows GBK 控制台下 print emoji 报错
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


# 共用参数（两种模式都要）
_COMMON_ARGS = [
    "--windowed",                   # 窗口模式（不显示控制台）
    "--noconfirm",                  # 覆盖输出目录
    "--clean",                      # 清理缓存
    "--noupx",                      # 关闭 UPX 压缩 —— UPX 压缩的 DLL 启动时要
                                    # 先解压，反而拖慢冷启动；多占 ~30MB 磁盘换
                                    # ~0.5-1 秒启动提速，值得
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
]


def build(mode: str = "onedir"):
    """执行打包。mode: 'onedir' | 'onefile'"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_dir, "app_icon.ico")
    splash_png = os.path.join(base_dir, "app_icon.png")

    mode_flag = "--onedir" if mode == "onedir" else "--onefile"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        mode_flag,
        "--icon", icon_path,
        "--name", "专利标记助手",
    ]

    # Splash screen：启动时先显示一张图，遮住 Python/PyQt 初始化的 1-2 秒。
    # 仅当 app_icon.png 存在时启用；PyInstaller 6.x 通过 pyi_splash 模块
    # 在主窗口 show() 后手动关闭（见 main.py）。
    if os.path.exists(splash_png):
        cmd.extend(["--splash", splash_png])

    cmd.extend(_COMMON_ARGS)
    cmd.append("main.py")

    print("=" * 60)
    print(f"开始打包「专利标记助手」...  模式：{mode}")
    print("=" * 60)

    result = subprocess.run(cmd, cwd=base_dir)

    if result.returncode == 0:
        if mode == "onedir":
            out = os.path.join(base_dir, "dist", "专利标记助手")
            print("\n" + "=" * 60)
            print("✅ 打包成功！")
            print(f"   输出目录: {out}")
            print(f"   exe 位置: {os.path.join(out, '专利标记助手.exe')}")
            print(f"   分发建议: 把整个「专利标记助手」文件夹打成 zip 发给用户")
            print("=" * 60)
        else:
            out = os.path.join(base_dir, "dist", "专利标记助手.exe")
            print("\n" + "=" * 60)
            print("✅ 打包成功！")
            print(f"   输出文件: {out}")
            print(f"   注意: onefile 模式首次启动会解压到 %TEMP% 需等 5-15 秒")
            print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print(f"❌ 打包失败，退出码: {result.returncode}")
        print("=" * 60)

    return result.returncode


if __name__ == "__main__":
    mode = "onedir"
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg in ("onedir", "onefile"):
            mode = arg
        else:
            print(f"未知模式：{arg}，使用默认 onedir")
    sys.exit(build(mode))
