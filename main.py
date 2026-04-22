# -*- coding: utf-8 -*-
"""
main.py — 专利标记助手 入口文件
"""
import sys
import os

# 设置高DPI支持
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFont

from main_window import MainWindow
from styles import LIGHT_THEME_QSS


def main():
    app = QApplication(sys.argv)

    # 设置应用属性
    app.setApplicationName("专利标记助手")
    app.setApplicationDisplayName("专利标记助手")

    # 设置默认字体
    font = QFont("Microsoft YaHei UI", 10)
    app.setFont(font)

    # 应用浅色主题样式作为默认
    app.setStyleSheet(LIGHT_THEME_QSS)

    # 创建并显示主窗口
    window = MainWindow()

    # 如果命令行传入了文件路径，直接打开
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if os.path.isfile(file_path) and file_path.lower().endswith('.docx'):
            # 延迟加载，等窗口完全显示后再打开文件
            from PyQt6.QtCore import QTimer
            QTimer.singleShot(500, lambda: window._load_document(file_path))

    window.show()

    # 若由 PyInstaller 打包并启用了 splash，窗口显示后立即关闭启动图
    try:
        import pyi_splash  # type: ignore  # 仅在 PyInstaller 打包产物中存在
        pyi_splash.close()
    except ImportError:
        pass

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
