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
from PyQt6.QtCore import QTimer

from main_window import MainWindow
from styles import LIGHT_THEME_QSS


def _close_pyi_splash():
    """关闭 PyInstaller 启动画面（仅在打包产物中可用）"""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except ImportError:
        pass


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
            QTimer.singleShot(500, lambda: window._load_document(file_path))

    window.show()

    # 等主窗口首次绘制完成再关闭 splash —— singleShot(0) 会被排到首个 paint
    # 事件之后执行，使 logo 显示时长恰好等于真实启动时间：启动越快 logo 越短，
    # 且不会出现「splash 已消失但主窗口尚未绘制」的视觉空档。
    QTimer.singleShot(0, _close_pyi_splash)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
