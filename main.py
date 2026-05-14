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
from version import __version__
from single_instance import try_send_to_running, install_listener


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
    app.setApplicationVersion(__version__)

    # 设置默认字体
    font = QFont("Microsoft YaHei UI", 10)
    app.setFont(font)

    # 应用浅色主题样式作为默认
    app.setStyleSheet(LIGHT_THEME_QSS)

    # 解析命令行带入的 docx 路径（若有）
    incoming_file = ""
    if len(sys.argv) > 1:
        candidate = sys.argv[1]
        if os.path.isfile(candidate) and candidate.lower().endswith('.docx'):
            incoming_file = candidate

    # 单实例：若已有实例在跑，把路径转交给它然后立刻退出
    if incoming_file and try_send_to_running(incoming_file):
        sys.exit(0)

    # 创建并显示主窗口
    window = MainWindow()

    if incoming_file:
        # 延迟加载，等窗口完全显示后再打开文件
        QTimer.singleShot(500, lambda: window._load_document(incoming_file))

    window.show()

    # 第一实例：挂上 IPC 监听，句柄挂在 window 上避免 GC
    window._single_instance_server = install_listener(window)

    # 等主窗口首次绘制完成再关闭 splash —— singleShot(0) 会被排到首个 paint
    # 事件之后执行，使 logo 显示时长恰好等于真实启动时间：启动越快 logo 越短，
    # 且不会出现「splash 已消失但主窗口尚未绘制」的视觉空档。
    QTimer.singleShot(0, _close_pyi_splash)

    # 启动 1.5s 后后台检查更新；frozen 产物默认开，dev 模式默认关
    # —— 句柄挂在 window 上避免被 gc，window 销毁时同步销毁
    from updater import UpdateChecker, should_check
    if should_check():
        window._update_checker = UpdateChecker(window, __version__)
        QTimer.singleShot(1500, window._update_checker.start)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
