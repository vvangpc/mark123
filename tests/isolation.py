# -*- coding: utf-8 -*-
"""
isolation.py — 让测试不碰用户真实的应用设置与配置目录

MainWindow 关闭时会把勾选状态 / 窗口几何写进 QSettings（Windows 下是注册表），
测试里切过的开关会原样带进用户下次启动的软件。导入本模块即：
  · AppSettings 改用临时目录里的 INI 文件；
  · QStandardPaths 进入测试模式，词库等配置文件读写落到测试专用目录。
"""
import os
import tempfile

from PyQt6.QtCore import QSettings, QStandardPaths

import config.config_manager as _cm

QStandardPaths.setTestModeEnabled(True)
_INI = os.path.join(tempfile.mkdtemp(prefix="patent_marker_test_"), "settings.ini")


class _TempSettings(_cm.AppSettings):
    def __init__(self):
        self._s = QSettings(_INI, QSettings.Format.IniFormat)


_cm.AppSettings = _TempSettings
