# -*- coding: utf-8 -*-
"""
styles.py — Qt 深色主题样式表
为专利标记助手桌面应用提供现代深色玻璃态界面样式。
"""

DARK_THEME_QSS = """
/* ===== 全局基础 ===== */
QMainWindow {
    background-color: #0f1419;
}

QWidget {
    color: #e0e0e0;
    font-family: "Microsoft YaHei UI", "Segoe UI", "Inter", sans-serif;
    font-size: 13px;
}

/* ===== 标题标签 ===== */
QLabel#titleLabel {
    font-size: 22px;
    font-weight: bold;
    color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00bfa5, stop:1 #2962ff);
    padding: 8px 0;
}

QLabel#subtitleLabel {
    font-size: 13px;
    color: #9e9e9e;
    padding: 2px 0;
}

QLabel#sectionLabel {
    font-size: 14px;
    font-weight: bold;
    color: #00bfa5;
    padding: 6px 0;
}

/* ===== 玻璃态面板 ===== */
QFrame#glassPanel {
    background-color: rgba(15, 20, 25, 0.85);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 12px;
    padding: 16px;
}

QFrame#headerPanel {
    background-color: rgba(15, 20, 25, 0.9);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 8px;
}

/* ===== 文本框 ===== */
QTextEdit, QPlainTextEdit {
    background-color: rgba(15, 20, 25, 0.7);
    color: #ffffff;
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 10px;
    padding: 14px;
    font-size: 13px;
    line-height: 1.6;
    selection-background-color: #00bfa5;
    selection-color: #ffffff;
}

QTextEdit:focus, QPlainTextEdit:focus {
    border-color: #00bfa5;
    background-color: rgba(20, 25, 32, 0.85);
}

QTextEdit:hover, QPlainTextEdit:hover {
    border-color: rgba(0, 191, 165, 0.4);
}

/* ===== 按钮 ===== */
QPushButton {
    background-color: rgba(15, 20, 25, 0.8);
    color: #ffffff;
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 22px;
    padding: 12px 24px;
    font-size: 14px;
    font-weight: bold;
    min-height: 20px;
}

QPushButton:hover {
    background-color: #00bfa5;
    border-color: #00bfa5;
}

QPushButton:pressed {
    background-color: #00897b;
}

QPushButton:disabled {
    background-color: rgba(60, 60, 60, 0.5);
    color: rgba(255, 255, 255, 0.3);
    border-color: rgba(255, 255, 255, 0.05);
}

QPushButton#primaryBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #00bfa5, stop:1 #00897b);
    border: none;
    font-size: 15px;
    padding: 14px 32px;
}

QPushButton#primaryBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #26d9c0, stop:1 #00bfa5);
}

QPushButton#primaryBtn:pressed {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #00897b, stop:1 #00695c);
}

QPushButton#accentBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #2962ff, stop:1 #5e35b1);
    border: none;
    font-size: 14px;
}

QPushButton#accentBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #448aff, stop:1 #7e57c2);
}

QPushButton#fileBtn {
    background-color: rgba(41, 98, 255, 0.15);
    border: 2px dashed rgba(41, 98, 255, 0.4);
    border-radius: 14px;
    padding: 14px;
    font-size: 15px;
    color: #90caf9;
    min-height: 40px;
}

QPushButton#fileBtn:hover {
    background-color: rgba(41, 98, 255, 0.25);
    border-color: rgba(41, 98, 255, 0.7);
    color: #ffffff;
}

QPushButton#smallBtn {
    padding: 6px 16px;
    font-size: 12px;
    border-radius: 16px;
    min-height: 10px;
    background-color: rgba(255, 255, 255, 0.05);
    border: 1px solid rgba(255, 255, 255, 0.15);
}

QPushButton#smallBtn:hover {
    background-color: #00bfa5;
    border-color: #00bfa5;
}

QPushButton#dangerBtn {
    background-color: rgba(255, 82, 82, 0.15);
    border: 1px solid rgba(255, 82, 82, 0.3);
    color: #ff8a80;
}

QPushButton#dangerBtn:hover {
    background-color: #ff5252;
    border-color: #ff5252;
    color: #ffffff;
}

/* ===== 进度条 ===== */
QProgressBar {
    background-color: rgba(255, 255, 255, 0.05);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 8px;
    text-align: center;
    color: #ffffff;
    font-size: 11px;
    height: 18px;
}

QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00bfa5, stop:1 #2962ff);
    border-radius: 7px;
}

/* ===== 状态栏 ===== */
QStatusBar {
    background-color: rgba(10, 14, 18, 0.95);
    color: #9e9e9e;
    border-top: 1px solid rgba(255, 255, 255, 0.05);
    font-size: 12px;
    padding: 4px;
}

/* ===== 滚动条 ===== */
QScrollBar:vertical {
    background: rgba(0, 0, 0, 0.2);
    width: 10px;
    border-radius: 5px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: rgba(255, 255, 255, 0.15);
    border-radius: 5px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background: rgba(255, 255, 255, 0.3);
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QScrollBar:horizontal {
    background: rgba(0, 0, 0, 0.2);
    height: 10px;
    border-radius: 5px;
}

QScrollBar::handle:horizontal {
    background: rgba(255, 255, 255, 0.15);
    border-radius: 5px;
    min-width: 30px;
}

QScrollBar::handle:horizontal:hover {
    background: rgba(255, 255, 255, 0.3);
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}

/* ===== 标签页 ===== */
QTabWidget::pane {
    background-color: rgba(15, 20, 25, 0.85);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 10px;
    padding: 8px;
}

QTabBar::tab {
    background-color: rgba(255, 255, 255, 0.05);
    color: #9e9e9e;
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 10px 20px;
    margin-right: 4px;
    font-size: 13px;
}

QTabBar::tab:selected {
    background-color: rgba(15, 20, 25, 0.85);
    color: #00bfa5;
    font-weight: bold;
    border-color: rgba(0, 191, 165, 0.3);
}

QTabBar::tab:hover:!selected {
    background-color: rgba(255, 255, 255, 0.08);
    color: #e0e0e0;
}

/* ===== 分组框 ===== */
QGroupBox {
    background-color: rgba(15, 20, 25, 0.5);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 10px;
    margin-top: 12px;
    padding: 16px;
    padding-top: 28px;
    font-size: 13px;
    font-weight: bold;
    color: #00bfa5;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 4px 12px;
    color: #00bfa5;
}

/* ===== 表格 ===== */
QTableWidget, QTableView {
    background-color: rgba(15, 20, 25, 0.6);
    alternate-background-color: rgba(255, 255, 255, 0.04);
    color: #e0e0e0;
    gridline-color: rgba(255, 255, 255, 0.08);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 10px;
    selection-background-color: rgba(0, 191, 165, 0.35);
    selection-color: #ffffff;
}

QTableWidget::item, QTableView::item {
    padding: 4px 6px;
    border: none;
}

QTableWidget::item:selected, QTableView::item:selected {
    background-color: rgba(0, 191, 165, 0.35);
    color: #ffffff;
}

QHeaderView::section {
    background-color: rgba(15, 20, 25, 0.9);
    color: #00bfa5;
    border: none;
    border-right: 1px solid rgba(255, 255, 255, 0.06);
    border-bottom: 1px solid rgba(255, 255, 255, 0.12);
    padding: 6px 8px;
    font-weight: bold;
}

QHeaderView::section:last {
    border-right: none;
}

QTableCornerButton::section {
    background-color: rgba(15, 20, 25, 0.9);
    border: none;
    border-bottom: 1px solid rgba(255, 255, 255, 0.12);
}

/* ===== 消息框 ===== */
QMessageBox {
    background-color: #0f1419;
}

QMessageBox QLabel {
    color: #e0e0e0;
    font-size: 13px;
}

QMessageBox QPushButton {
    min-width: 80px;
}

/* ===== 文件对话框 ===== */
QFileDialog {
    background-color: #0f1419;
}

/* ===== 工具提示 ===== */
QToolTip {
    background-color: rgba(30, 30, 30, 0.95);
    color: #ffffff;
    border: 1px solid rgba(0, 191, 165, 0.3);
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
}

/* ===== 分割器 ===== */
QSplitter::handle {
    background-color: rgba(255, 255, 255, 0.05);
    width: 3px;
    height: 3px;
    border-radius: 1px;
}

QSplitter::handle:hover {
    background-color: #00bfa5;
}

/* ===== 检查字数选择栏（claim 检查 Tab） ===== */
QFrame#claimNBar {
    background-color: rgba(255, 255, 255, 0.04);
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-radius: 18px;
    padding: 4px 10px;
}

QFrame#claimNBar QLabel {
    color: #9e9e9e;
    font-size: 13px;
    padding: 0 2px;
}

QPushButton#nPresetBtn {
    min-width: 34px;
    max-width: 34px;
    min-height: 26px;
    max-height: 26px;
    padding: 0;
    font-size: 13px;
    font-weight: 600;
    color: #b0bec5;
    background-color: rgba(255, 255, 255, 0.04);
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 13px;
}

QPushButton#nPresetBtn:hover {
    background-color: rgba(0, 191, 165, 0.18);
    border-color: rgba(0, 191, 165, 0.55);
    color: #ffffff;
}

QPushButton#nPresetBtn:checked {
    background-color: #00bfa5;
    border-color: #00bfa5;
    color: #ffffff;
}

QSpinBox#nCustomSpin {
    min-height: 28px;
    padding: 2px 10px;
    font-size: 13px;
    font-weight: 600;
    color: #ffffff;
    background-color: rgba(0, 0, 0, 0.30);
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 13px;
    selection-background-color: #00bfa5;
}

QSpinBox#nCustomSpin:focus {
    border-color: #00bfa5;
}
"""

LIGHT_THEME_QSS = """
/* ===== 全局基础 ===== */
QMainWindow {
    background-color: #f5f7fa;
}

QWidget {
    color: #2c3e50;
    font-family: "Microsoft YaHei UI", "Segoe UI", "Inter", sans-serif;
    font-size: 13px;
}

/* ===== 标题标签 ===== */
QLabel#titleLabel {
    font-size: 22px;
    font-weight: bold;
    color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #009688, stop:1 #1976d2);
    padding: 8px 0;
}

QLabel#subtitleLabel {
    font-size: 13px;
    color: #7f8c8d;
    padding: 2px 0;
}

QLabel#sectionLabel {
    font-size: 14px;
    font-weight: bold;
    color: #009688;
    padding: 6px 0;
}

/* ===== 玻璃态面板 ===== */
QFrame#glassPanel {
    background-color: rgba(255, 255, 255, 0.85);
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 12px;
    padding: 16px;
}

QFrame#headerPanel {
    background-color: rgba(255, 255, 255, 0.9);
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 8px;
}

/* ===== 文本框 ===== */
QTextEdit, QPlainTextEdit {
    background-color: #ffffff;
    color: #2c3e50;
    border: 1px solid rgba(0, 0, 0, 0.1);
    border-radius: 10px;
    padding: 14px;
    font-size: 13px;
    line-height: 1.6;
    selection-background-color: #00bfa5;
    selection-color: #ffffff;
}

QTextEdit:focus, QPlainTextEdit:focus {
    border-color: #00bfa5;
    background-color: #fafafa;
}

QTextEdit:hover, QPlainTextEdit:hover {
    border-color: rgba(0, 191, 165, 0.4);
}

/* ===== 按钮 ===== */
QPushButton {
    background-color: #ffffff;
    color: #34495e;
    border: 1px solid rgba(0, 0, 0, 0.1);
    border-radius: 22px;
    padding: 12px 24px;
    font-size: 14px;
    font-weight: bold;
    min-height: 20px;
}

QPushButton:hover {
    background-color: #e0f2f1;
    border-color: #00bfa5;
    color: #00897b;
}

QPushButton:pressed {
    background-color: #b2dfdb;
}

QPushButton:disabled {
    background-color: rgba(220, 220, 220, 0.5);
    color: rgba(100, 100, 100, 0.4);
    border-color: rgba(0, 0, 0, 0.05);
}

QPushButton#primaryBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #00bfa5, stop:1 #00897b);
    color: #ffffff;
    border: none;
    font-size: 15px;
    padding: 14px 32px;
}

QPushButton#primaryBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #26d9c0, stop:1 #00bfa5);
    color: #ffffff;
}

QPushButton#primaryBtn:pressed {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #00897b, stop:1 #00695c);
}

QPushButton#accentBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #2962ff, stop:1 #5e35b1);
    color: #ffffff;
    border: none;
    font-size: 14px;
}

QPushButton#accentBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #448aff, stop:1 #7e57c2);
    color: #ffffff;
}

QPushButton#fileBtn {
    background-color: rgba(41, 98, 255, 0.05);
    border: 2px dashed rgba(41, 98, 255, 0.4);
    border-radius: 14px;
    padding: 14px;
    font-size: 15px;
    color: #1976d2;
    min-height: 40px;
}

QPushButton#fileBtn:hover {
    background-color: rgba(41, 98, 255, 0.15);
    border-color: rgba(41, 98, 255, 0.7);
    color: #0d47a1;
}

QPushButton#smallBtn {
    padding: 6px 16px;
    font-size: 12px;
    border-radius: 16px;
    min-height: 10px;
    background-color: #ffffff;
    border: 1px solid rgba(0, 0, 0, 0.15);
    color: #34495e;
}

QPushButton#smallBtn:hover {
    background-color: #e0f2f1;
    border-color: #00bfa5;
}

QPushButton#dangerBtn {
    background-color: rgba(255, 82, 82, 0.05);
    border: 1px solid rgba(255, 82, 82, 0.3);
    color: #d32f2f;
}

QPushButton#dangerBtn:hover {
    background-color: #ffebee;
    border-color: #ff5252;
    color: #c62828;
}

/* ===== 进度条 ===== */
QProgressBar {
    background-color: rgba(0, 0, 0, 0.05);
    border: 1px solid rgba(0, 0, 0, 0.1);
    border-radius: 8px;
    text-align: center;
    color: #2c3e50;
    font-size: 11px;
    height: 18px;
}

QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00bfa5, stop:1 #2962ff);
    border-radius: 7px;
}

/* ===== 状态栏 ===== */
QStatusBar {
    background-color: #eceff1;
    color: #546e7a;
    border-top: 1px solid rgba(0, 0, 0, 0.05);
    font-size: 12px;
    padding: 4px;
}

/* ===== 滚动条 ===== */
QScrollBar:vertical {
    background: rgba(0, 0, 0, 0.05);
    width: 10px;
    border-radius: 5px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: rgba(0, 0, 0, 0.2);
    border-radius: 5px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background: rgba(0, 0, 0, 0.35);
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QScrollBar:horizontal {
    background: rgba(0, 0, 0, 0.05);
    height: 10px;
    border-radius: 5px;
}

QScrollBar::handle:horizontal {
    background: rgba(0, 0, 0, 0.2);
    border-radius: 5px;
    min-width: 30px;
}

QScrollBar::handle:horizontal:hover {
    background: rgba(0, 0, 0, 0.35);
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}

/* ===== 标签页 ===== */
QTabWidget::pane {
    background-color: #ffffff;
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 10px;
    padding: 8px;
}

QTabBar::tab {
    background-color: #f5f7fa;
    color: #7f8c8d;
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 10px 20px;
    margin-right: 4px;
    font-size: 13px;
}

QTabBar::tab:selected {
    background-color: #ffffff;
    color: #009688;
    font-weight: bold;
    border-color: rgba(0, 150, 136, 0.3);
}

QTabBar::tab:hover:!selected {
    background-color: #e0f2f1;
    color: #00897b;
}

/* ===== 分组框 ===== */
QGroupBox {
    background-color: rgba(255, 255, 255, 0.6);
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 10px;
    margin-top: 12px;
    padding: 16px;
    padding-top: 28px;
    font-size: 13px;
    font-weight: bold;
    color: #009688;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 4px 12px;
    color: #009688;
}

/* ===== 表格 ===== */
QTableWidget, QTableView {
    background-color: #ffffff;
    alternate-background-color: #f5f7fa;
    color: #2c3e50;
    gridline-color: rgba(0, 0, 0, 0.08);
    border: 1px solid rgba(0, 0, 0, 0.1);
    border-radius: 10px;
    selection-background-color: rgba(0, 191, 165, 0.25);
    selection-color: #263238;
}

QTableWidget::item, QTableView::item {
    padding: 4px 6px;
    border: none;
}

QTableWidget::item:selected, QTableView::item:selected {
    background-color: rgba(0, 191, 165, 0.25);
    color: #263238;
}

QHeaderView::section {
    background-color: #eceff1;
    color: #00796b;
    border: none;
    border-right: 1px solid rgba(0, 0, 0, 0.05);
    border-bottom: 1px solid rgba(0, 0, 0, 0.1);
    padding: 6px 8px;
    font-weight: bold;
}

QHeaderView::section:last {
    border-right: none;
}

QTableCornerButton::section {
    background-color: #eceff1;
    border: none;
    border-bottom: 1px solid rgba(0, 0, 0, 0.1);
}

/* ===== 消息框 ===== */
QMessageBox {
    background-color: #ffffff;
}

QMessageBox QLabel {
    color: #2c3e50;
    font-size: 13px;
}

QMessageBox QPushButton {
    min-width: 80px;
}

/* ===== 文件对话框 ===== */
QFileDialog {
    background-color: #ffffff;
}

/* ===== 工具提示 ===== */
QToolTip {
    background-color: #ffffff;
    color: #2c3e50;
    border: 1px solid rgba(0, 150, 136, 0.3);
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
}

/* ===== 分割器 ===== */
QSplitter::handle {
    background-color: rgba(0, 0, 0, 0.05);
    width: 3px;
    height: 3px;
    border-radius: 1px;
}

QSplitter::handle:hover {
    background-color: #009688;
}

/* ===== 检查字数选择栏（claim 检查 Tab） ===== */
QFrame#claimNBar {
    background-color: #ffffff;
    border: 1px solid rgba(0, 0, 0, 0.12);
    border-radius: 18px;
    padding: 4px 10px;
}

QFrame#claimNBar QLabel {
    color: #546e7a;
    font-size: 13px;
    padding: 0 2px;
}

QPushButton#nPresetBtn {
    min-width: 34px;
    max-width: 34px;
    min-height: 26px;
    max-height: 26px;
    padding: 0;
    font-size: 13px;
    font-weight: 600;
    color: #455a64;
    background-color: #f5f7fa;
    border: 1px solid rgba(0, 0, 0, 0.12);
    border-radius: 13px;
}

QPushButton#nPresetBtn:hover {
    background-color: #e0f2f1;
    border-color: #00bfa5;
    color: #00796b;
}

QPushButton#nPresetBtn:checked {
    background-color: #00bfa5;
    border-color: #00bfa5;
    color: #ffffff;
}

QSpinBox#nCustomSpin {
    min-height: 28px;
    padding: 2px 10px;
    font-size: 13px;
    font-weight: 600;
    color: #263238;
    background-color: #ffffff;
    border: 1px solid rgba(0, 0, 0, 0.20);
    border-radius: 13px;
    selection-background-color: #00bfa5;
    selection-color: #ffffff;
}

QSpinBox#nCustomSpin:focus {
    border-color: #00bfa5;
}
"""
