# -*- coding: utf-8 -*-
"""
ui/nav_panel.py — 右侧两列导航（master-detail）

第一列（3列，固定）：操作模块（标记 / 清洗 / 错别字 / 权项）。
第二列（4列，动态）：随第一列选择变化，内容有两种形态——
  · 列表型模块：列出该模块的子功能；选中子功能 → 发 page_selected(page_index)，
    由主窗口切换左下 QStackedWidget（2框），只显示该单个子功能。
  · 控件型模块（如「标记」）：第二列直接放该模块的操作按钮组；选中该模块即
    发 page_selected(其面板页)，左下 2框 同步显示该模块面板（如附图标记字典）。
"""
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QListWidget, QListWidgetItem, QStackedWidget,
)

_QSS = """
QListWidget {
    border: 1px solid rgba(120,120,120,0.25);
    border-radius: 8px;
    background: transparent;
    outline: 0;
    padding: 2px;
}
QListWidget::item {
    padding: 7px 6px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background: rgba(58,142,230,0.20);
    border: 1px solid #3a8ee6;
    color: palette(text);
}
QListWidget::item:hover { background: rgba(58,142,230,0.10); }
"""

# 第二列项目里存放目标页码的角色
_PAGE_ROLE = Qt.ItemDataRole.UserRole


class NavPanel(QWidget):
    """两列导航：col1 模块（固定，3列） + col2 子功能 / 操作按钮（动态，4列）。"""

    page_selected = pyqtSignal(int)

    def __init__(self, modules, parent=None):
        """modules: 每项描述一个固定模块及其第二列（4列）内容，两种形态：
          (label, [(sub_label, page_index), ...])    —— 第二列为子功能列表；
          (label, widget, page_index)                —— 第二列为自定义控件（如标记
                                                         模块的操作按钮组），选中该模块
                                                         即切到 page_index。
        """
        super().__init__(parent)
        self.setObjectName("navPanel")
        self.setStyleSheet(_QSS)
        self._modules = modules

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # 第一列（3列）：固定模块
        self.col1 = QListWidget()
        self.col1.setObjectName("navCol1")
        self.col1.setFixedWidth(78)
        for entry in modules:
            self.col1.addItem(entry[0])
        layout.addWidget(self.col1)

        # 第二列（4列）：用 QStackedWidget 容纳「共用子功能列表」与各控件型模块的自定义控件
        self.col2_stack = QStackedWidget()
        self.col2_stack.setObjectName("navCol2")
        self.col2_stack.setMinimumWidth(150)
        layout.addWidget(self.col2_stack, 1)

        # stack 第 0 页：被所有「列表型」模块共用的子功能列表
        self.col2_list = QListWidget()
        self.col2_list.currentRowChanged.connect(self._on_sub_changed)
        self.col2_stack.addWidget(self.col2_list)

        # 为每个「控件型」模块登记其自定义控件页与目标面板页
        self._widget_pages: dict[int, tuple[int, int]] = {}  # 模块行 -> (stack 页, 面板页)
        for row, entry in enumerate(modules):
            if isinstance(entry[1], QWidget):
                widget, page_index = entry[1], entry[2]
                stack_index = self.col2_stack.addWidget(widget)
                self._widget_pages[row] = (stack_index, page_index)

        self.col1.currentRowChanged.connect(self._on_module_changed)
        self.col1.setCurrentRow(0)  # 触发填充 col2 并选中首项

    def _on_module_changed(self, row: int):
        if row < 0 or row >= len(self._modules):
            return
        # 控件型模块：第二列显示其自定义控件，并直接切到该模块面板页
        if row in self._widget_pages:
            stack_index, page_index = self._widget_pages[row]
            self.col2_stack.setCurrentIndex(stack_index)
            self.page_selected.emit(page_index)
            return
        # 列表型模块：填充共用列表并切到列表页
        self.col2_stack.setCurrentWidget(self.col2_list)
        self.col2_list.clear()
        for sub_label, page_index in self._modules[row][1]:
            item = QListWidgetItem(sub_label)
            item.setData(_PAGE_ROLE, page_index)
            self.col2_list.addItem(item)
        if self.col2_list.count():
            self.col2_list.setCurrentRow(0)  # 触发 _on_sub_changed → 切页

    def _on_sub_changed(self, row: int):
        if row < 0:
            return
        item = self.col2_list.item(row)
        if item is not None:
            self.page_selected.emit(int(item.data(_PAGE_ROLE)))
