# -*- coding: utf-8 -*-
"""
ui/nav_panel.py — 右侧两列导航（master-detail）

第一列（固定）：操作模块（标记 / 清洗 / 错别字 / 权项）。
第二列（动态）：随第一列选择变化，列出该模块的子功能。
选中第二列某子功能 → 发 page_selected(page_index)，由主窗口切换左下 QStackedWidget，
左下只显示该单个子功能。
"""
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QWidget, QHBoxLayout, QListWidget, QListWidgetItem

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
    """两列导航：col1 模块（固定） + col2 子功能（动态）。"""

    page_selected = pyqtSignal(int)

    def __init__(self, modules, parent=None):
        """modules: [(module_label, [(sub_label, page_index), ...]), ...]"""
        super().__init__(parent)
        self.setObjectName("navPanel")
        self.setStyleSheet(_QSS)
        self._modules = modules

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.col1 = QListWidget()
        self.col1.setObjectName("navCol1")
        self.col1.setFixedWidth(78)
        for label, _subs in modules:
            self.col1.addItem(label)
        layout.addWidget(self.col1)

        self.col2 = QListWidget()
        self.col2.setObjectName("navCol2")
        self.col2.setMinimumWidth(150)
        layout.addWidget(self.col2, 1)

        self.col1.currentRowChanged.connect(self._on_module_changed)
        self.col2.currentRowChanged.connect(self._on_sub_changed)
        self.col1.setCurrentRow(0)  # 触发填充 col2 并选中首个子功能

    def _on_module_changed(self, row: int):
        if row < 0 or row >= len(self._modules):
            return
        self.col2.clear()
        for sub_label, page_index in self._modules[row][1]:
            item = QListWidgetItem(sub_label)
            item.setData(_PAGE_ROLE, page_index)
            self.col2.addItem(item)
        if self.col2.count():
            self.col2.setCurrentRow(0)  # 触发 _on_sub_changed → 切页

    def _on_sub_changed(self, row: int):
        if row < 0:
            return
        item = self.col2.item(row)
        if item is not None:
            self.page_selected.emit(int(item.data(_PAGE_ROLE)))
