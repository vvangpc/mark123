# -*- coding: utf-8 -*-
"""
ui/activity_bar.py — 右侧竖条模块切换条（activity bar）

竖排若干可选中按钮，点击发 switched(index) 信号，由主窗口连到
QStackedWidget.setCurrentIndex 切换左下的模块面板。
"""
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QButtonGroup

_QSS = """
#activityBtn {
    border: 1px solid rgba(120,120,120,0.25);
    border-radius: 8px;
    background: transparent;
    font-size: 11px;
    padding: 2px;
}
#activityBtn:hover { background: rgba(58,142,230,0.12); }
#activityBtn:checked {
    background: rgba(58,142,230,0.18);
    border: 1px solid #3a8ee6;
    font-weight: bold;
}
"""


class ActivityBar(QWidget):
    """右侧模块切换竖条。点击按钮发 switched(index)。"""

    switched = pyqtSignal(int)

    def __init__(self, items: list[tuple[str, str]], parent=None):
        """items: [(icon, label), ...]，顺序与 QStackedWidget 页码一致。"""
        super().__init__(parent)
        self.setObjectName("activityBar")
        self.setStyleSheet(_QSS)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 8, 4, 8)
        layout.setSpacing(8)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)

        for i, (icon, label) in enumerate(items):
            btn = QPushButton(f"{icon}\n{label}")
            btn.setObjectName("activityBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFixedSize(64, 58)
            if i == 0:
                btn.setChecked(True)
            btn.clicked.connect(lambda _c, idx=i: self.switched.emit(idx))
            self._group.addButton(btn, i)
            layout.addWidget(btn)

        layout.addStretch()

    def set_current(self, index: int):
        btn = self._group.button(index)
        if btn is not None:
            btn.setChecked(True)
