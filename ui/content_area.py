# -*- coding: utf-8 -*-
"""
ui/content_area.py — 左上常驻专利内容区

上方 4 个标签页：权利要求书 / 说明书 / 说明书附图 / 说明书摘要。
其中「说明书」是技术领域 + 背景技术 + 发明内容 + 附图说明 + 具体实施方式的合集。

阶段二·增量1：只读显示（基于 get_section_text 的纯文本）。
阶段二·增量2（任务6）：升级为结构化可编辑（行=段、行数守恒回写、双击定位高亮），
                        并按 doc_parser._has_image 把含图片/公式的段置为只读。
"""
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTabWidget, QTextEdit

from core.doc_parser import get_section_text

# 「说明书」合集所含子章节（按文档常规顺序）
_SPEC_ORDER = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]


class ContentArea(QWidget):
    """常驻内容区：4 标签页显示专利各部分。"""

    TAB_NAMES = ["权利要求书", "说明书", "说明书附图", "说明书摘要"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("contentArea")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("contentTabs")
        self._edits: list[QTextEdit] = []
        for name in self.TAB_NAMES:
            edit = QTextEdit()
            edit.setReadOnly(True)
            edit.setObjectName("contentEdit")
            edit.setPlaceholderText("加载文档后在此显示专利内容…")
            self._edits.append(edit)
            self.tabs.addTab(edit, name)
        layout.addWidget(self.tabs)

    def clear(self):
        for e in self._edits:
            e.clear()

    def load(self, doc_data: dict):
        """从 doc_data 填充 4 个标签页。"""
        if not doc_data:
            self.clear()
            return
        sections = doc_data.get("sections", {})
        paras = doc_data.get("paragraphs", [])
        self._edits[0].setPlainText(self._join(paras, sections, ["权利要求书"]))
        self._edits[1].setPlainText(self._spec(paras, sections))
        self._edits[2].setPlainText(self._join(paras, sections, ["说明书附图"]))
        self._edits[3].setPlainText(self._join(paras, sections, ["说明书摘要", "摘要附图"]))

    # ── 内部 ──
    @staticmethod
    def _join(paras, sections, names) -> str:
        blocks = []
        for n in names:
            sec = sections.get(n)
            if sec is not None:
                text = get_section_text(paras, sec)
                if text:
                    blocks.append(text)
        return "\n\n".join(blocks)

    @staticmethod
    def _spec(paras, sections) -> str:
        """说明书合集：按文档顺序拼接存在的子章节，每节带【名称】小标题。"""
        present = sorted(
            (sections[n].start_idx, n) for n in _SPEC_ORDER if n in sections
        )
        blocks = []
        for _idx, name in present:
            text = get_section_text(paras, sections[name])
            blocks.append(f"【{name}】\n{text}" if text else f"【{name}】")
        return "\n\n".join(blocks)
