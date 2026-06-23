# -*- coding: utf-8 -*-
"""
ui/content_area.py — 左上常驻内容区（主舞台）

标签页：
  权利要求书 / 说明书 / 说明书附图 / 说明书摘要 — 专利内容（「说明书」是技术领域 +
  背景技术 + 发明内容 + 附图说明 + 具体实施方式的合集）；
  操作日志 / 操作历史 — 由主窗口写入的工具页（各模块的日志与历史统一汇总到这里，
  以腾空下方操作区、让专利内容占据主要空间）。

阶段二·增量1：专利内容只读显示。
阶段二·增量2（任务6）：升级为结构化可编辑（行=段、行数守恒回写、双击定位高亮），
                        并按 doc_parser._has_image 把含图片/公式的段置为只读。
"""
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTabWidget, QTextEdit

from core.doc_parser import get_section_text

# 「说明书」合集所含子章节（按文档常规顺序）
_SPEC_ORDER = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]


class ContentArea(QWidget):
    """常驻内容区：专利内容 4 标签页 + 操作日志 / 操作历史 工具页。"""

    TAB_NAMES = ["权利要求书", "说明书", "说明书附图", "说明书摘要"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("contentArea")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("contentTabs")

        # —— 专利内容标签页 ——
        self._edits: list[QTextEdit] = []
        for name in self.TAB_NAMES:
            edit = QTextEdit()
            edit.setReadOnly(True)
            edit.setObjectName("contentEdit")
            edit.setPlaceholderText("加载文档后在此显示专利内容…")
            self._edits.append(edit)
            self.tabs.addTab(edit, name)

        # —— 工具标签页：操作日志 / 操作历史 ——
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setObjectName("contentEdit")
        self.log_edit.setPlaceholderText("各模块的操作日志将在此统一显示…")
        self.tabs.addTab(self.log_edit, "📋 操作日志")

        self.history_edit = QTextEdit()
        self.history_edit.setReadOnly(True)
        self.history_edit.setObjectName("contentEdit")
        self.history_edit.setPlaceholderText(
            "尚未对当前文档进行修改。\n"
            "所有「标注 / 删除标记 / 清洗 / 错别字修正」等操作都会先记录在这里，\n"
            "待确认后点右侧「💾 文件生成」一次性写入新 docx。"
        )
        self.tabs.addTab(self.history_edit, "📜 操作历史")

        layout.addWidget(self.tabs)

    def clear(self):
        for e in self._edits:
            e.clear()

    def load(self, doc_data: dict):
        """从 doc_data 填充 4 个专利内容标签页（日志 / 历史不在此清空）。"""
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
