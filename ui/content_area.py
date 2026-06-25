# -*- coding: utf-8 -*-
"""
ui/content_area.py — 左上常驻内容区（主舞台）

标签页：
  权利要求书 / 说明书 / 说明书附图 / 说明书摘要 — 专利内容（「说明书」是技术领域 +
  背景技术 + 发明内容 + 附图说明 + 具体实施方式的合集）；
  操作日志 / 操作历史 — 由主窗口写入的工具页（各模块的日志与历史统一汇总到这里，
  以腾空下方操作区、让专利内容占据主要空间）。

阶段二·增量1：专利内容只读显示。
阶段二·增量2（任务6）：升级为**结构化实时编辑**——
  · 行=段：每个标签页里一行文本对应内存中的一个 docx 段落（`_para_maps`）；
  · 行数守恒：编辑禁止增删行（=增删段），否则拒绝回写并提示（守住 para_idx 红线）；
  · 手动确认回写：标签栏右上角「✓ 确认修改」按钮（有未保存编辑时才启用），点击后用
    `core.paragraph_edit.set_paragraph_text` 逐段写回内存（只改 `<w:t>`、保留公式/图片
    等非文本节点）；检查 / 导出前主窗口也会调用 `flush_all()` 兜底回写；
  · 含图片/公式的段（`doc_parser._has_image`）只读，不参与回写，被改动时给出提示；
  · 双击定位高亮：`locate_paragraph` 按 para_idx 反查行号并整行高亮（供权项结果双击跳转）。
"""
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QTextCursor, QColor, QTextCharFormat, QTextFormat
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QTabWidget, QTextEdit, QPushButton,
)

from core.doc_parser import _has_image
from core.paragraph_edit import set_paragraph_text

# 「说明书」合集所含子章节（按文档常规顺序）
_SPEC_ORDER = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]


class ContentArea(QWidget):
    """常驻内容区：专利内容 4 标签页（结构化可编辑）+ 操作日志 / 操作历史 工具页。"""

    TAB_NAMES = ["权利要求书", "说明书", "说明书附图", "说明书摘要"]
    HIGHLIGHT = "#ffd966"          # 双击定位行高亮色（沿用权项预览框配色）

    # 任一专利段落被回写到内存后发出（主窗口据此让检查缓存失效）
    contentEdited = pyqtSignal()
    # 编辑被拒绝 / 只读段被改动时发出，主窗口转成 toast 提示
    editWarning = pyqtSignal(str)
    # 「确认修改」点击后发出，携带本次实际写回内存的段落数（主窗口转成 toast）
    editConfirmed = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("contentArea")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("contentTabs")

        # 标签栏右上角：「确认修改」按钮（有未保存编辑时才启用）
        self._confirm_btn = QPushButton("✓ 确认修改")
        self._confirm_btn.setObjectName("confirmEditBtn")
        self._confirm_btn.setToolTip(
            "把当前对专利文本的编辑写入内存（行数须与原文一致，含图片/公式的段只读）"
        )
        self._confirm_btn.setEnabled(False)
        self._confirm_btn.clicked.connect(self._on_confirm_clicked)
        self.tabs.setCornerWidget(self._confirm_btn, Qt.Corner.TopRightCorner)

        self._doc_data: dict | None = None
        # 每个专利标签页：行号 → 全文段落索引 的映射（行=段）
        self._para_maps: list[list[int]] = [[] for _ in self.TAB_NAMES]
        self._suppress = False           # 程序性填充时抑制 textChanged
        self._dirty: set[int] = set()    # 待回写的标签页索引

        # —— 专利内容标签页（可编辑） ——
        self._edits: list[QTextEdit] = []
        for i, name in enumerate(self.TAB_NAMES):
            edit = QTextEdit()
            edit.setReadOnly(True)       # load 前无内容 → 只读，load 后开放编辑
            edit.setObjectName("contentEdit")
            edit.setPlaceholderText("加载文档后在此显示专利内容，可直接编辑纯文本段…")
            edit.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
            edit.textChanged.connect(lambda _i=i: self._on_text_changed(_i))
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

    # ===== 加载 / 清空 =====
    def clear(self):
        self._doc_data = None
        self._dirty.clear()
        for i, e in enumerate(self._edits):
            self._para_maps[i] = []
            self._suppress = True
            e.clear()
            e.setExtraSelections([])
            e.setReadOnly(True)
            self._suppress = False
        self._update_confirm_enabled()

    def load(self, doc_data: dict):
        """从 doc_data 填充 4 个专利内容标签页，并建立行=段映射（日志 / 历史不在此清空）。"""
        self.clear()
        if not doc_data:
            return
        self._doc_data = doc_data
        sections = doc_data.get("sections", {})
        paras = doc_data.get("paragraphs", [])

        # 每个标签页对应的全文段落索引列表（行=段）
        self._para_maps[0] = self._range_of(sections, ["权利要求书"])
        self._para_maps[1] = self._spec_indices(sections)
        self._para_maps[2] = self._range_of(sections, ["说明书附图"])
        self._para_maps[3] = self._range_of(sections, ["说明书摘要", "摘要附图"])

        for i, pmap in enumerate(self._para_maps):
            edit = self._edits[i]
            self._suppress = True
            edit.setPlainText("\n".join(paras[idx].text for idx in pmap))
            self._suppress = False
            edit.setReadOnly(not pmap)   # 仅在该标签页有内容时开放编辑

    # ===== 编辑 → 手动确认回写 =====
    def _on_text_changed(self, tab: int):
        if self._suppress:
            return
        self._dirty.add(tab)
        self._update_confirm_enabled()

    def _update_confirm_enabled(self):
        if hasattr(self, "_confirm_btn"):
            self._confirm_btn.setEnabled(bool(self._dirty))

    def _on_confirm_clicked(self):
        """「✓ 确认修改」：把所有未保存编辑写回内存。"""
        if not self._dirty:
            self.editConfirmed.emit(0)
            return
        changed, rejected = self.flush_all()
        if rejected:
            # 行数不守恒已由 editWarning 提示，保持 dirty 让用户修正后再确认
            return
        self.editConfirmed.emit(changed)

    def flush_all(self):
        """把所有待回写的标签页写回内存（检查 / 导出前也会调用，确保内存最新）。

        返回 (本次实际写回的段落数, 因行数不守恒被拒的标签页数)。
        """
        total_changed = 0
        rejected = 0
        for tab in list(self._dirty):
            changed, ok = self._flush(tab)
            total_changed += changed
            if not ok:
                rejected += 1
        self._update_confirm_enabled()
        return total_changed, rejected

    def _flush(self, tab: int):
        """把第 tab 个标签页的文本逐段回写到内存（行数守恒；只读段跳过）。

        返回 (写回段落数, 是否成功)；行数不守恒返回 (0, False) 且保留 dirty。
        """
        if self._doc_data is None:
            self._dirty.discard(tab)
            return 0, True
        pmap = self._para_maps[tab]
        if not pmap:
            self._dirty.discard(tab)
            return 0, True
        paras = self._doc_data.get("paragraphs", [])
        lines = self._edits[tab].toPlainText().split("\n")
        if len(lines) != len(pmap):
            self.editWarning.emit(
                f"「{self.TAB_NAMES[tab]}」结构化编辑禁止增删段落"
                f"（当前 {len(lines)} 行 / 应为 {len(pmap)} 行）；"
                "请撤销增删行操作（Ctrl+Z）后再点确认，本次改动未保存。"
            )
            return 0, False

        changed = 0
        readonly_touched = False
        for line, idx in zip(lines, pmap):
            para = paras[idx]
            if _has_image(para):
                # 含图片/公式段只读：不回写；若被改动则提示
                if line != para.text:
                    readonly_touched = True
                continue
            if set_paragraph_text(para, line):
                changed += 1

        self._dirty.discard(tab)
        if readonly_touched:
            self.editWarning.emit("含图片/公式的段落为只读，对其的改动不会被保存。")
        if changed:
            self.contentEdited.emit()
        return changed, True

    # ===== 双击定位高亮 =====
    def locate_paragraph(self, tab: int, para_idx: int, search_key: str = "") -> bool:
        """切到第 tab 个标签页，按段落索引反查行号并整行高亮（供结果双击跳转）。

        para_idx 不在该标签页时，回退为按 search_key 文本查找。
        """
        if tab < 0 or tab >= len(self._edits):
            return False
        pmap = self._para_maps[tab]
        try:
            line = pmap.index(para_idx)
        except ValueError:
            return self._locate_by_text(tab, search_key)

        self.tabs.setCurrentIndex(tab)
        ed = self._edits[tab]
        block = ed.document().findBlockByNumber(line)
        if not block.isValid():
            return self._locate_by_text(tab, search_key)
        cur = QTextCursor(block)
        ed.setTextCursor(cur)
        ed.ensureCursorVisible()
        self._highlight_block(ed, block)
        return True

    def locate_in_claims(self, search_key: str) -> bool:
        """兼容旧接口：在「权利要求书」标签页按文本查找 / 高亮。"""
        return self._locate_by_text(0, search_key)

    def _locate_by_text(self, tab: int, search_key: str) -> bool:
        if tab < 0 or tab >= len(self._edits) or not search_key:
            return False
        self.tabs.setCurrentIndex(tab)
        ed = self._edits[tab]
        cur = ed.textCursor()
        cur.movePosition(QTextCursor.MoveOperation.Start)
        ed.setTextCursor(cur)
        found = ed.find(search_key)
        if found:
            ed.ensureCursorVisible()
        return found

    def _highlight_block(self, ed: QTextEdit, block) -> None:
        sel = QTextEdit.ExtraSelection()
        fmt = QTextCharFormat()
        fmt.setBackground(QColor(self.HIGHLIGHT))
        fmt.setProperty(QTextFormat.Property.FullWidthSelection, True)
        sel.format = fmt
        sel.cursor = QTextCursor(block)
        ed.setExtraSelections([sel])

    # ── 内部：构建行=段映射 ──
    def _range_of(self, sections: dict, names: list[str]) -> list[int]:
        """按给定章节名顺序，收集其段落区间的全文段落索引（含空段，保证行=段）。"""
        out: list[int] = []
        for n in names:
            sec = sections.get(n)
            if sec is not None:
                out.extend(range(sec.start_idx, sec.end_idx))
        return out

    def _spec_indices(self, sections: dict) -> list[int]:
        """说明书合集：按文档顺序拼接存在的子章节段落区间（章节标题段本就在区间内，充当小标题）。"""
        present = sorted(
            (sections[n].start_idx, n) for n in _SPEC_ORDER if n in sections
        )
        out: list[int] = []
        for _idx, name in present:
            sec = sections[name]
            out.extend(range(sec.start_idx, sec.end_idx))
        return out
