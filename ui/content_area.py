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

阶段二·错别字/重复字内联高亮：`highlight_issues` 检查后把所有问题文字在对应标签页
  标黄（仅覆盖 `wrong` 字符段，行=段 → 段内偏移==行内偏移，靠 typo 的 `occurrence`/
  dup 首次出现反推偏移）；`locate_issue` 单击「修改前」时跳转到该段并把这一条改红强调。

阶段三·内联富显示：含图片/公式的标签页改用富文本构建（`_build_rich`）——按段内文档顺序
  插入文本与图片（附图 PNG/JPEG、公式 WMF/EMF 预览经 GDI 转 QImage），**每段仍是一个 block、
  内联对象是块内 U+FFFC 不增行**，故行=段/回写/高亮模型不变；含对象段沿用 `_has_image` 只读跳过
  回写。无对象的标签页保持 `setPlainText` 快速可编辑路径。
"""
from PyQt6.QtCore import Qt, QTimer, QUrl, pyqtSignal
from PyQt6.QtGui import (
    QTextCursor, QColor, QTextCharFormat, QTextFormat, QTextDocument,
    QTextImageFormat,
)
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QTabWidget, QTextEdit, QPushButton,
)

from ui.render.media import has_renderable_object, iter_content, scale_to_width

from core.doc_parser import _has_image
from core.paragraph_edit import set_paragraph_text

# 「说明书」合集所含子章节（按文档常规顺序）
_SPEC_ORDER = ["技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式"]


class ContentArea(QWidget):
    """常驻内容区：专利内容 4 标签页（结构化可编辑）+ 操作日志 / 操作历史 工具页。"""

    TAB_NAMES = ["权利要求书", "说明书", "说明书附图", "说明书摘要"]
    HIGHLIGHT = "#ffd966"          # 双击定位行高亮色（沿用权项预览框配色）
    ISSUE_YELLOW = "#ffd54f"       # 错别字/重复字：全部检查项标黄
    ISSUE_RED = "#ff5252"          # 错别字/重复字：当前点击项标红（配白色前景）

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
        # 错别字/重复字内联高亮：每个标签页一组 (line, offset, length)；当前点击项单独记
        self._issue_ranges: list[list[tuple]] = [[] for _ in self.TAB_NAMES]
        self._active_issue: tuple | None = None   # (tab, line, offset, length)
        # 阶段三富显示：图片/公式预览 QImage 缓存（rId→QImage，含负缓存）；
        # 哪些标签页走了富文本（含对象）以便窗口缩放时按宽重排
        self._img_cache: dict = {}
        self._rich_tabs: set[int] = set()
        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(200)
        self._resize_timer.timeout.connect(self._relayout_rich_tabs)

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
        self._issue_ranges = [[] for _ in self.TAB_NAMES]
        self._active_issue = None
        self._rich_tabs = set()
        self._img_cache = {}
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

        document = doc_data.get("document")
        for i, pmap in enumerate(self._para_maps):
            self._fill_tab(i, pmap, paras, document)

    def _fill_tab(self, i: int, pmap: list, paras: list, document) -> None:
        """含图片/公式的标签页走富文本，否则走纯文本快速路径。"""
        edit = self._edits[i]
        needs_rich = bool(pmap) and document is not None and any(
            has_renderable_object(paras[idx], document, self._img_cache) for idx in pmap
        )
        self._suppress = True
        if needs_rich:
            self._build_rich(edit, pmap, paras, document)
            self._rich_tabs.add(i)
        else:
            edit.setPlainText("\n".join(paras[idx].text for idx in pmap))
            self._rich_tabs.discard(i)
        self._suppress = False
        edit.setReadOnly(not pmap)   # 仅在该标签页有内容时开放编辑

    def _build_rich(self, edit: QTextEdit, pmap: list, paras: list, document) -> None:
        """逐段构建富文本：每段一个 block，段内按文档顺序插文本/内联图片（不增行）。"""
        edit.clear()
        doc = edit.document()
        max_w = max(64, edit.viewport().width() - 24)
        cursor = QTextCursor(doc)
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        for n, idx in enumerate(pmap):
            if n > 0:
                cursor.insertBlock()
            for kind, payload in iter_content(paras[idx], document, self._img_cache):
                if kind == "text":
                    if payload:
                        cursor.insertText(payload)
                elif kind == "placeholder":
                    cursor.insertText(payload)
                elif kind == "image":
                    img = scale_to_width(payload, max_w)
                    name = f"mem://{id(payload)}"
                    doc.addResource(QTextDocument.ResourceType.ImageResource,
                                    QUrl(name), img)
                    fmt = QTextImageFormat()
                    fmt.setName(name)
                    fmt.setWidth(img.width())
                    fmt.setHeight(img.height())
                    cursor.insertImage(fmt)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        # 含图片/公式的标签页防抖重排，使附图/公式随 1框 宽度自适应
        if self._rich_tabs:
            self._resize_timer.start()

    def _relayout_rich_tabs(self):
        """窗口缩放后按新宽度重建富文本标签页（QImage 走缓存，不重解码/重渲染 GDI）。"""
        if self._doc_data is None or not self._rich_tabs:
            return
        paras = self._doc_data.get("paragraphs", [])
        document = self._doc_data.get("document")
        if document is None:
            return
        for i in list(self._rich_tabs):
            pmap = self._para_maps[i]
            if not pmap:
                continue
            self._suppress = True
            self._build_rich(self._edits[i], pmap, paras, document)
            self._suppress = False
        # 重排后此前的内联高亮失效（block 重建）→ 清掉，避免错位
        self.clear_issue_highlights()

    # ===== 编辑 → 手动确认回写 =====
    def _on_text_changed(self, tab: int):
        if self._suppress:
            return
        self._dirty.add(tab)
        self._update_confirm_enabled()
        # 编辑后字符偏移失效 → 清除内联高亮，避免标错位置
        if self._issue_ranges[tab] or (self._active_issue and self._active_issue[0] == tab):
            self.clear_issue_highlights()

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
                # （内联对象在 QTextEdit 里是 U+FFFC 占位符，比较时先剔除，免误报）
                if line.replace("￼", "") != para.text:
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

    # ===== 错别字 / 重复字 内联高亮 =====
    def highlight_issues(self, results: list) -> None:
        """检查后把当前结果集里所有问题文字在对应标签页标黄（仅覆盖 wrong 字符段）。

        results 为 typo / dup 结果列表，元素含 para_idx / wrong / occurrence(可选)。
        反查或偏移失败的条目静默跳过；本调用整体重建高亮（清掉上一轮）。
        """
        self._issue_ranges = [[] for _ in self.TAB_NAMES]
        self._active_issue = None
        for item in results or []:
            r = self._range_for_issue(item)
            if r is None:
                continue
            tab, line, offset, length = r
            self._issue_ranges[tab].append((line, offset, length))
        for tab in range(len(self._edits)):
            self._apply_issue_selections(tab)

    def locate_issue(self, para_idx: int, wrong: str, occurrence: int = 1) -> bool:
        """单击「修改前」：跳转到该问题所在段并把这一条标红强调（叠在黄底之上）。"""
        r = self._range_for_issue(
            {"para_idx": para_idx, "wrong": wrong, "occurrence": occurrence}
        )
        if r is None:
            return False
        tab, line, offset, length = r
        self._active_issue = (tab, line, offset, length)
        self.tabs.setCurrentIndex(tab)
        ed = self._edits[tab]
        block = ed.document().findBlockByNumber(line)
        if block.isValid():
            cur = QTextCursor(block)
            cur.setPosition(block.position() + offset)
            ed.setTextCursor(cur)
            ed.ensureCursorVisible()
        self._apply_issue_selections(tab)
        return True

    def highlight_claim_issues(self, items: list) -> None:
        """权项专用标黄：items 元素为 {para_idx, wrong}。

        与 typo/dup 不同——权项的 para_idx 是「权项首段」，术语可能在该权项的后续
        段落，故从首段所在行起**向后**找术语首次出现（行=段、各段为连续行）。
        """
        self._issue_ranges = [[] for _ in self.TAB_NAMES]
        self._active_issue = None
        for it in items or []:
            r = self._range_for_issue_forward(it.get("para_idx", -1), it.get("wrong") or "")
            if r is None:
                continue
            tab, line, offset, length = r
            self._issue_ranges[tab].append((line, offset, length))
        for tab in range(len(self._edits)):
            self._apply_issue_selections(tab)

    def locate_claim_issue(self, para_idx: int, wrong: str) -> bool:
        """权项单击「说明」：从首段起向后找术语，跳转并标红。"""
        r = self._range_for_issue_forward(para_idx, wrong)
        if r is None:
            return False
        tab, line, offset, length = r
        self._active_issue = (tab, line, offset, length)
        self.tabs.setCurrentIndex(tab)
        ed = self._edits[tab]
        block = ed.document().findBlockByNumber(line)
        if block.isValid():
            cur = QTextCursor(block)
            cur.setPosition(block.position() + offset)
            ed.setTextCursor(cur)
            ed.ensureCursorVisible()
        self._apply_issue_selections(tab)
        return True

    def clear_issue_highlights(self) -> None:
        """清除全部内联高亮（黄+红）。"""
        self._issue_ranges = [[] for _ in self.TAB_NAMES]
        self._active_issue = None
        for ed in self._edits:
            ed.setExtraSelections([])

    def _range_for_issue_forward(self, para_idx: int, wrong: str):
        """从 para_idx 所在行起，向后逐行找 wrong 首次出现；返回 (tab,line,offset,len)。"""
        if para_idx is None or para_idx < 0 or not wrong:
            return None
        tl = self._find_tab_line(para_idx)
        if tl is None:
            return None
        tab, start_line = tl
        doc = self._edits[tab].document()
        total = doc.blockCount()
        for line in range(start_line, total):
            block = doc.findBlockByNumber(line)
            if not block.isValid():
                break
            off = block.text().find(wrong)
            if off >= 0:
                return tab, line, off, len(wrong)
        return None

    def _range_for_issue(self, item: dict):
        """把一条结果解析为 (tab, line, offset, length)；失败返回 None。"""
        para_idx = item.get("para_idx", -1)
        wrong = item.get("wrong") or ""
        if para_idx is None or para_idx < 0 or not wrong:
            return None
        tl = self._find_tab_line(para_idx)
        if tl is None:
            return None
        tab, line = tl
        block = self._edits[tab].document().findBlockByNumber(line)
        if not block.isValid():
            return None
        offset = self._nth_occurrence(block.text(), wrong, int(item.get("occurrence", 1) or 1))
        if offset < 0:
            return None
        return tab, line, offset, len(wrong)

    def _apply_issue_selections(self, tab: int) -> None:
        """按 _issue_ranges[tab] 铺黄底，再把 _active_issue（若属本 tab）铺红底。"""
        ed = self._edits[tab]
        sels = []
        for (line, offset, length) in self._issue_ranges[tab]:
            sel = self._make_selection(ed, line, offset, length, self.ISSUE_YELLOW)
            if sel is not None:
                sels.append(sel)
        if self._active_issue and self._active_issue[0] == tab:
            _t, line, offset, length = self._active_issue
            sel = self._make_selection(
                ed, line, offset, length, self.ISSUE_RED, fg="#ffffff"
            )
            if sel is not None:
                sels.append(sel)
        ed.setExtraSelections(sels)

    def _make_selection(self, ed: QTextEdit, line: int, offset: int, length: int,
                        bg: str, fg: str | None = None):
        """造一个覆盖某行 [offset, offset+length) 字符段的 ExtraSelection。"""
        block = ed.document().findBlockByNumber(line)
        if not block.isValid():
            return None
        sel = QTextEdit.ExtraSelection()
        fmt = QTextCharFormat()
        fmt.setBackground(QColor(bg))
        if fg:
            fmt.setForeground(QColor(fg))
        sel.format = fmt
        cur = QTextCursor(block)
        cur.setPosition(block.position() + offset)
        cur.setPosition(
            block.position() + offset + length, QTextCursor.MoveMode.KeepAnchor
        )
        sel.cursor = cur
        return sel

    def _find_tab_line(self, para_idx: int):
        """反查 para_idx 所在的 (tab, line)；不在任何标签页返回 None。"""
        for tab, pmap in enumerate(self._para_maps):
            try:
                return tab, pmap.index(para_idx)
            except ValueError:
                continue
        return None

    @staticmethod
    def _nth_occurrence(text: str, needle: str, n: int) -> int:
        """返回 needle 在 text 中第 n 次出现（1-indexed）的偏移；找不到返回 -1。"""
        if not text or not needle or n < 1:
            return -1
        idx = -1
        for _ in range(n):
            idx = text.find(needle, idx + 1)
            if idx == -1:
                return -1
        return idx

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
