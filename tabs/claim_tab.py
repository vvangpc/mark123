# -*- coding: utf-8 -*-
"""
tabs/claim_tab.py — 权利要求书检查 Tab

以 Mixin 形式抽离自 main_window.py，由 MainWindow 继承。保持方法体内所有
self.xxx 引用不变，避免重写信号/状态。Mixin 依赖 MainWindow 提供：
  - self.doc_data / self.current_marks
  - self._show_toast / self._add_history
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QPlainTextEdit, QFrame, QSplitter, QGroupBox, QCheckBox,
    QSpinBox, QButtonGroup, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QDialog, QTextEdit,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextCursor, QTextCharFormat, QColor

from workers import _longest_nonspace_run


class ClaimTabMixin:
    """权利要求书检查 Tab 的方法集合，供 MainWindow 继承。"""

    def _create_claim_check_tab(self) -> QWidget:
        """创建「权利要求书检查」标签页"""
        widget = QWidget()
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(0, 8, 0, 0)
        outer.setSpacing(10)

        # ── 顶部工具栏 ──────────────────────────────────────
        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)

        # 检查字数选择栏：pill 式容器，左侧标签 + 6 个预设按钮 + 自定义框
        n_bar = QFrame()
        n_bar.setObjectName("claimNBar")
        n_bar_layout = QHBoxLayout(n_bar)
        n_bar_layout.setContentsMargins(10, 2, 10, 2)
        n_bar_layout.setSpacing(6)

        n_label = QLabel("检查字数")
        n_bar_layout.addWidget(n_label)

        self._claim_n_buttons = QButtonGroup(widget)
        self._claim_n_buttons.setExclusive(True)
        for val in (2, 3, 4, 5, 6, 7):
            b = QPushButton(str(val))
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setObjectName("nPresetBtn")
            if val == 2:
                b.setChecked(True)
            b.clicked.connect(lambda _=False, v=val: self._on_claim_n_preset(v))
            self._claim_n_buttons.addButton(b, val)
            n_bar_layout.addWidget(b)

        sep = QLabel("|")
        sep.setStyleSheet("color: rgba(128,128,128,0.4); padding: 0 4px;")
        n_bar_layout.addWidget(sep)

        n_bar_layout.addWidget(QLabel("自定义"))
        self.claim_n_custom = QSpinBox()
        self.claim_n_custom.setObjectName("nCustomSpin")
        self.claim_n_custom.setRange(2, 30)
        self.claim_n_custom.setValue(8)
        self.claim_n_custom.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # 去掉上下调节按钮，只保留输入框（仍可用滚轮 / 键盘 ↑↓ 调整）
        self.claim_n_custom.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.claim_n_custom.setToolTip(
            "自定义检查字数。数值变化会取消上方 6 个按钮的选中态，"
            "以此处填入的数值为准。"
        )
        self.claim_n_custom.valueChanged.connect(self._on_claim_n_custom_changed)
        n_bar_layout.addWidget(self.claim_n_custom)

        toolbar.addWidget(n_bar)
        toolbar.addSpacing(6)

        self.claim_check_btn = QPushButton("▶  开始检查")
        self.claim_check_btn.setObjectName("accentBtn")
        self.claim_check_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_check_btn.setEnabled(False)
        self.claim_check_btn.clicked.connect(self._on_claim_check_start)
        toolbar.addWidget(self.claim_check_btn)

        # ── 引用基础降噪：动态截断 / 动态回退 两个上下分布的勾选框 ──
        dyn_box = QFrame()
        dyn_box.setObjectName("dynBox")
        dyn_layout = QVBoxLayout(dyn_box)
        dyn_layout.setContentsMargins(6, 0, 6, 0)
        dyn_layout.setSpacing(2)

        # 第一行：动态截断
        trunc_row = QHBoxLayout()
        trunc_row.setContentsMargins(0, 0, 0, 0)
        trunc_row.setSpacing(4)
        self.claim_dyn_trunc_cb = QCheckBox()
        self.claim_dyn_trunc_cb.setChecked(False)
        self.claim_dyn_trunc_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        trunc_row.addWidget(self.claim_dyn_trunc_cb)
        # 可点击的标签：左键点击 → 弹介绍；[黑名单] → 打开黑名单词库
        self.claim_dyn_trunc_label = QLabel(
            '<a href="info" style="text-decoration:none;color:inherit;">动态截断</a>'
            '&nbsp;<a href="bl" style="text-decoration:none;color:#3a8ee6;">[黑名单]</a>'
        )
        self.claim_dyn_trunc_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_dyn_trunc_label.setToolTip(
            "点「动态截断」查看功能说明；点「[黑名单]」编辑边界词库"
        )
        self.claim_dyn_trunc_label.linkActivated.connect(self._on_dyn_trunc_link)
        trunc_row.addWidget(self.claim_dyn_trunc_label)
        trunc_row.addStretch()
        dyn_layout.addLayout(trunc_row)

        # 第二行：动态回退
        fb_row = QHBoxLayout()
        fb_row.setContentsMargins(0, 0, 0, 0)
        fb_row.setSpacing(4)
        self.claim_dyn_fb_cb = QCheckBox()
        self.claim_dyn_fb_cb.setChecked(False)
        self.claim_dyn_fb_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        fb_row.addWidget(self.claim_dyn_fb_cb)
        self.claim_dyn_fb_label = QLabel(
            '<a href="info" style="text-decoration:none;color:inherit;">动态回退</a>'
        )
        self.claim_dyn_fb_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_dyn_fb_label.setToolTip("点击查看功能说明")
        self.claim_dyn_fb_label.linkActivated.connect(self._on_dyn_fb_link)
        fb_row.addWidget(self.claim_dyn_fb_label)
        fb_row.addStretch()
        dyn_layout.addLayout(fb_row)

        toolbar.addWidget(dyn_box)

        # ── 不确定用语检查 / 术语不一致检查 上下分布的勾选框 ──
        check_box = QFrame()
        check_box.setObjectName("checkBox")
        check_layout = QVBoxLayout(check_box)
        check_layout.setContentsMargins(6, 0, 6, 0)
        check_layout.setSpacing(2)

        # 第一行：不确定用语检查（默认勾选，保持原有行为）
        vague_row = QHBoxLayout()
        vague_row.setContentsMargins(0, 0, 0, 0)
        vague_row.setSpacing(4)
        self.claim_vague_cb = QCheckBox()
        self.claim_vague_cb.setChecked(True)
        self.claim_vague_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        vague_row.addWidget(self.claim_vague_cb)
        self.claim_vague_label = QLabel(
            '<a href="info" style="text-decoration:none;color:inherit;">不确定用语检查</a>'
            '&nbsp;<a href="wb" style="text-decoration:none;color:#3a8ee6;">[词库]</a>'
        )
        self.claim_vague_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_vague_label.setToolTip(
            "点「不确定用语检查」查看功能说明；点「[词库]」编辑不确定用语词库"
        )
        self.claim_vague_label.linkActivated.connect(self._on_vague_link)
        vague_row.addWidget(self.claim_vague_label)
        vague_row.addStretch()
        check_layout.addLayout(vague_row)

        # 第二行：术语不一致检查（噪音较大，默认关闭）
        term_row = QHBoxLayout()
        term_row.setContentsMargins(0, 0, 0, 0)
        term_row.setSpacing(4)
        self.claim_term_cb = QCheckBox()
        self.claim_term_cb.setChecked(False)
        self.claim_term_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        term_row.addWidget(self.claim_term_cb)
        self.claim_term_label = QLabel(
            '<a href="info" style="text-decoration:none;color:inherit;">术语不一致检查</a>'
        )
        self.claim_term_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_term_label.setToolTip("点击查看功能说明")
        self.claim_term_label.linkActivated.connect(self._on_term_link)
        term_row.addWidget(self.claim_term_label)
        term_row.addStretch()
        check_layout.addLayout(term_row)

        toolbar.addWidget(check_box)

        toolbar.addStretch()

        self.claim_status_label = QLabel("请先打开 docx 文件")
        self.claim_status_label.setObjectName("subtitleLabel")
        toolbar.addWidget(self.claim_status_label)

        outer.addLayout(toolbar)

        # ── 中部：左预览（可编辑） / 右结果表（可拖动分隔）──
        self.claim_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.claim_splitter.setChildrenCollapsible(False)

        # 左侧：预览编辑 + 确认按钮
        left_panel = QGroupBox("📄 权利要求书（可编辑）")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setSpacing(6)

        self.claim_preview_edit = QPlainTextEdit()
        self.claim_preview_edit.setPlaceholderText(
            "打开文档后将在这里展示权利要求书。\n"
            "您可以直接编辑；修改后务必点下方的「✔ 确认修改」把改动写回内存，\n"
            "最终「💾 文件生成」时才会落盘到 .docx。"
        )
        self.claim_preview_edit.textChanged.connect(self._on_claim_text_changed)
        left_layout.addWidget(self.claim_preview_edit, 1)

        self.claim_confirm_btn = QPushButton("✔  确认修改")
        self.claim_confirm_btn.setObjectName("primaryBtn")
        self.claim_confirm_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_confirm_btn.setEnabled(False)
        self.claim_confirm_btn.setToolTip(
            "将预览框的当前内容写回内存中的权利要求书段落。\n"
            "段落数必须保持不变（每行对应一段），最终 .docx 会反映此处的修改。"
        )
        self.claim_confirm_btn.clicked.connect(self._on_claim_confirm_edits)
        left_layout.addWidget(self.claim_confirm_btn)

        self.claim_splitter.addWidget(left_panel)

        # 右侧：检查结果表
        right_panel = QGroupBox("🔍 检查结果")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setSpacing(6)

        right_hint = QLabel(
            "表中仅展示问题，不会写入最终文件；双击「上下文」可跳转到左侧预览框"
            "对应位置并高亮；双击「说明」可弹出完整描述；改完再次「开始检查」"
            "重扫，直到清零。"
        )
        right_hint.setObjectName("subtitleLabel")
        right_hint.setWordWrap(True)
        right_layout.addWidget(right_hint)

        self.claim_result_table = QTableWidget(0, 5)
        self.claim_result_table.setHorizontalHeaderLabels(
            ["类型", "权项", "上下文", "说明", "操作"]
        )
        h = self.claim_result_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        # 上下文：Interactive，用户可拖动它与「说明」之间的分界线
        h.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        # 说明：Stretch，自动填满剩余空间，右边界不可拖动，操作列始终可见
        h.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        # 操作：固定宽度，吸附在右侧，不会被遮挡
        h.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.claim_result_table.setColumnWidth(4, 76)
        self.claim_result_table.setColumnWidth(2, 240)
        h.setStretchLastSection(False)
        self.claim_result_table.setAlternatingRowColors(True)
        self.claim_result_table.verticalHeader().setVisible(False)
        # 行高下限 34px，避免"忽略"等按钮文字在压缩行高下被裁剪
        self.claim_result_table.verticalHeader().setDefaultSectionSize(34)
        self.claim_result_table.verticalHeader().setMinimumSectionSize(34)
        # 双击「上下文」格（列 2） → 跳转并高亮左侧预览框对应位置
        self.claim_result_table.cellDoubleClicked.connect(
            self._on_claim_result_double_clicked
        )
        right_layout.addWidget(self.claim_result_table, 1)

        self.claim_splitter.addWidget(right_panel)
        self.claim_splitter.setSizes([560, 640])

        outer.addWidget(self.claim_splitter, 1)
        return widget

    def _on_claim_n_preset(self, value: int):
        """点击 N 预设按钮：更新当前 N 值，并清空自定义框的"高亮态"。"""
        self._claim_n = int(value)
        # 自定义框的值不影响当前 N，但为 UX 一致把它设回与按钮相同的值
        self.claim_n_custom.blockSignals(True)
        self.claim_n_custom.setValue(int(value))
        self.claim_n_custom.blockSignals(False)

    def _on_claim_n_custom_changed(self, value: int):
        """自定义 N 字数变化：取消所有按钮选中，以自定义值为准。"""
        # 如果恰好是 2~7，且已经在对应按钮上，就让按钮保持选中态一致
        btn = self._claim_n_buttons.button(int(value))
        if btn is not None:
            btn.setChecked(True)
        else:
            # 清空独占组的选中态（QButtonGroup exclusive 不允许全部未选，
            # 所以先切 exclusive 为 False，清空，再切回 True）
            self._claim_n_buttons.setExclusive(False)
            for b in self._claim_n_buttons.buttons():
                b.setChecked(False)
            self._claim_n_buttons.setExclusive(True)
        self._claim_n = int(value)

    def _claim_tab_load_from_doc(self):
        """_load_document 成功后：把权利要求书章节填入预览框。"""
        if not self.doc_data:
            return
        # 新文档加载 → 清空本次会话的忽略记录
        self._claim_session_ignore = set()
        sections = self.doc_data.get('sections', {})
        section = sections.get('权利要求书')
        if section is None:
            self.claim_preview_edit.blockSignals(True)
            self.claim_preview_edit.setPlainText("")
            self.claim_preview_edit.blockSignals(False)
            self._claim_start_idx = None
            self._claim_end_idx = None
            self._claim_para_count = 0
            self._claim_dirty = False
            self._claim_loaded = False
            self.claim_check_btn.setEnabled(False)
            self.claim_confirm_btn.setEnabled(False)
            self.claim_status_label.setText("未识别到权利要求书章节")
            self.claim_result_table.setRowCount(0)
            self._claim_results = []
            return

        paragraphs = self.doc_data['paragraphs']
        self._claim_start_idx = section.start_idx
        self._claim_end_idx = section.end_idx
        self._claim_para_count = section.end_idx - section.start_idx

        # 每段一行（含空段）→ 行数严格等于 para_count
        lines = []
        for i in range(section.start_idx, section.end_idx):
            text = paragraphs[i].text if paragraphs[i].text else ""
            lines.append(text)
        content = "\n".join(lines)

        self.claim_preview_edit.blockSignals(True)
        self.claim_preview_edit.setPlainText(content)
        self.claim_preview_edit.blockSignals(False)

        self._claim_dirty = False
        self._claim_loaded = True
        self.claim_check_btn.setEnabled(True)
        self.claim_confirm_btn.setEnabled(False)
        self._claim_results = []
        self.claim_result_table.setRowCount(0)
        self.claim_status_label.setText(
            f"已加载权利要求书：共 {self._claim_para_count} 段  ·  点「开始检查」开始"
        )

    def _on_claim_text_changed(self):
        """预览框内容变化 → 标记 dirty，高亮确认按钮"""
        if not self._claim_loaded:
            return
        self._claim_dirty = True
        self.claim_confirm_btn.setEnabled(True)
        self._update_claim_status_bar()

    def _update_claim_status_bar(self):
        """刷新权利要求书 Tab 的状态栏文字"""
        if not self._claim_loaded:
            return
        n_results = len(self._claim_results)
        parts = [f"共 {self._claim_para_count} 段"]
        if n_results:
            parts.append(f"问题 {n_results} 条")
        if self._claim_dirty:
            parts.append("⚠ 未确认修改")
        self.claim_status_label.setText("  ·  ".join(parts))

    def _on_claim_check_start(self):
        """点「开始检查」：先从预览框收集文本并做检查（不会自动写回内存）"""
        if not self._claim_loaded or self._claim_start_idx is None:
            self._show_toast("请先加载包含权利要求书的文档", "warning")
            return

        # 从预览框拿当前内容（可能是用户修改过但未确认的）
        text = self.claim_preview_edit.toPlainText()
        lines = text.split("\n")
        if len(lines) != self._claim_para_count:
            QMessageBox.warning(
                self, "段落数变化",
                f"预览框现有 {len(lines)} 行，而加载时为 {self._claim_para_count} 段。\n\n"
                "本功能要求每行对应一个段落（v1 限制）。请：\n"
                "• 不要增减整行（不要按回车新增空行，也不要删除整行）；\n"
                "• 仅在行内修改文字；\n"
                "• 如需大幅调整结构，请使用其它功能完成后再回来。"
            )
            return

        # 临时构造 dummy paragraphs：用简单壳类，只提供 .text 属性给 parse_claims 用
        class _Shell:
            __slots__ = ("text",)

            def __init__(self, t):
                self.text = t

        # 为了让 para_idx 的偏移与全文一致，我们只把权利要求书区间替换成 shell
        doc_paragraphs = self.doc_data['paragraphs']
        shell_paragraphs = list(doc_paragraphs)
        for i, line in enumerate(lines):
            shell_paragraphs[self._claim_start_idx + i] = _Shell(line)

        try:
            from claim_check import run_all_checks
            from config_manager import load_vague_wordbank, load_boundary_blacklist
            vague_words = load_vague_wordbank()
            use_trunc = self.claim_dyn_trunc_cb.isChecked()
            use_fb = self.claim_dyn_fb_cb.isChecked()
            boundary_bl = load_boundary_blacklist() if use_trunc else None
            n = int(self._claim_n)
            results = run_all_checks(
                shell_paragraphs,
                self._claim_start_idx,
                self._claim_end_idx,
                n=n,
                ignore_set=set(self._claim_session_ignore),
                vague_words=vague_words,
                check_vague=self.claim_vague_cb.isChecked(),
                check_term=self.claim_term_cb.isChecked(),
                use_dynamic_truncate=use_trunc,
                use_dynamic_fallback=use_fb,
                boundary_blacklist=boundary_bl,
            )
        except Exception as e:
            import traceback as tb
            QMessageBox.critical(
                self, "检查失败",
                f"权利要求书检查出现异常：\n{e}\n\n{tb.format_exc()}"
            )
            return

        self._claim_results = results
        self._render_claim_results(results)
        self._update_claim_status_bar()
        if results:
            self._show_toast(f"发现 {len(results)} 条问题", "warning")
        else:
            self._show_toast("未发现问题", "success")

    def _render_claim_results(self, results: list):
        """渲染检查结果到右侧表格"""
        KIND_LABELS = {
            "antecedent": "引用基础",
            "dependency": "引用关系",
            "term":       "术语不一致",
            "vague":      "不确定用语",
            "numbering":  "序号",
            "multi_dep":  "多引合法性",
        }
        self.claim_result_table.setRowCount(0)
        for row_idx, item in enumerate(results):
            self.claim_result_table.insertRow(row_idx)

            kind_item = QTableWidgetItem(KIND_LABELS.get(item.get("kind"), item.get("kind", "")))
            kind_item.setFlags(kind_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.claim_result_table.setItem(row_idx, 0, kind_item)

            claim_no = item.get("claim_no")
            claim_str = f"权{claim_no}" if claim_no else "-"
            no_item = QTableWidgetItem(claim_str)
            no_item.setFlags(no_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.claim_result_table.setItem(row_idx, 1, no_item)

            ctx_item = QTableWidgetItem(item.get("context", ""))
            ctx_item.setFlags(ctx_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.claim_result_table.setItem(row_idx, 2, ctx_item)

            msg_item = QTableWidgetItem(item.get("message", ""))
            msg_item.setFlags(msg_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.claim_result_table.setItem(row_idx, 3, msg_item)

            # 操作列：忽略 — 使用 setFixedSize 强制尺寸，绕开 QSS 级联压缩
            btn = QPushButton("忽略")
            btn.setObjectName("rowActionBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFixedSize(68, 28)
            btn.clicked.connect(lambda _=False, r=row_idx: self._on_claim_ignore_row(r))
            self.claim_result_table.setCellWidget(row_idx, 4, btn)
            # 显式给当前行一个下限行高，避免 Qt 以 item 的 sizeHint 压缩行高
            self.claim_result_table.setRowHeight(row_idx, 34)

    def _show_claim_result_detail(self, row: int):
        """弹出只读对话框显示指定结果行的完整字段（避免 elide 截断）。"""
        if row < 0 or row >= len(self._claim_results):
            return
        item = self._claim_results[row]
        kind_map = {
            "antecedent": "引用基础",
            "dependency": "引用关系",
            "term": "术语一致性",
            "vague": "不确定用语",
            "numbering": "独立权项序号",
            "multi_dep": "多项引用合法性",
        }
        kind_label = kind_map.get(item.get("kind"), item.get("kind") or "—")
        claim_no = item.get("claim_no")
        claim_label = f"权利要求{claim_no}" if claim_no else "—"
        context = item.get("context") or ""
        message = item.get("message") or ""
        suggestion = item.get("suggestion") or ""

        dlg = QDialog(self)
        dlg.setWindowTitle("问题详情")
        dlg.setMinimumSize(560, 360)
        dlg.setModal(True)

        lay = QVBoxLayout(dlg)
        lay.setSpacing(10)

        meta = QLabel(f"<b>类型：</b>{kind_label}　　<b>位置：</b>{claim_label}")
        meta.setTextFormat(Qt.TextFormat.RichText)
        lay.addWidget(meta)

        lay.addWidget(QLabel("上下文："))
        ctx_view = QPlainTextEdit()
        ctx_view.setReadOnly(True)
        ctx_view.setPlainText(context)
        ctx_view.setMaximumHeight(90)
        lay.addWidget(ctx_view)

        lay.addWidget(QLabel("说明："))
        msg_view = QPlainTextEdit()
        msg_view.setReadOnly(True)
        msg_view.setPlainText(message)
        lay.addWidget(msg_view, 1)

        if suggestion:
            lay.addWidget(QLabel("建议："))
            sug_view = QPlainTextEdit()
            sug_view.setReadOnly(True)
            sug_view.setPlainText(suggestion)
            sug_view.setMaximumHeight(80)
            lay.addWidget(sug_view)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(close_btn)
        lay.addLayout(btn_row)

        dlg.exec()

    def _on_claim_result_double_clicked(self, row: int, col: int):
        """
        双击结果表：
          - 第 2 列「上下文」→ 在左侧预览框定位到对应段落并高亮关键片段
          - 第 3 列「说明」  → 弹出完整说明对话框（避免 elide 显示不全）
          - 其它列           → 不响应

        上下文定位策略：
          1. 以结果的 para_idx 锚定 preview 中的行号
          2. 尝试在该行文本里找到 context（去首尾空格）的最长非空白片段
          3. 若找不到就整行选中
        高亮用 QPlainTextEdit.ExtraSelection，下次双击时自动替换。
        """
        if row < 0 or row >= len(self._claim_results):
            return
        if col == 3:
            self._show_claim_result_detail(row)
            return
        if col != 2:
            return
        if not self._claim_loaded or self._claim_start_idx is None:
            return

        item = self._claim_results[row]
        para_idx = item.get("para_idx")
        if para_idx is None or para_idx < 0:
            return
        line_no = para_idx - self._claim_start_idx
        if line_no < 0 or line_no >= self._claim_para_count:
            return

        doc = self.claim_preview_edit.document()
        block = doc.findBlockByNumber(line_no)
        if not block.isValid():
            return
        line_text = block.text()

        # 从 context 中抽一个"搜索锚"：取最长的非空白子串（通常是术语本身）
        context = (item.get("context") or "").strip()
        search_key = _longest_nonspace_run(context)

        start_in_line = -1
        if search_key:
            start_in_line = line_text.find(search_key)
        # 二次兜底：优先用术语字面量（从 message 里抽）
        if start_in_line < 0:
            import re as _re
            msg = item.get("message", "")
            m = _re.search(r'『所述(.+?)』', msg) or _re.search(r'『(.+?)』', msg)
            if m:
                search_key = m.group(1)
                start_in_line = line_text.find(search_key)

        cursor = QTextCursor(block)
        if start_in_line >= 0 and search_key:
            cursor.setPosition(block.position() + start_in_line)
            cursor.setPosition(
                block.position() + start_in_line + len(search_key),
                QTextCursor.MoveMode.KeepAnchor,
            )
        else:
            cursor.movePosition(QTextCursor.MoveOperation.StartOfBlock)
            cursor.movePosition(
                QTextCursor.MoveOperation.EndOfBlock,
                QTextCursor.MoveMode.KeepAnchor,
            )

        # 持久高亮（QPlainTextEdit 的 ExtraSelection 类型来自 QTextEdit）
        extra = QTextEdit.ExtraSelection()
        fmt = QTextCharFormat()
        fmt.setBackground(QColor("#ffd966"))  # 柔和黄
        fmt.setForeground(QColor("#1a237e"))
        extra.format = fmt
        extra.cursor = QTextCursor(cursor)
        self.claim_preview_edit.setExtraSelections([extra])

        # 光标 + 滚动 + 聚焦
        self.claim_preview_edit.setTextCursor(cursor)
        self.claim_preview_edit.centerCursor()
        self.claim_preview_edit.setFocus()

    def _on_claim_ignore_row(self, row: int):
        """
        点击行内「忽略」：本次会话忽略。不写入任何持久词库。

        - 从 message 中解析出术语（antecedent 的『所述X』里的 X；term 的两个相似词）
        - 加入 self._claim_session_ignore（仅当次会话有效，新开文档即清空）
        - 同类型且术语相同的其它行一并移除
        - vague / dependency / numbering 等其它类型的"忽略"只是从表格里移除该行
        """
        if row < 0 or row >= len(self._claim_results):
            return
        item = self._claim_results[row]
        kind = item.get("kind")
        msg = item.get("message", "")
        import re as _re

        session_adds = []
        if kind == "antecedent":
            m = _re.search(r'『所述(.+?)』', msg)
            if m:
                session_adds.append(m.group(1))
        elif kind == "term":
            m = _re.search(r'『(.+?)』与『(.+?)』', msg)
            if m:
                session_adds.extend([m.group(1), m.group(2)])

        for w in session_adds:
            if w:
                self._claim_session_ignore.add(w)

        if session_adds:
            self._show_toast(
                f"本次忽略：{'、'.join(session_adds)}（未加入词库）",
                "info",
            )

        # 移除本行；若是术语类，顺带把剩余结果里同术语的同类型行也清掉，
        # 免得用户点一下还要再点一堆
        removed_indices = {row}
        if session_adds:
            for i, r in enumerate(self._claim_results):
                if i == row or r.get("kind") != kind:
                    continue
                r_msg = r.get("message", "")
                if any(w and w in r_msg for w in session_adds):
                    removed_indices.add(i)
        self._claim_results = [
            r for i, r in enumerate(self._claim_results) if i not in removed_indices
        ]
        self._render_claim_results(self._claim_results)
        self._update_claim_status_bar()

    def _on_claim_ignore_dialog(self):
        """打开忽略词库编辑对话框"""
        try:
            from claim_ignore_dialog import ClaimIgnoreDialog
        except Exception as e:
            QMessageBox.critical(self, "无法打开", f"加载忽略词库编辑器失败：\n{e}")
            return
        dlg = ClaimIgnoreDialog(self)
        dlg.exec()

    # ── 引用基础降噪：动态截断 / 动态回退 信息弹窗与黑名单入口 ──
    def _on_dyn_trunc_link(self, href: str):
        if href == "bl":
            self._on_open_boundary_blacklist()
            return
        QMessageBox.information(
            self, "动态截断 — 功能说明",
            "【动态截断】黑名单边界识别法\n\n"
            "默认的引用基础检查会取『所述』后面的固定 N 个字作为术语，\n"
            "但 N 偏大时常把后续动词/方位词一起吞进来，造成误判。\n\n"
            "动态截断的做法：\n"
            "  • 从『所述』往后扫 CJK 字符\n"
            "  • 一旦命中黑名单词的首字（如『安装』『上方』『与』），\n"
            "    立即停止，把之前累积的字串作为术语\n"
            "  • 例：「所述齿轮安装在主轴上」→ 命中『安装』→ 术语 = 齿轮\n\n"
            "黑名单可点本行右侧的「[黑名单]」按钮编辑。\n\n"
            "适用场景：词库越完善，识别越精准，是降噪首选方案。"
        )

    def _on_dyn_fb_link(self, href: str):
        QMessageBox.information(
            self, "动态回退 — 功能说明",
            "【动态回退】渐进式前缀匹配法\n\n"
            "默认检查只比对固定 N 字术语，若 N 字串没在前文出现就报错；\n"
            "但实际上前文可能用更短的形式定义过这个部件。\n\n"
            "动态回退的做法：\n"
            "  • 取『所述』后的 N 字术语\n"
            "  • 如果找不到对应定义，就把末尾字砍掉，换成 N-1 字再试\n"
            "  • 继续缩短到 N-2、N-3 …，任一前缀匹配上就放过\n"
            "  • 全部缩到 2 字仍不匹配才报错\n\n"
            "例：N=6 时『所述齿轮安装在主』→ 砍 → 齿轮安装在 → 齿轮安装\n"
            "    → 齿轮（在前文出现过！）→ 放过\n\n"
            "适用场景：作为容错机制，能极大降低代理人改字数后产生的噪音。\n\n"
            "── 与「动态截断」同时勾选时 ──\n"
            "先用截断得到一个干净的最长术语，再对该术语应用回退\n"
            "（从右向左缩短前缀），仅当所有前缀都没匹配上时才报错。\n"
            "这是误判最低的组合策略。"
        )

    def _on_open_boundary_blacklist(self):
        try:
            from boundary_blacklist_dialog import BoundaryBlacklistDialog
        except Exception as e:
            QMessageBox.critical(self, "无法打开", f"加载黑名单词库编辑器失败：\n{e}")
            return
        dlg = BoundaryBlacklistDialog(self)
        dlg.exec()

    # ── 不确定用语检查 / 术语不一致检查 的信息弹窗 + 词库入口 ──
    def _on_vague_link(self, href: str):
        if href == "wb":
            self._on_claim_ignore_dialog()
            return
        QMessageBox.information(
            self, "不确定用语检查 — 功能说明",
            "【不确定用语检查】\n\n"
            "扫描权利要求书全文，查找属于「不确定 / 含糊」用语词库的词汇，\n"
            "例如：约、大概、可能、优选、左右、基本、通常 …\n\n"
            "权利要求书中应避免不确定用语，否则常被审查员以「保护范围不清楚」\n"
            "为由发出审查意见 (OA)。\n\n"
            "• 词库可点本行右侧的「[词库]」按钮编辑，支持增删 / 导入 / 导出\n"
            "• 默认勾选；如需关闭检查可取消勾选\n"
            "• 内置 30+ 条常见词，可恢复默认"
        )

    def _on_term_link(self, href: str):
        QMessageBox.information(
            self, "术语不一致检查 — 功能说明",
            "【术语不一致检查（同一术语多种写法）】\n\n"
            "以『所述』后面的 N 字术语作为锚点，在权利要求书范围内查找\n"
            "「长度相同但仅一字之差」的相似术语对，作为可能的术语漂移上报。\n\n"
            "例：权 1 写「齿圈」、权 2 写「齿环」→ 报『齿圈』vs『齿环』疑似同义\n\n"
            "• 该检查噪音相对较大，默认不勾选\n"
            "• 仅在做权利要求书术语一致性复盘时建议启用\n"
            "• 受工具栏「检查字数 N」与「忽略词库」共同影响"
        )

    def _on_claim_confirm_edits(self) -> bool:
        """
        确认修改：把预览框中的内容写回内存 paragraphs[start:end]。
        返回 True 表示成功或无需写回；False 表示失败（段落数变化等）。
        """
        if not self._claim_loaded or self._claim_start_idx is None:
            return True
        if not self._claim_dirty:
            return True

        text = self.claim_preview_edit.toPlainText()
        lines = text.split("\n")
        if len(lines) != self._claim_para_count:
            QMessageBox.warning(
                self, "段落数变化",
                f"预览框现有 {len(lines)} 行，而原文档为 {self._claim_para_count} 段。\n\n"
                "本功能要求每行对应一个段落（v1 限制）。请恢复为原段数后再确认，"
                "或使用撤销 (Ctrl+Z) 回到上一状态。"
            )
            return False

        paragraphs = self.doc_data['paragraphs']
        from claim_check import set_paragraph_text
        changed_count = 0
        changed_lines = []
        for i, new_line in enumerate(lines):
            para = paragraphs[self._claim_start_idx + i]
            old_line = para.text if para.text else ""
            if old_line != new_line:
                try:
                    if set_paragraph_text(para, new_line):
                        changed_count += 1
                        preview_old = old_line.strip()[:40]
                        preview_new = new_line.strip()[:40]
                        changed_lines.append(
                            f"第{i+1}段：{preview_old} → {preview_new}"
                        )
                except Exception as e:
                    QMessageBox.critical(
                        self, "写回失败",
                        f"第 {i+1} 段写回时出错：\n{e}"
                    )
                    return False

        self._claim_dirty = False
        self.claim_confirm_btn.setEnabled(False)
        self._update_claim_status_bar()

        if changed_count:
            detail = "\n".join(changed_lines[:20])
            if len(changed_lines) > 20:
                detail += f"\n…（共 {len(changed_lines)} 处改动）"
            self._add_history(
                f"编辑权利要求书（{changed_count} 段）",
                detail,
            )
            self._show_toast(f"已确认 {changed_count} 段修改", "success")
        else:
            self._show_toast("无实际改动", "info")
        return True
