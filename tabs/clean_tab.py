# -*- coding: utf-8 -*-
"""
tabs/clean_tab.py — 清理 Tab（所述 / 标点 / 孤立标记）

以 Mixin 形式抽离自 main_window.py，由 MainWindow 继承。保持方法体内所有
self.xxx 引用不变。Mixin 依赖 MainWindow 提供：
  - self.doc_data / self.current_marks
  - self.suoshu_checkboxes / self.status_bar / self.progress_bar
  - self._show_toast / self._add_history / self._log
  - CleanWorker, SUOSHU_ALLOWED_SECTIONS, SUOSHU_DEFAULT_CHECKED
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QGroupBox, QCheckBox, QFrame, QMessageBox,
)
from PyQt6.QtCore import Qt, QTimer

from workers import CleanWorker

# 「删除所述」功能允许处理的章节白名单
SUOSHU_ALLOWED_SECTIONS = ("权利要求书", "背景技术", "具体实施方式")
# 默认勾选的章节
SUOSHU_DEFAULT_CHECKED = ("具体实施方式",)


class CleanTabMixin:
    """清理 Tab 方法集合，供 MainWindow 继承。"""

    def _create_clean_tab(self) -> QWidget:
        """创建文本清洗标签页：三个清洗功能并排展示"""
        outer = QWidget()
        outer_layout = QVBoxLayout(outer)
        outer_layout.setContentsMargins(0, 8, 0, 0)
        outer_layout.setSpacing(10)

        # ── 顶部：三功能并排 ─────────────────────────────
        cards_row = QHBoxLayout()
        cards_row.setSpacing(12)

        cards_row.addWidget(self._build_suoshu_card(), 1)
        cards_row.addWidget(self._build_punct_card(), 1)
        cards_row.addWidget(self._build_orphan_card(), 1)

        outer_layout.addLayout(cards_row)

        # ── 底部：清洗日志 ───────────────────────────────
        clean_log_group = QGroupBox("📋 清洗操作日志")
        clean_log_layout = QVBoxLayout(clean_log_group)
        self.clean_log_text = QTextEdit()
        self.clean_log_text.setReadOnly(True)
        self.clean_log_text.setPlaceholderText("清洗操作的日志将在此显示...")
        clean_log_layout.addWidget(self.clean_log_text)
        outer_layout.addWidget(clean_log_group, 1)

        return outer

    def _build_suoshu_card(self) -> QGroupBox:
        '''卡片①：删除"所述"'''
        group = QGroupBox('🗑️  删除"所述"')
        group.setObjectName("cleanCard")
        v = QVBoxLayout(group)
        v.setSpacing(8)

        hint = QLabel('勾选章节后，所选章节中的"所述"将被直接删除。')
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        v.addWidget(hint)

        # 章节勾选区
        cb_box = QFrame()
        cb_box.setObjectName("checkBoxFrame")
        self.suoshu_cb_layout = QVBoxLayout(cb_box)
        self.suoshu_cb_layout.setSpacing(6)
        self.suoshu_cb_layout.setContentsMargins(10, 8, 10, 8)
        self._suoshu_placeholder = QLabel("（请先加载文档）")
        self._suoshu_placeholder.setObjectName("subtitleLabel")
        self.suoshu_cb_layout.addWidget(self._suoshu_placeholder)
        v.addWidget(cb_box)

        v.addStretch()

        self.suoshu_btn = QPushButton('▶  执行删除"所述"')
        self.suoshu_btn.setObjectName("accentBtn")
        self.suoshu_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.suoshu_btn.setEnabled(False)
        self.suoshu_btn.clicked.connect(self._on_clean_suoshu)
        v.addWidget(self.suoshu_btn)
        return group

    def _build_punct_card(self) -> QGroupBox:
        """卡片②：标点检查"""
        group = QGroupBox("🔤  标点检查")
        group.setObjectName("cleanCard")
        v = QVBoxLayout(group)
        v.setSpacing(8)

        hint = QLabel(
            "将中文正文中的半角标点（, ; : ? ! . ( ) ' \"）替换为对应的全角，"
            "仅在紧邻中文字符时替换，不影响英文/数字；不处理 < > 以免误伤数学符号。\n"
            "执行顺序：半角→全角 → 全角→半角 → 修正连续重复标点。"
        )
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        v.addWidget(hint)

        # 选项区
        opt_box = QFrame()
        opt_box.setObjectName("checkBoxFrame")
        opt_layout = QVBoxLayout(opt_box)
        opt_layout.setSpacing(6)
        opt_layout.setContentsMargins(10, 8, 10, 8)

        self.punct_halfwidth_cb = QCheckBox("半角 → 全角标点")
        self.punct_halfwidth_cb.setChecked(True)
        opt_layout.addWidget(self.punct_halfwidth_cb)

        self.punct_fullwidth_cb = QCheckBox("全角 → 半角标点")
        self.punct_fullwidth_cb.setChecked(False)
        self.punct_fullwidth_cb.setToolTip(
            "无条件将段落中的全角标点替换为半角；与「半角→全角」勾选时后者优先生效"
        )
        opt_layout.addWidget(self.punct_fullwidth_cb)

        self.fix_punctuation_cb = QCheckBox("修正连续重复标点")
        self.fix_punctuation_cb.setChecked(True)
        self.fix_punctuation_cb.setToolTip("会在前两步之后再执行，避免出现 `.。` 这类混合连续标点遗漏")
        opt_layout.addWidget(self.fix_punctuation_cb)

        v.addWidget(opt_box)
        v.addStretch()

        self.punct_btn = QPushButton("▶  执行标点检查")
        self.punct_btn.setObjectName("accentBtn")
        self.punct_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.punct_btn.setEnabled(False)
        self.punct_btn.clicked.connect(self._on_clean_punct)
        v.addWidget(self.punct_btn)
        return group

    def _build_orphan_card(self) -> QGroupBox:
        """卡片③：孤立附图标记检测"""
        group = QGroupBox("🔍  孤立附图标记检测")
        group.setObjectName("cleanCard")
        v = QVBoxLayout(group)
        v.setSpacing(8)

        hint = QLabel(
            "找出「附图说明中提及、但具体实施方式中未出现」的内容，包括：\n"
            "  • 附图标记名称（如齿圈、夹指…）\n"
            "  • 图编号（如图1、图5…）\n"
            "该功能仅做检测，不会修改文档。"
        )
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        v.addWidget(hint)

        self.orphan_btn = QPushButton("🔍  检测孤立标记")
        self.orphan_btn.setObjectName("accentBtn")
        self.orphan_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.orphan_btn.setEnabled(False)
        self.orphan_btn.clicked.connect(self._on_detect_orphans)
        v.addWidget(self.orphan_btn)

        # 该卡片专属的小型结果显示框
        self.orphan_result_text = QTextEdit()
        self.orphan_result_text.setReadOnly(True)
        self.orphan_result_text.setPlaceholderText("孤立标记检测结果将在此显示...")
        self.orphan_result_text.setMinimumHeight(120)
        v.addWidget(self.orphan_result_text, 1)
        return group

    def _update_suoshu_section_checkboxes(self, sections: dict):
        """文档加载后动态生成「删除所述」章节勾选框（仅白名单章节）"""
        # 清除旧控件
        while self.suoshu_cb_layout.count():
            item = self.suoshu_cb_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.suoshu_checkboxes.clear()

        # 仅按白名单顺序展示，且只展示文档中实际存在的章节
        any_added = False
        for name in SUOSHU_ALLOWED_SECTIONS:
            if name not in sections:
                continue
            cb = QCheckBox(name)
            cb.setChecked(name in SUOSHU_DEFAULT_CHECKED)
            self.suoshu_checkboxes[name] = cb
            self.suoshu_cb_layout.addWidget(cb)
            any_added = True

        if not any_added:
            tip = QLabel("（文档中未识别到可处理的章节）")
            tip.setObjectName("subtitleLabel")
            self.suoshu_cb_layout.addWidget(tip)

    def _get_selected_suoshu_sections(self) -> list:
        """返回已勾选的章节名列表"""
        return [name for name, cb in self.suoshu_checkboxes.items() if cb.isChecked()]

    def _log_clean(self, message: str):
        """向清洗日志区追加一行"""
        self.clean_log_text.append(message)
        sb = self.clean_log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _set_clean_buttons_enabled(self, enabled: bool):
        self.suoshu_btn.setEnabled(enabled)
        self.punct_btn.setEnabled(enabled)
        self.orphan_btn.setEnabled(enabled)
        self.typo_check_btn.setEnabled(enabled)
        self.dup_check_btn.setEnabled(enabled)
        # 应用按钮：恢复时仅在当前显示的检查列表非空时点亮
        if enabled:
            active = self._active_cache_list() if hasattr(self, "_current_check_kind") else []
            self.typo_apply_btn.setEnabled(bool(active))

    def _start_clean_worker(self, action: str, log_prefix: str, history_label: str = None, **kwargs):
        """通用：启动 CleanWorker
        history_label 不为 None 时，操作完成后自动写入操作历史框。
        """
        if not self.doc_data:
            self._show_toast("请先加载文档！", "error")
            return
        self._set_clean_buttons_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self._log_clean(f"{'='*40}")
        self._log_clean(f"▶ {log_prefix}...")

        # 暂存当前历史标签 / action，供 finished 回调使用
        self._pending_history_label = history_label
        self._pending_clean_action = action

        self.clean_worker = CleanWorker(self.doc_data, action, **kwargs)
        self.clean_worker.progress.connect(self.progress_bar.setValue)
        self.clean_worker.finished.connect(self._on_clean_finished)
        self.clean_worker.error.connect(self._on_clean_error)
        if action == "typo_check":
            self.clean_worker.typo_results.connect(self._on_typo_results_ready)
        elif action == "dup_check":
            self.clean_worker.typo_results.connect(self._on_dup_results_ready)
        self.clean_worker.start()

    def _on_clean_suoshu(self):
        selected = self._get_selected_suoshu_sections()
        if not selected:
            self._show_toast('请至少勾选一个章节！', "error")
            return
        label = f'删除"所述"（{"、".join(selected)}）'
        self._start_clean_worker("suoshu", label, history_label=label,
                                  selected_sections=selected)

    def _on_clean_punct(self):
        do_half = self.punct_halfwidth_cb.isChecked()
        do_full = self.punct_fullwidth_cb.isChecked()
        do_consec = self.fix_punctuation_cb.isChecked()
        if not (do_half or do_full or do_consec):
            self._show_toast("请至少勾选一种标点处理方式！", "warning")
            return
        parts = []
        if do_half:
            parts.append("半角→全角")
        if do_full:
            parts.append("全角→半角")
        if do_consec:
            parts.append("修正连续标点")
        label = f"标点检查（{' + '.join(parts)}）"
        self._start_clean_worker(
            "punct",
            label,
            history_label=label,
            do_halfwidth=do_half,
            do_fullwidth=do_full,
            do_consecutive=do_consec,
        )

    def _on_detect_orphans(self):
        if not self.current_marks:
            self._show_toast("标记字典为空，无法检测！", "error")
            return
        # 检测类不写入历史（不修改文档）
        self._start_clean_worker("orphan", "孤立附图标记检测", marks=self.current_marks)

    # ─── 错别字 / 重复字词共用同一张表 ────────────────────
    def _on_clean_finished(self, message: str):
        """清洗操作完成"""
        self._set_clean_buttons_enabled(True)
        self.progress_bar.setVisible(False)
        action = getattr(self, "_pending_clean_action", None)

        # 孤立标记检测：结果只显示在自己卡片的小日志框，不写入全局清洗日志
        if action == "orphan" and hasattr(self, "orphan_result_text"):
            from datetime import datetime
            stamp = datetime.now().strftime("%H:%M:%S")
            self.orphan_result_text.append(f"[{stamp}] {message}\n")
        else:
            self._log_clean(f"✅ {message}")

        # 应用修正完成后，清空当前显示的检查结果并刷新表格（提示用户修改已生效）
        if action == "typo_apply":
            kind = getattr(self, "_current_check_kind", None)
            if kind == "typo":
                self.typo_data = []
            elif kind == "dup":
                self.dup_data = []
            if hasattr(self, "_render_table_from_data"):
                self._render_table_from_data([])

        self.status_bar.showMessage(message)
        self._show_toast(message[:40], "success")

        # 写入操作历史（仅当 _start_clean_worker 提供了 history_label）
        label = getattr(self, "_pending_history_label", None)
        if label:
            self._add_history(label, message)
        self._pending_history_label = None
        self._pending_clean_action = None

        QTimer.singleShot(1500, lambda: self.progress_bar.setVisible(False))

    def _on_clean_error(self, error_msg: str):
        """清洗操作失败"""
        self._set_clean_buttons_enabled(True)
        self.progress_bar.setVisible(False)
        self._log_clean(f"❌ {error_msg}")
        self.status_bar.showMessage("清洗操作失败")
        self._show_toast("操作失败！", "error")
