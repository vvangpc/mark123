# -*- coding: utf-8 -*-
"""
tabs/mark_tab.py — 标记提取 / 标注 / 文件生成 Tab + 章节预览

以 Mixin 形式抽离自 main_window.py，由 MainWindow 继承。保持方法体内所有
self.xxx 引用不变。Mixin 依赖 MainWindow 提供：
  - self.doc_data / self.current_marks / self.current_file_path
  - self.history_entries / self.worker / self.progress_bar / self.status_bar
  - self._show_toast / self._add_history / self._log
  - AnnotateWorker, ToastWidget
"""
import os

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QPlainTextEdit, QFileDialog, QMessageBox,
    QFrame, QSplitter, QGroupBox, QLineEdit, QCheckBox,
)
from PyQt6.QtCore import Qt

from doc_parser import get_section_text
from mark_extractor import (
    extract_marks_from_paragraph, extract_marks_from_paragraphs,
    marks_to_display_text, parse_marks_from_display_text,
)
from annotator import (
    build_claims_replace_dict, build_implementation_replace_dict,
    annotate_section, annotate_paragraph_safe,
)
from workers import AnnotateWorker


class MarkTabMixin:
    """标记提取 / 标注 / 文件生成 Tab 的方法集合。"""

    def _create_mark_tab(self) -> QWidget:
        """创建标记提取与标注标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(12)

        # ── 上方: 标记字典编辑区 ──────────────────────────
        marks_group = QGroupBox("📌 附图标记字典（自动提取，可手动编辑）")
        marks_layout = QVBoxLayout(marks_group)

        self.marks_edit = QPlainTextEdit()
        self.marks_edit.setPlaceholderText(
            "打开 docx 文件后将自动提取附图标记...\n"
            "格式示例: 1-齿圈，2-夹指，3-转盘，4-定位销"
        )
        self.marks_edit.setMaximumHeight(90)
        marks_layout.addWidget(self.marks_edit)

        # 标记操作按钮行
        marks_btn_row = QHBoxLayout()

        self.confirm_marks_btn = QPushButton("✅  重新确认标记")
        self.confirm_marks_btn.setObjectName("primaryBtn")
        self.confirm_marks_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.confirm_marks_btn.setEnabled(False)
        self.confirm_marks_btn.setToolTip("将编辑框中的词典写回内存中的附图标记段落，并记入操作历史")
        self.confirm_marks_btn.clicked.connect(self._on_confirm_marks)
        marks_btn_row.addWidget(self.confirm_marks_btn)

        self.mark_count_label = QLabel("")
        self.mark_count_label.setObjectName("subtitleLabel")
        marks_btn_row.addWidget(self.mark_count_label)

        marks_btn_row.addStretch()

        self.refresh_marks_btn = QPushButton("🔄  重新提取标记")
        self.refresh_marks_btn.setObjectName("primaryBtn")
        self.refresh_marks_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.refresh_marks_btn.setEnabled(False)
        self.refresh_marks_btn.setToolTip("从原始 docx 文档重新提取附图标记到编辑框")
        self.refresh_marks_btn.clicked.connect(self._on_refresh_marks)
        marks_btn_row.addWidget(self.refresh_marks_btn)

        self.clear_marks_btn = QPushButton("🗑️ 清空标记")
        self.clear_marks_btn.setObjectName("dangerBtn")
        self.clear_marks_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_marks_btn.clicked.connect(lambda: self.marks_edit.clear())
        marks_btn_row.addWidget(self.clear_marks_btn)

        marks_layout.addLayout(marks_btn_row)
        layout.addWidget(marks_group)

        # ── 中间: 操作行 ─────────────────────────────────
        action_layout = QHBoxLayout()
        action_layout.setSpacing(12)

        self.annotate_btn = QPushButton("⚡  一键标注")
        self.annotate_btn.setObjectName("primaryBtn")
        self.annotate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.annotate_btn.setEnabled(False)
        self.annotate_btn.setToolTip("权利要求书 + 具体实施方式 全部自动标注（仅修改内存）")
        self.annotate_btn.clicked.connect(self._on_annotate)
        action_layout.addWidget(self.annotate_btn)

        self.annotate_claims_btn = QPushButton("📋 仅标注权利要求书")
        self.annotate_claims_btn.setObjectName("accentBtn")
        self.annotate_claims_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.annotate_claims_btn.setEnabled(False)
        self.annotate_claims_btn.clicked.connect(lambda: self._on_annotate_section("claims"))
        action_layout.addWidget(self.annotate_claims_btn)

        self.annotate_impl_btn = QPushButton("📝 仅标注具体实施方式")
        self.annotate_impl_btn.setObjectName("accentBtn")
        self.annotate_impl_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.annotate_impl_btn.setEnabled(False)
        self.annotate_impl_btn.clicked.connect(lambda: self._on_annotate_section("implementation"))
        action_layout.addWidget(self.annotate_impl_btn)

        self.remove_marks_btn = QPushButton("🧹 删除所有标记")
        self.remove_marks_btn.setObjectName("dangerBtn")
        self.remove_marks_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.remove_marks_btn.setEnabled(False)
        self.remove_marks_btn.setToolTip("基于标记字典，扫描并清洗正文中的编号（仅修改内存）")
        self.remove_marks_btn.clicked.connect(self._on_remove_marks)
        action_layout.addWidget(self.remove_marks_btn)

        action_layout.addStretch()

        # 文件生成按钮（与上述操作分离）
        self.generate_btn = QPushButton("💾  文件生成")
        self.generate_btn.setObjectName("primaryBtn")
        self.generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setToolTip("将内存中累计的所有修改保存为新的 docx 文件")
        self.generate_btn.clicked.connect(self._on_generate_file)
        action_layout.addWidget(self.generate_btn)

        # 「打开文件所在目录」勾选框：默认不勾选，避免自动弹资源管理器
        self.open_dir_cb = QCheckBox("打开文件所在目录")
        self.open_dir_cb.setChecked(False)
        self.open_dir_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_dir_cb.setToolTip(
            "勾选后，文件生成成功时会自动打开输出目录。\n"
            "默认不勾选，避免频繁弹窗打断工作。"
        )
        action_layout.addWidget(self.open_dir_cb)

        layout.addLayout(action_layout)

        # ── 下方：左日志 / 右历史 双栏 ───────────────────
        bottom_splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左：操作日志
        log_group = QGroupBox("📋 操作日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("操作的详细日志将在此显示...")
        log_layout.addWidget(self.log_text)
        bottom_splitter.addWidget(log_group)

        # 右：操作历史
        history_group = QGroupBox("📜 操作历史")
        history_layout = QVBoxLayout(history_group)
        self.history_text = QTextEdit()
        self.history_text.setReadOnly(True)
        self.history_text.setPlaceholderText(
            "尚未对当前文档进行修改。\n"
            "所有「标注 / 删除标记 / 清洗 / 错别字修正」等操作都会先记录在这里，\n"
            "待您确认后点击右上方的「💾 文件生成」按钮一次性写入新 docx。"
        )
        history_layout.addWidget(self.history_text)
        bottom_splitter.addWidget(history_group)

        bottom_splitter.setSizes([500, 400])
        layout.addWidget(bottom_splitter, 1)

        return widget

    def _create_preview_tab(self) -> QWidget:
        """创建文档内容预览标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)

        # 使用分割器
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左侧: 章节列表
        sections_panel = QFrame()
        sections_panel.setObjectName("glassPanel")
        sections_layout = QVBoxLayout(sections_panel)

        sections_title = QLabel("📑 文档章节")
        sections_title.setObjectName("sectionLabel")
        sections_layout.addWidget(sections_title)

        # 章节按钮容器
        self.section_buttons_layout = QVBoxLayout()
        sections_layout.addLayout(self.section_buttons_layout)
        sections_layout.addStretch()

        splitter.addWidget(sections_panel)

        # 右侧: 内容预览
        preview_panel = QFrame()
        preview_panel.setObjectName("glassPanel")
        preview_layout = QVBoxLayout(preview_panel)

        preview_title_row = QHBoxLayout()
        self.preview_title = QLabel("📄 内容预览（可编辑）")
        self.preview_title.setObjectName("sectionLabel")
        preview_title_row.addWidget(self.preview_title)
        preview_title_row.addStretch()
        self.preview_status_label = QLabel("")
        self.preview_status_label.setObjectName("subtitleLabel")
        preview_title_row.addWidget(self.preview_status_label)
        preview_layout.addLayout(preview_title_row)

        self.preview_text = QTextEdit()
        self.preview_text.setPlaceholderText(
            "选择左侧章节查看内容。\n"
            "可直接编辑本区域；修改后点下方的「✔ 确认修改」把改动写回内存，"
            "最终「💾 文件生成」时才会落盘到 .docx。"
        )
        self.preview_text.textChanged.connect(self._on_preview_text_changed)
        preview_layout.addWidget(self.preview_text)

        self.preview_confirm_btn = QPushButton("✔  确认修改")
        self.preview_confirm_btn.setObjectName("primaryBtn")
        self.preview_confirm_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.preview_confirm_btn.setEnabled(False)
        self.preview_confirm_btn.setToolTip(
            "将预览框的当前内容写回内存中的章节段落。\n"
            "段落数必须保持不变（每行对应一段），最终 .docx 会反映此处的修改。"
        )
        self.preview_confirm_btn.clicked.connect(self._on_preview_confirm_edits)
        preview_layout.addWidget(self.preview_confirm_btn)

        # 预览编辑状态
        self._preview_section_name = None
        self._preview_start_idx = None
        self._preview_end_idx = None
        self._preview_para_count = 0
        self._preview_dirty = False
        self._preview_loading = False

        splitter.addWidget(preview_panel)

        # 设置分割比例
        splitter.setSizes([220, 800])

        layout.addWidget(splitter)

        return widget

    def _extract_and_display_marks(self):
        """提取标记并显示在编辑框"""
        if not self.doc_data:
            return

        mark_paras = self.doc_data.get('mark_paras', [])
        mark_para = self.doc_data.get('mark_para')
        if mark_paras:
            marks = extract_marks_from_paragraphs(mark_paras)
        elif mark_para:
            marks = extract_marks_from_paragraph(mark_para)
        else:
            marks = {}

        self.current_marks = marks
        if marks:
            self.marks_edit.setPlainText(marks_to_display_text(marks))
            self.mark_count_label.setText(f"共 {len(marks)} 个标记")
        else:
            self.marks_edit.setPlainText("")
            self.mark_count_label.setText("未找到附图标记段落")

    def _on_refresh_marks(self):
        """重新从原始 docx 文档提取标记到编辑框（不影响内存修改）"""
        if not self.current_file_path:
            self._show_toast("请先打开文档！", "error")
            return
        try:
            fresh = parse_document(self.current_file_path)
        except Exception as e:
            self._show_toast(f"重新解析失败：{e}", "error")
            return
        mark_paras = fresh.get('mark_paras', [])
        mark_para = fresh.get('mark_para')
        if not mark_para:
            self._show_toast("原始文档中未找到附图标记段落", "warning")
            self.marks_edit.setPlainText("")
            self.current_marks = {}
            self.mark_count_label.setText("未找到附图标记段落")
            return

        marks = extract_marks_from_paragraphs(mark_paras) if mark_paras else extract_marks_from_paragraph(mark_para)
        self.current_marks = marks
        self.marks_edit.setPlainText(marks_to_display_text(marks))
        self.mark_count_label.setText(f"共 {len(marks)} 个标记")
        self._log(f"🔄 从原始文档重新提取标记：共 {len(marks)} 个")
        self._show_toast(f"重新提取成功，共 {len(marks)} 个", "success")

    def _on_annotate(self):
        """一键标注权利要求书和具体实施方式（仅修改内存）"""
        self._start_annotate_worker(action="add", scope="all")

    def _on_remove_marks(self):
        """删除所有标记（仅修改内存）"""
        self._start_annotate_worker(action="remove", scope="all")

    def _on_annotate_section(self, mode: str):
        """单独标注某个章节（mode: 'claims' 或 'implementation'）"""
        section_name = "权利要求书" if mode == "claims" else "具体实施方式"
        if not self._validate_before_annotate():
            return
        if section_name not in self.doc_data['sections']:
            self._show_toast(f"未找到 {section_name} 章节！", "error")
            return
        self._start_annotate_worker(action="add", scope=mode)

    def _start_annotate_worker(self, action: str, scope: str):
        """通用：启动标注 Worker（仅修改内存中的 doc_data）"""
        if not self._validate_before_annotate():
            return
        self._sync_marks_from_editor()

        action_name = "标注" if action == "add" else "删除标记"
        scope_name = {"all": "全文", "claims": "权利要求书", "implementation": "具体实施方式"}[scope]

        self._log("=" * 50)
        self._log(f"🚀 开始{action_name}（{scope_name}）...")
        self._log(f"   标记数量: {len(self.current_marks)}")

        self._set_buttons_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.worker = AnnotateWorker(self.doc_data, self.current_marks, action=action, scope=scope)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(self._on_annotate_finished)
        self.worker.error.connect(self._on_annotate_error)
        self.worker.start()

    def _on_annotate_finished(self, summary: str, detail: str):
        """标注操作完成（仅内存修改完成）"""
        self._set_buttons_enabled(True)
        self._log(f"✅ {summary}")
        for line in detail.split("\n"):
            if line.strip():
                self._log(f"   {line}")

        self._add_history(summary, detail)
        self.status_bar.showMessage(f"{summary} — 已写入内存，待生成文件")
        self._show_toast(f"{summary} 已写入内存", "success")
        QTimer.singleShot(1500, lambda: self.progress_bar.setVisible(False))

    def _on_annotate_error(self, error_msg: str):
        """标注失败回调"""
        self._set_buttons_enabled(True)
        self.progress_bar.setVisible(False)
        self._log(f"❌ {error_msg}")
        self.status_bar.showMessage("操作失败")
        self._show_toast("操作失败！", "error")
        QMessageBox.critical(self, "操作失败", error_msg)

    # ===== 文件生成 =====

    def _on_generate_file(self):
        """将内存中累计的所有修改保存为新的 docx"""
        if not self.doc_data:
            self._show_toast("请先打开文档！", "error")
            return
        # 权利要求书 Tab 若有未确认修改，提示用户先确认
        if getattr(self, "_claim_dirty", False):
            reply = QMessageBox.question(
                self, "权利要求书有未确认修改",
                "「权利要求书检查」Tab 的预览框中有未确认的编辑。\n"
                "是否立即确认并把这些修改写入内存？\n\n"
                "• 选「Yes」：先确认再继续生成文件；\n"
                "• 选「No」：放弃这些未确认修改继续生成（文件不会包含它们）；\n"
                "• 选「Cancel」：中断，回到界面手动处理。",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
                | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Cancel:
                return
            if reply == QMessageBox.StandardButton.Yes:
                if not self._on_claim_confirm_edits():
                    return  # 段落数不匹配等 → 中断
            # No：放弃未确认修改，继续向下
        if not self.history_entries:
            reply = QMessageBox.question(
                self, "未检测到修改",
                "操作历史为空，确定要生成文件吗？\n（将复制原始文档为新文件）",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        # ── 弹出生成确认卡片，允许自定义后缀 ──
        default_action = self._infer_output_action_name()
        base_name = os.path.splitext(
            os.path.basename(self.current_file_path)
        )[0]
        if base_name.endswith("_已标注") or base_name.endswith("_已清洗"):
            base_name = base_name.rsplit("_", 1)[0]

        dlg = QDialog(self)
        dlg.setWindowTitle("文件生成")
        dlg.setMinimumWidth(420)
        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.setSpacing(12)

        dlg_layout.addWidget(QLabel(
            "即将生成文件。可在下方自定义后缀名称，留空则使用自动推断的默认值。"
        ))

        suffix_row = QHBoxLayout()
        suffix_row.addWidget(QLabel("文件后缀："))
        suffix_edit = QLineEdit()
        suffix_edit.setPlaceholderText(f"默认: {default_action}")
        suffix_edit.setToolTip(
            f"留空将自动使用「{default_action}」作为后缀。\n"
            "输入自定义内容后，文件名变为 原名_自定义.docx"
        )
        suffix_row.addWidget(suffix_edit, 1)
        dlg_layout.addLayout(suffix_row)

        preview_label = QLabel()
        preview_label.setObjectName("subtitleLabel")
        preview_label.setWordWrap(True)

        def _update_preview():
            custom = suffix_edit.text().strip()
            action = custom if custom else default_action
            preview_label.setText(f"预览：{base_name}_{action}.docx")

        suffix_edit.textChanged.connect(lambda _: _update_preview())
        _update_preview()
        dlg_layout.addWidget(preview_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("生成")
        ok_btn.setObjectName("primaryBtn")
        ok_btn.clicked.connect(dlg.accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        dlg_layout.addLayout(btn_row)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        custom_suffix = suffix_edit.text().strip()
        action_name = custom_suffix if custom_suffix else default_action
        output_path = self._generate_output_path(action_name)
        try:
            self.doc_data['document'].save(output_path)
        except PermissionError:
            QMessageBox.critical(
                self, "保存失败",
                f"文件 '{os.path.basename(output_path)}' 正被其他程序占用！\n\n"
                "请先在 Word 或 WPS 中【关闭】该文档，然后重试。"
            )
            return
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"无法保存文件：\n{e}")
            return

        self._log("=" * 50)
        self._log(f"💾 文件已生成: {output_path}")
        self.status_bar.showMessage(f"文件已生成 — {os.path.basename(output_path)}")
        self._show_toast("文件生成成功！", "success")

        QMessageBox.information(
            self, "文件生成完成",
            f"文件已保存至：\n\n{output_path}",
            QMessageBox.StandardButton.Ok,
        )
        if self.open_dir_cb.isChecked():
            try:
                os.startfile(os.path.dirname(output_path))
            except Exception:
                pass


    def _on_confirm_marks(self):
        """重新确认标记：将编辑框中的词典写回内存中的附图标记段落"""
        if not self.doc_data:
            self._show_toast("请先打开文档！", "error")
            return
        text = self.marks_edit.toPlainText().strip()
        if not text:
            self._show_toast("标记字典为空！", "error")
            return

        new_marks = parse_marks_from_display_text(text)
        self.current_marks = new_marks
        self.mark_count_label.setText(f"共 {len(new_marks)} 个标记")

        # 写回 mark 段落
        from annotator import update_mark_paragraph_text
        mark_para = self.doc_data.get('mark_para')
        if mark_para is None:
            self._show_toast("未找到附图标记段落，无法同步到文档", "warning")
            self._log("⚠️ 文档中没有附图标记段落，仅更新内存词典")
            return

        new_text_for_para = marks_to_display_text(new_marks)
        try:
            update_mark_paragraph_text(mark_para, new_text_for_para)
        except Exception as e:
            self._show_toast(f"同步失败：{e}", "error")
            return

        self._log(f"✅ 已重新确认标记并写回内存：共 {len(new_marks)} 项")
        self._add_history(
            f"重新确认标记 ({len(new_marks)} 项)",
            f"已将编辑框中的词典写回附图标记段落：\n{new_text_for_para}",
        )
        self._show_toast("标记已同步到内存", "success")

    # ===== 辅助方法 =====

    def _validate_before_annotate(self) -> bool:
        """标注前验证"""
        if not self.doc_data:
            self._show_toast("请先打开 docx 文件！", "error")
            return False

        if not self.marks_edit.toPlainText().strip():
            self._show_toast("标记字典为空，请先输入标记！", "error")
            return False

        return True

    def _sync_marks_from_editor(self):
        """从编辑框同步标记到内存"""
        text = self.marks_edit.toPlainText().strip()
        if text:
            self.current_marks = parse_marks_from_display_text(text)
            self.mark_count_label.setText(f"共 {len(self.current_marks)} 个标记")

    def _generate_output_path(self, action_name="已标注") -> str:
        """生成输出文件路径"""
        dir_name = os.path.dirname(self.current_file_path)
        base_name = os.path.splitext(os.path.basename(self.current_file_path))[0]
        # 去掉旧后缀如果存在
        if base_name.endswith("_已标注") or base_name.endswith("_已清洗"):
            base_name = base_name.rsplit("_", 1)[0]
        output_name = f"{base_name}_{action_name}.docx"
        return os.path.join(dir_name, output_name)

    def _infer_output_action_name(self) -> str:
        """
        根据操作历史判断输出文件后缀：
          • 最近一次「标注 / 删除标记」动作决定后缀；
          • 若最后一次是「删除标记」→「_已清洗」；
          • 若最后一次是「标注」或没有相关动作 → 「_已标注」。
        历史里夹杂的清洗 / 错别字 / 权利要求书修改等不影响判断，
        只有「标注 / 删除标记」这两个互斥动作起决定作用。
        """
        for entry in reversed(self.history_entries):
            summary = (entry.get("summary") or "").strip()
            if summary.startswith("删除标记"):
                return "已清洗"
            if summary.startswith("标注"):
                return "已标注"
        return "已标注"


    def _update_section_buttons(self):
        """更新章节列表按钮"""
        # 清空现有按钮
        while self.section_buttons_layout.count():
            item = self.section_buttons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 新文档：重置预览编辑状态
        self._preview_loading = True
        self.preview_text.setPlainText("")
        self._preview_loading = False
        self._preview_section_name = None
        self._preview_start_idx = None
        self._preview_end_idx = None
        self._preview_para_count = 0
        self._preview_dirty = False
        self.preview_title.setText("📄 内容预览（可编辑）")
        self.preview_status_label.setText("")
        self.preview_confirm_btn.setEnabled(False)

        if not self.doc_data:
            return

        sections = self.doc_data['sections']
        for name, section in sections.items():
            para_count = section.end_idx - section.start_idx
            btn = QPushButton(f"  {name}  ({para_count}段)")
            btn.setObjectName("smallBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda checked, n=name: self._preview_section(n))
            self.section_buttons_layout.addWidget(btn)

    def _preview_section(self, section_name: str):
        """预览指定章节内容（可编辑）。"""
        if not self.doc_data:
            return

        sections = self.doc_data['sections']
        if section_name not in sections:
            return

        # 切换章节前提示未保存改动
        if self._preview_dirty and self._preview_section_name and \
                self._preview_section_name != section_name:
            resp = QMessageBox.question(
                self, "有未确认的改动",
                f"「{self._preview_section_name}」的预览框有未确认改动，"
                "切换章节将丢弃这些改动。是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if resp != QMessageBox.StandardButton.Yes:
                return

        section = sections[section_name]
        paragraphs = self.doc_data['paragraphs']

        # 每段一行（含空段）→ 行数严格等于段数，便于回写
        lines = []
        for i in range(section.start_idx, section.end_idx):
            text = paragraphs[i].text if paragraphs[i].text else ""
            lines.append(text)
        content = "\n".join(lines)

        self._preview_section_name = section_name
        self._preview_start_idx = section.start_idx
        self._preview_end_idx = section.end_idx
        self._preview_para_count = section.end_idx - section.start_idx

        self._preview_loading = True
        self.preview_text.setPlainText(content)
        self._preview_loading = False
        self._preview_dirty = False

        self.preview_title.setText(f"📄 {section_name}（可编辑）")
        self.preview_status_label.setText(f"共 {self._preview_para_count} 段")
        self.preview_confirm_btn.setEnabled(False)

    def _on_preview_text_changed(self):
        """预览框内容变化 → 标记 dirty，启用确认按钮。"""
        if self._preview_loading or self._preview_section_name is None:
            return
        self._preview_dirty = True
        self.preview_confirm_btn.setEnabled(True)
        self.preview_status_label.setText(
            f"共 {self._preview_para_count} 段 · 有未确认改动"
        )

    def _on_preview_confirm_edits(self) -> bool:
        """确认修改：把预览框的内容写回内存 paragraphs[start:end]。"""
        if self._preview_section_name is None or self._preview_start_idx is None:
            return True
        if not self._preview_dirty:
            return True

        text = self.preview_text.toPlainText()
        lines = text.split("\n")
        if len(lines) != self._preview_para_count:
            QMessageBox.warning(
                self, "段落数变化",
                f"预览框现有 {len(lines)} 行，而原章节为 {self._preview_para_count} 段。\n\n"
                "本功能要求每行对应一个段落。请恢复为原段数后再确认，"
                "或使用撤销 (Ctrl+Z) 回到上一状态。"
            )
            return False

        paragraphs = self.doc_data['paragraphs']
        from claim_check import set_paragraph_text
        changed_count = 0
        changed_lines = []
        for i, new_line in enumerate(lines):
            para = paragraphs[self._preview_start_idx + i]
            old_line = para.text if para.text else ""
            if old_line == new_line:
                continue
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

        self._preview_dirty = False
        self.preview_confirm_btn.setEnabled(False)
        self.preview_status_label.setText(f"共 {self._preview_para_count} 段")

        if changed_count:
            detail = "\n".join(changed_lines[:20])
            if len(changed_lines) > 20:
                detail += f"\n…（共 {len(changed_lines)} 处改动）"
            self._add_history(
                f"编辑「{self._preview_section_name}」（{changed_count} 段）",
                detail,
            )
            self._show_toast(f"已确认 {changed_count} 段修改", "success")
        else:
            self._show_toast("无实际改动", "info")
        return True


