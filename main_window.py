# -*- coding: utf-8 -*-
"""
main_window.py — PyQt6 主窗口
专利附图标记桌面软件的GUI界面。
"""
import os
import sys
import traceback
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTextEdit, QPlainTextEdit, QFileDialog, QMessageBox,
    QStatusBar, QProgressBar, QFrame, QSplitter, QTabWidget,
    QGroupBox, QApplication, QSizePolicy
)
from PyQt6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, pyqtSignal, QThread
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QDragEnterEvent, QDropEvent

from doc_parser import parse_document, get_section_text
from mark_extractor import extract_marks_from_paragraph, marks_to_display_text, parse_marks_from_display_text
from annotator import (
    smart_annotate_section, build_claims_replace_dict,
    build_implementation_replace_dict, annotate_section
)


class AnnotateWorker(QThread):
    """在后台线程中执行标注操作，避免界面卡顿"""
    finished = pyqtSignal(str, int, int)  # (消息, 权利要求替换数, 实施方式替换数)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)  # 进度百分比

    def __init__(self, doc_data: dict, marks: dict, output_path: str):
        super().__init__()
        self.doc_data = doc_data
        self.marks = marks
        self.output_path = output_path

    def run(self):
        try:
            sections = self.doc_data['sections']
            paragraphs = self.doc_data['paragraphs']
            doc = self.doc_data['document']

            claims_count = 0
            impl_count = 0

            self.progress.emit(10)

            # 标注权利要求书
            if '权利要求书' in sections:
                # 权利要求书使用 名称 → 名称（数字） 格式
                # 注意：需要确保不重复标注
                section = sections['权利要求书']
                # 先检查是否已标注
                result = smart_annotate_section(
                    paragraphs, section, self.marks, mode="claims"
                )
                if result == -1:
                    claims_count = -1  # 表示已经标注过
                else:
                    claims_count = result

            self.progress.emit(50)

            # 标注具体实施方式
            if '具体实施方式' in sections:
                section = sections['具体实施方式']
                result = smart_annotate_section(
                    paragraphs, section, self.marks, mode="implementation"
                )
                if result == -1:
                    impl_count = -1
                else:
                    impl_count = result

            self.progress.emit(80)

            # 保存文件
            doc.save(self.output_path)

            self.progress.emit(100)

            msg = self._build_result_message(claims_count, impl_count)
            self.finished.emit(msg, claims_count, impl_count)

        except Exception as e:
            self.error.emit(f"标注失败：{str(e)}\n{traceback.format_exc()}")

    def _build_result_message(self, claims_count, impl_count):
        parts = []
        if claims_count == -1:
            parts.append("权利要求书：已有标注，跳过")
        elif claims_count == 0:
            parts.append("权利要求书：未找到需要标注的内容")
        else:
            parts.append(f"权利要求书：成功标注 {claims_count} 个段落")

        if impl_count == -1:
            parts.append("具体实施方式：已有标注，跳过")
        elif impl_count == 0:
            parts.append("具体实施方式：未找到需要标注的内容")
        else:
            parts.append(f"具体实施方式：成功标注 {impl_count} 个段落")

        return "\n".join(parts)


class ToastWidget(QLabel):
    """悬浮Toast提示"""

    def __init__(self, parent, message: str, toast_type: str = "info"):
        super().__init__(message, parent)
        self.setFixedHeight(42)
        self.setMinimumWidth(280)
        self.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        # 根据类型设置颜色
        colors = {
            "info": "#00bfa5",
            "success": "#00e676",
            "error": "#ff5252",
            "warning": "#ffab40"
        }
        border_color = colors.get(toast_type, colors["info"])

        self.setStyleSheet(f"""
            QLabel {{
                background-color: rgba(30, 30, 30, 0.95);
                color: #ffffff;
                padding: 10px 20px;
                border-radius: 8px;
                border-left: 5px solid {border_color};
                font-size: 13px;
                font-weight: 500;
            }}
        """)

        # 定位到右上角
        parent_width = parent.width() if parent else 800
        self.move(parent_width - self.width() - 30, 30)
        self.show()

        # 3秒后自动消失
        QTimer.singleShot(3000, self._fade_out)

    def _fade_out(self):
        self.close()
        self.deleteLater()


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.doc_data = None         # 解析后的文档数据
        self.current_marks = {}      # 当前标记字典
        self.current_file_path = ""  # 当前打开的文件路径
        self.worker = None           # 后台工作线程

        self.setWindowTitle("📌 专利附图标记助手")
        self.setMinimumSize(1100, 750)
        self.resize(1280, 860)

        # 接受拖放
        self.setAcceptDrops(True)

        self._init_ui()

    def _init_ui(self):
        """初始化界面布局"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 16, 20, 12)
        main_layout.setSpacing(12)

        # ===== 顶部区域：标题 + 文件选择 =====
        header = self._create_header()
        main_layout.addWidget(header)

        # ===== 中间区域：标签页 =====
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("mainTabs")

        # 标签页1: 标记与标注
        tab1 = self._create_mark_tab()
        self.tab_widget.addTab(tab1, "📌 标记提取与标注")

        # 标签页2: 文档预览
        tab2 = self._create_preview_tab()
        self.tab_widget.addTab(tab2, "📄 文档内容预览")

        main_layout.addWidget(self.tab_widget, 1)

        # ===== 底部：状态栏 =====
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪 — 请选择或拖入 .docx 文件")

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

    def _create_header(self) -> QFrame:
        """创建顶部区域"""
        header = QFrame()
        header.setObjectName("headerPanel")
        layout = QVBoxLayout(header)
        layout.setSpacing(8)

        # 标题行
        title_row = QHBoxLayout()
        title_label = QLabel("专利附图标记助手")
        title_label.setObjectName("titleLabel")
        title_row.addWidget(title_label)
        title_row.addStretch()

        # 文件信息标签
        self.file_info_label = QLabel("未打开文件")
        self.file_info_label.setObjectName("subtitleLabel")
        title_row.addWidget(self.file_info_label)
        layout.addLayout(title_row)

        # 副标题
        subtitle = QLabel("直接操作 .docx 专利文件，自动提取附图标记并一键标注权利要求书和具体实施方式")
        subtitle.setObjectName("subtitleLabel")
        layout.addWidget(subtitle)

        # 文件选择按钮
        self.file_btn = QPushButton("📂  点击选择 docx 文件，或将文件拖入此区域")
        self.file_btn.setObjectName("fileBtn")
        self.file_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.file_btn.clicked.connect(self._on_select_file)
        layout.addWidget(self.file_btn)

        return header

    def _create_mark_tab(self) -> QWidget:
        """创建标记提取与标注标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(12)

        # 上方: 标记字典编辑区
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

        self.refresh_marks_btn = QPushButton("🔄 重新提取标记")
        self.refresh_marks_btn.setObjectName("smallBtn")
        self.refresh_marks_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.refresh_marks_btn.setEnabled(False)
        self.refresh_marks_btn.clicked.connect(self._on_refresh_marks)
        marks_btn_row.addWidget(self.refresh_marks_btn)

        self.mark_count_label = QLabel("")
        self.mark_count_label.setObjectName("subtitleLabel")
        marks_btn_row.addWidget(self.mark_count_label)

        marks_btn_row.addStretch()

        self.clear_marks_btn = QPushButton("🗑️ 清空标记")
        self.clear_marks_btn.setObjectName("dangerBtn")
        self.clear_marks_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_marks_btn.clicked.connect(lambda: self.marks_edit.clear())
        marks_btn_row.addWidget(self.clear_marks_btn)

        marks_layout.addLayout(marks_btn_row)
        layout.addWidget(marks_group)

        # 中间: 操作区
        action_layout = QHBoxLayout()
        action_layout.setSpacing(16)

        # 一键标注按钮
        self.annotate_btn = QPushButton("⚡  一键标注")
        self.annotate_btn.setObjectName("primaryBtn")
        self.annotate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.annotate_btn.setEnabled(False)
        self.annotate_btn.setToolTip("自动标注权利要求书（名称+括号编号）和具体实施方式（名称+数字）")
        self.annotate_btn.clicked.connect(self._on_annotate)
        action_layout.addWidget(self.annotate_btn)

        # 单独标注权利要求书
        self.annotate_claims_btn = QPushButton("📋 仅标注权利要求书")
        self.annotate_claims_btn.setObjectName("accentBtn")
        self.annotate_claims_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.annotate_claims_btn.setEnabled(False)
        self.annotate_claims_btn.clicked.connect(lambda: self._on_annotate_section("claims"))
        action_layout.addWidget(self.annotate_claims_btn)

        # 单独标注具体实施方式
        self.annotate_impl_btn = QPushButton("📝 仅标注具体实施方式")
        self.annotate_impl_btn.setObjectName("accentBtn")
        self.annotate_impl_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.annotate_impl_btn.setEnabled(False)
        self.annotate_impl_btn.clicked.connect(lambda: self._on_annotate_section("implementation"))
        action_layout.addWidget(self.annotate_impl_btn)

        layout.addLayout(action_layout)

        # 下方: 操作日志
        log_group = QGroupBox("📋 操作日志")
        log_layout = QVBoxLayout(log_group)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("标注操作的日志将在此显示...")
        log_layout.addWidget(self.log_text)

        layout.addWidget(log_group, 1)

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
        self.preview_title = QLabel("📄 内容预览")
        self.preview_title.setObjectName("sectionLabel")
        preview_title_row.addWidget(self.preview_title)
        preview_title_row.addStretch()
        preview_layout.addLayout(preview_title_row)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText("选择左侧章节查看内容...")
        preview_layout.addWidget(self.preview_text)

        splitter.addWidget(preview_panel)

        # 设置分割比例
        splitter.setSizes([220, 800])

        layout.addWidget(splitter)

        return widget

    # ===== 事件处理 =====

    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖入文件时的处理"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.docx') for url in urls):
                event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        """释放拖入文件时的处理"""
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.docx'):
                self._load_document(file_path)
                break

    # ===== 操作回调 =====

    def _on_select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择专利 docx 文件",
            "",
            "Word 文档 (*.docx);;所有文件 (*)"
        )
        if file_path:
            self._load_document(file_path)

    def _load_document(self, file_path: str):
        """加载并解析docx文档"""
        self.status_bar.showMessage(f"正在解析文档: {os.path.basename(file_path)} ...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        QApplication.processEvents()

        try:
            self.progress_bar.setValue(20)
            QApplication.processEvents()

            # 解析文档
            self.doc_data = parse_document(file_path)
            self.current_file_path = file_path

            self.progress_bar.setValue(60)
            QApplication.processEvents()

            # 更新文件信息
            filename = os.path.basename(file_path)
            self.file_info_label.setText(f"📄 {filename}")
            self.file_btn.setText(f"📂  当前文件: {filename}   (点击更换)")

            # 提取标记
            self._extract_and_display_marks()

            self.progress_bar.setValue(80)
            QApplication.processEvents()

            # 更新章节预览
            self._update_section_buttons()

            # 启用按钮
            self.annotate_btn.setEnabled(True)
            self.annotate_claims_btn.setEnabled(True)
            self.annotate_impl_btn.setEnabled(True)
            self.refresh_marks_btn.setEnabled(True)

            self.progress_bar.setValue(100)

            # 日志
            sections = self.doc_data['sections']
            section_info = "、".join(sections.keys())
            total_paras = len(self.doc_data['paragraphs'])
            self._log(f"✅ 文档解析成功: {filename}")
            self._log(f"   总段落数: {total_paras}，识别到的章节: {section_info}")

            if self.current_marks:
                self._log(f"   提取到 {len(self.current_marks)} 个附图标记")
            else:
                self._log("⚠️ 未找到附图标记，请手动在标记字典中输入")

            self.status_bar.showMessage(f"文档解析完成 — {filename}")
            self._show_toast("文档解析成功！", "success")

        except Exception as e:
            self._log(f"❌ 文档解析失败: {str(e)}")
            self._log(traceback.format_exc())
            self.status_bar.showMessage("文档解析失败")
            self._show_toast(f"解析失败: {str(e)}", "error")
            QMessageBox.critical(self, "解析失败", f"无法解析该 docx 文件:\n\n{str(e)}")

        finally:
            QTimer.singleShot(1500, lambda: self.progress_bar.setVisible(False))

    def _extract_and_display_marks(self):
        """提取标记并显示在编辑框"""
        if not self.doc_data:
            return

        mark_para = self.doc_data.get('mark_para')
        if mark_para:
            marks = extract_marks_from_paragraph(mark_para)
            self.current_marks = marks
            display = marks_to_display_text(marks)
            self.marks_edit.setPlainText(display)
            self.mark_count_label.setText(f"共 {len(marks)} 个标记")
        else:
            self.marks_edit.setPlainText("")
            self.current_marks = {}
            self.mark_count_label.setText("未找到附图标记段落")

    def _on_refresh_marks(self):
        """重新从编辑框文本解析标记"""
        text = self.marks_edit.toPlainText().strip()
        if not text:
            self._show_toast("标记字典为空！", "error")
            return

        self.current_marks = parse_marks_from_display_text(text)
        count = len(self.current_marks)
        self.mark_count_label.setText(f"共 {count} 个标记")
        self._log(f"🔄 从编辑框重新解析标记: 共 {count} 个")
        self._show_toast(f"解析成功，共 {count} 个标记", "success")

    def _on_annotate(self):
        """一键标注权利要求书和具体实施方式"""
        if not self._validate_before_annotate():
            return

        # 从编辑框重新读取标记（用户可能已编辑）
        self._sync_marks_from_editor()

        # 生成输出文件路径
        output_path = self._generate_output_path()

        self._log("=" * 50)
        self._log("🚀 开始一键标注...")
        self._log(f"   标记数量: {len(self.current_marks)}")
        self._log(f"   输出文件: {os.path.basename(output_path)}")

        # 禁用按钮
        self._set_buttons_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # 重新解析文档（因为需要对新的Document对象操作）
        self.doc_data = parse_document(self.current_file_path)

        # 启动后台标注线程
        self.worker = AnnotateWorker(self.doc_data, self.current_marks, output_path)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(lambda msg, c, i: self._on_annotate_finished(msg, output_path))
        self.worker.error.connect(self._on_annotate_error)
        self.worker.start()

    def _on_annotate_section(self, mode: str):
        """单独标注某个章节"""
        if not self._validate_before_annotate():
            return

        self._sync_marks_from_editor()

        section_name = "权利要求书" if mode == "claims" else "具体实施方式"
        if section_name not in self.doc_data['sections']:
            self._show_toast(f"未找到 {section_name} 章节！", "error")
            return

        output_path = self._generate_output_path()

        self._log("=" * 50)
        self._log(f"🚀 开始标注 {section_name}...")

        self._set_buttons_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # 重新解析
        self.doc_data = parse_document(self.current_file_path)

        self.worker = AnnotateWorker(self.doc_data, self.current_marks, output_path)
        # 修改worker只标注一个章节
        original_run = self.worker.run

        def custom_run():
            try:
                sections = self.doc_data['sections']
                paragraphs = self.doc_data['paragraphs']
                doc = self.doc_data['document']

                self.worker.progress.emit(20)

                if mode == "claims":
                    replace_dict = build_claims_replace_dict(self.current_marks)
                else:
                    replace_dict = build_implementation_replace_dict(self.current_marks)

                section = sections[section_name]
                count = annotate_section(paragraphs, section, replace_dict,
                                         marks=self.current_marks, mode=mode)

                self.worker.progress.emit(70)
                doc.save(output_path)
                self.worker.progress.emit(100)

                if count > 0:
                    msg = f"{section_name}：成功标注 {count} 个段落"
                else:
                    msg = f"{section_name}：未找到需要标注的内容（可能已标注）"

                self.worker.finished.emit(msg, count, 0)
            except Exception as e:
                self.worker.error.emit(f"标注失败：{str(e)}\n{traceback.format_exc()}")

        self.worker.run = custom_run
        self.worker.finished.connect(lambda msg, c, i: self._on_annotate_finished(msg, output_path))
        self.worker.error.connect(self._on_annotate_error)
        self.worker.start()

    def _on_annotate_finished(self, message: str, output_path: str):
        """标注完成回调"""
        self._set_buttons_enabled(True)
        self._log(f"✅ 标注完成！")
        self._log(f"   {message.replace(chr(10), chr(10) + '   ')}")
        self._log(f"   文件已保存至: {output_path}")

        self.status_bar.showMessage(f"标注完成 — {os.path.basename(output_path)}")
        self._show_toast("标注完成！文件已保存", "success")

        QTimer.singleShot(1500, lambda: self.progress_bar.setVisible(False))

        # 提示用户
        reply = QMessageBox.information(
            self, "标注完成",
            f"文件已成功标注并保存至：\n\n{output_path}\n\n{message}\n\n是否打开文件所在目录？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )

        if reply == QMessageBox.StandardButton.Yes:
            os.startfile(os.path.dirname(output_path))

    def _on_annotate_error(self, error_msg: str):
        """标注失败回调"""
        self._set_buttons_enabled(True)
        self.progress_bar.setVisible(False)
        self._log(f"❌ {error_msg}")
        self.status_bar.showMessage("标注失败")
        self._show_toast("标注失败！", "error")
        QMessageBox.critical(self, "标注失败", error_msg)

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

    def _generate_output_path(self) -> str:
        """生成输出文件路径: 原名_已标注.docx"""
        dir_name = os.path.dirname(self.current_file_path)
        base_name = os.path.splitext(os.path.basename(self.current_file_path))[0]
        output_name = f"{base_name}_已标注.docx"
        return os.path.join(dir_name, output_name)

    def _set_buttons_enabled(self, enabled: bool):
        """设置按钮状态"""
        self.annotate_btn.setEnabled(enabled)
        self.annotate_claims_btn.setEnabled(enabled)
        self.annotate_impl_btn.setEnabled(enabled)
        self.file_btn.setEnabled(enabled)

    def _update_section_buttons(self):
        """更新章节列表按钮"""
        # 清空现有按钮
        while self.section_buttons_layout.count():
            item = self.section_buttons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

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
        """预览指定章节内容"""
        if not self.doc_data:
            return

        sections = self.doc_data['sections']
        if section_name not in sections:
            return

        section = sections[section_name]
        paragraphs = self.doc_data['paragraphs']
        text = get_section_text(paragraphs, section)

        self.preview_title.setText(f"📄 {section_name}")
        self.preview_text.setPlainText(text)

    def _log(self, message: str):
        """添加日志"""
        self.log_text.append(message)
        # 滚动到底部
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _show_toast(self, message: str, toast_type: str = "info"):
        """显示Toast提示"""
        try:
            toast = ToastWidget(self, message, toast_type)
            # 重新定位到右上角
            toast.adjustSize()
            x = self.width() - toast.width() - 30
            toast.move(max(x, 10), 30)
        except Exception:
            pass  # Toast显示失败不影响主流程
