# -*- coding: utf-8 -*-
"""
main_window.py — PyQt6 主窗口
专利附图标记桌面软件的GUI界面。
"""
import os
import traceback
from datetime import datetime

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTextEdit, QPlainTextEdit, QFileDialog, QMessageBox,
    QStatusBar, QProgressBar, QFrame, QSplitter, QTabWidget,
    QGroupBox, QApplication, QCheckBox, QDialog, QSpinBox,
    QButtonGroup, QTableWidget, QTableWidgetItem, QHeaderView,
    QScrollArea, QAbstractItemView, QLineEdit,
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt6.QtGui import (
    QDragEnterEvent, QDropEvent, QTextCursor, QTextCharFormat, QColor,
)

from doc_parser import parse_document, get_section_text
from mark_extractor import extract_marks_from_paragraph, extract_marks_from_paragraphs, marks_to_display_text, parse_marks_from_display_text
from annotator import (
    smart_annotate_section, build_claims_replace_dict,
    build_implementation_replace_dict, annotate_section,
    smart_remove_section, annotate_paragraph_safe
)
from styles import DARK_THEME_QSS, LIGHT_THEME_QSS
from config_manager import AppSettings
from cleaner import (
    remove_suoshu, unify_halfwidth_punct, convert_fullwidth_to_halfwidth,
    detect_orphan_marks,
    check_typos_wordbank, check_typos_pycorrector, check_duplicate_words,
    merge_typo_results, apply_typo_corrections
)
from workers import (
    _longest_nonspace_run, _is_pycorrector_available,
    AnnotateWorker, CleanWorker, ToastWidget,
)
from tabs.claim_tab import ClaimTabMixin
from tabs.typo_tab import TypoTabMixin
from tabs.clean_tab import CleanTabMixin
from tabs.mark_tab import MarkTabMixin


class MainWindow(ClaimTabMixin, TypoTabMixin, CleanTabMixin, MarkTabMixin, QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.doc_data = None         # 解析后的文档数据
        self.current_marks = {}      # 当前标记字典
        self.current_file_path = ""  # 当前打开的文件路径
        self.worker = None           # 后台标注线程
        self.clean_worker = None     # 后台清洗线程
        self.current_theme = "light"  # 默认浅色
        self.suoshu_checkboxes = {}  # {section_name: QCheckBox}
        self.typo_data = []          # 当前错别字检查结果
        self.dup_data = []           # 当前重复字词检查结果
        self.history_entries = []    # 内存中累计的操作历史
        # 权利要求书检查 Tab 的状态
        self._claim_start_idx = None       # 权利要求书段落起始索引
        self._claim_end_idx = None         # 权利要求书段落结束索引（不含）
        self._claim_para_count = 0         # 权利要求书段落总数（用于确认时校验）
        self._claim_dirty = False          # 预览框是否有未确认修改
        self._claim_results = []           # 当前检查结果缓存
        self._claim_loaded = False         # 当前文档是否已加载过权利要求书内容
        self._claim_n = 2                  # 当前滑窗字数（按钮 / 自定义共享）
        self._claim_session_ignore = set() # 本次会话内结果行「忽略」记录（不持久化）

        # 配置管理器
        self.settings = AppSettings()

        # 读取已保存的主题（默认浅色）
        self.current_theme = self.settings.get_theme() or "light"

        self.setWindowTitle("📌 专利标记助手")
        self.setMinimumSize(1100, 750)
        self.resize(1280, 860)

        # 接受拖放
        self.setAcceptDrops(True)

        self._init_ui()

        # 应用持久化的窗口几何
        geom = self.settings.get_geometry()
        if geom is not None:
            try:
                self.restoreGeometry(geom)
            except Exception:
                pass

        # 应用持久化的主题
        try:
            app = QApplication.instance()
            if app is not None:
                app.setStyleSheet(DARK_THEME_QSS if self.current_theme == "dark" else LIGHT_THEME_QSS)
        except Exception:
            pass

        # 应用持久化的清洗 Tab 勾选状态
        try:
            self.punct_halfwidth_cb.setChecked(
                self.settings.get_bool("clean/punct_halfwidth", True)
            )
            self.punct_fullwidth_cb.setChecked(
                self.settings.get_bool("clean/punct_fullwidth", False)
            )
            self.fix_punctuation_cb.setChecked(
                self.settings.get_bool("clean/fix_consecutive_punct", True)
            )
        except Exception:
            pass

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

        # 标签页3: 文本清洗
        tab3 = self._create_clean_tab()
        self.tab_widget.addTab(tab3, "🧹 文本清洗")

        # 标签页4: 错别字检查
        tab4 = self._create_typo_tab()
        self.tab_widget.addTab(tab4, "📝 错别字检查")

        # 标签页5: 权利要求书检查（引用基础 / 引用关系 / 术语一致性）
        tab5 = self._create_claim_check_tab()
        self.tab_widget.addTab(tab5, "⚖️ 权利要求书检查")

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
        title_label = QLabel("专利标记助手")
        title_label.setObjectName("titleLabel")
        title_row.addWidget(title_label)
        
        # 主题切换按钮
        self.theme_btn = QPushButton("🌓 切换主题")
        self.theme_btn.setObjectName("smallBtn")
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.clicked.connect(self._toggle_theme)
        title_row.addWidget(self.theme_btn)

        title_row.addStretch()

        # 文件信息标签
        self.file_info_label = QLabel("未打开文件")
        self.file_info_label.setObjectName("subtitleLabel")
        title_row.addWidget(self.file_info_label)
        layout.addLayout(title_row)

        # 文件选择按钮
        self.file_btn = QPushButton("📂  点击选择 docx 文件，或将文件拖入此区域")
        self.file_btn.setObjectName("fileBtn")
        self.file_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.file_btn.clicked.connect(self._on_select_file)
        layout.addWidget(self.file_btn)

        return header



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
        start_dir = ""
        try:
            start_dir = self.settings.get_last_dir() or ""
        except Exception:
            pass
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择专利 docx 文件",
            start_dir,
            "Word 文档 (*.docx);;所有文件 (*)"
        )
        if file_path:
            try:
                self.settings.set_last_dir(os.path.dirname(file_path))
            except Exception:
                pass
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

            # 更新清洗Tab的章节勾选框
            self._update_suoshu_section_checkboxes(self.doc_data['sections'])

            # 启用按钮
            self.annotate_btn.setEnabled(True)
            self.annotate_claims_btn.setEnabled(True)
            self.annotate_impl_btn.setEnabled(True)
            self.remove_marks_btn.setEnabled(True)
            self.refresh_marks_btn.setEnabled(True)
            self.confirm_marks_btn.setEnabled(True)
            self.suoshu_btn.setEnabled(True)
            self.punct_btn.setEnabled(True)
            self.orphan_btn.setEnabled(True)
            self.typo_check_btn.setEnabled(True)
            self.dup_check_btn.setEnabled(True)

            # 把权利要求书内容加载到新 Tab
            self._claim_tab_load_from_doc()

            # 加载新文档时清空历史与禁用「文件生成」
            self._clear_history()

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

    # ===== 操作历史 =====

    def _add_history(self, summary: str, detail: str = ""):
        """向历史框追加一条记录"""
        ts = datetime.now().strftime("%H:%M:%S")
        entry = {"time": ts, "summary": summary, "detail": detail}
        self.history_entries.append(entry)
        self._render_history()
        # 文件生成按钮：有了历史就启用
        self.generate_btn.setEnabled(True)

    def _render_history(self):
        """重新渲染历史框"""
        if not self.history_entries:
            self.history_text.clear()
            return
        lines = []
        for i, e in enumerate(self.history_entries, 1):
            lines.append(f"<b>#{i}  [{e['time']}]  {e['summary']}</b>")
            if e.get("detail"):
                for d in e["detail"].split("\n"):
                    if d.strip():
                        lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;{d}")
            lines.append("")
        self.history_text.setHtml("<br>".join(lines))
        sb = self.history_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _clear_history(self):
        self.history_entries = []
        self.history_text.clear()
        self.generate_btn.setEnabled(False)

    # ===== 标记同步 =====

    def _set_buttons_enabled(self, enabled: bool):
        """设置标注操作按钮状态"""
        self.annotate_btn.setEnabled(enabled)
        self.annotate_claims_btn.setEnabled(enabled)
        self.annotate_impl_btn.setEnabled(enabled)
        self.remove_marks_btn.setEnabled(enabled)
        self.confirm_marks_btn.setEnabled(enabled)
        self.refresh_marks_btn.setEnabled(enabled)
        self.file_btn.setEnabled(enabled)
        # generate_btn 仅在有历史时启用
        if enabled and self.history_entries:
            self.generate_btn.setEnabled(True)
        elif not enabled:
            self.generate_btn.setEnabled(False)

    def _toggle_theme(self):
        """切换深色/浅色主题"""
        app = QApplication.instance()
        if self.current_theme == "dark":
            app.setStyleSheet(LIGHT_THEME_QSS)
            self.current_theme = "light"
        else:
            app.setStyleSheet(DARK_THEME_QSS)
            self.current_theme = "dark"
        try:
            self.settings.set_theme(self.current_theme)
        except Exception:
            pass
        self._show_toast(f"已切换为{'浅色' if self.current_theme == 'light' else '深色'}主题")

    def closeEvent(self, event):
        """窗口关闭时持久化配置"""
        try:
            self.settings.set_theme(self.current_theme)
            self.settings.set_geometry(self.saveGeometry())
            if hasattr(self, "punct_halfwidth_cb"):
                self.settings.set_bool("clean/punct_halfwidth", self.punct_halfwidth_cb.isChecked())
            if hasattr(self, "punct_fullwidth_cb"):
                self.settings.set_bool("clean/punct_fullwidth", self.punct_fullwidth_cb.isChecked())
            if hasattr(self, "fix_punctuation_cb"):
                self.settings.set_bool("clean/fix_consecutive_punct", self.fix_punctuation_cb.isChecked())
            self.settings.sync()
        except Exception:
            pass
        super().closeEvent(event)

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

    # ===== 文本清洗 =====


    # ═══════════════════════════════════════════════════
    # 权利要求书检查 Tab — 槽函数
    # ═══════════════════════════════════════════════════

