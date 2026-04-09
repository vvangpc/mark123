# -*- coding: utf-8 -*-
"""
main_window.py — PyQt6 主窗口
专利附图标记桌面软件的GUI界面。
"""
import os
import traceback
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTextEdit, QPlainTextEdit, QFileDialog, QMessageBox,
    QStatusBar, QProgressBar, QFrame, QSplitter, QTabWidget,
    QGroupBox, QApplication, QCheckBox, QDialog, QSpinBox,
    QButtonGroup, QTableWidget, QTableWidgetItem, QHeaderView,
    QScrollArea, QAbstractItemView,
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt6.QtGui import (
    QDragEnterEvent, QDropEvent, QTextCursor, QTextCharFormat, QColor,
)

from doc_parser import parse_document, get_section_text
from mark_extractor import extract_marks_from_paragraph, marks_to_display_text, parse_marks_from_display_text
from annotator import (
    smart_annotate_section, build_claims_replace_dict,
    build_implementation_replace_dict, annotate_section,
    smart_remove_section, annotate_paragraph_safe
)
from styles import DARK_THEME_QSS, LIGHT_THEME_QSS
from cleaner import (
    remove_suoshu, unify_halfwidth_punct, convert_fullwidth_to_halfwidth,
    detect_orphan_marks,
    check_typos_wordbank, check_typos_pycorrector, check_duplicate_words,
    merge_typo_results, apply_typo_corrections
)

# 「删除所述」功能允许处理的章节白名单
SUOSHU_ALLOWED_SECTIONS = ("权利要求书", "背景技术", "具体实施方式")
# 默认勾选的章节
SUOSHU_DEFAULT_CHECKED = ("具体实施方式",)


def _longest_nonspace_run(s: str) -> str:
    """从字符串中抽取最长的一段连续非空白字符（用作搜索锚点）。"""
    if not s:
        return ""
    best = ""
    cur_start = -1
    for i, ch in enumerate(s):
        if ch.isspace():
            if cur_start >= 0:
                seg = s[cur_start:i]
                if len(seg) > len(best):
                    best = seg
                cur_start = -1
        else:
            if cur_start < 0:
                cur_start = i
    if cur_start >= 0:
        seg = s[cur_start:]
        if len(seg) > len(best):
            best = seg
    return best


def _is_pycorrector_available() -> bool:
    """检测 pycorrector 是否已安装"""
    try:
        import pycorrector  # noqa: F401
        return True
    except ImportError:
        return False


class AnnotateWorker(QThread):
    """在后台线程中执行标注/清除操作（仅修改内存，不保存文件）"""
    finished = pyqtSignal(str, str)  # (历史摘要, 详细消息)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, doc_data: dict, marks: dict, action: str, scope: str = "all"):
        """
        action: "add" 或 "remove"
        scope: "all" / "claims" / "implementation"
        """
        super().__init__()
        self.doc_data = doc_data
        self.marks = marks
        self.action = action
        self.scope = scope

    def run(self):
        try:
            sections = self.doc_data['sections']
            paragraphs = self.doc_data['paragraphs']

            claims_count = 0
            impl_count = 0
            self.progress.emit(20)

            do_claims = self.scope in ("all", "claims")
            do_impl = self.scope in ("all", "implementation")

            if do_claims and '权利要求书' in sections:
                section = sections['权利要求书']
                if self.action == "add":
                    claims_count = smart_annotate_section(paragraphs, section, self.marks, mode="claims")
                else:
                    claims_count = smart_remove_section(paragraphs, section, self.marks, mode="claims")

            self.progress.emit(60)

            if do_impl and '具体实施方式' in sections:
                section = sections['具体实施方式']
                if self.action == "add":
                    impl_count = smart_annotate_section(paragraphs, section, self.marks, mode="implementation")
                else:
                    impl_count = smart_remove_section(paragraphs, section, self.marks, mode="implementation")

            self.progress.emit(100)

            summary, detail = self._build_messages(claims_count, impl_count)
            self.finished.emit(summary, detail)

        except Exception as e:
            self.error.emit(f"操作失败：{str(e)}\n{traceback.format_exc()}")

    def _build_messages(self, claims_count, impl_count):
        action_name = "标注" if self.action == "add" else "删除标记"
        parts = []

        if self.scope in ("all", "claims"):
            if claims_count == -1:
                parts.append("权利要求书：已有标注，跳过")
            elif claims_count == 0:
                parts.append(f"权利要求书：未找到需{action_name}的内容")
            else:
                parts.append(f"权利要求书：成功{action_name} {claims_count} 段")

        if self.scope in ("all", "implementation"):
            if impl_count == -1:
                parts.append("具体实施方式：已有标注，跳过")
            elif impl_count == 0:
                parts.append(f"具体实施方式：未找到需{action_name}的内容")
            else:
                parts.append(f"具体实施方式：成功{action_name} {impl_count} 段")

        scope_name = {"all": "全文", "claims": "权利要求书", "implementation": "具体实施方式"}[self.scope]
        summary = f"{action_name}（{scope_name}）"
        return summary, "\n".join(parts)


class CleanWorker(QThread):
    """在后台线程中执行清洗/检查操作"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)
    typo_results = pyqtSignal(list)  # 仅 typo_check 动作使用

    def __init__(self, doc_data: dict, action: str, **kwargs):
        super().__init__()
        self.doc_data = doc_data
        self.action = action
        self.kwargs = kwargs  # 额外参数按 action 传入

    def run(self):
        try:
            paragraphs = self.doc_data['paragraphs']
            sections = self.doc_data['sections']

            self.progress.emit(10)

            if self.action == "suoshu":
                selected = self.kwargs.get("selected_sections", [])
                count = remove_suoshu(paragraphs, sections, selected)
                self.progress.emit(100)
                self.finished.emit(f"删除「所述」完成，共处理 {count} 个段落")

            elif self.action == "punct":
                do_half = self.kwargs.get("do_halfwidth", True)
                do_full = self.kwargs.get("do_fullwidth", False)
                do_consec = self.kwargs.get("do_consecutive", True)
                half_n = full_n = consec_n = 0

                # 顺序：1) 半角→全角  2) 全角→半角(可选)  3) 修正连续重复
                # 把 1/2 放在 3 之前可避免出现 ".。" 这种混合连续无法被修正
                if do_half:
                    half_n = unify_halfwidth_punct(paragraphs, sections)
                self.progress.emit(40)
                if do_full:
                    full_n = convert_fullwidth_to_halfwidth(paragraphs, sections)
                self.progress.emit(70)
                if do_consec:
                    from cleaner import fix_consecutive_punct
                    consec_n = fix_consecutive_punct(paragraphs, sections=sections)
                self.progress.emit(100)

                parts = []
                if do_half:
                    parts.append(f"半角→全角 {half_n} 段")
                if do_full:
                    parts.append(f"全角→半角 {full_n} 段")
                if do_consec:
                    parts.append(f"修正连续标点 {consec_n} 段")
                self.finished.emit("标点检查完成：" + "，".join(parts))

            elif self.action == "orphan":
                marks = self.kwargs.get("marks", {})
                orphans = detect_orphan_marks(paragraphs, sections, marks)
                self.progress.emit(60)
                from cleaner import detect_orphan_figures
                missing_figs = detect_orphan_figures(paragraphs, sections)
                self.progress.emit(100)

                parts = []
                if orphans:
                    lines = [f"  {num} — {name}" for num, name in orphans]
                    parts.append(
                        "⚠️ 孤立附图标记（附图说明有、具体实施方式无）：\n"
                        + "\n".join(lines)
                    )
                else:
                    parts.append("✅ 附图标记：所有标记均在具体实施方式中出现")

                if missing_figs:
                    fig_lines = "、".join(f"图{n}" for n in missing_figs)
                    parts.append(
                        f"⚠️ 未引用图编号（附图说明提及但具体实施方式未出现）：{fig_lines}"
                    )
                else:
                    parts.append("✅ 图编号：附图说明中的图编号均在具体实施方式中出现")

                msg = "\n".join(parts)
                self.finished.emit(msg)

            elif self.action == "typo_check":
                wb_results = check_typos_wordbank(paragraphs, sections)
                self.progress.emit(50)
                pc_results = check_typos_pycorrector(paragraphs, sections)
                self.progress.emit(90)
                merged = merge_typo_results(wb_results, pc_results)
                self.progress.emit(100)
                self.typo_results.emit(merged)
                count = len(merged)
                self.finished.emit(f"错别字检查完成，发现 {count} 处疑似问题")

            elif self.action == "dup_check":
                ignore_list = self.kwargs.get("ignore_list", [])
                dup_results = check_duplicate_words(paragraphs, sections, ignore_list=ignore_list)
                self.progress.emit(100)
                self.typo_results.emit(dup_results)
                count = len(dup_results)
                self.finished.emit(f"重复字词检查完成，发现 {count} 处疑似问题")

            elif self.action == "typo_apply":
                corrections = self.kwargs.get("corrections", [])
                count = apply_typo_corrections(paragraphs, corrections)
                self.progress.emit(100)
                self.finished.emit(f"已应用 {count} 处修正")

        except Exception as e:
            import traceback as tb
            self.error.emit(f"操作失败：{str(e)}\n{tb.format_exc()}")


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
        from config_manager import AppSettings
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
            from styles import DARK_THEME_QSS, LIGHT_THEME_QSS
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

    def _create_typo_tab(self) -> QWidget:
        """创建错别字检查标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(10)

        # ── 顶部：引擎状态条（两个引擎同一行）─────────────
        engine_group = QGroupBox("🧠 检查引擎状态")
        engine_layout = QHBoxLayout(engine_group)
        engine_layout.setSpacing(20)

        wb_count = self._get_wordbank_count()
        self.wb_label = QLabel()
        self.wb_label.setTextFormat(Qt.TextFormat.RichText)
        self.wb_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.wb_label.setToolTip("点击打开词库编辑器，可添加 / 修改 / 删除自定义词条")
        self.wb_label.setText(
            f"✅  内置词库引擎  —  已加载 <b>{wb_count}</b> 条规则  "
            f"（<a href='#edit'>点击编辑词库</a>）"
        )
        self.wb_label.linkActivated.connect(lambda _: self._on_open_wordbank_dialog())
        self.wb_label.mousePressEvent = lambda e: self._on_open_wordbank_dialog()
        engine_layout.addWidget(self.wb_label)

        # 注：「离线 NLP 引擎（pycorrector）」入口已隐藏到错别字词库编辑器内
        # （专业用户三级菜单），主界面不再展示，避免给小白用户造成困扰。

        # 重复字词忽略词库入口
        self.dup_ignore_label = QLabel()
        self.dup_ignore_label.setTextFormat(Qt.TextFormat.RichText)
        self.dup_ignore_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dup_ignore_label.setToolTip("点击打开「重复字词忽略词库」编辑器")
        self.dup_ignore_label.mousePressEvent = lambda e: self._on_open_dup_ignore_dialog()
        self._refresh_dup_ignore_label()
        engine_layout.addWidget(self.dup_ignore_label)

        engine_layout.addStretch()
        layout.addWidget(engine_group)

        # ── 中间：合并的检查区（错别字 / 重复字词共用一张表）─
        action_group = QGroupBox("📝 错别字 / 重复字词检查")
        action_v = QVBoxLayout(action_group)

        hint = QLabel(
            "在「建议修改」列编辑后，点「应用所有修改」写入内存；"
            "切回「标注」页点「💾 文件生成」生成最终文件。"
        )
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        action_v.addWidget(hint)

        # 同一行：错别字检查 / 重复字词检查 / 计数 / 应用所有修改
        btn_row = QHBoxLayout()

        self.typo_check_btn = QPushButton("🔍  错别字检查")
        self.typo_check_btn.setObjectName("accentBtn")
        self.typo_check_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.typo_check_btn.setEnabled(False)
        self.typo_check_btn.setToolTip("已检查过的结果会缓存，再次点击可在两种模式间切换显示")
        self.typo_check_btn.clicked.connect(self._on_typo_check)
        btn_row.addWidget(self.typo_check_btn)

        self.dup_check_btn = QPushButton("🔁  重复字词检查")
        self.dup_check_btn.setObjectName("accentBtn")
        self.dup_check_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.dup_check_btn.setEnabled(False)
        self.dup_check_btn.setToolTip("已检查过的结果会缓存，再次点击可在两种模式间切换显示")
        self.dup_check_btn.clicked.connect(self._on_dup_check)
        btn_row.addWidget(self.dup_check_btn)

        self.typo_count_label = QLabel("")
        self.typo_count_label.setObjectName("subtitleLabel")
        btn_row.addWidget(self.typo_count_label)

        btn_row.addStretch()

        self.typo_apply_btn = QPushButton("✅  应用所有修改")
        self.typo_apply_btn.setObjectName("primaryBtn")
        self.typo_apply_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.typo_apply_btn.setEnabled(False)
        self.typo_apply_btn.clicked.connect(self._on_apply_corrections)
        btn_row.addWidget(self.typo_apply_btn)
        action_v.addLayout(btn_row)

        # 单一结果表格，两类检查共用
        self.typo_table = QTableWidget(0, 4)
        self.typo_table.setHorizontalHeaderLabels(["章节", "原文片段", "建议修改", "操作"])
        self.typo_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.typo_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.typo_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.typo_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.typo_table.setAlternatingRowColors(True)
        self.typo_table.verticalHeader().setVisible(False)
        action_v.addWidget(self.typo_table, 1)

        layout.addWidget(action_group, 1)

        # 当前显示的检查类型："typo" 或 "dup" 或 None
        self._current_check_kind = None
        return widget

    # ─────────────────────────────────────────
    # 权利要求书检查 Tab
    # ─────────────────────────────────────────
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

        self.claim_ignore_btn = QPushButton("📕  不确定用语词库")
        self.claim_ignore_btn.setObjectName("smallBtn")
        self.claim_ignore_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_ignore_btn.setToolTip(
            "编辑权利要求书中「不应出现」的不确定 / 含糊用语词库\n"
            "（约、大概、可能、优选…），可增删、导入导出、恢复内置。"
        )
        self.claim_ignore_btn.clicked.connect(self._on_claim_ignore_dialog)
        toolbar.addWidget(self.claim_ignore_btn)

        # 「术语不一致」检查的开关：该检查噪音较大，默认关闭
        self.claim_term_cb = QCheckBox("检查术语不一致")
        self.claim_term_cb.setChecked(False)
        self.claim_term_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_term_cb.setToolTip(
            "启用「同一术语多种写法」检查。\n"
            "该检查以『所述』后面的 N 字术语作为锚点，可能产生一定噪音，\n"
            "默认不启用；需要做权利要求书术语一致性复盘时再勾选。"
        )
        toolbar.addWidget(self.claim_term_cb)

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
        # 上下文 / 说明 两列改为 Interactive：用户可以左右拖动分界线调节宽度
        h.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        h.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        h.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        # 给一个合理的初始宽度，避免 Interactive 模式下列宽为 0
        self.claim_result_table.setColumnWidth(2, 260)
        self.claim_result_table.setColumnWidth(3, 320)
        h.setStretchLastSection(False)
        self.claim_result_table.setAlternatingRowColors(True)
        self.claim_result_table.verticalHeader().setVisible(False)
        # 双击「上下文」格（列 2） → 跳转并高亮左侧预览框对应位置
        self.claim_result_table.cellDoubleClicked.connect(
            self._on_claim_result_double_clicked
        )
        right_layout.addWidget(self.claim_result_table, 1)

        self.claim_splitter.addWidget(right_panel)
        self.claim_splitter.setSizes([560, 640])

        outer.addWidget(self.claim_splitter, 1)
        return widget

    def _get_wordbank_count(self) -> int:
        """读取当前生效词库条目数（合并内置 + 用户自定义）"""
        try:
            from config_manager import get_merged_wordbank
            return len(get_merged_wordbank())
        except Exception:
            try:
                from typo_wordbank import WORDBANK
                return len(WORDBANK)
            except Exception:
                return 0

    def _refresh_wordbank_label(self):
        """重新计算词库条目数并刷新右上角标签"""
        if not hasattr(self, "wb_label"):
            return
        wb_count = self._get_wordbank_count()
        self.wb_label.setText(
            f"✅  内置词库引擎  —  已加载 <b>{wb_count}</b> 条规则  "
            f"（<a href='#edit'>点击编辑词库</a>）"
        )

    def _on_open_wordbank_dialog(self):
        """打开词库编辑对话框"""
        try:
            from wordbank_dialog import WordbankDialog
        except Exception as e:
            QMessageBox.critical(self, "无法打开", f"加载词库编辑器失败：\n{e}")
            return
        dlg = WordbankDialog(self)
        dlg.exec()
        # 无论保存与否都刷新一次（用户可能取消但已删除若干行；保险起见重读）
        self._refresh_wordbank_label()
        # 失效错别字检查的缓存 —— 下次点「错别字检查」时强制重新扫描，
        # 确保词库修改立即生效，而不是显示旧的陈旧结果
        self._invalidate_typo_cache()

    def _refresh_dup_ignore_label(self):
        """刷新重复字词忽略词库标签"""
        if not hasattr(self, "dup_ignore_label"):
            return
        try:
            from config_manager import load_dup_ignore_list
            n = len(load_dup_ignore_list())
        except Exception:
            n = 0
        self.dup_ignore_label.setText(
            f"🙈  重复字词忽略词库  —  <b>{n}</b> 条  "
            f"（<a href='#edit'>点击编辑</a>）"
        )

    def _on_open_dup_ignore_dialog(self):
        """打开「重复字词忽略词库」编辑对话框"""
        try:
            from dup_ignore_dialog import DupIgnoreDialog
        except Exception as e:
            QMessageBox.critical(self, "无法打开", f"加载忽略词库编辑器失败：\n{e}")
            return
        dlg = DupIgnoreDialog(self)
        dlg.exec()
        self._refresh_dup_ignore_label()
        # 失效重复字词检查的缓存 —— 下次点「重复字词检查」时强制重新扫描，
        # 确保新添加的忽略词立即生效
        self._invalidate_dup_cache()

    def _invalidate_typo_cache(self):
        """清空错别字检查缓存，若当前显示的就是错别字则清空表格"""
        self.typo_data = []
        if getattr(self, "_current_check_kind", None) == "typo":
            self._render_table_from_data([])

    def _invalidate_dup_cache(self):
        """清空重复字词检查缓存，若当前显示的就是重复字词则清空表格"""
        self.dup_data = []
        if getattr(self, "_current_check_kind", None) == "dup":
            self._render_table_from_data([])

    def _on_open_pycorrector_dialog(self):
        """弹出 pycorrector 安装 / 模型 / 教程说明卡片"""
        installed = _is_pycorrector_available()
        dlg = QDialog(self)
        dlg.setWindowTitle("离线 NLP 引擎 — pycorrector 说明")
        dlg.setMinimumSize(640, 520)

        v = QVBoxLayout(dlg)
        v.setSpacing(10)

        # 状态行
        status = QLabel()
        status.setTextFormat(Qt.TextFormat.RichText)
        if installed:
            status.setText("<b style='color:#00bfa5;'>✅ 当前状态：pycorrector 已安装</b>")
        else:
            status.setText("<b style='color:#ffab40;'>⚠️ 当前状态：pycorrector 未安装</b>")
        v.addWidget(status)

        # 教程区（可复制）
        tutorial_label = QLabel("以下命令可直接复制到 PowerShell / CMD / 终端中执行：")
        tutorial_label.setObjectName("subtitleLabel")
        v.addWidget(tutorial_label)

        tutorial_text = QPlainTextEdit()
        tutorial_text.setReadOnly(False)  # 设置可编辑以便复制（不可保存）
        tutorial_text.setPlainText(
            "═══════════════════════════════════════════════════\n"
            "  pycorrector 安装教程（零基础版 · 跟着做即可）\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "▎它是什么？\n"
            "  pycorrector 是一个免费的中文错别字检测库。安装后本软件会用它做\n"
            "  更深度的错别字检查。不安装也不影响基本使用，只是只能用内置词库。\n"
            "\n"
            "═══════════════════════════════════════════════════\n"
            "  第 1 步：先确认电脑上有 Python\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  ① 同时按下键盘的【Win 键】+ 【R 键】（Win 键就是 Ctrl 旁边那个\n"
            "     带 Windows 标志的键）。\n"
            "  ② 弹出「运行」小窗口，里面输入：\n"
            "         cmd\n"
            "     然后按【回车】。会出现一个黑底白字的窗口（叫「命令提示符」）。\n"
            "  ③ 在黑窗口里把下面这行复制粘贴进去，按【回车】：\n"
            "\n"
            "         python --version\n"
            "\n"
            "  ④ 看返回结果：\n"
            "     • 如果显示「Python 3.x.x」（比如 Python 3.11.5）→ 已安装，\n"
            "       直接跳到「第 2 步」。\n"
            "     • 如果提示「不是内部或外部命令」「找不到 python」→ 还没装，\n"
            "       请先去官网下载安装：\n"
            "           https://www.python.org/downloads/\n"
            "       下载后双击安装包，安装界面最下方一定要勾选\n"
            "       【Add Python to PATH】这一项，然后一路点 Next 装完。\n"
            "       装完后关掉黑窗口，重新做一次第 1 步。\n"
            "\n"
            "═══════════════════════════════════════════════════\n"
            "  第 2 步：安装 PyTorch (CPU版) 与 pycorrector\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  在同一个黑窗口里，先复制粘贴下面这行并按【回车】（安装底层依赖）：\n"
            "\n"
            "         pip install torch --index-url https://download.pytorch.org/whl/cpu\n"
            "\n"
            "  等待安装完成（因为没有显卡驱动包，速度较快），接着再执行下面这行：\n"
            "\n"
            "         pip install pycorrector -i https://pypi.tuna.tsinghua.edu.cn/simple\n"
            "\n"
            "  说明：看到最后出现 Successfully installed pycorrector-x.x.x 就表示装好了。\n"
            "\n"
            "═══════════════════════════════════════════════════\n"
            "  第 3 步：验证安装成功\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  在黑窗口里运行：\n"
            "\n"
            "         python -c \"import pycorrector; print('OK')\"\n"
            "\n"
            "  • 如果只输出一行 OK，说明安装成功。\n"
            "  • 如果报错 ModuleNotFoundError，说明第 2 步没装成功，重做第 2 步。\n"
            "\n"
            "═══════════════════════════════════════════════════\n"
            "  第 4 步：让本软件认到它\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  ① 关掉本软件（直接点右上角 ×）。\n"
            "  ② 重新打开本软件。\n"
            "  ③ 不需要查看任何状态 —— 安装成功后，错别字检查会自动\n"
            "     额外调用 NLP 引擎进行更深度的检测。\n"
            "\n"
            "  ※ 也可以不重启软件，点本卡片左下角的【🔄 重新检测安装状态】\n"
            "     按钮立即刷新；下方状态行变绿即表示已识别。\n"
            "═══════════════════════════════════════════════════\n"
            "  附：模型文件存在哪？怎么清理？\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  pycorrector 第一次运行时会自动下载几百 MB 的语言模型。\n"
            "  默认存放路径（Windows）：\n"
            "         C:\\Users\\你的用户名\\.pycorrector\n"
            "\n"
            "  如果模型坏了想重下，把上面那个文件夹整个删掉，下次运行就会\n"
            "  自动重新下载。也可以在黑窗口里运行：\n"
            "\n"
            "         rmdir /s /q %USERPROFILE%\\.pycorrector\n"
            "\n"
            "═══════════════════════════════════════════════════\n"
            "  附：怎么卸载 pycorrector？\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  在黑窗口里运行：\n"
            "\n"
            "         pip uninstall pycorrector\n"
            "\n"
            "  跳出「Proceed (Y/n)?」就输入 y 再按回车。\n"
            "\n"
            "═══════════════════════════════════════════════════\n"
            "  常见问题\n"
            "═══════════════════════════════════════════════════\n"
            "\n"
            "  Q：报错「pip 不是内部或外部命令」？\n"
            "  A：第 1 步装 Python 时没勾选【Add Python to PATH】。\n"
            "     重新装一次 Python，记得勾上。\n"
            "\n"
            "  Q：安装很慢、一直转圈？\n"
            "  A：换网络试试，或确认命令里有 -i 清华镜像那一段。\n"
            "\n"
            "  Q：装完软件还是显示「未安装」？\n"
            "  A：先点本卡片【🔄 重新检测安装状态】；如果还不行，\n"
            "     完全关闭软件再重新打开。"
        )
        tutorial_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        v.addWidget(tutorial_text, 1)

        # 操作按钮行
        btn_row = QHBoxLayout()

        copy_btn = QPushButton("📋 复制全部")
        copy_btn.clicked.connect(lambda: (
            QApplication.clipboard().setText(tutorial_text.toPlainText()),
            self._show_toast("已复制到剪贴板", "success"),
        ))
        btn_row.addWidget(copy_btn)

        reload_btn = QPushButton("🔄 重新检测安装状态")
        reload_btn.setToolTip("不重启软件即可重新探测 pycorrector 是否可导入")
        def _reload():
            ok = _is_pycorrector_available()
            if ok:
                status.setText("<b style='color:#00bfa5;'>✅ 当前状态：pycorrector 已安装</b>")
            else:
                status.setText("<b style='color:#ffab40;'>⚠️ 当前状态：pycorrector 未安装</b>")
        reload_btn.clicked.connect(_reload)
        btn_row.addWidget(reload_btn)

        btn_row.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(close_btn)

        v.addLayout(btn_row)

        dlg.exec()

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
        """重新从原始 docx 文档提取标记到编辑框（不影响内存修改）"""
        if not self.current_file_path:
            self._show_toast("请先打开文档！", "error")
            return
        try:
            fresh = parse_document(self.current_file_path)
        except Exception as e:
            self._show_toast(f"重新解析失败：{e}", "error")
            return
        mark_para = fresh.get('mark_para')
        if not mark_para:
            self._show_toast("原始文档中未找到附图标记段落", "warning")
            self.marks_edit.setPlainText("")
            self.current_marks = {}
            self.mark_count_label.setText("未找到附图标记段落")
            return

        marks = extract_marks_from_paragraph(mark_para)
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

        action_name = self._infer_output_action_name()
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

    # ===== 操作历史 =====

    def _add_history(self, summary: str, detail: str = ""):
        """向历史框追加一条记录"""
        from datetime import datetime
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

    # ===== 文本清洗 =====

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
    def _on_typo_check(self):
        """点击「错别字检查」：切换显示错别字结果；首次点击会跑一次扫描"""
        # 先把当前表格上的编辑回写到对应缓存（避免切换时丢失）
        self._snapshot_table_to_active_cache()
        if self.typo_data:
            # 已有缓存：直接切换显示，不重新扫描
            self._current_check_kind = "typo"
            self._render_table_from_data(self.typo_data)
        else:
            # 启动后台扫描；finished 回调里会自动渲染
            self._current_check_kind = "typo"
            self._start_clean_worker("typo_check", "错别字检查")

    def _on_dup_check(self):
        """点击「重复字词检查」：切换显示重复字词结果；首次点击会跑一次扫描"""
        self._snapshot_table_to_active_cache()
        if self.dup_data:
            self._current_check_kind = "dup"
            self._render_table_from_data(self.dup_data)
        else:
            self._current_check_kind = "dup"
            try:
                from config_manager import load_dup_ignore_list
                ignore_list = load_dup_ignore_list()
            except Exception:
                ignore_list = []
            self._start_clean_worker("dup_check", "重复字词检查", ignore_list=ignore_list)

    def _on_typo_results_ready(self, results: list):
        self.typo_data = results
        if self._current_check_kind == "typo":
            self._render_table_from_data(results)

    def _on_dup_results_ready(self, results: list):
        self.dup_data = results
        if self._current_check_kind == "dup":
            self._render_table_from_data(results)

    def _on_apply_corrections(self):
        """单一应用按钮：把当前显示的检查结果应用到内存"""
        if not self.doc_data or self._current_check_kind is None:
            self._show_toast("请先点击一种检查按钮！", "warning")
            return

        # ① 把表格中的用户编辑回写到对应缓存（持久化建议）
        self._snapshot_table_to_active_cache()
        data = self.typo_data if self._current_check_kind == "typo" else self.dup_data
        label_prefix = "错别字修正" if self._current_check_kind == "typo" else "重复字词修正"

        # ② 从缓存中收集所有需要修改的项
        corrections = []
        for item in data:
            confirmed = (item.get("suggestion") or "").strip()
            wrong = item.get("wrong", "")
            if not (wrong and confirmed) or confirmed == wrong:
                continue
            if item.get("_ignored"):
                continue
            corrections.append({
                "para_idx": item["para_idx"],
                "wrong": wrong,
                "confirmed_fix": confirmed,
            })

        if not corrections:
            self._show_toast("没有可应用的修正！", "warning")
            return

        label = f"{label_prefix} ({len(corrections)} 处)"
        self._start_clean_worker("typo_apply", label, history_label=label,
                                  corrections=corrections)

    def _on_check_ignore_row(self):
        """点击「忽略」按钮：从表格与缓存中移除该行"""
        sender = self.sender()
        if sender is None:
            return
        target_row = -1
        for r in range(self.typo_table.rowCount()):
            if self.typo_table.cellWidget(r, 3) is sender:
                target_row = r
                break
        if target_row < 0:
            return
        # 先回写当前所有编辑，再移除
        self._snapshot_table_to_active_cache()
        data = self._active_cache_list()
        if 0 <= target_row < len(data):
            data.pop(target_row)
        self._render_table_from_data(data)

    def _active_cache_list(self) -> list:
        """返回当前模式对应的缓存列表"""
        if self._current_check_kind == "typo":
            return self.typo_data
        if self._current_check_kind == "dup":
            return self.dup_data
        return []

    def _snapshot_table_to_active_cache(self):
        """把表格 column-2（建议修改）的当前值写回到对应缓存的 suggestion 字段
        —— 这样切换检查类型 / 重渲染时，用户已编辑的内容不会丢失。
        """
        if self._current_check_kind is None:
            return
        data = self._active_cache_list()
        for r in range(self.typo_table.rowCount()):
            if r >= len(data):
                break
            fix_item = self.typo_table.item(r, 2)
            if fix_item is not None:
                data[r]["suggestion"] = fix_item.text()

    def _render_table_from_data(self, results: list):
        """把缓存列表渲染到共用表格"""
        self.typo_table.setRowCount(0)
        for row_idx, item in enumerate(results):
            self.typo_table.insertRow(row_idx)

            # 列0：章节名（para_idx 存 UserRole）
            section_text = item.get("section") or "（未归类）"
            pos_item = QTableWidgetItem(section_text)
            pos_item.setFlags(pos_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            pos_item.setData(Qt.ItemDataRole.UserRole, item["para_idx"])
            self.typo_table.setItem(row_idx, 0, pos_item)

            # 列1：原文片段
            ctx_item = QTableWidgetItem(item["context"])
            ctx_item.setFlags(ctx_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.typo_table.setItem(row_idx, 1, ctx_item)

            # 列2：建议修改（可编辑）
            fix_item = QTableWidgetItem(item.get("suggestion", ""))
            self.typo_table.setItem(row_idx, 2, fix_item)

            # 列3：忽略按钮
            ignore_btn = QPushButton("忽略")
            ignore_btn.setObjectName("smallBtn")
            ignore_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            ignore_btn.clicked.connect(self._on_check_ignore_row)
            self.typo_table.setCellWidget(row_idx, 3, ignore_btn)

        # 计数标签 + 应用按钮启用状态
        kind_text = "错别字" if self._current_check_kind == "typo" else "重复字词"
        self.typo_count_label.setText(f"  当前显示：{kind_text}  ·  共 {len(results)} 处")
        self.typo_apply_btn.setEnabled(len(results) > 0)

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

    # ═══════════════════════════════════════════════════
    # 权利要求书检查 Tab — 槽函数
    # ═══════════════════════════════════════════════════

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
            from config_manager import load_vague_wordbank
            vague_words = load_vague_wordbank()
            n = int(self._claim_n)
            results = run_all_checks(
                shell_paragraphs,
                self._claim_start_idx,
                self._claim_end_idx,
                n=n,
                ignore_set=set(self._claim_session_ignore),
                vague_words=vague_words,
                check_term=self.claim_term_cb.isChecked(),
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

            # 操作列：忽略（按本条 context 里的关键词）
            btn = QPushButton("忽略")
            btn.setObjectName("smallBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _=False, r=row_idx: self._on_claim_ignore_row(r))
            self.claim_result_table.setCellWidget(row_idx, 4, btn)

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
