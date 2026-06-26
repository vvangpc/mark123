# -*- coding: utf-8 -*-
"""
main_window.py — PyQt6 主窗口
专利附图标记桌面软件的GUI界面。
"""
import os
import traceback
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton,
    QLabel, QTextEdit, QPlainTextEdit, QFileDialog, QMessageBox,
    QStatusBar, QProgressBar, QFrame, QSplitter, QTabWidget, QStackedWidget,
    QGroupBox, QApplication, QCheckBox, QDialog, QSpinBox, QAbstractSpinBox,
    QButtonGroup, QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QMenu, QSizePolicy,
    QStyle, QStyleOptionButton, QStylePainter,
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread, QPoint, QSize
from PyQt6.QtGui import (
    QDragEnterEvent, QDropEvent, QTextCursor, QTextCharFormat, QColor, QFont,
)

from core.doc_parser import parse_document
from core.mark_extractor import extract_marks_from_paragraph, extract_marks_from_paragraphs, marks_to_display_text, parse_marks_from_display_text
from core.annotator import smart_annotate_section, smart_remove_section
from ui.styles import DARK_THEME_QSS, LIGHT_THEME_QSS
from ui.content_area import ContentArea
from ui.nav_panel import NavPanel
from core.cleaner import (
    remove_suoshu, unify_halfwidth_punct, convert_fullwidth_to_halfwidth,
    detect_orphan_marks,
    check_typos_wordbank, check_duplicate_words,
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


class MarqueeButton(QPushButton):
    """文件上传按钮：长文件名跑马灯滚动。

    滚动只通过 paintEvent 重绘实现（`set_marquee_text` → `update()`），
    **不** 调用 setText/updateGeometry，因此每帧不触发任何布局重排，
    避免与 1框/2框/3列 分隔条拖拽、文档导入相互卡顿。
    QSS 背景 / 颜色 / hover / 左对齐 全部由样式表绘制（drawControl）保留。
    """

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self._marquee_text = None   # None=普通绘制；非空=画此覆盖文本

    # 宽度不随文件名增长：仅填充所在列的可用宽度，绝不因超长文件名撑大按钮 /
    # 撑大右侧列 / 让窗口无法缩小（高度仍按样式）。超长部分由跑马灯滚动展示。
    def sizeHint(self):
        return QSize(0, super().sizeHint().height())

    def minimumSizeHint(self):
        return QSize(0, super().minimumSizeHint().height())

    def set_marquee_text(self, s):
        if s == self._marquee_text:
            return
        self._marquee_text = s
        self.update()               # 仅重绘，不重排

    def paintEvent(self, ev):
        if self._marquee_text is None:
            super().paintEvent(ev)
            return
        opt = QStyleOptionButton()
        self.initStyleOption(opt)
        opt.text = self._marquee_text
        QStylePainter(self).drawControl(QStyle.ControlElement.CE_PushButton, opt)


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
                    from core.cleaner import fix_consecutive_punct
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
                from core.cleaner import detect_orphan_figures
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
                self.progress.emit(80)
                merged = merge_typo_results(wb_results)
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
    """悬浮Toast提示。定位由 MainWindow 统一管理（堆叠 + resize 跟随），
    自身只负责样式、尺寸与自动消失。"""

    def __init__(self, parent, message: str, toast_type: str = "info",
                 on_closed=None):
        super().__init__(message, parent)
        self._on_closed = on_closed
        self.setWordWrap(True)                 # 长文本/多行自动换行，避免被裁切遮挡
        self.setMinimumWidth(280)
        self.setMaximumWidth(460)              # 过宽则换行，不撑出屏幕
        self.setMinimumHeight(40)              # 单行时维持原观感
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

        # 宽度按最长一行决定（夹在 280~460），再按该宽度自适应高度，
        # 避免多行/长文本被固定高度裁切（用户反馈的「遮挡」问题）
        fm = self.fontMetrics()
        longest = max((fm.horizontalAdvance(ln) for ln in str(message).split("\n")), default=0)
        self.setFixedWidth(min(max(longest + 56, 280), 460))
        self.adjustSize()

        # 3秒后自动消失
        QTimer.singleShot(3000, self._fade_out)

    def _fade_out(self):
        if self._on_closed is not None:
            try:
                self._on_closed(self)
            except Exception:
                pass
        self.close()
        self.deleteLater()


class _ClickableLabel(QLabel):
    """整体可点击的标签：点击任意位置发出 clicked 信号。
    不要用 `label.mousePressEvent = lambda ...` 覆盖实例方法——那会绕过
    QLabel 自身的事件处理（linkActivated 等信号会因此永不触发）。"""
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


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
        self._active_toasts = []     # 当前显示中的 toast（用于堆叠与 resize 重排）
        # 权利要求书检查 Tab 的状态
        self._claim_start_idx = None       # 权利要求书段落起始索引
        self._claim_end_idx = None         # 权利要求书段落结束索引（不含）
        self._claim_para_count = 0         # 权利要求书段落总数（用于确认时校验）
        self._claim_dirty = False          # 预览框是否有未确认修改
        self._claim_results = []           # 当前检查结果缓存
        self._claim_loaded = False         # 当前文档是否已加载过权利要求书内容
        self._claim_n = 2                  # 当前滑窗字数（按钮 / 自定义共享）
        self._claim_session_ignore = set() # 本次会话内结果行「忽略」记录（不持久化）
        # 说明书检查（实施例编号 / 摘要字数）状态
        self._spec_impl_ok = False         # 当前文档是否有「具体实施方式」章节
        self._spec_abs_ok = False          # 当前文档是否有「说明书摘要」章节
        self._spec_results = []            # 当前说明书检查结果缓存
        self._spec_kind = None             # 当前结果属于哪个检查（embodiment / abstract）

        # 配置管理器
        from config.config_manager import AppSettings
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
            from ui.styles import DARK_THEME_QSS, LIGHT_THEME_QSS
            app = QApplication.instance()
            if app is not None:
                app.setStyleSheet(DARK_THEME_QSS if self.current_theme == "dark" else LIGHT_THEME_QSS)
        except Exception:
            pass

        # 应用持久化的复选框勾选状态
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
            self.open_dir_cb.setChecked(
                self.settings.get_bool("gen/open_dir", False)
            )
            self.claim_dyn_trunc_cb.setChecked(
                self.settings.get_bool("claim/dyn_truncate", False)
            )
            self.claim_dyn_fb_cb.setChecked(
                self.settings.get_bool("claim/dyn_fallback", False)
            )
            self.claim_vague_cb.setChecked(
                self.settings.get_bool("claim/check_vague", True)
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

        # ===== 主体：三区布局（无顶部标题栏，最大化内容 / 操作面积）=====
        #   左上：常驻内容区（4 标签页）   左下：当前模块面板（QStackedWidget）
        #   右侧：模块切换竖条 + 文件操作
        # 主体水平分割：左区（1框/2框） ↔ 右区（3列/4列导航 + 文件操作）可拖拽调宽
        body = QSplitter(Qt.Orientation.Horizontal)
        body.setObjectName("bodySplit")
        body.setChildrenCollapsible(False)
        body.setHandleWidth(6)

        left_split = QSplitter(Qt.Orientation.Vertical)
        left_split.setObjectName("mainSplit")

        # 左上：专利内容常驻区（任务6：结构化实时编辑）
        self.content_area = ContentArea()
        self.content_area.contentEdited.connect(self._on_content_edited)
        self.content_area.editWarning.connect(
            lambda msg: self._show_toast(msg, "warning")
        )
        self.content_area.editConfirmed.connect(self._on_content_confirmed)
        left_split.addWidget(self.content_area)

        # 左下：各「子功能」面板（每页一个子功能，由右侧两列导航切换；
        #        索引须与 NavPanel 的 page_index 对应）
        self.panel_stack = QStackedWidget()
        self.panel_stack.addWidget(self._create_mark_tab())                      # 0 标记：附图标记标注
        self.panel_stack.addWidget(self._wrap_card(self._build_suoshu_card()))   # 1 清洗：删除“所述”
        self.panel_stack.addWidget(self._wrap_card(self._build_punct_card()))    # 2 清洗：标点检查
        self.panel_stack.addWidget(self._wrap_card(self._build_orphan_card()))   # 3 清洗：孤立标记检测
        self.panel_stack.addWidget(self._create_typo_tab())                      # 4 错别字 / 重复字
        self.panel_stack.addWidget(self._create_claim_check_tab())               # 5 权利要求书检查
        self.panel_stack.addWidget(self._create_replace_page())                  # 6 清洗：全文替换（输入框）
        self.panel_stack.addWidget(self._create_spec_tab())                      # 7 说明书检查（实施例 / 摘要，共享结果表）
        left_split.addWidget(self.panel_stack)
        # 2框允许被分隔条自由收缩 / 收起：QStackedWidget 默认最小高度取最高页（如结果表），
        # 显式置 0，避免显示矮小卡片时 2框 仍被撑高、分隔条拖不下去（红框区收不掉）。
        self.panel_stack.setMinimumHeight(0)
        # 内容区为主舞台：默认占据约 2/3 高度，确保专利内容清晰可读
        left_split.setStretchFactor(0, 2)
        left_split.setStretchFactor(1, 1)
        left_split.setSizes([560, 280])
        left_split.setHandleWidth(6)
        # 1框（内容区）不可收起；2框（操作面板）可收起——向下拖到一定位置即收起消失（无极调节）
        left_split.setCollapsible(0, False)
        left_split.setCollapsible(1, True)
        body.addWidget(left_split)

        # 右侧：两列导航（模块 → 子功能） + 文件操作
        right_widget = QWidget()
        right_col = QVBoxLayout(right_widget)
        right_col.setContentsMargins(0, 0, 0, 0)
        right_col.setSpacing(8)

        # 「标记」模块的操作按钮组放进右侧第二列（4列）；其余模块第二列仍是子功能列表。
        self.mark_actions_panel = self._create_mark_actions()
        self.nav_panel = NavPanel([
            ("📌 标记", self.mark_actions_panel, 0),
            ("🧹 清洗", self._build_clean_nav(), 1),   # 控件型：4列放 清洗各功能 + 确认替换
            ("📝 错别字", self._build_typo_nav(), 4),  # 控件型：错别字检查 + 错别字词库
            ("🔁 重复字", self._build_dup_nav(), 4),   # 控件型：重复字检查 + 忽略词库（共用 2框 结果表）
            ("⚖️ 权项", self._build_claim_nav(), 5),  # 控件型：检查参数 + 开始检查在 4列
            ("📑 说明书", self._build_spec_nav(), 7),  # 控件型：实施例编号 / 摘要字数（共用 2框 结果表）
        ])
        self.nav_panel.page_selected.connect(self.panel_stack.setCurrentIndex)
        right_col.addWidget(self.nav_panel, 1)

        # 文件上传（兼当前文件名提示），置于「文件生成」上方，与其等宽对齐
        self.file_btn = MarqueeButton("📂 选择 / 拖入 docx")
        self.file_btn.setObjectName("fileBtn")
        self.file_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.file_btn.setToolTip("点击选择，或把 .docx 拖到窗口任意位置")
        self.file_btn.clicked.connect(self._on_select_file)
        right_col.addWidget(self.file_btn)
        # 长文件名跑马灯滚动显示（字符窗口滑动，不触发省略号、保留 QSS）
        self._file_marquee_full = ""
        self._file_marquee_pos = 0
        self._file_marquee_timer = QTimer(self)
        self._file_marquee_timer.setInterval(120)
        self._file_marquee_timer.timeout.connect(self._file_marquee_tick)

        # 文件生成
        self.generate_btn = QPushButton("💾 文件生成")
        self.generate_btn.setObjectName("primaryBtn")
        self.generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setToolTip("将内存中累计的所有修改保存为新的 docx 文件")
        self.generate_btn.clicked.connect(self._on_generate_file)
        right_col.addWidget(self.generate_btn)

        # 打开目录 + 设置（同一行）
        gen_row = QHBoxLayout()
        gen_row.setSpacing(6)
        self.open_dir_cb = QCheckBox("打开目录")
        self.open_dir_cb.setChecked(False)
        self.open_dir_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_dir_cb.setToolTip("勾选后，文件生成成功时自动打开输出目录")
        gen_row.addWidget(self.open_dir_cb)
        gen_row.addStretch()
        self.settings_btn = QPushButton("⚙ 设置")
        self.settings_btn.setObjectName("smallBtn")
        self.settings_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.settings_btn.clicked.connect(self._show_settings_menu)
        gen_row.addWidget(self.settings_btn)
        right_col.addLayout(gen_row)

        body.addWidget(right_widget)
        body.setStretchFactor(0, 1)   # 左区随窗口伸缩
        body.setStretchFactor(1, 0)   # 右区（导航 + 文件）保持窄
        body.setSizes([900, 250])

        main_layout.addWidget(body, 1)

        # ===== 底部：状态栏 =====
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪 — 请选择或拖入 .docx 文件")

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

    def _create_mark_tab(self) -> QWidget:
        """左下（2框）「标记」面板：只保留「附图标记字典」。
        所有标记操作按钮已移至右侧第二列（4列），见 _create_mark_actions()。"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(8)

        marks_group = QGroupBox("📌 附图标记字典（自动提取，可手动编辑）")
        marks_layout = QVBoxLayout(marks_group)

        self.marks_edit = QPlainTextEdit()
        self.marks_edit.setPlaceholderText(
            "打开 docx 文件后将自动提取附图标记...\n"
            "格式示例: 1-齿圈，2-夹指，3-转盘，4-定位销"
        )
        marks_layout.addWidget(self.marks_edit, 1)

        self.mark_count_label = QLabel("")
        self.mark_count_label.setObjectName("subtitleLabel")
        marks_layout.addWidget(self.mark_count_label)

        layout.addWidget(marks_group, 1)
        return widget

    def _create_mark_actions(self) -> QWidget:
        """右侧第二列（4列）的「标记」操作按钮组：竖排、风格统一、分三组。
        组① 附图标记字典：重新确认标记 / 重新提取标记；
        组② 批量标注：一键标注（主操作，实心强调）/ 仅标注权利要求书 / 仅标注具体实施方式；
        组③ 清除标记：删除所有标记 / 清空标记（危险样式）。
        （附图标记字典编辑框留在左下 2框，见 _create_mark_tab()。）"""

        def _caption(text: str) -> QLabel:
            lab = QLabel(text)
            lab.setObjectName("markCaption")
            return lab

        def _btn(text: str, kind: str = "") -> QPushButton:
            b = QPushButton(text)
            b.setObjectName("navActionBtn")
            if kind:
                b.setProperty("kind", kind)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            return b

        widget = QWidget()
        widget.setObjectName("markActions")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        # —— 组① 附图标记字典 ——
        layout.addWidget(_caption("附图标记字典"))

        self.confirm_marks_btn = _btn("✅ 重新确认标记")
        self.confirm_marks_btn.setEnabled(False)
        self.confirm_marks_btn.setToolTip("将左下字典编辑框的内容写回内存中的附图标记段落，并记入操作历史")
        self.confirm_marks_btn.clicked.connect(self._on_confirm_marks)
        layout.addWidget(self.confirm_marks_btn)

        self.refresh_marks_btn = _btn("🔄 重新提取标记")
        self.refresh_marks_btn.setEnabled(False)
        self.refresh_marks_btn.setToolTip("从原始 docx 文档重新提取附图标记到编辑框")
        self.refresh_marks_btn.clicked.connect(self._on_refresh_marks)
        layout.addWidget(self.refresh_marks_btn)

        # —— 组② 批量标注 ——
        layout.addWidget(_caption("批量标注"))

        self.annotate_btn = _btn("⚡ 一键标注", kind="primary")
        self.annotate_btn.setEnabled(False)
        self.annotate_btn.setToolTip("权利要求书 + 具体实施方式 全部自动标注（仅修改内存）")
        self.annotate_btn.clicked.connect(self._on_annotate)
        layout.addWidget(self.annotate_btn)

        self.annotate_claims_btn = _btn("📋 仅标注权利要求书")
        self.annotate_claims_btn.setEnabled(False)
        self.annotate_claims_btn.clicked.connect(lambda: self._on_annotate_section("claims"))
        layout.addWidget(self.annotate_claims_btn)

        self.annotate_impl_btn = _btn("📝 仅标注具体实施方式")
        self.annotate_impl_btn.setEnabled(False)
        self.annotate_impl_btn.clicked.connect(lambda: self._on_annotate_section("implementation"))
        layout.addWidget(self.annotate_impl_btn)

        # —— 组③ 清除标记 ——
        layout.addWidget(_caption("清除标记"))

        self.remove_marks_btn = _btn("🧹 删除所有标记", kind="danger")
        self.remove_marks_btn.setEnabled(False)
        self.remove_marks_btn.setToolTip("基于标记字典，扫描并清洗正文中的编号（仅修改内存）")
        self.remove_marks_btn.clicked.connect(self._on_remove_marks)
        layout.addWidget(self.remove_marks_btn)

        self.clear_marks_btn = _btn("🗑️ 清空标记", kind="danger")
        self.clear_marks_btn.setToolTip("清空左下「附图标记字典」编辑框")
        self.clear_marks_btn.clicked.connect(lambda: self.marks_edit.clear())
        layout.addWidget(self.clear_marks_btn)

        layout.addStretch(1)
        return widget

    def _wrap_card(self, card) -> QWidget:
        """把一张清洗卡片包成 QStackedWidget 的一页（顶部对齐、留白）。"""
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 8, 0, 0)
        lay.addWidget(card)
        lay.addStretch(1)
        return page

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

    def _build_clean_nav(self) -> QWidget:
        """清洗模块 4列：删除所述 / 标点 / 孤立 / 全文替换（导航）+ 确认替换（执行）。"""
        w = QWidget()
        w.setObjectName("markActions")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(5)

        v.addWidget(self._nav_caption("清洗功能"))
        b1 = self._nav_btn("🗑️ 删除“所述”")
        b1.clicked.connect(lambda: self.panel_stack.setCurrentIndex(1))
        v.addWidget(b1)
        b2 = self._nav_btn("🔣 标点检查")
        b2.clicked.connect(lambda: self.panel_stack.setCurrentIndex(2))
        v.addWidget(b2)
        b3 = self._nav_btn("🔍 孤立标记检测")
        b3.clicked.connect(lambda: self.panel_stack.setCurrentIndex(3))
        v.addWidget(b3)

        v.addWidget(self._nav_caption("全文替换"))
        b4 = self._nav_btn("🔁 全文替换")
        b4.setToolTip("在 2框 填写「替换前 / 替换后」")
        b4.clicked.connect(lambda: self.panel_stack.setCurrentIndex(6))
        v.addWidget(b4)
        self.replace_confirm_btn = self._nav_btn("✅ 确认替换", kind="primary")
        self.replace_confirm_btn.setToolTip("把 2框「替换前」的文本全文替换为「替换后」")
        self.replace_confirm_btn.clicked.connect(self._on_clean_replace)
        v.addWidget(self.replace_confirm_btn)

        v.addStretch(1)
        return w

    def _create_replace_page(self) -> QWidget:
        """2框「全文替换」页：替换前 / 替换后 输入框（执行按钮「确认替换」在 4列）。"""
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(10)

        group = QGroupBox("🔁 全文替换")
        g = QVBoxLayout(group)
        hint = QLabel(
            "对全文做文本替换（格式安全，保留公式 / 图片）。填好「替换前 / 替换后」，"
            "再点右侧 4列的「✅ 确认替换」。例：把全文「发明」替换为「实用新型」。"
        )
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        g.addWidget(hint)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("替换前："))
        self.replace_from_edit = QLineEdit()
        self.replace_from_edit.setPlaceholderText("要被替换的文本，如：发明")
        row1.addWidget(self.replace_from_edit, 1)
        g.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("替换后："))
        self.replace_to_edit = QLineEdit()
        self.replace_to_edit.setPlaceholderText("替换成的文本，如：实用新型（留空＝删除）")
        row2.addWidget(self.replace_to_edit, 1)
        g.addLayout(row2)

        layout.addWidget(group)
        layout.addStretch(1)
        return w

    def _on_clean_replace(self):
        """执行全文替换：把「替换前」文本在所有段落中替换为「替换后」（格式安全）。"""
        if not self.doc_data:
            self._show_toast("请先打开文档！", "error")
            return
        if self._is_busy():
            self._show_toast("正在处理中，请稍候", "warning")
            return
        src = self.replace_from_edit.text()
        dst = self.replace_to_edit.text()
        if not src:
            self._show_toast("请先在 2框 填写「替换前」内容", "warning")
            return
        from core.annotator import annotate_paragraph_safe
        n = 0
        for p in self.doc_data["paragraphs"]:
            try:
                if annotate_paragraph_safe(p, {src: dst}):
                    n += 1
            except Exception:
                pass
        self.content_area.load(self.doc_data)   # 刷新 1框 显示
        self._invalidate_typo_cache()
        self._invalidate_dup_cache()
        self._log_clean(f"🔁 全文替换：「{src}」→「{dst}」，{n} 段发生替换")
        if n:
            self._add_history(f"全文替换「{src}」→「{dst}」", f"共 {n} 段")
            self._show_toast(f"已替换 {n} 段", "success")
        else:
            self._show_toast("未找到可替换内容", "info")

    @staticmethod
    def _nav_btn(text: str, kind: str = "") -> QPushButton:
        """4列操作按钮：与「标记」模块同款 navActionBtn 风格（左对齐、整宽，
        kind='primary' 实心强调 / 'danger' 危险）。"""
        b = QPushButton(text)
        b.setObjectName("navActionBtn")
        if kind:
            b.setProperty("kind", kind)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        return b

    @staticmethod
    def _nav_caption(text: str) -> QLabel:
        """4列分组小标题（与「标记」模块 markCaption 同款）。"""
        lab = QLabel(text)
        lab.setObjectName("markCaption")
        return lab

    def _set_apply_enabled(self, flag: bool):
        """同时启停 错别字 / 重复字 两个模块 4列 的「应用所有修改」按钮。"""
        for b in (getattr(self, "typo_apply_btn", None), getattr(self, "dup_apply_btn", None)):
            if b is not None:
                b.setEnabled(flag)

    def _build_typo_nav(self) -> QWidget:
        """错别字模块 4列：检查 + 应用所有修改 + 词库（navActionBtn 风格）。"""
        w = QWidget()
        w.setObjectName("markActions")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(5)

        v.addWidget(self._nav_caption("检查"))
        self.typo_check_btn = self._nav_btn("🔍 错别字检查", kind="primary")
        self.typo_check_btn.setEnabled(False)
        self.typo_check_btn.clicked.connect(self._on_typo_check)
        v.addWidget(self.typo_check_btn)

        self.typo_apply_btn = self._nav_btn("✅ 应用所有修改")
        self.typo_apply_btn.setEnabled(False)
        self.typo_apply_btn.setToolTip("把 2框「建议修改」列的内容写回内存")
        self.typo_apply_btn.clicked.connect(self._on_apply_corrections)
        v.addWidget(self.typo_apply_btn)

        v.addWidget(self._nav_caption("词库"))
        self.wb_btn = self._nav_btn("")
        self.wb_btn.setToolTip("打开错别字词库编辑器（可增删 / 导入 / 导出）")
        self.wb_btn.clicked.connect(self._on_open_wordbank_dialog)
        v.addWidget(self.wb_btn)

        v.addStretch(1)
        self._refresh_wordbank_label()
        return w

    def _build_dup_nav(self) -> QWidget:
        """重复字模块 4列：检查 + 忽略词库（参照标记模块的 navActionBtn 分组风格）。"""
        w = QWidget()
        w.setObjectName("markActions")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(5)

        v.addWidget(self._nav_caption("检查"))
        self.dup_check_btn = self._nav_btn("🔁 重复字词检查", kind="primary")
        self.dup_check_btn.setEnabled(False)
        self.dup_check_btn.clicked.connect(self._on_dup_check)
        v.addWidget(self.dup_check_btn)

        self.dup_apply_btn = self._nav_btn("✅ 应用所有修改")
        self.dup_apply_btn.setEnabled(False)
        self.dup_apply_btn.setToolTip("把 2框「建议修改」列的内容写回内存")
        self.dup_apply_btn.clicked.connect(self._on_apply_corrections)
        v.addWidget(self.dup_apply_btn)

        v.addWidget(self._nav_caption("词库"))
        self.dup_ignore_btn = self._nav_btn("")
        self.dup_ignore_btn.setToolTip("打开「重复字词忽略词库」编辑器")
        self.dup_ignore_btn.clicked.connect(self._on_open_dup_ignore_dialog)
        v.addWidget(self.dup_ignore_btn)

        v.addStretch(1)
        self._refresh_dup_ignore_label()
        return w

    def _create_typo_tab(self) -> QWidget:
        """创建错别字检查标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(10)

        # 检查 / 词库 / 应用 按钮都在右侧 4列；2框 标题随当前检查类型动态变化，只放结果表
        self.typo_result_group = QGroupBox("📝 错别字 / 重复字词检查结果")
        action_group = self.typo_result_group
        action_v = QVBoxLayout(action_group)

        # 单一结果表格，两类检查共用（计数并入标题，不再单列计数行）
        # 去掉「原文片段」列——原文已在 1框 常驻显示，结果直接在 1框 内联高亮；
        # 单击「修改前」会跳转到 1框 对应位置并标红。
        self.typo_table = QTableWidget(0, 5)
        self.typo_table.setHorizontalHeaderLabels(
            ["章节", "修改前", "修改后", "修改", "忽略"]
        )
        _th = self.typo_table.horizontalHeader()
        _th.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # 章节：自适应
        # 修改前 / 修改后 均为 Interactive，可独立拖拽列边框；不设 Stretch（同前理由）。
        _th.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)       # 修改前
        _th.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)       # 修改后
        _th.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)            # 修改：固定
        _th.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)            # 忽略：固定，吸附右侧
        _th.setStretchLastSection(False)
        self.typo_table.setColumnWidth(1, 200)
        self.typo_table.setColumnWidth(2, 240)
        self.typo_table.setColumnWidth(3, 64)
        self.typo_table.setColumnWidth(4, 64)
        self.typo_table.setAlternatingRowColors(True)
        self.typo_table.verticalHeader().setVisible(False)
        # 行高给文字留足空间
        self.typo_table.verticalHeader().setDefaultSectionSize(30)
        self.typo_table.verticalHeader().setMinimumSectionSize(28)
        # 「修改前」单击跳转 1框；「修改」「忽略」列单击触发应用单条 / 忽略
        self.typo_table.cellClicked.connect(self._on_typo_cell_clicked)
        action_v.addWidget(self.typo_table, 1)

        layout.addWidget(action_group, 1)

        # 当前显示的检查类型："typo" 或 "dup" 或 None
        self._current_check_kind = None
        return widget

    # ─────────────────────────────────────────
    # 权利要求书检查 Tab
    # ─────────────────────────────────────────
    def _build_claim_nav(self) -> QWidget:
        """权项模块的 4列控件：检查字数 / 动态截断 / 动态回退 / 不确定用语 / 开始检查。"""
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)

        # 检查字数：6 个预设 + 自定义
        self._claim_n = 2
        v.addWidget(QLabel("检查字数"))
        n_grid = QGridLayout()
        n_grid.setSpacing(3)
        self._claim_n_buttons = QButtonGroup(w)
        self._claim_n_buttons.setExclusive(True)
        for i, val in enumerate((2, 3, 4, 5, 6, 7)):
            b = QPushButton(str(val))
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setObjectName("nPresetBtn")
            if val == 2:
                b.setChecked(True)
            b.clicked.connect(lambda _=False, vv=val: self._on_claim_n_preset(vv))
            self._claim_n_buttons.addButton(b, val)
            n_grid.addWidget(b, i // 2, i % 2)   # 两列：每行 2 个
        v.addLayout(n_grid)

        custom_row = QHBoxLayout()
        custom_row.setSpacing(4)
        custom_row.addWidget(QLabel("自定义"))
        self.claim_n_custom = QSpinBox()
        self.claim_n_custom.setObjectName("nCustomSpin")
        self.claim_n_custom.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.claim_n_custom.setRange(2, 30)
        self.claim_n_custom.setValue(8)
        self.claim_n_custom.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.claim_n_custom.setToolTip("自定义检查字数；填入后以此为准（取消上方预设选中）")
        self.claim_n_custom.valueChanged.connect(self._on_claim_n_custom_changed)
        custom_row.addWidget(self.claim_n_custom, 1)
        v.addLayout(custom_row)

        # 引用基础降噪：动态截断（带黑名单） / 动态回退
        trunc_row = QHBoxLayout()
        trunc_row.setSpacing(4)
        self.claim_dyn_trunc_cb = QCheckBox()
        self.claim_dyn_trunc_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        trunc_row.addWidget(self.claim_dyn_trunc_cb)
        self.claim_dyn_trunc_label = QLabel(
            '<a href="info" style="text-decoration:none;color:inherit;">动态截断</a>'
            '&nbsp;<a href="bl" style="text-decoration:none;color:#3a8ee6;">[黑名单]</a>'
        )
        self.claim_dyn_trunc_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.claim_dyn_trunc_label.setToolTip("点「动态截断」查看说明；点「[黑名单]」编辑边界词库")
        self.claim_dyn_trunc_label.linkActivated.connect(self._on_dyn_trunc_link)
        trunc_row.addWidget(self.claim_dyn_trunc_label)
        trunc_row.addStretch()
        v.addLayout(trunc_row)

        fb_row = QHBoxLayout()
        fb_row.setSpacing(4)
        self.claim_dyn_fb_cb = QCheckBox()
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
        v.addLayout(fb_row)

        # 不确定用语检查（默认勾选）
        vague_row = QHBoxLayout()
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
        self.claim_vague_label.setToolTip("点「不确定用语检查」查看说明；点「[词库]」编辑词库")
        self.claim_vague_label.linkActivated.connect(self._on_vague_link)
        vague_row.addWidget(self.claim_vague_label)
        vague_row.addStretch()
        v.addLayout(vague_row)

        self.claim_check_btn = self._nav_btn("▶ 开始检查")
        self.claim_check_btn.setEnabled(False)
        self.claim_check_btn.clicked.connect(self._on_claim_check_start)
        v.addWidget(self.claim_check_btn)

        v.addStretch()
        return w

    def _create_claim_check_tab(self) -> QWidget:
        """权利要求书检查结果页（2框）。检查参数 / 开始检查在 4列（见 _build_claim_nav）；
        权利要求书正文在 1框「权利要求书」标签页查看 / 编辑。"""
        widget = QWidget()
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(0, 8, 0, 0)
        outer.setSpacing(8)

        self.claim_status_label = QLabel("请先打开 docx 文件")
        self.claim_status_label.setObjectName("subtitleLabel")
        outer.addWidget(self.claim_status_label)

        hint = QLabel(
            "表中仅展示问题，不写入最终文件；可定位的问题已在 1框「权利要求书」标黄，"
            "单击「说明」跳到 1框 并把该条标红；在 1框改完正文后再次「开始检查」重扫。"
        )
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        outer.addWidget(hint)

        # 去「上下文」列——原文已在 1框 常驻并内联标黄；单击「说明」跳转 1框 标红。
        self.claim_result_table = QTableWidget(0, 4)
        self.claim_result_table.setHorizontalHeaderLabels(
            ["类型", "权项", "说明", "操作"]
        )
        h = self.claim_result_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)         # 说明
        h.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)           # 操作
        self.claim_result_table.setColumnWidth(3, 72)
        h.setStretchLastSection(False)
        self.claim_result_table.setAlternatingRowColors(True)
        self.claim_result_table.verticalHeader().setVisible(False)
        self.claim_result_table.verticalHeader().setDefaultSectionSize(30)
        self.claim_result_table.verticalHeader().setMinimumSectionSize(28)
        # 单击「说明」跳转 1框；单击「操作」忽略（不再有双击行为）
        self.claim_result_table.cellClicked.connect(self._on_claim_cell_clicked)
        outer.addWidget(self.claim_result_table, 1)
        return widget

    def _get_wordbank_count(self) -> int:
        """读取当前生效词库条目数（合并内置 + 用户自定义）"""
        try:
            from config.config_manager import get_merged_wordbank
            return len(get_merged_wordbank())
        except Exception:
            try:
                from config.typo_wordbank import WORDBANK
                return len(WORDBANK)
            except Exception:
                return 0

    def _refresh_wordbank_label(self):
        """刷新 4列「错别字词库」按钮文字（含条目数）"""
        if not hasattr(self, "wb_btn"):
            return
        self.wb_btn.setText(f"📔 错别字词库 ({self._get_wordbank_count()})")

    def _on_open_wordbank_dialog(self):
        """打开词库编辑对话框"""
        try:
            from ui.dialogs.wordbank_dialog import WordbankDialog
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
        """刷新 4列「重复字忽略词库」按钮文字（含条目数）"""
        if not hasattr(self, "dup_ignore_btn"):
            return
        try:
            from config.config_manager import load_dup_ignore_list
            n = len(load_dup_ignore_list())
        except Exception:
            n = 0
        self.dup_ignore_btn.setText(f"🙈 重复字忽略词库 ({n})")

    def _on_open_dup_ignore_dialog(self):
        """打开「重复字词忽略词库」编辑对话框"""
        try:
            from ui.dialogs.dup_ignore_dialog import DupIgnoreDialog
        except Exception as e:
            QMessageBox.critical(self, "无法打开", f"加载忽略词库编辑器失败：\n{e}")
            return
        dlg = DupIgnoreDialog(self)
        dlg.exec()
        self._refresh_dup_ignore_label()
        # 失效重复字词检查的缓存 —— 下次点「重复字词检查」时强制重新扫描，
        # 确保新添加的忽略词立即生效
        self._invalidate_dup_cache()

    def _on_content_edited(self):
        """1框 专利内容被结构化编辑回写到内存后：失效错别字/重复字检查缓存，
        下次检查时按编辑后的最新内存重新扫描（权项检查每次都读内存，无需额外处理）。"""
        self._invalidate_typo_cache()
        self._invalidate_dup_cache()

    def _on_content_confirmed(self, count: int):
        """点击 1框「✓ 确认修改」后的反馈提示。"""
        if count > 0:
            self._show_toast(f"已写入内存：{count} 段修改", "success")
        else:
            self._show_toast("没有需要写入的文本改动", "info")

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

    # ===== 事件处理 =====

    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖入文件时的处理"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.docx') for url in urls):
                event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        """释放拖入文件时的处理"""
        if self._is_busy():
            self._show_toast("正在处理中，请稍候再切换文件", "info")
            return
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.docx'):
                self._load_document(file_path)
                break

    # ===== 操作回调 =====

    def _set_file_button_name(self, filename: str):
        """设置 file_btn 上的文件名，并启动跑马灯（溢出时滚动，由 tick 自行判定）。

        这里 setText 一次确定基准文本；之后滚动只走 set_marquee_text(重绘不重排)，
        避免每帧触发布局重排而与分隔条拖拽 / 文档导入相互卡顿。
        """
        self._file_marquee_full = filename or ""
        self._file_marquee_pos = 0
        self.file_btn.setText(f"📄 {self._file_marquee_full}")
        self.file_btn.set_marquee_text(None)
        if self._file_marquee_full:
            self._file_marquee_timer.start()
        else:
            self._file_marquee_timer.stop()

    def _file_marquee_tick(self):
        """文件名跑马灯：前缀「📄 」固定，文件名在剩余宽度内左对齐；
        溢出时按字符窗口滑动——每次只显示一段恰好不超出可用宽度的子串，
        既不触发省略号(…)，也只重绘不重排（QSS 完整保留）。"""
        full = self._file_marquee_full
        if not full:
            self._file_marquee_timer.stop()
            return
        prefix = "📄 "
        fm = self.file_btn.fontMetrics()
        # 28 ≈ 左右内边距(12*2) + 虚线边框等余量
        name_avail = self.file_btn.width() - 28 - fm.horizontalAdvance(prefix)
        if name_avail <= 0 or fm.horizontalAdvance(full) <= name_avail:
            # 不溢出：恢复普通绘制（显示 setText 的完整文本）
            self.file_btn.set_marquee_text(None)
            self._file_marquee_pos = 0
            return
        loop = full + "      "          # 循环间隔
        s = loop + loop
        self._file_marquee_pos = (self._file_marquee_pos + 1) % len(loop)
        start = self._file_marquee_pos
        end = start
        while end < start + len(loop) and fm.horizontalAdvance(s[start:end + 1]) <= name_avail:
            end += 1
        self.file_btn.set_marquee_text(prefix + s[start:end])

    def _on_select_file(self):
        """选择文件"""
        if self._is_busy():
            self._show_toast("正在处理中，请稍候再切换文件", "info")
            return
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

            # 更新文件信息（按钮自身既是上传入口也是文件名提示）
            filename = os.path.basename(file_path)
            self._set_file_button_name(filename)
            self.file_btn.setToolTip(file_path)

            # 提取标记
            self._extract_and_display_marks()

            self.progress_bar.setValue(80)
            QApplication.processEvents()

            # 填充左上常驻内容区
            self.content_area.load(self.doc_data)

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
            # 说明书检查：按章节存在性启停「实施例编号 / 摘要字数」按钮
            self._spec_tab_load_from_doc()

            # 作废上一份文档的错别字/重复字词检查缓存——
            # 旧结果的 para_idx 指向旧文档，残留会导致新文档显示
            # 旧结果、甚至按旧位置应用修正
            self.typo_data = []
            self.dup_data = []
            self._current_check_kind = None
            self.typo_table.setRowCount(0)
            self.typo_result_group.setTitle("📝 错别字 / 重复字词检查结果")
            self._set_apply_enabled(False)

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

    # ─────────────── 单实例：接收远端转交的文件 ───────────────
    def _receive_remote_file(self, file_path: str):
        """另一个进程通过 QLocalSocket 转交过来的文件路径。
        行为对齐拖入 / 点击：直接替换当前文件。"""
        if not file_path or not os.path.isfile(file_path):
            self._raise_to_front()
            return
        if not file_path.lower().endswith(".docx"):
            self._raise_to_front()
            return

        # 后台 worker 正在跑时拒绝重新加载，避免半截改写 doc_data
        if self._is_busy():
            self._raise_to_front()
            self._show_toast("正在处理中，请稍候再切换文件", "info")
            return

        self._raise_to_front()
        self._load_document(file_path)

    def _raise_to_front(self):
        """把窗口从最小化恢复并置顶激活"""
        if self.isMinimized():
            self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.show()
        self.raise_()
        self.activateWindow()

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
            self.marks_edit.setPlainText("")
            self.current_marks = {}
            self.mark_count_label.setText("未找到附图标记段落")
            return
        self.current_marks = marks
        self.marks_edit.setPlainText(marks_to_display_text(marks))
        self.mark_count_label.setText(f"共 {len(marks)} 个标记")

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
        if self._is_busy():
            self._show_toast("正在处理中，请等待当前操作完成", "warning")
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

        # 标注/删除标记已写入内存 → 刷新 1框 显示最新效果，并失效检查缓存
        self.content_area.load(self.doc_data)
        self._invalidate_typo_cache()
        self._invalidate_dup_cache()

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
        if self._is_busy():
            self._show_toast("正在处理中，请等待当前操作完成后再生成文件", "warning")
            return
        # 落盘前先把 1框 里待回写的结构化编辑写入内存
        self.content_area.flush_all()
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
        """重新渲染历史（内容区「操作历史」标签页）"""
        edit = self.content_area.history_edit
        if not self.history_entries:
            edit.clear()
            return
        lines = []
        for i, e in enumerate(self.history_entries, 1):
            lines.append(f"<b>#{i}  [{e['time']}]  {e['summary']}</b>")
            if e.get("detail"):
                for d in e["detail"].split("\n"):
                    if d.strip():
                        lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;{d}")
            lines.append("")
        edit.setHtml("<br>".join(lines))
        sb = edit.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _clear_history(self):
        self.history_entries = []
        self.content_area.history_edit.clear()
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
        from core.annotator import update_mark_paragraph_text
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

    def _is_busy(self) -> bool:
        """是否有后台 worker 正在读写内存中的文档树"""
        w = getattr(self, "worker", None)
        if w is not None and w.isRunning():
            return True
        w = getattr(self, "clean_worker", None)
        if w is not None and w.isRunning():
            return True
        return False

    def _set_doc_ops_enabled(self, enabled: bool):
        """任一后台 worker 运行期间，禁掉所有会读写文档树的入口
        （标注组 / 清洗检查组 / 权要 Tab / 章节预览 / 生成文件），
        防止两个线程并发读写同一棵 lxml 树导致文档损坏或崩溃。
        恢复时按各控件自身的状态条件点亮。"""
        # 标注组
        self.annotate_btn.setEnabled(enabled)
        self.annotate_claims_btn.setEnabled(enabled)
        self.annotate_impl_btn.setEnabled(enabled)
        self.remove_marks_btn.setEnabled(enabled)
        self.confirm_marks_btn.setEnabled(enabled)
        self.refresh_marks_btn.setEnabled(enabled)
        self.file_btn.setEnabled(enabled)
        # 清洗 / 检查组
        self.suoshu_btn.setEnabled(enabled)
        self.punct_btn.setEnabled(enabled)
        self.orphan_btn.setEnabled(enabled)
        self.typo_check_btn.setEnabled(enabled)
        self.dup_check_btn.setEnabled(enabled)
        self._set_apply_enabled(enabled and bool(self._active_cache_list()))
        # 权要 Tab
        self.claim_check_btn.setEnabled(enabled and self._claim_loaded)
        # 说明书检查（按各自章节存在性）
        self.spec_emb_btn.setEnabled(enabled and self._spec_impl_ok)
        self.spec_abs_btn.setEnabled(enabled and self._spec_abs_ok)
        # generate_btn 仅在有历史时启用
        self.generate_btn.setEnabled(enabled and bool(self.history_entries))

    def _set_buttons_enabled(self, enabled: bool):
        """设置标注操作按钮状态（与清洗组互锁，见 _set_doc_ops_enabled）"""
        self._set_doc_ops_enabled(enabled)

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

    # ─────────────── 设置菜单 ───────────────
    def _show_settings_menu(self):
        """在设置按钮下方弹出菜单：切换主题 / 关于 / 检查更新"""
        menu = QMenu(self)
        act_theme = menu.addAction("🌓  切换主题")
        act_about = menu.addAction("ℹ  关于")
        act_update = menu.addAction("🔄  检查更新")

        act_theme.triggered.connect(self._toggle_theme)
        act_about.triggered.connect(self._show_about_dialog)
        act_update.triggered.connect(self._on_check_updates_manual)

        btn = self.settings_btn
        menu.exec(btn.mapToGlobal(QPoint(0, btn.height())))

    def _show_about_dialog(self):
        """关于对话框"""
        try:
            from version import __version__
        except Exception:
            __version__ = "?"
        QMessageBox.about(
            self,
            "关于 专利标记助手",
            f"<b>专利标记助手</b> V{__version__}<br/><br/>"
            "面向中国专利代理人的 .docx 辅助工具：<br/>"
            "附图标记自动提取与标注 · 文本清洗 · 错别字检查 · 权利要求书检查。<br/><br/>"
            "<span style='color:#888'>本地运行，不上传任何文件内容。</span>"
        )

    def _on_check_updates_manual(self):
        """用户主动触发的更新检查（无更新/失败均给反馈）"""
        try:
            from infra.updater import UpdateChecker
            from version import __version__
        except Exception as e:
            QMessageBox.warning(self, "检查更新", f"无法加载更新模块：{e}")
            return
        # 句柄挂在 self 上避免被 GC
        self._manual_update_checker = UpdateChecker(self, __version__, manual=True)
        self._manual_update_checker.start()

    def closeEvent(self, event):
        """窗口关闭时持久化配置；先等运行中的后台线程退出，
        避免 QThread 随窗口销毁时崩溃（Destroyed while thread is still running）"""
        for w in (getattr(self, "worker", None), getattr(self, "clean_worker", None)):
            if w is not None and w.isRunning():
                w.wait(5000)
        try:
            self.settings.set_theme(self.current_theme)
            self.settings.set_geometry(self.saveGeometry())
            if hasattr(self, "punct_halfwidth_cb"):
                self.settings.set_bool("clean/punct_halfwidth", self.punct_halfwidth_cb.isChecked())
            if hasattr(self, "punct_fullwidth_cb"):
                self.settings.set_bool("clean/punct_fullwidth", self.punct_fullwidth_cb.isChecked())
            if hasattr(self, "fix_punctuation_cb"):
                self.settings.set_bool("clean/fix_consecutive_punct", self.fix_punctuation_cb.isChecked())
            if hasattr(self, "open_dir_cb"):
                self.settings.set_bool("gen/open_dir", self.open_dir_cb.isChecked())
            if hasattr(self, "claim_dyn_trunc_cb"):
                self.settings.set_bool("claim/dyn_truncate", self.claim_dyn_trunc_cb.isChecked())
            if hasattr(self, "claim_dyn_fb_cb"):
                self.settings.set_bool("claim/dyn_fallback", self.claim_dyn_fb_cb.isChecked())
            if hasattr(self, "claim_vague_cb"):
                self.settings.set_bool("claim/check_vague", self.claim_vague_cb.isChecked())
            if hasattr(self, "suoshu_checkboxes"):
                for _name, _cb in self.suoshu_checkboxes.items():
                    self.settings.set_bool(f"clean/suoshu/{_name}", _cb.isChecked())
            self.settings.sync()
        except Exception:
            pass
        super().closeEvent(event)

    def _log(self, message: str):
        """添加日志（统一汇总到内容区「操作日志」标签页）"""
        edit = self.content_area.log_edit
        edit.append(message)
        sb = edit.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _show_toast(self, message: str, toast_type: str = "info"):
        """显示Toast提示（多条纵向堆叠，不互相遮挡）"""
        try:
            toast = ToastWidget(self, message, toast_type,
                                on_closed=self._on_toast_closed)
            self._active_toasts.append(toast)
            self._reposition_toasts()
            toast.show()
        except Exception:
            pass  # Toast显示失败不影响主流程

    def _reposition_toasts(self):
        """把所有活跃 toast 右对齐并自上而下堆叠"""
        y = 30
        for t in self._active_toasts:
            x = self.width() - t.width() - 30
            t.move(max(x, 10), y)
            y += t.height() + 8

    def _on_toast_closed(self, toast):
        """toast 消失时从活跃列表移除并重排剩余 toast"""
        try:
            self._active_toasts.remove(toast)
        except ValueError:
            pass
        self._reposition_toasts()

    def resizeEvent(self, event):
        """窗口缩放时让活跃 toast 跟随右上角"""
        super().resizeEvent(event)
        if getattr(self, "_active_toasts", None):
            self._reposition_toasts()

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
            try:
                default = name in SUOSHU_DEFAULT_CHECKED
                cb.setChecked(self.settings.get_bool(f"clean/suoshu/{name}", default))
            except Exception:
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
        """清洗日志统一汇总到内容区「操作日志」标签页"""
        self._log(message)

    def _set_clean_buttons_enabled(self, enabled: bool):
        """设置清洗操作按钮状态（与标注组互锁，见 _set_doc_ops_enabled）"""
        self._set_doc_ops_enabled(enabled)

    def _start_clean_worker(self, action: str, log_prefix: str, history_label: str = None, **kwargs):
        """通用：启动 CleanWorker
        history_label 不为 None 时，操作完成后自动写入操作历史框。
        """
        if not self.doc_data:
            self._show_toast("请先加载文档！", "error")
            return
        if self._is_busy():
            self._show_toast("正在处理中，请等待当前操作完成", "warning")
            return
        # 先把 1框 里待回写的结构化编辑落到内存，确保检查基于最新内容
        self.content_area.flush_all()
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
                from config.config_manager import load_dup_ignore_list
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

    def _on_typo_cell_clicked(self, row: int, col: int):
        """单击「修改前」列(1) → 跳转 1框 高亮；「修改」列(3) → 应用本条；
        「忽略」列(4) → 移除该行。其它列不处理。"""
        if row < 0 or row >= self.typo_table.rowCount():
            return
        if col == 1:
            self._locate_typo_in_content(row)
            return
        if col == 3:
            self._apply_single_correction(row)
            return
        if col != 4:
            return
        # 先回写当前所有编辑，再移除
        self._snapshot_table_to_active_cache()
        data = self._active_cache_list()
        if 0 <= row < len(data):
            data.pop(row)
        self._render_table_from_data(data)

    def _locate_typo_in_content(self, row: int):
        """单击「修改前」：在 1框 跳转到该错别字/重复字位置并标红。"""
        data = self._active_cache_list()
        if not (0 <= row < len(data)):
            return
        item = data[row]
        ok = self.content_area.locate_issue(
            item.get("para_idx", -1),
            item.get("wrong", ""),
            int(item.get("occurrence", 1) or 1),
        )
        if not ok:
            self._show_toast("未能在原文中定位该处（内容可能已变动）", "warning")

    def _apply_single_correction(self, row: int):
        """「修改」：把单条结果的「修改后」写回内存（不影响其它行）。"""
        if not self.doc_data or self._current_check_kind is None:
            return
        # 先把表格里的即时编辑回写到缓存，保证取到用户最新输入的「修改后」
        self._snapshot_table_to_active_cache()
        data = self._active_cache_list()
        if not (0 <= row < len(data)):
            return
        item = data[row]
        confirmed = (item.get("suggestion") or "").strip()
        wrong = item.get("wrong", "")
        if not (wrong and confirmed) or confirmed == wrong:
            self._show_toast("该条没有可应用的修改（「修改后」为空或与原词相同）", "warning")
            return

        from core.cleaner import apply_typo_corrections
        count = apply_typo_corrections(
            self.doc_data["paragraphs"],
            [{"para_idx": item["para_idx"], "wrong": wrong, "confirmed_fix": confirmed}],
        )
        if not count:
            self._show_toast("未找到可替换的文本，可能内容已变动", "warning")
            return

        label_prefix = "错别字修正" if self._current_check_kind == "typo" else "重复字词修正"
        self._add_history(f"{label_prefix}（1 处）", f"{wrong} → {confirmed}")
        # 本条已应用 → 从缓存与表格移除；其余行的 para_idx 不受影响仍有效
        data.pop(row)
        self._render_table_from_data(data)
        # 内存已变动 → 让 1框 与另一类检查缓存保持一致
        self.content_area.load(self.doc_data)
        if self._current_check_kind == "typo":
            self._invalidate_dup_cache()
        else:
            self._invalidate_typo_cache()
        self._show_toast(f"已修改：{wrong} → {confirmed}", "success")

    def _active_cache_list(self) -> list:
        """返回当前模式对应的缓存列表"""
        if self._current_check_kind == "typo":
            return self.typo_data
        if self._current_check_kind == "dup":
            return self.dup_data
        return []

    def _snapshot_table_to_active_cache(self):
        """把表格「修改后」列(col 2)的当前值写回到对应缓存的 suggestion 字段
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

    def _fg(self, light_hex: str, dark_hex: str) -> QColor:
        """按当前主题返回表格前景色（深色主题需要更亮的色调保证对比度）"""
        return QColor(dark_hex if self.current_theme == "dark" else light_hex)

    def _render_table_from_data(self, results: list):
        """把缓存列表渲染到共用表格"""
        # 批量渲染：一次分配全部行 + 暂停重绘，避免大结果集逐行 insertRow 卡顿
        self.typo_table.setUpdatesEnabled(False)
        try:
            self._fill_typo_table(results)
        finally:
            self.typo_table.setUpdatesEnabled(True)

        # 2框 标题随检查类型动态变化（计数并入），并更新应用按钮启用状态
        if self._current_check_kind == "typo":
            self.typo_result_group.setTitle(f"📝 错别字检查结果（共 {len(results)} 处）")
        else:
            self.typo_result_group.setTitle(f"🔁 重复字词检查结果（共 {len(results)} 处）")
        self._set_apply_enabled(len(results) > 0)

        # 1框 内联高亮：把当前结果集全部标黄（空表则清空高亮）
        self.content_area.highlight_issues(results)

    def _fill_typo_table(self, results: list):
        self.typo_table.setRowCount(0)
        self.typo_table.setRowCount(len(results))
        for row_idx, item in enumerate(results):

            # 列0：章节名（para_idx 存 UserRole）
            section_text = item.get("section") or "（未归类）"
            pos_item = QTableWidgetItem(section_text)
            pos_item.setFlags(pos_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            pos_item.setData(Qt.ItemDataRole.UserRole, item["para_idx"])
            self.typo_table.setItem(row_idx, 0, pos_item)

            # 列1：修改前（只读，显示原始错词；单击跳转到 1框 并标红）
            wrong_item = QTableWidgetItem(item.get("wrong", ""))
            wrong_item.setFlags(wrong_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            wrong_item.setForeground(self._fg("#c62828", "#ff7b72"))
            wrong_item.setToolTip("点击跳转到原文并高亮")
            self.typo_table.setItem(row_idx, 1, wrong_item)

            # 列2：修改后（可编辑）
            fix_item = QTableWidgetItem(item.get("suggestion", ""))
            self.typo_table.setItem(row_idx, 2, fix_item)

            _act_font = QFont()
            _act_font.setBold(True)

            # 列3：修改 —— 仅把本条「修改后」写回内存（点击由 cellClicked 捕获）
            fix_btn = QTableWidgetItem("修改")
            fix_btn.setFlags(
                Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
            )
            fix_btn.setTextAlignment(
                Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
            )
            fix_btn.setForeground(self._fg("#1565c0", "#79b8ff"))
            fix_btn.setFont(_act_font)
            fix_btn.setToolTip("把本条「修改后」写入内存（只应用这一条）")
            self.typo_table.setItem(row_idx, 3, fix_btn)

            # 列4：忽略 —— 普通文字单元格（点击由 cellClicked 捕获）
            ig_item = QTableWidgetItem("忽略")
            ig_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
            )
            ig_item.setTextAlignment(
                Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
            )
            ig_item.setForeground(self._fg("#00897b", "#4dd0c4"))
            ig_item.setFont(_act_font)
            ig_item.setToolTip("点击忽略此条")
            self.typo_table.setItem(row_idx, 4, ig_item)

    def _on_clean_finished(self, message: str):
        """清洗操作完成"""
        self._set_clean_buttons_enabled(True)
        self.progress_bar.setVisible(False)
        action = getattr(self, "_pending_clean_action", None)

        # 应用错别字 / 重复字修正后：修正已写入内存，作废两类检查缓存，使下次检查
        # 重新扫描已修正的内容，避免再次报出已修复的问题。
        if action == "typo_apply":
            self._invalidate_typo_cache()
            self._invalidate_dup_cache()

        # 改动文档内容的清洗操作（删"所述"/标点/应用错别字修正）已写入内存
        # → 刷新 1框 显示最新效果。检测类（orphan/typo_check/dup_check）不改文档、
        #   且会自带高亮渲染，故不在此刷新以免清掉高亮。
        if action in ("suoshu", "punct", "typo_apply") and self.doc_data:
            self.content_area.load(self.doc_data)

        # 孤立标记检测：结果只显示在自己卡片的小日志框，不写入全局清洗日志
        if action == "orphan" and hasattr(self, "orphan_result_text"):
            from datetime import datetime
            stamp = datetime.now().strftime("%H:%M:%S")
            self.orphan_result_text.append(f"[{stamp}] {message}\n")
        else:
            self._log_clean(f"✅ {message}")

        self.status_bar.showMessage(message)
        # Toast 已支持多行自适应高度；过长（孤立标记很多时）截断并提示去结果框看详情
        _lines = message.split("\n")
        _toast = message if len(_lines) <= 10 else "\n".join(_lines[:10]) + "\n  …（详见下方结果框）"
        self._show_toast(_toast, "warning" if "⚠️" in message else "success")

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
        """_load_document 成功后：记录权利要求书段落区间（正文在 1框「权利要求书」展示 / 编辑）。"""
        if not self.doc_data:
            return
        # 新文档加载 → 清空本次会话的忽略记录
        self._claim_session_ignore = set()
        sections = self.doc_data.get('sections', {})
        section = sections.get('权利要求书')
        if section is None:
            self._claim_start_idx = None
            self._claim_end_idx = None
            self._claim_para_count = 0
            self._claim_loaded = False
            self.claim_check_btn.setEnabled(False)
            self.claim_status_label.setText("未识别到权利要求书章节")
            self.claim_result_table.setRowCount(0)
            self._claim_results = []
            return

        self._claim_start_idx = section.start_idx
        self._claim_end_idx = section.end_idx
        self._claim_para_count = section.end_idx - section.start_idx
        self._claim_loaded = True
        self.claim_check_btn.setEnabled(True)
        self._claim_results = []
        self.claim_result_table.setRowCount(0)
        self.claim_status_label.setText(
            f"已加载权利要求书：共 {self._claim_para_count} 段  ·  在 1框「权利要求书」查看，点「开始检查」"
        )

    def _update_claim_status_bar(self):
        """刷新权利要求书检查的状态文字"""
        if not self._claim_loaded:
            return
        n_results = len(self._claim_results)
        parts = [f"共 {self._claim_para_count} 段"]
        if n_results:
            parts.append(f"问题 {n_results} 条")
        self.claim_status_label.setText("  ·  ".join(parts))

    def _on_claim_check_start(self):
        """点「开始检查」：基于内存中的权利要求书段落运行检查
        （1框 对正文的编辑实时反映到内存，故无需预览框）。"""
        if not self._claim_loaded or self._claim_start_idx is None:
            self._show_toast("请先加载包含权利要求书的文档", "warning")
            return

        # 先把 1框 权利要求书的结构化编辑落到内存，确保检查基于最新内容
        self.content_area.flush_all()

        try:
            from core.claim_check import run_all_checks
            from config.config_manager import load_vague_wordbank, load_boundary_blacklist
            vague_words = load_vague_wordbank()
            use_trunc = self.claim_dyn_trunc_cb.isChecked()
            use_fb = self.claim_dyn_fb_cb.isChecked()
            boundary_bl = load_boundary_blacklist() if use_trunc else None
            n = int(self._claim_n)
            results = run_all_checks(
                self.doc_data['paragraphs'],
                self._claim_start_idx,
                self._claim_end_idx,
                n=n,
                ignore_set=set(self._claim_session_ignore),
                vague_words=vague_words,
                check_vague=self.claim_vague_cb.isChecked(),
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
            "vague":      "不确定用语",
            "numbering":  "序号",
            "multi_dep":  "多引合法性",
            "ending":     "句号结尾",
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

            # 列2：说明（message；单击跳转 1框 标红；tooltip 显示完整说明 + 建议，
            # 替代被删的「双击弹窗看全文」）
            msg = item.get("message", "")
            msg_item = QTableWidgetItem(msg)
            msg_item.setFlags(msg_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            tip = msg
            sug = (item.get("suggestion") or "").strip()
            if sug:
                tip = f"{msg}\n建议：{sug}"
            msg_item.setToolTip(tip)
            self.claim_result_table.setItem(row_idx, 2, msg_item)

            # 列3：操作（忽略）—— 普通文字单元格（点击由 cellClicked 捕获）
            ig_item = QTableWidgetItem("忽略")
            ig_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
            )
            ig_item.setTextAlignment(
                Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
            )
            ig_item.setForeground(self._fg("#00897b", "#4dd0c4"))
            ig_font = QFont()
            ig_font.setBold(True)
            ig_item.setFont(ig_font)
            ig_item.setToolTip("点击忽略此条")
            self.claim_result_table.setItem(row_idx, 3, ig_item)

        # 1框 内联高亮：把所有能定位的权项问题标黄（空结果→清空高亮）。
        # 权项 para_idx 是「权项首段」，术语可能在后续段 → 用向后搜索版。
        self.content_area.highlight_claim_issues(self._claim_highlight_items(results))

    def _claim_anchor(self, item: dict) -> str:
        """从一条权项结果里抽取要在 1框 高亮的锚点文本。

        antecedent / vague 优先取 message 里『…』内的术语（所述X / W）；
        其余类型回退取 context 的最长非空白片段（引用片段、句末片段等）。
        """
        kind = item.get("kind")
        msg = item.get("message", "")
        if kind in ("antecedent", "vague"):
            import re as _re
            m = _re.search(r'『(.+?)』', msg)
            if m:
                return m.group(1)
        return _longest_nonspace_run((item.get("context") or "").strip())

    def _claim_highlight_items(self, results: list) -> list:
        """把权项结果转成内联高亮所需的 {para_idx, wrong, occurrence} 列表。

        锚点在对应段内找不到的（如序号类问题）由 content_area 自动跳过 →
        即「只标能定位的问题」。
        """
        out = []
        for it in results:
            anchor = self._claim_anchor(it)
            pid = it.get("para_idx", -1)
            if anchor and isinstance(pid, int) and pid >= 0:
                out.append({"para_idx": pid, "wrong": anchor, "occurrence": 1})
        return out

    def _locate_claim_in_content(self, row: int):
        """单击「说明」：在 1框「权利要求书」跳转到该问题位置并标红。"""
        if row < 0 or row >= len(self._claim_results):
            return
        item = self._claim_results[row]
        anchor = self._claim_anchor(item)
        ok = False
        if anchor:
            ok = self.content_area.locate_claim_issue(item.get("para_idx", -1), anchor)
        if not ok:
            self._show_toast("未能在原文中定位该处（可能是序号类问题或内容已变动）", "warning")

    def _on_claim_cell_clicked(self, row: int, col: int):
        """单击「说明」(col 2)→跳转 1框 标红；「操作」(col 3)→忽略该行。"""
        if col == 2:
            self._locate_claim_in_content(row)
        elif col == 3:
            self._on_claim_ignore_row(row)

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

    # ─────────────────────────────────────────
    # 说明书检查（实施例编号 / 摘要字数；共用 2框 结果表 + 1框 内联高亮）
    # ─────────────────────────────────────────
    def _build_spec_nav(self) -> QWidget:
        """说明书模块 4列：实施例编号 / 摘要字数（后续检查在此加按钮即可）。"""
        w = QWidget()
        w.setObjectName("markActions")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(5)

        v.addWidget(self._nav_caption("检查"))
        self.spec_emb_btn = self._nav_btn("📑 实施例编号", kind="primary")
        self.spec_emb_btn.setEnabled(False)
        self.spec_emb_btn.setToolTip("校验「具体实施方式」中 实施例一/二/三… 是否从一连续、无重复、无颠倒")
        self.spec_emb_btn.clicked.connect(lambda: self._on_spec_check("embodiment"))
        v.addWidget(self.spec_emb_btn)

        self.spec_abs_btn = self._nav_btn("📊 摘要字数")
        self.spec_abs_btn.setEnabled(False)
        self.spec_abs_btn.setToolTip("校验「说明书摘要」文字部分是否超过 300 字")
        self.spec_abs_btn.clicked.connect(lambda: self._on_spec_check("abstract"))
        v.addWidget(self.spec_abs_btn)

        v.addStretch(1)
        return w

    def _create_spec_tab(self) -> QWidget:
        """说明书检查结果页（2框）。检查在 4列；结果在此表，组标题随检查类型动态。"""
        widget = QWidget()
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(0, 8, 0, 0)
        outer.setSpacing(8)

        self.spec_status_label = QLabel("请先打开 docx 文件")
        self.spec_status_label.setObjectName("subtitleLabel")
        outer.addWidget(self.spec_status_label)

        hint = QLabel("可定位的问题已在 1框 标黄；单击「说明」跳到对应标签并把该条标红。")
        hint.setObjectName("subtitleLabel")
        hint.setWordWrap(True)
        outer.addWidget(hint)

        self.spec_result_group = QGroupBox("📑 说明书检查结果")
        g = QVBoxLayout(self.spec_result_group)
        self.spec_result_table = QTableWidget(0, 4)
        self.spec_result_table.setHorizontalHeaderLabels(["类型", "定位", "说明", "操作"])
        h = self.spec_result_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)          # 说明
        h.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)            # 操作
        self.spec_result_table.setColumnWidth(3, 72)
        h.setStretchLastSection(False)
        self.spec_result_table.setAlternatingRowColors(True)
        self.spec_result_table.verticalHeader().setVisible(False)
        self.spec_result_table.verticalHeader().setDefaultSectionSize(30)
        self.spec_result_table.verticalHeader().setMinimumSectionSize(28)
        self.spec_result_table.cellClicked.connect(self._on_spec_cell_clicked)
        g.addWidget(self.spec_result_table, 1)
        outer.addWidget(self.spec_result_group, 1)
        return widget

    def _spec_tab_load_from_doc(self):
        """_load_document 成功后：按章节存在性启停说明书检查按钮、清空旧结果。"""
        if not self.doc_data:
            return
        sections = self.doc_data.get("sections", {})
        self._spec_impl_ok = sections.get("具体实施方式") is not None
        self._spec_abs_ok = (sections.get("说明书摘要") or sections.get("摘要")) is not None
        self._spec_results = []
        self._spec_kind = None
        self.spec_result_table.setRowCount(0)
        self.spec_result_group.setTitle("📑 说明书检查结果")
        self.spec_emb_btn.setEnabled(self._spec_impl_ok)
        self.spec_abs_btn.setEnabled(self._spec_abs_ok)
        bits = []
        if self._spec_impl_ok:
            bits.append("具体实施方式")
        if self._spec_abs_ok:
            bits.append("说明书摘要")
        if bits:
            self.spec_status_label.setText("已加载：" + " / ".join(bits) + "  ·  点上方「检查」")
        else:
            self.spec_status_label.setText("未识别到 具体实施方式 / 说明书摘要 章节")

    def _on_spec_check(self, kind: str):
        """运行说明书检查（embodiment / abstract）：填共享结果表 + 1框 内联高亮。"""
        if not self.doc_data:
            self._show_toast("请先打开文档！", "error")
            return
        # 落 1框 结构化编辑到内存，确保检查基于最新内容
        self.content_area.flush_all()
        try:
            from core.spec_check import check_embodiment_numbering, check_abstract_length
            paras = self.doc_data["paragraphs"]
            sections = self.doc_data.get("sections", {})
            if kind == "embodiment":
                results = check_embodiment_numbering(paras, sections)
            else:
                results = check_abstract_length(paras, sections)
        except Exception as e:
            import traceback as tb
            QMessageBox.critical(
                self, "检查失败",
                f"说明书检查出现异常：\n{e}\n\n{tb.format_exc()}"
            )
            return

        self._spec_results = results
        self._spec_kind = kind
        self._render_spec_results(results, kind)
        if results:
            self._show_toast(f"发现 {len(results)} 处问题", "warning")
        else:
            name = "实施例编号" if kind == "embodiment" else "摘要字数"
            self._show_toast(f"{name}检查：未发现问题", "success")

    def _render_spec_results(self, results: list, kind: str):
        """渲染说明书检查结果到共享表，并把可定位项在 1框 标黄。"""
        KIND_LABELS = {
            "emb_start": "起始", "emb_gap": "缺号", "emb_dup": "重号",
            "emb_order": "顺序", "abstract_len": "字数",
        }
        title = "实施例编号检查" if kind == "embodiment" else "摘要字数检查"
        self.spec_result_group.setTitle(f"📑 {title}（{len(results)}）")
        self.spec_result_table.setRowCount(0)
        for row_idx, item in enumerate(results):
            self.spec_result_table.insertRow(row_idx)

            kind_item = QTableWidgetItem(KIND_LABELS.get(item.get("kind"), item.get("kind", "")))
            kind_item.setFlags(kind_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.spec_result_table.setItem(row_idx, 0, kind_item)

            # 列1：定位（实施例标题文本；摘要类不展示越界长串，固定显示「摘要」）
            loc = "摘要" if item.get("kind") == "abstract_len" else (item.get("wrong") or "")
            loc_item = QTableWidgetItem(loc)
            loc_item.setFlags(loc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.spec_result_table.setItem(row_idx, 1, loc_item)

            msg = item.get("message", "")
            msg_item = QTableWidgetItem(msg)
            msg_item.setFlags(msg_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            sug = (item.get("suggestion") or "").strip()
            msg_item.setToolTip(f"{msg}\n建议：{sug}" if sug else msg)
            self.spec_result_table.setItem(row_idx, 2, msg_item)

            ig_item = QTableWidgetItem("忽略")
            ig_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            ig_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            ig_item.setForeground(self._fg("#00897b", "#4dd0c4"))
            ig_font = QFont()
            ig_font.setBold(True)
            ig_item.setFont(ig_font)
            ig_item.setToolTip("点击忽略此条")
            self.spec_result_table.setItem(row_idx, 3, ig_item)

        # 1框 内联高亮（实施例→说明书 tab / 摘要→说明书摘要 tab；空结果→清空高亮）
        self.content_area.highlight_issues(self._spec_highlight_items(results))

    def _spec_highlight_items(self, results: list) -> list:
        """把说明书结果转成内联高亮 {para_idx, wrong, occurrence}；不可定位项跳过。"""
        out = []
        for r in results:
            wrong = r.get("wrong")
            pid = r.get("para_idx", -1)
            if wrong and isinstance(pid, int) and pid >= 0:
                out.append({"para_idx": pid, "wrong": wrong, "occurrence": 1})
        return out

    def _on_spec_cell_clicked(self, row: int, col: int):
        """单击「说明」(col 2)→跳转 1框 标红；「操作」(col 3)→移除该行。"""
        if col == 2:
            self._locate_spec_in_content(row)
        elif col == 3:
            self._on_spec_ignore_row(row)

    def _locate_spec_in_content(self, row: int):
        """单击「说明」：在 1框 对应标签跳转到该处并标红（实施例→说明书 / 摘要→摘要）。"""
        if row < 0 or row >= len(self._spec_results):
            return
        item = self._spec_results[row]
        wrong = item.get("wrong") or ""
        ok = False
        if wrong:
            ok = self.content_area.locate_issue(item.get("para_idx", -1), wrong, 1)
        if not ok:
            self._show_toast("未能在原文中定位该处（可能内容已变动）", "warning")

    def _on_spec_ignore_row(self, row: int):
        """忽略本条：仅从当前结果表移除并刷新（不写词库；编号类无术语可记）。"""
        if row < 0 or row >= len(self._spec_results):
            return
        self._spec_results = [r for i, r in enumerate(self._spec_results) if i != row]
        self._render_spec_results(self._spec_results, self._spec_kind or "embodiment")

    def _on_claim_ignore_dialog(self):
        """打开忽略词库编辑对话框"""
        try:
            from ui.dialogs.claim_ignore_dialog import ClaimIgnoreDialog
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
            from ui.dialogs.boundary_blacklist_dialog import BoundaryBlacklistDialog
        except Exception as e:
            QMessageBox.critical(self, "无法打开", f"加载黑名单词库编辑器失败：\n{e}")
            return
        dlg = BoundaryBlacklistDialog(self)
        dlg.exec()

    # ── 不确定用语检查 的信息弹窗 + 词库入口 ──
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

