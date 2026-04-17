# -*- coding: utf-8 -*-
"""
workers.py — 后台线程与通用小部件
抽离自 main_window.py，包含：
- _longest_nonspace_run / _is_pycorrector_available 两个工具函数
- AnnotateWorker / CleanWorker 两个后台 QThread
- ToastWidget 右上角悬浮提示
"""
import traceback

from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt6.QtWidgets import QLabel

from annotator import (
    smart_annotate_section, smart_remove_section,
)
from cleaner import (
    remove_suoshu, unify_halfwidth_punct, convert_fullwidth_to_halfwidth,
    detect_orphan_marks, fix_consecutive_punct, detect_orphan_figures,
    check_typos_wordbank, check_typos_pycorrector, check_duplicate_words,
    merge_typo_results, apply_typo_corrections,
)


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
            self.error.emit(f"操作失败：{str(e)}\n{traceback.format_exc()}")


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
