# -*- coding: utf-8 -*-
"""
tabs/typo_tab.py — 错字 / 重复字词检查 Tab

以 Mixin 形式抽离自 main_window.py，由 MainWindow 继承。保持方法体内所有
self.xxx 引用不变。Mixin 依赖 MainWindow 提供：
  - self.doc_data
  - self.typo_data / self.dup_data / self._current_check_kind
  - self._show_toast / self._add_history / self._start_clean_worker
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QPlainTextEdit, QGroupBox, QDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QAbstractItemView, QFrame, QCheckBox,
)
from PyQt6.QtCore import Qt

from workers import _is_pycorrector_available


class TypoTabMixin:
    """错字 / 重复字词检查 Tab 的方法集合，供 MainWindow 继承。"""

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
        self.typo_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.typo_table.setColumnWidth(3, 76)
        self.typo_table.setAlternatingRowColors(True)
        self.typo_table.verticalHeader().setVisible(False)
        # 行高下限 32px，避免"忽略"等按钮文字在压缩行高下被裁剪
        self.typo_table.verticalHeader().setDefaultSectionSize(32)
        self.typo_table.verticalHeader().setMinimumSectionSize(32)
        action_v.addWidget(self.typo_table, 1)

        layout.addWidget(action_group, 1)

        # 当前显示的检查类型："typo" 或 "dup" 或 None
        self._current_check_kind = None
        return widget

    # ─────────────────────────────────────────
    # 权利要求书检查 Tab
    # ─────────────────────────────────────────

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
            ignore_btn.setObjectName("rowActionBtn")
            ignore_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            ignore_btn.setMinimumHeight(28)
            ignore_btn.clicked.connect(self._on_check_ignore_row)
            self.typo_table.setCellWidget(row_idx, 3, ignore_btn)
            # 显式给当前行一个下限行高，避免 Qt 以 item 的 sizeHint 压缩行高
            self.typo_table.setRowHeight(row_idx, 32)

        # 计数标签 + 应用按钮启用状态
        kind_text = "错别字" if self._current_check_kind == "typo" else "重复字词"
        self.typo_count_label.setText(f"  当前显示：{kind_text}  ·  共 {len(results)} 处")
        self.typo_apply_btn.setEnabled(len(results) > 0)

