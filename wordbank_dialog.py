# -*- coding: utf-8 -*-
"""
wordbank_dialog.py — 错别字词库编辑对话框
- 4 列双条目紧凑布局（每行显示 2 个条目）
- 内置 = 灰底，自定义 = 白底（无类型列）
- 支持搜索 / 添加 / 删除 / 导入 / 导出
- 内置条目可删除（加入禁用名单，仅本机生效）
"""
import json
import os
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QAbstractItemView,
    QFileDialog, QLineEdit
)

from typo_wordbank import WORDBANK as BUILTIN_WORDBANK
from config_manager import (
    load_user_wordbank, save_user_wordbank,
    load_disabled_builtin_wrongs, save_disabled_builtin_wrongs,
)

# 视觉常量
_BUILTIN_BG = QColor(240, 240, 240)      # 内置：浅灰底
_BUILTIN_FG = QColor(110, 110, 110)      # 内置：深灰字
_USER_BG = QColor(255, 255, 255)         # 自定义：白底
_USER_FG = QColor(30, 30, 30)            # 自定义：黑字


class WordbankDialog(QDialog):
    """错别字词库编辑器（4 列双条目布局 + 搜索）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("错别字词库编辑器")
        self.setMinimumSize(860, 600)
        self.setModal(True)

        # 数据模型：扁平条目列表
        # 每项: {"wrong": str, "suggestion": str, "kind": "builtin"|"user"}
        self._entries: list = []

        # 记录：用户删除的内置条目 wrong（本机生效）
        self._disabled_builtin_wrongs: set = set(load_disabled_builtin_wrongs())

        # 当前搜索关键字（小写）
        self._search_text: str = ""

        self._build_ui()
        self._load_entries_from_disk()
        self._rebuild_table()

    # ─────────────────────────────────────────
    # UI 构建
    # ─────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # 顶部说明
        hint = QLabel(
            "下表显示当前生效的错别字词库。<br>"
            "• <b>灰底</b>：内置规则（只读，可删除，删除仅对本机生效）<br>"
            "• <b>白底</b>：用户自定义规则（可自由添加 / 修改 / 删除） · 保存后即生效<br>"
            "<a href='#nlp' style='color:#0066cc;'>增加离线NLP引擎（专业用户）</a>"
        )
        hint.setTextFormat(Qt.TextFormat.RichText)
        hint.setWordWrap(True)
        hint.linkActivated.connect(self._on_open_nlp_tutorial)
        layout.addWidget(hint)

        # 搜索行
        search_row = QHBoxLayout()
        search_label = QLabel("🔎 搜索：")
        search_row.addWidget(search_label)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入关键字过滤 wrong / suggestion（区分大小写无关）")
        self.search_edit.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_edit, 1)

        clear_search_btn = QPushButton("清除")
        clear_search_btn.clicked.connect(lambda: self.search_edit.clear())
        search_row.addWidget(clear_search_btn)
        layout.addLayout(search_row)

        # 表格（4 列：wrong1 | sug1 | wrong2 | sug2）
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["错误写法", "建议正确", "错误写法", "建议正确"]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.verticalHeader().setDefaultSectionSize(22)
        # 2 / 3 列之间加粗分界线（通过左边框制造视觉分隔）
        self.table.setStyleSheet(
            "QTableWidget::item { padding: 1px 4px; }"
            "QHeaderView::section { padding: 2px 6px; }"
            "QTableWidget::item:selected { background-color: #4a90e2; color: white; }"
        )
        # 关键：第 3 列（索引 2）左侧加粗分界线 —— 用 item 绘制边框实现
        # 由于 Qt 样式表不支持单独列的 border-left，这里在 _append_row 时为该列单元格设置粗边框字符
        # 替代方案：调整 header section 1 的固定宽度？不可行。
        # 我们通过在第 2/3 列之间 insert 一个窄分隔列，然后隐藏不可行。
        # 采用另一种方式：通过自定义 delegate 或在第 3 列 item 上画左边框。
        # 简单起见：为第 3 列每个 cell 设置左侧文本前空格 + 让 header 第 3 列标题左侧加符号。
        # 这里直接通过 setSpan 不适用，改为使用垂直 frame 分隔符 —— 但 QTableWidget 不支持。
        # 最终方案：在样式表中通过 nth-item 不支持。
        # => 用「列间距 + 列头文字 + 分隔颜色」：通过在样式表为整表的网格线加深实现整体感，
        #    然后把 2 列之间的分隔特别渲染 —— 采用 item delegate 最稳。
        self._install_divider_delegate()
        layout.addWidget(self.table, 1)

        # 按钮行
        btn_row = QHBoxLayout()

        add_btn = QPushButton("➕ 添加条目")
        add_btn.clicked.connect(self._on_add_row)
        btn_row.addWidget(add_btn)

        del_btn = QPushButton("🗑️ 删除选中")
        del_btn.clicked.connect(self._on_delete_row)
        btn_row.addWidget(del_btn)

        import_btn = QPushButton("📥 导入…")
        import_btn.setToolTip("从 JSON / CSV 文件导入用户词条")
        import_btn.clicked.connect(self._on_import)
        btn_row.addWidget(import_btn)

        export_btn = QPushButton("📤 导出…")
        export_btn.setToolTip("导出当前自定义词条到 JSON 或 CSV")
        export_btn.clicked.connect(self._on_export)
        btn_row.addWidget(export_btn)

        self.count_label = QLabel("")
        self.count_label.setObjectName("subtitleLabel")
        btn_row.addWidget(self.count_label)

        btn_row.addStretch()

        save_btn = QPushButton("💾 保存并关闭")
        save_btn.setObjectName("primaryBtn")
        save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(save_btn)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        layout.addLayout(btn_row)

    def _on_open_nlp_tutorial(self, *_):
        """点击「专业用户可增加离线 NLP 引擎」时弹出教程对话框（三级菜单）。
        优先复用主窗口的实现，避免重复代码；找不到则给出降级提示。"""
        parent = self.parent()
        # parent 可能不是 MainWindow 实例（例如未来作为独立工具调用）
        if parent is not None and hasattr(parent, "_on_open_pycorrector_dialog"):
            try:
                parent._on_open_pycorrector_dialog()
                return
            except Exception as e:
                QMessageBox.warning(self, "无法打开教程", f"教程对话框加载失败：\n{e}")
                return
        QMessageBox.information(
            self, "离线 NLP 引擎",
            "本功能仅供专业用户：需在系统中额外安装 pycorrector 离线 NLP 库。\n"
            "详细安装步骤请在主程序中查看（当前对话框无法独立访问教程）。"
        )

    def _install_divider_delegate(self):
        """在第 2/3 列之间画一条加粗分界线（通过 item delegate 实现）"""
        from PyQt6.QtWidgets import QStyledItemDelegate
        from PyQt6.QtGui import QPen

        class _DividerDelegate(QStyledItemDelegate):
            def paint(self, painter, option, index):
                super().paint(painter, option, index)
                # 第 3 列（col==2）左侧画一条深色粗竖线
                if index.column() == 2:
                    painter.save()
                    pen = QPen(QColor(80, 80, 80))
                    pen.setWidth(2)
                    painter.setPen(pen)
                    rect = option.rect
                    painter.drawLine(rect.left(), rect.top(), rect.left(), rect.bottom())
                    painter.restore()

        self._divider_delegate = _DividerDelegate(self.table)
        self.table.setItemDelegate(self._divider_delegate)

    # ─────────────────────────────────────────
    # 数据加载与渲染
    # ─────────────────────────────────────────

    def _load_entries_from_disk(self):
        """把内置 + 用户词库合并为扁平条目列表（用户条目置前）。"""
        self._entries = []
        user_entries = load_user_wordbank()
        user_wrongs = {e["wrong"] for e in user_entries}

        for e in user_entries:
            self._entries.append({
                "wrong": e["wrong"],
                "suggestion": e["suggestion"],
                "kind": "user",
            })

        for e in BUILTIN_WORDBANK:
            if e["wrong"] in user_wrongs:
                continue
            if e["wrong"] in self._disabled_builtin_wrongs:
                continue
            self._entries.append({
                "wrong": e["wrong"],
                "suggestion": e["suggestion"],
                "kind": "builtin",
            })

    def _filter_entries(self) -> list:
        """按搜索关键字过滤（返回 (index, entry) 列表）"""
        if not self._search_text:
            return list(enumerate(self._entries))
        kw = self._search_text
        result = []
        for i, e in enumerate(self._entries):
            if kw in e["wrong"].lower() or kw in e["suggestion"].lower():
                result.append((i, e))
        return result

    def _rebuild_table(self):
        """按当前 _entries + 过滤条件重绘整个表格（2 条/行）"""
        self.table.blockSignals(True)
        self.table.setRowCount(0)

        filtered = self._filter_entries()
        # 每 2 条一行
        for row_idx in range((len(filtered) + 1) // 2):
            self.table.insertRow(row_idx)
            left = filtered[row_idx * 2]
            right = filtered[row_idx * 2 + 1] if row_idx * 2 + 1 < len(filtered) else None
            self._set_cell_pair(row_idx, 0, left)
            if right is not None:
                self._set_cell_pair(row_idx, 2, right)
            else:
                # 空占位，禁止编辑
                for col in (2, 3):
                    placeholder = QTableWidgetItem("")
                    placeholder.setFlags(placeholder.flags() & ~Qt.ItemFlag.ItemIsEditable
                                          & ~Qt.ItemFlag.ItemIsSelectable)
                    placeholder.setBackground(QBrush(_USER_BG))
                    self.table.setItem(row_idx, col, placeholder)

        self.table.blockSignals(False)
        self.table.itemChanged.connect(self._on_item_changed)
        self._update_count()

    def _set_cell_pair(self, row: int, col_start: int, indexed_entry):
        """在指定 (row, col_start..col_start+1) 设置 wrong+suggestion 两个单元格。
        indexed_entry = (entry_index, entry_dict)
        """
        entry_idx, entry = indexed_entry
        is_builtin = entry["kind"] == "builtin"

        wrong_item = QTableWidgetItem(entry["wrong"])
        sug_item = QTableWidgetItem(entry["suggestion"])

        # 存储条目索引用于回写
        wrong_item.setData(Qt.ItemDataRole.UserRole, entry_idx)
        sug_item.setData(Qt.ItemDataRole.UserRole, entry_idx)

        bg = _BUILTIN_BG if is_builtin else _USER_BG
        fg = _BUILTIN_FG if is_builtin else _USER_FG
        for it in (wrong_item, sug_item):
            it.setBackground(QBrush(bg))
            it.setForeground(QBrush(fg))

        if is_builtin:
            wrong_item.setFlags(wrong_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            sug_item.setFlags(sug_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        self.table.setItem(row, col_start, wrong_item)
        self.table.setItem(row, col_start + 1, sug_item)

    def _on_item_changed(self, item: QTableWidgetItem):
        """用户编辑白底单元格 → 回写 _entries"""
        entry_idx = item.data(Qt.ItemDataRole.UserRole)
        if entry_idx is None or entry_idx >= len(self._entries):
            return
        entry = self._entries[entry_idx]
        if entry["kind"] != "user":
            return
        col = item.column()
        text = item.text()
        if col % 2 == 0:  # wrong 列
            entry["wrong"] = text
        else:  # suggestion 列
            entry["suggestion"] = text

    def _update_count(self):
        builtin_n = sum(1 for e in self._entries if e["kind"] == "builtin")
        user_n = len(self._entries) - builtin_n
        visible = len(self._filter_entries())
        if self._search_text:
            self.count_label.setText(
                f"  共 {len(self._entries)} 条（内置 {builtin_n}，自定义 {user_n}） · 匹配 {visible} 条"
            )
        else:
            self.count_label.setText(
                f"  共 {len(self._entries)} 条（内置 {builtin_n}，自定义 {user_n}）"
            )

    # ─────────────────────────────────────────
    # 交互：搜索 / 添加 / 删除
    # ─────────────────────────────────────────

    def _on_search_changed(self, text: str):
        self._search_text = text.strip().lower()
        self._rebuild_table()

    def _on_add_row(self):
        """在 _entries 开头插入一个空白用户条目，然后重绘并聚焦"""
        self._entries.insert(0, {"wrong": "", "suggestion": "", "kind": "user"})
        # 添加新行时清空搜索以便立即看到
        if self._search_text:
            self.search_edit.clear()  # 会触发重绘
        else:
            self._rebuild_table()
        # 聚焦到第一行第一列
        self.table.setCurrentCell(0, 0)
        item = self.table.item(0, 0)
        if item is not None:
            self.table.editItem(item)

    def _on_delete_row(self):
        """根据选中单元格删除对应条目
        - 一个选中 cell 对应一个条目（cell_pair = col0-1 或 col2-3）
        - 自定义 → 直接删除
        - 内置 → 加入禁用名单并删除
        """
        selected = self.table.selectedIndexes()
        if not selected:
            QMessageBox.information(self, "提示", "请先选中要删除的单元格（支持多选）")
            return

        # 收集所有唯一 entry_idx
        entry_idxs = set()
        for idx in selected:
            item = self.table.item(idx.row(), idx.column())
            if item is None:
                continue
            eidx = item.data(Qt.ItemDataRole.UserRole)
            if eidx is not None:
                entry_idxs.add(int(eidx))

        if not entry_idxs:
            QMessageBox.information(self, "提示", "选中的单元格不是有效条目")
            return

        deleted_user = 0
        deleted_builtin = 0
        # 从大到小删除以保持索引稳定
        for eidx in sorted(entry_idxs, reverse=True):
            if eidx >= len(self._entries):
                continue
            entry = self._entries[eidx]
            if entry["kind"] == "builtin":
                if entry["wrong"]:
                    self._disabled_builtin_wrongs.add(entry["wrong"])
                deleted_builtin += 1
            else:
                deleted_user += 1
            self._entries.pop(eidx)

        self._rebuild_table()

        if deleted_builtin:
            QMessageBox.information(
                self, "已删除",
                f"已删除 {deleted_user} 条自定义、{deleted_builtin} 条内置规则。\n"
                "内置规则的删除仅对本机生效，保存后下次启动仍隐藏。"
            )

    # ─────────────────────────────────────────
    # 保存
    # ─────────────────────────────────────────

    def _on_save(self):
        """收集所有 user 条目，写入用户词库 JSON；同时持久化禁用名单"""
        user_entries = []
        for e in self._entries:
            if e["kind"] != "user":
                continue
            w = (e["wrong"] or "").strip()
            s = (e["suggestion"] or "").strip()
            if not w or not s or w == s:
                continue
            user_entries.append({"wrong": w, "suggestion": s})

        try:
            save_user_wordbank(user_entries)
            save_disabled_builtin_wrongs(self._disabled_builtin_wrongs)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"无法写入用户词库文件：\n{e}")
            return

        disabled_n = len(self._disabled_builtin_wrongs)
        extra = f"，已禁用 {disabled_n} 条内置规则" if disabled_n else ""
        QMessageBox.information(
            self, "已保存",
            f"用户词库已保存，共 {len(user_entries)} 条自定义规则{extra}。\n"
            "再次点击「开始检查」即可生效。"
        )
        self.accept()

    # ─────────────────────────────────────────
    # 导入 / 导出
    # ─────────────────────────────────────────

    def _collect_user_rows(self) -> list:
        rows = []
        for e in self._entries:
            if e["kind"] != "user":
                continue
            w = (e["wrong"] or "").strip()
            s = (e["suggestion"] or "").strip()
            if w and s and w != s:
                rows.append({"wrong": w, "suggestion": s})
        return rows

    def _on_export(self):
        rows = self._collect_user_rows()
        if not rows:
            QMessageBox.information(self, "无内容", "当前没有可导出的「自定义」词条。")
            return

        path, selected_filter = QFileDialog.getSaveFileName(
            self, "导出用户词库", "user_wordbank.json",
            "JSON 文件 (*.json);;CSV 文件 (*.csv);;所有文件 (*)"
        )
        if not path:
            return

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".csv" or "csv" in selected_filter.lower():
                if not path.lower().endswith(".csv"):
                    path += ".csv"
                self._write_csv(path, rows)
            else:
                if not path.lower().endswith(".json"):
                    path += ".json"
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(rows, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"无法写入文件：\n{e}")
            return

        QMessageBox.information(self, "已导出", f"已导出 {len(rows)} 条用户词条至：\n{path}")

    def _on_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "导入用户词库", "",
            "词库文件 (*.json *.csv);;JSON 文件 (*.json);;CSV 文件 (*.csv);;所有文件 (*)"
        )
        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                imported = self._read_csv(path)
            else:
                with open(path, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                if not isinstance(raw, list):
                    raise ValueError("JSON 顶层必须为数组")
                imported = [
                    {"wrong": str(x.get("wrong", "")).strip(),
                     "suggestion": str(x.get("suggestion", "")).strip()}
                    for x in raw if isinstance(x, dict)
                ]
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"解析文件出错：\n{e}")
            return

        imported = [e for e in imported if e["wrong"] and e["suggestion"] and e["wrong"] != e["suggestion"]]
        if not imported:
            QMessageBox.warning(self, "无有效条目", "文件中未发现有效的词条。")
            return

        # 合并到现有自定义条目：相同 wrong 以导入版本覆盖
        existing = {e["wrong"]: idx for idx, e in enumerate(self._entries) if e["kind"] == "user"}
        added = 0
        updated = 0
        for item in imported:
            if item["wrong"] in existing:
                idx = existing[item["wrong"]]
                if self._entries[idx]["suggestion"] != item["suggestion"]:
                    self._entries[idx]["suggestion"] = item["suggestion"]
                    updated += 1
            else:
                self._entries.insert(0, {
                    "wrong": item["wrong"],
                    "suggestion": item["suggestion"],
                    "kind": "user",
                })
                added += 1

        self._rebuild_table()
        QMessageBox.information(
            self, "导入完成",
            f"成功导入 {len(imported)} 条（新增 {added}，更新 {updated}）。\n点击「保存并关闭」后生效。"
        )

    @staticmethod
    def _write_csv(path: str, rows: list):
        import csv
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["wrong", "suggestion"])
            for r in rows:
                w.writerow([r["wrong"], r["suggestion"]])

    @staticmethod
    def _read_csv(path: str) -> list:
        import csv
        results = []
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames and "wrong" in reader.fieldnames and "suggestion" in reader.fieldnames:
                for row in reader:
                    results.append({
                        "wrong": (row.get("wrong") or "").strip(),
                        "suggestion": (row.get("suggestion") or "").strip(),
                    })
            else:
                f.seek(0)
                for row in csv.reader(f):
                    if len(row) >= 2:
                        results.append({
                            "wrong": row[0].strip(),
                            "suggestion": row[1].strip(),
                        })
        return results
