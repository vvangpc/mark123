# -*- coding: utf-8 -*-
"""
dialogs/base_wordbank_dialog.py — 同构词库对话框基类

为以下三类"单列字符串"词库对话框提供公共 UI + 行为：
  - 不确定用语词库 (ClaimIgnoreDialog)
  - 动态截断黑名单词库 (BoundaryBlacklistDialog)
  - 重复字词忽略词库 (DupIgnoreDialog)

子类通过覆写类属性 + 抽象方法声明差异：
  - TITLE / HINT_HTML / ADD_PLACEHOLDER / EXPORT_FILENAME / MIN_SIZE
  - HAS_RESTORE_DEFAULTS (是否显示"恢复内置"按钮)
  - load_items() / save_items(items) / get_builtin()

UI 结构、过滤、增删、导入导出逻辑完全复用。
"""
import json

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QFileDialog, QLineEdit, QAbstractItemView,
)

# 网格列数：词条在对话框中以 N 列展示
GRID_COLS = 4


class BaseWordbankDialog(QDialog):
    """单列字符串词库对话框基类。子类通过类属性 + 抽象方法声明差异。"""

    # ── 子类需覆写的配置 ──────────────────────────────────────
    TITLE: str = "词库"
    HINT_HTML: str = ""
    ADD_PLACEHOLDER: str = "输入要添加的词后按回车或点「添加」"
    EXPORT_FILENAME: str = "wordbank.json"
    MIN_SIZE: tuple = (520, 600)
    HAS_RESTORE_DEFAULTS: bool = False

    # 各类业务文本（可选覆写）
    SAVE_SUFFIX: str = "下次使用时即生效。"
    DUPLICATE_ITEM_HINT: str = "「{}」已在词库中。"
    EXPORT_TITLE: str = "导出词库"
    IMPORT_TITLE: str = "导入词库"
    EMPTY_EXPORT_HINT: str = "当前词库为空。"

    # ── 抽象接口 ──────────────────────────────────────────────
    def load_items(self) -> list:
        raise NotImplementedError

    def save_items(self, items: list) -> None:
        raise NotImplementedError

    def get_builtin(self) -> list:
        return []

    # ── 生命周期 ──────────────────────────────────────────────
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(self.TITLE)
        self.setMinimumSize(*self.MIN_SIZE)
        self.setModal(True)

        self._items: list = list(self.load_items())
        self._search_text: str = ""

        self._build_ui()
        self._rebuild_list()

    # ── UI 构建 ──────────────────────────────────────────────
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        if self.HINT_HTML:
            hint = QLabel(self.HINT_HTML)
            hint.setWordWrap(True)
            layout.addWidget(hint)

        # 搜索行
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("🔎 搜索："))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入关键字过滤列表")
        self.search_edit.textChanged.connect(self._on_search_changed)
        search_row.addWidget(self.search_edit, 1)
        clear_btn = QPushButton("清除")
        clear_btn.clicked.connect(lambda: self.search_edit.clear())
        search_row.addWidget(clear_btn)
        layout.addLayout(search_row)

        # 快速添加
        add_row = QHBoxLayout()
        add_row.addWidget(QLabel("➕ 添加："))
        self.add_edit = QLineEdit()
        self.add_edit.setPlaceholderText(self.ADD_PLACEHOLDER)
        self.add_edit.returnPressed.connect(self._on_add_clicked)
        add_row.addWidget(self.add_edit, 1)
        add_btn = QPushButton("添加")
        add_btn.setObjectName("accentBtn")
        add_btn.clicked.connect(self._on_add_clicked)
        add_row.addWidget(add_btn)
        layout.addLayout(add_row)

        # 网格
        self.grid = QTableWidget(0, GRID_COLS)
        self.grid.horizontalHeader().setVisible(False)
        self.grid.verticalHeader().setVisible(False)
        self.grid.setShowGrid(False)
        self.grid.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.grid.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.grid.setAlternatingRowColors(True)
        self.grid.verticalHeader().setDefaultSectionSize(30)
        hh = self.grid.horizontalHeader()
        for c in range(GRID_COLS):
            hh.setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.grid, 1)

        # 底部按钮
        btn_row = QHBoxLayout()

        del_btn = QPushButton("🗑️ 删除选中")
        del_btn.clicked.connect(self._on_delete_clicked)
        btn_row.addWidget(del_btn)

        if self.HAS_RESTORE_DEFAULTS:
            restore_btn = QPushButton("🔄 恢复内置")
            restore_btn.setToolTip("把缺失的内置默认词合并回当前列表（不会删除已有条目）")
            restore_btn.clicked.connect(self._on_restore_defaults)
            btn_row.addWidget(restore_btn)

        import_btn = QPushButton("📥 导入…")
        import_btn.setToolTip("从 JSON / TXT 文件导入（TXT 每行一个词）")
        import_btn.clicked.connect(self._on_import)
        btn_row.addWidget(import_btn)

        export_btn = QPushButton("📤 导出…")
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

    # ── 列表渲染 ──────────────────────────────────────────────
    def _rebuild_list(self):
        kw = self._search_text
        filtered = [s for s in self._items if (not kw) or (kw in s.lower())]

        rows = (len(filtered) + GRID_COLS - 1) // GRID_COLS
        self.grid.clearContents()
        self.grid.setRowCount(rows)
        for idx, s in enumerate(filtered):
            r, c = divmod(idx, GRID_COLS)
            item = QTableWidgetItem(s)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.grid.setItem(r, c, item)
        # 最后一行右侧补空占位（禁选），避免拖选出现假格
        if rows:
            last_idx = len(filtered)
            if last_idx % GRID_COLS:
                last_r = rows - 1
                for c in range(last_idx % GRID_COLS, GRID_COLS):
                    placeholder = QTableWidgetItem("")
                    placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
                    self.grid.setItem(last_r, c, placeholder)
        self._update_count(len(filtered))

    def _update_count(self, visible: int):
        if self._search_text:
            self.count_label.setText(f"  共 {len(self._items)} 条 · 匹配 {visible} 条")
        else:
            self.count_label.setText(f"  共 {len(self._items)} 条")

    # ── 事件处理 ──────────────────────────────────────────────
    def _on_search_changed(self, text: str):
        self._search_text = text.strip().lower()
        self._rebuild_list()

    def _on_add_clicked(self):
        text = self.add_edit.text().strip()
        if not text:
            return
        if text in self._items:
            QMessageBox.information(self, "已存在", self.DUPLICATE_ITEM_HINT.format(text))
            self.add_edit.clear()
            return
        self._items.insert(0, text)
        self.add_edit.clear()
        if self._search_text:
            self.search_edit.clear()
        else:
            self._rebuild_list()

    def _on_delete_clicked(self):
        selected = self.grid.selectedItems()
        to_remove = {it.text() for it in selected if it and it.text()}
        if not to_remove:
            QMessageBox.information(self, "提示", "请先选中要删除的条目")
            return
        self._items = [s for s in self._items if s not in to_remove]
        self._rebuild_list()

    def _on_restore_defaults(self):
        builtin = self.get_builtin()
        existing = set(self._items)
        added = 0
        for w in builtin:
            if w and w not in existing:
                self._items.append(w)
                existing.add(w)
                added += 1
        self._rebuild_list()
        if added:
            QMessageBox.information(self, "已合并", f"新增 {added} 条内置默认词。")
        else:
            QMessageBox.information(self, "无变化", "当前列表已包含所有内置默认词。")

    def _on_save(self):
        try:
            self.save_items(self._items)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"无法写入文件：\n{e}")
            return
        QMessageBox.information(
            self, "已保存",
            f"词库已保存，共 {len(self._items)} 条。\n{self.SAVE_SUFFIX}"
        )
        self.accept()

    def _on_export(self):
        if not self._items:
            QMessageBox.information(self, "无内容", self.EMPTY_EXPORT_HINT)
            return
        path, selected_filter = QFileDialog.getSaveFileName(
            self, self.EXPORT_TITLE, self.EXPORT_FILENAME,
            "JSON 文件 (*.json);;文本文件 (*.txt);;所有文件 (*)"
        )
        if not path:
            return
        try:
            if path.lower().endswith(".txt") or "txt" in selected_filter.lower():
                if not path.lower().endswith(".txt"):
                    path += ".txt"
                with open(path, "w", encoding="utf-8") as f:
                    for s in self._items:
                        f.write(s + "\n")
            else:
                if not path.lower().endswith(".json"):
                    path += ".json"
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(self._items, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"无法写入文件：\n{e}")
            return
        QMessageBox.information(self, "已导出", f"已导出 {len(self._items)} 条至：\n{path}")

    def _on_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self, self.IMPORT_TITLE, "",
            "词库文件 (*.json *.txt);;JSON 文件 (*.json);;文本文件 (*.txt);;所有文件 (*)"
        )
        if not path:
            return
        try:
            if path.lower().endswith(".json"):
                with open(path, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                if not isinstance(raw, list):
                    raise ValueError("JSON 顶层必须为数组")
                imported = [str(x).strip() for x in raw if str(x).strip()]
            else:
                with open(path, "r", encoding="utf-8") as f:
                    imported = [line.strip() for line in f if line.strip()]
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"解析文件出错：\n{e}")
            return

        if not imported:
            QMessageBox.warning(self, "无有效条目", "文件中未发现有效的词条。")
            return

        existing = set(self._items)
        added = 0
        for s in imported:
            if s not in existing:
                self._items.insert(0, s)
                existing.add(s)
                added += 1
        self._rebuild_list()
        QMessageBox.information(
            self, "导入完成",
            f"成功导入 {len(imported)} 条（新增 {added}，重复 {len(imported) - added}）。"
        )
