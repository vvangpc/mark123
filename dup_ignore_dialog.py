# -*- coding: utf-8 -*-
"""
dup_ignore_dialog.py — 重复字词忽略词库编辑器
- 用户添加到此列表中的字 / 词，将在「重复字词检查」中被忽略
- 支持搜索 / 添加 / 删除 / 导入 / 导出
"""
import json
import os
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QFileDialog, QLineEdit, QAbstractItemView,
)

from config_manager import load_dup_ignore_list, save_dup_ignore_list

# 列数：词条在对话框中以 N 列网格展示
GRID_COLS = 4


class DupIgnoreDialog(QDialog):
    """重复字词忽略词库编辑器"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("重复字词忽略词库")
        self.setMinimumSize(480, 560)
        self.setModal(True)

        # 数据：扁平字符串列表
        self._items: list = list(load_dup_ignore_list())
        self._search_text: str = ""

        self._build_ui()
        self._rebuild_list()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        hint = QLabel(
            "添加到此处的字 / 词，将在「重复字词检查」中被忽略。\n"
            "例如：将「所述」加入忽略后，「所述所述」不再被标红。\n"
            "• 支持按单字 / 词组匹配（同时匹配「重复单元」和「完整重复串」）\n"
            "• 保存后下次点「重复字词检查」即生效"
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        # 搜索 + 快速添加行
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

        # 快速添加输入框
        add_row = QHBoxLayout()
        add_row.addWidget(QLabel("➕ 添加："))
        self.add_edit = QLineEdit()
        self.add_edit.setPlaceholderText("输入要忽略的字 / 词后按回车或点「添加」")
        self.add_edit.returnPressed.connect(self._on_add_clicked)
        add_row.addWidget(self.add_edit, 1)
        add_btn = QPushButton("添加")
        add_btn.setObjectName("accentBtn")
        add_btn.clicked.connect(self._on_add_clicked)
        add_row.addWidget(add_btn)
        layout.addLayout(add_row)

        # 4 列网格
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

    # ─────────────────────────────────────────

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

    def _on_search_changed(self, text: str):
        self._search_text = text.strip().lower()
        self._rebuild_list()

    def _on_add_clicked(self):
        text = self.add_edit.text().strip()
        if not text:
            return
        if text in self._items:
            QMessageBox.information(self, "已存在", f"「{text}」已在忽略列表中。")
            self.add_edit.clear()
            return
        self._items.insert(0, text)
        self.add_edit.clear()
        # 清空搜索以便立即可见
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

    def _on_save(self):
        try:
            save_dup_ignore_list(self._items)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"无法写入文件：\n{e}")
            return
        QMessageBox.information(
            self, "已保存",
            f"忽略词库已保存，共 {len(self._items)} 条。\n下次点「重复字词检查」即生效。"
        )
        self.accept()

    def _on_export(self):
        if not self._items:
            QMessageBox.information(self, "无内容", "当前忽略词库为空。")
            return
        path, selected_filter = QFileDialog.getSaveFileName(
            self, "导出忽略词库", "dup_ignore.json",
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
            self, "导入忽略词库", "",
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
