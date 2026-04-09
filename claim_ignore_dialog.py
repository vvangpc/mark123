# -*- coding: utf-8 -*-
"""
claim_ignore_dialog.py — 不确定用语词库编辑器

维护权利要求书中"不应出现"的含糊 / 不确定用语（例如：约、大概、可能、
左右、优选…）。勾选「开始检查」时，会把这些词当作 `vague` 类问题报出。

- 首次打开：自动以内置默认列表填充；
- 用户可增删、导入、导出；
- 「恢复内置」按钮把内置词合并回当前列表；
- 保存时会持久化到 `~/.config/PatentMarker/vague_wordbank.json`。
"""
import json
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QMessageBox, QFileDialog, QLineEdit,
    QAbstractItemView,
)

from config_manager import (
    load_vague_wordbank, save_vague_wordbank, get_builtin_vague_wordbank,
)


class ClaimIgnoreDialog(QDialog):
    """不确定用语词库编辑器

    类名保留 `ClaimIgnoreDialog` 是为了兼容历史导入；其作用已变为
    「权利要求书中不应出现的不确定用语」词库的编辑。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("不确定用语忽略词库")
        self.setMinimumSize(520, 600)
        self.setModal(True)

        self._items: list = list(load_vague_wordbank())
        self._search_text: str = ""

        self._build_ui()
        self._rebuild_list()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        hint = QLabel(
            "此词库列出权利要求书中「不应出现」的不确定 / 含糊用语。\n"
            "「开始检查」时，这些词一旦出现在权利要求书中，就会在右侧结果表里以\n"
            "「不确定用语」类型报出。\n\n"
            "• 内置默认已包含常见词（约、大概、可能、优选、左右…）\n"
            "• 可自由增删 / 导入 / 导出；点「恢复内置」可把默认词合并回来\n"
            "• 保存后下次点「开始检查」即生效"
        )
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

        # 快速添加输入框
        add_row = QHBoxLayout()
        add_row.addWidget(QLabel("➕ 添加："))
        self.add_edit = QLineEdit()
        self.add_edit.setPlaceholderText("输入要检测的不确定用语后按回车或点「添加」")
        self.add_edit.returnPressed.connect(self._on_add_clicked)
        add_row.addWidget(self.add_edit, 1)
        add_btn = QPushButton("添加")
        add_btn.setObjectName("accentBtn")
        add_btn.clicked.connect(self._on_add_clicked)
        add_row.addWidget(add_btn)
        layout.addLayout(add_row)

        # 列表
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.list_widget.setAlternatingRowColors(True)
        layout.addWidget(self.list_widget, 1)

        # 底部按钮
        btn_row = QHBoxLayout()

        del_btn = QPushButton("🗑️ 删除选中")
        del_btn.clicked.connect(self._on_delete_clicked)
        btn_row.addWidget(del_btn)

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

    # ─────────────────────────────────────────

    def _rebuild_list(self):
        self.list_widget.clear()
        kw = self._search_text
        visible = 0
        for s in self._items:
            if kw and kw not in s.lower():
                continue
            item = QListWidgetItem(s)
            self.list_widget.addItem(item)
            visible += 1
        self._update_count(visible)

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
            QMessageBox.information(self, "已存在", f"「{text}」已在词库中。")
            self.add_edit.clear()
            return
        self._items.insert(0, text)
        self.add_edit.clear()
        if self._search_text:
            self.search_edit.clear()
        else:
            self._rebuild_list()

    def _on_delete_clicked(self):
        selected = self.list_widget.selectedItems()
        if not selected:
            QMessageBox.information(self, "提示", "请先选中要删除的条目")
            return
        to_remove = {item.text() for item in selected}
        self._items = [s for s in self._items if s not in to_remove]
        self._rebuild_list()

    def _on_restore_defaults(self):
        builtin = get_builtin_vague_wordbank()
        existing = set(self._items)
        added = 0
        for w in builtin:
            if w and w not in existing:
                self._items.append(w)
                existing.add(w)
                added += 1
        self._rebuild_list()
        if added:
            QMessageBox.information(
                self, "已合并", f"新增 {added} 条内置默认词。"
            )
        else:
            QMessageBox.information(
                self, "无变化", "当前列表已包含所有内置默认词。"
            )

    def _on_save(self):
        try:
            save_vague_wordbank(self._items)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"无法写入文件：\n{e}")
            return
        QMessageBox.information(
            self, "已保存",
            f"不确定用语词库已保存，共 {len(self._items)} 条。\n下次点「开始检查」即生效。"
        )
        self.accept()

    def _on_export(self):
        if not self._items:
            QMessageBox.information(self, "无内容", "当前词库为空。")
            return
        path, selected_filter = QFileDialog.getSaveFileName(
            self, "导出不确定用语词库", "vague_wordbank.json",
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
            self, "导入不确定用语词库", "",
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
