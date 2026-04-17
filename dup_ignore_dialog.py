# -*- coding: utf-8 -*-
"""
dup_ignore_dialog.py — 重复字词忽略词库编辑器
- 用户添加到此列表中的字 / 词，将在「重复字词检查」中被忽略
- 支持搜索 / 添加 / 删除 / 导入 / 导出
"""
from config_manager import load_dup_ignore_list, save_dup_ignore_list
from dialogs.base_wordbank_dialog import BaseWordbankDialog


class DupIgnoreDialog(BaseWordbankDialog):
    """重复字词忽略词库编辑器"""

    TITLE = "重复字词忽略词库"
    MIN_SIZE = (480, 560)
    HAS_RESTORE_DEFAULTS = False
    HINT_HTML = (
        "添加到此处的字 / 词，将在「重复字词检查」中被忽略。\n"
        "例如：将「所述」加入忽略后，「所述所述」不再被标红。\n"
        "• 支持按单字 / 词组匹配（同时匹配「重复单元」和「完整重复串」）\n"
        "• 保存后下次点「重复字词检查」即生效"
    )
    ADD_PLACEHOLDER = "输入要忽略的字 / 词后按回车或点「添加」"
    EXPORT_FILENAME = "dup_ignore.json"
    EXPORT_TITLE = "导出忽略词库"
    IMPORT_TITLE = "导入忽略词库"
    SAVE_SUFFIX = "下次点「重复字词检查」即生效。"
    DUPLICATE_ITEM_HINT = "「{}」已在忽略列表中。"
    EMPTY_EXPORT_HINT = "当前忽略词库为空。"

    def load_items(self) -> list:
        return list(load_dup_ignore_list())

    def save_items(self, items: list) -> None:
        save_dup_ignore_list(items)
