# -*- coding: utf-8 -*-
"""
claim_ignore_dialog.py — 不确定用语词库编辑器

维护权利要求书中"不应出现"的含糊 / 不确定用语（例如：约、大概、可能、
左右、优选…）。勾选「开始检查」时，会把这些词当作 `vague` 类问题报出。
"""
from config_manager import (
    load_vague_wordbank, save_vague_wordbank, get_builtin_vague_wordbank,
)
from dialogs.base_wordbank_dialog import BaseWordbankDialog


class ClaimIgnoreDialog(BaseWordbankDialog):
    """不确定用语词库编辑器。

    类名保留 `ClaimIgnoreDialog` 是为了兼容历史导入；其作用已变为
    「权利要求书中不应出现的不确定用语」词库的编辑。
    """

    TITLE = "不确定用语忽略词库"
    MIN_SIZE = (520, 600)
    HAS_RESTORE_DEFAULTS = True
    HINT_HTML = (
        "此词库列出权利要求书中「不应出现」的不确定 / 含糊用语。\n"
        "「开始检查」时，这些词一旦出现在权利要求书中，就会在右侧结果表里以\n"
        "「不确定用语」类型报出。\n\n"
        "• 内置默认已包含常见词（约、大概、可能、优选、左右…）\n"
        "• 可自由增删 / 导入 / 导出；点「恢复内置」可把默认词合并回来\n"
        "• 保存后下次点「开始检查」即生效"
    )
    ADD_PLACEHOLDER = "输入要检测的不确定用语后按回车或点「添加」"
    EXPORT_FILENAME = "vague_wordbank.json"
    EXPORT_TITLE = "导出不确定用语词库"
    IMPORT_TITLE = "导入不确定用语词库"
    SAVE_SUFFIX = "下次点「开始检查」即生效。"

    def load_items(self) -> list:
        return list(load_vague_wordbank())

    def save_items(self, items: list) -> None:
        save_vague_wordbank(items)

    def get_builtin(self) -> list:
        return get_builtin_vague_wordbank()
