# -*- coding: utf-8 -*-
"""
boundary_blacklist_dialog.py — 引用基础检查「动态截断黑名单」词库编辑器

「动态截断」从『所述』后向后扫描 CJK 字符，遇到黑名单词的首字立即停下，
把之前累积的字串当作术语提取出来。例如：
    「所述齿轮安装在主轴上」 → 命中 "安装" → 术语 = "齿轮"
"""
from config_manager import (
    load_boundary_blacklist, save_boundary_blacklist,
    get_builtin_boundary_blacklist,
)
from dialogs.base_wordbank_dialog import BaseWordbankDialog


class BoundaryBlacklistDialog(BaseWordbankDialog):
    """动态截断黑名单词库编辑器"""

    TITLE = "动态截断黑名单词库"
    MIN_SIZE = (560, 620)
    HAS_RESTORE_DEFAULTS = True
    HINT_HTML = (
        "此词库列出引用基础检查「动态截断」模式所用的边界词。\n"
        "提取『所述X』中的 X 时，从「所述」往后扫，遇到任一黑名单词的首字\n"
        "就立即停下，把之前累积的字串作为术语。\n\n"
        "例：「所述齿轮安装在主轴上」 → 命中『安装』→ 术语 = 齿轮\n\n"
        "• 内置默认包含动词类（安装/连接/设置…）、方位词（上/下/内/外/端…）\n"
        "  以及连词类（与/和/或…）共 100+ 条\n"
        "• 可自由增删 / 导入 / 导出；点「恢复内置」可把默认词合并回来\n"
        "• 保存后下次点「开始检查」即生效"
    )
    ADD_PLACEHOLDER = "输入要加入黑名单的词后按回车或点「添加」"
    EXPORT_FILENAME = "boundary_blacklist.json"
    EXPORT_TITLE = "导出动态截断黑名单"
    IMPORT_TITLE = "导入动态截断黑名单"
    SAVE_SUFFIX = "下次点「开始检查」即生效。"

    def load_items(self) -> list:
        return list(load_boundary_blacklist())

    def save_items(self, items: list) -> None:
        save_boundary_blacklist(items)

    def get_builtin(self) -> list:
        return get_builtin_boundary_blacklist()
