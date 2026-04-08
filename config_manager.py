# -*- coding: utf-8 -*-
"""
config_manager.py — 软件配置与用户词库持久化管理
- 使用 QSettings 保存界面设置（主题、勾选状态、窗口大小、最近目录）
- 用户词库以 JSON 文件存储于用户配置目录，与内置词库合并
"""
import json
import os
from PyQt6.QtCore import QSettings, QStandardPaths

ORG_NAME = "PatentMarker"
APP_NAME = "MarkAssistant"


def get_config_dir() -> str:
    """返回用户级配置目录（自动创建）"""
    cfg_dir = QStandardPaths.writableLocation(
        QStandardPaths.StandardLocation.AppConfigLocation
    )
    if not cfg_dir:
        cfg_dir = os.path.join(os.path.expanduser("~"), ".patent_marker")
    os.makedirs(cfg_dir, exist_ok=True)
    return cfg_dir


def get_wordbank_path() -> str:
    """用户自定义错别字词库 JSON 路径"""
    return os.path.join(get_config_dir(), "user_wordbank.json")


def get_disabled_builtin_path() -> str:
    """用户禁用的内置词库条目 JSON 路径"""
    return os.path.join(get_config_dir(), "disabled_builtin_wordbank.json")


def get_dup_ignore_path() -> str:
    """重复字词检查的忽略词库 JSON 路径"""
    return os.path.join(get_config_dir(), "dup_ignore_wordbank.json")


# ─────────────────────────────────────────
# QSettings 封装：界面配置
# ─────────────────────────────────────────

class AppSettings:
    """简单的 QSettings 封装，提供类型安全的 get/set"""

    def __init__(self):
        self._s = QSettings(ORG_NAME, APP_NAME)

    # ---- 主题 ----
    def get_theme(self) -> str:
        return self._s.value("ui/theme", "light", type=str)

    def set_theme(self, theme: str):
        self._s.setValue("ui/theme", theme)

    # ---- 勾选状态 ----
    def get_bool(self, key: str, default: bool) -> bool:
        v = self._s.value(key, default)
        if isinstance(v, bool):
            return v
        if isinstance(v, str):
            return v.lower() in ("true", "1", "yes")
        return bool(v)

    def set_bool(self, key: str, value: bool):
        self._s.setValue(key, bool(value))

    # ---- 窗口几何 ----
    def get_geometry(self):
        return self._s.value("ui/geometry")

    def set_geometry(self, geom):
        self._s.setValue("ui/geometry", geom)

    # ---- 最近打开目录 ----
    def get_last_dir(self) -> str:
        return self._s.value("io/last_dir", "", type=str)

    def set_last_dir(self, path: str):
        self._s.setValue("io/last_dir", path)

    def sync(self):
        self._s.sync()


# ─────────────────────────────────────────
# 用户词库：JSON 持久化
# ─────────────────────────────────────────

def load_user_wordbank() -> list:
    """读取用户自定义错别字词库 JSON。文件不存在时返回空列表。"""
    path = get_wordbank_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            # 校验每项的格式
            return [
                {"wrong": str(x.get("wrong", "")), "suggestion": str(x.get("suggestion", ""))}
                for x in data
                if isinstance(x, dict) and x.get("wrong")
            ]
    except Exception:
        return []
    return []


def save_user_wordbank(entries: list):
    """保存用户词库到 JSON 文件"""
    path = get_wordbank_path()
    cleaned = [
        {"wrong": e["wrong"], "suggestion": e["suggestion"]}
        for e in entries
        if e.get("wrong") and e.get("suggestion")
    ]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)


def load_disabled_builtin_wrongs() -> set:
    """读取被用户禁用的内置词库 wrong 列表（返回 set）"""
    path = get_disabled_builtin_path()
    if not os.path.exists(path):
        return set()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return {str(x) for x in data if x}
    except Exception:
        return set()
    return set()


def save_disabled_builtin_wrongs(wrongs) -> None:
    """保存被用户禁用的内置词库 wrong 列表"""
    path = get_disabled_builtin_path()
    cleaned = sorted({str(w) for w in wrongs if w})
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)


def load_dup_ignore_list() -> list:
    """读取用户自定义的「重复字词忽略词库」（返回去重后的字符串列表）"""
    path = get_dup_ignore_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            seen = set()
            result = []
            for x in data:
                s = str(x).strip()
                if s and s not in seen:
                    seen.add(s)
                    result.append(s)
            return result
    except Exception:
        return []
    return []


def save_dup_ignore_list(items) -> None:
    """保存重复字词忽略词库"""
    path = get_dup_ignore_path()
    cleaned = []
    seen = set()
    for x in items:
        s = str(x).strip()
        if s and s not in seen:
            seen.add(s)
            cleaned.append(s)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)


def get_merged_wordbank() -> list:
    """合并内置词库与用户词库：
    - 剔除用户禁用的内置条目
    - 用户条目优先（若 wrong 与内置相同，则覆盖内置）
    """
    from typo_wordbank import WORDBANK as BUILTIN
    user = load_user_wordbank()
    user_wrongs = {e["wrong"] for e in user}
    disabled = load_disabled_builtin_wrongs()
    merged = [
        e for e in BUILTIN
        if e["wrong"] not in user_wrongs and e["wrong"] not in disabled
    ] + user
    return merged
