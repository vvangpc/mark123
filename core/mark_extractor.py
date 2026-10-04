# -*- coding: utf-8 -*-
"""
mark_extractor.py — 标记提取器
从附图说明段落中提取标记定义（数字-名称映射）。
"""
import re


_SEP = r"[-\-—–、.．,:：\s]"
_NAME_CH = r"[\u4e00-\u9fa5a-zA-Z]"
_ENTRY_RE = re.compile(
    r"(?<!\d)(\d+(?:\s*[、,，和及]\s*\d+)*)"   # 编号（可多个）
    + r"\s*" + _SEP + r"\s*"
    + r"(" + _NAME_CH + r"(?:" + _NAME_CH + r"|\d+(?!\s*[-—–、.．,:：]|\s+\S))*)"
)
_NUM_RE = re.compile(r"\d+")


def extract_marks_from_text(text: str) -> dict[int, str]:
    """
    从附图标记文本中提取标记字典。

    支持多种常见格式：
    - "1-齿圈，2-夹指，3-转盘"
    - "1、齿圈；2、夹指；3、转盘"
    - "齿圈1，夹指2，转盘3"
    - 混合格式

    参数:
        text: 附图标记文本，如 "附图标记：1-齿圈，2-夹指，3-转盘..."

    返回:
        {数字: 名称} 如 {1: '齿圈', 2: '夹指', 3: '转盘'}
    """
    if not text or not text.strip():
        return {}

    marks = {}

    # 如果文本包含"附图标记"前缀，去掉
    text = re.sub(r'^.*?附图标记\s*[:：]\s*', '', text.strip())

    # 策略1: "编号+分隔符+名称" 模式
    # 匹配: 1-齿圈, 1、齿圈, 1.齿圈, 1 齿圈；多个编号共用一个名称（11、12-侧板）；
    # 名称内可含数字（2-第3连杆），但「数字 + 分隔符」视为下一条的开头而不吞进名称
    for m in _ENTRY_RE.finditer(text):
        name = m.group(2).strip()
        if not name:
            continue
        for num_s in _NUM_RE.findall(m.group(1)):
            num = int(num_s)
            if num not in marks:
                marks[num] = name

    # 策略2: "名称+数字" 模式（备用，如果策略1没有结果）
    if not marks:
        pattern2 = r'([\u4e00-\u9fa5a-zA-Z]{2,20})(\d+)'
        for m in re.finditer(pattern2, text):
            name = m.group(1).strip()
            num = int(m.group(2))
            # 过滤掉介词前缀
            name = re.sub(r'^(所述|该|此|及|和|与|的|以及|图|包括|连接)+', '', name)
            if name and len(name) >= 1:
                if num not in marks:
                    marks[num] = name

    return marks


def extract_marks_from_paragraph(paragraph) -> dict[int, str]:
    """
    从python-docx的Paragraph对象中提取标记。

    参数:
        paragraph: python-docx Paragraph对象

    返回:
        {数字: 名称}
    """
    if paragraph is None:
        return {}
    return extract_marks_from_text(paragraph.text)


def extract_marks_from_paragraphs(paragraphs) -> dict[int, str]:
    """
    从多个 Paragraph 对象中提取标记（合并文本后统一解析）。

    参数:
        paragraphs: python-docx Paragraph 对象列表

    返回:
        {数字: 名称}
    """
    if not paragraphs:
        return {}
    combined = "；".join(p.text for p in paragraphs if p.text)
    return extract_marks_from_text(combined)


def marks_to_display_text(marks: dict[int, str]) -> str:
    """
    将标记字典转为显示文本。

    参数:
        marks: {数字: 名称}

    返回:
        如 "1-齿圈，2-夹指，3-转盘"
    """
    if not marks:
        return ""
    sorted_nums = sorted(marks.keys())
    parts = [f"{num}-{marks[num]}" for num in sorted_nums]
    return "，".join(parts)


def parse_marks_from_display_text(text: str) -> dict[int, str]:
    """
    从用户编辑后的显示文本中重新解析标记字典。
    支持格式: "1-齿圈，2-夹指" 或 "1、齿圈；2、夹指" 等

    参数:
        text: 用户编辑后的标记文本

    返回:
        {数字: 名称}
    """
    return extract_marks_from_text(text)


# ── 忽略附图标记 ──────────────────────────────────────────
# 部件名后的括号编号：齿圈（1）/ 连接件 (3a) / 侧板（11、12）。前一个字必须是汉字或字母，
# 段首的「（1）」步骤号、「权利要求1(2)」这类不动。
_PAREN_MARK_RE = re.compile(
    r'(?<=[' + _NAME_CH[1:-1] + r'])\s*[（(]\s*\d+[A-Za-z]*'
    r'(?:\s*[、,，\-－~～]\s*\d+[A-Za-z]*)*\s*[）)]'
)


def strip_reference_marks(text: str, marks: dict = None) -> tuple:
    """去掉文本里的附图标记，返回 (去标记文本, 位置映射)。

    映射 idx[k] 为去标记文本第 k 个字在原文中的下标，末尾另有一个 len(text) 哨兵。
    除括号编号外，marks（{编号: 名称}）里「名称+编号」的无括号写法（齿圈1）也去掉编号。
    """
    spans = [m.span() for m in _PAREN_MARK_RE.finditer(text)]
    for num, name in (marks or {}).items():
        if not name:
            continue
        num_s = str(num)
        for m in re.finditer(re.escape(name) + num_s + r'(?!\d)', text):
            spans.append((m.end() - len(num_s), m.end()))
    if not spans:
        return text, list(range(len(text) + 1))
    drop = [False] * len(text)
    for a, b in spans:
        for k in range(a, b):
            drop[k] = True
    keep = [k for k in range(len(text)) if not drop[k]]
    return "".join(text[k] for k in keep), keep + [len(text)]
