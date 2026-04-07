# -*- coding: utf-8 -*-
"""
mark_extractor.py — 标记提取器
从附图说明段落中提取标记定义（数字-名称映射）。
"""
import re


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

    # 策略1: "数字+分隔符+名称" 模式
    # 匹配: 1-齿圈, 1、齿圈, 1.齿圈, 1 齿圈 等
    pattern1 = r'(\d+)\s*[-\-—–、.．,:：\s]\s*([\u4e00-\u9fa5a-zA-Z][\u4e00-\u9fa5a-zA-Z]*)'
    for m in re.finditer(pattern1, text):
        num = int(m.group(1))
        name = m.group(2).strip()
        if name and len(name) >= 1:
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
