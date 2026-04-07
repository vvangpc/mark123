# -*- coding: utf-8 -*-
"""
annotator.py — 格式安全标注引擎
在Word段落中执行文本替换，同时完整保留所有run级别的格式（字体、加粗、颜色、字号等）。
"""
import re
import copy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _copy_run_format(source_run, target_run):
    """
    将源run的所有格式属性复制到目标run。
    直接复制xml属性以确保完整保留所有格式。
    """
    # 复制rPr（run properties）元素
    source_rpr = source_run._element.find(qn('w:rPr'))
    if source_rpr is not None:
        # 先移除目标的rPr
        target_rpr = target_run._element.find(qn('w:rPr'))
        if target_rpr is not None:
            target_run._element.remove(target_rpr)
        # 深拷贝源的rPr并添加到目标
        new_rpr = copy.deepcopy(source_rpr)
        target_run._element.insert(0, new_rpr)


def _build_char_map(paragraph):
    """
    构建段落中每个字符到run索引和字符偏移的映射。

    返回:
        full_text: 段落完整文本
        char_info: [(run_index, char_offset_in_run), ...] 每个字符的来源信息
    """
    full_text = ""
    char_info = []

    for run_idx, run in enumerate(paragraph.runs):
        for char_offset, char in enumerate(run.text):
            full_text += char
            char_info.append((run_idx, char_offset))

    return full_text, char_info


def annotate_paragraph(paragraph, replace_dict: dict[str, str]) -> bool:
    """
    在段落中执行文本替换，完整保留所有格式。

    核心算法：
    1. 拼接所有run文本得到完整段落文本
    2. 在完整文本上执行替换，记录替换操作
    3. 重新计算每个run应该包含的文本
    4. 仅修改run.text，保留所有格式属性

    参数:
        paragraph: python-docx Paragraph对象
        replace_dict: {原文: 替换后文本}

    返回:
        是否发生了替换
    """
    if not paragraph.runs or not replace_dict:
        return False

    # 步骤1: 获取完整段落文本
    full_text, char_info = _build_char_map(paragraph)

    if not full_text.strip():
        return False

    # 步骤2: 在完整文本上计算所有替换操作
    # 按长度降序排列key，确保长字符串优先匹配
    sorted_keys = sorted(replace_dict.keys(), key=len, reverse=True)

    # 转义正则特殊字符并构建模式
    escaped_keys = [re.escape(k) for k in sorted_keys if k]
    if not escaped_keys:
        return False

    pattern = '|'.join(escaped_keys)
    compiled = re.compile(pattern)

    # 检查是否有匹配
    if not compiled.search(full_text):
        return False

    # 执行替换，获取新文本
    new_text = compiled.sub(lambda m: replace_dict[m.group(0)], full_text)

    if new_text == full_text:
        return False

    # 步骤3: 将新文本按原run边界重新分配
    # 使用差异对比算法来精确定位替换位置
    _redistribute_text(paragraph, full_text, new_text, char_info)

    return True


def _redistribute_text(paragraph, old_text: str, new_text: str, char_info: list):
    """
    将替换后的新文本重新分配到各个run中，保留原始格式。

    策略：
    - 计算每个替换的位置和偏移量
    - 对于跨run的替换，将新文本放入第一个涉及的run
    - 保持未修改部分的run归属不变
    """
    runs = paragraph.runs
    if not runs:
        return

    # 收集每个run的文本起止位置
    run_boundaries = []
    pos = 0
    for run in runs:
        length = len(run.text)
        run_boundaries.append((pos, pos + length))
        pos += length

    # 使用简洁的方法：
    # 找出所有替换操作的位置，然后逐个run计算新文本
    sorted_keys = sorted(
        [k for k in _current_replace_dict.keys() if k],
        key=len, reverse=True
    )
    escaped_keys = [re.escape(k) for k in sorted_keys]
    pattern = '|'.join(escaped_keys)
    compiled = re.compile(pattern)

    # 找到所有匹配的位置
    replacements = []  # [(old_start, old_end, new_substring), ...]
    for m in compiled.finditer(old_text):
        old_start = m.start()
        old_end = m.end()
        new_substr = _current_replace_dict[m.group(0)]
        replacements.append((old_start, old_end, new_substr))

    if not replacements:
        return

    # 从后往前处理替换（避免位置偏移）
    # 构建位置映射: old_pos -> new_pos
    # 直接对每个run的文本进行修改

    # 转换为分段列表：[原始文本段, ...]
    segments = []  # [(old_start, old_end, new_text, is_replaced), ...]
    prev_end = 0
    for old_start, old_end, new_substr in replacements:
        if prev_end < old_start:
            segments.append((prev_end, old_start, old_text[prev_end:old_start], False))
        segments.append((old_start, old_end, new_substr, True))
        prev_end = old_end
    if prev_end < len(old_text):
        segments.append((prev_end, len(old_text), old_text[prev_end:], False))

    # 为每个run分配新文本
    new_run_texts = [""] * len(runs)

    for seg_old_start, seg_old_end, seg_text, is_replaced in segments:
        if not is_replaced:
            # 未替换的文本段，保持原归属
            for i, char in enumerate(seg_text):
                old_pos = seg_old_start + i
                if old_pos < len(char_info):
                    run_idx, _ = char_info[old_pos]
                    new_run_texts[run_idx] += char
        else:
            # 替换的文本段，将新文本放入起始run
            if seg_old_start < len(char_info):
                run_idx, _ = char_info[seg_old_start]
                new_run_texts[run_idx] += seg_text

    # 更新每个run的文本
    for i, run in enumerate(runs):
        run.text = new_run_texts[i]


# 模块级变量，用于在闭包中传递替换字典
_current_replace_dict = {}


def annotate_paragraph_safe(paragraph, replace_dict: dict[str, str]) -> bool:
    """
    安全版本的段落标注函数，使用更可靠的替换策略。

    参数:
        paragraph: python-docx Paragraph对象
        replace_dict: {原文: 替换后文本}

    返回:
        是否发生了替换
    """
    global _current_replace_dict

    if not paragraph.runs or not replace_dict:
        return False

    # 步骤1: 获取完整段落文本
    full_text, char_info = _build_char_map(paragraph)
    if not full_text.strip():
        return False

    # 步骤2: 排序key（长的优先）
    sorted_keys = sorted(replace_dict.keys(), key=len, reverse=True)
    escaped_keys = [re.escape(k) for k in sorted_keys if k]
    if not escaped_keys:
        return False

    pattern = '|'.join(escaped_keys)
    compiled = re.compile(pattern)

    # 查找所有匹配
    matches = list(compiled.finditer(full_text))
    if not matches:
        return False

    # 设置全局替换字典
    _current_replace_dict = replace_dict

    # 步骤3: 构建未替换/已替换文本段
    segments = []
    prev_end = 0
    for m in matches:
        old_start = m.start()
        old_end = m.end()
        new_substr = replace_dict[m.group(0)]
        if prev_end < old_start:
            segments.append((prev_end, old_start, full_text[prev_end:old_start], False))
        segments.append((old_start, old_end, new_substr, True))
        prev_end = old_end
    if prev_end < len(full_text):
        segments.append((prev_end, len(full_text), full_text[prev_end:], False))

    # 步骤4: 为每个run分配新文本
    runs = paragraph.runs
    new_run_texts = [""] * len(runs)

    for seg_old_start, seg_old_end, seg_text, is_replaced in segments:
        if not is_replaced:
            # 未替换的部分，保持原run归属
            for i in range(len(seg_text)):
                old_pos = seg_old_start + i
                if old_pos < len(char_info):
                    run_idx, _ = char_info[old_pos]
                    new_run_texts[run_idx] += seg_text[i]
        else:
            # 替换的部分，新文本放入匹配起始位置所在的run
            if seg_old_start < len(char_info):
                run_idx, _ = char_info[seg_old_start]
                new_run_texts[run_idx] += seg_text

    # 步骤5: 仅修改run.text（格式完整保留）
    for i, run in enumerate(runs):
        run.text = new_run_texts[i]

    return True


def build_claims_replace_dict(marks: dict[int, str]) -> dict[str, str]:
    """
    构建权利要求书的替换字典。
    权利要求书中，部件名称无标注 → 添加全角括号编号。
    如：齿圈 → 齿圈（1）

    注意：需要避免重复标注（如果已有标注则跳过）。

    参数:
        marks: {数字: 名称}

    返回:
        {原文: 替换后文本}
    """
    replace_dict = {}
    for num, name in marks.items():
        # 将纯名称替换为 名称（数字）
        replacement = f"{name}（{num}）"
        replace_dict[name] = replacement
    return replace_dict


def build_implementation_replace_dict(marks: dict[int, str]) -> dict[str, str]:
    """
    构建具体实施方式的替换字典。
    具体实施方式中，部件名称无标注 → 添加数字编号。
    如：齿圈 → 齿圈1

    参数:
        marks: {数字: 名称}

    返回:
        {原文: 替换后文本}
    """
    replace_dict = {}
    for num, name in marks.items():
        replacement = f"{name}{num}"
        replace_dict[name] = replacement
    return replace_dict


def _is_paragraph_already_annotated(text: str, replace_dict: dict[str, str],
                                     marks: dict[int, str], mode: str) -> bool:
    """
    判断段落是否已经包含标注。

    检测策略：
    1. 替换后的完整文本已经存在于段落中
    2. 对于权利要求书：名称后紧跟全角或半角括号+数字
    3. 对于具体实施方式：名称后紧跟数字

    参数:
        text: 段落文本
        replace_dict: 替换字典
        marks: 标记字典
        mode: "claims" 或 "implementation"
    """
    # 检查方式1: 替换后文本已存在
    for old_text, new_text in replace_dict.items():
        if new_text in text:
            return True

    # 检查方式2: 按模式用正则检测已有标注
    for num, name in marks.items():
        if name not in text:
            continue

        if mode == "claims":
            # 权利要求书：检查 名称（数字） 或 名称(数字) 模式
            pattern = re.escape(name) + r'[（(]\s*' + str(num) + r'\s*[）)]'
            if re.search(pattern, text):
                return True
        else:
            # 具体实施方式：检查 名称+数字 模式（数字紧跟名称）
            pattern = re.escape(name) + str(num)
            if pattern in text:
                return True

    return False


def annotate_section(paragraphs, section, replace_dict: dict[str, str],
                     skip_already_annotated: bool = True,
                     marks: dict[int, str] = None,
                     mode: str = "claims") -> int:
    """
    对指定章节中的所有段落执行标注。

    参数:
        paragraphs: 文档所有段落列表
        section: DocSection对象
        replace_dict: 替换字典
        skip_already_annotated: 是否跳过已有标注的段落
        marks: 标记字典 {数字: 名称}，用于辅助检测已标注
        mode: "claims" 或 "implementation"

    返回:
        成功替换的段落数量
    """
    replaced_count = 0

    for i in range(section.start_idx, section.end_idx):
        para = paragraphs[i]
        text = para.text.strip()
        if not text:
            continue

        if skip_already_annotated and marks:
            if _is_paragraph_already_annotated(text, replace_dict, marks, mode):
                continue

        if annotate_paragraph_safe(para, replace_dict):
            replaced_count += 1

    return replaced_count


def smart_annotate_section(paragraphs, section, marks: dict[int, str],
                           mode: str = "claims") -> int:
    """
    智能标注章节 - 自动检测已有标注并避免重复。

    参数:
        paragraphs: 文档所有段落列表
        section: DocSection对象
        marks: {数字: 名称} 标记字典
        mode: "claims" 权利要求书模式 或 "implementation" 具体实施方式模式

    返回:
        成功替换的段落数量，-1表示已经全部标注过
    """
    if mode == "claims":
        replace_dict = build_claims_replace_dict(marks)
    else:
        replace_dict = build_implementation_replace_dict(marks)

    # 采样整个章节（而非仅前10段）来判断是否已有标注
    annotated_para_count = 0
    para_with_marks_count = 0

    for i in range(section.start_idx, section.end_idx):
        text = paragraphs[i].text.strip()
        if not text:
            continue

        # 检查该段落是否包含任何标记名称
        has_any_mark = any(name in text for name in marks.values())
        if not has_any_mark:
            continue

        para_with_marks_count += 1

        # 检查是否已标注
        if _is_paragraph_already_annotated(text, replace_dict, marks, mode):
            annotated_para_count += 1

    # 如果包含标记的段落中，超过70%已经标注，视为已全部标注
    if para_with_marks_count > 0 and annotated_para_count / para_with_marks_count > 0.7:
        return -1  # 返回-1表示已经标注过

    return annotate_section(paragraphs, section, replace_dict,
                           skip_already_annotated=True,
                           marks=marks, mode=mode)
