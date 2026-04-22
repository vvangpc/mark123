# -*- coding: utf-8 -*-
"""
annotator.py — 格式安全标注引擎
在Word段落中执行文本替换，通过直接操作底层的 <w:t> XML 标签，
确保完整保留所有对象（如 MathType/AXmath 公式节点、图形等）以及 run 级别的格式。
"""
import re
from docx.oxml.ns import qn


def _build_xml_char_map(paragraph):
    """
    构建段落中基于底层 <w:t> (text) 元素的字符映射表。
    
    返回:
        full_text: 段落内的纯文本拼接
        char_info: [(wt_element, char_offset_in_wt), ...]
    """
    full_text = ""
    char_info = []

    for run in paragraph.runs:
        # 获取基础 XML 元素 w:r
        r_elem = run._r
        # 仅查找该 run 内的所有 w:t 节点（跳过 w:object, m:oMath 等公式图表）
        wt_elements = r_elem.xpath('.//w:t')
        
        for wt in wt_elements:
            text = wt.text
            if text:
                for idx, char in enumerate(text):
                    full_text += char
                    char_info.append((wt, idx))

    return full_text, char_info


def annotate_paragraph_safe(paragraph, replace_dict: dict[str, str],
                            preamble_regex: re.Pattern = None) -> bool:
    """
    安全版本的段落标注函数，直接操作 <w:t> 以保留所有非文本子节点。

    参数:
        paragraph: python-docx Paragraph对象
        replace_dict: {原文: 替换后文本}
        preamble_regex: 若提供且在段落内命中，则其 end() 之前的匹配视为前言部分
            （如权利要求"其特征在于"之前的"一种X的装置"），不进行替换。

    返回:
        是否发生了替换
    """
    if not paragraph.runs or not replace_dict:
        return False

    # 步骤1: 获取仅基于 w:t 的内容映射
    full_text, char_info = _build_xml_char_map(paragraph)
    if not full_text:
        return False

    # 计算"前言保护区"长度：权利要求书首段中"其特征在于"之前的专利名称部分
    protected_prefix_len = 0
    if preamble_regex is not None:
        m_pre = preamble_regex.search(full_text)
        if m_pre:
            protected_prefix_len = m_pre.end()

    # 步骤2: 排序key（长的优先）
    sorted_keys = sorted(replace_dict.keys(), key=len, reverse=True)
    escaped_keys = [re.escape(k) for k in sorted_keys if k]
    if not escaped_keys:
        return False

    pattern = '|'.join(escaped_keys)
    compiled = re.compile(pattern)

    # 查找所有匹配；前言部分（专利名称）不进行替换
    matches = [m for m in compiled.finditer(full_text)
               if m.start() >= protected_prefix_len]
    if not matches:
        return False

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

    # 步骤4: 重新向每个 <w:t> 分配新文本
    # 提取所有使用到的 wt 节点去重，防止有些节点没清空残余字符
    involved_wts = list(dict.fromkeys(wt for wt, _ in char_info))
    
    # 准备每个 wt 的新写入缓冲区
    wt_buffers = {wt: [] for wt in involved_wts}

    for seg_old_start, seg_old_end, seg_text, is_replaced in segments:
        if not is_replaced:
            # 未替换的部分，按字面字符原位放回
            for i in range(len(seg_text)):
                old_pos = seg_old_start + i
                wt, _ = char_info[old_pos]
                wt_buffers[wt].append(seg_text[i])
        else:
            # 替换的部分：把这部分的新文本全挂在匹配字串开始时的那个 wt 节点下
            wt, _ = char_info[seg_old_start]
            wt_buffers[wt].append(seg_text)

    # 将缓冲区内容写回 wt.text
    for wt in involved_wts:
        wt.text = "".join(wt_buffers[wt])

    return True


# ---------------- 替换与生成规则 ----------------

def build_claims_replace_dict(marks: dict[int, str]) -> dict[str, str]:
    """权利要求书：齿圈 → 齿圈（1）"""
    replace_dict = {}
    for num, name in marks.items():
        replace_dict[name] = f"{name}（{num}）"
    return replace_dict

def build_implementation_replace_dict(marks: dict[int, str]) -> dict[str, str]:
    """具体实施方式：齿圈 → 齿圈1"""
    replace_dict = {}
    for num, name in marks.items():
        replace_dict[name] = f"{name}{num}"
    return replace_dict

def build_claims_remove_dict(marks: dict[int, str]) -> dict[str, str]:
    """权利要求书清洗：齿圈（1）/ 齿圈(1) → 齿圈"""
    remove_dict = {}
    for num, name in marks.items():
        remove_dict[f"{name}（{num}）"] = name
        remove_dict[f"{name}({num})"] = name
    return remove_dict

def build_implementation_remove_dict(marks: dict[int, str]) -> dict[str, str]:
    """具体实施方式清洗：齿圈1 → 齿圈"""
    remove_dict = {}
    for num, name in marks.items():
        remove_dict[f"{name}{num}"] = name
    return remove_dict


# 权利要求书前言识别：命中后其 end() 之前视为"专利名称"部分，不参与标注
# 兼容"其特征在于"与"其特征是"，允许后跟中/英文标点
_CLAIMS_PREAMBLE_RE = re.compile(r'其特征(?:在于|是)[，：:,]?')


# ---------------- 检查 ----------------

def _is_paragraph_already_annotated(text: str, marks: dict[int, str], mode: str) -> bool:
    """检查是否已经标注过 (简化基于正则)。"""
    for num, name in marks.items():
        if name not in text:
            continue
        if mode == "claims":
            if re.search(re.escape(name) + r'[（(]\s*' + str(num) + r'\s*[）)]', text):
                return True
        else:
            if re.escape(name) + str(num) in text:
                return True
    return False


# ---------------- 主干操作流程 ----------------

def annotate_section(paragraphs, section, replace_dict: dict[str, str],
                     skip_already_annotated: bool = True,
                     marks: dict[int, str] = None,
                     mode: str = "claims") -> int:
    """对指定章节的所有段落执行替换操作（增加标记）"""
    replaced_count = 0
    # 权利要求书中首段（独立权项）常含"一种X，其特征在于，..."前言，
    # X 里若包含已标注部件名，应避免被错误标注
    preamble_re = _CLAIMS_PREAMBLE_RE if mode == "claims" else None
    for i in range(section.start_idx, section.end_idx):
        para = paragraphs[i]
        text = para.text.strip()
        if not text:
            continue

        if skip_already_annotated and marks:
            if _is_paragraph_already_annotated(text, marks, mode):
                continue

        if annotate_paragraph_safe(para, replace_dict, preamble_regex=preamble_re):
            replaced_count += 1
    return replaced_count


def remove_section_marks(paragraphs, section, remove_dict: dict[str, str]) -> int:
    """对指定章节清除标记"""
    removed_count = 0
    for i in range(section.start_idx, section.end_idx):
        para = paragraphs[i]
        if not para.text.strip():
            continue
        
        # 使用相同的清洗引擎安全替换
        if annotate_paragraph_safe(para, remove_dict):
            removed_count += 1
    return removed_count


def smart_annotate_section(paragraphs, section, marks: dict[int, str],
                           mode: str = "claims") -> int:
    """智能增加标注"""
    replace_dict = build_claims_replace_dict(marks) if mode == "claims" else build_implementation_replace_dict(marks)
    
    annotated_para_count = 0
    para_with_marks_count = 0

    for i in range(section.start_idx, section.end_idx):
        text = paragraphs[i].text.strip()
        if not text:
            continue
        has_any_mark = any(name in text for name in marks.values())
        if not has_any_mark:
            continue
        para_with_marks_count += 1
        if _is_paragraph_already_annotated(text, marks, mode):
            annotated_para_count += 1

    if para_with_marks_count > 0 and annotated_para_count / para_with_marks_count > 0.7:
        return -1
        
    return annotate_section(paragraphs, section, replace_dict, True, marks, mode)

def smart_remove_section(paragraphs, section, marks: dict[int, str], mode: str = "claims") -> int:
    """智能删除标注"""
    remove_dict = build_claims_remove_dict(marks) if mode == "claims" else build_implementation_remove_dict(marks)
    return remove_section_marks(paragraphs, section, remove_dict)


def update_mark_paragraph_text(mark_para, new_text: str) -> bool:
    """
    将附图标记段落的文本替换为 new_text（保留段落格式，丢弃多余 run）。
    用于用户编辑词典后同步回文档段落。
    返回是否修改成功。
    """
    if mark_para is None:
        return False
    runs = list(mark_para.runs)
    if not runs:
        mark_para.add_run(new_text)
        return True
    # 第一个 run 写入新内容
    runs[0].text = new_text
    # 后续 run 清空（保留 run 节点以保留任何指向它们的引用）
    for r in runs[1:]:
        r.text = ""
    return True
