# -*- coding: utf-8 -*-
"""
doc_parser.py — docx 五书分段器
负责读取docx文件，按标题关键词识别专利五书各部分的段落范围。
"""
import re
from docx import Document


class DocSection:
    """表示文档中的一个章节"""
    def __init__(self, name: str, start_idx: int, end_idx: int):
        self.name = name
        self.start_idx = start_idx  # 起始段落索引（包含标题段落）
        self.end_idx = end_idx      # 结束段落索引（不包含下一节标题）

    def __repr__(self):
        return f"DocSection(name='{self.name}', start={self.start_idx}, end={self.end_idx})"


# 五书章节标题关键词（按文档中出现的顺序排列）
# 注意：这里只列出我们关心的章节
SECTION_KEYWORDS = [
    ("权利要求书", ["权利要求书"]),
    ("技术领域", ["技术领域"]),
    ("背景技术", ["背景技术"]),
    ("发明内容", ["发明内容"]),
    ("附图说明", ["附图说明"]),
    ("具体实施方式", ["具体实施方式"]),
    ("说明书附图", ["说明书附图"]),
    ("说明书摘要", ["说明书摘要", "摘要"]),
    ("摘要附图", ["摘要附图"]),
]


def is_section_title(paragraph, keywords: list[str]) -> bool:
    """
    判断一个段落是否是指定的章节标题。
    判断条件：
    1. 段落文字去除空白后与关键词完全匹配
    2. 或者段落文字以关键词开头且为加粗样式
    """
    text = paragraph.text.strip()
    if not text:
        return False

    for kw in keywords:
        # 精确匹配
        if text == kw:
            return True
        # 加粗标题匹配（有些标题可能带前缀）
        if text == kw and paragraph.runs:
            first_run = paragraph.runs[0]
            if first_run.font.bold:
                return True

    return False


def parse_document(doc_path: str) -> dict:
    """
    解析docx文档，识别五书各章节的段落范围。

    参数:
        doc_path: docx文件路径

    返回:
        {
            'document': Document对象,
            'sections': {章节名: DocSection对象},
            'claims_paras': 权利要求书段落列表,
            'implementation_paras': 具体实施方式段落列表,
            'mark_description_para': 附图标记所在段落,
        }
    """
    doc = Document(doc_path)
    paragraphs = doc.paragraphs

    # 第一步：找到所有章节标题的位置
    title_positions = []  # [(段落索引, 章节名)]

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue
        for section_name, keywords in SECTION_KEYWORDS:
            if is_section_title(para, keywords):
                title_positions.append((i, section_name))
                break

    # 如果没有找到明确的"权利要求书"标题，需要通过内容特征来识别
    # 权利要求书的特征：段落以 "1." 或 "1、" 开头，且包含"其特征在于"
    section_names_found = [name for _, name in title_positions]

    if "权利要求书" not in section_names_found:
        # 尝试通过内容特征定位权利要求书
        for i, para in enumerate(paragraphs):
            text = para.text.strip()
            if re.match(r'^1[.、．]', text) and '其特征在于' in text:
                # 找到了权利要求书的第一项，向前查找标题
                # 权利要求书可能没有显式标题（隐含在分页符中）
                title_positions.append((i, "权利要求书"))
                break

    # 第二步：根据标题位置划分段落区间
    # 按段落索引排序
    title_positions.sort(key=lambda x: x[0])

    sections = {}
    for idx, (pos, name) in enumerate(title_positions):
        if idx + 1 < len(title_positions):
            next_pos = title_positions[idx + 1][0]
        else:
            next_pos = len(paragraphs)

        # 对于权利要求书，起始段落就是内容段落（可能没有单独标题）
        if name == "权利要求书":
            # 检查当前段落是否是"权利要求书"标题（纯标题文字）
            if paragraphs[pos].text.strip() == "权利要求书":
                content_start = pos + 1
            else:
                content_start = pos
        else:
            content_start = pos + 1

        sections[name] = DocSection(name, content_start, next_pos)

    # 第三步：在"附图说明"章节中寻找附图标记段落
    mark_para = None
    mark_para_idx = None

    if "附图说明" in sections:
        sec = sections["附图说明"]
        for i in range(sec.start_idx, sec.end_idx):
            text = paragraphs[i].text.strip()
            if "附图标记" in text or re.search(r'[1-9]\s*[-\-—–]\s*[\u4e00-\u9fa5]', text):
                mark_para = paragraphs[i]
                mark_para_idx = i
                break

    # 如果在附图说明中没找到，全文搜索
    if mark_para is None:
        for i, para in enumerate(paragraphs):
            text = para.text.strip()
            if "附图标记" in text and re.search(r'[1-9]', text):
                mark_para = para
                mark_para_idx = i
                break

    return {
        'document': doc,
        'sections': sections,
        'paragraphs': paragraphs,
        'mark_para': mark_para,
        'mark_para_idx': mark_para_idx,
        'title_positions': title_positions,
    }


def get_section_text(paragraphs, section: DocSection) -> str:
    """获取指定章节的完整文本（用于预览）"""
    lines = []
    for i in range(section.start_idx, section.end_idx):
        text = paragraphs[i].text.strip()
        if text:
            lines.append(text)
    return "\n".join(lines)
