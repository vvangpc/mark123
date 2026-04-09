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
    ("说明书摘要", ["说明书摘要", "摘要", "摘 要"]),
    ("具体实施方式", ["具体实施方式"]),
    ("说明书附图", ["说明书附图"]),
    ("摘要附图", ["摘要附图"]),
]


def is_section_title(paragraph, keywords: list[str]) -> bool:
    """
    增强版：判断一个段落是否是指定的章节标题。
    无视中间的空格、全半角空格，以及末尾可能存在的冒号。
    """
    raw_text = paragraph.text.strip()
    if not raw_text:
        return False

    # 暴力清理：去除字符串中的所有空白字符（包括全角空格、半角空格、制表符）
    clean_text = re.sub(r'\s+', '', raw_text)

    for kw in keywords:
        clean_kw = re.sub(r'\s+', '', kw)
        
        # 1. 完全匹配
        if clean_text == clean_kw:
            return True
        
        # 2. 匹配带中英文冒号的情况 (如 "摘要：" 或 "摘要:")
        if clean_text == f"{clean_kw}：" or clean_text == f"{clean_kw}:":
            return True
            
        # 3. 字体加粗匹配 (如果首个Run是加粗的，且文本以关键词开头)
        # 例如："说明书附图如下" 如果加粗了也可以算
        if clean_text.startswith(clean_kw) and paragraph.runs:
            first_run = paragraph.runs[0]
            if first_run.font and first_run.font.bold:
                return True
                
        # 4. 如果是找摘要，单独给个直接识别前缀为“摘要”但紧跟附图或冒号的条件
        if "摘要" in clean_kw and (clean_text.startswith("摘要附图") or clean_text.startswith("摘要:")):
             return True

    return False


def _has_image(para) -> bool:
    """检查段落是否包含图片/OLE对象"""
    ns_w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    ns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    ns_draw = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    el = para._element
    if el.findall(f'.//{ns_w}object'):
        return True
    if el.findall(f'.//{ns_w}pict'):
        return True
    if el.findall(f'.//{ns_draw}blip'):
        return True
    # drawing 元素（内联图片）
    ns_wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    if el.findall(f'.//{ns_wp}inline') or el.findall(f'.//{ns_wp}anchor'):
        return True
    return False


def _trim_patent_title_from_claims(paragraphs, sections: dict) -> None:
    """
    把混入到"权利要求书"尾部的专利名称（说明书首页发明名称标题）修剪掉。

    判定尾段是否为专利名称标题（需同时满足）：
      • 非空；
      • 不以权项序号 ^\\d+[.、．] 开头；
      • 不含"所述"/"其特征"等权项常见结构词；
      • 不以句读（。，；；,. ;）收尾；
      • 长度小于 30 字（典型的专利名称多为十几字以内）。

    若满足则连同其前的空段一并剪掉。为安全起见，最多只剪掉 3 个非空尾段。
    """
    if "权利要求书" not in sections:
        return
    sec = sections["权利要求书"]
    new_end = sec.end_idx
    trimmed_nonempty = 0
    max_trim_nonempty = 3

    while new_end > sec.start_idx:
        text = paragraphs[new_end - 1].text.strip()
        if not text:
            # 尾部空段：直接跳过但不计入数量
            new_end -= 1
            continue

        looks_like_claim = (
            bool(re.match(r'^\s*\d+\s*[.、．]', text))
            or text.endswith(("。", "，", "；", ";", ",", "."))
            or ("其特征" in text)
            or ("所述" in text)
            or len(text) >= 30
        )
        if looks_like_claim:
            break

        # 视为专利名称标题，剪掉
        new_end -= 1
        trimmed_nonempty += 1
        if trimmed_nonempty >= max_trim_nonempty:
            break

    if new_end != sec.end_idx and new_end > sec.start_idx:
        sections["权利要求书"] = DocSection(
            "权利要求书", sec.start_idx, new_end
        )


def _infer_abstract_boundary(paragraphs, sections: dict, title_positions: list):
    """
    后处理：如果"具体实施方式"的区间一直延伸到了文档末尾附近，  
    检查其末尾是否存在"说明书附图 → 说明书摘要"的典型结构。
    如果找到，自动切分出"说明书附图"和"说明书摘要"区间。
    
    典型的专利文档末尾结构（无独立标题）:
        ... 具体实施方式正文 ...
        [空行若干]
        [图片]           <- 说明书附图起点
        图1
        [图片]
        图2
        [空行若干]
        摘要正文         <- 说明书摘要起点
        [空行若干]
        摘要附图指定为图1
    """
    if "具体实施方式" not in sections:
        return
    
    impl_sec = sections["具体实施方式"]
    total = len(paragraphs)
    
    # 只在"具体实施方式"延伸到文档末尾时才做推断
    # （如果后面已经有其他正确切分的章节就不需要处理了）
    if impl_sec.end_idx < total - 5:
        return
    
    # 从"具体实施方式"区间的末尾向前扫描，寻找图片区域
    # 图片区域的特征：连续的图片段落和"图N"标签
    drawing_start = None  # 说明书附图区域的起始索引
    abstract_start = None  # 说明书摘要的起始索引
    
    # 从末尾向前找到最后一个正文段落（非空、非图号、非"摘要附图指定为…"）
    # 这就是摘要正文
    last_content_idx = None
    for i in range(impl_sec.end_idx - 1, impl_sec.start_idx, -1):
        text = paragraphs[i].text.strip()
        clean = re.sub(r'\s+', '', text)
        # 跳过空段落
        if not text:
            continue
        # 跳过"摘要附图指定为图N"这种末尾标记
        if clean.startswith("摘要附图"):
            continue
        # 找到最后一个实质正文段落
        last_content_idx = i
        break
    
    if last_content_idx is None:
        return
    
    # 从 last_content_idx 向前找，看其前方是否有图片区域
    # 先定位图片区域（图片段落 + "图N" 标签段落）
    # 从 last_content_idx 往前扫
    scan_start = last_content_idx - 1
    found_figure_label = False
    figure_region_end = None
    
    for i in range(scan_start, impl_sec.start_idx, -1):
        text = paragraphs[i].text.strip()
        has_img = _has_image(paragraphs[i])
        
        # 匹配"图1"、"图 2"这样的图号标签
        is_figure_label = bool(re.match(r'^图\s*\d+$', text))
        
        if is_figure_label or has_img:
            if not found_figure_label:
                figure_region_end = i
            found_figure_label = True
            drawing_start = i
        elif text == "":
            # 空行可以穿越
            if found_figure_label:
                drawing_start = i
            continue
        else:
            # 碰到了非图片、非空的正文段落 -> 停止
            if found_figure_label:
                break
            else:
                # 在到达图片区域之前就碰到了正文 -> 没有附图区域
                return
                
    if not found_figure_label or drawing_start is None:
        return
    
    # 摘要正文起点 = 图片区域之后、非空的第一个段落
    for i in range(figure_region_end + 1, impl_sec.end_idx):
        text = paragraphs[i].text.strip()
        if text:
            abstract_start = i
            break
    
    if abstract_start is None:
        return
    
    # 验证摘要正文确实像是摘要（通常是一段较长的概述性文字）
    abstract_text = paragraphs[abstract_start].text.strip()
    if len(abstract_text) < 15:
        # 太短了，可能不是摘要
        return

    # ===== 执行切分 =====
    # 1. 缩短"具体实施方式"的范围
    sections["具体实施方式"] = DocSection("具体实施方式", impl_sec.start_idx, drawing_start)
    
    # 2. 新增"说明书附图"区间（如果尚不存在）
    if "说明书附图" not in sections:
        sections["说明书附图"] = DocSection("说明书附图", drawing_start, abstract_start)
        title_positions.append((drawing_start, "说明书附图"))
    
    # 3. 新增"说明书摘要"区间（如果尚不存在或范围为空）
    if "说明书摘要" not in sections or sections["说明书摘要"].start_idx >= sections["说明书摘要"].end_idx:
        sections["说明书摘要"] = DocSection("说明书摘要", abstract_start, impl_sec.end_idx)
        # 移除旧的空摘要 title_position（如果有的话）
        title_positions[:] = [(p, n) for p, n in title_positions if n != "说明书摘要"]
        title_positions.append((abstract_start, "说明书摘要"))
    
    title_positions.sort(key=lambda x: x[0])


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
    # 权利要求书的特征：段落以 "1." 或 "1、" 开头，且包含"其特征"
    # （放宽匹配：原先强制要求 "其特征在于"，但部分专利文本写作
    #  "其特征是" / "其特征为" / "其特征包括" 等，会导致权利要求书识别失败）
    section_names_found = [name for _, name in title_positions]

    if "权利要求书" not in section_names_found:
        # 尝试通过内容特征定位权利要求书
        for i, para in enumerate(paragraphs):
            text = para.text.strip()
            if re.match(r'^1[.、．]', text) and '其特征' in text:
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

    # 第三步（后处理 A）：修剪"权利要求书"尾部混入的专利名称
    # 许多专利文档中，权项最后一条与"技术领域"之间会夹一行【发明名称 / 专利
    # 名称】（说明书首页标题），导致权利要求书区间把它当成最后一段保留。
    # 判定尾段是否为专利名称标题：
    #   • 非空
    #   • 不以权项序号 \d+[.、．] 开头
    #   • 不含"所述"/"其特征"等权项常用词
    #   • 不以句读（。，；,. ;）收尾
    #   • 长度小于 30 字（典型专利名称通常只有十几个字）
    _trim_patent_title_from_claims(paragraphs, sections)

    # 第三步（后处理 B）：修正"具体实施方式"可能吞掉"说明书摘要"的问题
    # 许多专利文档中,"说明书摘要"没有独立的标题段落，其正文直接跟在说明书附图
    # 的图片之后。此时"具体实施方式"区间会一路延伸到文档末尾，把摘要正文一并包含。
    # 解决策略：如果检测到"具体实施方式"且其范围延伸到文档末尾附近，从末尾向前
    # 扫描，寻找"说明书附图"区域（IMG段落 + 图号标签 如"图1"），然后把其后的
    # 非空正文段落标记为"说明书摘要"的起点。
    _infer_abstract_boundary(paragraphs, sections, title_positions)

    # 第四步：在"附图说明"章节中寻找附图标记段落
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
