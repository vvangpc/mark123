# -*- coding: utf-8 -*-
"""
core/paragraph_edit.py — 段落级格式安全写回工具

把内容区/预览框里编辑后的整段文本写回 python-docx 段落，尽量保留段落中的
非文本节点（公式 m:oMath、图片 w:drawing/w:pict、OLE w:object 等）。

阶段一：承载从 claim_check 抽出的 set_paragraph_text。
阶段二：在此扩展结构化编辑所需的更完善回写（含对象段只读判定等），
        与 annotator._build_xml_char_map 复用同一套 <w:t> 字符映射理念。
"""


def set_paragraph_text(paragraph, new_text: str) -> bool:
    """
    把一个段落的可见文本整体覆盖为 new_text，
    同时尽量保留段落中的非文本节点（公式、图片等）。
    实现策略：把所有 w:t 节点的文字清空，再把 new_text 全部写到首个 w:t。

    返回: 是否发生了实际改动
    """
    # 先收集所有 w:t 节点
    wt_nodes = []
    for run in paragraph.runs:
        for wt in run._r.xpath('.//w:t'):
            wt_nodes.append(wt)

    # 计算旧文本
    old_text = "".join((wt.text or "") for wt in wt_nodes)
    if old_text == new_text:
        return False

    if not wt_nodes:
        # 整段没有任何 w:t：通过 add_run 插入
        paragraph.add_run(new_text)
        return True

    # 写入首节点，其他清空
    wt_nodes[0].text = new_text
    for wt in wt_nodes[1:]:
        wt.text = ""
    return True
