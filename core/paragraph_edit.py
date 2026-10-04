# -*- coding: utf-8 -*-
"""
core/paragraph_edit.py — 段落级格式安全写回工具

把内容区里编辑后的整段文本写回 python-docx 段落：

  · 显示文本（display_text）与 para.text 取同一范围（直属 w:r + 超链接里的 w:r），
    逐字对齐，只把软回车 w:br / w:cr 显示成「↵」——para.text 里它们是 "\\n"，
    放进编辑器会把一段拆成两行，破坏「行 = 段」。
  · 回写（set_paragraph_text）按字符差分只改动用户真正改过的字符：未改动的字留在
    原 w:t（原 run 格式、上下标都在），制表符 / 软回车这类结构元素保持原位，
    公式、图片等非文本节点不碰。新增的文字跟随其前一个字的 run 格式。
"""
from difflib import SequenceMatcher

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_WT = qn("w:t")
_WR = qn("w:r")
_BR = qn("w:br")
_CR = qn("w:cr")
_TAB = qn("w:tab")
_PTAB = qn("w:ptab")
_NBH = qn("w:noBreakHyphen")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

LINE_BREAK_MARK = "↵"
# 只能由文档结构产生、不能靠在编辑器里打字新增的字符
_STRUCTURAL_CHARS = {"\t", LINE_BREAK_MARK}


def token_text(el) -> str | None:
    """run 内结构元素的显示字符；w:t 与非文本元素返回 None。"""
    tag = el.tag
    if tag in (_TAB, _PTAB):
        return "\t"
    if tag == _BR:
        # 与 python-docx 一致：分页 / 分栏符不产生文字
        return LINE_BREAK_MARK if el.get(qn("w:type"), "textWrapping") == "textWrapping" else ""
    if tag == _CR:
        return LINE_BREAK_MARK
    if tag == _NBH:
        return "-"
    return None


def iter_text_runs(paragraph):
    """按文档顺序产出与 para.text 同范围的 w:r（直属 + 超链接内）。"""
    return paragraph._p.xpath("w:r | w:hyperlink/w:r")


def build_char_map(paragraph):
    """返回 (显示文本, 槽位列表)。

    槽位与显示文本逐字对应：w:t 里的字为 (w:t 元素, None)，结构元素为 (None, 元素)。
    """
    chars = []
    slots = []
    for r in iter_text_runs(paragraph):
        for el in r.iterchildren():
            if el.tag == _WT:
                for ch in el.text or "":
                    chars.append(ch)
                    slots.append((el, None))
                continue
            tok = token_text(el)
            if tok:
                chars.append(tok)
                slots.append((None, el))
    return "".join(chars), slots


def display_text(paragraph) -> str:
    """内容区里一行显示的文本（与 para.text 逐字对齐，软回车为「↵」）。"""
    return build_char_map(paragraph)[0]


def replace_chars(paragraph, edits: dict) -> bool:
    """按显示文本位置原位改字：edits = {位置: 新字符串}。

    只改 w:t 里的字（结构元素位置的条目忽略），每处改动留在原字所在的 w:t，
    run 格式不变。用于标点转换这类「逐个字符」的替换。
    """
    if not edits:
        return False
    _, slots = build_char_map(paragraph)
    by_wt: dict = {}
    offset = 0
    prev = None
    for pos, (wt, _) in enumerate(slots):
        if wt is None:
            continue
        offset = offset + 1 if wt is prev else 0
        prev = wt
        if pos in edits:
            by_wt.setdefault(wt, {})[offset] = edits[pos]
    for wt, items in by_wt.items():
        chars = list(wt.text or "")
        for k, s in items.items():
            chars[k] = s
        wt.text = "".join(chars)
    return bool(by_wt)


def set_paragraph_text(paragraph, new_text: str) -> bool:
    """把段落的显示文本改成 new_text，只改动差异部分。

    new_text 里新出现的制表符 / 「↵」无法凭空造出结构元素，会被忽略；
    调用方可用 display_text 回读判断是否与预期一致。
    返回：是否发生了实际改动。
    """
    old, slots = build_char_map(paragraph)
    if old == new_text:
        return False

    if not slots:
        insert = "".join(ch for ch in new_text if ch not in _STRUCTURAL_CHARS)
        if not insert:
            return False
        paragraph.add_run(insert)
        return True

    buffers: dict = {}
    for wt, _ in slots:
        if wt is not None:
            buffers.setdefault(wt, [])
    removed_tokens = []
    anchor = None          # 上一个保留下来的槽位：("wt", 元素) / ("tok", 元素)

    def _target(i1: int, i2: int):
        """新文字写入哪个 w:t。

        紧跟前一个保留的字；否则用被替换区间里（或紧接其后）的第一个 w:t；
        插入点挨着一个保留的结构元素时，就在它旁边新建 w:t（继承该 run 的格式）。
        """
        if anchor is not None and anchor[0] == "wt":
            return anchor[1]
        k = i1
        while k < len(slots):
            if slots[k][0] is not None:
                return slots[k][0]
            if k >= i2:
                break
            k += 1
        wt = OxmlElement("w:t")
        if anchor is not None:
            anchor[1].addnext(wt)
        else:
            # 段首插入，其后是结构元素（或整段都是将被删掉的结构元素）
            (slots[k] if k < len(slots) else slots[i1])[1].addprevious(wt)
        buffers[wt] = []
        return wt

    matcher = SequenceMatcher(None, old, new_text, autojunk=False)
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op == "equal":
            for i in range(i1, i2):
                wt, tok = slots[i]
                if wt is not None:
                    buffers[wt].append(old[i])
                    anchor = ("wt", wt)
                else:
                    anchor = ("tok", tok)
            continue
        for i in range(i1, i2):
            tok = slots[i][1]
            if tok is not None:
                removed_tokens.append(tok)
        insert = "".join(ch for ch in new_text[j1:j2] if ch not in _STRUCTURAL_CHARS)
        if insert:
            wt = _target(i1, i2)
            buffers[wt].append(insert)
            anchor = ("wt", wt)

    for wt, buf in buffers.items():
        text = "".join(buf)
        wt.text = text
        if text != text.strip():
            wt.set(_XML_SPACE, "preserve")
    for tok in removed_tokens:
        parent = tok.getparent()
        if parent is not None:
            parent.remove(tok)
    return True
