# -*- coding: utf-8 -*-
"""
spec_check.py — 说明书 / 具体实施方式 只读审计

与 claim_check.py 对称：纯函数、只读、不改文档树，可逐步生长为说明书检查族。
当前提供两项：

    check_embodiment_numbering(paragraphs, sections)
        扫描「具体实施方式」，校验实施例标题序号：实施例一、实施例二…
        应从「一」开始、连续、不重复、不颠倒。

    check_abstract_length(paragraphs, sections, limit=300)
        校验「说明书摘要」文字部分是否超过细则规定的 300 字。

结果 dict 沿用全仓统一 schema（见 core/claim_check.py），并补一个 `wrong`
字段作为 1框 内联高亮的锚点（main_window 直接拿去喂 content_area.highlight_issues）：

    {"kind": str, "para_idx": int, "context": str,
     "message": str, "suggestion": str, "wrong": str}
"""
import re

# ── 中文数字 ↔ 整数 ────────────────────────────────────────────────
_CN_VAL = {
    "零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
    "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
}
_FW2HW = str.maketrans("０１２３４５６７８９", "0123456789")
_CN_DIGITS = "零一二三四五六七八九"


def _cn_to_int(s: str):
    """中文 / 阿拉伯（含全角）数字串 → int；无法解析返回 None。
    覆盖 一–九、十、十一–十九、二十–九十九（含「两」=2、「〇/零」）。"""
    if not s:
        return None
    s = s.strip().translate(_FW2HW)
    if s.isascii() and s.isdigit():
        return int(s)
    if s == "十":
        return 10
    if s.startswith("十"):                       # 十一 ~ 十九
        rest = s[1:]
        if len(rest) == 1 and rest in _CN_VAL:
            return 10 + _CN_VAL[rest]
        return None
    if "十" in s:                                # 二十 ~ 九十九
        tens, ones = s.split("十", 1)
        if tens not in _CN_VAL:
            return None
        val = _CN_VAL[tens] * 10
        if ones:
            if ones not in _CN_VAL:
                return None
            val += _CN_VAL[ones]
        return val
    if len(s) == 1 and s in _CN_VAL:             # 一 ~ 九
        return _CN_VAL[s]
    return None


def _int_to_cn(n: int) -> str:
    """1–99 的 int → 中文数字（用于提示文案，实施例极少超过 99）。"""
    if n <= 0:
        return str(n)
    if n < 10:
        return _CN_DIGITS[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + (_CN_DIGITS[n % 10] if n % 10 else "")
    if n < 100:
        out = _CN_DIGITS[n // 10] + "十"
        if n % 10:
            out += _CN_DIGITS[n % 10]
        return out
    return str(n)


# ── 实施例标题识别 ────────────────────────────────────────────────
_NUM = r"(?:[一二三四五六七八九十百零〇两]+|\d+)"
# 起锚：可选「第N」+ 实施例/实施方式 + 可选「N」。两处数字至多有一处。
_HEAD_RE = re.compile(rf"^\s*(?:第\s*({_NUM})\s*)?实施(?:例|方式)(?:\s*({_NUM}))?")
# 标题token 之后必须是边界（标点 / 右括号 / 空白 / 串尾）才算「标题」；
# 若紧跟续写汉字（如 实施例一「中」/「所述」/「的」）则判为行内引用，丢弃。
_BOUNDARY_CHARS = set("：:、，。,.；;（）()【】[]｜|/\\-－–—~～　")


def _heading_ordinal(text: str):
    """把一段文字判定为实施例标题；返回 (序号int, 标题文本) 或 None。

    标题文本用于 1框 高亮锚点，取自原始 text（保留其中空格），保证 find 命中。
    """
    if not text:
        return None
    m = _HEAD_RE.match(text)
    if not m:
        return None
    raw = m.group(2) or m.group(1)        # 实施例X 的 X，或 第X实施例 的 X
    if raw is None:
        return None                        # 裸「实施例」无序号 → 不计入
    end = m.end()
    if end < len(text):
        nxt = text[end]
        if nxt not in _BOUNDARY_CHARS and not nxt.isspace():
            return None                    # 续写汉字/字母 → 行内引用
    val = _cn_to_int(raw)
    if val is None:
        return None
    token = text[m.start():end].strip()    # 如「实施例三」「第一实施例」
    return val, token


def check_embodiment_numbering(paragraphs, sections) -> list:
    """校验「具体实施方式」中实施例标题序号：从一连续、不重、不颠倒。

    返回结果列表，kind ∈ {emb_start, emb_gap, emb_dup, emb_order}。
    每条按文档顺序、每个标题至多报一类（dup > order > gap），起始问题单列。
    """
    results: list = []
    sec = sections.get("具体实施方式") if sections else None
    if sec is None or sec.start_idx >= sec.end_idx:
        return results

    # 文档顺序收集所有实施例标题（不排序——顺序即信号）
    found = []   # (序号, para_idx, 标题文本)
    for i in range(sec.start_idx, sec.end_idx):
        try:
            text = paragraphs[i].text
        except Exception:
            continue
        r = _heading_ordinal(text)
        if r is not None:
            found.append((r[0], i, r[1]))
    if not found:
        return results

    seq_ctx = "、".join(_int_to_cn(v) for v, _, _ in found)

    # 起始非「一」：单列一条，并抑制由此引出的「缺 1..k」gap，避免和起始双报
    suppress_lead_gap = False
    first_val, first_pid, first_token = found[0]
    if first_val != 1:
        results.append({
            "kind": "emb_start", "para_idx": first_pid, "context": seq_ctx,
            "message": f"实施例起始为「{first_token}」，应从「实施例一」开始",
            "suggestion": "将第一个实施例改为 实施例一",
            "wrong": first_token,
        })
        suppress_lead_gap = True

    expected, prev_max, seen = 1, 0, set()
    for idx, (val, pid, token) in enumerate(found):
        if val in seen:
            results.append({
                "kind": "emb_dup", "para_idx": pid, "context": seq_ctx,
                "message": f"实施例「{token}」重复出现",
                "suggestion": "删除或修正重复的实施例编号", "wrong": token,
            })
        elif val < prev_max:
            results.append({
                "kind": "emb_order", "para_idx": pid, "context": seq_ctx,
                "message": f"实施例「{token}」顺序颠倒（出现在「实施例{_int_to_cn(prev_max)}」之后）",
                "suggestion": "按升序排列实施例", "wrong": token,
            })
        else:  # val > prev_max，升序新值
            if val > expected and not (idx == 0 and suppress_lead_gap):
                miss = "、".join(f"实施例{_int_to_cn(m)}" for m in range(expected, val))
                results.append({
                    "kind": "emb_gap", "para_idx": pid, "context": seq_ctx,
                    "message": f"实施例序号不连续：缺失 {miss}",
                    "suggestion": "补齐缺失的实施例或重排编号", "wrong": token,
                })
            expected = val + 1
            prev_max = val
        seen.add(val)
    return results


# ── 摘要字数 ──────────────────────────────────────────────────────
_WS_RE = re.compile(r"\s+")


def check_abstract_length(paragraphs, sections, limit: int = 300) -> list:
    """校验「说明书摘要」文字部分是否超过 limit（默认 300）字。

    超限时返回一条 abstract_len，wrong 锚点落在「跨过第 limit 字」那一段的
    越界起点到段尾，便于 1框 高亮超出部分。"""
    sec = (sections.get("说明书摘要") or sections.get("摘要")) if sections else None
    if sec is None or sec.start_idx >= sec.end_idx:
        return []

    body = []   # (para_idx, text)，跳过标题段与空段
    for i in range(sec.start_idx, sec.end_idx):
        try:
            text = paragraphs[i].text
        except Exception:
            continue
        clean = _WS_RE.sub("", text)
        if not clean or clean in ("说明书摘要", "摘要"):
            continue
        body.append((i, text))

    total = sum(len(_WS_RE.sub("", t)) for _, t in body)
    if total <= limit:
        return []

    # 定位跨过第 limit 字的那段与段内偏移（越界起点）
    cum = 0
    anchor_pid, anchor_off = -1, 0
    for pid, text in body:
        nws = 0
        for off, ch in enumerate(text):
            if not ch.isspace():
                nws += 1
                if cum + nws == limit + 1:
                    anchor_pid, anchor_off = pid, off
                    break
        if anchor_pid >= 0:
            break
        cum += nws

    wrong = ""
    if anchor_pid >= 0:
        wrong = paragraphs[anchor_pid].text[anchor_off:]
    return [{
        "kind": "abstract_len", "para_idx": anchor_pid, "context": f"{total}字",
        "message": f"说明书摘要约 {total} 字，超出规定上限 {limit} 字（多 {total - limit} 字）",
        "suggestion": f"精简摘要至 {limit} 字以内", "wrong": wrong,
    }]
