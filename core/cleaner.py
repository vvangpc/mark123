# -*- coding: utf-8 -*-
"""
cleaner.py — 文本清洗功能模块
提供：删除"所述"、半角→全角标点统一、孤立附图标记检测、错别字检查与应用。
所有文本写入操作复用 annotator.annotate_paragraph_safe()，保证格式安全。
"""
import re
from functools import lru_cache
from core.annotator import annotate_paragraph_safe, _build_xml_char_map

# 权利要求序号行头（如 "1." / "2、" / "3．"）；与 claim_check._CLAIM_HEAD_RE 语义一致
_CLAIM_HEAD_RE = re.compile(r'^\s*(\d+)\s*[\.\．\、]')


@lru_cache(maxsize=64)
def _dup_pattern(min_len: int, max_len: int) -> re.Pattern:
    """缓存「连续重复字词」正则：(X){min..max} 紧跟 1 次以上 X。"""
    return re.compile(r'(.{%d,%d})\1+' % (min_len, max_len))


def _get_active_wordbank() -> list:
    """获取当前生效的词库（合并内置 + 用户自定义，每次调用都重新读取以反映最新修改）"""
    try:
        from config.config_manager import get_merged_wordbank
        return get_merged_wordbank()
    except Exception:
        from config.typo_wordbank import WORDBANK
        return WORDBANK


# ─────────────────────────────────────────
# 1. 删除"所述"
# ─────────────────────────────────────────

def remove_suoshu(paragraphs, sections: dict, selected_section_names: list) -> int:
    """
    对 selected_section_names 中存在于 sections 的章节，删除所有"所述"。
    返回实际处理（发生替换）的段落数。
    """
    replace_dict = {"所述": ""}
    count = 0
    for name in selected_section_names:
        section = sections.get(name)
        if section is None:
            continue
        for i in range(section.start_idx, section.end_idx):
            para = paragraphs[i]
            if not para.text.strip():
                continue
            if annotate_paragraph_safe(para, replace_dict):
                count += 1
    return count


# ─────────────────────────────────────────
# 2. 半角→全角标点统一
# ─────────────────────────────────────────

# 半角 → 全角（仅在紧邻中文字符时替换，避免误伤英文/数字）
# 注意：不包含 < > ，避免误伤撰写中作为「大于/小于」使用的数学符号
# 注意：半角句点 “.” 不放在此 map 里，改由 _safe_replace_dot 单独处理，
#        以避免把权利要求书序号 “1.” “2.” 错误替换成 “1。” “2。”
# 注意：直引号 ' " 也不放在此 map 里——它们无方向，恒映射左引号会把
#        "高强度" 变成 “高强度“；改由 _safe_replace_quotes_in_paragraph
#        按「先开后闭」交替配对处理
_HALFWIDTH_MAP = {
    ",": "，",
    ";": "；",
    ":": "：",
    "?": "？",
    "!": "！",
    "(": "（",
    ")": "）",
}

# 全角 → 半角（默认不启用；用户在 UI 勾选时执行）
# 同样不处理 《》 ，避免与数学符号混淆
_FULLWIDTH_MAP = {
    "，": ",",
    "；": ";",
    "：": ":",
    "？": "?",
    "！": "!",
    "。": ".",
    "（": "(",
    "）": ")",
    "“": '"',
    "”": '"',
    "‘": "'",
    "’": "'",
}

# 匹配：中文字符 + 半角标点（可选空格）或 半角标点 + 中文字符
_CJK = r'[\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef]'


def _build_punct_replace_dict(paragraph_text: str) -> dict:
    """
    针对当前段落文本，找出需要替换的半角→全角条目（仅在中文上下文中替换）。
    返回 {half: full} 字典（仅含在本段中实际存在且紧邻中文的条目）。
    """
    result = {}
    for half, full in _HALFWIDTH_MAP.items():
        # 检测是否存在"中文+半角"或"半角+中文"模式
        pattern = f'(?:{_CJK}{re.escape(half)}|{re.escape(half)}{_CJK})'
        if re.search(pattern, paragraph_text):
            result[half] = full
    return result


# 半角句点 "." 的安全替换正则：
# 匹配"中文 + ."或". + 中文"，但排除"数字 + ."（权利要求序号 1. 2. 等）
_SAFE_DOT_RE = re.compile(
    rf'(?<!\d)\.(?={_CJK})'      # 非数字（或行首）+ . + 中文
    rf'|(?<={_CJK})\.'           # 中文 + .
)


def _safe_replace_dot_in_paragraph(paragraph) -> bool:
    """
    将段落中紧邻中文的半角句点 . 替换为全角句号 。，
    但跳过「数字 + .」的序号格式（如 "1." "12."）。

    直接操作 paragraph.runs 的 w:t 文本；返回是否有替换发生。
    """
    changed = False
    for run in paragraph.runs:
        for wt in run._r.xpath('.//w:t'):
            old = wt.text or ""
            new = _SAFE_DOT_RE.sub("。", old)
            if new != old:
                wt.text = new
                changed = True
    return changed


# 直引号配对替换：'X' → ‘X’，"X" → “X”
# 仅当本段中该引号数量为偶数（可完整配对）且至少一处紧邻中文时才转换，
# 避免误伤英寸/英尺记号（5'6"）或纯英文片段。
_QUOTE_PAIRS = {
    '"': ("“", "”"),
    "'": ("‘", "’"),
}
_CJK_CHAR_RE = re.compile(_CJK)


def _safe_replace_quotes_in_paragraph(paragraph) -> bool:
    """
    将段落中成对的半角直引号按「先开后闭」交替替换为全角弯引号。
    跨 run / w:t 的引号对也能正确配对（基于段落级字符映射原位改写）。
    返回是否有替换发生。
    """
    full_text, char_info = _build_xml_char_map(paragraph)
    if not full_text:
        return False

    edits = {}  # 位置 -> 新字符
    for ch, (open_q, close_q) in _QUOTE_PAIRS.items():
        positions = [i for i, c in enumerate(full_text) if c == ch]
        if not positions or len(positions) % 2 != 0:
            continue
        has_cjk_context = any(
            (p > 0 and _CJK_CHAR_RE.match(full_text[p - 1]))
            or (p + 1 < len(full_text) and _CJK_CHAR_RE.match(full_text[p + 1]))
            for p in positions
        )
        if not has_cjk_context:
            continue
        for k, p in enumerate(positions):
            edits[p] = open_q if k % 2 == 0 else close_q

    if not edits:
        return False

    # 替换为等长字符，可按 wt 分组原位改写
    by_wt = {}
    for pos, new_ch in edits.items():
        wt, idx = char_info[pos]
        by_wt.setdefault(wt, []).append((idx, new_ch))
    for wt, items in by_wt.items():
        chars = list(wt.text or "")
        for idx, new_ch in items:
            chars[idx] = new_ch
        wt.text = "".join(chars)
    return True


def _build_fullwidth_replace_dict(paragraph_text: str) -> dict:
    """
    针对当前段落文本，找出需要替换的全角→半角条目。
    与半角→全角不同的是：全角标点本身就是中文字符的一部分，因此无需上下文限定，
    只要段落中存在该全角字符即可纳入替换字典。
    """
    result = {}
    for full, half in _FULLWIDTH_MAP.items():
        if full in paragraph_text:
            result[full] = half
    return result


# 连续重复标点：覆盖全角 + 半角常见标点
_CONSECUTIVE_PUNCT_MAP = {
    "。。": "。",
    "，，": "，",
    "、、": "、",
    "！！": "！",
    "？？": "？",
    "；；": "；",
    "：：": "：",
    ",,": ",",
    "..": ".",
    ";;": ";",
    "::": ":",
    "??": "?",
    "!!": "!",
}


def fix_consecutive_punct(paragraphs, max_passes: int = 3, sections: dict = None) -> int:
    """
    修正连续重复的标点（如 ,,、。。、！！）。
    多次循环以处理三连及以上情况（如 。。。 → 。。 → 。）。
    若传入 sections，仅处理各章节范围内的段落。
    返回受影响的段落数（去重计数）。
    """
    if sections:
        indices = set()
        for sec in sections.values():
            indices.update(range(sec.start_idx, sec.end_idx))
        targets = [(i, paragraphs[i]) for i in sorted(indices)]
    else:
        targets = list(enumerate(paragraphs))

    affected = set()
    for _ in range(max_passes):
        changed_this_pass = False
        for idx, para in targets:
            if not para.text.strip():
                continue
            if annotate_paragraph_safe(para, _CONSECUTIVE_PUNCT_MAP):
                affected.add(idx)
                changed_this_pass = True
        if not changed_this_pass:
            break
    return len(affected)


def unify_halfwidth_punct(paragraphs, sections: dict = None) -> int:
    """
    将中文上下文中的半角标点替换为全角。
    若传入 sections，则只处理各章节范围内的段落；否则处理全文。
    返回实际发生替换的段落数。

    半角句点 "." 单独处理：跳过「数字 + .」序号格式（如 1. 2.），
    避免把权利要求书编号替换成 1。2。
    """
    count = 0

    if sections:
        indices = set()
        for sec in sections.values():
            indices.update(range(sec.start_idx, sec.end_idx))
        target_paras = [(i, paragraphs[i]) for i in sorted(indices)]
    else:
        target_paras = list(enumerate(paragraphs))

    for _, para in target_paras:
        text = para.text
        if not text.strip():
            continue
        touched = False
        # 1) 常规半角 → 全角（不含句点 "." 与引号）
        replace_dict = _build_punct_replace_dict(text)
        if replace_dict and annotate_paragraph_safe(para, replace_dict):
            touched = True
        # 2) 半角句点 "." → "。" 的安全替换（跳过 "数字." 序号格式）
        if _safe_replace_dot_in_paragraph(para):
            touched = True
        # 3) 成对直引号 → 全角弯引号（按先开后闭交替配对）
        if _safe_replace_quotes_in_paragraph(para):
            touched = True
        if touched:
            count += 1
    return count


def convert_fullwidth_to_halfwidth(paragraphs, sections: dict = None) -> int:
    """
    将段落中的全角标点替换为半角（无上下文限定）。
    用于用户主动勾选时使用，默认不启用。
    返回实际发生替换的段落数。
    """
    count = 0

    if sections:
        indices = set()
        for sec in sections.values():
            indices.update(range(sec.start_idx, sec.end_idx))
        target_paras = [(i, paragraphs[i]) for i in sorted(indices)]
    else:
        target_paras = list(enumerate(paragraphs))

    for _, para in target_paras:
        text = para.text
        if not text.strip():
            continue
        replace_dict = _build_fullwidth_replace_dict(text)
        if replace_dict and annotate_paragraph_safe(para, replace_dict):
            count += 1
    return count


# ─────────────────────────────────────────
# 3. 孤立附图标记检测
# ─────────────────────────────────────────

def detect_orphan_marks(paragraphs, sections: dict, marks: dict) -> list:
    """
    找出「在附图说明中出现、但在具体实施方式中未出现」的标记。

    参数:
        paragraphs: 全文段落列表
        sections:   解析后的章节字典
        marks:      {编号(int): 名称(str)}

    返回:
        [(编号, 名称), ...] — 孤立标记列表，按编号排序
    """
    def collect_names_in_section(section_name):
        section = sections.get(section_name)
        if section is None:
            return set()
        text = " ".join(
            paragraphs[i].text
            for i in range(section.start_idx, section.end_idx)
        )
        return {name for name in marks.values() if name in text}

    in_captions = collect_names_in_section("附图说明")
    in_impl = collect_names_in_section("具体实施方式")

    orphans = []
    for num, name in sorted(marks.items()):
        if name in in_captions and name not in in_impl:
            orphans.append((num, name))
    return orphans


def _get_section_text(paragraphs, sections: dict, section_name: str) -> str:
    """拼接指定章节的段落文本（空字符串安全）。"""
    section = sections.get(section_name)
    if section is None:
        return ""
    return " ".join(
        paragraphs[i].text
        for i in range(section.start_idx, section.end_idx)
        if 0 <= i < len(paragraphs)
    )


# 匹配 "图1" / "图 1" / "图1a" / "图1A" / "图1-2" / "图1（a）" / "图1(a)" 等形式；
# 核心编号只取阿拉伯数字，字母 / 括号后缀作为同编号的子图统一纳入该编号。
_FIGURE_REF_PATTERN = re.compile(
    r'图\s*(\d+)(?:\s*[-－–\u2013]\s*\d+)?(?:\s*[a-zA-Z])?'
)


def detect_orphan_figures(paragraphs, sections: dict) -> list:
    """
    找出「在附图说明中被提及、但在具体实施方式中没有出现」的图编号。

    例如附图说明写了『图 5 为 …』，若具体实施方式全篇没有『图5』字样，
    则图 5 属于孤立图编号，需要提醒代理人补写。

    参数:
        paragraphs: 全文段落列表
        sections:   解析后的章节字典

    返回:
        [图编号:int, ...]，按编号升序
    """
    caption_text = _get_section_text(paragraphs, sections, "附图说明")
    impl_text = _get_section_text(paragraphs, sections, "具体实施方式")
    if not caption_text:
        return []

    caption_nums = {
        int(m.group(1))
        for m in _FIGURE_REF_PATTERN.finditer(caption_text)
    }
    if not caption_nums:
        return []

    impl_nums = {
        int(m.group(1))
        for m in _FIGURE_REF_PATTERN.finditer(impl_text)
    } if impl_text else set()

    missing = sorted(n for n in caption_nums if n not in impl_nums)
    return missing


# ─────────────────────────────────────────
# 4. 错别字检查
# ─────────────────────────────────────────

def _make_locator_with_paragraphs(sections: dict, paragraphs):
    """
    工厂：基于 paragraphs 构造段落 → 章节定位函数。
    """
    if not sections:
        return lambda i: ""

    para_to_section = {}
    claims_section = None
    for name, sec in sections.items():
        for i in range(sec.start_idx, sec.end_idx):
            para_to_section[i] = name
        if name == "权利要求书":
            claims_section = sec

    para_to_claim_no = {}
    if claims_section is not None:
        current_no = None
        for i in range(claims_section.start_idx, claims_section.end_idx):
            text = paragraphs[i].text if 0 <= i < len(paragraphs) else ""
            m = _CLAIM_HEAD_RE.match(text) if text else None
            if m:
                current_no = m.group(1)
            if current_no is not None:
                para_to_claim_no[i] = current_no

    def locate(i: int) -> str:
        sect_name = para_to_section.get(i, "")
        if sect_name == "权利要求书":
            no = para_to_claim_no.get(i)
            if no:
                return f"权利要求{no}"
        return sect_name or "（未归类）"

    return locate


def check_typos_wordbank(paragraphs, sections: dict = None) -> list:
    """
    使用内置词库扫描全文（或指定章节），返回疑似错别字列表。

    返回格式:
        [{"para_idx": int, "section": str, "context": str,
          "wrong": str, "suggestion": str, "kind": "wordbank"}, ...]
    """
    results = []

    if sections:
        indices = set()
        for sec in sections.values():
            indices.update(range(sec.start_idx, sec.end_idx))
        target = sorted(indices)
    else:
        target = range(len(paragraphs))

    wordbank = _get_active_wordbank()
    locate = _make_locator_with_paragraphs(sections or {}, paragraphs)

    # 全词库合成单个 alternation 正则，每段只扫一遍（原实现为 段落数×词条数
    # 次子串查找）。长词优先排序使「权力要求书」只命中最长词条，
    # 而不是同时命中「权力要求」与「权力要求书」重复报告。
    sug_map = {e["wrong"]: e["suggestion"] for e in wordbank if e.get("wrong")}
    if not sug_map:
        return results
    pattern = re.compile("|".join(
        re.escape(w) for w in sorted(sug_map, key=len, reverse=True)
    ))

    for i in target:
        text = paragraphs[i].text
        if not text.strip():
            continue
        # 同一段中每处出现各报一条（应用时会整段全部替换，
        # 逐条上报使显示数量与实际替换处数一致）；
        # occurrence 按各 wrong 词在段内的出现次序计数
        occ_counter: dict = {}
        for m in pattern.finditer(text):
            wrong = m.group(0)
            occ = occ_counter.get(wrong, 0)
            occ_counter[wrong] = occ + 1
            pos = m.start()
            # 提取上下文（前后各15字）
            start = max(0, pos - 15)
            end = min(len(text), pos + len(wrong) + 15)
            results.append({
                "para_idx": i,
                "section": locate(i),
                "context": text[start:end],
                "wrong": wrong,
                "suggestion": sug_map[wrong],
                "kind": "wordbank",
                "occurrence": occ,
            })
    return results


# ─────────────────────────────────────────
# 4b. 连续重复字词检测（AA / ABCABC / 所述所述 等）
# ─────────────────────────────────────────

# 不参与重复检测的字符（多为标点、空白、编号符号），避免「、、」「——」等被误判
_DUP_IGNORE_CHARS = set(
    " \t\u3000，。、；：？！,.;:?!\"'“”‘’（）()【】[]{}《》<>—-—_…·"
    "0123456789０１２３４５６７８９"
)


def check_duplicate_words(paragraphs, sections: dict = None,
                          min_len: int = 1, max_len: int = 6,
                          ignore_list: list = None) -> list:
    """
    检测段落中连续重复出现的字符或词组，例如：
        "所述所述" (len=2 重复)
        "AA"     (len=1 重复)
        "ABCABC" (len=3 重复)

    参数:
        min_len / max_len: 待检测的「单元长度」范围
        ignore_list: 忽略词列表 — 若匹配的 unit 或 full 命中其中任一项，则跳过
    返回:
        [{"para_idx": int, "section": str, "context": str,
          "wrong": str, "suggestion": str, "kind": "duplicate"}, ...]
    """
    results = []
    ignore_set = set(s.strip() for s in (ignore_list or []) if s and s.strip())

    if sections:
        indices = set()
        for sec in sections.values():
            indices.update(range(sec.start_idx, sec.end_idx))
        target = sorted(indices)
    else:
        target = range(len(paragraphs))

    locate = _make_locator_with_paragraphs(sections or {}, paragraphs)
    pattern = _dup_pattern(min_len, max_len)

    seen_keys = set()  # 同段落同 wrong 去重

    for i in target:
        text = paragraphs[i].text
        if not text or not text.strip():
            continue
        # 忽略词在本段中的出现区间：每段只扫一遍，供所有 match 复用
        # （原实现对每个 match × 每个忽略词重复 find 扫描）
        ignore_spans = []
        if ignore_set:
            for ig in ignore_set:
                idx = text.find(ig)
                while idx != -1:
                    ignore_spans.append((idx, idx + len(ig)))
                    idx = text.find(ig, idx + 1)
        for m in pattern.finditer(text):
            unit = m.group(1)
            full = m.group(0)
            # 过滤掉单元全部由标点/空白/数字构成的情况
            if all(ch in _DUP_IGNORE_CHARS for ch in unit):
                continue
            # 过滤纯空白
            if not unit.strip():
                continue
            # 过滤用户自定义忽略词库（匹配 unit 或 full）
            if ignore_set and (unit in ignore_set or full in ignore_set):
                continue
            # 检查重复片段是否被忽略词在原文中的出现所覆盖
            if ignore_spans:
                match_s, match_e = m.start(), m.end()
                if any(s <= match_s and e >= match_e for s, e in ignore_spans):
                    continue
            # 过滤单字「的的、了了」等高频虚词时可后期再加（暂保留）
            key = (i, full)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            pos = m.start()
            start = max(0, pos - 15)
            end = min(len(text), pos + len(full) + 15)
            context = text[start:end]
            results.append({
                "para_idx": i,
                "section": locate(i),
                "context": context,
                "wrong": full,
                "suggestion": unit,  # 默认建议：保留一份
                "kind": "duplicate",
            })
    return results


def merge_typo_results(*result_lists) -> list:
    """
    合并多个来源（词库 / 重复词等）的结果，
    去除重复项（同一段落同一 wrong 的同一处出现只保留一条；
    occurrence 标记同段第几处出现，缺省视为第 0 处）。
    """
    seen = set()
    merged = []
    for lst in result_lists:
        if not lst:
            continue
        for item in lst:
            key = (item["para_idx"], item["wrong"], item.get("occurrence", 0))
            if key in seen:
                continue
            seen.add(key)
            merged.append(item)
    return merged


def apply_typo_corrections(paragraphs, corrections: list) -> int:
    """
    将用户确认的修正写回内存中的段落对象（不保存文件）。

    参数:
        paragraphs:  全文段落列表
        corrections: [{"para_idx": int, "wrong": str, "confirmed_fix": str}, ...]
                     confirmed_fix 为空字符串时跳过该条。

    返回:
        实际发生替换的段落数。
    """
    # 按段落分组，批量替换（同一段落可能有多处修正）
    from collections import defaultdict
    para_replace: dict[int, dict] = defaultdict(dict)
    for item in corrections:
        fix = item.get("confirmed_fix", "").strip()
        if not fix:
            continue
        if item["wrong"] != fix:
            para_replace[item["para_idx"]][item["wrong"]] = fix

    count = 0
    for para_idx, replace_dict in para_replace.items():
        if annotate_paragraph_safe(paragraphs[para_idx], replace_dict):
            count += 1
    return count
