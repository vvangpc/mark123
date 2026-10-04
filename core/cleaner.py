# -*- coding: utf-8 -*-
"""
cleaner.py — 文本清洗功能模块
提供：删除"所述"、半角→全角标点统一、孤立附图标记检测、错别字检查与应用。
文本写入走 annotator.annotate_paragraph_safe() 或 core.paragraph_edit（只动 w:t），保证格式安全。
"""
import re
from functools import lru_cache
from core.annotator import annotate_paragraph_safe, _build_xml_char_map
from core.paragraph_edit import display_text, replace_chars, set_paragraph_text

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

# "中文上下文"字符：汉字 + 中文标点 + 全角标点符号。
# 刻意不含全角数字 / 全角字母（FF10-19、FF21-3A、FF41-5A）：「１.」是序号，
# 「Ａ,Ｂ」是字母列举，都不该被当成中文语境改掉。
_CJK = (r'[\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f'
        r'\uff01-\uff0f\uff1a-\uff20\uff3b-\uff40\uff5b-\uff65]')
_CJK_CHAR_RE = re.compile(_CJK)


def _is_cjk(ch: str) -> bool:
    return bool(ch) and bool(_CJK_CHAR_RE.match(ch))


def _halfwidth_edits(text: str) -> dict:
    """算出本段需要改成全角的位置 → 全角字符。

    逐处判断而不是"本段有一处挨着中文就全段替换"——后者会把
    「采用公式f(x,y)计算,结果为1:2」改成「f(x，y）…1：2」。
      · , ; : ? !   只改紧邻中文的那一处；
      · ( )         成对判断：括号内有中文，或括号两侧都是中文 / 段首段尾
                    （壳体(1)的 → 壳体（1）的），整对一起改；f(x,y) 这类保持半角。
                    落单的括号按紧邻中文处理。
    """
    edits = {}
    stack, lone, pairs = [], [], []
    for i, ch in enumerate(text):
        if ch == "(":
            stack.append(i)
        elif ch == ")":
            if stack:
                pairs.append((stack.pop(), i))
            else:
                lone.append(i)
    lone.extend(stack)

    def _near_cjk(i: int) -> bool:
        return (i > 0 and _is_cjk(text[i - 1])) or (i + 1 < len(text) and _is_cjk(text[i + 1]))

    for o, c in pairs:
        left = text[o - 1] if o > 0 else ""
        right = text[c + 1] if c + 1 < len(text) else ""
        outer = (not left or _is_cjk(left)) and (not right or _is_cjk(right))
        if outer or any(_is_cjk(ch) for ch in text[o + 1:c]):
            edits[o] = "（"
            edits[c] = "）"
    for i in lone:
        if _near_cjk(i):
            edits[i] = _HALFWIDTH_MAP[text[i]]
    for i, ch in enumerate(text):
        if ch in _HALFWIDTH_MAP and ch not in "()" and _near_cjk(i):
            edits[i] = _HALFWIDTH_MAP[ch]
    return edits


# 半角句点 "." 的安全替换正则：
# 匹配"中文 + ."或". + 中文"，但排除"数字 + ."（权利要求序号 1. 2. 等）
_SAFE_DOT_RE = re.compile(
    rf'(?<!\d)\.(?={_CJK})'      # 非数字（或行首）+ . + 中文
    rf'|(?<={_CJK})\.'           # 中文 + .
)


def _dot_edits(text: str) -> dict:
    """紧邻中文的半角句点 → 「。」，跳过「数字 + .」序号。

    在整段显示文本上匹配：「1」与「.一种」分在两个 run 时，逐个 w:t 匹配的旧实现
    看不到前一个 run 的数字，会把权项编号改成「1。一种」。
    """
    return {m.start(): "。" for m in _SAFE_DOT_RE.finditer(text)}


# 直引号配对替换：'X' → ‘X’，"X" → “X”
# 仅当本段中该引号数量为偶数（可完整配对）且至少一处紧邻中文时才转换，
# 避免误伤英寸/英尺记号（5'6"）或纯英文片段。
_QUOTE_PAIRS = {
    '"': ("“", "”"),
    "'": ("‘", "’"),
}


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


# 注：曾经的 fix_consecutive_punct（"修正连续重复标点"勾选项，不问自动改）
# 已在 V4.2.1 移除 —— 同一件事由「🔣 标点 → 重复标点检查」承担，
# 后者会先把每一处列进 2框 结果表让用户逐条确认，比闷头替换安全得多。
# 检出与建议逻辑见本文件 4c 节的 check_duplicate_punct。


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
        text = display_text(para)
        if not text.strip():
            continue
        touched = False
        # 1) 常规半角 → 全角（不含引号）+ 2) 半角句点 "." → "。"（跳过 "数字." 序号）
        #    均为等长逐字替换，位置在同一份文本上算，可一次写回
        edits = _halfwidth_edits(text)
        edits.update(_dot_edits(text))
        if replace_chars(para, edits):
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
        # 同一段中每处出现各报一条，应用时按 occurrence 只改那一处；
        # occurrence 为 1 基、与定位 / 应用共用 occurrence_at / nth_occurrence 计数。
        for m in pattern.finditer(text):
            wrong = m.group(0)
            pos = m.start()
            occ = occurrence_at(text, wrong, pos)
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
            # 同段同词每处各报一条（应用时按 occurrence 只改这一处）
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
                "occurrence": occurrence_at(text, full, pos),
            })
    return results


# ─────────────────────────────────────────
# 4c. 重复标点检测（先看后改，结果进共用结果表）
# ─────────────────────────────────────────

# 参与检测的标点集合：全角常见停顿 + 对应半角 + 半角句点。
# 逐条确认后经 apply_typo_corrections 写回，与错别字 / 重复字词同一条通路。
_DUP_PUNCT_RUN_RE = re.compile(r'[。，、；：？！,;:?!.]{2,}')

# 归一化：半角 → 对应全角；顿号在"重复"意义上与逗号同类
# （「，、」「、，」属于同一个停顿被写了两遍，应一并报出）
_PUNCT_NORM = {
    ",": "，", ";": "；", ":": "：", "?": "？", "!": "！", ".": "。",
    "、": "，",
}


def _dup_punct_suggestion(run: str) -> str:
    """给一段重复标点选一个保留字符：优先全角，其次出现最多者，再次取首个。"""
    full = [ch for ch in run if ch not in _PUNCT_NORM]   # 不在归一化表里的即全角本身
    pool = full or list(run)
    best, best_n = pool[0], 0
    for ch in dict.fromkeys(pool):          # 保持首次出现顺序，平局取靠前的
        n = pool.count(ch)
        if n > best_n:
            best, best_n = ch, n
    return best


def check_duplicate_punct(paragraphs, sections: dict = None) -> list:
    """
    检测「重复标点」——同一个停顿被连写了两遍以上，包括：
        "。。" "，，" "、、"      同字符连写
        "，," ",，" "。." ".。"   全角/半角混写
        "，、" "、，"             逗号与顿号混写

    只检测不修改；返回结构与错别字 / 重复字词检查一致，可直接进共用结果表，
    并经 apply_typo_corrections 写回内存。

    返回:
        [{"para_idx": int, "section": str, "context": str,
          "wrong": str, "suggestion": str, "kind": "punct_dup",
          "occurrence": int(1 基)}, ...]
    """
    results = []

    if sections:
        indices = set()
        for sec in sections.values():
            indices.update(range(sec.start_idx, sec.end_idx))
        target = sorted(indices)
    else:
        target = range(len(paragraphs))

    locate = _make_locator_with_paragraphs(sections or {}, paragraphs)

    for i in target:
        text = paragraphs[i].text
        if not text or not text.strip():
            continue
        for m in _DUP_PUNCT_RUN_RE.finditer(text):
            run = m.group(0)
            # 英文省略号 "..."（3 个及以上半角句点）是合法写法，跳过
            if set(run) == {"."} and len(run) >= 3:
                continue
            # 归一化后必须全部相同，才算"同一个停顿写了两遍"；
            # 「，。」这类不同停顿的连写含义不明确，不在本检查范围内
            norm = {_PUNCT_NORM.get(ch, ch) for ch in run}
            if len(norm) != 1:
                continue
            keep = _dup_punct_suggestion(run)
            pos = m.start()
            occ = occurrence_at(text, run, pos)
            start = max(0, pos - 15)
            end = min(len(text), pos + len(run) + 15)
            results.append({
                "para_idx": i,
                "section": locate(i),
                "context": text[start:end],
                "wrong": run,
                "suggestion": keep,
                "kind": "punct_dup",
                "occurrence": occ,
            })
    return results


def merge_typo_results(*result_lists) -> list:
    """
    合并多个来源（词库 / 重复词等）的结果，
    去除重复项（同一段落同一 wrong 的同一处出现只保留一条；
    occurrence 标记同段第几处出现（1 基），缺省视为第 1 处）。
    """
    seen = set()
    merged = []
    for lst in result_lists:
        if not lst:
            continue
        for item in lst:
            key = (item["para_idx"], item["wrong"], item.get("occurrence", 1))
            if key in seen:
                continue
            seen.add(key)
            merged.append(item)
    return merged


def nth_occurrence(text: str, needle: str, n: int) -> int:
    """needle 在 text 中第 n 次出现（1 基，允许重叠）的偏移；找不到返回 -1。"""
    if not text or not needle or n < 1:
        return -1
    idx = -1
    for _ in range(n):
        idx = text.find(needle, idx + 1)
        if idx == -1:
            return -1
    return idx


def occurrence_at(text: str, needle: str, pos: int) -> int:
    """pos 处的 needle 是第几次出现（与 nth_occurrence 同一套计数，1 基）。

    检查结果的 occurrence 必须用这套计数：按正则 finditer 自己数会和定位 / 应用
    时的 str.find 计数对不上——「权力要求书…权力要求」里第二处「权力要求」
    按 finditer 是第 1 次，按 find 却是第 2 次（第 1 次藏在「权力要求书」里）。
    """
    n = 0
    idx = text.find(needle)
    while 0 <= idx <= pos:
        n += 1
        idx = text.find(needle, idx + 1)
    return n


def apply_typo_corrections(paragraphs, corrections: list) -> int:
    """
    将用户确认的修正写回内存中的段落对象（不保存文件）。

    参数:
        paragraphs:  全文段落列表
        corrections: [{"para_idx": int, "wrong": str, "confirmed_fix": str,
                       "occurrence": int(可选，1 基)}, ...]
                     有 occurrence 时只改那一处；没有则改本段全部出现。
                     confirmed_fix 为空字符串时跳过该条。

    返回:
        实际替换的处数。
    """
    from collections import defaultdict
    by_para: dict = defaultdict(list)
    for item in corrections:
        fix = (item.get("confirmed_fix") or "").strip()
        wrong = item.get("wrong") or ""
        if not fix or not wrong or fix == wrong:
            continue
        by_para[item["para_idx"]].append((wrong, fix, item.get("occurrence")))

    count = 0
    for para_idx, items in by_para.items():
        para = paragraphs[para_idx]
        text = display_text(para)
        spans = []
        for wrong, fix, occ in items:
            if occ:
                pos = nth_occurrence(text, wrong, int(occ))
                if pos >= 0:
                    spans.append((pos, pos + len(wrong), fix))
            else:
                pos = text.find(wrong)
                while pos >= 0:
                    spans.append((pos, pos + len(wrong), fix))
                    pos = text.find(wrong, pos + len(wrong))
        # 区间重叠（两条结果指向交叠的字）时只保留靠前的一条
        spans.sort()
        kept, end = [], -1
        for s, e, f in spans:
            if s >= end:
                kept.append((s, e, f))
                end = e
        if not kept:
            continue
        new_text = text
        for s, e, f in reversed(kept):
            new_text = new_text[:s] + f + new_text[e:]
        if set_paragraph_text(para, new_text):
            count += len(kept)
    return count


def remap_after_edit(old_text: str, new_text: str, pos: int, old_len: int,
                     new_len: int, items: list) -> list:
    """同段一处 [pos, pos+old_len) 被改成 new_len 个字之后，重算其余结果的 occurrence。

    返回与 items 等长的列表：已失效（与改动区间交叠、或改后原位置已不是该词）
    的为 None，其余为 occurrence 更新后的副本。
    """
    out = []
    for it in items:
        w = it.get("wrong") or ""
        p = nth_occurrence(old_text, w, int(it.get("occurrence", 1) or 1))
        np = p if p < pos else p + new_len - old_len
        if (p < 0 or (p < pos + old_len and pos < p + len(w))
                or new_text[np:np + len(w)] != w):
            out.append(None)
        else:
            out.append({**it, "occurrence": occurrence_at(new_text, w, np)})
    return out
