# -*- coding: utf-8 -*-
"""
claim_check.py — 权利要求书引用检查

提供 6 项针对权利要求书的检查：
  1. 引用基础（antecedent basis）           — N 字滑窗
  2. 权利要求引用关系                         — 解析"根据权利要求X所述"
  3. 同一术语多种写法                         — N 字滑窗 + 相似度
  4. 多值/不确定用语                          — 内置词库
  5. 独立权利要求序号连续性                   — 正则扫编号
  6. 单引/多引合法性                          — 解析"或"结构

所有函数均为纯函数：输入权利要求书段落列表与参数，输出 list[dict]。
每条结果格式：
    {
        "kind":        "antecedent" / "dependency" / "term" /
                       "vague" / "numbering" / "multi_dep",
        "claim_no":    int | None,
        "para_idx":    int  (全文中的段落索引),
        "context":     str  (原文片段),
        "message":     str  (问题描述),
        "suggestion":  str  (改法提示，允许为空),
    }
"""
import re
from dataclasses import dataclass, field


# ─────────────────────────────────────────
# 内置"多值/不确定用语"词库
# ─────────────────────────────────────────
VAGUE_WORDBANK = [
    "大约", "大致", "大体", "大概", "约为",
    "左右", "上下", "前后", "附近",
    "优选", "优选地", "优选的", "优选为",
    "较佳", "较好", "最好",
    "基本", "基本上", "基本为",
    "一般", "通常", "往往",
    "可能", "或许", "也许",
    "少许", "少量", "多个",   # "多个"有时合法，但在权利要求中常被审查员抓
    "等等", "诸如",
]


# ─────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────
@dataclass
class ClaimInfo:
    no: int                            # 权利要求序号
    para_indices: list                 # 归属此权项的全文段落索引
    text: str                          # 拼接后的完整文本（去掉首部编号）
    raw_text: str                      # 含编号的原始文本
    cites: set = field(default_factory=set)  # 引用的其它权项序号
    cite_groups: list = field(default_factory=list)
    # cite_groups: [{"raw": "权利要求1或2", "nums": {1,2}, "mode": "or"/"and"/"single"}]
    is_independent: bool = True        # 是否独立权项（没有 cites 即为独立）


# ─────────────────────────────────────────
# 解析
# ─────────────────────────────────────────
_CLAIM_HEAD_RE = re.compile(r'^\s*(\d+)\s*[\.\．\、]\s*')
# 匹配"根据权利要求X所述" / "如权利要求X所述" / "按照权利要求X所述" 等
# 捕获紧跟的编号串（允许 "1", "1、2", "1或2", "1-3" 等）
_CITE_RE = re.compile(
    r'(?:根据|如|按照|依据)?权利要求\s*'
    r'([0-9０-９]+(?:\s*(?:[,，、和或至\-－~～]|到|或者)\s*[0-9０-９]+)*)'
    r'\s*(?:任一?项?)?\s*所述'
)


def _norm_digits(s: str) -> str:
    """全角数字 → 半角"""
    out = []
    for c in s:
        code = ord(c)
        if 0xFF10 <= code <= 0xFF19:
            out.append(chr(code - 0xFF10 + ord('0')))
        else:
            out.append(c)
    return "".join(out)


def _extract_cite_nums(num_str: str) -> tuple:
    """
    从形如 "1、2" / "1或2" / "1至3" 的字符串里提取出整数集合与模式。
    返回 (nums: set[int], mode: 'or'/'and'/'range'/'single')
    """
    s = _norm_digits(num_str)
    mode = "single"
    if "或" in s or "或者" in s:
        mode = "or"
    elif "、" in s or "," in s or "，" in s or "和" in s:
        mode = "and"
    elif "-" in s or "－" in s or "~" in s or "～" in s or "至" in s or "到" in s:
        mode = "range"
    # 处理范围
    nums = set()
    if mode == "range":
        m = re.match(r'\s*(\d+)\s*(?:[-－~～]|至|到)\s*(\d+)', s)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b:
                nums.update(range(a, b + 1))
            else:
                nums.update(range(b, a + 1))
            return nums, mode
    # 通用：抓所有数字
    for m in re.finditer(r'\d+', s):
        nums.add(int(m.group(0)))
    return nums, mode


def parse_claims(paragraphs, start_idx: int, end_idx: int) -> dict:
    """
    从全文 paragraphs 的 [start_idx, end_idx) 区间中解析权利要求。

    返回: {claim_no(int): ClaimInfo}
    """
    claims: dict = {}
    current_no = None
    current_paras: list = []
    current_lines: list = []

    def _flush():
        nonlocal current_no, current_paras, current_lines
        if current_no is None:
            return
        raw = "\n".join(current_lines).strip()
        # 去掉开头 "1." / "1、"
        stripped = _CLAIM_HEAD_RE.sub("", raw, count=1)
        info = ClaimInfo(
            no=current_no,
            para_indices=list(current_paras),
            text=stripped,
            raw_text=raw,
        )
        # 解析所有引用组
        for m in _CITE_RE.finditer(raw):
            nums, mode = _extract_cite_nums(m.group(1))
            if nums:
                info.cite_groups.append({
                    "raw": m.group(0),
                    "nums": nums,
                    "mode": mode,
                })
                info.cites.update(nums)
        info.is_independent = (len(info.cites) == 0)
        claims[current_no] = info
        current_no = None
        current_paras = []
        current_lines = []

    for i in range(start_idx, end_idx):
        if i < 0 or i >= len(paragraphs):
            continue
        text = paragraphs[i].text if paragraphs[i].text else ""
        if not text.strip():
            # 空段落：属于当前权项
            if current_no is not None:
                current_paras.append(i)
                current_lines.append(text)
            continue
        m = _CLAIM_HEAD_RE.match(text)
        if m:
            # 新权项开始
            _flush()
            try:
                current_no = int(_norm_digits(m.group(1)))
            except ValueError:
                current_no = None
            current_paras = [i]
            current_lines = [text]
        else:
            # 续行
            if current_no is not None:
                current_paras.append(i)
                current_lines.append(text)
    _flush()
    return claims


# ─────────────────────────────────────────
# 检查 1: 权利要求引用关系
# ─────────────────────────────────────────
def check_claim_dependency(claims: dict) -> list:
    results = []
    max_no = max(claims.keys()) if claims else 0
    for no in sorted(claims.keys()):
        info = claims[no]
        for grp in info.cite_groups:
            for cited in sorted(grp["nums"]):
                msg = None
                if cited == no:
                    msg = f"权利要求 {no} 自引（引用了自身）"
                elif cited > no:
                    msg = f"权利要求 {no} 非法前引：引用了序号更大的权利要求 {cited}"
                elif cited not in claims:
                    msg = f"权利要求 {no} 引用了不存在的权利要求 {cited}（最大序号为 {max_no}）"
                if msg:
                    results.append({
                        "kind": "dependency",
                        "claim_no": no,
                        "para_idx": info.para_indices[0] if info.para_indices else -1,
                        "context": grp["raw"],
                        "message": msg,
                        "suggestion": "",
                    })
    return results


# ─────────────────────────────────────────
# 检查 2: 单引/多引合法性
# ─────────────────────────────────────────
def check_multi_dependency(claims: dict) -> list:
    """
    常见规则：
      - 多项引用只能用"或"，不能用"和/及/与"来并列（因为"和"会被解读为同时满足）
      - "权利要求1-3任一项所述" 合法；"权利要求1和2所述" 不合法
    """
    results = []
    for no in sorted(claims.keys()):
        info = claims[no]
        for grp in info.cite_groups:
            if len(grp["nums"]) <= 1:
                continue
            mode = grp["mode"]
            if mode == "and":
                results.append({
                    "kind": "multi_dep",
                    "claim_no": no,
                    "para_idx": info.para_indices[0] if info.para_indices else -1,
                    "context": grp["raw"],
                    "message": f"权利要求 {no} 的多项引用使用了'和/、'连接，应改为'或'",
                    "suggestion": grp["raw"].replace("和", "或").replace("、", "或"),
                })
    return results


# ─────────────────────────────────────────
# 检查 3: 独立权利要求序号连续性
# ─────────────────────────────────────────
def check_claim_numbering(claims: dict) -> list:
    results = []
    if not claims:
        return results
    nums = sorted(claims.keys())
    expected = list(range(1, len(nums) + 1))
    if nums != expected:
        # 找出缺号/重号/起始不为1等
        msgs = []
        if nums[0] != 1:
            msgs.append(f"起始序号为 {nums[0]}，应从 1 开始")
        missing = [x for x in range(nums[0], nums[-1] + 1) if x not in nums]
        if missing:
            msgs.append(f"缺失序号：{missing}")
        if msgs:
            first_no = nums[0]
            info = claims[first_no]
            results.append({
                "kind": "numbering",
                "claim_no": None,
                "para_idx": info.para_indices[0] if info.para_indices else -1,
                "context": "、".join(str(x) for x in nums),
                "message": "权项序号不连续：" + "；".join(msgs),
                "suggestion": "",
            })
    return results


# ─────────────────────────────────────────
# 检查 4: 多值/不确定用语
# ─────────────────────────────────────────
def check_vague_terms(claims: dict, vague_words=None) -> list:
    results = []
    vague_words = list(vague_words or VAGUE_WORDBANK)
    for no in sorted(claims.keys()):
        info = claims[no]
        text = info.text
        for w in vague_words:
            if not w:
                continue
            pos = 0
            while True:
                idx = text.find(w, pos)
                if idx < 0:
                    break
                start = max(0, idx - 12)
                end = min(len(text), idx + len(w) + 12)
                ctx = text[start:end].replace("\n", " ")
                results.append({
                    "kind": "vague",
                    "claim_no": no,
                    "para_idx": info.para_indices[0] if info.para_indices else -1,
                    "context": ctx,
                    "message": f"权利要求 {no} 含不确定用语『{w}』",
                    "suggestion": "",
                })
                pos = idx + len(w)
    return results


# ─────────────────────────────────────────
# N 字滑窗工具
# ─────────────────────────────────────────
_CJK_RE = re.compile(r'[\u4e00-\u9fff]')

# 常见中文语法虚词 / 高频停用字：包含这些字的 ngram 大概率不是技术术语
_STOPCHARS = set(
    "的地得之了所在是为而与和或及若如也有就都亦且其此于对从被将把使以至于"
    "上下中内外前后左右里间边旁侧面处"
    "一二三四五六七八九十百千万两几每各种个件只次项条"
    "并其余他她它们我你您大小多少些第即已还又则但故只仅"
)


def _is_noisy_ngram(seg: str) -> bool:
    """判断 ngram 是否为噪声（含停用字）"""
    return any(ch in _STOPCHARS for ch in seg)


def _sliding_cjk_ngrams(text: str, n: int, skip_noise: bool = False) -> list:
    """
    从文本中提取所有长度为 n 的"全中文"子串。
    非中文字符视为分隔符（不产生包含它们的子串）。
    返回 (ngram, start_idx) 的列表。

    skip_noise=True 时过滤掉含停用字的 ngram（用于术语提取以降噪）。
    """
    out = []
    L = len(text)
    for i in range(L - n + 1):
        seg = text[i:i + n]
        if not all(_CJK_RE.match(ch) for ch in seg):
            continue
        if skip_noise and _is_noisy_ngram(seg):
            continue
        out.append((seg, i))
    return out


# 用于识别"X所述"是否为"权利要求N所述"的引用语：若"所述"前 12 字内
# 出现"权利要求\d+"，视为引用语公式，本处的"所述"不参与引用基础检查。
_CLAIM_CITE_PREFIX_RE = re.compile(r'权利要求\s*\d')


def _is_in_citation_formula(text: str, suoshu_pos: int) -> bool:
    lookback_start = max(0, suoshu_pos - 12)
    return bool(_CLAIM_CITE_PREFIX_RE.search(text[lookback_start:suoshu_pos]))


# ─────────────────────────────────────────
# 检查 5: 引用基础（antecedent basis）
# ─────────────────────────────────────────
def check_antecedent_basis(claims: dict, n: int, ignore_set: set) -> list:
    """
    原则：以"所述X"形式出现的 n 字子串 X，必须在之前（同权项内或该权项所引用
    的更早权项中）以"非所述"方式出现过一次（视为首次定义）。

    注：n 字滑窗自然会产生噪声，配合 ignore_set 过滤。
    """
    results = []
    ignore_set = set(ignore_set or ())
    # 先给每个权项建立"已定义的 n 字术语集合"（不包含所述前缀）
    # 遍历时，按权项从小到大，继承该权项所依赖的权项的定义集
    defined_by_claim: dict = {}
    for no in sorted(claims.keys()):
        info = claims[no]
        # 起始集合 = 所有被其引用的前序权项的 defined 集合的并集
        base: set = set()
        for cited in info.cites:
            if cited in defined_by_claim:
                base |= defined_by_claim[cited]
        # 遍历文本，识别"非所述的 n 字 CJK 子串"作为首次定义
        text = info.text
        cur_defined = set(base)
        # 过滤掉"权利要求N所述"中的"所述"（属于引用语公式，不是反向引用）
        suoshu_positions = [
            m.start() for m in re.finditer(r'所述', text)
            if not _is_in_citation_formula(text, m.start())
        ]
        suoshu_span = set()
        for p in suoshu_positions:
            # 认为"所述"后紧跟的 n 个 CJK 字作为被引用的术语
            for k in range(p + 2, min(p + 2 + n, len(text))):
                suoshu_span.add(k)
        # 先扫一遍"非所述"上下文中的 n 字 CJK 子串 → 记入定义集
        # 此处用 skip_noise=True 过滤掉含"的/在/是"等停用字的子串
        for seg, idx in _sliding_cjk_ngrams(text, n, skip_noise=True):
            if seg in ignore_set:
                continue
            # 如果该子串位于"所述 + seg"的窗口内 → 视为引用，不当作首次定义
            range_positions = set(range(idx, idx + n))
            if range_positions & suoshu_span:
                continue
            cur_defined.add(seg)
        # 再扫一遍"所述 + n 字子串"，检查是否在 cur_defined 里
        # 注：此检查使用已经包含本段首次定义的集合 —— 合理：在同一权项中
        # 先定义后"所述"是合规的。
        for p in suoshu_positions:
            # 抓取紧跟"所述"的 n 个 CJK 字
            term_chars = []
            for k in range(p + 2, len(text)):
                ch = text[k]
                if _CJK_RE.match(ch):
                    term_chars.append(ch)
                    if len(term_chars) == n:
                        break
                else:
                    break
            if len(term_chars) < n:
                continue
            term = "".join(term_chars)
            if term in ignore_set:
                continue
            # 若"所述"后紧跟的 n 字本身含停用字（如"第一/两侧/所述"），
            # 视为虚词序列而非技术术语，不参与引用基础检查。
            if _is_noisy_ngram(term):
                continue
            if term not in cur_defined:
                start = max(0, p - 8)
                end = min(len(text), p + 2 + n + 8)
                ctx = text[start:end].replace("\n", " ")
                results.append({
                    "kind": "antecedent",
                    "claim_no": no,
                    "para_idx": info.para_indices[0] if info.para_indices else -1,
                    "context": ctx,
                    "message": (
                        f"权利要求 {no} 中『所述{term}』缺少引用基础"
                        f"（在本权项及所引用的前序权项中未找到对『{term}』的首次定义）"
                    ),
                    "suggestion": "",
                })
        # 写回
        defined_by_claim[no] = cur_defined
    return results


# ─────────────────────────────────────────
# 检查 6: 同一术语多种写法
# ─────────────────────────────────────────
def _similar(a: str, b: str) -> bool:
    """简易相似：长度相同且仅 1 字不同；或其中一个是另一个的包含串。"""
    if a == b:
        return False
    if len(a) == len(b):
        diff = sum(1 for x, y in zip(a, b) if x != y)
        return diff == 1
    return False


def check_term_consistency(claims: dict, n: int, ignore_set: set) -> list:
    """
    只收集「所述」后紧跟的 n 字 CJK 术语，在权利要求书范围内找出相似对
    （长度相同且仅一字之差），作为"同一术语多种写法"疑点上报。

    之所以只看「所述」后面：
      • 这是权利要求书中"反向引用"的标准句式，"所述X"里的 X 必然是一个
        已定义的术语，天然排除了 n 字滑窗扫全文带来的大量噪声；
      • 代理人撰写时最常见的术语漂移（如"齿圈"/"齿环"）就发生在"所述"后。
    """
    ignore_set = set(ignore_set or ())
    # term -> {count, first: (claim_no, para_idx, context)}
    term_locs: dict = {}

    for no in sorted(claims.keys()):
        info = claims[no]
        text = info.text
        # 找所有"所述"且不在"权利要求N所述"引用公式里
        suoshu_positions = [
            m.start() for m in re.finditer(r'所述', text)
            if not _is_in_citation_formula(text, m.start())
        ]
        for p in suoshu_positions:
            # 紧跟"所述"的 n 个 CJK 字组成术语
            term_chars = []
            for k in range(p + 2, len(text)):
                ch = text[k]
                if _CJK_RE.match(ch):
                    term_chars.append(ch)
                    if len(term_chars) == n:
                        break
                else:
                    break
            if len(term_chars) < n:
                continue
            term = "".join(term_chars)
            if term in ignore_set:
                continue
            # 含停用字（如"第一/两侧"）视为虚词序列，不参与术语一致性
            if _is_noisy_ngram(term):
                continue

            if term not in term_locs:
                start = max(0, p - 8)
                end = min(len(text), p + 2 + n + 8)
                ctx = text[start:end].replace("\n", " ")
                term_locs[term] = {
                    "count": 1,
                    "first": (no, info.para_indices[0] if info.para_indices else -1, ctx),
                }
            else:
                term_locs[term]["count"] += 1

    results = []
    seen_pairs = set()
    terms = list(term_locs.keys())
    for i, a in enumerate(terms):
        for b in terms[i + 1:]:
            if _similar(a, b):
                pair = (a, b) if a < b else (b, a)
                if pair in seen_pairs:
                    continue
                seen_pairs.add(pair)
                loc = term_locs[a]["first"]
                results.append({
                    "kind": "term",
                    "claim_no": loc[0],
                    "para_idx": loc[1],
                    "context": loc[2],
                    "message": (
                        f"发现相似术语『{a}』与『{b}』"
                        f"（{a}:{term_locs[a]['count']}次 · {b}:{term_locs[b]['count']}次），"
                        f"疑似同一术语多种写法"
                    ),
                    "suggestion": "",
                })
    return results


# ─────────────────────────────────────────
# 聚合入口
# ─────────────────────────────────────────
def run_all_checks(paragraphs, start_idx: int, end_idx: int,
                   n: int, ignore_set=None, vague_words=None) -> list:
    """
    一次性运行全部 6 项检查，返回合并后的结果列表。

    参数:
        paragraphs:   全文段落列表（python-docx Paragraph 对象）
        start_idx, end_idx: 权利要求书段落区间
        n:            术语类检查的滑窗字数（2~6）
        ignore_set:   用户自定义忽略词集合（仅作用于术语类检查）
        vague_words:  覆盖默认 VAGUE_WORDBANK
    """
    claims = parse_claims(paragraphs, start_idx, end_idx)
    if not claims:
        return []

    results = []
    results.extend(check_claim_dependency(claims))
    results.extend(check_multi_dependency(claims))
    results.extend(check_claim_numbering(claims))
    results.extend(check_vague_terms(claims, vague_words))
    results.extend(check_antecedent_basis(claims, n, ignore_set or set()))
    results.extend(check_term_consistency(claims, n, ignore_set or set()))

    # 按 claim_no、para_idx 排序以稳定输出
    def sort_key(r):
        return (r.get("claim_no") or 0, r.get("para_idx") or 0, r.get("kind") or "")
    results.sort(key=sort_key)
    return results


# ─────────────────────────────────────────
# 段落写回辅助（供 main_window 使用）
# ─────────────────────────────────────────
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
