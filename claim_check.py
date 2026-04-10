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
# 内置"动态截断黑名单"词库
# ─────────────────────────────────────────
# 在"所述 X"中，X 的真实边界通常是这些"词性词"（动词/方位词/虚词）。
# 动态截断时，从"所述"往后扫，遇到标点或任一黑名单词的首字时立刻停下，
# 把之前累积的 CJK 字当作术语返回。
DEFAULT_BOUNDARY_BLACKLIST = [
    # ── 动词类（结构性动词，常出现在"所述X + 动词"里）──
    "安装", "连接", "设置", "设于", "设在", "位于", "固定", "固连",
    "套设", "套装", "套接", "套在", "插入", "插设", "嵌入", "嵌设",
    "抵接", "抵靠", "贴合", "贴附", "贴设", "焊接", "铰接", "铰连",
    "粘接", "粘贴", "粘连", "卡接", "卡设", "卡在", "卡合", "啮合",
    "紧固", "紧贴", "挤压", "按压", "压接", "压紧", "压合", "螺接",
    "螺纹", "铆接", "铆合", "钩接", "钩挂", "悬挂", "吊装",
    "包括", "包含", "包围", "围绕", "环绕", "环设", "环抱",
    "用于", "以便", "以使", "使得", "能够", "可以",
    "穿过", "穿设", "穿出", "贯穿", "贯通", "通过",
    "朝向", "面向", "背向", "指向", "延伸", "伸出", "伸入",
    "形成", "构成", "组成", "具有", "带有", "设有",
    # ── 方位词类（"所述X + 的上/下/内/外"）──
    "上方", "下方", "上部", "下部", "上端", "下端", "上侧", "下侧",
    "上表", "下表", "顶部", "底部", "顶端", "底端", "顶面", "底面",
    "内部", "外部", "内侧", "外侧", "内端", "外端", "内壁", "外壁",
    "前方", "后方", "前部", "后部", "前端", "后端", "前侧", "后侧",
    "左方", "右方", "左部", "右部", "左端", "右端", "左侧", "右侧",
    "中部", "中间", "中央", "中心", "周侧", "周缘", "周向", "径向",
    "轴向", "端部", "端面", "一端", "另一", "两端", "两侧", "两者",
    # ── 连词/助词/介词类 ──
    "与其", "和其", "及其", "或者",
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
# 动态截断 / 动态回退 辅助
# ─────────────────────────────────────────
def _extract_term_dynamic_truncate(text: str, start: int,
                                    blacklist_first_chars: set,
                                    max_len: int = 12) -> str:
    """
    从 text[start:] 起向后扫描 CJK 字符，直到遇到下列任一情况就停下：
      • 非 CJK 字符（标点、空格、英文、数字…）
      • 黑名单词的首字（即位置 k 起 text[k:k+L] 在黑名单里）
      • 累计达到 max_len（保护用，防止整段被吞）
    返回累积到的术语字符串（可能为空）。
    """
    out = []
    k = start
    L = len(text)
    while k < L and len(out) < max_len:
        ch = text[k]
        if not _CJK_RE.match(ch):
            break
        # 检查从位置 k 起是否命中黑名单的某个词
        # 黑名单首字命中是必要条件，再做一次完整匹配以避免误伤
        if ch in blacklist_first_chars:
            hit = False
            # 尝试匹配 2~4 字的黑名单词（黑名单词长度大多是 2，少数 3~4）
            for L_word in (2, 3, 4):
                if k + L_word <= L and text[k:k + L_word] in _BLACKLIST_WORDS_CACHE:
                    hit = True
                    break
            if hit:
                break
        out.append(ch)
        k += 1
    return "".join(out)


# 全局缓存：把当前调用使用的黑名单词集合放进去，加速 _extract_term_dynamic_truncate
_BLACKLIST_WORDS_CACHE: set = set()


def _build_blacklist_lookup(blacklist) -> tuple:
    """从词列表构造 (首字 set, 完整词 set)"""
    words = {w for w in (blacklist or ()) if w}
    first_chars = {w[0] for w in words if w}
    return first_chars, words


# ─────────────────────────────────────────
# 检查 5: 引用基础（antecedent basis）
# ─────────────────────────────────────────
def check_antecedent_basis(claims: dict, n: int, ignore_set: set,
                            use_dynamic_truncate: bool = False,
                            use_dynamic_fallback: bool = False,
                            boundary_blacklist=None) -> list:
    """
    原则：以"所述X"形式出现的术语 X，必须在之前（同权项内或该权项所引用的
    更早权项中）以"非所述"方式出现过一次（视为首次定义）。

    术语 X 的提取策略：
      • 默认（两个开关都关）：紧跟"所述"后取 n 个 CJK 字
      • 仅 use_dynamic_truncate=True：从"所述"往后扫，遇到标点 / 黑名单词
        立刻停下，把累积的 CJK 字串作为术语（自适应长度）
      • 仅 use_dynamic_fallback=True：取 n 字后，若不在已定义集中，则不断
        把末尾砍掉一个字，重试，直到匹配到或缩到 1 字仍不匹配才报错
      • 两个都开：先用截断得到一个"干净"的最长术语，再对该术语应用回退
        （从右向左缩短）。仅当所有前缀都没匹配时才报错——这是误判最低的组合

    注：滑窗式定义集仍然按 n 字累积；回退/截断只影响"所述"侧的术语形态。
    """
    results = []
    ignore_set = set(ignore_set or ())
    # 准备黑名单查找表（即便没启用截断也无副作用）
    bl_first_chars, bl_words = _build_blacklist_lookup(boundary_blacklist or [])
    global _BLACKLIST_WORDS_CACHE
    _BLACKLIST_WORDS_CACHE = bl_words
    dyn_mode = use_dynamic_truncate or use_dynamic_fallback
    # 动态模式下，术语长度可变；用更宽松的最大长度做候选
    DYN_MAX_LEN = 12

    def _collect_freeform_terms(text: str, suoshu_span: set) -> set:
        """
        把文本拆成"连续 CJK 段"，对每段产生所有长度 2..DYN_MAX_LEN 的子串
        作为"已定义术语候选"。位于 suoshu_span 内的位置视为引用，不参与定义。
        """
        out: set = set()
        L = len(text)
        i = 0
        while i < L:
            if not _CJK_RE.match(text[i]):
                i += 1
                continue
            # 找到一段连续 CJK
            j = i
            while j < L and _CJK_RE.match(text[j]):
                j += 1
            # 对这段 [i, j) 抽所有长度 2..DYN_MAX_LEN 的子串，
            # 但需要剔除与 suoshu_span 完全重叠的子串
            seg_len = j - i
            for L_sub in range(2, min(DYN_MAX_LEN, seg_len) + 1):
                for k in range(i, j - L_sub + 1):
                    if any((pos in suoshu_span) for pos in range(k, k + L_sub)):
                        continue
                    sub = text[k:k + L_sub]
                    out.add(sub)
            i = j
        return out

    # 先给每个权项建立"已定义的 n 字术语集合"（不包含所述前缀）
    # 遍历时，按权项从小到大，继承该权项所依赖的权项的定义集
    defined_by_claim: dict = {}
    defined_free_by_claim: dict = {}  # 动态模式下使用
    for no in sorted(claims.keys()):
        info = claims[no]
        # 起始集合 = 所有被其引用的前序权项的 defined 集合的并集
        base: set = set()
        base_free: set = set()
        for cited in info.cites:
            if cited in defined_by_claim:
                base |= defined_by_claim[cited]
            if cited in defined_free_by_claim:
                base_free |= defined_free_by_claim[cited]
        # 遍历文本，识别"非所述的 n 字 CJK 子串"作为首次定义
        text = info.text
        cur_defined = set(base)
        cur_defined_free = set(base_free)
        # 过滤掉"权利要求N所述"中的"所述"（属于引用语公式，不是反向引用）
        suoshu_positions = [
            m.start() for m in re.finditer(r'所述', text)
            if not _is_in_citation_formula(text, m.start())
        ]
        # 引用区窗口：
        #   • 默认（无动态模式）  → n 字
        #   • 仅 use_dynamic_fallback → DYN_MAX_LEN 字（回退会试 n..2 字所有前缀）
        #   • 启用 use_dynamic_truncate → 用同一套截断逻辑算出的实际术语长度
        #     这避免了「12 字窗口把后面真正定义的术语也吞掉」的问题
        #     例：「所述垂直延伸板段的端部设有挂钩（10），所述挂钩…」
        #         如果窗口是固定 12 字，会把"挂"扣掉，导致"挂钩"收不进定义集；
        #         用截断逻辑算出真实长度 7（停在"端部"），"挂钩"就能正常入集。
        suoshu_span = set()
        for p in suoshu_positions:
            if use_dynamic_truncate:
                ref_term = _extract_term_dynamic_truncate(
                    text, p + 2, bl_first_chars, max_len=DYN_MAX_LEN
                )
                ref_len = len(ref_term) if ref_term else n
            else:
                ref_len = DYN_MAX_LEN if dyn_mode else n
            for k in range(p + 2, min(p + 2 + ref_len, len(text))):
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
        # 动态模式：额外收集变长子串作为"自由形式"定义集
        if dyn_mode:
            cur_defined_free |= _collect_freeform_terms(text, suoshu_span)

        # ── 提取每个 "所述" 后的术语并校验 ──
        for p in suoshu_positions:
            # 1) 默认 n 字术语（用于既不开截断也不开回退、或回退单独使用时）
            term_chars = []
            for k in range(p + 2, len(text)):
                ch = text[k]
                if _CJK_RE.match(ch):
                    term_chars.append(ch)
                    if len(term_chars) == n:
                        break
                else:
                    break
            n_term = "".join(term_chars) if len(term_chars) == n else ""

            # 2) 动态截断术语（只要开了截断就要算）
            trunc_term = ""
            if use_dynamic_truncate:
                trunc_term = _extract_term_dynamic_truncate(
                    text, p + 2, bl_first_chars, max_len=DYN_MAX_LEN
                )

            # 决定本次"所述"的报告策略
            if not use_dynamic_truncate and not use_dynamic_fallback:
                # 默认：n 字定值
                if not n_term:
                    continue
                if n_term in ignore_set or _is_noisy_ngram(n_term):
                    continue
                if n_term in cur_defined:
                    continue
                missing_term = n_term
                ctx_term_len = n
            elif use_dynamic_truncate and not use_dynamic_fallback:
                # 仅截断：以截断结果为准
                if not trunc_term or len(trunc_term) < 2:
                    continue
                if trunc_term in ignore_set or _is_noisy_ngram(trunc_term):
                    continue
                if (trunc_term in cur_defined_free) or (trunc_term in cur_defined):
                    continue
                missing_term = trunc_term
                ctx_term_len = len(trunc_term)
            elif use_dynamic_fallback and not use_dynamic_truncate:
                # 仅回退：从 n 字开始往下缩
                if not n_term:
                    continue
                if n_term in ignore_set or _is_noisy_ngram(n_term):
                    continue
                hit = False
                for L_try in range(n, 1, -1):
                    sub = n_term[:L_try]
                    if sub in ignore_set:
                        hit = True
                        break
                    if (sub in cur_defined) or (sub in cur_defined_free):
                        hit = True
                        break
                if hit:
                    continue
                missing_term = n_term
                ctx_term_len = n
            else:
                # 同时开启：先截断，再回退；只要任一前缀匹配就放过
                # 这是误判最低的组合策略
                base_term = trunc_term
                if not base_term or len(base_term) < 2:
                    # 截断失败时退化为 n 字
                    base_term = n_term
                if not base_term or len(base_term) < 2:
                    continue
                if base_term in ignore_set or _is_noisy_ngram(base_term):
                    continue
                hit = False
                for L_try in range(len(base_term), 1, -1):
                    sub = base_term[:L_try]
                    if sub in ignore_set:
                        hit = True
                        break
                    if (sub in cur_defined) or (sub in cur_defined_free):
                        hit = True
                        break
                if hit:
                    continue
                # 还要再多一道兜底：base_term 的最末字往往是边界字，
                # 单独再用 n_term（n 字定值）兜一遍可避免漏放过同义噪声
                if n_term and (
                    n_term in cur_defined or n_term in cur_defined_free
                ):
                    continue
                missing_term = base_term
                ctx_term_len = len(base_term)

            start = max(0, p - 8)
            end = min(len(text), p + 2 + ctx_term_len + 8)
            ctx = text[start:end].replace("\n", " ")
            results.append({
                "kind": "antecedent",
                "claim_no": no,
                "para_idx": info.para_indices[0] if info.para_indices else -1,
                "context": ctx,
                "message": f"『所述{missing_term}』缺少引用基础",
                "suggestion": "",
            })
        # 写回
        defined_by_claim[no] = cur_defined
        defined_free_by_claim[no] = cur_defined_free
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
                   n: int, ignore_set=None, vague_words=None,
                   check_term: bool = False,
                   check_vague: bool = True,
                   use_dynamic_truncate: bool = False,
                   use_dynamic_fallback: bool = False,
                   boundary_blacklist=None) -> list:
    """
    一次性运行全部检查，返回合并后的结果列表。

    参数:
        paragraphs:   全文段落列表（python-docx Paragraph 对象）
        start_idx, end_idx: 权利要求书段落区间
        n:            术语类检查的滑窗字数（2~6）
        ignore_set:   用户自定义忽略词集合（仅作用于术语类检查）
        vague_words:  覆盖默认 VAGUE_WORDBANK
        check_term:   是否执行「术语不一致」检查；默认 False，因为该检查
                      噪音较大，只有用户在 UI 中明确勾选时才会运行
        use_dynamic_truncate / use_dynamic_fallback / boundary_blacklist:
                      引用基础检查的两个降噪开关，详见 check_antecedent_basis
    """
    claims = parse_claims(paragraphs, start_idx, end_idx)
    if not claims:
        return []

    results = []
    results.extend(check_claim_dependency(claims))
    results.extend(check_multi_dependency(claims))
    results.extend(check_claim_numbering(claims))
    if check_vague:
        results.extend(check_vague_terms(claims, vague_words))
    results.extend(check_antecedent_basis(
        claims, n, ignore_set or set(),
        use_dynamic_truncate=use_dynamic_truncate,
        use_dynamic_fallback=use_dynamic_fallback,
        boundary_blacklist=boundary_blacklist,
    ))
    if check_term:
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
