# -*- coding: utf-8 -*-
"""
claim_check.py — 权利要求书引用检查

提供 6 项针对权利要求书的检查：
  1. 引用基础（antecedent basis）           — N 字滑窗（可叠加动态截断 / 动态回退）
  2. 权利要求引用关系                         — 解析"根据权利要求X所述"
  3. 多值/不确定用语                          — 内置词库
  4. 权项序号连续性 / 重号                     — 正则扫编号
  5. 单引/多引合法性                          — 解析"或"结构
  6. 每条权项以「。」结尾                      — 撰写规范

所有函数均为纯函数：输入权利要求书段落列表与参数，输出 list[dict]。
每条结果格式：
    {
        "kind":        "antecedent" / "dependency" /
                       "vague" / "numbering" / "multi_dep" / "ending",
        "claim_no":    int | None,
        "para_idx":    int  (全文中的段落索引),
        "context":     str  (原文片段),
        "message":     str  (问题描述),
        "suggestion":  str  (改法提示，允许为空),
    }
"""
import re
from dataclasses import dataclass, field
from types import SimpleNamespace

from core.mark_extractor import strip_reference_marks


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
    # ── 方法类权项的高频动词（"所述X + 动词"，装置权项黑名单覆盖不到）──
    # 只收"几乎不构成部件名前缀"的词：像 采集 / 控制 / 驱动 / 计算 这类
    # 同时能组成 采集单元 / 控制器 / 驱动电机 / 计算机 的词**故意不收**，
    # 收了会把「所述光谱采集单元」截成「光谱」。
    "进行", "设定", "发出", "具体", "简化", "转换", "施加",
    "判断", "得到", "转动", "配置", "调制", "遮挡",
    "乘以", "除以", "作用于", "作用下",
    # ── 连词/助词/介词类 ──
    "与其", "和其", "及其", "或者", "分别",
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
# 捕获紧跟的编号串（允许 "1", "1、2", "1或2", "1-3", "1或权利要求2" 等；
# 分隔符后允许重复出现"权利要求"字样，使"权利要求1或权利要求2所述"
# 解析为一个引用组——否则会拆成两个 single 组，多项引用规则全部失效）
_CITE_RE = re.compile(
    r'(?:根据|如|按照|依据)?权利要求\s*'
    r'([0-9０-９]+(?:\s*(?:[,，、和或至\-－~～]|到|或者)\s*(?:权利要求)?\s*[0-9０-９]+)*)'
    r'\s*中?\s*的?\s*(?:任(?:意)?一?项?|之一)?\s*所述'
)
_RANGE_RE = re.compile(r'(\d+)\s*(?:[-－~～]|至|到)\s*(\d+)')
_NUM_RE = re.compile(r'\d+')
_SUOSHU_RE = re.compile(r'所述')

# 动态截断时术语最大保留字符数，超过视为冗余描述
DYN_TERM_MAX_LEN = 12

# 动态截断的"合理术语长度"软上限。超过它仍没遇到边界，说明黑名单没覆盖住这段
# 文字里的动词——方法类权项尤其常见（简化为 / 转换为 / 施加在 / 进行校核 …），
# 截出的长串几乎不可能在前文原样出现，直接报"缺少引用基础"就是纯误判。
# 仅「动态截断」单开时用它做溢出判定并回落到 n 字定值术语；
# 「截断 + 回退」同开时由前缀回退自行消化长串，不受此上限影响。
DYN_TRUNC_SOFT_MAX = 6


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
    从形如 "1、2" / "1或2" / "1至3" / "1、3-5" / "1至3或5" 的字符串里
    提取出整数集合与模式。
    返回 (nums: set[int], mode: 'or'/'and'/'range'/'single')

    范围区段（a至b / a-b）无论与何种连接词混用都会展开为完整区间，
    避免 "1、3-5" 只取到 {1,3,5} 漏掉 4。
    """
    s = _norm_digits(num_str)
    mode = "single"
    if "或" in s or "或者" in s:
        mode = "or"
    elif "、" in s or "," in s or "，" in s or "和" in s:
        mode = "and"
    elif "-" in s or "－" in s or "~" in s or "～" in s or "至" in s or "到" in s:
        mode = "range"
    nums = set()
    # 先展开所有 "a至b" / "a-b" 范围区段（与 mode 无关）
    for m in _RANGE_RE.finditer(s):
        a, b = int(m.group(1)), int(m.group(2))
        lo, hi = (a, b) if a <= b else (b, a)
        nums.update(range(lo, hi + 1))
    # 再补散号（范围端点重复加入 set，无害）
    for m in _NUM_RE.finditer(s):
        nums.add(int(m.group(0)))
    return nums, mode


def parse_claims_ex(paragraphs, start_idx: int, end_idx: int) -> tuple:
    """
    从全文 paragraphs 的 [start_idx, end_idx) 区间中解析权利要求。

    返回: (claims, duplicates)
        claims:     {claim_no(int): ClaimInfo}（重号时保留首次出现的权项）
        duplicates: [{"no": int, "para_idx": int, "context": str}]
                    —— 序号与已有权项重复的后续权项
    """
    claims: dict = {}
    duplicates: list = []
    current_no = None
    current_paras: list = []
    current_lines: list = []

    def _flush():
        nonlocal current_no, current_paras, current_lines
        if current_no is None:
            return
        raw = "\n".join(current_lines).strip()
        # 重号：dict 直接赋值会静默覆盖前一条，导致重号查不出来；
        # 这里保留首条，把后续同号权项记入 duplicates 供编号检查上报
        if current_no in claims:
            duplicates.append({
                "no": current_no,
                "para_idx": current_paras[0] if current_paras else -1,
                "context": raw.replace("\n", " ")[:30],
            })
            current_no = None
            current_paras = []
            current_lines = []
            return
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
    return claims, duplicates


def parse_claims(paragraphs, start_idx: int, end_idx: int) -> dict:
    """兼容包装：只返回 {claim_no: ClaimInfo}（重号检测请用 parse_claims_ex）。"""
    claims, _ = parse_claims_ex(paragraphs, start_idx, end_idx)
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
    多项引用合法性检查：
      1. 多项引用只能用"或"，不能用"和/及/与"来并列（"和"会被解读为同时满足）；
         "权利要求1-3任一项所述" 合法；"权利要求1和2所述" 不合法
      2. 多引多（实施细则第 25 条第 2 款，2023 修订前为第 22 条第 2 款）：多项引用权利要求不得作为另一项
         多项引用权利要求的引用基础
      3. 多项引用建议写明「中任一项」，否则保护范围表述不清，常被审查员指出
    """
    results = []
    # 多项引用权项集合：任一引用组包含 2 个以上编号即视为多项引用权利要求
    multi_dep_nos = {
        no for no, info in claims.items()
        if any(len(grp["nums"]) > 1 for grp in info.cite_groups)
    }
    for no in sorted(claims.keys()):
        info = claims[no]
        for grp in info.cite_groups:
            if len(grp["nums"]) <= 1:
                continue
            para_idx = info.para_indices[0] if info.para_indices else -1
            mode = grp["mode"]
            if mode == "and":
                results.append({
                    "kind": "multi_dep",
                    "claim_no": no,
                    "para_idx": para_idx,
                    "context": grp["raw"],
                    "message": f"权利要求 {no} 的多项引用使用了'和/、'连接，应改为'或'",
                    "suggestion": grp["raw"].replace("和", "或").replace("、", "或"),
                })
            elif not any(k in grp["raw"] for k in ("任一", "任意", "之一")):
                # 多项引用缺「任一项」（and 模式已在上面单独报，不重复提示）
                results.append({
                    "kind": "multi_dep",
                    "claim_no": no,
                    "para_idx": para_idx,
                    "context": grp["raw"],
                    "message": f"权利要求 {no} 的多项引用未写明『中任一项』，保护范围表述不清",
                    "suggestion": grp["raw"].replace("所述", "中任一项所述", 1),
                })
            # 多引多：本组引用的权项中存在另一个多项引用权项
            hit_multi = sorted(grp["nums"] & multi_dep_nos)
            if hit_multi:
                hit_str = "、".join(str(x) for x in hit_multi)
                results.append({
                    "kind": "multi_dep",
                    "claim_no": no,
                    "para_idx": para_idx,
                    "context": grp["raw"],
                    "message": (
                        f"权利要求 {no} 多项引用了多项引用权利要求 {hit_str}"
                        f"（多引多，不符合专利法实施细则第二十五条第二款）"
                    ),
                    "suggestion": "改写被引权项为单项引用，或拆分本权项的引用关系",
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
    """扫描不确定用语。

    词库里存在包含 / 交叠关系的词条（优选 ⊂ 优选地、基本 ⊂ 基本上、大约 ∩ 约为），
    逐词独立 find 会把同一处问题重复上报 2~3 条。这里改为「长词优先 + 位置遮罩」：
    已被某个词占用的字符区间不再被更短 / 交叠的词命中，一处问题只报一条。
    排序次键取词面，保证结果顺序稳定（set 迭代序不可依赖）。
    """
    results = []
    words = sorted(
        {w for w in (vague_words or VAGUE_WORDBANK) if w},
        key=lambda w: (-len(w), w),
    )
    for no in sorted(claims.keys()):
        info = claims[no]
        text = info.text
        taken = [False] * len(text)
        hits = []
        for w in words:
            pos = 0
            while True:
                idx = text.find(w, pos)
                if idx < 0:
                    break
                if not any(taken[idx:idx + len(w)]):
                    for k in range(idx, idx + len(w)):
                        taken[k] = True
                    hits.append((idx, w))
                pos = idx + len(w)
        for idx, w in sorted(hits):
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
    """判断定长 ngram 是否为噪声（含任一停用字即丢弃）。

    只适用于「定长滑窗」术语：长度固定为 n，夹到停用字的多半确实不是术语。
    """
    return any(ch in _STOPCHARS for ch in seg)


# 方位 / 大小字打头时几乎总是部件名本身（上盖、下壳体、内筒、侧板、大齿轮…），
# 夹在中间或尾部时才多半是越界（壳体上设…），故定长模式只对首字放行这几个字。
_TERM_HEAD_OK = set("上下中内外前后左右侧大小")


def _is_noisy_fixed_term(core: str) -> bool:
    """定长 n 字模式下"所述"侧术语的噪声判据（core 为去掉序数前缀后的部分）。"""
    if not core:
        return True
    if core[0] in _STOPCHARS and core[0] not in _TERM_HEAD_OK:
        return True
    return _is_noisy_ngram(core[1:])


# 序数前缀「第X」里 X 的取值（「第一连接件」的 min_keep 跨越判定用）
_ORDINAL_CHARS = set("一二三四五六七八九十0123456789０１２３４５６７８９")

# 动态截断的「单字右边界」：只收结构助词 / 连词 / 介词 / 指示词。
# 刻意**不**复用 _STOPCHARS——后者含量词与数词（件 / 个 / 条 / 项 / 一 / 二 …），
# 而「连接件 / 紧固件 / 弹性件 / 第一…」正是高频部件名，拿它做边界会把术语砍断。
_TRUNC_BOUNDARY_CHARS = set("的地得之与和或及并且而则但若如在于对从被将把以由向其该此这那等为是有")

# 术语尾部可剪掉的字：助词 + 单字方位词。同样不含量词 / 数词，
# 所以「第一连接件」不会被剪成「第一连接」，而「壳体上」「连接杆的」能剪干净。
_TRIM_TAIL_CHARS = set("的地得之与和或及并且而在于对从被把以由向上下中内外前后左右里间边旁侧面处")


def _is_all_stopchars(seg: str) -> bool:
    """整串都是停用字 → 肯定不是技术术语。

    动态模式（截断 / 回退）的术语长度可变，越长越容易夹带一个停用字
    （第一连接件 / 上盖 / 安装座上 …）。沿用 _is_noisy_ngram 的「含任一停用字
    即丢弃」会把大量真实部件名整条跳过、造成漏报，故动态术语改用这条更宽松的
    判据；漏报的代价由动态模式自带的自由形式定义集 + 前缀回退来对冲。
    """
    return bool(seg) and all(ch in _STOPCHARS for ch in seg)


def _trim_stopchars(seg: str, keep: int = 0) -> str:
    """去掉术语尾部的助词 / 单字方位词，至多剪到剩 keep 个字。

    动态截断常停在双字方位词之前，尾字往往是「的 / 上 / 中」等
    （安装座上、连接杆的），剪掉之后才是真正的部件名。
    keep 用来护住「第一侧 / 第二面」这类序数 + 单字方位名词的那一个字。
    """
    i = len(seg)
    while i > keep and seg[i - 1] in _TRIM_TAIL_CHARS:
        i -= 1
    return seg[:i]


def _cjk_run(text: str, start: int, limit: int) -> str:
    """从 start 起取连续 CJK 字，最多 limit 个。"""
    end = start
    L = len(text)
    while end < L and end - start < limit and _CJK_RE.match(text[end]):
        end += 1
    return text[start:end]


# 序数前缀：所述第一连接件 / 所述第三弹簧
_ORDINAL_RE = re.compile(r'第[一二三四五六七八九十百零〇两]+')

# 数量前缀：所述多个连接杆 / 所述至少两个光源 / 所述若干支架 / 所述各连接杆。
# 量词只收不构成技术形容词的那几个——层 / 段 / 级 / 孔 会组成 多层结构 / 三段式 /
# 多孔板 这类术语本体，剥掉反而丢信息。
_QUANT_CLS = "个根对组条块片只件套台支"
_QUANTIFIER_RE = re.compile(
    r'(?:至少|至多|不少于|不多于)?'
    rf'(?:(?:[一二两三四五六七八九十]+|多|数)[{_QUANT_CLS}]|若干[{_QUANT_CLS}]?)'
    rf'|(?:各|每)[{_QUANT_CLS}一]?'
)


def _ordinal_len(text: str, start: int) -> int:
    m = _ORDINAL_RE.match(text, start)
    return m.end() - start if m else 0


# 兜底识别"权利要求N所述"引用语：_CITE_RE 解析不了的写法，若"所述"前 12 字内
# 出现"权利要求\d"，也视为引用语公式。
_CLAIM_CITE_PREFIX_RE = re.compile(r'权利要求\s*\d')


def _is_in_citation_formula(text: str, suoshu_pos: int) -> bool:
    lookback_start = max(0, suoshu_pos - 12)
    return bool(_CLAIM_CITE_PREFIX_RE.search(text[lookback_start:suoshu_pos]))


def _find_references(text: str) -> list:
    """返回 [(p, s)]：p 为"所述"位置，s 为被引术语起点（"所述的X"跳过"的"）。

    权利要求引用语里的"所述"（根据权利要求1、2、3、4或5中任一项所述…）不算：
    先按 _CITE_RE 的命中区间排除，长引用串超出 12 字回看窗口也不会漏判。
    """
    cite_spans = [m.span() for m in _CITE_RE.finditer(text)]
    out = []
    for m in _SUOSHU_RE.finditer(text):
        p = m.start()
        if any(a <= p < b for a, b in cite_spans) or _is_in_citation_formula(text, p):
            continue
        s = p + 2
        if s < len(text) and text[s] == "的":
            s += 1
        out.append((p, s))
    return out


def _collect_definitions(text: str, ref_starts: set, max_len: int) -> dict:
    """连续 CJK 段里所有 2..max_len 字子串 → 首次出现位置。

    起点落在被引术语起点（ref_starts）上的子串是引用、不是定义，跳过。
    只排除起点而不盖住一整段固定宽度：固定掩码会越过较短的被引术语，把紧随其后
    真正首次出现的术语也盖掉——「所述多源合束器（5）为二向色镜」里的 二向色镜、
    「所述梁单元力学模型上的最大弯曲应力」里的 最大弯曲应力 都因此登记不进来。
    """
    out: dict = {}
    L = len(text)
    i = 0
    while i < L:
        if not _CJK_RE.match(text[i]):
            i += 1
            continue
        j = i
        while j < L and _CJK_RE.match(text[j]):
            j += 1
        for k in range(i, j - 1):
            if k in ref_starts:
                continue
            for e in range(k + 2, min(j, k + max_len) + 1):
                out.setdefault(text[k:e], k)
        i = j
    return out


def _collect_standalone_definitions(text: str, ref_starts: set, max_len: int,
                                    bl_words: set, bl_lengths) -> dict:
    """与 _collect_definitions 同口径，但只收**右侧到边界**的子串 → 首次出现位置。

    边界：非 CJK / 单字虚词 / 黑名单词 / 「所述」。用来认「较短术语」：
    「对所述装配体的第一侧和第二侧分别加热」里的 第一侧、第二侧 右侧都到了边界，
    能为「所述第一侧加热 / 所述第二侧朝下」提供引用基础；而「安装板」里的 安装
    右侧是"板"，不会被拿来放过「所述安装座」。
    """
    out: dict = {}
    L = len(text)
    i = 0
    while i < L:
        if not _CJK_RE.match(text[i]):
            i += 1
            continue
        j = i
        while j < L and _CJK_RE.match(text[j]):
            j += 1
        for e in range(i + 2, j + 1):
            if e < j and not (
                text[e] in _TRUNC_BOUNDARY_CHARS
                or text.startswith("所述", e)
                or any(text[e:e + n] in bl_words for n in bl_lengths)
            ):
                continue
            for k in range(max(i, e - max_len), e - 1):
                if k not in ref_starts:
                    out.setdefault(text[k:e], k)
        i = j
    return out


# ─────────────────────────────────────────
# 动态截断 / 动态回退 辅助
# ─────────────────────────────────────────
def _extract_term_dynamic_truncate(text: str, start: int,
                                    blacklist_first_chars: set,
                                    blacklist_words: set,
                                    blacklist_lengths,
                                    max_len: int = 12,
                                    min_keep: int = 2) -> str:
    """
    从 text[start:] 起向后扫描 CJK 字符，直到遇到下列任一情况就停下：
      • 非 CJK 字符（标点、空格、英文、数字…）
      • 黑名单词的首字（即位置 k 起 text[k:k+L] 在黑名单里）
      • 单字虚词（的 / 与 / 在 / 于 …，见 _TRUNC_BOUNDARY_CHARS）——黑名单收的是
        双字动词与方位词，单字虚词从来不在其中，导致
        「所述流量调节阀的入口与气源连通」一路吞到 max_len 才停，
        截出的长串当然找不到引用基础；单字虚词恰恰是最可靠的术语右边界
      • 累计达到 max_len（保护用，防止整段被吞）
    blacklist_lengths 为当前黑名单实际出现的词长（降序）；按真实词长匹配，
    使用户自定义的 1 字 / 5+ 字词条也能生效，为空时只按非 CJK 边界截断。
    返回累积到的术语字符串（可能为空）。

    min_keep：前 min_keep 个字**不**触发黑名单中断。黑名单里的动词多数同时是
    高频部件名的构词前缀（安装座 / 连接杆 / 固定板 / 设置槽 …），若在起点就
    中断会截出空串、整条被跳过而漏检。保留前 2 个字后，
    「所述安装座上设有孔」截出「安装座上」，再配合动态回退可继续缩到
    「安装座」→「安装」，不再漏报。
    """
    out = []
    k = start
    L = len(text)
    # 序数前缀「第一 / 第二 / 第3 …」本身不是术语，保护长度要跨过它再起算，
    # 否则「所述第一连接件」会在黑名单词"连接"处断成"第一"。
    keep = min_keep
    if start + 1 < L and text[start] == "第" and text[start + 1] in _ORDINAL_CHARS:
        keep = min_keep + 2
    while k < L and len(out) < max_len:
        ch = text[k]
        if not _CJK_RE.match(ch) or text.startswith("所述", k):
            break
        # 检查从位置 k 起是否命中黑名单的某个词
        # 黑名单首字命中是必要条件，再做一次完整匹配以避免误伤
        # （前 min_keep 个字不判，见 docstring）
        if len(out) >= keep:
            if ch in _TRUNC_BOUNDARY_CHARS:
                break
            if ch in blacklist_first_chars:
                hit = False
                # 按黑名单实际出现的词长尝试完整匹配（长优先）
                for L_word in blacklist_lengths:
                    if k + L_word <= L and text[k:k + L_word] in blacklist_words:
                        hit = True
                        break
                if hit:
                    break
        out.append(ch)
        k += 1
    return "".join(out)


def _build_blacklist_lookup(blacklist) -> tuple:
    """从词列表构造 (首字 set, 完整词 set, 词长列表)。
    词长列表按实际黑名单词长去重并降序排列，供截断时按真实长度匹配（长优先）。"""
    words = {w for w in (blacklist or ()) if w}
    first_chars = {w[0] for w in words if w}
    lengths = sorted({len(w) for w in words}, reverse=True)
    return first_chars, words, lengths


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

    术语 X 的提取策略（X 从"所述"/"所述的"之后起算）：
      • 默认（两个开关都关）：取 n 个 CJK 字；不足 n 字时取实际可用的 ≥2 字
      • 仅 use_dynamic_truncate：按标点 / 单字虚词 / 黑名单词截断，剪掉尾部停用字；
        截断退化（<2 字）或溢出（> DYN_TRUNC_SOFT_MAX）时回落到 n 字
      • 仅 use_dynamic_fallback：取 n 字后从右往左逐字缩短，任一前缀有定义即放过
      • 两个都开：先截断，再对截断结果做回退——误判最低的组合
    序数前缀「第X」不计入 n、回退也不会缩进序数之后不足 2 字（否则「所述第三弹簧」
    会被「第三光路」里的「第三」放过）；数量前缀（多个 / 至少两个 / 各…）不匹配时
    剥掉再试一次。

    定义集：本权项"非引用"位置出现过的全部 2..max_len 字 CJK 子串（见
    _collect_definitions），加上引用链继承。择一引用（或 / 至 / -）要求术语在
    **每个**被引权项的引用链里都有定义——「根据权利要求1或2所述」里只有权利要求 2
    引入的特征，与权利要求 1 组合时就没有引用基础。

    同一缺失术语在本权项及其从属权项里只报首处。
    """
    results = []
    ignore_set = set(ignore_set or ())
    bl_first_chars, bl_words, bl_lengths = _build_blacklist_lookup(boundary_blacklist or [])
    # 认「独立出现的较短术语」时的右边界：不管截断开没开，都用内置 + 用户黑名单
    _, st_words, st_lengths = _build_blacklist_lookup(
        set(DEFAULT_BOUNDARY_BLACKLIST) | set(boundary_blacklist or ()))
    # 定义集子串上限：要能容纳「序数 + n 字」与截断术语
    max_len = max(DYN_TERM_MAX_LEN, n) + 4

    defs: dict = {}        # {claim_no: {term: 本权项内首次位置}}
    stand: dict = {}       # {claim_no: {右侧到边界的 term: 本权项内首次位置}}
    groups_of: dict = {}   # {claim_no: [(须全部满足?, [被引权项…])]}
    reported: dict = {}    # {claim_no: 本权项及其引用链上已报过的术语}
    memo: dict = {}

    def _has(c: int, term: str, tbl=defs) -> bool:
        """term 是否在权项 c 的完整引用链（含 c 自身全文）中有定义（tbl 为 defs / stand）。"""
        key = (c, term, tbl is stand)
        hit = memo.get(key)
        if hit is None:
            hit = term in tbl[c] or _inherited(c, term, tbl)
            memo[key] = hit
        return hit

    def _inherited(no: int, term: str, tbl=defs) -> bool:
        for need_all, nums in groups_of[no]:
            if (all if need_all else any)(_has(c, term, tbl) for c in nums):
                return True
        return False

    for no in sorted(claims.keys()):
        info = claims[no]
        text = info.text
        groups = []
        for g in info.cite_groups:
            nums = sorted(c for c in g["nums"] if c in claims and c < no)
            if nums:
                groups.append((g["mode"] in ("or", "range"), nums))
        groups_of[no] = groups

        refs = _find_references(text)
        ref_starts = set()
        for _, s in refs:
            ref_starts.add(s)
            q = _QUANTIFIER_RE.match(text, s)
            if q:
                ref_starts.add(q.end())
        local = _collect_definitions(text, ref_starts, max_len)
        defs[no] = local
        local_st = _collect_standalone_definitions(text, ref_starts, max_len, st_words, st_lengths)
        stand[no] = local_st
        seen = set()
        for _, nums in groups:
            for c in nums:
                seen |= reported[c]
        reported[no] = seen

        # 权项文本里第 k 行 ↔ para_indices[k]；开头被剥掉的编号部分可能含换行
        line_base = info.raw_text[:len(info.raw_text) - len(text)].count("\n")

        def _candidate(st: int):
            """从术语起点 st 按当前模式取 (起点, 术语, 回退下限 | None, 额外兜底术语)。"""
            ord_len = _ordinal_len(text, st)
            run = _cjk_run(text, st, ord_len + n)
            core_len = len(run) - ord_len
            fixed = run if len(run) >= 2 and core_len >= 1 else ""
            n_term = run if core_len == n else ""
            trunc = ""
            if use_dynamic_truncate:
                trunc = _trim_stopchars(_extract_term_dynamic_truncate(
                    text, st, bl_first_chars, bl_words, bl_lengths, max_len=DYN_TERM_MAX_LEN
                ), ord_len + 1 if ord_len else 0)
                if len(trunc) < 2 or len(trunc) <= ord_len:
                    trunc = ""

            if not use_dynamic_truncate and not use_dynamic_fallback:
                if not fixed or _is_noisy_fixed_term(fixed[ord_len:]):
                    return None
                return st, fixed, None, ""
            if use_dynamic_truncate and not use_dynamic_fallback:
                base = trunc if trunc and len(trunc) - ord_len <= DYN_TRUNC_SOFT_MAX else fixed
                if not base or _is_all_stopchars(base):
                    return None
                return st, base, None, ""
            base = trunc or fixed
            if not base or _is_all_stopchars(base):
                return None
            floor = min(len(base), ord_len + 2) if ord_len else 2
            return st, base, floor, (n_term if use_dynamic_truncate else "")

        for p, s in refs:
            starts = [s]
            q = _QUANTIFIER_RE.match(text, s)
            if q and q.end() > s:
                starts.append(q.end())
            cands = [c for c in map(_candidate, starts) if c]
            if not cands:
                continue

            def _local_before(term: str, tbl=local) -> bool:
                dp = tbl.get(term)
                return dp is not None and dp < p

            def _passes(cand, lookup, lookup_st) -> bool:
                _, base, floor, extra = cand
                if floor is None:
                    tries = (base,)
                else:
                    tries = [base[:k] for k in range(len(base), floor - 1, -1)]
                if any(t in ignore_set or lookup(t) for t in tries):
                    return True
                if extra and lookup(extra):
                    return True
                # 较短前缀在前文**独立**出现过（右侧到边界）也算有基础：
                # 定长 n 字 / 截断越界时「所述温区沿传送方向」「所述第一侧加热」
                # 多带的字不该让真实术语 温区 / 第一侧 判成缺基础。
                # 序数术语可缩到「第X + 1 字」（第一侧 / 第二面），裸序数不行。
                o = _ordinal_len(base, 0)
                return any(lookup_st(base[:k])
                           for k in range(len(base) - 1, (o + 1 if o else 2) - 1, -1))

            lookup_all = lambda t: _local_before(t) or _inherited(no, t)
            lookup_all_st = lambda t: _local_before(t, local_st) or _inherited(no, t, stand)
            if any(_passes(c, lookup_all, lookup_all_st) for c in cands):
                continue
            # 回退模式下报错意味着连最短前缀都没定义，问题实质是那个前缀：
            # 用它做去重 / 忽略键，「所述齿圈套…」「所述齿圈与…」只报一条
            keys = [c[1][:c[2]] if c[2] else c[1] for c in cands]
            if any(k in seen for k in keys):
                continue

            st, base, _, _ = cands[0]
            term = keys[0]
            seen.add(term)
            anchor = text[p:st + len(base)]

            # 择一引用：找出在哪些被引权项的引用链里有、哪些没有
            have, miss = [], []
            for need_all, nums in groups:
                if not need_all or len(nums) < 2:
                    continue
                for c in nums:
                    ok = any(
                        _passes(cd, lambda t, c=c: _local_before(t) or _has(c, t),
                                lambda t, c=c: _local_before(t, local_st) or _has(c, t, stand))
                        for cd in cands
                    )
                    (have if ok else miss).append(c)
            if have and miss:
                message = (
                    f"『{anchor}』在引用权利要求 {'、'.join(map(str, miss))} 时缺少引用基础"
                    f"（仅权利要求 {'、'.join(map(str, have))} 中有）"
                )
            else:
                message = f"『{anchor}』缺少引用基础"

            line = line_base + text.count("\n", 0, p)
            if info.para_indices:
                para_idx = info.para_indices[min(line, len(info.para_indices) - 1)]
            else:
                para_idx = -1
            start = max(0, p - 8)
            end = min(len(text), st + len(base) + 8)
            results.append({
                "kind": "antecedent",
                "claim_no": no,
                "para_idx": para_idx,
                "context": text[start:end].replace("\n", " "),
                "message": message,
                "suggestion": "",
                "term": term,
                "anchor": anchor,
            })
    return results


# ─────────────────────────────────────────
# 句号结尾检查
# ─────────────────────────────────────────
def check_claim_ending_punctuation(claims: dict) -> list:
    """每条权利要求必须以「。」结尾（中国专利撰写规范）"""
    results = []
    for no in sorted(claims.keys()):
        info = claims[no]
        stripped = (info.text or "").rstrip()
        if not stripped:
            continue
        last = stripped[-1]
        if last == "。":
            continue
        # 末段（权项尾部的空段已被 raw_text 剥掉）：定位锚点取它的最后 20 字
        lines = info.raw_text.split("\n")
        k = min(len(lines), len(info.para_indices)) - 1
        results.append({
            "kind": "ending",
            "claim_no": no,
            "para_idx": info.para_indices[k] if k >= 0 else -1,
            "context": stripped[-20:],
            "message": f"权利要求 {no} 未以「。」结尾（结尾字符：{last!r}）",
            "suggestion": "在末尾补「。」",
            "anchor": lines[-1].rstrip()[-20:],
        })
    return results


# ─────────────────────────────────────────
# 聚合入口
# ─────────────────────────────────────────
def run_all_checks(paragraphs, start_idx: int, end_idx: int,
                   n: int, ignore_set=None, vague_words=None,
                   check_vague: bool = True,
                   use_dynamic_truncate: bool = False,
                   use_dynamic_fallback: bool = False,
                   boundary_blacklist=None,
                   ignore_marks: bool = False,
                   marks: dict = None) -> list:
    """
    一次性运行全部检查，返回合并后的结果列表。

    参数:
        paragraphs:   全文段落列表（python-docx Paragraph 对象）
        start_idx, end_idx: 权利要求书段落区间
        n:            术语类检查的滑窗字数（2~6）
        ignore_set:   用户自定义忽略词集合（仅作用于术语类检查）
        vague_words:  覆盖默认 VAGUE_WORDBANK
        use_dynamic_truncate / use_dynamic_fallback / boundary_blacklist:
                      引用基础检查的两个降噪开关，详见 check_antecedent_basis
        ignore_marks: 在去掉附图标记的文本上检查（见 strip_reference_marks），
                      结果与未标注时一致；anchor 映射回带标记的原文，定位 / 高亮照常
        marks:        {编号: 名称}，用于识别无括号的「齿圈1」写法
    """
    view = paragraphs
    stripped_maps = {}   # {段落索引: (去标记文本, 位置映射, 原文)}
    if ignore_marks:
        view = list(paragraphs)
        for i in range(max(0, start_idx), min(end_idx, len(paragraphs))):
            orig = paragraphs[i].text or ""
            text, idx = strip_reference_marks(orig, marks)
            if text != orig:
                view[i] = SimpleNamespace(text=text)
                stripped_maps[i] = (text, idx, orig)

    results = _run_checks(view, start_idx, end_idx, n, ignore_set, vague_words, check_vague,
                          use_dynamic_truncate, use_dynamic_fallback, boundary_blacklist)
    if stripped_maps:
        for r in results:
            _map_anchor_to_original(r, view, end_idx, stripped_maps)
    return results


def _map_anchor_to_original(r: dict, view, end_idx: int, stripped_maps: dict) -> None:
    """把去标记文本上的 anchor 换成原文里对应的那一段（含中间的标记）。

    与 1框 定位同一口径：从 para_idx 所在段起向后找 anchor 首次出现。
    """
    anchor = r.get("anchor")
    pid = r.get("para_idx")
    if not anchor or not isinstance(pid, int) or pid < 0:
        return
    for i in range(pid, min(end_idx, len(view))):
        pos = (view[i].text or "").find(anchor)
        if pos < 0:
            continue
        if i in stripped_maps:
            _text, idx, orig = stripped_maps[i]
            r["anchor"] = orig[idx[pos]:idx[pos + len(anchor) - 1] + 1]
        return


def _run_checks(paragraphs, start_idx, end_idx, n, ignore_set, vague_words, check_vague,
                use_dynamic_truncate, use_dynamic_fallback, boundary_blacklist) -> list:
    claims, duplicates = parse_claims_ex(paragraphs, start_idx, end_idx)
    if not claims:
        return []

    results = []
    # 重号权项（如出现两个 "3."）
    for dup in duplicates:
        results.append({
            "kind": "numbering",
            "claim_no": dup["no"],
            "para_idx": dup["para_idx"],
            "context": dup["context"],
            "message": f"权利要求 {dup['no']} 出现多次（序号重复）",
            "suggestion": "重新编号，保证权项序号从 1 开始连续且不重复",
        })
    # 无独立权利要求：所有权项都引用了其它权项
    if all(info.cites for info in claims.values()):
        first_no = min(claims.keys())
        first_info = claims[first_no]
        results.append({
            "kind": "numbering",
            "claim_no": None,
            "para_idx": first_info.para_indices[0] if first_info.para_indices else -1,
            "context": "、".join(str(x) for x in sorted(claims.keys())),
            "message": "未发现独立权利要求（所有权项均引用其它权项）",
            "suggestion": "检查权利要求 1 是否误写了引用语，独立权项不应含『根据权利要求N所述』",
        })
    results.extend(check_claim_dependency(claims))
    results.extend(check_multi_dependency(claims))
    results.extend(check_claim_numbering(claims))
    results.extend(check_claim_ending_punctuation(claims))
    if check_vague:
        results.extend(check_vague_terms(claims, vague_words))
    results.extend(check_antecedent_basis(
        claims, n, ignore_set or set(),
        use_dynamic_truncate=use_dynamic_truncate,
        use_dynamic_fallback=use_dynamic_fallback,
        boundary_blacklist=boundary_blacklist,
    ))

    # 按 claim_no、para_idx 排序以稳定输出
    def sort_key(r):
        return (r.get("claim_no") or 0, r.get("para_idx") or 0, r.get("kind") or "")
    results.sort(key=sort_key)
    return results
